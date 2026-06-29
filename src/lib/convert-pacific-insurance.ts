/**
 * 太平洋货箱清单转换 — 将易通货箱清单映射为太平洋投保清单，并按单箱货值区间拆分
 *
 * 映射关系（来自模板 太平洋货箱清单转换模板.xlsx 的公式）：
 *   入仓编号        ← FBA ID
 *   DESCRIPTION     ← 中文品名 + 英文品名
 *   QTY PCS         ← 申报总数量
 *   UNIT VALUE      ← 单个产品申报货值(USD)
 *   TOTAL VALUE     ← 总申报货值
 *   币种            ← 申报币种
 *   CTNS            ← 总箱数(CTN)
 *   G.W.(KG)        ← 单箱货物毛重(KG)
 *   MEASUREMENT(CBM)← 长(CM) × 宽(CM) × 高(CM) / 1,000,000
 */

import ExcelJS from "exceljs";
import JSZip from "jszip";

// ============================================================
// Types
// ============================================================

/** 用户可配置的汇率 */
export interface ExchangeRates {
  USD: number;
  EUR: number;
  GBP: number;
  JPY: number;
}

/** 一个区间桶的描述 */
interface IntervalSpec {
  name: string;      // e.g. "5000-10000RMB"
  min: number;       // inclusive lower bound
  max: number;       // exclusive upper bound (or Infinity for catch-all)
}

/** 一行解析后的源数据，附加 perBoxRMB */
export interface PacificDataRow {
  fbaId: string;
  description: string;
  qtyPcs: number;
  unitValue: number;
  totalValue: number;
  currency: string;
  ctns: number;
  grossWeight: number;
  measurement: number;
  perBoxRMB: number;
  exchangeRate: number;
}

/** 一个区间的完整结果 */
export interface PacificIntervalResult {
  name: string;
  fileName: string;
  rowCount: number;
  rows: PacificDataRow[];
  buffer: Buffer;
}

/** 顶层拆分结果（存入 session，不含 buffers 返回给客户端） */
export interface PacificSplitResult {
  sourceFile: string;
  totalRows: number;
  skippedRows: number;
  intervals: PacificIntervalResult[];
}

/** 源列头 → 投保清单列的映射 */
interface ColumnMap {
  fbaId: number;
  cnName: number;
  enName: number;
  totalQty: number;
  unitValue: number;
  totalValue: number;
  currency: number;
  totalBoxes: number;
  grossWeight: number;
  lengthCm: number;
  widthCm: number;
  heightCm: number;
}

// ============================================================
// Header-based column detection
// ============================================================

const HEADER_PATTERNS: Record<keyof ColumnMap, string[]> = {
  fbaId:        ["FBA ID", "fba id", "FBAID"],
  cnName:       ["中文品名"],
  enName:       ["英文品名"],
  totalQty:     ["申报总数量"],
  unitValue:    ["单个产品申报货值", "单个产品申报货值(USD)"],
  totalValue:   ["总申报货值"],
  currency:     ["申报币种"],
  totalBoxes:   ["总箱数", "总箱数(CTN)"],
  grossWeight:  ["单箱货物毛重", "单箱货物毛重(KG)", "毛重"],
  lengthCm:     ["长(CM)", "长(cm)", "长"],
  widthCm:      ["宽(CM)", "宽(cm)", "宽"],
  heightCm:     ["高(CM)", "高(cm)", "高"],
};

function buildColumnMap(headers: (string | null | undefined)[]): ColumnMap {
  const map: Partial<ColumnMap> = {};

  for (const [key, patterns] of Object.entries(HEADER_PATTERNS)) {
    for (let i = 0; i < headers.length; i++) {
      const h = (headers[i] ?? "").toString().trim();
      if (patterns.some((p) => h === p || h.includes(p))) {
        (map as Record<string, number>)[key] = i;
        break;
      }
    }
  }

  const required: (keyof ColumnMap)[] = [
    "fbaId", "cnName", "enName", "totalQty", "unitValue",
    "totalValue", "currency", "totalBoxes", "grossWeight",
    "lengthCm", "widthCm", "heightCm",
  ];
  const missing = required.filter((k) => map[k] === undefined);
  if (missing.length > 0) {
    throw new Error(
      `无法在表头中找到以下列: ${missing.join(", ")}。请确认上传的是易通货箱清单模板。`
    );
  }

  return map as ColumnMap;
}

// ============================================================
// Output workbook definition (shared)
// ============================================================

const OUT_HEADERS = [
  "入仓编号\n（FBA仓和菜鸟仓必填）",
  "DESCRIPTION  OF GOODS                              \n(中英文 货物品名)",
  "QTY  PCS\n（总数量/总个数）",
  "UNIT VALUE \n单个单价",
  "TOTAL VALUE \n 投保总货值\n（不加成总金额）",
  "币种\n（下拉款选择）",
  "CTNS\n(总件数/总箱数)",
  "G.W.(KG)\n（毛重）",
  "MEASUREMENT(CBM)\n（体积）",
  "提单号",
  "船名航次（海运）\n航次（空运）\n班列（铁路必填）\n卡航（国内/国外车牌号）\n全程快递（无需填）",
  "*柜号\n（买单的可以填供应商名称）",
  "后端派送方式\n（按实际情况改填）",
  "快递单号\n（快递派送必填，\n卡派无需填）",
  "起运地-中转-目的地",
  "渠道【头程+尾端】\n(海运/空运/快递/卡航）",
  "详细地址（目的地）",
  "特殊请备注",
];

const OUT_COL_WIDTHS = [22, 40, 12, 14, 16, 8, 10, 12, 14, 16, 22, 22, 18, 18, 22, 22, 24, 18];

// ============================================================
// Interval generation
// ============================================================

function generateDynamicIntervals(maxRMB: number): IntervalSpec[] {
  const intervals: IntervalSpec[] = [
    { name: "不足5000RMB",         min: 0,     max: 5000 },
    { name: "5000-10000RMB",       min: 5000,  max: 10000 },
    { name: "10000-20000RMB",      min: 10000, max: 20000 },
    { name: "20000-30000RMB",      min: 20000, max: 30000 },
    { name: "30000-40000RMB",      min: 30000, max: 40000 },
  ];

  let lower = 40000;
  while (lower < maxRMB) {
    const upper = lower + 10000;
    intervals.push({
      name: `${lower / 10000}万-${upper / 10000}万RMB`,
      min: lower,
      max: upper,
    });
    lower = upper;
  }

  // Last interval is a catch-all
  intervals[intervals.length - 1].max = Number.POSITIVE_INFINITY;

  return intervals;
}

// ============================================================
// Build one interval workbook from scratch
// ============================================================

function buildIntervalWorkbook(rows: PacificDataRow[]): ExcelJS.Workbook {
  const wb = new ExcelJS.Workbook();
  const sheet = wb.addWorksheet("太平洋投保清单");

  // Header row
  const headerRow = sheet.getRow(1);
  OUT_HEADERS.forEach((h, i) => {
    headerRow.getCell(i + 1).value = h;
  });
  headerRow.font = { bold: true, size: 10 };
  headerRow.alignment = { wrapText: true, vertical: "middle", horizontal: "center" };
  headerRow.height = 50;

  // Column widths
  OUT_COL_WIDTHS.forEach((w, i) => {
    sheet.getColumn(i + 1).width = w;
  });

  // Data rows
  for (let i = 0; i < rows.length; i++) {
    const d = rows[i];
    const row = sheet.getRow(i + 2);

    row.getCell(1).value = d.fbaId;
    row.getCell(2).value = d.description;
    row.getCell(3).value = d.qtyPcs;
    row.getCell(4).value = d.unitValue;
    row.getCell(5).value = d.totalValue;
    row.getCell(6).value = d.currency;
    row.getCell(7).value = d.ctns;
    row.getCell(8).value = d.grossWeight;
    row.getCell(9).value = Math.round(d.measurement * 1000000) / 1000000;

    row.alignment = { vertical: "middle" };
  }

  // Borders
  const lastRow = rows.length + 1;
  for (let r = 1; r <= lastRow; r++) {
    for (let c = 1; c <= 18; c++) {
      sheet.getCell(r, c).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
    }
  }

  // Freeze header
  sheet.views = [{ state: "frozen", ySplit: 1 }];

  return wb;
}

// ============================================================
// Main entry: convert + split
// ============================================================

export async function convertAndSplitPacific(
  filePath: string,
  sourceFileName: string,
  rates: ExchangeRates,
): Promise<PacificSplitResult> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(filePath);

  const srcSheet = wb.worksheets[0];
  if (!srcSheet) {
    throw new Error("Excel 文件中没有找到工作表");
  }

  // Find header row — scan first 10 rows for "FBA ID"
  let headerRow: ExcelJS.Row | null = null;
  let headerRowNum = 0;
  for (let r = 1; r <= Math.min(srcSheet.rowCount, 10); r++) {
    const row = srcSheet.getRow(r);
    let hasFba = false;
    row.eachCell((cell) => {
      const v = cell.value?.toString().trim() ?? "";
      if (v === "FBA ID" || v === "fba id") hasFba = true;
    });
    if (hasFba) {
      headerRow = row;
      headerRowNum = r;
      break;
    }
  }

  if (!headerRow) {
    throw new Error("找不到表头行（需要包含 'FBA ID' 列）。请确认上传的是易通货箱清单模板。");
  }

  // Read headers
  const headers: (string | null | undefined)[] = [];
  headerRow.eachCell((cell, cn) => {
    headers[cn - 1] = cell.value?.toString().trim() ?? null;
  });

  const colMap = buildColumnMap(headers);
  console.log(`📋 Header row: ${headerRowNum}, columns mapped: ${Object.keys(colMap).length}`);

  // =========================================================
  // Phase 1: read all raw rows (including ctns=0)
  // =========================================================
  interface RawRow {
    fbaId: string;
    cnName: string;
    enName: string;
    qtyPcs: number;
    unitValue: number;
    totalValue: number;
    currency: string;
    ctns: number;
    grossWeight: number;
    lengthCm: number;
    widthCm: number;
    heightCm: number;
    rowNum: number;
  }

  const rawRows: RawRow[] = [];

  for (let r = headerRowNum + 1; r <= srcSheet.rowCount; r++) {
    const row = srcSheet.getRow(r);
    const fbaCell = row.getCell(colMap.fbaId + 1);
    const fbaId = fbaCell.value?.toString().trim() ?? "";

    if (!fbaId) break;

    const cnName = row.getCell(colMap.cnName + 1).value?.toString().trim() ?? "";
    const enName = row.getCell(colMap.enName + 1).value?.toString().trim() ?? "";
    const qtyPcs      = parseFloat(String(row.getCell(colMap.totalQty + 1).value ?? 0)) || 0;
    const unitValue   = parseFloat(String(row.getCell(colMap.unitValue + 1).value ?? 0)) || 0;
    const totalValue  = parseFloat(String(row.getCell(colMap.totalValue + 1).value ?? 0)) || 0;
    const currency    = row.getCell(colMap.currency + 1).value?.toString().trim() ?? "USD";
    const ctns        = parseFloat(String(row.getCell(colMap.totalBoxes + 1).value ?? 0)) || 0;
    const grossWeight = parseFloat(String(row.getCell(colMap.grossWeight + 1).value ?? 0)) || 0;
    const lengthCm    = parseFloat(String(row.getCell(colMap.lengthCm + 1).value ?? 0)) || 0;
    const widthCm     = parseFloat(String(row.getCell(colMap.widthCm + 1).value ?? 0)) || 0;
    const heightCm    = parseFloat(String(row.getCell(colMap.heightCm + 1).value ?? 0)) || 0;

    rawRows.push({
      fbaId, cnName, enName, qtyPcs, unitValue, totalValue,
      currency, ctns, grossWeight, lengthCm, widthCm, heightCm,
      rowNum: r,
    });
  }

  // =========================================================
  // Phase 2: compute merged perBoxRMB for each box group
  //
  // Rows with ctns=0 share a box with the preceding ctns>0 row.
  // We merge the TOTAL VALUE across the group to compute perBoxRMB
  // (used for interval assignment), but each row stays individual
  // in the output Excel.
  // =========================================================
  // Map: rawRow index → perBoxRMB (same value for all rows in a box group)
  const perBoxRMBByIndex = new Map<number, number>();

  for (let i = 0; i < rawRows.length; i++) {
    const raw = rawRows[i];

    if (raw.ctns > 0) {
      // Start of a box group — sum totalValue of this row + any following ctns=0 rows
      let mergedTotalValue = raw.totalValue;
      let mergeCount = 0;

      let j = i + 1;
      while (j < rawRows.length && rawRows[j].ctns === 0) {
        mergedTotalValue += rawRows[j].totalValue;
        mergeCount++;
        j++;
      }

      if (mergeCount > 0) {
        console.log(
          `🔗 Row ${raw.rowNum}: box group with ${mergeCount} ctns=0 row(s) — merged totalValue=${mergedTotalValue.toFixed(2)}`
        );
      }

      const exchangeRate = rates[raw.currency as keyof ExchangeRates] ?? 1;
      const perBoxRMB = (mergedTotalValue / raw.ctns) * exchangeRate;

      // Assign the same perBoxRMB to all rows in this group
      for (let k = i; k < j; k++) {
        perBoxRMBByIndex.set(k, perBoxRMB);
      }
    }
    // ctns=0 rows: handled in the look-ahead above; if orphan (first row ctns=0),
    // they won't appear in perBoxRMBByIndex at all.
  }

  // =========================================================
  // Phase 3: build allRows — every raw row becomes an output row
  // =========================================================
  const allRows: PacificDataRow[] = [];
  let skippedRows = 0;

  for (let i = 0; i < rawRows.length; i++) {
    const raw = rawRows[i];
    const perBoxRMB = perBoxRMBByIndex.get(i);

    if (perBoxRMB === undefined) {
      // Orphan ctns=0 row at the very beginning — no group to join
      skippedRows++;
      console.log(`⚠️  Row ${raw.rowNum}: ctns=0 without preceding box — skipping`);
      continue;
    }

    const exchangeRate = rates[raw.currency as keyof ExchangeRates] ?? 1;
    const measurement = (raw.lengthCm * raw.widthCm * raw.heightCm) / 1000000;

    allRows.push({
      fbaId: raw.fbaId,
      description: raw.cnName + raw.enName,
      qtyPcs: raw.qtyPcs,
      unitValue: raw.unitValue,
      totalValue: raw.totalValue,
      currency: raw.currency,
      ctns: raw.ctns,
      grossWeight: raw.grossWeight,
      measurement,
      perBoxRMB,
      exchangeRate,
    });
  }

  if (allRows.length === 0) {
    throw new Error("未找到有效数据行。请确认上传的文件包含货箱清单数据。");
  }

  console.log(`📦 Parsed ${allRows.length} data rows, skipped ${skippedRows}`);

  // Generate dynamic intervals
  const maxRMB = Math.max(...allRows.map((r) => r.perBoxRMB));
  const intervals = generateDynamicIntervals(maxRMB);
  console.log(`📊 Generated ${intervals.length} intervals (max perBoxRMB = ${maxRMB.toFixed(2)})`);

  // Assign rows to intervals
  const buckets: PacificDataRow[][] = intervals.map(() => []);
  for (const row of allRows) {
    const idx = intervals.findIndex(
      (iv) => row.perBoxRMB >= iv.min && row.perBoxRMB < iv.max
    );
    if (idx >= 0) {
      buckets[idx].push(row);
    }
  }

  // Build output per interval
  const intervalResults: PacificIntervalResult[] = [];
  for (let i = 0; i < intervals.length; i++) {
    const iv = intervals[i];
    const rows = buckets[i];
    const wb = buildIntervalWorkbook(rows);
    const buffer = Buffer.from(await wb.xlsx.writeBuffer());

    // Generate file name matching insurance split naming convention
    const fileName = iv.name.replace(/\//g, "-") + ".xlsx";

    intervalResults.push({
      name: iv.name,
      fileName,
      rowCount: rows.length,
      rows,
      buffer,
    });
  }

  console.log(
    `✅ Split complete: ${intervalResults.filter((iv) => iv.rowCount > 0).length} non-empty intervals`
  );

  return {
    sourceFile: sourceFileName,
    totalRows: allRows.length,
    skippedRows,
    intervals: intervalResults,
  };
}

// ============================================================
// ZIP generation
// ============================================================

export async function generatePacificZip(result: PacificSplitResult): Promise<Buffer> {
  const zip = new JSZip();

  for (const iv of result.intervals) {
    zip.file(iv.fileName, iv.buffer);
  }

  const summary = {
    sourceFile: result.sourceFile,
    generatedAt: new Date().toISOString(),
    totalRows: result.totalRows,
    skippedRows: result.skippedRows,
    intervals: result.intervals.map((iv) => ({
      name: iv.name,
      fileName: iv.fileName,
      rowCount: iv.rowCount,
    })),
  };
  zip.file("pacific_split_summary.json", JSON.stringify(summary, null, 2));

  return Buffer.from(await zip.generateAsync({ type: "nodebuffer" }));
}
