/**
 * 太平洋货箱清单转换 — 将易通货箱清单映射为太平洋投保清单
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

// ============================================================
// Types
// ============================================================

export interface PacificConversionResult {
  sourceFile: string;
  rowCount: number;
  buffer: Buffer;
}

/** 源列头 → 投保清单列的映射 */
interface ColumnMap {
  fbaId: number;          // FBA ID → 入仓编号
  cnName: number;         // 中文品名 → DESCRIPTION (前半)
  enName: number;         // 英文品名 → DESCRIPTION (后半)
  totalQty: number;       // 申报总数量 → QTY PCS
  unitValue: number;      // 单个产品申报货值(USD) → UNIT VALUE
  totalValue: number;     // 总申报货值 → TOTAL VALUE
  currency: number;       // 申报币种 → 币种
  totalBoxes: number;     // 总箱数(CTN) → CTNS
  grossWeight: number;    // 单箱货物毛重(KG) → G.W.(KG)
  lengthCm: number;       // 长(CM) → MEASUREMENT 计算
  widthCm: number;        // 宽(CM) → MEASUREMENT 计算
  heightCm: number;       // 高(CM) → MEASUREMENT 计算
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
        (map as Record<string, number>)[key] = i; // 0-based index
        break;
      }
    }
  }

  // Validate required fields
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
// Main conversion
// ============================================================

export async function convertPacificInsurance(
  filePath: string,
  sourceFileName: string,
): Promise<PacificConversionResult> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(filePath);

  // Use the first sheet (source data)
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

  // Collect data rows (after header, until first empty FBA ID)
  interface DataRow {
    fbaId: string;
    description: string;
    qtyPcs: number;
    unitValue: number;
    totalValue: number;
    currency: string;
    ctns: number;
    grossWeight: number;
    measurement: number;
  }

  const dataRows: DataRow[] = [];

  for (let r = headerRowNum + 1; r <= srcSheet.rowCount; r++) {
    const row = srcSheet.getRow(r);
    const fbaCell = row.getCell(colMap.fbaId + 1);
    const fbaId = fbaCell.value?.toString().trim() ?? "";

    // Stop at empty FBA ID
    if (!fbaId) break;

    const cnName = row.getCell(colMap.cnName + 1).value?.toString().trim() ?? "";
    const enName = row.getCell(colMap.enName + 1).value?.toString().trim() ?? "";
    const description = cnName + enName;

    const qtyPcs        = parseFloat(String(row.getCell(colMap.totalQty + 1).value ?? 0)) || 0;
    const unitValue     = parseFloat(String(row.getCell(colMap.unitValue + 1).value ?? 0)) || 0;
    const totalValue    = parseFloat(String(row.getCell(colMap.totalValue + 1).value ?? 0)) || 0;
    const currency      = row.getCell(colMap.currency + 1).value?.toString().trim() ?? "USD";
    const ctns          = parseFloat(String(row.getCell(colMap.totalBoxes + 1).value ?? 0)) || 0;
    const grossWeight   = parseFloat(String(row.getCell(colMap.grossWeight + 1).value ?? 0)) || 0;
    const lengthCm      = parseFloat(String(row.getCell(colMap.lengthCm + 1).value ?? 0)) || 0;
    const widthCm       = parseFloat(String(row.getCell(colMap.widthCm + 1).value ?? 0)) || 0;
    const heightCm      = parseFloat(String(row.getCell(colMap.heightCm + 1).value ?? 0)) || 0;

    // MEASUREMENT(CBM) = 长 × 宽 × 高 / 1,000,000
    const measurement = (lengthCm * widthCm * heightCm) / 1000000;

    dataRows.push({
      fbaId,
      description,
      qtyPcs,
      unitValue,
      totalValue,
      currency,
      ctns,
      grossWeight,
      measurement,
    });
  }

  if (dataRows.length === 0) {
    throw new Error("未找到数据行。请确认上传的文件包含货箱清单数据。");
  }

  console.log(`📦 Parsed ${dataRows.length} data rows`);

  // ============================================================
  // Build output workbook
  // ============================================================

  const outWb = new ExcelJS.Workbook();
  const outSheet = outWb.addWorksheet("太平洋投保清单");

  // Define output columns
  const outHeaders = [
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

  // Write header row
  const headerRowOut = outSheet.getRow(1);
  outHeaders.forEach((h, i) => {
    headerRowOut.getCell(i + 1).value = h;
  });
  headerRowOut.font = { bold: true, size: 10 };
  headerRowOut.alignment = { wrapText: true, vertical: "middle", horizontal: "center" };
  headerRowOut.height = 50;

  // Set column widths
  const colWidths = [22, 40, 12, 14, 16, 8, 10, 12, 14, 16, 22, 22, 18, 18, 22, 22, 24, 18];
  colWidths.forEach((w, i) => {
    outSheet.getColumn(i + 1).width = w;
  });

  // Write data rows
  for (let i = 0; i < dataRows.length; i++) {
    const d = dataRows[i];
    const row = outSheet.getRow(i + 2);

    row.getCell(1).value = d.fbaId;
    row.getCell(2).value = d.description;
    row.getCell(3).value = d.qtyPcs;
    row.getCell(4).value = d.unitValue;
    row.getCell(5).value = d.totalValue;
    row.getCell(6).value = d.currency;
    row.getCell(7).value = d.ctns;
    row.getCell(8).value = d.grossWeight;
    row.getCell(9).value = Math.round(d.measurement * 1000000) / 1000000; // 保留6位小数

    // Columns 10-18 left empty for manual fill

    row.alignment = { vertical: "middle" };
  }

  // Add borders
  const lastRow = dataRows.length + 1;
  for (let r = 1; r <= lastRow; r++) {
    for (let c = 1; c <= 18; c++) {
      outSheet.getCell(r, c).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
    }
  }

  // Freeze header row
  outSheet.views = [{ state: "frozen", ySplit: 1 }];

  const buffer = await outWb.xlsx.writeBuffer();
  return {
    sourceFile: sourceFileName,
    rowCount: dataRows.length,
    buffer: Buffer.from(buffer),
  };
}
