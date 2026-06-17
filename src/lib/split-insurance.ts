/**
 * 保单拆分核心逻辑
 *
 * 使用 exceljs 完整保留原模板格式（字体、颜色、合并单元格、图片等）。
 * 策略：每个区间从原始 buffer 反序列化一份克隆 → 删掉不属于该区间的行 →
 * 修改头部 → 添加辅助列 W → 输出。
 */

import ExcelJS from "exceljs";
import JSZip from "jszip";
import { readFileSync } from "fs";

// ============================================================
// Constants
// ============================================================

const SOURCE_SHEET = "ETTON电商物流 下单模板";

const RATES: Record<string, number> = { USD: 7, EUR: 8 };

const INTERVALS = [
  { name: "不足5000RMB", min: 0, max: 5000 },
  { name: "5000-10000RMB", min: 5000, max: 10000 },
  { name: "10000-20000RMB", min: 10000, max: 20000 },
  { name: "20000-30000RMB", min: 20000, max: 30000 },
  { name: "30000-40000RMB", min: 30000, max: 40000 },
];

// 0-indexed column indices
const COL_A = 0; // 订单号
const COL_G = 6; // 箱数
const COL_J = 9; // 总价
const COL_K = 10; // 币种
const COL_Q = 16; // 重量(KG)
const COL_R = 17; // 长(CM)
const COL_S = 18; // 宽(CM)
const COL_T = 19; // 高(CM)

// ============================================================
// Types
// ============================================================

export interface BoxGroup {
  rows: number[]; // 1-indexed Excel row numbers
  boxes: number;
  totalPrice: number;
  currency: string;
  rate: number;
  perBoxOrig: number;
  perBoxRMB: number;
}

export interface IntervalResult {
  name: string;
  fileName: string;
  totalBoxes: number;
  totalWeight: number;
  totalVolume: number;
  groupCount: number;
  groups: BoxGroup[];
  buffer: Buffer;
}

export interface SplitResult {
  sourceFile: string;
  totalBoxes: number;
  totalGroups: number;
  intervals: IntervalResult[];
}

// ============================================================
// Helpers
// ============================================================

/** Safely read a numeric cell value from an ExcelJS row */
function numVal(row: ExcelJS.Row, colIdx: number): number {
  const cell = row.getCell(colIdx + 1); // ExcelJS is 1-indexed
  const v = cell.value;
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return v;
  if (typeof v === "string") {
    const n = parseFloat(v);
    return isNaN(n) ? 0 : n;
  }
  // Could be a rich-text or formula result object
  if (typeof v === "object") {
    if ("result" in v) return numValFromAny(v.result);
    if ("richText" in v) {
      const text = (v as { richText: Array<{ text: string }> }).richText
        .map((r) => r.text)
        .join("");
      const n = parseFloat(text);
      return isNaN(n) ? 0 : n;
    }
  }
  return 0;
}

function numValFromAny(v: unknown): number {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return v;
  if (typeof v === "string") {
    const n = parseFloat(v);
    return isNaN(n) ? 0 : n;
  }
  return 0;
}

function strVal(row: ExcelJS.Row, colIdx: number): string {
  const cell = row.getCell(colIdx + 1);
  const v = cell.value;
  if (v === null || v === undefined) return "";
  if (typeof v === "string") return v.trim();
  if (typeof v === "number") return String(v);
  if (typeof v === "object") {
    if ("richText" in v && Array.isArray((v as { richText: unknown }).richText)) {
      return (v as { richText: Array<{ text: string }> }).richText
        .map((r) => r.text)
        .join("")
        .trim();
    }
    if ("result" in v) return String((v as { result: unknown }).result ?? "").trim();
  }
  return String(v ?? "").trim();
}

/** Determine which interval a perBoxRMB value belongs to */
function findInterval(rmb: number): number {
  for (let i = 0; i < INTERVALS.length; i++) {
    if (i === INTERVALS.length - 1) {
      // Last interval: >= min
      if (rmb >= INTERVALS[i].min) return i;
    } else {
      if (rmb >= INTERVALS[i].min && rmb < INTERVALS[i].max) return i;
    }
  }
  return INTERVALS.length - 1; // fallback
}

// ============================================================
// Phase 1: Parse & Group
// ============================================================

function parseAndGroup(worksheet: ExcelJS.Worksheet) {
  const groups: BoxGroup[] = [];
  let currentGroup: BoxGroup | null = null;

  // Data starts at row 25 (1-indexed)
  let rowNum = 25;
  while (true) {
    const row = worksheet.getRow(rowNum);
    const aVal = strVal(row, COL_A);
    if (aVal === "") break; // A column empty → end of data

    const gVal = numVal(row, COL_G);
    const jVal = numVal(row, COL_J);
    const kVal = strVal(row, COL_K) || "USD";
    const rate = RATES[kVal] || 7;
    const boxes = gVal > 0 ? gVal : 0;

    if (boxes > 0) {
      // New box group
      if (currentGroup) groups.push(currentGroup);
      currentGroup = {
        rows: [rowNum],
        boxes,
        totalPrice: jVal,
        currency: kVal,
        rate,
        perBoxOrig: 0,
        perBoxRMB: 0,
      };
    } else {
      // Mixed row (混箱)
      if (!currentGroup) {
        currentGroup = {
          rows: [rowNum],
          boxes: 1,
          totalPrice: jVal,
          currency: "USD",
          rate: 7,
          perBoxOrig: 0,
          perBoxRMB: 0,
        };
      } else {
        currentGroup.rows.push(rowNum);
        currentGroup.totalPrice += jVal;
      }
    }
    rowNum++;
  }
  if (currentGroup) groups.push(currentGroup);

  // Calculate per-box RMB
  for (const g of groups) {
    g.perBoxOrig = g.totalPrice / g.boxes;
    g.perBoxRMB = g.perBoxOrig * g.rate;
  }

  return { groups, lastDataRow: rowNum - 1 };
}

// ============================================================
// Phase 2: Calculate interval summaries
// ============================================================

function calcIntervalData(groups: BoxGroup[], worksheet: ExcelJS.Worksheet) {
  const intervalGroups: BoxGroup[][] = INTERVALS.map(() => []);

  for (const g of groups) {
    const idx = findInterval(g.perBoxRMB);
    intervalGroups[idx].push(g);
  }

  return INTERVALS.map((iv, idx) => {
    const ivGroups = intervalGroups[idx];
    const totalBoxes = ivGroups.reduce((s, g) => s + g.boxes, 0);

    let totalWeight = 0;
    for (const g of ivGroups) {
      const firstRow = worksheet.getRow(g.rows[0]);
      const firstQ = numVal(firstRow, COL_Q);
      totalWeight += firstQ * g.boxes;
      for (let i = 1; i < g.rows.length; i++) {
        const mixRow = worksheet.getRow(g.rows[i]);
        totalWeight += numVal(mixRow, COL_Q);
      }
    }

    let totalVolume = 0;
    for (const g of ivGroups) {
      const firstRow = worksheet.getRow(g.rows[0]);
      const rVal = numVal(firstRow, COL_R);
      const sVal = numVal(firstRow, COL_S);
      const tVal = numVal(firstRow, COL_T);
      totalVolume += ((rVal * sVal * tVal) / 1000000) * g.boxes;
    }

    return {
      name: iv.name,
      fileName: `${iv.name}.xlsx`,
      totalBoxes,
      totalWeight: Math.round(totalWeight * 100) / 100,
      totalVolume: Math.round(totalVolume * 10000) / 10000,
      groupCount: ivGroups.length,
      groups: ivGroups,
    };
  });
}

// ============================================================
// Phase 3: Generate output buffer for one interval
//
// Strategy: clone the source workbook via buffer round-trip, then delete
// data rows that don't belong.  This preserves formatting, merged cells,
// floating images, and auxiliary sheets natively.
//
// DISPIMG ("Place in Cell") images are NOT supported by exceljs at all —
// their data is lost on load.  We extract them from the original xlsx zip
// BEFORE exceljs ever touches them, then re-place them as floating images
// in the clone.
// ============================================================

/**
 * Flatten shared formulas so spliceRows doesn't break them.
 * DISPIMG cells were already cleared + re-imaged by embedMediaImagesInWorkbook
 * before cloning, so we only need to handle shared formulas here.
 */
function flattenFormulas(ws: ExcelJS.Worksheet) {
  ws.eachRow({ includeEmpty: false }, (row) => {
    row.eachCell({ includeEmpty: false }, (cell) => {
      const v = cell.value;
      if (!v || typeof v !== "object") return;

      const obj = v as unknown as Record<string, unknown>;

      // Shared-formula clone → just keep the computed result
      if ("sharedFormula" in obj) {
        cell.value = (obj.result ?? null) as ExcelJS.CellValue;
        return;
      }

      // Shared-formula master → turn into regular formula
      const formula = obj.formula;
      if (typeof formula === "string" && ("shareType" in obj || "ref" in obj)) {
        cell.value = {
          formula,
          result: obj.result,
        } as ExcelJS.CellFormulaValue;
      }
    });
  });
}

// ============================================================
// DISPIMG preservation: inject cellImages.xml + media back into
// the output xlsx so =DISPIMG() formulas resolve natively.
// ============================================================

/**
 * Inject cellImages.xml, its .rels, all media files, and required
 * [Content_Types].xml entries from the original xlsx into the output buffer.
 * This makes =DISPIMG() formulas work natively in Excel / WPS.
 */
async function injectCellImagesIntoOutput(
  outBuf: Buffer,
  src: string | Uint8Array
): Promise<Buffer> {
  const srcData = typeof src === "string" ? readFileSync(src) : src;
  const srcZip = await JSZip.loadAsync(srcData);
  const outZip = await JSZip.loadAsync(outBuf);

  // --- Copy cellImages.xml ---
  const ciPath = Object.keys(srcZip.files).find(
    (f) => /xl\/cellimages\.xml$/i.test(f) && !srcZip.files[f].dir
  );
  if (!ciPath) return outBuf;
  outZip.file(ciPath, await srcZip.file(ciPath)!.async("nodebuffer"));

  // --- Copy cellImages.xml.rels ---
  const ciRelsPath = Object.keys(srcZip.files).find(
    (f) => /xl\/_rels\/cellimages\.xml\.rels$/i.test(f) && !srcZip.files[f].dir
  );
  if (ciRelsPath) {
    outZip.file(ciRelsPath, await srcZip.file(ciRelsPath)!.async("nodebuffer"));
  }

  // --- Copy all media files from xl/media/ ---
  const mediaFolder = srcZip.folder("xl/media");
  if (mediaFolder) {
    for (const [name, file] of Object.entries(mediaFolder.files)) {
      if (file.dir) continue;
      outZip.file(name, await file.async("nodebuffer"));
    }
  }

  // --- Update [Content_Types].xml ---
  const ctPath = "[Content_Types].xml";
  let ctXml = await outZip.file(ctPath)!.async("string");

  // Add cellImages override if missing
  if (!ctXml.includes("cellimages")) {
    ctXml = ctXml.replace(
      "</Types>",
      '<Override PartName="/xl/cellImages.xml" ContentType="application/vnd.ms-excel.cellimages+xml"/></Types>'
    );
  }

  // Add Default entries for image types if missing
  for (const [ext, mime] of [
    ["jpeg", "image/jpeg"],
    ["png", "image/png"],
    ["gif", "image/gif"],
  ]) {
    if (!ctXml.includes(`Extension="${ext}"`)) {
      ctXml = ctXml.replace(
        "</Types>",
        `<Default Extension="${ext}" ContentType="${mime}"/></Types>`
      );
    }
  }

  outZip.file(ctPath, ctXml);

  return Buffer.from(await outZip.generateAsync({ type: "nodebuffer" }));
}

/**
 * Generate one split file from a pre-serialized clone buffer.
 * `srcBuffer` must be the output of `workbook.xlsx.writeBuffer()` from the
 * original workbook so each interval starts from an identical pristine clone.
 */
async function generateIntervalFile(
  srcBuffer: Buffer,
  intervalData: ReturnType<typeof calcIntervalData>[number]
): Promise<Buffer> {
  // --- Work from a pristine clone ---
  const wb = new ExcelJS.Workbook();
  // @ts-expect-error — exceljs Buffer type vs @types/node generic Buffer
  await wb.xlsx.load(srcBuffer);

  const ws = wb.getWorksheet(SOURCE_SHEET);
  if (!ws) throw new Error(`Sheet "${SOURCE_SHEET}" not found`);

  // --- Flatten shared formulas BEFORE any row mutation ---
  flattenFormulas(ws);

  // --- Determine source rows to keep (sorted) ---
  const keptRows = intervalData.groups.flatMap((g) => g.rows).sort((a, b) => a - b);

  // --- Find ALL data rows in the clone (row 25+) ---
  const allDataRows: number[] = [];
  for (let r = 25; ; r++) {
    const aVal = strVal(ws.getRow(r), COL_A);
    if (aVal === "") break;
    allDataRows.push(r);
  }

  // --- Delete unwanted data rows from BOTTOM to TOP ---
  const keptSet = new Set(keptRows);
  const toDelete = allDataRows.filter((r) => !keptSet.has(r));
  for (let i = toDelete.length - 1; i >= 0; i--) {
    ws.spliceRows(toDelete[i], 1);
  }

  // --- Build source-row → new-row mapping (after deletion) ---
  const rowMap = new Map<number, number>();
  keptRows.forEach((srcRow, idx) => {
    rowMap.set(srcRow, 25 + idx);
  });
  for (let r = 1; r <= 24; r++) {
    rowMap.set(r, r);
  }

  // --- Update header cells ---
  // Row 1, Col A: title
  const origTitle = strVal(ws.getRow(1), COL_A) || "电商物流下单信息模板";
  ws.getCell("A1").value = `${origTitle} - ${intervalData.name} (${intervalData.totalBoxes}箱)`;

  // Row 18, Col B: total boxes
  ws.getCell("B18").value = intervalData.totalBoxes;

  // Row 19, Col B: total weight
  ws.getCell("B19").value = intervalData.totalWeight;

  // Row 20, Col B: total volume
  ws.getCell("B20").value = intervalData.totalVolume;

  // --- Add column W (每箱RMB) ---
  // Row 24 header
  const wHeader = ws.getCell("W24");
  wHeader.value = "每箱RMB";
  wHeader.font = { bold: true, size: 9 };

  // Build row → perBoxRMB map (source row numbers)
  const rowRmbMap = new Map<number, number>();
  for (const g of intervalData.groups) {
    for (const r of g.rows) {
      rowRmbMap.set(r, Math.round(g.perBoxRMB * 100) / 100);
    }
  }

  // Write column W values at remapped row positions
  for (const [srcRow, rmb] of rowRmbMap) {
    const dstRow = rowMap.get(srcRow);
    if (dstRow !== undefined) {
      const cell = ws.getCell(`W${dstRow}`);
      cell.value = rmb;
      cell.numFmt = "0.00";
      cell.font = { size: 9 };
    }
  }

  // --- Ensure column W has a reasonable width ---
  const wCol = ws.getColumn(23); // W is col 23 (1-indexed)
  if (!wCol.width || (wCol.width as number) < 10) {
    wCol.width = 12;
  }

  // --- Write output ---
  const outBuffer = await wb.xlsx.writeBuffer();
  return Buffer.from(outBuffer);
}

// ============================================================
// Main: Process uploaded file
// ============================================================

export async function processSplit(
  filePathOrBuffer: string | Uint8Array,
  fileName: string
): Promise<SplitResult> {
  // Load original workbook
  const wb = new ExcelJS.Workbook();
  if (typeof filePathOrBuffer === "string") {
    await wb.xlsx.readFile(filePathOrBuffer);
  } else {
    // @ts-expect-error — exceljs Buffer type incompatible with @types/node generic Buffer
    await wb.xlsx.load(filePathOrBuffer);
  }

  const ws = wb.getWorksheet(SOURCE_SHEET);
  if (!ws) {
    throw new Error(
      `找不到 Sheet "${SOURCE_SHEET}"。请确认上传的是易通下单模板文件。`
    );
  }

  // Phase 1: Parse and group
  const { groups } = parseAndGroup(ws);

  // Phase 2: Calculate interval data
  const intervalDataList = calcIntervalData(groups, ws);

  // Serialize the source workbook ONCE — each interval deserializes a
  // pristine clone and deletes its unwanted rows.
  console.log("🔄 Cloning workbook for interval generation…");
  const cloneBuffer = Buffer.from(await wb.xlsx.writeBuffer());

  // Phase 3: Generate output per interval from the clone buffer
  const intervals: IntervalResult[] = [];
  for (const ivData of intervalDataList) {
    let buffer = await generateIntervalFile(cloneBuffer, ivData);
    // Inject cellImages.xml + media so DISPIMG formulas resolve natively
    buffer = await injectCellImagesIntoOutput(buffer, filePathOrBuffer);
    intervals.push({
      name: ivData.name,
      fileName: ivData.fileName,
      totalBoxes: ivData.totalBoxes,
      totalWeight: ivData.totalWeight,
      totalVolume: ivData.totalVolume,
      groupCount: ivData.groupCount,
      groups: ivData.groups.map((g) => ({
        ...g,
        perBoxRMB: Math.round(g.perBoxRMB * 100) / 100,
        perBoxOrig: Math.round(g.perBoxOrig * 100) / 100,
      })),
      buffer,
    });
  }

  return {
    sourceFile: fileName,
    totalBoxes: groups.reduce((s, g) => s + g.boxes, 0),
    totalGroups: groups.length,
    intervals,
  };
}

/** Generate combined ZIP buffer from split result */
export async function generateZip(
  result: SplitResult
): Promise<Buffer> {
  const zip = new JSZip();
  for (const iv of result.intervals) {
    zip.file(iv.fileName, iv.buffer);
  }
  const summary = {
    sourceFile: result.sourceFile,
    generatedAt: new Date().toISOString(),
    totalBoxes: result.totalBoxes,
    totalGroups: result.totalGroups,
    intervals: result.intervals.map((iv) => ({
      name: iv.name,
      fileName: iv.fileName,
      totalBoxes: iv.totalBoxes,
      totalWeight: iv.totalWeight,
      totalVolume: iv.totalVolume,
      groupCount: iv.groupCount,
      groups: iv.groups,
    })),
  };
  zip.file("split_summary.json", JSON.stringify(summary, null, 2));
  return Buffer.from(await zip.generateAsync({ type: "nodebuffer" }));
}
