/**
 * TR入仓数据整理 — 客户/供应商箱规比对 → 建议「出给客户」箱规
 *
 * 数据流：
 *   客户数据（一行一个产品：FBA ID + 品名 + 总箱数 + 长宽高 + 实重）
 *   供应商数据（逐箱：货箱编号 + 货箱长/高/宽 + 货箱实重/材积重）
 *     ↓
 *   选数算法 + 差异校验报警 + 历史对比
 *     ↓
 *   建议「出给客户」箱规（前端可微调）+ 历史库累积
 *
 * 已确认规则（2026-09-02）：
 *   - 产品唯一键 = 品名 + 长 + 宽 + 高 + 实重
 *   - FBA 从供应商箱号取「U+流水号」前的部分（如 FBA19MYJ057TU000001 → FBA19MYJ057T）
 *   - 箱规匹配：长宽高向下取整后相等
 *   - 选数：有历史同款时，供应商第 1 大计费重 ≤ 历史最大 → 取第 1 大，否则退取第 2 大；新品（无历史）取第 2 大
 *   - 实重差异 ≤ 0.3kg 视为同产品（匹配容差）
 *   - 出给客户 = 供应商选数箱规（尺寸 + 该箱规最大实重 + 公式材积重），再放大最短边 +1 作安全余量（尽量不放大最大边）
 *   - 放大约束：放大后材积重 − 客户材积重 < 2；计费重不超过历史最大计费重
 *   - 差异校验（供应商原始箱规 vs 客户申报，区分材积/实重主导）：
 *      材积主导（体积重 ≥ 实重）：三边和差 ≥ 6、材积重差 ≥ 2 → 核查过机图
 *      实重主导（体积重 < 实重）：实重差 ≥ 0.5 → 核查过机图
 */

import ExcelJS from "exceljs";
import fs from "fs";
import path from "path";
import { HISTORY_SEED } from "./history-seed";

// ============================================================
// 常量
// ============================================================

/** 材积重除数：材积重 = 长 × 宽 × 高 ÷ VOLUME_DIVISOR */
const VOLUME_DIVISOR = 6000;
/** 放大比例：单行/全局不满足计费重约束时的放大倍数 */
const AMPLIFY_RATIO = 1.02;
/** 三边和差报警阈值（材积主导时，|供应商原始三边和 − 客户三边和| >= 6 报警） */
const SUM_SIDES_THRESHOLD = 6;
/** 材积重差报警阈值（|差| >= 2 报警） */
const VOLUME_DIFF_THRESHOLD = 2;
/** 供应商过大箱报警阈值（最大材积重 − 建议材积重 >= 2） */
const SUPPLIER_EXCESS_THRESHOLD = 2;
/** 实重差报警阈值（实重主导时，供应商原始实重 − 客户实重 差 >= 0.5 报警） */
const ACTUAL_DIFF_THRESHOLD = 0.5;
/** 历史库同款判定：单边尺寸差容差（cm），|Δ长|、|Δ宽|、|Δ高| ≤ 该值视为同一款 */
const HISTORY_DIM_TOLERANCE = 1;
/** 历史库同款判定：实重差容差（kg），|Δ实重| ≤ 该值视为同一款 */
const HISTORY_WEIGHT_TOLERANCE = 1;

// ============================================================
// 类型
// ============================================================

export interface CustomerRow {
  fbaId: string;
  productName: string;
  /** 客户元信息（导出列 B-F）：系统SO / 客户渠道 / 国家 / 仓库代码 / 单证报关 */
  so: string;
  channel: string;
  country: string;
  warehouse: string;
  customs: string;
  totalBoxes: number;
  lengthCm: number;
  widthCm: number;
  heightCm: number;
  actualWeight: number;
  volumeWeight: number;
  chargeableWeight: number;
  sumSides: number;
}

export interface SupplierBox {
  fbaId: string;
  lengthCm: number;
  widthCm: number;
  heightCm: number;
  actualWeight: number;
  volumeWeight: number;
  chargeableWeight: number;
}

export interface HistoryEntry {
  productName: string;
  /** 客户申报箱规（同款判定基准） */
  customerLengthCm: number;
  customerWidthCm: number;
  customerHeightCm: number;
  customerActualWeight: number;
  /** 出给客户建议（历史最大值对比基准） */
  lengthCm: number;
  widthCm: number;
  heightCm: number;
  actualWeight: number;
  volumeWeight: number;
  chargeableWeight: number;
  updatedAt: string;
}

/** 历史库：扁平数组，每条记录一份历史产品（品名 + 客户箱规标识 + 出给客户建议值） */
export type HistoryLibrary = HistoryEntry[];

/** 供应商代表箱（选数命中的那箱，导出时用于「供应商」对比列） */
export interface SupplierRepresentative {
  lengthCm: number;
  widthCm: number;
  heightCm: number;
  actualWeight: number;
  volumeWeight: number;
}

/** 单行建议（前端可编辑 suggestion 部分） */
export interface SuggestionRow {
  fbaId: string;
  productName: string;
  /** 客户元信息（导出列 B-F，随客户数据透传） */
  so: string;
  channel: string;
  country: string;
  warehouse: string;
  customs: string;
  totalBoxes: number;
  customer: {
    lengthCm: number;
    widthCm: number;
    heightCm: number;
    actualWeight: number;
    volumeWeight: number;
    chargeableWeight: number;
    sumSides: number;
  };
  /** 供应商代表箱（选数命中箱，未匹配时全 0） */
  supplier: SupplierRepresentative;
  supplierChargeable: number;
  supplierMaxVolumeWeight: number;
  suggestion: {
    lengthCm: number;
    widthCm: number;
    heightCm: number;
    actualWeight: number;
    volumeWeight: number;
    chargeableWeight: number;
    sumSides: number;
  };
  alarms: string[];
  historyMax: HistoryEntry | null;
  /** 命中的供应商箱排名（1=第1大，2=第2大…）；null 表示未自动选择或未匹配 */
  pickedRank: number | null;
}

export interface BuildResult {
  rows: SuggestionRow[];
  supplierTotal: number;
}

// ============================================================
// 工具函数
// ============================================================

/** 安全读取文本（处理 null/string/number/richText/formula） */
function cellText(cell: ExcelJS.Cell): string {
  const v = cell.value;
  if (v == null) return "";
  if (typeof v === "string") return v.trim();
  if (typeof v === "number") return String(v);
  if (typeof v === "object") {
    const o = v as {
      richText?: { text: string }[];
      text?: unknown;
      result?: unknown;
    };
    if (o.richText) return o.richText.map((r) => r.text).join("").trim();
    if (o.result !== undefined && o.result !== null) return String(o.result).trim();
    if (o.text !== undefined && o.text !== null) return String(o.text).trim();
  }
  return String(v).trim();
}

/** 安全读取数值（处理文本/公式） */
function cellNum(cell: ExcelJS.Cell): number {
  const v = cell.value;
  if (typeof v === "number") return v;
  if (typeof v === "object" && v !== null) {
    const o = v as { result?: unknown };
    if (o.result !== undefined && o.result !== null) {
      const n = Number(o.result);
      return isNaN(n) ? 0 : n;
    }
  }
  const n = parseFloat(String(v ?? "").replace(/[^\d.\-]/g, ""));
  return isNaN(n) ? 0 : n;
}

function round2(n: number): number {
  return Math.round(n * 100) / 100;
}

function sumSides(l: number, w: number, h: number): number {
  return round2(l + w + h);
}

/** 材积重 = 长×宽×高 ÷ 6000（保留 2 位小数，与供应商文件材积重一致） */
function calcVolumeWeight(l: number, w: number, h: number): number {
  return round2((l * w * h) / VOLUME_DIVISOR);
}

/** 箱规 key：长宽高向下取整 */
function makeSpecKey(b: { lengthCm: number; widthCm: number; heightCm: number }): string {
  return `${Math.floor(b.lengthCm)}_${Math.floor(b.widthCm)}_${Math.floor(b.heightCm)}`;
}

/**
 * 从供应商箱号提取 FBA：去掉末尾「U + 流水号」。
 * 例：`FBA19MYJ057TU000001` → `FBA19MYJ057T`、`FBA15M8F4YZRU000001` → `FBA15M8F4YZR`。
 * 规则按用户口径「取 U000 前面的值」，比固定取前 12 位更健壮（FBA 长度不恒为 12 也能正确切分）。
 */
function extractFba(boxNo: string): string {
  return boxNo.trim().replace(/U\d+$/, "");
}

// ============================================================
// 表头定位与列映射
// ============================================================

type FieldPatterns = Record<string, string[]>;

const CUSTOMER_PATTERNS: FieldPatterns = {
  fbaId: ["FBA ID", "FBAID", "FBA"],
  productName: ["中文品名", "品名"],
  so: ["系统SO"],
  channel: ["客户渠道", "渠道"],
  country: ["国家"],
  warehouse: ["仓库代码", "仓库"],
  customs: ["单证报关", "报关"],
  totalBoxes: ["总箱数", "箱数"],
  lengthCm: ["长(CM)", "长（CM", "长(cm)", "长（cm", "长("],
  widthCm: ["宽(CM)", "宽（CM", "宽(cm)", "宽("],
  heightCm: ["高(CM)", "高（CM", "高(cm)", "高("],
  actualWeight: ["实重"],
  volumeWeight: ["材积重"],
};

// 易通发票格式（如「8月第4周（易通发票）...xlsx」）：
// 表单头 + 明细表（中文表头，FBA 在 FBA货箱编号 列、重量在 货箱重量 列）。
// 注意不能复用 CUSTOMER_PATTERNS：此格式无「实重」列（用「货箱重量」），
// 且「品名」兜底会先命中「英文品名」列导致取错列。
const CUSTOMER_INVOICE_PATTERNS: FieldPatterns = {
  fbaId: ["FBA货箱编号"],
  productName: ["中文品名"],
  so: ["系统SO"],
  channel: ["客户渠道", "渠道"],
  country: ["国家"],
  warehouse: ["仓库代码", "仓库"],
  customs: ["单证报关", "报关"],
  totalBoxes: ["箱数/件数", "箱数"],
  lengthCm: ["长(CM)", "长（CM", "长(cm)", "长("],
  widthCm: ["宽(CM)", "宽（CM", "宽(cm)", "宽("],
  heightCm: ["高(CM)", "高（CM", "高(cm)", "高("],
  actualWeight: ["货箱重量"],
  volumeWeight: ["材积重"],
};

// 货箱清单格式（如「客户数据-给英美-0824到0830.xlsx」）：
// 表单头「货箱清单」+ 明细表（FBA ID / 中文品名 / 总箱数(CTN) / 长宽高 / 单箱货物毛重）。
// 注意：此格式「货箱重量」列是该 SO 的总重（如 2442.45），不是单箱重；
// 单箱实重 = 「单箱货物毛重」列。故必须精确匹配「单箱货物毛重」，不能复用「货箱重量」。
const CUSTOMER_CARGO_PATTERNS: FieldPatterns = {
  fbaId: ["FBA ID", "FBAID", "FBA"],
  // 注意不能加「品名」兜底：此格式同时有「英文品名」「中文品名」两列，
  // 「品名」会先命中「英文品名」导致取错列（与易通发票格式同坑）
  productName: ["中文品名"],
  so: ["系统SO"],
  channel: ["客户渠道", "渠道"],
  country: ["国家"],
  warehouse: ["仓库代码", "仓库"],
  customs: ["单证报关", "报关"],
  totalBoxes: ["总箱数(CTN)", "总箱数", "箱数"],
  lengthCm: ["长(CM)", "长（CM", "长(cm)", "长("],
  widthCm: ["宽(CM)", "宽（CM", "宽(cm)", "宽("],
  heightCm: ["高(CM)", "高（CM", "高(cm)", "高("],
  actualWeight: ["单箱货物毛重"],
  volumeWeight: ["材积重"],
};

const SUPPLIER_PATTERNS: FieldPatterns = {
  // 注意：必须用「货箱编号」精确匹配，不能用「系统箱号/箱号」兜底，
  // 否则会误匹配到先出现的「系统箱号」列（UKSZ...U0001），导致 FBA 提取错误
  fbaId: ["货箱编号"],
  lengthCm: ["货箱长", "长(CM)", "长("],
  widthCm: ["货箱宽", "宽(CM)", "宽("],
  heightCm: ["货箱高", "高(CM)", "高("],
  actualWeight: ["货箱实重", "实重"],
  volumeWeight: ["货箱材积重", "材积重"],
};

// 英美入仓格式（如 TRKJ26080105-英美入仓数据.xlsx）：
// FBA 在「扩展箱号」列（FBA19MYJ057TU000001 → 取「U000」前的 FBA19MYJ057T），
// 箱规在 货箱重量(BI)/货箱长度(BJ)/货箱宽度(BK)/货箱高度(BL)/货箱材积重(BM)。
// 注意不能复用 SUPPLIER_PATTERNS：此格式的「货箱编号」是运单号（10593316U001）而非 FBA。
const SUPPLIER_ENTRY_PATTERNS: FieldPatterns = {
  fbaId: ["扩展箱号"],
  lengthCm: ["货箱长度"],
  widthCm: ["货箱宽度"],
  heightCm: ["货箱高度"],
  actualWeight: ["货箱重量"],
  volumeWeight: ["货箱材积重"],
};

// 给总部格式（如「供应商数据-给总部-0824到0830 更新.xlsx」）：
// 每个 SO 一个块（块头 4 行 + 表头 + 逐箱数据 + TOTAL），FBA 在「FBA号」列（FBA19MTJH5NPU000014 → 取「U+流水号」前）。
// 箱规在 长(CM)/宽(CM)/高(CM)，单箱实重=「单件重量（KGS)」、材积重=「单件材积(KGS)」。
// 注意不能复用 SUPPLIER_PATTERNS/SUPPLIER_ENTRY_PATTERNS：此格式 FBA 在「FBA号」列、重量在「单件重量」列。
const SUPPLIER_HEADQUARTERS_PATTERNS: FieldPatterns = {
  fbaId: ["FBA号"],
  lengthCm: ["长(CM)", "长（CM", "长(cm)", "长("],
  widthCm: ["宽(CM)", "宽（CM", "宽(cm)", "宽("],
  heightCm: ["高(CM)", "高（CM", "高(cm)", "高("],
  actualWeight: ["单件重量"],
  volumeWeight: ["单件材积"],
};

/** 在表头区（前 maxRow 行）定位包含所有关键词的表头行 */
function findHeaderRow(ws: ExcelJS.Worksheet, required: string[], maxRow = 40): number | null {
  for (let r = 1; r <= maxRow; r++) {
    const texts: string[] = [];
    ws.getRow(r).eachCell((cell) => texts.push(cellText(cell)));
    const joined = texts.join(" ");
    if (required.every((k) => joined.includes(k))) return r;
  }
  return null;
}

/** 依据表头行与字段关键词，建立「字段 → 列号」映射（首个命中优先） */
function buildColumnMap(
  headerRow: ExcelJS.Row,
  patterns: FieldPatterns,
): Record<string, number> {
  const map: Record<string, number> = {};
  headerRow.eachCell((cell, col) => {
    const t = cellText(cell);
    if (!t) return;
    for (const [field, pats] of Object.entries(patterns)) {
      if (map[field] !== undefined) continue;
      if (pats.some((p) => t.includes(p))) {
        map[field] = col;
        break;
      }
    }
  });
  return map;
}

function readRowText(row: ExcelJS.Row, col: number | undefined): string {
  if (col === undefined) return "";
  return cellText(row.getCell(col));
}

function readRowNum(row: ExcelJS.Row, col: number | undefined): number {
  if (col === undefined) return 0;
  return cellNum(row.getCell(col));
}

// ============================================================
// 解析客户 / 供应商文件
// ============================================================

/** 解析客户数据（一行一个产品） */
export async function parseCustomerFile(filePath: string): Promise<CustomerRow[]> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(filePath);
  const ws = wb.worksheets[0];
  if (!ws) return [];

  const headerRow = findHeaderRow(ws, ["FBA", "品名", "箱数"]);
  if (headerRow === null) {
    throw new Error("客户文件未识别到表头（需包含 FBA / 品名 / 箱数 列）");
  }

  // 格式检测：表头含「FBA货箱编号」→ 易通发票格式；含「单箱货物毛重」→ 货箱清单格式；否则标准客户格式
  const headerTexts: string[] = [];
  ws.getRow(headerRow).eachCell((cell) => headerTexts.push(cellText(cell)));
  const isInvoiceFormat = headerTexts.some((t) => t.includes("FBA货箱编号"));
  const isCargoListFormat = headerTexts.some((t) => t.includes("单箱货物毛重"));

  const map = buildColumnMap(
    ws.getRow(headerRow),
    isInvoiceFormat
      ? CUSTOMER_INVOICE_PATTERNS
      : isCargoListFormat
        ? CUSTOMER_CARGO_PATTERNS
        : CUSTOMER_PATTERNS
  );

  const rows: CustomerRow[] = [];
  ws.eachRow((row, r) => {
    if (r <= headerRow) return;
    const fbaId = readRowText(row, map.fbaId);
    const productName = readRowText(row, map.productName);
    if (!productName && !fbaId) return; // 跳过尾部汇总/空行
    if (!productName) return;

    const so = readRowText(row, map.so);
    const channel = readRowText(row, map.channel);
    const country = readRowText(row, map.country);
    const warehouse = readRowText(row, map.warehouse);
    const customs = readRowText(row, map.customs);
    const totalBoxes = Math.round(readRowNum(row, map.totalBoxes)) || 0;
    const lengthCm = readRowNum(row, map.lengthCm);
    const widthCm = readRowNum(row, map.widthCm);
    const heightCm = readRowNum(row, map.heightCm);
    const actualWeight = readRowNum(row, map.actualWeight);

    // 材积重：文件已含则直接用，否则按公式计算
    let volumeWeight = readRowNum(row, map.volumeWeight);
    if (volumeWeight <= 0) volumeWeight = calcVolumeWeight(lengthCm, widthCm, heightCm);

    const chargeableWeight = Math.max(actualWeight, volumeWeight);

    rows.push({
      fbaId,
      productName,
      so,
      channel,
      country,
      warehouse,
      customs,
      totalBoxes,
      lengthCm,
      widthCm,
      heightCm,
      actualWeight,
      volumeWeight,
      chargeableWeight,
      sumSides: sumSides(lengthCm, widthCm, heightCm),
    });
  });

  return rows;
}

/** 解析供应商逐箱数据 */
export async function parseSupplierFile(filePath: string): Promise<SupplierBox[]> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(filePath);
  const ws = wb.worksheets[0];
  if (!ws) return [];

  const headerRow =
    findHeaderRow(ws, ["货箱编号", "货箱长"]) ??
    findHeaderRow(ws, ["FBA号", "单件重量"]);
  if (headerRow === null) {
    throw new Error("供应商文件未识别到表头（需包含 货箱编号/货箱长 或 FBA号/单件重量 列）");
  }

  // 格式检测：表头含「扩展箱号」→ 英美入仓格式；含「单件重量」→ 给总部格式；否则天图格式
  const headerTexts: string[] = [];
  ws.getRow(headerRow).eachCell((cell) => headerTexts.push(cellText(cell)));
  const isEntryFormat = headerTexts.some((t) => t.includes("扩展箱号"));
  const isHeadquartersFormat = headerTexts.some((t) => t.includes("单件重量"));

  const map = buildColumnMap(
    ws.getRow(headerRow),
    isEntryFormat
      ? SUPPLIER_ENTRY_PATTERNS
      : isHeadquartersFormat
        ? SUPPLIER_HEADQUARTERS_PATTERNS
        : SUPPLIER_PATTERNS
  );

  const boxes: SupplierBox[] = [];
  ws.eachRow((row, r) => {
    if (r <= headerRow) return;
    const boxNo = readRowText(row, map.fbaId);
    if (!boxNo) return; // 箱号为空即停
    const fbaId = extractFba(boxNo); // 取「U+流水号」前的 FBA（如 FBA19MYJ057TU000001 → FBA19MYJ057T）
    if (!/^FBA[A-Z0-9]+$/i.test(fbaId)) return; // 跳过非数据行（给总部格式的块头/TOTAL/表头等）

    const lengthCm = readRowNum(row, map.lengthCm);
    const widthCm = readRowNum(row, map.widthCm);
    const heightCm = readRowNum(row, map.heightCm);
    const actualWeight = readRowNum(row, map.actualWeight);

    let volumeWeight = readRowNum(row, map.volumeWeight);
    if (volumeWeight <= 0) volumeWeight = calcVolumeWeight(lengthCm, widthCm, heightCm);

    boxes.push({
      fbaId,
      lengthCm,
      widthCm,
      heightCm,
      actualWeight,
      volumeWeight,
      chargeableWeight: Math.max(actualWeight, volumeWeight),
    });
  });

  return boxes;
}

// ============================================================
// 选数算法
// ============================================================

/**
 * 选数：根据历史参考值决定取计费重第 1 大还是第 2 大那箱。
 *   - 新品（无历史同款）：取第 2 大（避免取到偶发偏大的异常箱）
 *   - 有历史：第 1 大计费重 ≤ 历史最大计费重 → 取第 1 大；超过历史最大 → 退取第 2 大
 * 仅 1 箱时回退到第 1 大（sorted 非空，调用方已保证）。
 */
function selectBox(sorted: SupplierBox[], historyMaxChargeable: number | null): SupplierBox {
  const first = sorted[0];
  const second = sorted[1] ?? first;
  if (historyMaxChargeable == null) return second;
  if (first.chargeableWeight > historyMaxChargeable) return second;
  return first;
}

/**
 * 放大出给客户箱规：尽量不放大最大边，取最短边 +1 作安全余量。
 * 约束（任一不满足则不放大，返回原尺寸）：
 *   1. 放大后材积重 − 客户材积重 必须 < VOLUME_DIFF_THRESHOLD（2）
 *   2. 放大后计费重 不超过历史最大计费重（历史最大是上限）
 */
function amplifyDims(
  l: number,
  w: number,
  h: number,
  actualWeight: number,
  customerVolumeWeight: number,
  historyMaxChargeable: number | null,
): { lengthCm: number; widthCm: number; heightCm: number } {
  let nl = l;
  let nw = w;
  let nh = h;
  // 最短边 +1（相等时取先比较到的那条，保持确定性）
  if (l <= w && l <= h) nl += 1;
  else if (w <= h) nw += 1;
  else nh += 1;

  const newVol = calcVolumeWeight(nl, nw, nh);
  // 约束 1：材积重差必须 < 2
  if (newVol - customerVolumeWeight >= VOLUME_DIFF_THRESHOLD) {
    return { lengthCm: l, widthCm: w, heightCm: h };
  }
  // 约束 2：计费重不超过历史最大
  const newChargeable = Math.max(actualWeight, newVol);
  if (historyMaxChargeable != null && newChargeable > historyMaxChargeable) {
    return { lengthCm: l, widthCm: w, heightCm: h };
  }
  return { lengthCm: nl, widthCm: nw, heightCm: nh };
}

/**
 * 放大建议值，使材积重/计费重超过 target。
 * 材积主导（材积重 ≥ 实重）时按单边整数放大（每次把最短边 +1）；实重主导时直接放大实重。
 */
function forceAmplify(
  s: SuggestionRow["suggestion"],
  target: number,
): void {
  if (s.volumeWeight >= s.actualWeight) {
    // 材积主导：单边整数放大（每次把最短边 +1，直到材积重 > target）
    let L = s.lengthCm;
    let W = s.widthCm;
    let H = s.heightCm;

    let guard = 0;
    while (calcVolumeWeight(L, W, H) <= target && guard < 500) {
      if (L <= W && L <= H) L += 1;
      else if (W <= H) W += 1;
      else H += 1;
      guard++;
    }

    s.lengthCm = L;
    s.widthCm = W;
    s.heightCm = H;
    s.volumeWeight = calcVolumeWeight(L, W, H);
    s.chargeableWeight = Math.max(s.actualWeight, s.volumeWeight);
  } else {
    // 实重主导：直接放大实重
    s.actualWeight = round2(target);
    s.chargeableWeight = Math.max(s.actualWeight, s.volumeWeight);
  }
  s.sumSides = sumSides(s.lengthCm, s.widthCm, s.heightCm);
}

// ============================================================
// 建议生成
// ============================================================

export function buildSuggestions(
  customers: CustomerRow[],
  boxes: SupplierBox[],
  history: HistoryLibrary,
): BuildResult {
  // 供应商箱按 FBA 分组
  const boxGroups = new Map<string, SupplierBox[]>();
  for (const b of boxes) {
    const arr = boxGroups.get(b.fbaId) ?? [];
    arr.push(b);
    boxGroups.set(b.fbaId, arr);
  }

  const supplierTotal = boxes.reduce((s, b) => s + b.chargeableWeight, 0);

  const rows: SuggestionRow[] = [];

  for (const c of customers) {
    const group = boxGroups.get(c.fbaId);
    const customer = {
      lengthCm: c.lengthCm,
      widthCm: c.widthCm,
      heightCm: c.heightCm,
      actualWeight: c.actualWeight,
      volumeWeight: c.volumeWeight,
      chargeableWeight: c.chargeableWeight,
      sumSides: c.sumSides,
    };

    const alarms: string[] = [];

    // 历史库同款匹配：品名一致 且 客户箱规/实重相近（单边 ≤1cm、实重 ≤1kg）
    const historyMatches = history.filter((e) => isSameHistoryProduct(c, historyIdentity(e)));
    const historyMax = historyMatches.reduce<HistoryEntry | null>(
      (best, e) => (best === null || e.chargeableWeight > best.chargeableWeight ? e : best),
      null,
    );
    // 未匹配到供应商箱
    if (!group || group.length === 0) {
      rows.push({
        fbaId: c.fbaId,
        productName: c.productName,
        so: c.so,
        channel: c.channel,
        country: c.country,
        warehouse: c.warehouse,
        customs: c.customs,
        totalBoxes: c.totalBoxes,
        customer,
        supplier: { lengthCm: 0, widthCm: 0, heightCm: 0, actualWeight: 0, volumeWeight: 0 },
        supplierChargeable: 0,
        supplierMaxVolumeWeight: 0,
        suggestion: { ...customer },
        alarms: ["⚠需人工复核"],
        historyMax,
        pickedRank: null,
      });
      continue;
    }

    const sorted = [...group].sort((a, b) => b.chargeableWeight - a.chargeableWeight);
    const supplierChargeable = sorted[0].chargeableWeight;
    const supplierMaxVolumeWeight = Math.max(...group.map((b) => b.volumeWeight));

    // 选数（有历史按历史最大决定取第 1/第 2 大；新品取第 2 大）
    const picked = selectBox(sorted, historyMax ? historyMax.chargeableWeight : null);

    // 建议值 = 选中箱规的长宽高 + 该箱规所有箱的最大实重 + 公式材积重
    const specKey = makeSpecKey(picked);
    const specBoxes = group.filter((b) => makeSpecKey(b) === specKey);
    const maxActual = Math.max(...specBoxes.map((b) => b.actualWeight));
    const volumeWeight = calcVolumeWeight(picked.lengthCm, picked.widthCm, picked.heightCm);

    // 供应商代表箱 = 选数命中的箱规（未放大的原始值）
    const supplier: SupplierRepresentative = {
      lengthCm: picked.lengthCm,
      widthCm: picked.widthCm,
      heightCm: picked.heightCm,
      actualWeight: maxActual,
      volumeWeight,
    };

    // 出给客户 = 供应商选数箱规 + 放大最短边 +1 作安全余量（尽量不放大最大边）。
    // 约束：放大后材积重 − 客户材积重 < 2、计费重不超过历史最大；不满足则不放大。
    const amp = amplifyDims(
      picked.lengthCm,
      picked.widthCm,
      picked.heightCm,
      maxActual,
      c.volumeWeight,
      historyMax ? historyMax.chargeableWeight : null,
    );
    const ampVolumeWeight = calcVolumeWeight(amp.lengthCm, amp.widthCm, amp.heightCm);

    const suggestion: SuggestionRow["suggestion"] = {
      lengthCm: amp.lengthCm,
      widthCm: amp.widthCm,
      heightCm: amp.heightCm,
      actualWeight: maxActual,
      volumeWeight: ampVolumeWeight,
      chargeableWeight: Math.max(maxActual, ampVolumeWeight),
      sumSides: sumSides(amp.lengthCm, amp.widthCm, amp.heightCm),
    };
    const pickedRank: number | null = sorted.indexOf(picked) + 1;

    // 差异校验（供应商原始箱规 vs 客户申报，区分材积主导 / 实重主导），超限报警提示核查过机图。
    const supplierSumSides = sumSides(supplier.lengthCm, supplier.widthCm, supplier.heightCm);
    if (supplier.volumeWeight >= supplier.actualWeight) {
      // 材积主导（体积重 ≥ 实际重量）：查三边和差、材积重差
      if (Math.abs(supplierSumSides - c.sumSides) >= SUM_SIDES_THRESHOLD) {
        alarms.push("三边和差异超限，请核查过机图");
      }
      if (Math.abs(supplier.volumeWeight - c.volumeWeight) >= VOLUME_DIFF_THRESHOLD) {
        alarms.push("材积重差异超限，请核查过机图");
      }
    } else {
      // 实重主导（体积重 < 实际重量）：查实重差
      if (Math.abs(supplier.actualWeight - c.actualWeight) >= ACTUAL_DIFF_THRESHOLD) {
        alarms.push("实重差异超限，请核查过机图");
      }
    }
    // 报警：供应商过大箱（要求找供应商核查过机图，核实后再修改）
    if (supplierMaxVolumeWeight - suggestion.volumeWeight >= SUPPLIER_EXCESS_THRESHOLD) {
      alarms.push("供应商存在过大箱，请核查过机图");
    }

    // 供应商小于客户：单独提示，并取历史最大值（若有）作为出给客户建议，
    // 避免出给客户比客户申报还小（供应商装箱偏小 → 参考历史合理值）。
    // 用供应商真实最大计费重（第 1 大箱 supplierChargeable）判断，与前端「供应商计费重」列一致；
    // 不能用选数命中箱规（退选第 2 大时其计费重可能 < 客户，会误报）。
    if (supplierChargeable < c.chargeableWeight) {
      alarms.push("供应商计费重小于客户，请确认");
      if (historyMax) {
        suggestion.lengthCm = historyMax.lengthCm;
        suggestion.widthCm = historyMax.widthCm;
        suggestion.heightCm = historyMax.heightCm;
        suggestion.actualWeight = historyMax.actualWeight;
        suggestion.volumeWeight = historyMax.volumeWeight;
        suggestion.chargeableWeight = historyMax.chargeableWeight;
        suggestion.sumSides = sumSides(historyMax.lengthCm, historyMax.widthCm, historyMax.heightCm);
      }
    }

    // 历史对比：历史最大值 > 建议值 → 提示可参考历史最大值放大
    if (historyMax && historyMax.chargeableWeight > suggestion.chargeableWeight) {
      alarms.push("建议参考历史最大值放大");
    }

    rows.push({
      fbaId: c.fbaId,
      productName: c.productName,
      so: c.so,
      channel: c.channel,
      country: c.country,
      warehouse: c.warehouse,
      customs: c.customs,
      totalBoxes: c.totalBoxes,
      customer,
      supplier,
      supplierChargeable,
      supplierMaxVolumeWeight,
      suggestion,
      alarms,
      historyMax,
      pickedRank,
    });
  }

  return { rows, supplierTotal };
}

// ============================================================
// 历史库
// ============================================================

/** 客户产品标识（同款判定用：品名 + 客户申报长宽高 + 实重） */
interface CustomerIdentity {
  productName: string;
  lengthCm: number;
  widthCm: number;
  heightCm: number;
  actualWeight: number;
}

/** 提取历史记录对应的客户标识 */
function historyIdentity(e: HistoryEntry): CustomerIdentity {
  return {
    productName: e.productName,
    lengthCm: e.customerLengthCm,
    widthCm: e.customerWidthCm,
    heightCm: e.customerHeightCm,
    actualWeight: e.customerActualWeight,
  };
}

/**
 * 判断两份「客户标识」是否同一款。
 * 同款标准（2026-09-02 口径）：品名一致 且 长宽高排序后逐边差 ≤ 1cm 且 实重差 ≤ 1kg。
 * 长宽高先升序排序再逐边比，避免同一款箱因长宽高书写顺序不同
 * （如 `39.6×39.6×43.8` 与 `43.8×39.6×39.6`）被判成不同款。
 */
function isSameHistoryProduct(a: CustomerIdentity, b: CustomerIdentity): boolean {
  if (a.productName !== b.productName) return false;
  if (Math.abs(a.actualWeight - b.actualWeight) > HISTORY_WEIGHT_TOLERANCE) return false;
  const sa = [a.lengthCm, a.widthCm, a.heightCm].sort((x, y) => x - y);
  const sb = [b.lengthCm, b.widthCm, b.heightCm].sort((x, y) => x - y);
  return (
    Math.abs(sa[0] - sb[0]) <= HISTORY_DIM_TOLERANCE &&
    Math.abs(sa[1] - sb[1]) <= HISTORY_DIM_TOLERANCE &&
    Math.abs(sa[2] - sb[2]) <= HISTORY_DIM_TOLERANCE
  );
}

/**
 * 将一条历史记录并入历史库。
 * 默认同款取计费重更大者（自动累积用，避免建议值把历史最大值压小）；
 * 传入 `overwrite` 时，同款直接覆盖（手动导入「最终出给客户」数据用，以我们最终提供的值为准）。
 */
export function upsertHistoryEntry(
  lib: HistoryLibrary,
  entry: HistoryEntry,
  opts?: { overwrite?: boolean }
): void {
  const id = historyIdentity(entry);
  const idx = lib.findIndex((e) => isSameHistoryProduct(id, historyIdentity(e)));
  if (idx === -1) {
    lib.push(entry);
  } else if (opts?.overwrite || lib[idx].chargeableWeight < entry.chargeableWeight) {
    lib[idx] = entry;
  }
}

const HISTORY_DIR = () => path.join(process.cwd(), "data");
const HISTORY_FILE = () => path.join(HISTORY_DIR(), "history.json");

/** 读取历史库（文件不存在时返回内置种子数据，线上 K8s 无持久卷时兜底） */
export function loadHistory(): HistoryLibrary {
  try {
    const p = HISTORY_FILE();
    if (!fs.existsSync(p)) return HISTORY_SEED.map((e) => ({ ...e }));
    const raw = fs.readFileSync(p, "utf-8");
    const parsed = JSON.parse(raw);
    if (Array.isArray(parsed)) return parsed as HistoryLibrary;
    // 兼容旧格式 { 品名: {...} }：转为扁平数组，客户标识回填为出给客户值
    if (parsed && typeof parsed === "object") {
      return Object.entries(parsed).map(([name, v]) => {
        const e = v as Omit<HistoryEntry, "productName" | "customerLengthCm" | "customerWidthCm" | "customerHeightCm" | "customerActualWeight">;
        return {
          productName: name,
          customerLengthCm: e.lengthCm,
          customerWidthCm: e.widthCm,
          customerHeightCm: e.heightCm,
          customerActualWeight: e.actualWeight,
          lengthCm: e.lengthCm,
          widthCm: e.widthCm,
          heightCm: e.heightCm,
          actualWeight: e.actualWeight,
          volumeWeight: e.volumeWeight,
          chargeableWeight: e.chargeableWeight,
          updatedAt: e.updatedAt,
        };
      });
    }
    return [];
  } catch (e) {
    console.error("loadHistory failed:", e);
    return [];
  }
}

/** 写入历史库 */
export function saveHistory(lib: HistoryLibrary): void {
  try {
    fs.mkdirSync(HISTORY_DIR(), { recursive: true });
    fs.writeFileSync(HISTORY_FILE(), JSON.stringify(lib, null, 2), "utf-8");
  } catch (e) {
    console.error("saveHistory failed:", e);
  }
}

/** 每次导出后累积历史库（同款取计费重更大者） */
export function accumulateHistory(rows: SuggestionRow[]): HistoryLibrary {
  const lib = loadHistory();
  const now = new Date().toISOString();
  for (const r of rows) {
    if (!r.productName) continue;
    const entry: HistoryEntry = {
      productName: r.productName,
      customerLengthCm: r.customer.lengthCm,
      customerWidthCm: r.customer.widthCm,
      customerHeightCm: r.customer.heightCm,
      customerActualWeight: r.customer.actualWeight,
      lengthCm: r.suggestion.lengthCm,
      widthCm: r.suggestion.widthCm,
      heightCm: r.suggestion.heightCm,
      actualWeight: r.suggestion.actualWeight,
      volumeWeight: r.suggestion.volumeWeight,
      chargeableWeight: r.suggestion.chargeableWeight,
      updatedAt: now,
    };
    upsertHistoryEntry(lib, entry);
  }
  saveHistory(lib);
  return lib;
}

/**
 * 从历史 Excel 提取历史产品。支持两种格式（按列自动识别）：
 * 1. 双组格式（参考文件《拓锐入仓数据参考》等）：含「客户的 / 出给客户」两组列，
 *    客户组（长宽高实重）用于同款判定，出给客户组（长宽高实重材积重）用于历史最大值对比；
 * 2. 单组格式（《…出给客户的.xlsx》最终数据）：仅一组「长/宽/高/实重/材积重」列，
 *    该组即「最终提供给客户」的值，同款判定与历史最大值都用同一组（客户标识回填为最终值）。
 */
export async function importHistoryFromExcel(filePath: string): Promise<HistoryLibrary> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(filePath);
  const ws = wb.worksheets[0];
  if (!ws) return [];

  const headerRow = findHeaderRow(ws, ["品名"]);
  if (headerRow === null) {
    throw new Error("历史文件未识别到表头（需包含 品名 列）");
  }

  // 定位「客户组 / 出给客户组」列：优先按表头「客户X」「出给客户X」前缀，
  // 否则（参考文件两列同名「长(CM)」）按位置——客户取首次、出给客户取末次。
  const pickGroupCols = (keyword: string): { customer?: number; suggestion?: number } => {
    const hits: { col: number; text: string }[] = [];
    ws.getRow(headerRow).eachCell((cell, col) => {
      const t = cellText(cell);
      // 排除「总」开头的汇总列（总实重/总材积重/总计费重等），只取单箱值
      if (t && t.includes(keyword) && !t.startsWith("总")) hits.push({ col, text: t });
    });
    const cust = hits.find((h) => h.text.includes("客户") && !h.text.includes("出给客户"));
    const sugg = hits.find((h) => h.text.includes("出给客户"));
    return {
      customer: cust?.col ?? hits[0]?.col,
      suggestion: sugg?.col ?? hits[hits.length - 1]?.col,
    };
  };

  const productNameCol = (() => {
    const hits: number[] = [];
    ws.getRow(headerRow).eachCell((cell, col) => {
      if (cellText(cell).includes("品名")) hits.push(col);
    });
    return hits[0];
  })();

  const cLen = pickGroupCols("长");
  const cWid = pickGroupCols("宽");
  const cHei = pickGroupCols("高");
  const cAct = pickGroupCols("实重");
  const sVol = pickGroupCols("材积重");

  const lib: HistoryLibrary = [];
  const now = new Date().toISOString();
  ws.eachRow((row, r) => {
    if (r <= headerRow) return;
    const name = readRowText(row, productNameCol);
    if (!name) return;

    const customerLengthCm = readRowNum(row, cLen.customer);
    const customerWidthCm = readRowNum(row, cWid.customer);
    const customerHeightCm = readRowNum(row, cHei.customer);
    const customerActualWeight = readRowNum(row, cAct.customer);
    const lengthCm = readRowNum(row, cLen.suggestion);
    const widthCm = readRowNum(row, cWid.suggestion);
    const heightCm = readRowNum(row, cHei.suggestion);
    const actualWeight = readRowNum(row, cAct.suggestion);
    let volumeWeight = readRowNum(row, sVol.suggestion);
    if (volumeWeight <= 0) volumeWeight = calcVolumeWeight(lengthCm, widthCm, heightCm);

    // 客户组与出给客户组均无有效数据时跳过（尾部空行）
    const hasCustomer =
      customerLengthCm > 0 || customerWidthCm > 0 || customerHeightCm > 0 || customerActualWeight > 0;
    const hasSuggestion = lengthCm > 0 || widthCm > 0 || heightCm > 0 || actualWeight > 0;
    if (!hasCustomer && !hasSuggestion) return;

    const chargeableWeight = Math.max(actualWeight, volumeWeight);
    const entry: HistoryEntry = {
      productName: name,
      customerLengthCm,
      customerWidthCm,
      customerHeightCm,
      customerActualWeight,
      lengthCm,
      widthCm,
      heightCm,
      actualWeight,
      volumeWeight,
      chargeableWeight,
      updatedAt: now,
    };
    upsertHistoryEntry(lib, entry);
  });

  return lib;
}

/** 导出历史库为 Excel（备份用，含客户组与出给客户组，便于回导） */
export async function exportHistoryBuffer(history: HistoryLibrary): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("历史库");

  ws.columns = [
    { header: "中文品名", key: "name", width: 24 },
    { header: "客户长(CM)", key: "cLen", width: 10 },
    { header: "客户宽(CM)", key: "cWid", width: 10 },
    { header: "客户高(CM)", key: "cHei", width: 10 },
    { header: "客户实重", key: "cAct", width: 10 },
    { header: "出给客户长(CM)", key: "lengthCm", width: 12 },
    { header: "出给客户宽(CM)", key: "widthCm", width: 12 },
    { header: "出给客户高(CM)", key: "heightCm", width: 12 },
    { header: "出给客户实重", key: "actualWeight", width: 12 },
    { header: "出给客户材积重", key: "volumeWeight", width: 12 },
    { header: "出给客户计费重", key: "chargeableWeight", width: 12 },
    { header: "更新时间", key: "updatedAt", width: 24 },
  ];

  for (const e of history) {
    ws.addRow({
      name: e.productName,
      cLen: e.customerLengthCm,
      cWid: e.customerWidthCm,
      cHei: e.customerHeightCm,
      cAct: e.customerActualWeight,
      lengthCm: e.lengthCm,
      widthCm: e.widthCm,
      heightCm: e.heightCm,
      actualWeight: e.actualWeight,
      volumeWeight: e.volumeWeight,
      chargeableWeight: e.chargeableWeight,
      updatedAt: e.updatedAt,
    });
  }

  return Buffer.from(await wb.xlsx.writeBuffer());
}

// ============================================================
// 导出「出给客户」Excel
// ============================================================

/**
 * 导出建议结果为 Excel（导出前执行全局兜底校验）。
 * 全局兜底：Σ出给客户总计费重 > Σ供应商总计费重；否则等比例放大并标注 [全局调整]。
 *
 * 导出格式对齐参考文件《拓锐…入仓数据（成本）.xlsx》：表头两行（分组 + 列名），
 * 三组并排对比（客户的 / 供应商 / 出给客户），派生列（材积重/总重/计费重/差异）用
 * Excel 公式书写，方便人工调整尺寸后自动重算。
 */
export async function exportOutputBuffer(
  rows: SuggestionRow[],
  supplierTotal: number,
): Promise<Buffer> {
  // 全局兜底校验
  const customerTotal = rows.reduce(
    (s, r) => s + r.suggestion.chargeableWeight * r.totalBoxes,
    0,
  );
  if (rows.length > 0 && customerTotal > 0 && customerTotal <= supplierTotal) {
    const ratio = (supplierTotal / customerTotal) * AMPLIFY_RATIO;
    for (const r of rows) {
      forceAmplify(r.suggestion, r.suggestion.chargeableWeight * ratio);
      r.alarms.push("[全局调整]");
    }
  }

  // 列号常量（1 基，与参考文件列位一一对应）
  const C = {
    shipDate: 1, // A 出货日期
    so: 2, // B 系统SO
    channel: 3, // C 客户渠道
    country: 4, // D 国家
    warehouse: 5, // E 仓库代码
    customs: 6, // F 单证报关
    fbaId: 7, // G FBA ID
    productName: 8, // H 中文品名
    totalBoxes: 9, // I 总箱数
    cLen: 10, // J 客户长
    cWid: 11, // K 客户宽
    cHei: 12, // L 客户高
    cAct: 13, // M 客户实重
    cVol: 14, // N 客户材积重(公式)
    diff1: 16, // P 差异(客户材积重-供应商材积重)
    sLen: 17, // Q 供应商长
    sWid: 18, // R 供应商宽
    sHei: 19, // S 供应商高
    sAct: 20, // T 供应商实重
    sVol: 21, // U 供应商材积重(公式)
    sTotAct: 22, // V 供应商总实重(公式)
    sTotVol: 23, // W 供应商总材积重(公式)
    diff2: 25, // Y 差异(客户材积重-出给客户材积重)
    rLen: 26, // Z 出给客户长
    rWid: 27, // AA 出给客户宽
    rHei: 28, // AB 出给客户高
    rAct: 29, // AC 出给客户实重
    rVol: 30, // AD 出给客户材积重(公式)
    rTotAct: 31, // AE 出给客户总实重(公式)
    rTotVol: 32, // AF 出给客户总材积重(公式)
    rCharge: 33, // AG 出给客户计费重(公式)
    rTotCharge: 34, // AH 出给客户总计费重(公式)
    costKg: 35, // AI 成本KG
    remark: 36, // AJ 备注（无表头）
    channel2: 37, // AK 渠道
    boxes: 38, // AL 箱数
    totCharge: 39, // AM 总计费重
    totCost: 40, // AN 总成本重
  };

  const colLetter = (n: number): string => {
    let s = "";
    while (n > 0) {
      const m = (n - 1) % 26;
      s = String.fromCharCode(65 + m) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  };

  // 样式常量（对齐参考文件：宋体 11、细黑边框、居中）
  const FONT = { name: "宋体", size: 11, color: { argb: "FF000000" } };
  const FONT_BOLD = { name: "宋体", size: 11, bold: true, color: { argb: "FF000000" } };
  const CENTER = { horizontal: "center", vertical: "middle" } as const;
  const THIN_BORDER = {
    top: { style: "thin" as const, color: { argb: "FF000000" } },
    left: { style: "thin" as const, color: { argb: "FF000000" } },
    bottom: { style: "thin" as const, color: { argb: "FF000000" } },
    right: { style: "thin" as const, color: { argb: "FF000000" } },
  };
  const FILL = (argb: string) => ({ type: "pattern" as const, pattern: "solid" as const, fgColor: { argb } });
  const FMT_TWO_DEC = "0.00_);[Red](0.00)";

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("出给客户");
  ws.views = [{ state: "normal", showGridLines: false }];

  // ---- 第 1 行：分组表头（三色区分三组，宋体 11 不加粗）----
  const groupRow = ws.getRow(1);
  const setGroup = (start: number, end: number, label: string, fill: string) => {
    ws.mergeCells(`${colLetter(start)}1:${colLetter(end)}1`);
    for (let c = start; c <= end; c++) {
      const cell = groupRow.getCell(c);
      cell.value = c === start ? label : null;
      cell.font = FONT;
      cell.fill = FILL(fill);
      cell.alignment = CENTER;
    }
  };
  setGroup(C.cLen, C.cVol, "客户的", "FFDEEBF7");
  setGroup(C.sLen, C.sTotVol, "供应商", "FFFBE5D6");
  setGroup(C.rLen, C.rVol, "出给客户", "FFE2F0D9");

  // ---- 第 2 行：列名表头 ----
  const headerLabels: [number, string][] = [
    [C.shipDate, "出货日期"],
    [C.so, "系统SO"],
    [C.channel, "客户渠道"],
    [C.country, "国家"],
    [C.warehouse, "仓库代码"],
    [C.customs, "单证报关"],
    [C.fbaId, "FBA ID"],
    [C.productName, "中文品名"],
    [C.totalBoxes, "总箱数"],
    [C.cLen, "长(CM)"],
    [C.cWid, "宽(CM)"],
    [C.cHei, "高(CM)"],
    [C.cAct, "实重"],
    [C.cVol, "材积重"],
    [C.diff1, "差异"],
    [C.sLen, "长(CM)"],
    [C.sWid, "宽(CM)"],
    [C.sHei, "高(CM)"],
    [C.sAct, "实重"],
    [C.sVol, "材积重"],
    [C.sTotAct, "总实重"],
    [C.sTotVol, "总材积重"],
    [C.diff2, "差异"],
    [C.rLen, "长(CM)"],
    [C.rWid, "宽(CM)"],
    [C.rHei, "高(CM)"],
    [C.rAct, "实重"],
    [C.rVol, "材积重"],
    [C.rTotAct, "总实重"],
    [C.rTotVol, "总材积重"],
    [C.rCharge, "计费重"],
    [C.rTotCharge, "总计费重"],
    [C.costKg, "成本KG"],
    [C.channel2, "渠道"],
    [C.boxes, "箱数"],
    [C.totCharge, "总计费重"],
    [C.totCost, "总成本重"],
  ];
  const headerRow = ws.getRow(2);
  for (const [col, label] of headerLabels) {
    headerRow.getCell(col).value = label;
  }
  // 表头样式：主体灰蓝 FFADB9CA、分隔列(O=15/X=24)浅绿 FFE2F0D9、尾列(AK-AN)蓝 FF5B9BD5、备注(AJ)无填充无边框
  for (let c = 1; c <= 40; c++) {
    const cell = headerRow.getCell(c);
    cell.alignment = CENTER;
    if (c >= C.channel2 && c <= C.totCost) {
      cell.font = FONT;
      cell.fill = FILL("FF5B9BD5");
    } else if (c === 15 || c === 24) {
      cell.font = FONT;
      cell.fill = FILL("FFE2F0D9");
    } else if (c === C.remark) {
      cell.font = FONT;
    } else {
      cell.font = FONT_BOLD;
      cell.fill = FILL("FFADB9CA");
    }
    if (c !== C.remark) cell.border = THIN_BORDER;
  }

  // ---- 数据行（从第 3 行开始） ----
  const num = (n: number) => (Number.isFinite(n) && n !== 0 ? n : "");
  const formula = (f: string) => ({ formula: f });

  // 按 渠道 → SO 稳定排序（保持组内原始顺序），确保同 SO、同渠道相邻，便于合并与渠道汇总
  const channelOrder = new Map<string, number>();
  const soOrder = new Map<string, number>();
  rows.forEach((r) => {
    if (!channelOrder.has(r.channel)) channelOrder.set(r.channel, channelOrder.size);
    if (!soOrder.has(r.so)) soOrder.set(r.so, soOrder.size);
  });
  const data = [...rows].sort((a, b) => {
    const c = channelOrder.get(a.channel)! - channelOrder.get(b.channel)!;
    if (c !== 0) return c;
    return soOrder.get(a.so)! - soOrder.get(b.so)!;
  });

  // 分组（rows 下标，含两端）：SO 组 = 相邻相同系统SO；渠道组 = 相邻相同渠道
  type Span = { start: number; end: number };
  const soSpans: Span[] = [];
  for (let i = 0; i < data.length; i++) {
    if (i > 0 && data[i].so === data[i - 1].so) soSpans[soSpans.length - 1].end = i;
    else soSpans.push({ start: i, end: i });
  }
  const channelSpans: Span[] = [];
  for (let i = 0; i < data.length; i++) {
    if (i > 0 && data[i].channel === data[i - 1].channel) channelSpans[channelSpans.length - 1].end = i;
    else channelSpans.push({ start: i, end: i });
  }
  const dataStart = 3; // 明细首行（A-AJ 列从 R3 起，渠道汇总只占 AK-AN 列不挤占行）
  const dataEnd = dataStart + data.length - 1; // 明细末行号

  data.forEach((r, i) => {
    const rn = dataStart + i; // 数据起始行
    const row = ws.getRow(rn);
    const L = (col: number) => colLetter(col);
    const isSoFirst = i === 0 || data[i - 1].so !== data[i].so;

    const hasSupplier = r.supplier.lengthCm > 0;

    // 基础列：A 出货日期无来源留空；B(系统SO)/F(单证报关) 仅 SO 首行填值（其余由合并单元格承载）
    if (isSoFirst) {
      row.getCell(C.so).value = r.so;
      row.getCell(C.customs).value = r.customs;
    }
    row.getCell(C.channel).value = r.channel;
    row.getCell(C.country).value = r.country;
    row.getCell(C.warehouse).value = r.warehouse;
    row.getCell(C.fbaId).value = r.fbaId;
    row.getCell(C.productName).value = r.productName;
    row.getCell(C.totalBoxes).value = r.totalBoxes;

    // 客户的（J-N）
    row.getCell(C.cLen).value = num(r.customer.lengthCm);
    row.getCell(C.cWid).value = num(r.customer.widthCm);
    row.getCell(C.cHei).value = num(r.customer.heightCm);
    row.getCell(C.cAct).value = num(r.customer.actualWeight);
    row.getCell(C.cVol).value = formula(`${L(C.cLen)}${rn}*${L(C.cWid)}${rn}*${L(C.cHei)}${rn}/6000`);

    // 差异 P = 客户材积重 - 供应商材积重
    row.getCell(C.diff1).value = formula(`${L(C.cVol)}${rn}-${L(C.sVol)}${rn}`);

    // 供应商（Q-W）
    if (hasSupplier) {
      row.getCell(C.sLen).value = num(r.supplier.lengthCm);
      row.getCell(C.sWid).value = num(r.supplier.widthCm);
      row.getCell(C.sHei).value = num(r.supplier.heightCm);
      row.getCell(C.sAct).value = num(r.supplier.actualWeight);
      row.getCell(C.sVol).value = formula(`${L(C.sLen)}${rn}*${L(C.sWid)}${rn}*${L(C.sHei)}${rn}/6000`);
      row.getCell(C.sTotAct).value = formula(`${L(C.sAct)}${rn}*${L(C.totalBoxes)}${rn}`);
      row.getCell(C.sTotVol).value = formula(`${L(C.sVol)}${rn}*${L(C.totalBoxes)}${rn}`);
    }

    // 差异 Y = 客户材积重 - 出给客户材积重
    row.getCell(C.diff2).value = formula(`${L(C.cVol)}${rn}-${L(C.rVol)}${rn}`);

    // 出给客户（Z-AG）：AH(总计费重) 为 SO 级合计，合并后统一填写，不逐行赋值
    row.getCell(C.rLen).value = num(r.suggestion.lengthCm);
    row.getCell(C.rWid).value = num(r.suggestion.widthCm);
    row.getCell(C.rHei).value = num(r.suggestion.heightCm);
    row.getCell(C.rAct).value = num(r.suggestion.actualWeight);
    row.getCell(C.rVol).value = formula(`${L(C.rLen)}${rn}*${L(C.rWid)}${rn}*${L(C.rHei)}${rn}/6000`);
    row.getCell(C.rTotAct).value = formula(`${L(C.rAct)}${rn}*${L(C.totalBoxes)}${rn}`);
    row.getCell(C.rTotVol).value = formula(`${L(C.rVol)}${rn}*${L(C.totalBoxes)}${rn}`);
    row.getCell(C.rCharge).value = formula(`ROUND(MAX(${L(C.rTotAct)}${rn},${L(C.rTotVol)}${rn}),0)`);

    // 备注 = 报警（无表头列）
    if (r.alarms.length > 0) {
      row.getCell(C.remark).value = r.alarms.join("；");
    }
  });

  // ---- SO 级合并：B(系统SO)/F(单证报关)/AH(总计费重)/AI(成本KG) ----
  // AH = 该 SO 各产品计费重(AG)合计；AI 留空供人工填写（可能与供应商砍价后回填）
  for (const sp of soSpans) {
    const r1 = dataStart + sp.start;
    const r2 = dataStart + sp.end;
    const L = (col: number) => colLetter(col);
    if (r1 < r2) {
      ws.mergeCells(`${L(C.so)}${r1}:${L(C.so)}${r2}`);
      ws.mergeCells(`${L(C.customs)}${r1}:${L(C.customs)}${r2}`);
      ws.mergeCells(`${L(C.rTotCharge)}${r1}:${L(C.rTotCharge)}${r2}`);
      ws.mergeCells(`${L(C.costKg)}${r1}:${L(C.costKg)}${r2}`);
    }
    ws.getCell(`${L(C.rTotCharge)}${r1}`).value = formula(`SUM(${L(C.rCharge)}${r1}:${L(C.rCharge)}${r2})`);
  }

  // ---- 渠道汇总（AK-AN，从第 3 行起连续填，每渠道一行；A-AJ 列明细从 R3 起不受影响）----
  channelSpans.forEach((sp, idx) => {
    const rn = dataStart + idx; // 渠道汇总行连续排在 R3 起
    const L = (col: number) => colLetter(col);
    const r1 = dataStart + sp.start; // 该渠道明细首行
    const r2 = dataStart + sp.end;   // 该渠道明细末行
    ws.getCell(`${L(C.channel2)}${rn}`).value = data[sp.start].channel;
    ws.getCell(`${L(C.boxes)}${rn}`).value = formula(`SUM(${L(C.totalBoxes)}${r1}:${L(C.totalBoxes)}${r2})`);
    ws.getCell(`${L(C.totCharge)}${rn}`).value = formula(`SUM(${L(C.rTotCharge)}${r1}:${L(C.rTotCharge)}${r2})`);
    ws.getCell(`${L(C.totCost)}${rn}`).value = formula(`SUM(${L(C.costKg)}${r1}:${L(C.costKg)}${r2})`);
  });

  // ---- 渠道汇总「合计」行（第 K+1 行，汇总各渠道的箱数/总计费重/总成本重）----
  const sumRow = dataStart + channelSpans.length;
  {
    const L = (col: number) => colLetter(col);
    ws.getCell(`${L(C.channel2)}${sumRow}`).value = "合计";
    ws.getCell(`${L(C.boxes)}${sumRow}`).value = formula(`SUM(${L(C.boxes)}${dataStart}:${L(C.boxes)}${sumRow - 1})`);
    ws.getCell(`${L(C.totCharge)}${sumRow}`).value = formula(`SUM(${L(C.totCharge)}${dataStart}:${L(C.totCharge)}${sumRow - 1})`);
    ws.getCell(`${L(C.totCost)}${sumRow}`).value = formula(`SUM(${L(C.totCost)}${dataStart}:${L(C.totCost)}${sumRow - 1})`);
  }

  // ---- 总合计行（对齐参考文件：无边框无填充，宋体 11 居中）----
  const totalRow = dataEnd + 1;
  const row = ws.getRow(totalRow);
  const L = (col: number) => colLetter(col);
  row.getCell(C.totalBoxes).value = formula(`SUM(${L(C.totalBoxes)}${dataStart}:${L(C.totalBoxes)}${dataEnd})`);
  row.getCell(C.rCharge).value = formula(`SUM(${L(C.rCharge)}${dataStart}:${L(C.rCharge)}${dataEnd})`);
  row.getCell(C.rTotCharge).value = formula(`SUM(${L(C.rTotCharge)}${dataStart}:${L(C.rTotCharge)}${dataEnd})`);
  row.getCell(C.costKg).value = formula(`SUM(${L(C.costKg)}${dataStart}:${L(C.costKg)}${dataEnd})`);

  // ---- 列宽（精确对齐参考文件 40 列）----
  const widths: [number, number][] = [
    [1, 8.99], [2, 15.53], [3, 31.8], [4, 9.76], [5, 13.57], [6, 10.41],
    [7, 16.49], [8, 21.62], [9, 6.63], [10, 7.13], [11, 7.13], [12, 7.13],
    [13, 8.49], [14, 8.99], [15, 1.76], [16, 8.99], [17, 8.99], [18, 8.99],
    [19, 8.99], [20, 8.99], [21, 8.99], [22, 8.99], [23, 9.38], [24, 2.39],
    [25, 8.99], [26, 8.99], [27, 8.99], [28, 8.99], [29, 14.75], [30, 8.99],
    [31, 8.99], [32, 9.38], [33, 8.99], [34, 8.99], [35, 7.62], [36, 25.47],
    [37, 31.35], [38, 8.99], [39, 12.61], [40, 12.6],
  ];
  for (const [col, w] of widths) ws.getColumn(col).width = w;

  // ---- 统一字体/对齐/边框（数据区 R3 起：宋体 11 黑、居中、细黑边框；合计行与备注列 AJ 无边框；表头 R2 样式已单独设置）----
  for (let rn = 3; rn <= totalRow; rn++) {
    const rw = ws.getRow(rn);
    const isTotal = rn === totalRow;
    for (let c = 1; c <= 40; c++) {
      const cell = rw.getCell(c);
      cell.font = FONT;
      cell.alignment = CENTER;
      if (!isTotal && c !== C.remark) cell.border = THIN_BORDER;
    }
  }

  // ---- 数据行数字格式（材积重/总材积重/差异 两位小数负数红；总实重两位；计费重取整）----
  for (let i = 0; i < rows.length; i++) {
    const rw = ws.getRow(dataStart + i);
    for (const c of [C.cVol, C.diff1, C.sVol, C.sTotVol, C.diff2, C.rVol, C.rTotVol]) {
      rw.getCell(c).numFmt = FMT_TWO_DEC;
    }
    rw.getCell(C.rTotAct).numFmt = "0.00_";
    rw.getCell(C.rCharge).numFmt = "0_";
  }

  return Buffer.from(await wb.xlsx.writeBuffer());
}
