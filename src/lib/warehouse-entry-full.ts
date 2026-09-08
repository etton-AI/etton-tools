/**
 * TR全量入仓数据整理（逐箱版） — 客户/供应商箱规比对 → 建议「出给客户」箱规
 *
 * 与原「TR入仓数据整理」的区别：本功能按「逐箱」输出（一个 FBA 号下每个供应商箱各生成一行），
 * 不再做代表箱选数、不累积历史库、无历史对比。
 *
 * 数据流：
 *   客户数据（一行一个产品：FBA ID + 品名 + 总箱数 + 长宽高 + 实重）
 *   供应商数据（逐箱：货箱编号 + 货箱长/高/宽 + 货箱实重/材积重）
 *     ↓
 *   按 FBA 逐箱匹配 + 11 级尺寸修正 + 差异校验报警
 *     ↓
 *   建议「出给客户」箱规（每箱一行）
 *
 * 已确认规则（2026-09-02）：
 *   - FBA 从供应商箱号取「U+流水号」前的部分（如 FBA19MYJ057TU000001 → FBA19MYJ057T）
 *   - 逐箱输出：一个 FBA 号下每个供应商箱各生成一行，totalBoxes = 1
 *   - 出给客户尺寸：材积主导（材积重 > 实重）→ 11 级规则放大；实重主导 → 取该箱原值（长宽高降序）
 *   - 11 级规则：三边降序为 [最长, 次长, 短]，依次尝试规则 1–10，命中第一条「不冲突」即返回，
 *     冲突 = 修正后三边和 vs 客户三边和 差 ≥ 6，或 修正后材积重 vs 客户材积重 差 ≥ 2；规则 11「原值」兜底
 *   - 差异校验（该供应商箱 vs 客户申报，区分材积/实重主导）：
 *      材积主导：三边和差 ≥ 6、材积重差 ≥ 2 → 核查过机图
 *      实重主导：实重差 ≥ 0.5 → 核查过机图
 *   - 供应商过大箱：该 FBA 号最大材积重比当前箱大 ≥ 2 → 核查过机图
 *   - 建议计费重 < 客户计费重 → 提示「建议数据计费重小于客户」
 *   - 成本重校验：Σ出给客户总计费重 < Σ供应商总计费重 → 「成本重过大，请和供应商申请」
 */

import ExcelJS from "exceljs";

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
  /** 完整箱号编码（如 FBA19MYJ057TU000001），用于逐箱展示 */
  boxNo: string;
  /** 提取后的 FBA 号（箱号去掉末尾 U+流水号，如 FBA19MYJ057T） */
  fbaId: string;
  lengthCm: number;
  widthCm: number;
  heightCm: number;
  actualWeight: number;
  volumeWeight: number;
  chargeableWeight: number;
}

/** 供应商箱（逐箱输出时对应该行的那一箱，导出时用于「供应商」对比列） */
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
  /** 完整箱号编码（逐箱输出，未匹配时为空字符串） */
  boxNo: string;
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
  /** 该 FBA 号供应商总计费重（所有箱计费重之和），用于「成本重过大」提醒 */
  supplierFbaTotalChargeable: number;
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
  /** 命中的供应商箱排名（逐箱后不再使用，保留兼容，恒为 null） */
  pickedRank: number | null;
  /** 修正规则展示文本：null=未匹配；「规则1·三边各+1」…「规则11·原值」/「原值(实重主导)」 */
  pickedLabel: string | null;
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
      boxNo,
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
 * 11 级尺寸修正规则（材积主导时对供应商箱尺寸做安全放大）。
 * 先把三边降序排为 [最长, 次长, 短]，依次尝试 11 条规则，命中第一条「不冲突」的即返回；
 * 「冲突」= 修正后三边和 vs 客户三边和 差 ≥ 6，或 修正后材积重 vs 客户材积重 差 ≥ 2。
 * 输出保持 长≥宽≥高 降序；第 11 条「直接取供应商值」为兜底，无论是否冲突都返回原值。
 */
function applyDimensionRules(
  box: { lengthCm: number; widthCm: number; heightCm: number },
  customer: { sumSides: number; volumeWeight: number },
): { lengthCm: number; widthCm: number; heightCm: number; ruleLabel: string } {
  const [longest, middle, shortest] = [box.lengthCm, box.widthCm, box.heightCm].sort(
    (a, b) => b - a,
  );

  // 每条规则对 [最长, 次长, 短] 三边的增量 [ΔL, ΔM, ΔS]
  const RULES: { label: string; dl: number; dm: number; ds: number }[] = [
    { label: "规则1·三边各+1", dl: 1, dm: 1, ds: 1 },
    { label: "规则2·短边+2", dl: 0, dm: 0, ds: 2 },
    { label: "规则3·次长边+短边各+1", dl: 0, dm: 1, ds: 1 },
    { label: "规则4·最长边+短边各+1", dl: 1, dm: 0, ds: 1 },
    { label: "规则5·次长边+2", dl: 0, dm: 2, ds: 0 },
    { label: "规则6·最长边+次长边各+1", dl: 1, dm: 1, ds: 0 },
    { label: "规则7·最长边+2", dl: 2, dm: 0, ds: 0 },
    { label: "规则8·短边+1", dl: 0, dm: 0, ds: 1 },
    { label: "规则9·次长边+1", dl: 0, dm: 1, ds: 0 },
    { label: "规则10·最长边+1", dl: 1, dm: 0, ds: 0 },
  ];

  for (const rule of RULES) {
    const nl = longest + rule.dl;
    const nm = middle + rule.dm;
    const ns = shortest + rule.ds;
    const newSum = sumSides(nl, nm, ns);
    const newVol = calcVolumeWeight(nl, nm, ns);
    const conflict =
      Math.abs(newSum - customer.sumSides) >= SUM_SIDES_THRESHOLD ||
      Math.abs(newVol - customer.volumeWeight) >= VOLUME_DIFF_THRESHOLD;
    if (!conflict) {
      return { lengthCm: nl, widthCm: nm, heightCm: ns, ruleLabel: rule.label };
    }
  }

  // 规则 11：直接取供应商原值（前面全部冲突，兜底不放大）
  return {
    lengthCm: longest,
    widthCm: middle,
    heightCm: shortest,
    ruleLabel: "规则11·原值",
  };
}

/**
 * 放大建议值，使材积重/计费重超过 target。
 * 材积主导（材积重 ≥ 实重）时按单边整数放大（每次把最短边 +1）；实重主导时直接放大实重。
 * 放大受差异约束限制（出给客户 vs 客户申报，不能突破差异报警阈值）：
 *   材积主导：边和差 < 6 且 材积重差 < 2；
 *   实重主导：实重差 < 0.5。
 * 触达约束上限即停止放大。
 */
function forceAmplify(
  s: SuggestionRow["suggestion"],
  customer: SuggestionRow["customer"],
  target: number,
): void {
  if (s.volumeWeight >= s.actualWeight) {
    // 材积主导：单边整数放大（每次把最短边 +1），直到材积重 > target 或触达差异约束上限
    let L = s.lengthCm;
    let W = s.widthCm;
    let H = s.heightCm;

    let guard = 0;
    while (calcVolumeWeight(L, W, H) <= target && guard < 500) {
      let nl = L;
      let nw = W;
      let nh = H;
      if (L <= W && L <= H) nl += 1;
      else if (W <= H) nw += 1;
      else nh += 1;

      const nextVol = calcVolumeWeight(nl, nw, nh);
      const nextSum = sumSides(nl, nw, nh);
      // 差异约束：边和差 < 6 且 材积重差 < 2，超限则停止放大
      if (
        Math.abs(nextSum - customer.sumSides) >= SUM_SIDES_THRESHOLD ||
        Math.abs(nextVol - customer.volumeWeight) >= VOLUME_DIFF_THRESHOLD
      ) {
        break;
      }

      L = nl;
      W = nw;
      H = nh;
      guard++;
    }

    s.lengthCm = L;
    s.widthCm = W;
    s.heightCm = H;
    s.volumeWeight = calcVolumeWeight(L, W, H);
    s.chargeableWeight = Math.max(s.actualWeight, s.volumeWeight);
  } else {
    // 实重主导：放大实重，但实重差 < 0.5（不突破差异报警阈值）
    const cap = customer.actualWeight + ACTUAL_DIFF_THRESHOLD - 0.01;
    s.actualWeight = Math.max(s.actualWeight, Math.min(round2(target), cap));
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

    // 未匹配到供应商箱：保持一行，需人工复核
    if (!group || group.length === 0) {
      rows.push({
        fbaId: c.fbaId,
        boxNo: "",
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
        supplierFbaTotalChargeable: 0,
        supplierMaxVolumeWeight: 0,
        suggestion: { ...customer },
        alarms: ["⚠需人工复核"],
        pickedRank: null,
        pickedLabel: null,
      });
      continue;
    }

    const supplierFbaTotalChargeable = group.reduce((s, b) => s + b.chargeableWeight, 0);
    const supplierMaxVolumeWeight = Math.max(...group.map((b) => b.volumeWeight));

    // 客户材积主导：材积重 > 实重 → 对每箱按 11 级规则放大；否则尺寸/实重直接取该箱原值。
    const volumeLed = c.volumeWeight > c.actualWeight;

    // 逐箱输出：一个 FBA 号下每个供应商箱各生成一行（每行 = 一箱）。
    for (const b of group) {
      const boxAlarms: string[] = [];

      // 供应商列 = 该箱本身（不再是「代表箱」）
      const supplier: SupplierRepresentative = {
        lengthCm: b.lengthCm,
        widthCm: b.widthCm,
        heightCm: b.heightCm,
        actualWeight: b.actualWeight,
        volumeWeight: b.volumeWeight,
      };

      // 出给客户尺寸：材积主导 → 11 级规则；实重主导 → 原值不放大（降序）
      let dims: { lengthCm: number; widthCm: number; heightCm: number };
      let ruleLabel: string;
      if (volumeLed) {
        const res = applyDimensionRules(b, customer);
        dims = { lengthCm: res.lengthCm, widthCm: res.widthCm, heightCm: res.heightCm };
        ruleLabel = res.ruleLabel;
      } else {
        const d = [b.lengthCm, b.widthCm, b.heightCm].sort((x, y) => y - x);
        dims = { lengthCm: d[0], widthCm: d[1], heightCm: d[2] };
        ruleLabel = "原值(实重主导)";
      }

      const volWeight = calcVolumeWeight(dims.lengthCm, dims.widthCm, dims.heightCm);
      const suggestion: SuggestionRow["suggestion"] = {
        lengthCm: dims.lengthCm,
        widthCm: dims.widthCm,
        heightCm: dims.heightCm,
        actualWeight: b.actualWeight,
        volumeWeight: volWeight,
        chargeableWeight: Math.max(b.actualWeight, volWeight),
        sumSides: sumSides(dims.lengthCm, dims.widthCm, dims.heightCm),
      };

      // 逐箱差异校验（该箱 vs 客户申报）
      if (volumeLed) {
        if (Math.abs(suggestion.sumSides - c.sumSides) >= SUM_SIDES_THRESHOLD) {
          boxAlarms.push("三边和差异超限，请核查过机图");
        }
        if (Math.abs(suggestion.volumeWeight - c.volumeWeight) >= VOLUME_DIFF_THRESHOLD) {
          boxAlarms.push("材积重差异超限，请核查过机图");
        }
      } else {
        if (Math.abs(suggestion.actualWeight - c.actualWeight) >= ACTUAL_DIFF_THRESHOLD) {
          boxAlarms.push("实重差异超限，请核查过机图");
        }
      }
      // 供应商存在过大箱：该 FBA 号最大材积重比当前箱大 ≥2（供应商箱规不齐，混入异常大箱）
      if (supplierMaxVolumeWeight - suggestion.volumeWeight >= SUPPLIER_EXCESS_THRESHOLD) {
        boxAlarms.push("供应商存在过大箱，请核查过机图");
      }
      // 建议出给客户计费重 < 客户计费重：单独提示
      if (suggestion.chargeableWeight < c.chargeableWeight) {
        boxAlarms.push("建议数据计费重小于客户，请确认");
      }

      rows.push({
        fbaId: c.fbaId,
        boxNo: b.boxNo,
        productName: c.productName,
        so: c.so,
        channel: c.channel,
        country: c.country,
        warehouse: c.warehouse,
        customs: c.customs,
        totalBoxes: 1, // 逐箱：每行代表 1 箱
        customer,
        supplier,
        supplierChargeable: b.chargeableWeight,
        supplierFbaTotalChargeable,
        supplierMaxVolumeWeight,
        suggestion,
        alarms: boxAlarms,
        pickedRank: null,
        pickedLabel: ruleLabel,
      });
    }
  }

  // 按 FBA 号聚合：出给客户总计费重 vs 供应商该 FBA 号总计费重，负差提示成本重过大（紫色）。
  const fbaSuggestTotal = new Map<string, number>();
  for (const r of rows) {
    if (r.pickedLabel === null) continue;
    fbaSuggestTotal.set(r.fbaId, (fbaSuggestTotal.get(r.fbaId) ?? 0) + r.suggestion.chargeableWeight * r.totalBoxes);
  }
  for (const r of rows) {
    if (r.pickedLabel === null) continue;
    const suggestTotal = fbaSuggestTotal.get(r.fbaId) ?? 0;
    if (suggestTotal - r.supplierFbaTotalChargeable < 0) {
      r.alarms.push("成本重过大，请和供应商申请");
    }
  }

  return { rows, supplierTotal };
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
): Promise<{ buffer: Buffer; warning: string | null }> {
  // 全局兜底校验：出给客户总重 ≤ 供应商总重时，尝试放大（受差异约束限制）。
  let warning: string | null = null;
  const customerTotal = rows.reduce(
    (s, r) => s + r.suggestion.chargeableWeight * r.totalBoxes,
    0,
  );
  if (rows.length > 0 && customerTotal > 0 && customerTotal <= supplierTotal) {
    const ratio = (supplierTotal / customerTotal) * AMPLIFY_RATIO;
    for (const r of rows) {
      forceAmplify(r.suggestion, r.customer, r.suggestion.chargeableWeight * ratio);
      r.alarms.push("[全局调整]");
    }
    // 放大后重新计算总重：若差异约束挡住导致仍 < 供应商总重，提示可能亏损
    const amplifiedTotal = rows.reduce(
      (s, r) => s + r.suggestion.chargeableWeight * r.totalBoxes,
      0,
    );
    if (amplifiedTotal < supplierTotal) {
      warning = "出给客户总计费重小于供应商总计费重，可能亏损，请留意！";
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
    fbaId: 7, // G 箱号(完整编码)
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
    pickedLabel: 36, // AJ 选中名次
    remark: 37, // AK 备注（无表头）
    channel2: 38, // AL 渠道
    boxes: 39, // AM 箱数
    totCharge: 40, // AN 总计费重
    totCost: 41, // AO 总成本重
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
    [C.fbaId, "箱号"],
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
    [C.pickedLabel, "修正规则"],
    [C.channel2, "渠道"],
    [C.boxes, "箱数"],
    [C.totCharge, "总计费重"],
    [C.totCost, "总成本重"],
  ];
  const headerRow = ws.getRow(2);
  for (const [col, label] of headerLabels) {
    headerRow.getCell(col).value = label;
  }
  // 表头样式：主体灰蓝 FFADB9CA、分隔列(O=15/X=24)浅绿 FFE2F0D9、尾列(AL-AO)蓝 FF5B9BD5、备注(AK)无填充无边框
  for (let c = 1; c <= 41; c++) {
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
    row.getCell(C.fbaId).value = r.boxNo;
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

    // 选中名次（AJ）
    row.getCell(C.pickedLabel).value = r.pickedLabel ?? "";

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
    [31, 8.99], [32, 9.38], [33, 8.99], [34, 8.99], [35, 7.62], [36, 10],
    [37, 25.47], [38, 31.35], [39, 8.99], [40, 12.61], [41, 12.6],
  ];
  for (const [col, w] of widths) ws.getColumn(col).width = w;

  // ---- 统一字体/对齐/边框（数据区 R3 起：宋体 11 黑、居中、细黑边框；合计行与备注列 AK 无边框；表头 R2 样式已单独设置）----
  for (let rn = 3; rn <= totalRow; rn++) {
    const rw = ws.getRow(rn);
    const isTotal = rn === totalRow;
    for (let c = 1; c <= 41; c++) {
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

  // ---- 高亮样式（在统一字体/边框之后设置，避免被 cell.font = FONT 覆盖）----
  const RED_FONT = { name: "宋体", size: 11, color: { argb: "FFFF0000" } };
  const FILL_YELLOW_STRONG = FILL("FFFFFF00");
  const FILL_RED = FILL("FFFFC7CE");
  const FILL_YELLOW = FILL("FFFFEB9C");
  const FILL_BLUE = FILL("FFDDEBF7");
  const FILL_PURPLE = FILL("FFD9C2F0");
  for (let i = 0; i < data.length; i++) {
    const r = data[i];
    const rn = dataStart + i;
    const rw = ws.getRow(rn);
    const hasSupplier = r.supplier.lengthCm > 0;

    // Y 列（差异 = 客户材积重 − 出给客户材积重）差异绝对值 ≥ 2 标黄
    const customerVol = calcVolumeWeight(r.customer.lengthCm, r.customer.widthCm, r.customer.heightCm);
    const suggestionVol = calcVolumeWeight(r.suggestion.lengthCm, r.suggestion.widthCm, r.suggestion.heightCm);
    if (Math.abs(round2(customerVol - suggestionVol)) >= 2) {
      rw.getCell(C.diff2).fill = FILL_YELLOW_STRONG;
    }

    // 出给客户 vs 供应商 对比不同 → 出给客户列（Z/AA/AB/AC）红字
    if (hasSupplier) {
      if (r.suggestion.lengthCm !== r.supplier.lengthCm) rw.getCell(C.rLen).font = RED_FONT;
      if (r.suggestion.widthCm !== r.supplier.widthCm) rw.getCell(C.rWid).font = RED_FONT;
      if (r.suggestion.heightCm !== r.supplier.heightCm) rw.getCell(C.rHei).font = RED_FONT;
      if (r.suggestion.actualWeight !== r.supplier.actualWeight) rw.getCell(C.rAct).font = RED_FONT;
    }

    // AJ 备注底色（紫 > 红 > 黄 > 蓝）
    const remarkCell = rw.getCell(C.remark);
    const hasPurple = r.alarms.some((a) => a.includes("成本重过大"));
    const hasRed = r.alarms.some((a) => a.includes("核查过机图"));
    const hasYellow = r.alarms.some(
      (a) => a.includes("需人工复核") || a.includes("需人工确认") || a.includes("建议数据计费重小于客户"),
    );
    const hasBlue = r.alarms.some((a) => a.includes("建议参考历史"));
    if (hasPurple) remarkCell.fill = FILL_PURPLE;
    else if (hasRed) remarkCell.fill = FILL_RED;
    else if (hasYellow) remarkCell.fill = FILL_YELLOW;
    else if (hasBlue) remarkCell.fill = FILL_BLUE;
  }

  return {
    buffer: Buffer.from(await wb.xlsx.writeBuffer()),
    warning,
  };
}
