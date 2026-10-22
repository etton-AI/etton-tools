/**
 * 跨境堡批量投保 — 将「系统导出的货箱清单」按 SO 号分组，填充「跨境堡批量投保箱单模版」。
 *
 * 数据流：
 *   货箱清单（Sheet1，标题行 + 表头 + 数据行）
 *     ├─ 系统SO（仅每组首行，需向下填充）
 *     ├─ 国家 / 仓库代码 / FBA ID / 中文品名 / 材质 / 申报总数量 / 总申报货值 / 申报币种
 *     ├─ 总箱数(CTN) / 货箱重量（合并）/ 单箱货物毛重 / 客户渠道
 *       ↓
 *   跨境堡批量投保箱单模版（public/templates/跨境堡批量投保箱单模版.xlsx）
 *     ├─ R1 表头 / R2 说明
 *     └─ R3 起数据区，36 列，逐票（SO）一行
 *
 * 核心计算：
 *   1. 货物描述：按 SO 聚合「中文品名 + 申报总数量」→ `品名+数量PCS`，多品名以「、」连接
 *   2. 货值：总申报货值求和（四舍五入 2 位）
 *   3. 币种：申报币种 → 中文名（USD → 美元）
 *   4. 备注：仅「全美锁仓」渠道（客户渠道含「全美」）+ 入仓编号溢出的多余 FBA
 *
 * 审核（2026-10 新增）：
 *   a. 易碎品拆分：按「材质」判定易碎（玻璃/陶瓷/大理石/水晶/亚克力及其制品），
 *      一票内同时含易碎与非易碎时拆成两票（易碎那票货物类别「是易碎品」、非易碎「无易碎品」）。
 *   b. 币种校验：按目的国预期申报币种（如德国→EUR），与清单申报币种不一致时告警。
 *   c. FBA 仓校验：入仓编号不含「FBA」字眼时判定为非 FBA 仓（卫星仓/沃尔玛仓）告警。
 */

import ExcelJS from "exceljs";
import fs from "fs";
import path from "path";

// ============================================================
// Types
// ============================================================

export interface KuajingbaoConfig {
  /** 起运日期（YYYY-MM-DD 或 YYYY/MM/DD） */
  shipDate: string;
  /** 干线运输方式，默认 海运 */
  transportMode: string;
  /** 派送方式，默认 卡车派 */
  deliveryMethod: string;
  /** 起运地，默认 深圳 */
  originPlace: string;
  /** 包装类型，默认 纸箱 */
  packageType: string;
}

/** 一票（一个 SO，或拆分后的易碎/非易碎子票）的聚合结果 */
export interface SoGroup {
  so: string;              // 原始系统SO（用于告警/拆分标识，不直接写入运单）
  orderNo: string;         // 原单号（拆分子票带 -1/-2 后缀，保证唯一；同时填运输工具名称及航次 / 提单号）
  workOrder: string;       // 工作号（仅用于展示/排查）
  country: string;         // 国家 → 目的国
  warehouseCode: string;   // 仓库代码 → 目的地
  fbaList: string[];       // 该票全部 FBA 号（去箱序后缀、去重、保持出现顺序）
  fbaBase: string;         // 入仓编号（fbaList 以空格连接，≤350 字符；超出部分见 remark）
  channel: string;         // 客户渠道（判断是否全美锁仓）
  goodsDescription: string;// 货物描述
  goodsCategory: string;   // 货物类别（按材质自动判定：是易碎品 / 无易碎品）
  isFragile: boolean;      // 是否易碎票
  currency: string;        // 币种（中文）
  currencyCode: string;    // 币种代码（原始，如 USD/EUR，用于审核校验）
  totalValue: number;      // 货值
  totalBoxes: number;      // 总包装数量
  totalWeight: number;     // 总公斤数
  remark: string;          // 备注（全美锁仓说明 + 入仓编号溢出的多余 FBA）
}

/** 币种异常告警 */
export interface CurrencyWarning {
  so: string;          // 系统SO
  country: string;     // 目的国
  declared: string;    // 清单申报币种代码（如 USD）
  expected: string;    // 预期币种代码（如 EUR）
}

/** 非 FBA 仓告警 */
export interface NonFbaWarning {
  so: string;          // 系统SO
  fbaCodes: string[];  // 不含「FBA」字眼的入仓编号
}

export interface KuajingbaoResult {
  sourceFile: string;
  groups: SoGroup[];
  lockWarehouseCount: number;      // 全美锁仓（含备注）票数
  currencyWarnings: CurrencyWarning[]; // 币种异常告警
  nonFbaWarnings: NonFbaWarning[];     // 非 FBA 仓告警
  splitInfos: { so: string }[];        // 被拆成易碎/非易碎两票的 SO
  buffer: Buffer;
}

/** 货箱清单表头 → 列号 */
interface CargoColumnMap {
  so: number;
  workOrder: number;
  country: number;
  warehouseCode: number;
  fbaId: number;
  nameCh: number;
  boxWeight: number;
  boxCount: number;
  declaredQty: number;
  declaredValue: number;
  currency: number;
  channel: number;
  material: number;         // 材质（必填，易碎判定依据）
  boxGrossWeight?: number;  // 单箱货物毛重（可选，按行算公斤数用）
  boxNetWeight?: number;    // 单箱货物净重（可选，毛重缺失时兜底）
}

// ============================================================
// 常量
// ============================================================

/** 跨境堡模板主 sheet 名 */
const TEMPLATE_SHEET = "批量导入投保数据";

/** 入仓编号字数上限（含分隔符）；超出部分按「完整 FBA 号」为单位挪到备注，绝不截断单个 FBA */
const FBA_FIELD_LIMIT = 350;

/** 多个 FBA 号之间的分隔符（空格） */
const FBA_SEP = " ";

/** 模板内置常量（与示例数据一致） */
const INSURED = "易通科技物流（中山）有限公司";
const TYPE = "企业";
const ORIGIN_COUNTRY = "中国";
const CLAIM_PLACE = "HANGZHOU CHINA";
const DEST_TYPE = "FBA";
const SHELF_GUARANTEE = "保上架";
const RATIO = 1; // 加成比例 100%

/** 货物类别：按材质判定 */
const GOODS_CATEGORY_FRAGILE = "是易碎品";
const GOODS_CATEGORY_NORMAL = "无易碎品";

/** 币种映射：申报币种代码 → 中文名 */
const CURRENCY_MAP: Record<string, string> = {
  USD: "美元",
  EUR: "欧元",
  GBP: "英镑",
  JPY: "日元",
  CNY: "人民币",
  RMB: "人民币",
  AUD: "澳元",
  CAD: "加元",
  HKD: "港币",
  KRW: "韩元",
  SGD: "新加坡元",
  CHF: "瑞士法郎",
  MXN: "墨西哥比索",
  AED: "迪拉姆",
};

/**
 * 易碎材质关键词（中文 + 英文，小写匹配）。
 * 「及其制品」无需单独列举：玻璃制品/陶瓷制品等已包含关键词本身。
 */
const FRAGILE_MATERIAL_KEYWORDS = [
  "玻璃", "陶瓷", "大理石", "水晶", "亚克力",
  "glass", "ceramic", "marble", "crystal", "acrylic",
];

/**
 * 国家 → 预期申报币种代码。仅覆盖常见目的国；不在表中的国家不做币种校验。
 * 注意欧盟内非欧元区国家（波兰 PLN / 捷克 CZK / 匈牙利 HUF / 罗马尼亚 RON / 瑞典 SEK / 丹麦 DKK）单独列出。
 */
const COUNTRY_CURRENCY: Record<string, string> = {
  美国: "USD",
  英国: "GBP",
  // 欧元区
  德国: "EUR", 法国: "EUR", 西班牙: "EUR", 意大利: "EUR", 荷兰: "EUR",
  比利时: "EUR", 奥地利: "EUR", 葡萄牙: "EUR", 芬兰: "EUR", 爱尔兰: "EUR",
  希腊: "EUR", 卢森堡: "EUR", 斯洛伐克: "EUR", 斯洛文尼亚: "EUR", 塞浦路斯: "EUR",
  马耳他: "EUR", 爱沙尼亚: "EUR", 拉脱维亚: "EUR", 立陶宛: "EUR", 克罗地亚: "EUR",
  // 其他主要币种
  加拿大: "CAD",
  澳大利亚: "AUD",
  日本: "JPY",
  阿联酋: "AED",
  瑞士: "CHF",
  韩国: "KRW",
  新加坡: "SGD",
  墨西哥: "MXN",
  中国香港: "HKD",
  新西兰: "NZD",
  // 欧盟非欧元区
  波兰: "PLN", 捷克: "CZK", 匈牙利: "HUF", 罗马尼亚: "RON", 瑞典: "SEK", 丹麦: "DKK",
  // 其他
  挪威: "NOK",
  泰国: "THB",
  印度: "INR",
  印尼: "IDR", 印度尼西亚: "IDR",
  马来西亚: "MYR",
  菲律宾: "PHP",
  越南: "VND",
  巴西: "BRL",
  南非: "ZAR",
  土耳其: "TRY",
  俄罗斯: "RUB",
  以色列: "ILS",
  沙特阿拉伯: "SAR",
  卡塔尔: "QAR",
  科威特: "KWD",
};

/** 货箱清单表头关键词（trim 后 startsWith 匹配，兼容括号说明如「总箱数(CTN)」） */
const CARGO_HEADER_PATTERNS: Record<keyof CargoColumnMap, string[]> = {
  so: ["系统SO", "SO"],
  workOrder: ["工作号"],
  country: ["国家"],
  warehouseCode: ["仓库代码"],
  fbaId: ["FBA ID", "FBA"],
  nameCh: ["中文品名"],
  boxWeight: ["货箱重量"],
  boxCount: ["总箱数"],
  declaredQty: ["申报总数量"],
  declaredValue: ["总申报货值"],
  currency: ["申报币种"],
  channel: ["客户渠道"],
  material: ["材质"],
  boxGrossWeight: ["单箱货物毛重"],
  boxNetWeight: ["单箱货物净重"],
};

// ============================================================
// Helpers
// ============================================================

function cellText(cell: ExcelJS.Cell): string {
  const v = cell.value;
  if (v === null || v === undefined) return "";
  if (typeof v === "string") return v.trim();
  if (typeof v === "number") {
    // 大整数（如 FBA 号被存成数字）按整数输出，避免科学计数法
    if (v > 1e15) return String(Math.floor(v));
    return String(v);
  }
  if (typeof v === "object") {
    const o = v as { richText?: { text: string }[]; result?: unknown; text?: unknown };
    if (Array.isArray(o.richText)) return o.richText.map((r) => r.text).join("").trim();
    if (o.result !== undefined && o.result !== null) {
      const r = o.result;
      if (typeof r === "string") return r.trim();
      if (typeof r === "number") return r > 1e15 ? String(Math.floor(r)) : String(r);
      return "";
    }
    if (o.text !== undefined && o.text !== null) return String(o.text).trim();
  }
  return String(v).trim();
}

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

/** 四舍五入到 2 位小数，去掉浮点尾差 */
function round2(n: number): number {
  return Math.round(n * 100) / 100;
}

/** 币种代码 → 中文名；未知代码原样返回 */
function currencyToCn(code: string): string {
  const key = code.trim().toUpperCase();
  return CURRENCY_MAP[key] ?? code.trim() ?? "美元";
}

/** 材质是否为易碎（含玻璃/陶瓷/大理石/水晶/亚克力及其制品） */
function isFragileMaterial(material: string): boolean {
  const m = material.toLowerCase();
  return FRAGILE_MATERIAL_KEYWORDS.some((k) => m.includes(k.toLowerCase()));
}

/** 入仓编号是否属于 FBA 仓（含「FBA」字眼） */
function isFbaCode(fbaCode: string): boolean {
  return fbaCode.toUpperCase().includes("FBA");
}

/** 清洗文件名非法字符 */
function sanitizeFileName(name: string): string {
  const cleaned = name.replace(/[\\/:*?"<>|]/g, "_").trim();
  return cleaned || "未命名";
}

/**
 * 解析「YYYY-MM-DD」或「YYYY/MM/DD」为 Date，非法返回 today。
 * 注意用 Date.UTC 构造：exceljs 的 dateToExcel 直接用 getTime()（UTC 时基），
 * 若用本地时区构造（new Date(y,m-1,d)），在东八区会因时区偏移少一天。
 */
function parseShipDate(text: string): Date {
  const m = text.trim().match(/^(\d{4})[/-](\d{1,2})[/-](\d{1,2})/);
  if (m) {
    const y = Number(m[1]);
    const mo = Number(m[2]);
    const d = Number(m[3]);
    const date = new Date(Date.UTC(y, mo - 1, d));
    if (!isNaN(date.getTime())) return date;
  }
  return new Date();
}

// ============================================================
// 货箱清单解析
// ============================================================

/** 定位货箱清单表头行（含「系统SO」和「中文品名」） */
function findCargoHeaderRow(ws: ExcelJS.Worksheet): number {
  for (let r = 1; r <= Math.min(ws.rowCount, 10); r++) {
    const row = ws.getRow(r);
    let hasSo = false;
    let hasName = false;
    row.eachCell({ includeEmpty: false }, (cell) => {
      const t = cellText(cell);
      if (t === "系统SO") hasSo = true;
      if (t === "中文品名") hasName = true;
    });
    if (hasSo && hasName) return r;
  }
  throw new Error("找不到货箱清单表头（需同时包含「系统SO」和「中文品名」列）。请确认上传的是系统导出的货箱清单。");
}

/** 根据表头行构建列映射 */
function buildCargoColumnMap(headerRow: ExcelJS.Row): CargoColumnMap {
  const map: Partial<CargoColumnMap> = {};
  headerRow.eachCell({ includeEmpty: false }, (cell, cn) => {
    const t = cellText(cell);
    for (const [key, patterns] of Object.entries(CARGO_HEADER_PATTERNS)) {
      if (map[key as keyof CargoColumnMap] !== undefined) continue;
      if (patterns.some((p) => t.startsWith(p))) {
        map[key as keyof CargoColumnMap] = cn;
        break;
      }
    }
  });

  const required: (keyof CargoColumnMap)[] = [
    "so", "country", "warehouseCode", "fbaId", "nameCh",
    "boxWeight", "boxCount", "declaredQty", "declaredValue", "currency", "channel", "material",
  ];
  const missing = required.filter((k) => map[k] === undefined);
  if (missing.length > 0) {
    throw new Error(`货箱清单表头缺少以下列: ${missing.join("、")}。请确认上传的是系统导出的货箱清单。`);
  }
  return map as CargoColumnMap;
}

/** 货箱清单单个数据行（用于按材质拆分） */
interface CargoRow {
  nameCh: string;
  qty: number;
  value: number;
  boxes: number;
  material: string;
  isFragile: boolean;
  rowWeight: number; // 该行公斤数 = 单箱毛重 × 总箱数
  fbaId: string;     // 该行 FBA 号（去箱序后缀），按材质拆分时随行归到对应子票
}

/** 解析中间态：一个 SO 的原始行 + 汇总信息 */
interface SoAccumulator {
  so: string;
  workOrder: string;
  country: string;
  warehouseCode: string;
  channel: string;
  currencyCode: string;
  currency: string;
  boxWeight: number; // 货箱重量（合并单元格总值，仅兜底用）
  fbaList: string[];
  rows: CargoRow[];
}

interface ParseResult {
  groups: SoGroup[];
  currencyWarnings: CurrencyWarning[];
  nonFbaWarnings: NonFbaWarning[];
  splitInfos: { so: string }[];
}

/** 由一组行生成一票（SO 或拆分后的子票） */
function buildGroup(acc: SoAccumulator, rows: CargoRow[], isFragile: boolean, orderNo: string): SoGroup {
  // 货物描述：同品名数量累加，保持首次出现顺序
  const qtyMap = new Map<string, number>();
  for (const r of rows) {
    if (!r.nameCh) continue;
    qtyMap.set(r.nameCh, (qtyMap.get(r.nameCh) || 0) + r.qty);
  }
  const parts: string[] = [];
  for (const [name, q] of qtyMap) parts.push(`${name}+${q}PCS`);

  const totalValue = round2(rows.reduce((s, r) => s + r.value, 0));
  const totalBoxes = rows.reduce((s, r) => s + r.boxes, 0);

  // 入仓编号：该组（材质侧）各行 FBA 号去重、保持出现顺序，不再取整票 FBA
  const fbaList: string[] = [];
  for (const r of rows) {
    if (r.fbaId && !fbaList.includes(r.fbaId)) fbaList.push(r.fbaId);
  }

  // 公斤数：优先按行（单箱毛重×箱数）累加；无按行重量时用合并单元格货箱重量兜底
  let totalWeight = round2(rows.reduce((s, r) => s + r.rowWeight, 0));
  if (totalWeight === 0 && acc.boxWeight > 0) {
    // 兜底：拆分组按货值占比分摊总值，非拆分整取
    const groupValue = rows.reduce((s, r) => s + r.value, 0);
    const allValue = acc.rows.reduce((s, r) => s + r.value, 0);
    totalWeight = allValue > 0 ? round2((acc.boxWeight * groupValue) / allValue) : round2(acc.boxWeight);
  }

  return {
    so: acc.so,
    orderNo,
    workOrder: acc.workOrder,
    country: acc.country,
    warehouseCode: acc.warehouseCode,
    fbaList,
    fbaBase: "",
    channel: acc.channel,
    goodsDescription: parts.join("、"),
    goodsCategory: isFragile ? GOODS_CATEGORY_FRAGILE : GOODS_CATEGORY_NORMAL,
    isFragile,
    currency: acc.currency,
    currencyCode: acc.currencyCode,
    totalValue,
    totalBoxes,
    totalWeight,
    remark: "",
  };
}

/** 收尾：入仓编号（350 截断）+ 备注（全美锁仓 + 溢出 FBA） */
function finalizeGroup(g: SoGroup): void {
  // 入仓编号：全部 FBA 号以空格连接，限 350 字符（含分隔符）；
  // 按「完整 FBA 号」为单位截断，溢出部分（多余 FBA 号）整段挪到备注，绝不把单个 FBA 号从中间拆开。
  let fbaBase = "";
  let overflowFbas: string[] = [];
  for (let i = 0; i < g.fbaList.length; i++) {
    const f = g.fbaList[i];
    const candidate = fbaBase ? fbaBase + FBA_SEP + f : f;
    if (candidate.length <= FBA_FIELD_LIMIT) {
      fbaBase = candidate;
    } else {
      overflowFbas = g.fbaList.slice(i);
      break;
    }
  }
  g.fbaBase = fbaBase;

  // 备注 = 全美锁仓说明（若有）+ 入仓编号溢出的多余 FBA（若有）
  const remarkParts: string[] = [];
  if (g.channel.includes("全美")) {
    const code = g.warehouseCode;
    remarkParts.push(
      `本保单承保AGL（亚马逊物流）承运的物流运输服务，其亚马逊货件建仓地址为“${code}”，` +
      `投保的目的地为“${code}”，同时存在最后FBA调仓可能`
    );
  }
  if (overflowFbas.length > 0) {
    remarkParts.push(`入仓编号超出${FBA_FIELD_LIMIT}字符，多余FBA：${overflowFbas.join(FBA_SEP)}`);
  }
  g.remark = remarkParts.join("；");
}

/** 解析并聚合：按 SO 号（向下填充）分组，做审核（易碎拆分/币种/FBA） */
function parseCargoList(ws: ExcelJS.Worksheet): ParseResult {
  const headerRow = findCargoHeaderRow(ws);
  const col = buildCargoColumnMap(ws.getRow(headerRow));

  // 按 SO 分组，保持出现顺序
  const accs: SoAccumulator[] = [];
  const indexBySo = new Map<string, number>();

  let currentSo = "";

  for (let r = headerRow + 1; r <= ws.rowCount; r++) {
    const soCell = cellText(ws.getCell(r, col.so));
    if (soCell) currentSo = soCell; // 首行带 SO，续行向下填充
    if (!currentSo) continue; // 表头下方尚未遇到首个 SO

    if (!indexBySo.has(currentSo)) {
      indexBySo.set(currentSo, accs.length);
      accs.push({
        so: currentSo,
        workOrder: cellText(ws.getCell(r, col.workOrder)),
        country: cellText(ws.getCell(r, col.country)),
        warehouseCode: cellText(ws.getCell(r, col.warehouseCode)),
        channel: cellText(ws.getCell(r, col.channel)),
        currencyCode: cellText(ws.getCell(r, col.currency)).trim().toUpperCase(),
        currency: currencyToCn(cellText(ws.getCell(r, col.currency))),
        // 货箱重量是「每 SO 合并单元格」的总额（续行不重复），只取首行一次；
        // 若放进逐行累加，ExcelJS 会把合并格的值在合并区每行都读到 → 总公斤数被乘 N 倍。
        boxWeight: cellNum(ws.getCell(r, col.boxWeight)),
        fbaList: [],
        rows: [],
      });
    }

    const acc = accs[indexBySo.get(currentSo)!];

    // 币种/渠道：取非空值兜底
    const curRaw = cellText(ws.getCell(r, col.currency));
    if (curRaw && !acc.currencyCode) {
      acc.currencyCode = curRaw.trim().toUpperCase();
      acc.currency = currencyToCn(curRaw);
    }
    const ch = cellText(ws.getCell(r, col.channel));
    if (ch && !acc.channel) acc.channel = ch;

    // 入仓编号：收集该票全部 FBA ID（去箱序后缀 U+6 位数字，去重，保持首次出现顺序）
    const fbaId = cellText(ws.getCell(r, col.fbaId)).replace(/U\d{6}$/, "").trim();
    if (fbaId && !acc.fbaList.includes(fbaId)) {
      acc.fbaList.push(fbaId);
    }

    // 行数据（有品名或数量才计入，避免续行空品名误计）
    const nameCh = cellText(ws.getCell(r, col.nameCh));
    const qty = cellNum(ws.getCell(r, col.declaredQty));
    if (nameCh || qty > 0) {
      const boxes = cellNum(ws.getCell(r, col.boxCount));
      const material = cellText(ws.getCell(r, col.material));
      // 按行公斤数 = 单箱毛重 × 总箱数（毛重缺失用净重兜底）
      let perBoxWeight = 0;
      if (col.boxGrossWeight !== undefined) perBoxWeight = cellNum(ws.getCell(r, col.boxGrossWeight));
      if (perBoxWeight === 0 && col.boxNetWeight !== undefined) perBoxWeight = cellNum(ws.getCell(r, col.boxNetWeight));
      acc.rows.push({
        nameCh,
        qty,
        value: cellNum(ws.getCell(r, col.declaredValue)),
        boxes,
        material,
        isFragile: isFragileMaterial(material),
        rowWeight: round2(perBoxWeight * boxes),
        fbaId,
      });
    }
  }

  if (accs.length === 0) {
    throw new Error("货箱清单中未找到有效的 SO 数据行。");
  }

  // 逐 SO 生成最终票 + 审核告警
  const groups: SoGroup[] = [];
  const currencyWarnings: CurrencyWarning[] = [];
  const nonFbaWarnings: NonFbaWarning[] = [];
  const splitInfos: { so: string }[] = [];

  for (const acc of accs) {
    // 审核 b：币种校验（目的国预期币种 vs 清单申报币种）
    const expected = COUNTRY_CURRENCY[acc.country];
    if (expected && acc.currencyCode && acc.currencyCode !== expected) {
      currencyWarnings.push({
        so: acc.so,
        country: acc.country,
        declared: acc.currencyCode,
        expected,
      });
    }

    // 审核 c：FBA 仓校验（入仓编号不含「FBA」字眼 → 疑似非 FBA 仓）
    const nonFbaCodes = acc.fbaList.filter((f) => !isFbaCode(f));
    if (nonFbaCodes.length > 0) {
      nonFbaWarnings.push({ so: acc.so, fbaCodes: nonFbaCodes });
    }

    // 审核 a：按材质拆分易碎/非易碎
    const fragileRows = acc.rows.filter((r) => r.isFragile);
    const nonFragileRows = acc.rows.filter((r) => !r.isFragile);

    if (fragileRows.length > 0 && nonFragileRows.length > 0) {
      // 一票内同时含易碎与非易碎 → 拆两票：易碎 = -1，非易碎 = -2（顺序固定，与源行序无关）
      splitInfos.push({ so: acc.so });
      groups.push(buildGroup(acc, fragileRows, true, `${acc.so}-1`));
      groups.push(buildGroup(acc, nonFragileRows, false, `${acc.so}-2`));
    } else {
      const isFragile = fragileRows.length > 0;
      groups.push(buildGroup(acc, acc.rows, isFragile, acc.so));
    }
  }

  // 收尾：入仓编号 + 备注
  for (const g of groups) finalizeGroup(g);

  return { groups, currencyWarnings, nonFbaWarnings, splitInfos };
}

// ============================================================
// 模板填充
// ============================================================

/**
 * 将共享公式展平为独立公式/值，避免 spliceRows 删除行时破坏共享公式引用链。
 * 模板第 36 列「投保金额」以共享公式预填 AJ3:AJ499（分 8 段 si），必须先展平。
 */
function flattenFormulas(ws: ExcelJS.Worksheet): void {
  ws.eachRow({ includeEmpty: false }, (row) => {
    row.eachCell({ includeEmpty: false }, (cell) => {
      const v = cell.value;
      if (!v || typeof v !== "object") return;

      const obj = v as unknown as Record<string, unknown>;

      // 共享公式从格（sharedFormula）→ 只保留计算结果
      if ("sharedFormula" in obj) {
        cell.value = (obj.result ?? null) as ExcelJS.CellValue;
        return;
      }

      // 共享公式主格 → 转成普通公式
      const formula = obj.formula;
      if (typeof formula === "string" && ("shareType" in obj || "ref" in obj)) {
        cell.value = {
          formula,
          result: obj.result,
        } as unknown as ExcelJS.CellValue;
      }
    });
  });
}

/**
 * 将聚合结果写入模板。
 * 输出列顺序（36 列）与模板完全一致：
 *   被保险人 / 类型 / 证件号 / 原单号 / 干线运输方式 / 运输工具 / 提单号 / 派送方式 /
 *   目的地类型 / 快递公司 / 快递单号 / 入仓编号 / 上架保障 / 起运国 / 起运地 / 中转国 /
 *   中转地 / 目的国 / 目的地 / 赔付地 / 起运日期 / 货物类别 / 货物描述 / 包装类型 /
 *   总包装数量 / 总公斤数 / 保丢不保损 / 暴动保障 / 战争险保障 / 回邮附加责任 /
 *   送错地址附加责任 / 备注 / 币种 / 货值 / 加成比例 / 投保金额
 */
function fillTemplate(
  outWb: ExcelJS.Workbook,
  groups: SoGroup[],
  config: KuajingbaoConfig,
): void {
  const ws = outWb.getWorksheet(TEMPLATE_SHEET) || outWb.worksheets[0];
  if (!ws) throw new Error("模版中找不到「批量导入投保数据」工作表");

  const dataStartRow = 3;
  const shipDate = parseShipDate(config.shipDate);

  // 先展平共享公式，再逐格清空模板原有数据区（R3 起全部 36 列），保留 R1 表头 / R2 说明。
  // 用逐格置空而非 spliceRows：spliceRows 删除含公式的行会残留主格（exceljs 行位移 bug）。
  flattenFormulas(ws);
  for (let r = dataStartRow; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    for (let c = 1; c <= 36; c++) {
      row.getCell(c).value = null;
    }
  }

  groups.forEach((g, i) => {
    const rowNum = dataStartRow + i;
    const row = ws.getRow(rowNum);

    // 列 1~36，按模板顺序
    const vals: Array<string | number | Date | null> = [
      INSURED,                 // 1 被保险人
      TYPE,                    // 2 类型
      null,                    // 3 证件号
      g.orderNo,               // 4 原单号（拆分票带 -1/-2 后缀，保证唯一）
      config.transportMode,    // 5 干线运输方式
      g.orderNo,               // 6 运输工具名称及航次（与原单号一致）
      g.orderNo,               // 7 提单号（与原单号一致）
      config.deliveryMethod,   // 8 派送方式
      DEST_TYPE,               // 9 目的地类型
      null,                    // 10 快递公司
      null,                    // 11 快递单号
      g.fbaBase,               // 12 入仓编号
      SHELF_GUARANTEE,         // 13 上架保障
      ORIGIN_COUNTRY,          // 14 起运国
      config.originPlace,      // 15 起运地
      null,                    // 16 中转国
      null,                    // 17 中转地
      g.country,               // 18 目的国
      g.warehouseCode,         // 19 目的地
      CLAIM_PLACE,             // 20 赔付地
      shipDate,                // 21 起运日期
      g.goodsCategory,         // 22 货物类别（按材质自动判定）
      g.goodsDescription,      // 23 货物描述
      config.packageType,      // 24 包装类型
      g.totalBoxes,            // 25 总包装数量
      g.totalWeight,           // 26 总公斤数
      null,                    // 27 保丢不保损
      "否",                    // 28 暴动保障
      "是",                    // 29 战争险保障
      "否",                    // 30 回邮附加责任
      "否",                    // 31 送错地址附加责任
      g.remark || null,        // 32 备注
      g.currency,              // 33 币种
      g.totalValue,            // 34 货值
      RATIO,                   // 35 加成比例
      null,                    // 36 投保金额（公式，下方单独写入）
    ];

    vals.forEach((v, idx) => {
      if (v === null) return;
      row.getCell(idx + 1).value = v;
    });

    // 36 投保金额：保留模板公式 ROUND(货值,2)*加成比例（exceljs 公式串不带前导 =，
    // 写回 <f> 时原样输出，带 = 会导致读回时出现双 =）
    // 货值(34)/加成比例(35)/投保金额(36) 的数字格式由模板列样式提供，无需覆盖
    const amountCell = row.getCell(36);
    amountCell.value = {
      formula: `ROUND(AH${rowNum},2)*AI${rowNum}`,
      result: round2(g.totalValue * RATIO),
    } as unknown as ExcelJS.CellValue;
  });
}

// ============================================================
// 主入口
// ============================================================

export async function processKuajingbaoInsurance(
  filePath: string,
  sourceFileName: string,
  config: KuajingbaoConfig,
): Promise<KuajingbaoResult> {
  // ── 1. 解析货箱清单（含审核）──
  const srcWb = new ExcelJS.Workbook();
  await srcWb.xlsx.readFile(filePath);
  const srcSheet = srcWb.worksheets[0];
  if (!srcSheet) throw new Error("货箱清单中没有找到工作表");

  const { groups, currencyWarnings, nonFbaWarnings, splitInfos } = parseCargoList(srcSheet);

  // ── 2. 加载模板并填充 ──
  const tplPath = path.join(process.cwd(), "public", "templates", "跨境堡批量投保箱单模版.xlsx");
  if (!fs.existsSync(tplPath)) {
    throw new Error(`跨境堡投保模版不存在: ${tplPath}`);
  }
  const outWb = new ExcelJS.Workbook();
  await outWb.xlsx.readFile(tplPath);

  fillTemplate(outWb, groups, config);

  const buffer = Buffer.from(await outWb.xlsx.writeBuffer());
  const lockWarehouseCount = groups.filter((g) => g.remark).length;

  console.log(
    `[跨境堡] ${sourceFileName}: ${groups.length} 票（全美锁仓 ${lockWarehouseCount} 票带备注；` +
    `币种异常 ${currencyWarnings.length}；非FBA仓 ${nonFbaWarnings.length}；拆分 ${splitInfos.length}）`
  );

  return {
    sourceFile: sourceFileName,
    groups,
    lockWarehouseCount,
    currencyWarnings,
    nonFbaWarnings,
    splitInfos,
    buffer,
  };
}

/** 输出文件名：<原文件名去扩展名>_跨境堡投保箱单.xlsx */
export function kuajingbaoOutputName(sourceFileName: string): string {
  const base = sanitizeFileName(sourceFileName.replace(/\.(xlsx?|xls)$/i, ""));
  return `${base}_跨境堡投保箱单.xlsx`;
}
