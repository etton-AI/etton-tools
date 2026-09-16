/**
 * 跨境堡批量投保 — 将「系统导出的货箱清单」按 SO 号分组，填充「跨境堡批量投保箱单模版」。
 *
 * 数据流：
 *   货箱清单（Sheet1，标题行 + 表头 + 数据行）
 *     ├─ 系统SO（仅每组首行，需向下填充）
 *     ├─ 国家 / 仓库代码 / FBA ID / 中文品名 / 申报总数量 / 总申报货值 / 申报币种
 *     ├─ 总箱数(CTN) / 货箱重量 / 客户渠道
 *       ↓
 *   跨境堡批量投保箱单模版（public/templates/跨境堡批量投保箱单模版.xlsx）
 *     ├─ R1 表头 / R2 说明
 *     └─ R3 起数据区，36 列，逐票（SO）一行
 *
 * 核心计算的 4 个字段：
 *   1. 货物描述：按 SO 聚合「中文品名 + 申报总数量」→ `品名+数量PCS`，多品名以「、」连接
 *   2. 货值：总申报货值求和（四舍五入 2 位）
 *   3. 币种：申报币种 → 中文名（USD → 美元）
 *   4. 备注：仅「全美锁仓」渠道（客户渠道含「全美」）需要，仓库代码替换 AAAA
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
  /** 货物类别，默认 无易碎品 */
  goodsCategory: string;
  /** 包装类型，默认 纸箱 */
  packageType: string;
}

/** 一票（一个 SO）的聚合结果 */
export interface SoGroup {
  so: string;              // 系统SO → 原单号/运输工具/提单号
  workOrder: string;       // 工作号（仅用于展示/排查）
  country: string;         // 国家 → 目的国
  warehouseCode: string;   // 仓库代码 → 目的地
  fbaBase: string;         // FBA ID 去箱序 → 入仓编号
  channel: string;         // 客户渠道（判断是否全美锁仓）
  goodsDescription: string;// 货物描述
  currency: string;        // 币种（中文）
  totalValue: number;      // 货值
  totalBoxes: number;      // 总包装数量
  totalWeight: number;     // 总公斤数
  remark: string;          // 备注（非全美锁仓为空）
}

export interface KuajingbaoResult {
  sourceFile: string;
  groups: SoGroup[];
  lockWarehouseCount: number; // 全美锁仓（含备注）票数
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
}

// ============================================================
// 常量
// ============================================================

/** 跨境堡模板主 sheet 名 */
const TEMPLATE_SHEET = "批量导入投保数据";

/** 模板内置常量（与示例数据一致） */
const INSURED = "易通科技物流（中山）有限公司";
const TYPE = "企业";
const ORIGIN_COUNTRY = "中国";
const CLAIM_PLACE = "HANGZHOU CHINA";
const DEST_TYPE = "FBA";
const SHELF_GUARANTEE = "保上架";
const RATIO = 1; // 加成比例 100%

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
    "boxWeight", "boxCount", "declaredQty", "declaredValue", "currency", "channel",
  ];
  const missing = required.filter((k) => map[k] === undefined);
  if (missing.length > 0) {
    throw new Error(`货箱清单表头缺少以下列: ${missing.join("、")}。请确认上传的是系统导出的货箱清单。`);
  }
  return map as CargoColumnMap;
}

/** 解析并聚合：按 SO 号（向下填充）分组 */
function parseCargoList(ws: ExcelJS.Worksheet): SoGroup[] {
  const headerRow = findCargoHeaderRow(ws);
  const col = buildCargoColumnMap(ws.getRow(headerRow));

  // 按 SO 分组，保持出现顺序
  const groups: SoGroup[] = [];
  const indexBySo = new Map<string, number>();
  // 每个 SO 内品名 → 数量的累加（保持首次出现顺序）
  const qtyBySo = new Map<string, Map<string, number>>();

  let currentSo = "";

  for (let r = headerRow + 1; r <= ws.rowCount; r++) {
    const soCell = cellText(ws.getCell(r, col.so));
    if (soCell) currentSo = soCell; // 首行带 SO，续行向下填充
    if (!currentSo) continue; // 表头下方尚未遇到首个 SO

    const nameCh = cellText(ws.getCell(r, col.nameCh));
    const qty = cellNum(ws.getCell(r, col.declaredQty));

    if (!indexBySo.has(currentSo)) {
      indexBySo.set(currentSo, groups.length);
      qtyBySo.set(currentSo, new Map());
      groups.push({
        so: currentSo,
        workOrder: cellText(ws.getCell(r, col.workOrder)),
        country: cellText(ws.getCell(r, col.country)),
        warehouseCode: cellText(ws.getCell(r, col.warehouseCode)),
        fbaBase: "",
        channel: cellText(ws.getCell(r, col.channel)),
        goodsDescription: "",
        currency: currencyToCn(cellText(ws.getCell(r, col.currency))),
        totalValue: 0,
        totalBoxes: 0,
        totalWeight: 0,
        remark: "",
      });
    }

    const g = groups[indexBySo.get(currentSo)!];
    const qtyMap = qtyBySo.get(currentSo)!;

    // 品名数量累加（空品名跳过，避免把续行的空品名误计）
    if (nameCh) {
      qtyMap.set(nameCh, (qtyMap.get(nameCh) || 0) + qty);
    }

    // 数值聚合
    g.totalValue += cellNum(ws.getCell(r, col.declaredValue));
    g.totalBoxes += cellNum(ws.getCell(r, col.boxCount));
    g.totalWeight += cellNum(ws.getCell(r, col.boxWeight));

    // 入仓编号：首个 FBA ID 去掉箱序后缀（U+6 位数字）
    if (!g.fbaBase) {
      g.fbaBase = cellText(ws.getCell(r, col.fbaId)).replace(/U\d{6}$/, "");
    }

    // 币种/渠道：取非空值兜底
    const cur = cellText(ws.getCell(r, col.currency));
    if (cur && !g.currency) g.currency = currencyToCn(cur);
    const ch = cellText(ws.getCell(r, col.channel));
    if (ch && !g.channel) g.channel = ch;
  }

  if (groups.length === 0) {
    throw new Error("货箱清单中未找到有效的 SO 数据行。");
  }

  // 收尾：生成货物描述 / 备注 / 四舍五入货值
  for (const g of groups) {
    const qtyMap = qtyBySo.get(g.so)!;
    const parts: string[] = [];
    for (const [name, q] of qtyMap) {
      parts.push(`${name}+${q}PCS`);
    }
    g.goodsDescription = parts.join("、");
    g.totalValue = round2(g.totalValue);
    g.totalWeight = round2(g.totalWeight);
    // 备注：仅「全美锁仓」渠道（客户渠道含「全美」）
    if (g.channel.includes("全美")) {
      const code = g.warehouseCode;
      g.remark =
        `本保单承保AGL（亚马逊物流）承运的物流运输服务，其亚马逊货件建仓地址为“${code}”，` +
        `投保的目的地为“${code}”，同时存在最后FBA调仓可能${code}`;
    }
  }

  return groups;
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
      g.so,                    // 4 原单号
      config.transportMode,    // 5 干线运输方式
      g.so,                    // 6 运输工具名称及航次
      g.so,                    // 7 提单号
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
      config.goodsCategory,    // 22 货物类别
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
  // ── 1. 解析货箱清单 ──
  const srcWb = new ExcelJS.Workbook();
  await srcWb.xlsx.readFile(filePath);
  const srcSheet = srcWb.worksheets[0];
  if (!srcSheet) throw new Error("货箱清单中没有找到工作表");

  const groups = parseCargoList(srcSheet);

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
    `[跨境堡] ${sourceFileName}: ${groups.length} 票（全美锁仓 ${lockWarehouseCount} 票带备注）`
  );

  return {
    sourceFile: sourceFileName,
    groups,
    lockWarehouseCount,
    buffer,
  };
}

/** 输出文件名：<原文件名去扩展名>_跨境堡投保箱单.xlsx */
export function kuajingbaoOutputName(sourceFileName: string): string {
  const base = sanitizeFileName(sourceFileName.replace(/\.(xlsx?|xls)$/i, ""));
  return `${base}_跨境堡投保箱单.xlsx`;
}
