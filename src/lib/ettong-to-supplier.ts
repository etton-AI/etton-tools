/**
 * 易通下单模板 → 供应商模板 转换（朗胜 LS / 延讯 YX）
 *
 * 两个大客户共用同一种「易通下单模板」作为源（sheet: "ETTON电商物流 下单模板"），
 * 分别转换成两家供应商的下单模板：
 *   - 朗胜 LS：易通 → 「美线发票下单映射」；按 SKU 拆分下单（一个 SKU 一个 SO），
 *     依赖上传的「SO-SKU对应表」补 SO 号，输出多张发票打包成 ZIP。
 *   - 延讯 YX：易通标准模板 → 供应商发票（天图/星链/皓鹏/美琦/安邦）；
 *     支持单张或批量（批量须同国家），渠道名从供应商模板解析后下拉选择。
 *
 * 易通源顶部字段（固定坐标）：
 *   B6 业务类型 / B7 服务渠道 / B8 报关方式 / B9 是否合并报关 / B11 清关方式 /
 *   B18 总箱数 / B19 预计总重量 / B20 预计总体积 / B21 备注
 *   F2 带电 / F3 带磁 / F7 仓点类型 / F9 收件人国家 / F10 仓库代码 / F11 仓库地址 / F12 邮编
 * 易通源明细（R25 起，A 列空结束）：
 *   A 货箱编号 / B 货件追踪编码 / C 英文品名 / D 中文品名 / E 材质 / F 用途 / G 箱数 /
 *   H 总数量 / I 单价 / J 总价 / K 币种 / L 海关编码 / M 品牌 / N 型号 / O 品牌类型 /
 *   P 净重 / Q 毛重 / R 长 / S 宽 / T 高 / U 链接 / V 图片 / Z SKU / AA 地址
 */

import ExcelJS from "exceljs";
import JSZip from "jszip";
import path from "path";
import fs from "fs";

// ============================================================
// Types
// ============================================================

/** 易通源订单头 */
export interface EttongTopInfo {
  bizType: string;          // B6 业务类型
  channel: string;          // B7 服务渠道
  customs: string;          // B8 报关方式
  mergeDeclare: string;     // B9 是否合并报关
  clearance: string;        // B11 清关方式
  totalBoxes: number;       // B18 总箱数
  totalWeight: number;      // B19 预计总重量
  totalVolume: number;      // B20 预计总体积
  remark: string;           // B21 备注
  hasBattery: string;       // F2 带电
  hasMagnet: string;        // F3 带磁
  warehouseType: string;    // F7 仓点类型
  country: string;          // F9 收件人国家
  warehouseCode: string;    // F10 仓库代码
  warehouseAddress: string; // F11 仓库地址
  zipCode: string;          // F12 邮编
}

/** 一行易通源明细 */
export interface EttongItem {
  boxId: string;       // A 货箱编号
  refId: string;       // B 货件追踪编码(Reference ID)
  nameEn: string;      // C 英文品名
  nameCh: string;      // D 中文品名
  material: string;    // E 材质
  use: string;         // F 用途
  boxCount: number;    // G 箱数
  totalQty: number;    // H 总数量
  unitPrice: number;   // I 单价
  totalPrice: number;  // J 总价
  currency: string;    // K 币种
  hsCode: string;      // L 海关编码
  brand: string;       // M 品牌
  model: string;       // N 型号
  brandType: string;   // O 品牌类型
  netWeight: number;   // P 净重
  grossWeight: number; // Q 毛重
  lengthCm: number;    // R 长
  widthCm: number;     // S 宽
  heightCm: number;    // T 高
  link: string;        // U 链接
  image: string;       // V 图片
  sku: string;         // Z SKU
  address: string;     // AA 地址
}

/** LS 转换结果（一个 SO 一张发票，整体打包成 ZIP） */
export interface LsFileEntry {
  so: string;           // 客户订单号(SO)
  sku: string;          // 对应 SKU
  fileName: string;     // SO.xlsx
  warnings: string[];
}

export interface LsConvertResult {
  sourceFile: string;
  zipFileName: string;
  files: LsFileEntry[];
  skipped: { sku: string; reason: string }[];
  warnings: string[];
  buffer: Buffer;       // zip buffer
}

/** YX 转换结果 */
export interface YxConvertResult {
  sourceFile: string;
  fileName: string;
  supplier: YxSupplierKey;
  customerOrderNo: string;
  warehouseCode: string;
  totalBoxes: number;
  dataRows: number;
  warnings: string[];
  buffer: Buffer;
}

/** YX 批量转换结果 */
export interface YxBatchFile {
  sourceFile: string;
  fileName: string;
  country: string;
  warehouseCode: string;
  totalBoxes: number;
  dataRows: number;
}

export interface YxBatchResult {
  sourceFiles: string[];
  zipFileName: string;
  supplier: YxSupplierKey;
  country: string;
  files: YxBatchFile[];
  warnings: string[];
  buffer: Buffer;
}

// ============================================================
// Helpers
// ============================================================

function cellText(cell: ExcelJS.Cell): string {
  const v = cell.value;
  if (v == null) return "";
  if (typeof v === "string") return v.trim();
  if (typeof v === "number") return String(v);
  if (typeof v === "object") {
    const o = v as { richText?: { text: string }[]; text?: unknown; result?: unknown; error?: unknown };
    if (o.richText) return o.richText.map((r) => r.text).join("").trim();
    if (o.error !== undefined && o.error !== null) return ""; // 公式错误(#REF! 等)按空处理
    if (o.result !== undefined && o.result !== null) {
      const r = o.result;
      if (typeof r === "string") return r.trim();
      if (typeof r === "number") return String(r);
      if (typeof r === "boolean") return String(r);
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

/** 四舍五入到 2 位小数 */
function r2(n: number): number {
  return Math.round(n * 100) / 100;
}

function sanitizeFileName(name: string): string {
  const cleaned = name.replace(/[\\/:*?"<>|]/g, "_").trim();
  return cleaned || "未命名";
}

/** 规范化发票文件名（去扩展名/副本标记/空格），用于匹配 SO-SKU 表 A 列 */
function normalizeInvoiceName(name: string): string {
  return name
    .replace(/\.(xlsx|xls)$/i, "")
    .replace(/\(\d+\)/g, "")
    .replace(/\s+/g, "")
    .trim();
}

/**
 * 提取「日期 + 航司/仓代码」作为稳定匹配键（仅 ASCII）。
 *
 * 上传经过浏览器/curl 时，文件名里的中文（如「易通」）可能被按本地代码页编码成乱码，
 * 导致 rawFile.name 与 SO-SKU 表 A 列里的 UTF-8 中文不一致、整名匹配失败。
 * 而日期与代码（如 2026-10-15 LGB4）是纯 ASCII，不受编码影响，用作匹配键最稳。
 */
function invoiceKey(name: string): string {
  const m = name.match(/(\d{4}-\d{1,2}-\d{1,2})\s*([A-Z]{2,5}\d{1,2})/);
  return m ? m[1] + m[2] : "";
}

/** 今日 Excel 日期序列值（用于 LS 美线发票「预计提货时间」） */
function todayExcelSerial(): number {
  const now = new Date();
  const utc = Date.UTC(now.getFullYear(), now.getMonth(), now.getDate());
  const epoch = Date.UTC(1899, 11, 30);
  return Math.round((utc - epoch) / 86400000);
}

/** 写入「预计提货时间」：今天日期序列 + 日期格式，避免导出后显示成一串数字 */
function writePickupDate(cell: ExcelJS.Cell): void {
  cell.value = todayExcelSerial();
  cell.numFmt = "yyyy-mm-dd";
}

/** 海关编码：纯数字转数值（美线模板 J 列为数值格式） */
function hsCodeForLs(hsCode: string): string | number {
  return /^\d+$/.test(hsCode) ? Number(hsCode) : hsCode;
}

// ============================================================
// 易通源解析（共享）
// ============================================================

function parseTopInfo(ws: ExcelJS.Worksheet): EttongTopInfo {
  const s = (ref: string) => cellText(ws.getCell(ref));
  return {
    bizType: s("B6"),
    channel: s("B7"),
    customs: s("B8"),
    mergeDeclare: s("B9"),
    clearance: s("B11"),
    totalBoxes: cellNum(ws.getCell("B18")),
    totalWeight: cellNum(ws.getCell("B19")),
    totalVolume: cellNum(ws.getCell("B20")),
    remark: s("B21"),
    hasBattery: s("F2"),
    hasMagnet: s("F3"),
    warehouseType: s("F7"),
    country: s("F9"),
    warehouseCode: s("F10"),
    warehouseAddress: s("F11"),
    zipCode: s("F12"),
  };
}

function parseItems(ws: ExcelJS.Worksheet): EttongItem[] {
  const items: EttongItem[] = [];
  for (let r = 25; r <= ws.rowCount; r++) {
    const boxId = cellText(ws.getCell(r, 1));
    if (!boxId) break; // 货箱编号为空 = 结束
    items.push({
      boxId,
      refId: cellText(ws.getCell(r, 2)),
      nameEn: cellText(ws.getCell(r, 3)),
      nameCh: cellText(ws.getCell(r, 4)),
      material: cellText(ws.getCell(r, 5)),
      use: cellText(ws.getCell(r, 6)),
      boxCount: cellNum(ws.getCell(r, 7)),
      totalQty: cellNum(ws.getCell(r, 8)),
      unitPrice: cellNum(ws.getCell(r, 9)),
      totalPrice: cellNum(ws.getCell(r, 10)),
      currency: cellText(ws.getCell(r, 11)),
      hsCode: cellText(ws.getCell(r, 12)),
      brand: cellText(ws.getCell(r, 13)),
      model: cellText(ws.getCell(r, 14)),
      brandType: cellText(ws.getCell(r, 15)),
      netWeight: cellNum(ws.getCell(r, 16)),
      grossWeight: cellNum(ws.getCell(r, 17)),
      lengthCm: cellNum(ws.getCell(r, 18)),
      widthCm: cellNum(ws.getCell(r, 19)),
      heightCm: cellNum(ws.getCell(r, 20)),
      link: cellText(ws.getCell(r, 21)),
      image: cellText(ws.getCell(r, 22)),
      sku: cellText(ws.getCell(r, 26)),   // Z
      address: cellText(ws.getCell(r, 27)), // AA
    });
  }
  return items;
}

async function parseEttongSource(
  filePath: string
): Promise<{ top: EttongTopInfo; items: EttongItem[] }> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(filePath);
  const ws = wb.getWorksheet("ETTON电商物流 下单模板");
  if (!ws) throw new Error("找不到「ETTON电商物流 下单模板」工作表，请确认上传的是易通下单模板");
  const top = parseTopInfo(ws);
  const items = parseItems(ws);
  return { top, items };
}

// ============================================================
// SO-SKU 对应表解析（朗胜拆分用）
// ============================================================

interface SoSkuTable {
  /** 发票文件名(规范化) → [{so, sku}]（A 列为合并单元格，跨行沿用当前分组） */
  groups: Map<string, { so: string; sku: string }[]>;
  /** 稳定键(日期+代码) → [{so, sku}]（文件名中文乱码时的 fallback） */
  groupsByKey: Map<string, { so: string; sku: string }[]>;
  /** SKU → [so...]（全局，文件名匹配失败时的 fallback） */
  skuGlobal: Map<string, string[]>;
}

async function parseSoSkuTable(filePath: string): Promise<SoSkuTable> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(filePath);
  const ws = wb.getWorksheet("1") || wb.worksheets[0];
  if (!ws) throw new Error("SO-SKU 对应表中找不到数据工作表");

  const groups = new Map<string, { so: string; sku: string }[]>();
  const groupsByKey = new Map<string, { so: string; sku: string }[]>();
  const skuGlobal = new Map<string, string[]>();

  let currentGroup = "";
  let currentKey = "";
  for (let r = 2; r <= ws.rowCount; r++) {
    const invoice = cellText(ws.getCell(r, 1));
    if (invoice) {
      currentGroup = normalizeInvoiceName(invoice);
      currentKey = invoiceKey(invoice) || currentGroup;
    }

    const so = cellText(ws.getCell(r, 2));
    const sku = cellText(ws.getCell(r, 3));
    if (!sku && !so) continue;

    if (so && sku) {
      const arr = skuGlobal.get(sku) || [];
      arr.push(so);
      skuGlobal.set(sku, arr);
    }
    if (currentGroup && so && sku) {
      const arr = groups.get(currentGroup) || [];
      arr.push({ so, sku });
      groups.set(currentGroup, arr);

      const arrKey = groupsByKey.get(currentKey) || [];
      arrKey.push({ so, sku });
      groupsByKey.set(currentKey, arrKey);
    }
  }

  return { groups, groupsByKey, skuGlobal };
}

// ============================================================
// LS 朗胜：易通 → 美线发票下单映射
// ============================================================

/** 报关方式映射：需要退税(报关退税)→一般贸易，其余→代理报关 */
function lsCustomsFor(customs: string): string {
  return customs === "报关退税" ? "一般贸易" : "代理报关";
}

/**
 * 生成单张美线发票。item 为一行源明细（一个 SKU），so 为该 SKU 对应的 SO 号。
 */
async function buildLsInvoice(
  top: EttongTopInfo,
  item: EttongItem,
  so: string,
  opts: { channelName: string; pickupMethod: string; deliverySite: string },
  boxRangeStart: number
): Promise<{ fileName: string; buffer: Buffer }> {
  const tplPath = path.join(process.cwd(), "public", "templates", "ls-yx", "美线发票下单映射-模板.xlsx");
  if (!fs.existsSync(tplPath)) throw new Error(`美线发票模板不存在: ${tplPath}`);

  const outWb = new ExcelJS.Workbook();
  await outWb.xlsx.readFile(tplPath);
  const ws = outWb.getWorksheet("发票");
  if (!ws) throw new Error("美线发票模板中找不到「发票」工作表");

  // ── 顶部字段 ──
  ws.getCell(1, 2).value = so;                    // B1 客户订单号
  ws.getCell(2, 2).value = so;                    // B2 客户参考号
  ws.getCell(3, 2).value = opts.channelName;      // B3 服务
  ws.getCell(4, 2).value = top.warehouseCode;     // B4 地址库编码(触发 VLOOKUP 带出地址)
  ws.getCell(14, 2).value = top.hasBattery;       // B14 带电
  ws.getCell(15, 2).value = top.hasMagnet;        // B15 带磁
  ws.getCell(16, 2).value = lsCustomsFor(top.customs); // B16 报关方式
  ws.getCell(17, 2).value = item.boxCount;        // B17 箱数

  ws.getCell(1, 5).value = opts.deliverySite;     // E1 预计交货站点
  writePickupDate(ws.getCell(2, 5));              // E2 预计提货时间(今天)
  ws.getCell(3, 5).value = opts.pickupMethod;     // E3 提货方式
  ws.getCell(5, 5).value = "无";                   // E5 涉证类型
  ws.getCell(7, 5).value = "无";                   // E7 出证方式

  // ── 数据区（先清空模板示例数据行）──
  const dataStart = 19;
  const rowsToDelete = ws.actualRowCount - dataStart + 1;
  if (rowsToDelete > 0) ws.spliceRows(dataStart, rowsToDelete);

  const perBox = item.boxCount ? r2(item.totalQty / item.boxCount) : 0;
  const boxStart = boxRangeStart + 1;
  const boxEnd = boxRangeStart + item.boxCount;
  const row = ws.getRow(dataStart);
  const values: (string | number)[] = [
    `${boxStart}-${boxEnd}`, // A 货箱编号(累计区间)
    item.nameEn,                            // B 英文品名
    item.nameCh,                            // C 中文品名
    perBox,                                 // D 单箱数量
    item.unitPrice,                         // E 单价
    item.grossWeight,                       // F 货箱重量(毛重原值)
    r2(item.lengthCm - 1),                  // G 长 -1
    r2(item.widthCm - 1),                   // H 宽 -1
    r2(item.heightCm - 1),                  // I 高 -1
    hsCodeForLs(item.hsCode),               // J 海关编码
    item.brand,                             // K 品牌
    item.material,                          // L 材质
    item.model,                             // M 型号
    item.use,                               // N 用途
    0,                                      // O 产品图片(固定 0)
    item.refId,                             // P PO Number
    item.boxCount,                          // Q 物品箱号
    item.boxId,                             // R 物品 FBA ID
  ];
  values.forEach((v, ci) => {
    if (v !== null && v !== undefined && v !== "") row.getCell(ci + 1).value = v;
  });
  for (let c = 1; c <= 18; c++) {
    row.getCell(c).border = {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "thin" },
    };
  }

  const buffer = Buffer.from(await outWb.xlsx.writeBuffer());
  const fileName = `${sanitizeFileName(so || "未命名")}.xlsx`;
  return { fileName, buffer };
}

export async function convertLsToSupplier(
  filePath: string,
  sourceFileName: string,
  soSkuPath: string,
  opts: { channelName: string; pickupMethod: string; deliverySite: string }
): Promise<LsConvertResult> {
  // ── 1. 解析易通源 ──
  const { top, items } = await parseEttongSource(filePath);
  if (items.length === 0) throw new Error("易通下单模板中未找到有效货箱数据");

  // ── 2. 解析 SO-SKU 对应表 ──
  const soSku = await parseSoSkuTable(soSkuPath);

  // ── 3. 匹配：优先按发票文件名分组，其次按日期+代码稳定键，fallback 按 SKU 全局匹配 ──
  const srcKey = normalizeInvoiceName(sourceFileName);
  const srcInvoiceKey = invoiceKey(sourceFileName);
  let pairList = soSku.groups.get(srcKey);
  if (!pairList && srcInvoiceKey) {
    pairList = soSku.groupsByKey.get(srcInvoiceKey);
  }

  const warnings: string[] = [];
  const files: LsFileEntry[] = [];
  const skipped: { sku: string; reason: string }[] = [];
  const zip = new JSZip();

  if (pairList) {
    // 按文件名匹配到分组：为每个 SKU 建立 SO 队列（保持顺序，支持同一 SKU 多个 SO）
    const queue = new Map<string, string[]>();
    for (const p of pairList) {
      if (!p.sku) continue;
      const arr = queue.get(p.sku) || [];
      arr.push(p.so);
      queue.set(p.sku, arr);
    }

    let cumulative = 0;
    let currentGroup = ""; // 当前 FBA 分组键（同一货件追踪编码连续编号，换单号重新从 1 起）
    for (const item of items) {
      const sku = item.sku;
      if (!sku) {
        skipped.push({ sku: "", reason: "该行无 SKU" });
        continue;
      }
      const q = queue.get(sku);
      const so = q && q.length ? q.shift()! : "";
      if (!so) {
        skipped.push({ sku, reason: "SO-SKU 表中无对应 SO" });
        continue;
      }
      const gk = item.refId || item.boxId;
      if (gk !== currentGroup) {
        currentGroup = gk;
        cumulative = 0; // 换 FBA 货件，箱号重新从 1 起
      }
      const { fileName, buffer } = await buildLsInvoice(top, item, so, opts, cumulative);
      cumulative += item.boxCount;
      zip.file(fileName, buffer);
      files.push({ so, sku, fileName, warnings: [] });
    }
  } else {
    // fallback：文件名匹配不到（如用户改过名），按 SKU 全局匹配（唯一才用）
    warnings.push(
      `未在 SO-SKU 表中找到发票「${sourceFileName}」的映射，已按 SKU 全局匹配`
    );
    const usedSos = new Set<string>();
    let cumulative = 0;
    let currentGroup = ""; // 当前 FBA 分组键（同一货件追踪编码连续编号，换单号重新从 1 起）
    for (const item of items) {
      const sku = item.sku;
      if (!sku) {
        skipped.push({ sku: "", reason: "该行无 SKU" });
        continue;
      }
      const sos = soSku.skuGlobal.get(sku) || [];
      // 优先选尚未用过的 SO（避免同一 SKU 多 SO 时重复使用同一个）
      const so = sos.find((s) => !usedSos.has(s)) || sos[0] || "";
      if (!so) {
        skipped.push({ sku, reason: "SO-SKU 表中无对应 SO" });
        continue;
      }
      if (sos.length > 1) {
        warnings.push(`SKU「${sku}」在 SO-SKU 表中对应多个 SO(${sos.join("/")})，已取 ${so}`);
      }
      usedSos.add(so);
      const gk = item.refId || item.boxId;
      if (gk !== currentGroup) {
        currentGroup = gk;
        cumulative = 0; // 换 FBA 货件，箱号重新从 1 起
      }
      const { fileName, buffer } = await buildLsInvoice(top, item, so, opts, cumulative);
      cumulative += item.boxCount;
      zip.file(fileName, buffer);
      files.push({ so, sku, fileName, warnings: [] });
    }
  }

  if (files.length === 0) {
    throw new Error("未能生成任何美线发票：源明细的 SKU 与 SO-SKU 对应表均无法匹配");
  }

  // ── 4. 汇总 json 一并打包 ──
  zip.file(
    "转换结果汇总.json",
    JSON.stringify(
      {
        sourceFile: sourceFileName,
        total: files.length,
        skipped,
        warnings,
        files: files.map((f) => ({ so: f.so, sku: f.sku, fileName: f.fileName })),
      },
      null,
      2
    )
  );

  const buffer = Buffer.from(await zip.generateAsync({ type: "nodebuffer" }));
  const base = normalizeInvoiceName(sourceFileName) || "朗胜美线发票";
  return {
    sourceFile: sourceFileName,
    zipFileName: `${base}-美线发票.zip`,
    files,
    skipped,
    warnings,
    buffer,
  };
}

// ============================================================
// YX 延讯：易通标准模板 → 天图
// ============================================================

/** 收件人国家代码映射：中文国名 → 二字代码 */
const COUNTRY_CODE: Record<string, string> = {
  加拿大: "CA",
  美国: "US",
  英国: "GB",
  德国: "DE",
  法国: "FR",
  西班牙: "ES",
  意大利: "IT",
  日本: "JP",
  墨西哥: "MX",
  澳大利亚: "AU",
  荷兰: "NL",
  比利时: "BE",
};

function countryCode(country: string): string {
  return COUNTRY_CODE[country] || country;
}

/** 交税方式：清关方式「无」→「包税」 */
function yxTaxFor(clearance: string): string {
  return clearance === "无" || !clearance ? "包税" : clearance;
}

/** 报关方式映射：普通报关→托管报关，其余(报关退税等)透传 */
function yxCustomsFor(customs: string): string {
  return customs === "普通报关" ? "托管报关" : customs;
}

/** 链接/图片列：源为 0 或空时输出数值 0，否则输出原值 */
function urlOrZero(v: string): string | number {
  return !v || v === "0" ? 0 : v;
}

/** 对齐 Excel 的 15 位有效数字存储精度，消除源文件 J 列总价的浮点尾差（如 25.759999999999952 → 25.76） */
function excelPrec(v: number): number {
  return Number(v.toPrecision(15));
}

/** YX 供应商标识（安邦暂未提供模板，ready=false） */
export type YxSupplierKey = "天图" | "星链" | "皓鹏" | "美琦" | "安邦";

export interface YxConvertOptions {
  customerOrderNo: string; // 可空（批量时留空，按 FBA 编号命名）
  channelName: string;
}

/** 用户输入类错误（参数/校验），API 层映射为 400 */
export class YxValidationError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "YxValidationError";
  }
}

/** 星链「报关类型」映射：易通报关方式 → 星链不需/单独/合并报关 */
function yxXinglianCustomsFor(customs: string): string {
  if (customs === "无单证") return "不需报关";
  if (customs === "报关退税" || customs === "普通报关") return "单独报关";
  if (customs.includes("合并")) return "合并报关（两票及以上需要合并单独报关的）";
  return customs;
}

/** 皓鹏「报关方式」映射：易通报关方式 → 皓鹏单证/非单证报关 */
function yxHaopengCustomsFor(customs: string): string {
  if (customs === "报关退税") return "单证报关";
  if (customs === "普通报关") return "非单证报关";
  return customs;
}

/** 申报币种：ISO 三位码 → 中文名（皓鹏 E7 用） */
const YX_CURRENCY: Record<string, string> = {
  USD: "美元", CNY: "人民币", EUR: "欧元", GBP: "英镑",
  JPY: "日元", AUD: "澳元", CAD: "加元", HKD: "港币", MXN: "墨西哥比索",
};
function yxCurrencyFor(currency: string): string {
  return YX_CURRENCY[(currency || "").toUpperCase()] || currency || "";
}

/** 皓鹏货箱编号前缀：订单号优先，留空则取首个货箱编号去掉 Uxxxx 后缀 */
function yxBoxPrefix(orderNo: string, items: EttongItem[]): string {
  if (orderNo) return orderNo;
  const first = items[0]?.boxId || "";
  return first.replace(/U\d+$/i, "");
}

interface YxMapContext {
  index: number;            // 0-based 数据行序号
  cumulativeStart: number;  // 本行之前已累计的箱数（美琦 A 列累计区间用）
  boxPrefix: string;        // 皓鹏货箱编号前缀
}

interface YxSupplierSpec {
  key: YxSupplierKey;
  templateFile: string;    // 相对 public/templates/suppliers/
  sheetName: string;
  dataStartRow: number;    // 数据区起始行（表头下一行）
  maxCol: number;
  ready: boolean;          // 安邦=false
  fillHeader(ws: ExcelJS.Worksheet, top: EttongTopInfo, opts: YxConvertOptions, items: EttongItem[]): void;
  mapItem(item: EttongItem, top: EttongTopInfo, ctx: YxMapContext): (string | number)[];
  channelSheet?: { name: string; col: number; skipHeader?: boolean };
  /** 货箱编号累计区间按此键分组重置（美琦用：同一个 FBA 货件的单号连续累加，换单号重新从 1 起） */
  groupKey?: (item: EttongItem) => string;
}

const YX_SUPPLIERS: Record<YxSupplierKey, YxSupplierSpec> = {
  // ── 天图：B1 服务 / B2 仓库 / B11 国家 / B14 订单号 / B16 带电 / B17 带磁 / B18 交税 / B19 报关 / B22 总箱数 ──
  天图: {
    key: "天图",
    templateFile: "天图-模板.xlsx",
    sheetName: "Sheet1",
    dataStartRow: 30,
    maxCol: 18,
    ready: true,
    fillHeader(ws, top, opts) {
      ws.getCell(1, 2).value = opts.channelName;                 // B1 服务
      ws.getCell(2, 2).value = top.warehouseCode;                // B2 仓库代码(触发 VLOOKUP 带出地址)
      ws.getCell(11, 2).value = countryCode(top.country);        // B11 收件人国家代码
      ws.getCell(14, 2).value = opts.customerOrderNo;            // B14 客户订单号
      if (top.hasBattery) ws.getCell(16, 2).value = top.hasBattery; // B16 带电
      if (top.hasMagnet) ws.getCell(17, 2).value = top.hasMagnet;   // B17 带磁
      ws.getCell(18, 2).value = yxTaxFor(top.clearance);         // B18 交税方式
      ws.getCell(19, 2).value = yxCustomsFor(top.customs);       // B19 报关方式
      ws.getCell(22, 2).value = top.totalBoxes;                  // B22 总箱数
    },
    mapItem(it) {
      return [
        it.boxId, it.refId, it.nameEn, it.nameCh,
        excelPrec(it.unitPrice), it.totalQty, excelPrec(it.totalPrice),
        it.material, it.hsCode, it.use, it.brand, it.model,
        urlOrZero(it.image), urlOrZero(it.link),
        it.grossWeight, it.lengthCm, it.widthCm, it.heightCm,
      ];
    },
    channelSheet: { name: "Sheet2", col: 1 },
  },

  // ── 星链：sheet 美中北（多仓库同名结构，默认取美中北），数据 20 列 ──
  星链: {
    key: "星链",
    templateFile: "模板转换/星链-美国模板.xlsx",
    sheetName: "美中北",
    dataStartRow: 18,
    maxCol: 20,
    ready: true,
    fillHeader(ws, top, opts, items) {
      const first = items[0];
      ws.getCell(4, 2).value = first.boxId;                       // B4 FBA编号
      ws.getCell(5, 2).value = first.refId;                       // B5 货物追踪编号
      ws.getCell(6, 2).value = opts.channelName;                  // B6 收货渠道
      ws.getCell(7, 2).value = countryCode(top.country);          // B7 目的地(用于 INDIRECT 选地址表)
      ws.getCell(9, 2).value = yxXinglianCustomsFor(top.customs); // B9 报关类型
      ws.getCell(10, 2).value = "否";                             // B10 是否投保
      ws.getCell(11, 2).value = "不投保";                         // B11 投保类型
      ws.getCell(12, 6).value = top.warehouseCode;                // F12 FBA仓库码
    },
    mapItem(it, top, ctx) {
      return [
        ctx.index + 1,            // A 箱序号
        it.boxId,                 // B FBA ID
        it.refId,                 // C POAnumber/货物追踪编号
        it.nameCh,                // D 中文名
        it.nameEn,                // E 英文名
        it.netWeight,             // F (单个产品)净重
        it.totalQty,              // G (单箱)数量
        it.unitPrice,             // H 申报单价
        it.currency || "USD",     // I 申报币别
        it.material,              // J 材质(英文)
        it.use,                   // K 用途(英文)
        hsCodeForLs(it.hsCode),   // L 海关编码
        it.grossWeight,           // M 单件实重
        it.lengthCm,              // N 长
        it.widthCm,               // O 宽
        it.heightCm,              // P 高
        urlOrZero(it.image),      // Q 产品图片
        top.hasBattery,           // R 是否带电
        top.hasMagnet,            // S 是否带磁
        urlOrZero(it.link),       // T 销售链接
      ];
    },
    channelSheet: { name: "US", col: 1, skipHeader: true },
  },

  // ── 皓鹏：B1 服务 / B2 地址库编码 / B10 国家 / B13 订单号 / B14 箱数 / F5 报关 / F6 交税 / F7 币种 / F13 PO ──
  皓鹏: {
    key: "皓鹏",
    templateFile: "皓鹏-模板.xlsx",
    sheetName: "单票下单-填写表",
    dataStartRow: 19,
    maxCol: 20,
    ready: true,
    fillHeader(ws, top, opts, items) {
      const first = items[0];
      ws.getCell(1, 2).value = opts.channelName;                    // B1 服务
      ws.getCell(2, 2).value = top.warehouseCode;                   // B2 地址库编码
      ws.getCell(10, 2).value = countryCode(top.country);           // B10 收件人国家代码
      ws.getCell(13, 2).value = opts.customerOrderNo;               // B13 客户订单号
      ws.getCell(14, 2).value = top.totalBoxes;                     // B14 箱数
      ws.getCell(5, 6).value = yxHaopengCustomsFor(top.customs);    // F5 报关方式
      ws.getCell(6, 6).value = yxTaxFor(top.clearance);             // F6 交税方式
      ws.getCell(7, 6).value = yxCurrencyFor(first.currency);       // F7 申报币种
      ws.getCell(13, 6).value = first.refId;                        // F13 PO Number
    },
    mapItem(it, _top, ctx) {
      return [
        `${ctx.boxPrefix}U${String(ctx.index + 1).padStart(6, "0")}`, // A 货箱编号
        it.nameEn,                // B 英文品名
        it.nameCh,                // C 中文品名
        it.totalQty,              // D 申报数量(单箱)
        "个",                     // E 申报单位
        it.unitPrice,             // F 申报单价
        it.grossWeight,           // G 货箱重量
        it.lengthCm,              // H 长
        it.widthCm,               // I 宽
        it.heightCm,              // J 高
        hsCodeForLs(it.hsCode),   // K 海关编码
        it.brand,                 // L 品牌
        it.material,              // M 材质
        it.model,                 // N 型号
        it.use,                   // O 用途
        urlOrZero(it.link),       // P 销售链接
        urlOrZero(it.image),      // Q 产品图片
        it.netWeight,             // R 产品重量(每pcs)
        it.sku,                   // S SKU
        it.refId,                 // T po Number
      ];
    },
    channelSheet: { name: "2026.8.27 服务渠道表", col: 3 },
  },

  // ── 美琦：与 LS 美线发票同构(18 列，A 列均为累计区间「起-止」) ──
  美琦: {
    key: "美琦",
    templateFile: "模板转换/美琦-美线模板.xlsx",
    sheetName: "发票",
    dataStartRow: 19,
    maxCol: 18,
    ready: true,
    fillHeader(ws, top, opts) {
      ws.getCell(1, 2).value = opts.customerOrderNo;       // B1 客户订单号
      ws.getCell(2, 2).value = opts.customerOrderNo;       // B2 客户参考号
      ws.getCell(3, 2).value = opts.channelName;           // B3 服务
      ws.getCell(4, 2).value = top.warehouseCode;          // B4 地址库编码(触发 VLOOKUP)
      ws.getCell(14, 2).value = top.hasBattery;            // B14 带电
      ws.getCell(15, 2).value = top.hasMagnet;             // B15 带磁
      ws.getCell(16, 2).value = lsCustomsFor(top.customs); // B16 报关方式
      ws.getCell(17, 2).value = top.totalBoxes;            // B17 箱数
      ws.getCell(1, 5).value = "清溪仓";                   // E1 预计交货站点(默认)
      writePickupDate(ws.getCell(2, 5));                    // E2 预计提货时间(今天)
      ws.getCell(3, 5).value = "上门提货";                 // E3 提货方式
    },
    mapItem(it, _top, ctx) {
      const perBox = it.boxCount ? r2(it.totalQty / it.boxCount) : 0;
      const start = ctx.cumulativeStart + 1;
      const end = ctx.cumulativeStart + it.boxCount;
      return [
        `${start}-${end}`,        // A 货箱编号(累计区间)
        it.nameEn,                // B 英文品名
        it.nameCh,                // C 中文品名
        perBox,                   // D 单箱数量
        it.unitPrice,             // E 单价
        it.grossWeight,           // F 货箱重量(毛重原值)
        r2(it.lengthCm - 1),      // G 长 -1
        r2(it.widthCm - 1),       // H 宽 -1
        r2(it.heightCm - 1),      // I 高 -1
        hsCodeForLs(it.hsCode),   // J 海关编码
        it.brand,                 // K 品牌
        it.material,              // L 材质
        it.model,                 // M 型号
        it.use,                   // N 用途
        0,                        // O 产品图片(固定 0)
        it.refId,                 // P PO Number
        it.boxCount,              // Q 物品箱号
        it.boxId,                 // R 物品 FBA ID
      ];
    },
    channelSheet: { name: "服务渠道", col: 2, skipHeader: true },
    // 同一个 FBA 货件的单号(货件追踪编码)箱号连续累加；换单号重新从 1 起
    groupKey: (it) => it.refId || it.boxId,
  },

  // ── 安邦：暂无模板，占位 ──
  安邦: {
    key: "安邦",
    templateFile: "",
    sheetName: "",
    dataStartRow: 1,
    maxCol: 1,
    ready: false,
    fillHeader() {},
    mapItem() {
      return [];
    },
  },
};

/** 供应商列表（供前端渲染下拉） */
export const YX_SUPPLIER_KEYS: YxSupplierKey[] = ["天图", "星链", "皓鹏", "美琦", "安邦"];

const yxChannelCache = new Map<YxSupplierKey, string[]>();

/** 从供应商模板的渠道 sheet 解析渠道名（供前端下拉），结果去重并缓存 */
export async function parseYxChannels(supplier: YxSupplierKey): Promise<string[]> {
  const spec = YX_SUPPLIERS[supplier];
  if (!spec || !spec.ready || !spec.channelSheet) return [];
  const cached = yxChannelCache.get(supplier);
  if (cached) return cached;

  const tplPath = path.join(process.cwd(), "public", "templates", "suppliers", spec.templateFile);
  if (!fs.existsSync(tplPath)) return [];
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(tplPath);
  const ws = wb.getWorksheet(spec.channelSheet.name);
  if (!ws) return [];

  const channels: string[] = [];
  const seen = new Set<string>();
  const startRow = spec.channelSheet.skipHeader ? 2 : 1;
  for (let r = startRow; r <= ws.rowCount; r++) {
    const v = cellText(ws.getCell(r, spec.channelSheet.col));
    if (v && !seen.has(v)) {
      seen.add(v);
      channels.push(v);
    }
  }
  yxChannelCache.set(supplier, channels);
  return channels;
}

/** 单张转换：易通标准模板 → 指定供应商发票 */
export async function convertYxToSupplier(
  filePath: string,
  sourceFileName: string,
  supplier: YxSupplierKey,
  opts: YxConvertOptions
): Promise<YxConvertResult> {
  const spec = YX_SUPPLIERS[supplier];
  if (!spec) throw new YxValidationError(`未知供应商：${supplier}`);
  if (!spec.ready) throw new YxValidationError(`供应商「${supplier}」暂未提供模板，无法转换`);

  // ── 1. 解析易通源 ──
  const { top, items } = await parseEttongSource(filePath);
  if (items.length === 0) throw new Error("易通标准模板中未找到有效货箱数据");

  // ── 2. 加载供应商模板 ──
  const tplPath = path.join(process.cwd(), "public", "templates", "suppliers", spec.templateFile);
  if (!fs.existsSync(tplPath)) throw new Error(`供应商模板不存在: ${tplPath}`);
  const outWb = new ExcelJS.Workbook();
  await outWb.xlsx.readFile(tplPath);
  const ws = outWb.getWorksheet(spec.sheetName);
  if (!ws) throw new Error(`模板中找不到「${spec.sheetName}」工作表`);

  // ── 3. 填充顶部字段 ──
  spec.fillHeader(ws, top, opts, items);

  // ── 4. 数据区（先清空模板示例数据行）──
  const dataStart = spec.dataStartRow;
  const rowsToDelete = ws.actualRowCount - dataStart + 1;
  if (rowsToDelete > 0) ws.spliceRows(dataStart, rowsToDelete);

  const boxPrefix = yxBoxPrefix(opts.customerOrderNo, items);
  let cumulative = 0;
  let currentGroup = ""; // 当前 FBA 分组键（美琦货箱编号累计区间按此重置）
  for (let i = 0; i < items.length; i++) {
    const item = items[i];
    if (spec.groupKey) {
      const gk = spec.groupKey(item);
      if (gk !== currentGroup) {
        currentGroup = gk;
        cumulative = 0; // 换 FBA 货件，箱号重新从 1 起
      }
    }
    const row = ws.getRow(dataStart + i);
    const values = spec.mapItem(item, top, {
      index: i,
      cumulativeStart: cumulative,
      boxPrefix,
    });
    values.forEach((v, ci) => {
      if (v !== null && v !== undefined && v !== "") row.getCell(ci + 1).value = v;
    });
    for (let c = 1; c <= spec.maxCol; c++) {
      row.getCell(c).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
    }
    cumulative += item.boxCount;
  }

  // ── 5. 输出 ──
  const buffer = Buffer.from(await outWb.xlsx.writeBuffer());
  const baseName =
    opts.customerOrderNo ||
    yxBoxPrefix("", items) ||
    path.basename(sourceFileName, path.extname(sourceFileName)) ||
    "未命名";
  const fileName = `${sanitizeFileName(baseName)}.xlsx`;

  const warnings: string[] = [];
  if (!top.warehouseCode) warnings.push("源模板缺少仓库代码");
  if (supplier === "皓鹏") {
    warnings.push("皓鹏收件地址字段为模板示例值，需人工核对或从地址库编码表填充");
  }

  return {
    sourceFile: sourceFileName,
    fileName,
    supplier,
    customerOrderNo: opts.customerOrderNo,
    warehouseCode: top.warehouseCode,
    totalBoxes: top.totalBoxes,
    dataRows: items.length,
    warnings,
    buffer,
  };
}

/** 批量转换：同国家校验通过后逐张转换并打包 ZIP */
export async function convertYxBatch(
  files: { path: string; name: string }[],
  supplier: YxSupplierKey,
  opts: YxConvertOptions
): Promise<YxBatchResult> {
  const spec = YX_SUPPLIERS[supplier];
  if (!spec || !spec.ready) throw new YxValidationError(`供应商「${supplier}」暂未提供模板，无法转换`);
  if (files.length === 0) throw new YxValidationError("未提供待转换文件");

  // ── 校验：所有源文件国家一致 ──
  const countries = new Set<string>();
  for (const f of files) {
    const { top, items } = await parseEttongSource(f.path);
    if (items.length === 0) throw new YxValidationError(`「${f.name}」未找到有效货箱数据`);
    countries.add(top.country || "未识别");
  }
  if (countries.size > 1) {
    throw new YxValidationError(
      `批量转换仅支持同一国家的发票，检测到 ${[...countries].join("、")} 等多个国家，请分开上传`
    );
  }
  const country = [...countries][0] || "";

  // ── 逐张转换 + ZIP 打包 ──
  const zip = new JSZip();
  const outFiles: YxBatchFile[] = [];
  const warnings: string[] = [];
  const usedNames = new Map<string, number>();

  for (const f of files) {
    const result = await convertYxToSupplier(f.path, f.name, supplier, opts);

    let zipName = result.fileName;
    const dup = usedNames.get(result.fileName) || 0;
    if (dup > 0) {
      const ext = path.extname(zipName);
      const stem = zipName.slice(0, zipName.length - ext.length);
      zipName = `${stem}(${dup})${ext}`;
    }
    usedNames.set(result.fileName, (usedNames.get(result.fileName) || 0) + 1);

    zip.file(zipName, result.buffer);
    outFiles.push({
      sourceFile: result.sourceFile,
      fileName: result.fileName,
      country,
      warehouseCode: result.warehouseCode,
      totalBoxes: result.totalBoxes,
      dataRows: result.dataRows,
    });
    for (const w of result.warnings) warnings.push(`[${f.name}] ${w}`);
  }

  zip.file(
    "转换结果汇总.json",
    JSON.stringify(
      { supplier, channelName: opts.channelName, country, files: outFiles, warnings },
      null,
      2
    )
  );

  const buffer = Buffer.from(await zip.generateAsync({ type: "nodebuffer" }));
  const base = opts.customerOrderNo || sanitizeFileName(country) || "延讯批量";
  return {
    sourceFiles: files.map((f) => f.name),
    zipFileName: `${base}-${supplier}批量转换.zip`,
    supplier,
    country,
    files: outFiles,
    warnings,
    buffer,
  };
}

/** 兼容旧调用：天图单张转换 */
export async function convertYxToTiantu(
  filePath: string,
  sourceFileName: string,
  customerOrderNo: string,
  channelName: string
): Promise<YxConvertResult> {
  return convertYxToSupplier(filePath, sourceFileName, "天图", { customerOrderNo, channelName });
}
