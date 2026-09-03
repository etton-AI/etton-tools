/**
 * 延讯下单优化 — 将延讯下单发票(发货单)转换为易通下单模版
 *
 * 数据流：
 *   延讯下单发票(发货单 sheet)
 *     ├─ 顶部信息: 运输方式 / 正式报关 / 带电 / 目的地 / 渠道 / FBA号 / ReferenceID / 总箱数
 *     └─ 货箱清单: 箱号 / 品名 / 英文 / 材质 / 用途 / 发货数量 / 申报货值 / 海关编码 / 毛重 / 长宽高 / 币种
 *       ↓
 *   易通下单模版(public/templates/易通下单模版.xlsx)
 *     ├─ 顶部字段: 业务类型 / 报关方式 / 带电 / 仓点类型 / 收件人国家 / 仓库代码 / 备注 / 总箱数
 *     └─ 数据区(R25 起, 24 列): Shipment ID → 申报数量
 *
 * 关键规则：
 *   1. 报关方式映射: 公司自报→普通报关(买单), 永德吉报关→报关退税
 *   2. 带电: 不带电则不填
 *   3. 混箱: 同一箱号多行时，仅首行保留毛重/长宽高，其余行清空
 *   4. 输出文件名 = 导入文件名(去扩展名) + _ETTON
 */

import ExcelJS from "exceljs";
import JSZip from "jszip";
import path from "path";
import fs from "fs";

// ============================================================
// Types
// ============================================================

/** 海外仓收件人信息(易通「私人地址/海外仓」区块) */
export interface OverseasAddress {
  name: string;      // 收件人姓名
  company: string;   // 收件人公司
  address: string;   // 收件人地址
  phone: string;     // 收件人联系方式
  city: string;      // 收件人城市
  state: string;     // 收件人省份/州
  zip: string;       // 收件人邮编
}

/** 延讯发票顶部信息 */
export interface YanxunTopInfo {
  transportMode: string;      // 运输方式 → 易通「业务类型」
  customsRaw: string;         // 正式报关原文 → 映射后为「报关方式」
  customsMapped: string;      // 映射后的报关方式
  hasBattery: string;         // 带电 (是/否)
  country: string;            // 目的地国家 → 「收件人国家」
  warehouseCode: string;      // 仓点 → 「仓库代码」(FBA 场景) / 海外仓地址文本(海外仓场景)
  warehouseType: string;      // 场景判断 → FBA / 海外仓(据「FBA号/海外仓」字段值是否含「海外仓」)
  channel: string;            // 渠道 → 「备注」
  fbaId: string;              // FBA号/海外仓 原始值 → FBA 场景文件名
  transferOrderNo: string;    // 调拨单号 → 海外仓场景文件名
  overseas: OverseasAddress;  // 海外仓收件人信息(仅海外仓场景)
  referenceId: string;        // ReferenceID → 数据区「货件追踪编码」
  totalBoxes: number;         // 总箱数
}

/** 一行货箱数据(延讯箱单解析结果) */
export interface YanxunDataRow {
  boxId: string;              // 箱号 → Shipment ID
  nameEn: string;             // 英文 → Name(En)
  nameCh: string;             // 品名 → Name(Ch)
  material: string;           // 材质 → Material
  use: string;                // 用途 → Use
  quantity: number;           // 发货数量 → Quantity
  unitPrice: number;          // 申报货值 → Unit Price
  totalPrice: number;         // 数量 × 单价 → Total Price
  currency: string;           // 币种 → currency
  hsCode: string;             // 海关编码 → HS Code
  grossWeight: number;        // 毛重 → 净重 & 毛重
  volumeCbm: number;          // 体积CBM → 预计总体积
  lengthCm: number;           // 长 → 长
  widthCm: number;            // 宽 → 宽
  heightCm: number;           // 高 → 高
  link: string;               // 链接 → 链接
}

/** 顶层转换结果(存入 session，不含 buffer 返回给客户端) */
export interface YanxunConvertResult {
  sourceFile: string;
  fileName: string;           // FBA号.xlsx
  fbaId: string;
  totalBoxes: number;
  dataRows: number;
  mixedBoxGroups: number;     // 混箱组数
  topInfo: YanxunTopInfo;
  warnings: string[];         // 需人工处理的必填项提示
  buffer: Buffer;
}

/** 延讯箱单列映射 */
interface YanxunColumnMap {
  boxId: number;
  qty: number;
  lengthCm: number;
  widthCm: number;
  heightCm: number;
  grossWeight: number;
  volumeCbm: number;
  nameCh: number;
  nameEn: number;
  hsCode: number;
  unitPrice: number;
  material: number;
  use: number;
  currency?: number;  // 可选：海外仓备货单无「币种」列，按目的地国家默认(美国=USD)
  link: number;
}

// ============================================================
// 常量
// ============================================================

/** 延讯箱单表头关键词(trim 后 startsWith 匹配) */
const YANXUN_HEADER_PATTERNS: Record<keyof YanxunColumnMap, string[]> = {
  boxId:        ["箱号"],
  qty:          ["发货数量"],
  lengthCm:     ["长（", "长("],
  widthCm:      ["宽"],
  heightCm:     ["高"],
  grossWeight:  ["毛重"],
  volumeCbm:    ["体积"],
  nameCh:       ["品名"],
  nameEn:       ["英文"],
  hsCode:       ["海关编码"],
  unitPrice:    ["申报货值"],
  material:     ["材质"],
  use:          ["用途"],
  currency:     ["币种"],
  link:         ["链接"],
};

/** 报关方式映射: 延讯「正式报关」 → 易通「报关方式」 */
const CUSTOMS_MAP: Record<string, string> = {
  "公司自报": "普通报关",
  "永德吉报关": "报关退税",
  "永德吉": "报关退税",
  "否": "普通报关",  // 海外仓场景「正式报关=否」→ 普通报关(买单)
};

// ============================================================
// Helpers
// ============================================================

/** 安全读取单元格文本(处理 null / string / number / richText / 公式对象) */
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
    if (o.result !== undefined && o.result !== null) {
      // 公式结果：字符串/数字/布尔直接用；错误对象(如 {error:"#N/A"}) 视为空
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

/** 安全读取数值(处理文本/公式) */
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

/** 清洗文件名非法字符 */
function sanitizeFileName(name: string): string {
  const cleaned = name.replace(/[\\/:*?"<>|]/g, "_").trim();
  return cleaned || "未命名";
}

// ============================================================
// 延讯发票解析
// ============================================================

/** 在指定行范围内查找标签单元格(contains 匹配，返回第一个) */
function findLabel(
  ws: ExcelJS.Worksheet,
  keyword: string,
  maxRow: number,
  maxCol = 20,
): { row: number; col: number } | null {
  for (let r = 1; r <= maxRow; r++) {
    for (let c = 1; c <= maxCol; c++) {
      const t = cellText(ws.getCell(r, c));
      if (t.includes(keyword)) return { row: r, col: c };
    }
  }
  return null;
}

/** 查找所有匹配标签(按行列顺序) */
function findAllLabels(
  ws: ExcelJS.Worksheet,
  keyword: string,
  maxRow: number,
  maxCol = 20,
): { row: number; col: number }[] {
  const result: { row: number; col: number }[] = [];
  for (let r = 1; r <= maxRow; r++) {
    for (let c = 1; c <= maxCol; c++) {
      const t = cellText(ws.getCell(r, c));
      if (t.includes(keyword)) result.push({ row: r, col: c });
    }
  }
  return result;
}

/** 取标签下方第一个非空值 */
function valueBelow(ws: ExcelJS.Worksheet, row: number, col: number): string {
  for (let r = row + 1; r <= row + 3; r++) {
    const t = cellText(ws.getCell(r, col));
    if (t) return t;
  }
  return "";
}

/** 取标签右侧第一个非空值 */
function valueRight(ws: ExcelJS.Worksheet, row: number, col: number, maxScan = 8): string {
  for (let c = col + 1; c <= col + maxScan; c++) {
    const t = cellText(ws.getCell(row, c));
    if (t) return t;
  }
  return "";
}

/** 空海外仓收件人信息 */
function emptyOverseas(): OverseasAddress {
  return { name: "", company: "", address: "", phone: "", city: "", state: "", zip: "" };
}

/**
 * 解析海外仓地址文本 → 收件人信息(姓名/公司/地址/电话 + 城市/州/邮编)。
 *
 * 延讯海外仓发票的「目的地」字段填的是多行地址文本，形如：
 *   美西2号仓新仓：
 *   收件公司/收件人：HUA GAN-HYH-7283102
 *   派送地址：1449 W Industrial Park St Covina CA 91722
 *   电话：4695563956
 * 城市/州/邮编按美式地址尾部「City STATE ZIP」解析(州为 2 位大写字母，邮编为 5 位数字)。
 */
function parseOverseasAddress(text: string): OverseasAddress {
  const result = emptyOverseas();

  // 收件公司/收件人 → 姓名 & 公司(标准答案两者同值)
  const contact =
    text.match(/收件公司\/收件人[：:]\s*([^\n\r]+)/) ||
    text.match(/收件人[：:]\s*([^\n\r]+)/) ||
    text.match(/联系人[：:]\s*([^\n\r]+)/);
  if (contact) {
    result.name = contact[1].trim();
    result.company = contact[1].trim();
  }

  // 派送地址 / 地址
  const addr =
    text.match(/派送地址[：:]\s*([^\n\r]+)/) ||
    text.match(/地址[：:]\s*([^\n\r]+)/);
  if (addr) result.address = addr[1].trim();

  // 电话 / 联系方式
  const phone =
    text.match(/电话[：:]\s*([^\n\r]+)/) ||
    text.match(/联系方式[：:]\s*([^\n\r]+)/);
  if (phone) result.phone = phone[1].trim();

  // 从地址尾部解析 城市/州/邮编(美式：… 城市 州(2大写字母) 邮编(5数字))
  if (result.address) {
    const m = result.address.match(/(.+?)\s+([A-Z]{2})\s+(\d{5}(?:-\d{4})?)\s*$/);
    if (m) {
      result.state = m[2];
      result.zip = m[3];
      const cityWords = m[1].split(/\s+/);
      result.city = cityWords[cityWords.length - 1] || "";
    }
  }

  return result;
}

/** 定位延讯箱单表头行(同时含「箱号」和「发货数量」) */
function findHeaderRow(ws: ExcelJS.Worksheet): number {
  for (let r = 1; r <= Math.min(ws.rowCount, 60); r++) {
    const row = ws.getRow(r);
    let hasBox = false;
    let hasQty = false;
    row.eachCell({ includeEmpty: false }, (cell) => {
      const t = cellText(cell);
      if (t.startsWith("箱号")) hasBox = true;
      if (t.startsWith("发货数量")) hasQty = true;
    });
    if (hasBox && hasQty) return r;
  }
  throw new Error("找不到货箱清单表头（需同时包含「箱号」和「发货数量」列）。请确认上传的是延讯下单发票。");
}

/** 根据表头行构建列映射 */
function buildYanxunColumnMap(headerRow: ExcelJS.Row): YanxunColumnMap {
  const map: Partial<YanxunColumnMap> = {};

  headerRow.eachCell({ includeEmpty: false }, (cell, cn) => {
    const t = cellText(cell);
    for (const [key, patterns] of Object.entries(YANXUN_HEADER_PATTERNS)) {
      if (map[key as keyof YanxunColumnMap] !== undefined) continue;
      if (patterns.some((p) => t.startsWith(p))) {
        map[key as keyof YanxunColumnMap] = cn;
        break;
      }
    }
  });

  const required: (keyof YanxunColumnMap)[] = [
    "boxId", "qty", "lengthCm", "widthCm", "heightCm", "grossWeight",
    "nameCh", "nameEn", "hsCode", "unitPrice", "material", "use",
  ];
  const missing = required.filter((k) => map[k] === undefined);
  if (missing.length > 0) {
    throw new Error(`无法在箱单表头中找到以下列: ${missing.join(", ")}。请确认上传的是延讯下单发票。`);
  }

  return map as YanxunColumnMap;
}

/** 解析延讯发货单顶部信息 */
function parseTopInfo(ws: ExcelJS.Worksheet, headerRow: number): YanxunTopInfo {
  const maxRow = headerRow - 1; // 顶部区域在箱单表头之前

  // 运输方式(标签在上，值在下)
  const transportLabel = findLabel(ws, "运输方式", maxRow);
  const transportMode = transportLabel ? valueBelow(ws, transportLabel.row, transportLabel.col) : "";

  // 正式报关
  const customsLabel = findLabel(ws, "正式报关", maxRow);
  const customsRaw = customsLabel ? valueBelow(ws, customsLabel.row, customsLabel.col) : "";
  const customsMapped = CUSTOMS_MAP[customsRaw] ?? customsRaw ?? "";

  // 带电
  const batteryLabel = findLabel(ws, "带电", maxRow);
  const batteryRaw = batteryLabel ? valueBelow(ws, batteryLabel.row, batteryLabel.col) : "";
  const hasBattery = batteryRaw === "是" || batteryRaw === "带电" ? "是" : "";

  // 目的地(第一个=国家, 第二个=仓点/海外仓地址)
  const destLabels = findAllLabels(ws, "目的地", maxRow);
  const country = destLabels.length > 0 ? valueBelow(ws, destLabels[0].row, destLabels[0].col) : "";
  const warehouseCode = destLabels.length > 1 ? valueBelow(ws, destLabels[1].row, destLabels[1].col) : "";

  // FBA号/海外仓(标签在左，值在右；只取右侧紧邻一格，避免扫到「调拨单号」标签)
  const fbaLabel = findLabel(ws, "FBA号", maxRow);
  const fbaId = fbaLabel ? valueRight(ws, fbaLabel.row, fbaLabel.col, 1) : "";

  // 调拨单号(海外仓场景用于命名)
  const transferLabel = findLabel(ws, "调拨单号", maxRow);
  const transferOrderNo = transferLabel ? valueRight(ws, transferLabel.row, transferLabel.col, 1) : "";

  // 场景判断：据「FBA号/海外仓」字段值是否含「海外仓」字眼(目的地字段只是地址，不可用于判断)
  const warehouseType = /海外仓/.test(fbaId) ? "海外仓" : "FBA";

  // 渠道(「物流商/渠道」标签右侧多格；跳过地址类文本，取第一个渠道名)
  const channelLabel = findLabel(ws, "物流商/渠道", maxRow);
  let channel = "";
  if (channelLabel) {
    for (let c = channelLabel.col + 1; c <= channelLabel.col + 6; c++) {
      const t = cellText(ws.getCell(channelLabel.row, c));
      if (!t) continue;
      // 海外仓场景渠道右侧先是被目的地地址占据，需跳过地址类文本
      if (/发件人|地址|邮编|收件|电话|联系人/i.test(t)) continue;
      channel = t;
      break;
    }
  }

  // 海外仓收件人信息(从「目的地」第二格地址文本解析)；FBA 场景为空
  const overseas = warehouseType === "海外仓" ? parseOverseasAddress(warehouseCode) : emptyOverseas();

  // ReferenceID
  const refLabel = findLabel(ws, "ReferenceID", maxRow);
  const referenceId = refLabel ? valueRight(ws, refLabel.row, refLabel.col) : "";

  // 总箱数
  const boxLabel = findLabel(ws, "总箱数", maxRow);
  const totalBoxesRaw = boxLabel ? valueRight(ws, boxLabel.row, boxLabel.col) : "";
  const totalBoxes = parseInt(totalBoxesRaw, 10) || 0;

  return {
    transportMode,
    customsRaw,
    customsMapped,
    hasBattery,
    country,
    warehouseCode,
    warehouseType,
    channel,
    fbaId,
    transferOrderNo,
    overseas,
    referenceId,
    totalBoxes,
  };
}

/**
 * 校验必填项(缺失返回错误信息列表)。
 *
 * 易通模板必填项与延讯来源的对应：
 *   - 发货公司* / 是否合并报关* → 模板自带默认值，无需校验
 *   - 业务类型*(运输方式) / 报关方式*(正式报关) → 延讯顶部，缺失报错
 *   - 总箱数、预计总重量、预计总体积（均带星号必填）→ 由数据自动计算，不会缺失
 *   - 服务渠道* → 不做自动校正(延讯渠道名与易通服务渠道名非一一对应，由人工选定)
 *   - 二选一必填组「仓点类型+收件人国家+仓库代码」(FBA 地址库)与「私人地址/海外仓」二选一：
 *     FBA 场景走第一组(仓点类型=FBA，校验 FBA号/仓库代码)；
 *     海外仓场景走第二组(校验收件人姓名/地址 + 调拨单号，仓库代码不适用)
 */
function validateTopInfo(top: YanxunTopInfo): string[] {
  const errors: string[] = [];
  if (!top.transportMode) errors.push("无运输方式");
  if (!top.customsRaw) errors.push("无报关方式");
  if (!top.country) errors.push("无目的地国家");
  if (!top.channel) errors.push("无渠道名");

  if (top.warehouseType === "FBA") {
    // FBA 地址库组
    if (!top.warehouseCode) errors.push("无仓库代码");
    if (!top.fbaId) errors.push("无FBA号");
  } else {
    // 海外仓 → 「私人地址/海外仓」组(自动从地址文本解析)
    if (!top.overseas.name) errors.push("无收件人姓名");
    if (!top.overseas.address) errors.push("无收件人地址");
    if (!top.transferOrderNo) errors.push("无调拨单号");
  }

  return errors;
}

/** 解析延讯货箱清单数据行 */
function parseDataRows(
  ws: ExcelJS.Worksheet,
  headerRow: number,
  colMap: YanxunColumnMap,
): YanxunDataRow[] {
  const rows: YanxunDataRow[] = [];

  for (let r = headerRow + 1; r <= ws.rowCount; r++) {
    const boxId = cellText(ws.getCell(r, colMap.boxId));
    if (!boxId) break; // 箱号为空 = 结束(合计行)

    const quantity = cellNum(ws.getCell(r, colMap.qty));
    const unitPrice = cellNum(ws.getCell(r, colMap.unitPrice));

    rows.push({
      boxId,
      nameEn: cellText(ws.getCell(r, colMap.nameEn)),
      nameCh: cellText(ws.getCell(r, colMap.nameCh)),
      material: cellText(ws.getCell(r, colMap.material)),
      use: cellText(ws.getCell(r, colMap.use)),
      quantity,
      unitPrice,
      totalPrice: quantity * unitPrice,
      currency: colMap.currency
        ? cellText(ws.getCell(r, colMap.currency)) || "USD"
        : "USD", // 海外仓备货单无币种列，美国渠道默认美元
      hsCode: cellText(ws.getCell(r, colMap.hsCode)),
      grossWeight: cellNum(ws.getCell(r, colMap.grossWeight)),
      volumeCbm: cellNum(ws.getCell(r, colMap.volumeCbm)),
      lengthCm: cellNum(ws.getCell(r, colMap.lengthCm)),
      widthCm: cellNum(ws.getCell(r, colMap.widthCm)),
      heightCm: cellNum(ws.getCell(r, colMap.heightCm)),
      link: cellText(ws.getCell(r, colMap.link)),
    });
  }

  return rows;
}

/** 统计混箱组数(同一箱号出现多行 = 混箱) */
function countMixedBoxGroups(rows: YanxunDataRow[]): number {
  const count = new Map<string, number>();
  for (const r of rows) count.set(r.boxId, (count.get(r.boxId) || 0) + 1);
  return Array.from(count.values()).filter((c) => c > 1).length;
}

// ============================================================
// 易通模版填充
// ============================================================

/** 在易通模版中按标签定位并填入值 */
function setTemplateField(
  ws: ExcelJS.Worksheet,
  keyword: string,
  side: "left" | "right",
  value: string,
): boolean {
  const cols = side === "left" ? [1] : [5];
  for (let r = 1; r <= 22; r++) {
    for (const c of cols) {
      const t = cellText(ws.getCell(r, c));
      if (t.includes(keyword)) {
        ws.getCell(r, c + 1).value = value;
        return true;
      }
    }
  }
  return false;
}

// ============================================================
// 主入口
// ============================================================

export async function convertYanxunToEtton(
  filePath: string,
  sourceFileName: string,
): Promise<YanxunConvertResult> {
  // ── 1. 读取延讯发票 ──
  const srcWb = new ExcelJS.Workbook();
  await srcWb.xlsx.readFile(filePath);

  const srcSheet = srcWb.worksheets[0]; // 发货单(第一个 sheet)
  if (!srcSheet) throw new Error("延讯发票中没有找到工作表");

  // ── 2. 定位箱单表头并解析 ──
  const headerRow = findHeaderRow(srcSheet);
  const colMap = buildYanxunColumnMap(srcSheet.getRow(headerRow));
  const topInfo = parseTopInfo(srcSheet, headerRow);

  // 必填项校验(批量时失败项会汇总到 failed[]，提示客户及时修正)
  const topErrors = validateTopInfo(topInfo);
  if (topErrors.length > 0) {
    throw new Error(`必填项缺失: ${topErrors.join("、")}`);
  }

  const dataRows = parseDataRows(srcSheet, headerRow, colMap);

  if (dataRows.length === 0) {
    throw new Error("未找到有效货箱数据。请确认上传的是延讯下单发票。");
  }

  const mixedBoxGroups = countMixedBoxGroups(dataRows);

  // 总箱数: 优先顶部值，fallback 到去重箱号数
  const distinctBoxCount = new Set(dataRows.map((r) => r.boxId)).size;
  const totalBoxes = topInfo.totalBoxes || distinctBoxCount;

  console.log(
    `📦 延讯转换: 顶部FBA号=${topInfo.fbaId}, 数据${dataRows.length}行, 混箱${mixedBoxGroups}组, 总箱数${totalBoxes}`
  );

  // ── 3. 加载易通模版 ──
  const tplPath = path.join(process.cwd(), "public", "templates", "易通下单模版.xlsx");
  if (!fs.existsSync(tplPath)) {
    throw new Error(`易通下单模版不存在: ${tplPath}`);
  }
  const outWb = new ExcelJS.Workbook();
  await outWb.xlsx.readFile(tplPath);
  const outSheet = outWb.getWorksheet("ETTON电商物流 下单模板");
  if (!outSheet) throw new Error("易通模版中找不到「ETTON电商物流 下单模板」工作表");

  // ── 4. 填充顶部字段 ──
  setTemplateField(outSheet, "业务类型", "left", topInfo.transportMode || "海运");
  setTemplateField(outSheet, "报关方式", "left", topInfo.customsMapped || "普通报关");
  setTemplateField(outSheet, "总箱数", "left", String(totalBoxes));
  setTemplateField(outSheet, "备注", "left", topInfo.channel);

  // 带电: 不带电则不填(模板默认为空)
  if (topInfo.hasBattery === "是") {
    setTemplateField(outSheet, "带电", "right", "是");
  }

  // 二选一必填组：
  //   FBA 场景 → 第一组「仓点类型/收件人国家/仓库代码」
  //   海外仓场景 → 第二组「私人地址/海外仓」(R14-R22，值在 F 列=第 6 列，固定位置)
  if (topInfo.warehouseType === "FBA") {
    setTemplateField(outSheet, "仓点类型", "right", "FBA");
    setTemplateField(outSheet, "收件人国家", "right", topInfo.country);
    setTemplateField(outSheet, "仓库代码", "right", topInfo.warehouseCode);
  } else {
    // 海外仓场景：先清空 FBA 地址库组(仓点类型/收件人国家/仓库代码/仓库地址/邮编，模板自带默认值)
    for (let r = 7; r <= 12; r++) {
      outSheet.getCell(r, 6).value = null;
    }
    const ov = topInfo.overseas;
    outSheet.getCell(14, 6).value = ov.name;       // 收件人姓名
    outSheet.getCell(15, 6).value = ov.company;    // 收件人公司
    outSheet.getCell(16, 6).value = topInfo.country; // 收件人国家
    outSheet.getCell(17, 6).value = ov.city;       // 收件人城市
    outSheet.getCell(18, 6).value = ov.state;      // 收件人省份/州
    outSheet.getCell(19, 6).value = ov.zip;        // 收件人邮编
    outSheet.getCell(20, 6).value = ov.phone;      // 收件人联系方式
    // 收件人邮箱(R21)：延讯无此数据，保留模板原样(标准答案同样未动)
    outSheet.getCell(22, 6).value = ov.address;    // 收件人地址
  }

  // 预计总重量/总体积(从数据合计，混箱仅首行计入)
  let totalWeight = 0;
  let totalVolume = 0;
  const seenBox = new Set<string>();
  for (const d of dataRows) {
    if (seenBox.has(d.boxId)) continue; // 混箱后续行不计入重量/体积
    seenBox.add(d.boxId);
    totalWeight += d.grossWeight;
    totalVolume += d.volumeCbm; // 用延讯「体积CBM」列(每箱已 ROUND 到 2 位)
  }
  setTemplateField(outSheet, "预计总重量", "left", String(Math.round(totalWeight * 100) / 100));
  setTemplateField(outSheet, "预计总体积", "left", String(Math.round(totalVolume * 100) / 100));

  // ── 5. 填充数据区 ──
  const dataStartRow = 25;
  // 必须用 actualRowCount（最后一个有数据的行号）而非 rowCount。
  // rowCount 会把模板里带格式的空行也算进去（如本例 =53），据此算出 count=29 时
  // spliceRows 内部 nKeep = start+count = 54 > 实际行数，删除循环一次都不执行，
  // 导致模板示例数据行（如 FBA19MXX2JCWU000004/005）残留到输出文件底部。
  const rowsToDelete = outSheet.actualRowCount - dataStartRow + 1;
  if (rowsToDelete > 0) {
    outSheet.spliceRows(dataStartRow, rowsToDelete); // 删除模板示例数据行
  }

  const writtenBoxes = new Set<string>();
  for (let i = 0; i < dataRows.length; i++) {
    const d = dataRows[i];
    const row = outSheet.getRow(dataStartRow + i);
    const isFirstOfBox = !writtenBoxes.has(d.boxId);
    writtenBoxes.add(d.boxId);

    row.getCell(1).value = d.boxId;                       // Shipment ID
    row.getCell(2).value =
      topInfo.warehouseType === "海外仓" ? "/" : topInfo.referenceId; // Reference ID(海外仓无追踪编码，固定"/")
    row.getCell(3).value = d.nameEn;                      // Name(En)
    row.getCell(4).value = d.nameCh;                      // Name(Ch)
    row.getCell(5).value = d.material;                    // Material
    row.getCell(6).value = d.use;                         // Use
    row.getCell(7).value = isFirstOfBox ? 1 : 0;           // Number(箱数)，混箱后续行补 0
    row.getCell(8).value = d.quantity;                    // Quantity
    row.getCell(9).value = d.unitPrice;                   // Unit Price
    row.getCell(10).value = d.totalPrice;                 // Total Price
    row.getCell(11).value = d.currency;                   // currency
    row.getCell(12).value = d.hsCode;                     // HS Code
    row.getCell(13).value = "无";                          // brand
    row.getCell(14).value = "无";                          // Model
    row.getCell(15).value = "无";                          // Brand Type
    row.getCell(16).value = isFirstOfBox ? d.grossWeight : 0; // 净重，混箱后续行补 0
    row.getCell(17).value = isFirstOfBox ? d.grossWeight : 0; // 毛重，混箱后续行补 0
    row.getCell(18).value = isFirstOfBox ? d.lengthCm : 0;    // 长，混箱后续行补 0
    row.getCell(19).value = isFirstOfBox ? d.widthCm : 0;     // 宽，混箱后续行补 0
    row.getCell(20).value = isFirstOfBox ? d.heightCm : 0;    // 高，混箱后续行补 0
    row.getCell(21).value = d.link || 0;                  // 链接（销售链接，无数据填 0）
    row.getCell(22).value = 0;                            // 图片（产品图片，无数据填 0）
    // 是否申报、申报数量：延讯无对应数据，留空（不写值，与标准答案一致）
  }

  // ── 6. 数据区边框 ──
  for (let r = dataStartRow; r < dataStartRow + dataRows.length; r++) {
    for (let c = 1; c <= 24; c++) {
      outSheet.getCell(r, c).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
    }
  }

  // ── 7. 输出 ──
  const buffer = Buffer.from(await outWb.xlsx.writeBuffer());
  // 文件名：ETTON_FBA号 / ETTON_调拨单号
  const nameKey = topInfo.warehouseType === "FBA" ? topInfo.fbaId : topInfo.transferOrderNo;
  const fileName = `ETTON_${sanitizeFileName(nameKey || "未命名")}.xlsx`;
  const fbaId = topInfo.fbaId; // 仅用于展示

  // 需人工处理的提示（延讯无法自动映射到易通的字段）
  const warnings: string[] = [];

  return {
    sourceFile: sourceFileName,
    fileName,
    fbaId,
    totalBoxes,
    dataRows: dataRows.length,
    mixedBoxGroups,
    topInfo,
    warnings,
    buffer,
  };
}

// ============================================================
// 批量打包
// ============================================================

/** 将多个转换结果打包为 ZIP（重名 FBA号 自动追加序号，附带汇总 json） */
export async function generateYanxunZip(results: YanxunConvertResult[]): Promise<Buffer> {
  const zip = new JSZip();
  const nameCount = new Map<string, number>();

  for (const r of results) {
    let name = r.fileName;
    const seen = nameCount.get(r.fileName) || 0;
    if (seen > 0) {
      const base = r.fileName.replace(/\.xlsx$/i, "");
      name = `${base}_${seen + 1}.xlsx`;
    }
    nameCount.set(r.fileName, seen + 1);
    zip.file(name, r.buffer);
  }

  const summary = results.map((r) => ({
    sourceFile: r.sourceFile,
    fileName: r.fileName,
    fbaId: r.fbaId,
    totalBoxes: r.totalBoxes,
    dataRows: r.dataRows,
    mixedBoxGroups: r.mixedBoxGroups,
    warnings: r.warnings,
  }));
  zip.file("转换结果汇总.json", JSON.stringify(summary, null, 2));

  return Buffer.from(await zip.generateAsync({ type: "nodebuffer" }));
}
