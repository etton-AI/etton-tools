/**
 * 思锐客户下单发票 → 供应商下单模板 转换
 *
 * 数据流：
 *   思锐下单发票(sheet「ETTON电商物流 下单模板」)
 *     ├─ 订单头: F2带电/F3带磁/B9报关方式/B10保险/B11清关方式/B12-17 VAT-EORI/B18总箱数
 *     │         F14-F22 收件人(姓名/公司/国家/城市/省份州/邮编/联系方式/邮箱/地址)
 *     └─ 箱单明细(R25 起, A-V): 货箱编号/追踪编码/品名/材质/用途/箱数/总数量/单价/总价/币种/海关编码/品牌/型号/品牌类型/净重/毛重/长宽高/链接/图片
 *       ↓
 *   供应商模板(public/templates/sirui-suppliers/): 华威尔 / 美琦 / 英美 / 航乐
 *
 * 关键规则(以答案文件反推为准)：
 *   1. 混箱语义：G>=1 开 G 个新箱(单箱数量=H/G)；G=0/空=混箱(共享当前箱, 单箱数量=H)
 *   2. 华威尔(一箱一行 explode)：箱唛号 = boxId + "U" + 两位递增序号(跨全部箱累计, 混箱共享当前序号)
 *   3. 航乐(非FBA 一箱一行 explode)：Box No 顺序 1..N；混箱尺寸继承当前箱
 *   4. 美琦/英美(排序号)：累计区间 "起-止"(G=1 单个数字)，混箱共享当前累计箱号
 *   5. 单个产品净重 = 单箱净重 / 单箱数量
 *   6. 图片：华威尔/美琦源用 WPS DISPIMG(cellimages.xml)；航乐源用浮动图片(drawing1.xml)；英美无图片列
 *   7. 美琦 长/宽/高 = 源值 - 1(美琦特有)；英美地址字段为模板 VLOOKUP(按地址库编码带出)，
 *      商业/私人地址需清空地址库编码并手动填入收件信息
 *   8. 浮点计算统一 clean()(toPrecision(15)) 去噪, 对齐 Excel/WPS 精度(14.7/12 → 1.225)
 */

import ExcelJS from "exceljs";
import JSZip from "jszip";
import path from "path";
import fs from "fs";
import crypto from "crypto";

// ============================================================
// Types
// ============================================================

export type SiruiSupplierKey = "华威尔" | "美琦" | "英美" | "航乐";

/** 一行思锐箱单明细 */
export interface SiruiItem {
  boxId: string;         // A 货箱编号(非FBA 时为箱号区间如 "1-2"; FBA 时为 FBA 子单号)
  refId: string;         // B 货件追踪编码(Reference ID)
  nameEn: string;        // C 英文品名
  nameCh: string;        // D 中文品名
  material: string;      // E 材质
  use: string;           // F 用途
  boxCount: number;      // G 箱数/件数(>=1 开新箱; 0/空=混箱)
  totalQty: number;      // H 总产品数量
  unitPrice: number;     // I 申报单价
  totalPrice: number;    // J 申报总价
  currency: string;      // K 申报币种
  hsCode: string;        // L 海关编码
  brand: string;         // M 品牌
  model: string;         // N 型号
  brandType: string;     // O 品牌类型
  netWeight: number;     // P 单箱货物净重
  grossWeight: number;   // Q 单箱货物毛重
  lengthCm: number;      // R 长
  widthCm: number;       // S 宽
  heightCm: number;      // T 高
  link: string;          // U 销售链接
  dispimgId: string;     // V 列 DISPIMG 图片 ID(源为 WPS 嵌入图时)
  image: Buffer | null;  // 提取出的源产品图(用于 DISPIMG 注入)
}

/** 思锐发票订单头 */
export interface SiruiTopInfo {
  hasBattery: string;    // F2 带电
  hasMagnet: string;     // F3 带磁
  customs: string;       // B9 报关方式
  insurance: string;     // B10 是否购买保险
  clearance: string;     // B11 清关方式
  vatCompany: string;    // B12 VAT公司名
  vatAddress: string;    // B13 VAT注册地址
  vatNo: string;         // B14 VAT号
  eoriNo: string;        // B15 EORI号
  vatContact: string;    // B16 VAT公司联系人
  vatPhone: string;      // B17 VAT公司电话
  totalBoxes: number;    // B18 总箱数
  totalWeight: number;   // B19 预计总重量
  totalVolume: number;   // B20 预计总体积
  recipientName: string;    // F14 收件人姓名
  recipientCompany: string; // F15 收件人公司
  recipientCountry: string; // F16 收件人国家
  recipientCity: string;    // F17 收件人城市
  recipientState: string;   // F18 收件人省份/州
  recipientZip: string;     // F19 收件人邮编
  recipientPhone: string;   // F20 收件人联系方式
  recipientEmail: string;   // F21 收件人邮箱
  recipientAddress: string; // F22 收件人地址
  currency: string;         // 申报币种(取首个明细 K 列, 整票统一)
}

/** 顶层转换结果 */
export interface SiruiConvertResult {
  sourceFile: string;
  fileName: string;        // 客户订单号.xlsx
  supplier: SiruiSupplierKey;
  customerOrderNo: string;
  totalBoxes: number;
  dataRows: number;
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
    const o = v as { richText?: { text: string }[]; text?: unknown; result?: unknown };
    if (o.richText) return o.richText.map((r) => r.text).join("").trim();
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
  // 合并单元格的从属格(非左上主格)在 ExcelJS 中 type=Merge 且会继承主格值。
  // 思锐源用合并表达「混箱/共享」，从属格实际为空，此处视为 0。
  if (cell.type === ExcelJS.ValueType.Merge) return 0;
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

/** 去除浮点噪声(保留约 15 位有效数字，对齐 Excel/WPS 精度)。14.7/12 → 1.225 而非 1.2249999999999999 */
function clean(n: number): number {
  if (!isFinite(n)) return n;
  return Number(n.toPrecision(15));
}

function sanitizeFileName(name: string): string {
  const cleaned = name.replace(/[\\/:*?"<>|]/g, "_").trim();
  return cleaned || "未命名";
}

/** 提取 DISPIMG 单元格里的图片 ID(如 ID_xxx)，非 DISPIMG 返回空串 */
function dispimgIdFromCell(cell: ExcelJS.Cell): string {
  const v = cell.value;
  const formula =
    typeof v === "object" && v !== null ? ((v as { formula?: unknown }).formula as string | undefined) : undefined;
  const text = formula || cellText(cell);
  const m = /DISPIMG\("([^"]+)"/.exec(text);
  return m ? m[1] : "";
}

/** 海关编码：纯数字转数值(避免文本编码)，否则原样 */
function hsCodeFor(code: string): string | number {
  if (/^\d+$/.test(code)) return Number(code);
  return code;
}

/** 币种 → 华威尔模板中文币种(美金/英镑/欧元) */
function currencyCn(currency: string): string {
  const c = (currency || "").trim().toUpperCase();
  if (c === "USD" || c === "美金" || c === "美元") return "美金";
  if (c === "GBP" || c === "英镑") return "英镑";
  if (c === "EUR" || c === "欧元") return "欧元";
  return c || "美金";
}

/** 品牌类型 → 华威尔品牌类型下拉值(empty/自主品牌 → 境内自主品牌) */
function brandTypeFor(brandType: string): string {
  const b = (brandType || "").trim();
  if (!b || b === "自主品牌") return "境内自主品牌";
  return b;
}

/** 从地址串提取美国州缩写(如 "29 Commerce Ct, Verona, NJ 07044" → "NJ") */
function extractState(address: string): string {
  const m = /([A-Z]{2})\s+\d{4,5}/.exec(address || "");
  return m ? m[1] : "";
}

/** 地址第一行(逗号分隔首段) */
function addressLine1(address: string): string {
  return (address || "").split(/[,，]/)[0].trim();
}

/** 联系电话：取第一行(去掉换行备注) */
function phoneFirstLine(phone: string): string {
  return (phone || "").split(/\r?\n/)[0].trim();
}

const COUNTRY_CN_TO_CODE: Record<string, string> = {
  美国: "US", 德国: "DE", 英国: "GB", 法国: "FR", 意大利: "IT", 西班牙: "ES",
  日本: "JP", 加拿大: "CA", 澳大利亚: "AU", 荷兰: "NL", 墨西哥: "MX",
  比利时: "BE", 波兰: "PL", 捷克: "CZ", 奥地利: "AT", 瑞典: "SE", 丹麦: "DK",
  芬兰: "FI", 爱尔兰: "IE", 葡萄牙: "PT", 希腊: "GR", 匈牙利: "HU", 罗马尼亚: "RO",
  阿联酋: "AE", 迪拜: "AE",
};

const COUNTRY_EN_TO_CODE: Record<string, string> = {
  "United States": "US", USA: "US", America: "US",
  "United Kingdom": "GB", UK: "GB", Germany: "DE", France: "FR", Italy: "IT",
  Spain: "ES", Japan: "JP", Canada: "CA", Australia: "AU", Netherlands: "NL", Mexico: "MX",
};

const COUNTRY_CODE_TO_CN: Record<string, string> = {
  US: "美国", DE: "德国", GB: "英国", FR: "法国", IT: "意大利", ES: "西班牙",
  JP: "日本", CA: "加拿大", AU: "澳大利亚", NL: "荷兰", MX: "墨西哥", BE: "比利时",
};

/** 国家 → 二字代码(支持中文名/英文名/已有代码) */
function countryCode(country: string): string {
  const c = (country || "").trim();
  if (!c) return "";
  if (/^[A-Z]{2}$/i.test(c)) return c.toUpperCase();
  return COUNTRY_CN_TO_CODE[c] || COUNTRY_EN_TO_CODE[c] || c.toUpperCase();
}

/** 国家 → 中文名 */
function countryCn(country: string): string {
  const c = (country || "").trim();
  const code = countryCode(c);
  return COUNTRY_CODE_TO_CN[code] || c;
}

/** 华威尔「目的地/国家」: 美国 → 美东(默认, 人工可改), 其余用中文国名 */
function huaweiDestination(country: string): string {
  if (countryCode(country) === "US") return "美东";
  return countryCn(country);
}

// ============================================================
// 思锐源解析
// ============================================================

function parseTopInfo(ws: ExcelJS.Worksheet): SiruiTopInfo {
  const s = (ref: string) => cellText(ws.getCell(ref));
  return {
    hasBattery: s("F2"),
    hasMagnet: s("F3"),
    customs: s("B9"),
    insurance: s("B10"),
    clearance: s("B11"),
    vatCompany: s("B12"),
    vatAddress: s("B13"),
    vatNo: s("B14"),
    eoriNo: s("B15"),
    vatContact: s("B16"),
    vatPhone: s("B17"),
    totalBoxes: cellNum(ws.getCell("B18")),
    totalWeight: cellNum(ws.getCell("B19")),
    totalVolume: cellNum(ws.getCell("B20")),
    recipientName: s("F14"),
    recipientCompany: s("F15"),
    recipientCountry: s("F16"),
    recipientCity: s("F17"),
    recipientState: s("F18"),
    recipientZip: s("F19"),
    recipientPhone: s("F20"),
    recipientEmail: s("F21"),
    recipientAddress: s("F22"),
    currency: "",
  };
}

function parseItems(ws: ExcelJS.Worksheet): SiruiItem[] {
  const items: SiruiItem[] = [];
  for (let r = 25; r <= ws.rowCount; r++) {
    const boxId = cellText(ws.getCell(r, 1));
    if (!boxId) break; // 货箱编号为空 = 结束(其后为合计行)
    const vCell = ws.getCell(r, 22);
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
      dispimgId: dispimgIdFromCell(vCell),
      image: null,
    });
  }
  return items;
}

// ============================================================
// 源图片提取
// ============================================================

/**
 * 从源文件提取 WPS DISPIMG 嵌入图(cellimages.xml)。
 * cellimages.xml 中每个 <etc:cellImage> 的 name=DISPIMG ID, r:embed=rId;
 * 经 cellimages.xml.rels 映射到 xl/media/xxx。多图可共用同一 media(去重)。
 */
async function extractDispimgImages(filePath: string): Promise<Map<string, Buffer>> {
  const map = new Map<string, Buffer>();
  try {
    const raw = fs.readFileSync(filePath);
    const zip = await JSZip.loadAsync(raw);

    const ciFile = zip.file("xl/cellimages.xml");
    const relsFile = zip.file("xl/_rels/cellimages.xml.rels");
    if (!ciFile || !relsFile) return map;

    const ciXml = await ciFile.async("string");
    const relsXml = await relsFile.async("string");

    const relMap: Record<string, string> = {};
    for (const m of relsXml.matchAll(/<Relationship[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"/g)) {
      relMap[m[1]] = "xl/" + m[2].replace(/^\.\.\//, "");
    }

    for (const m of ciXml.matchAll(/<etc:cellImage>([\s\S]*?)<\/etc:cellImage>/g)) {
      const block = m[1];
      const nameM = /<xdr:cNvPr[^>]*name="([^"]+)"/.exec(block);
      const embedM = /<a:blip[^>]*r:embed="([^"]+)"/.exec(block);
      if (!nameM || !embedM) continue;
      const target = relMap[embedM[1]];
      if (!target) continue;
      const mediaFile = zip.file(target);
      if (mediaFile && !map.has(nameM[1])) {
        map.set(nameM[1], Buffer.from(await mediaFile.async("nodebuffer")));
      }
    }
  } catch {
    // 图片提取失败不致命(图片列留空)
  }
  return map;
}

/**
 * 从源文件浮动图片提取产品图(航乐: 图片锚定 V 列 col21, 数据行 R25=锚点行24)。
 */
async function extractSourceImages(filePath: string, itemCount: number): Promise<(Buffer | null)[]> {
  const result: (Buffer | null)[] = new Array(itemCount).fill(null);
  try {
    const raw = fs.readFileSync(filePath);
    const zip = await JSZip.loadAsync(raw);

    const drawingFile = zip.file("xl/drawings/drawing1.xml");
    const drawingRelsFile = zip.file("xl/drawings/_rels/drawing1.xml.rels");
    if (!drawingFile || !drawingRelsFile) return result;

    const drawingXml = await drawingFile.async("string");
    const relsXml = await drawingRelsFile.async("string");

    const relMap: Record<string, string> = {};
    for (const m of relsXml.matchAll(/<Relationship[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"/g)) {
      relMap[m[1]] = "xl/" + m[2].replace(/^\.\.\//, "");
    }

    const anchors: { row: number; col: number; rId: string }[] = [];
    // 兼容带属性的 <xdr:twoCellAnchor editAs="oneCell">
    for (const m of drawingXml.matchAll(/<xdr:twoCellAnchor[^>]*>([\s\S]*?)<\/xdr:twoCellAnchor>/g)) {
      const block = m[1];
      const fromCol = /<xdr:from>[\s\S]*?<xdr:col>(\d+)<\/xdr:col>/.exec(block);
      const fromRow = /<xdr:from>[\s\S]*?<xdr:row>(\d+)<\/xdr:row>/.exec(block);
      const embed = /<a:blip[^>]*r:embed="([^"]+)"/.exec(block);
      if (fromCol && fromRow && embed) {
        anchors.push({ col: parseInt(fromCol[1], 10), row: parseInt(fromRow[1], 10), rId: embed[1] });
      }
    }

    for (let i = 0; i < itemCount; i++) {
      const anchorRow = 24 + i; // 数据行 R25 对应 0 基行 24
      const anchor =
        anchors.find((a) => a.row === anchorRow && a.col === 21) ||
        anchors.find((a) => a.row === anchorRow && a.col === 22);
      if (!anchor) continue;
      const target = relMap[anchor.rId];
      if (!target) continue;
      const mediaFile = zip.file(target);
      if (mediaFile) result[i] = Buffer.from(await mediaFile.async("nodebuffer"));
    }
  } catch {
    // 图片提取失败不致命
  }
  return result;
}

// ============================================================
// 供应商配置
// ============================================================

type CellValue = string | number | boolean | null | ExcelJS.CellFormulaValue;

/** mapItem 可变上下文(跨明细行累计) */
interface MapCtx {
  boxSeq: number;              // 累计箱号(华威尔/航乐 explode 每箱+1; 美琦/英美排序号按箱数累加)
  curDim: { l: number; w: number; h: number }; // 当前箱尺寸(航乐混箱继承用)
}

interface SupplierSpec {
  key: SiruiSupplierKey;
  templateFile: string;   // public/templates/sirui-suppliers/ 下文件名
  sheetName: string;      // 数据所在 sheet(华威尔用空串→取首个 sheet)
  headerStartRow: number; // 数据区起始行
  imageCol: number;       // DISPIMG 图片列(0 = 不用 DISPIMG)
  rowHeight?: number;     // 数据行高(图片供应商需要)
  maxCol: number;         // 数据区最大列
  dataEndRow?: number;    // 数据区结束行(含, 其后为合计行), 设置后插行并更新合计公式
  sumCols?: string[];     // 合计行 SUM 公式列(插行后更新范围)
  fillHeader: (
    ws: ExcelJS.Worksheet,
    top: SiruiTopInfo,
    opts: {
      customerOrderNo: string;
      channelName: string;
      customsOverride?: string;
      station?: string;
    },
    items: SiruiItem[]
  ) => void;
  mapItem: (item: SiruiItem, top: SiruiTopInfo, ctx: MapCtx) => CellValue[][];
  styleDataRow?: (row: ExcelJS.Row, maxCol: number) => void; // 数据行逐列样式(模板数据区无格式时补全)
  countryVariants?: {
    match: (country: string) => boolean;
    templateFile: string;
    maxCol?: number;
    fillHeader?: SupplierSpec["fillHeader"];
  }[];
}

/** 报关方式自动映射(源 B9 有单证/无单证) */
function customsFor(supplier: SiruiSupplierKey, customs: string): string {
  if (customs === "有单证") {
    if (supplier === "华威尔") return "单独报关";
    if (supplier === "美琦") return "一般贸易";
    if (supplier === "英美") return "报关退税";
    if (supplier === "航乐") return "报关退税";
  }
  if (customs === "无单证") {
    if (supplier === "华威尔") return "买单报关";
    if (supplier === "美琦") return "代理报关";
    if (supplier === "英美") return "买单报关";
    if (supplier === "航乐") return "委托报关";
  }
  return customs;
}

/** 报关方式下拉选项(与各模板 data validation 值域一致) */
const CUSTOMS_OPTIONS: Record<SiruiSupplierKey, string[]> = {
  华威尔: ["单独报关", "买单报关", "合并单独报关"],
  美琦: ["一般贸易", "代理报关", "手册报关", "9710电商深圳报关单一窗口", "9810报关"],
  英美: ["报关退税", "买单报关", "合并报关"],
  航乐: ["报关退税", "委托报关"],
};

export function getSupplierCustomsOptions(supplier: SiruiSupplierKey): string[] {
  return CUSTOMS_OPTIONS[supplier] ?? [];
}

/** 清关方式 → 华威尔 D3 下拉(包税/不包税) */
function huaweiClearance(clearance: string): string {
  const c = (clearance || "").trim();
  if (c === "自税" || c === "不包税" || c === "不包税递延") return "不包税";
  return "包税"; // 包税 / 空 → 包税
}

/** 清关方式 → 英美 B18(包税→否, 其余→是) */
function yingmeiClearance(clearance: string): string {
  const c = (clearance || "").trim();
  if (c === "包税") return "否";
  return "是";
}

/** 清关方式 → 航乐 L7(自税→不包税递延, 不包税→不包税, 包税→包税) */
function hangyueClearance(clearance: string): string {
  const c = (clearance || "").trim();
  if (c === "自税") return "不包税递延";
  if (c === "不包税") return "不包税";
  return "包税";
}

/** 单箱数量 = 总数量 / 箱数(混箱 G=0 时为总数量) */
function perBoxQtyOf(item: SiruiItem): number {
  const n = item.boxCount >= 1 ? item.boxCount : 1;
  return clean(item.totalQty / n);
}

/**
 * 尺寸(长宽高)取值：
 *  - 开新箱(boxCount>=1)：用源尺寸并记录为当前箱尺寸
 *  - 混箱(boxCount=0)：源尺寸非空(0)用自身尺寸, 否则继承当前箱尺寸
 *    英美混箱行源尺寸留空 → 继承箱尺寸; 美琦混箱行源尺寸已填 → 用自身尺寸
 */
function dimOf(item: SiruiItem, ctx: MapCtx): { l: number; w: number; h: number } {
  if (item.boxCount >= 1) {
    ctx.curDim = { l: item.lengthCm, w: item.widthCm, h: item.heightCm };
    return ctx.curDim;
  }
  if (item.lengthCm > 0 || item.widthCm > 0 || item.heightCm > 0) {
    return { l: item.lengthCm, w: item.widthCm, h: item.heightCm };
  }
  return ctx.curDim;
}

/**
 * 英美地址块填充(默认模板与美线模板共用)。
 * 模板地址字段为 VLOOKUP(按 B2 地址库编码带出)，商业/私人地址需清空 B2 并手动填入。
 * 注意：空值也要显式写 null，否则 VLOOKUP 公式不会被清除。
 */
function fillYingmeiAddress(ws: ExcelJS.Worksheet, top: SiruiTopInfo): void {
  // 清掉模板 K 列(带电/带磁/报关/清关)残留示例值, 答案样本该列为空
  for (const r of [15, 16, 17, 18]) ws.getCell(r, 11).value = null;
  const state = top.recipientState || extractState(top.recipientAddress);
  const setB = (r: number, v: string | number | null) => {
    ws.getCell(r, 2).value = v == null || v === "" ? null : v;
  };
  setB(2, null);                              // 地址库编码 清空(商业地址)
  setB(3, top.recipientName);                 // 收件人姓名
  setB(4, top.recipientCompany);              // 收件人公司
  setB(5, addressLine1(top.recipientAddress)); // 地址一
  setB(6, null);                              // 地址二
  setB(7, null);                              // 地址三
  setB(8, top.recipientCity);                 // 城市
  setB(9, state);                             // 省份/州
  setB(10, top.recipientZip);                 // 邮编
  setB(11, countryCode(top.recipientCountry)); // 国家代码
  setB(12, phoneFirstLine(top.recipientPhone)); // 电话
  setB(13, top.recipientEmail);               // 邮箱
}

// 华威尔数据区逐列样式(与答案样本一致: 微软雅黑11/16 + Arial18, 中间列居中, V列无边框带换行)
const HUAWEI_COL_STYLES: { font: string; size: number; horiz: "left" | "center"; numFmt: string; border: boolean }[] = [
  { font: "微软雅黑", size: 11, horiz: "left",   numFmt: "@",       border: true },   // A 箱唛号
  { font: "微软雅黑", size: 16, horiz: "center", numFmt: "General", border: true },   // B Reference ID
  { font: "微软雅黑", size: 16, horiz: "center", numFmt: "General", border: true },   // C 中文品名
  { font: "微软雅黑", size: 16, horiz: "center", numFmt: "General", border: true },   // D 英文品名
  { font: "微软雅黑", size: 16, horiz: "center", numFmt: "General", border: true },   // E 材质
  { font: "微软雅黑", size: 16, horiz: "center", numFmt: "General", border: true },   // F 用途
  { font: "微软雅黑", size: 11, horiz: "left",   numFmt: "General", border: true },   // G 海关编码
  { font: "微软雅黑", size: 11, horiz: "left",   numFmt: "@",       border: true },   // H 产品类型
  { font: "微软雅黑", size: 16, horiz: "center", numFmt: "General", border: true },   // I 单箱产品数量
  { font: "Arial",   size: 18, horiz: "center", numFmt: "0_",      border: true },   // J 箱数
  { font: "微软雅黑", size: 11, horiz: "left",   numFmt: "General", border: true },   // K 币种
  { font: "微软雅黑", size: 11, horiz: "left",   numFmt: "General", border: true },   // L 单个产品申报单价
  { font: "微软雅黑", size: 11, horiz: "left",   numFmt: "General", border: true },   // M 单个产品净重
  { font: "微软雅黑", size: 11, horiz: "left",   numFmt: "General", border: true },   // N 产品图片
  { font: "微软雅黑", size: 11, horiz: "left",   numFmt: "@",       border: true },   // O 品牌
  { font: "微软雅黑", size: 11, horiz: "left",   numFmt: "@",       border: true },   // P 品牌类型
  { font: "微软雅黑", size: 11, horiz: "left",   numFmt: "@",       border: true },   // Q 型号
  { font: "微软雅黑", size: 11, horiz: "left",   numFmt: "General", border: true },   // R 销售链接
  { font: "微软雅黑", size: 11, horiz: "left",   numFmt: "@",       border: true },   // S 备注
  { font: "微软雅黑", size: 11, horiz: "left",   numFmt: "General", border: true },   // T 总产品数量
  { font: "微软雅黑", size: 11, horiz: "left",   numFmt: "General", border: true },   // U 总申报金额
  { font: "微软雅黑", size: 11, horiz: "left",   numFmt: "General", border: false },  // V 毛重(无边框)
];

const SUPPLIERS: Record<SiruiSupplierKey, SupplierSpec> = {
  // ── 华威尔(一箱一行 explode) ─────────────────────────────
  华威尔: {
    key: "华威尔",
    templateFile: "华威尔-模板.xlsx",
    sheetName: "", // 取首个 sheet「发票箱单」
    headerStartRow: 18,
    imageCol: 14, // N 列 产品高清图片
    rowHeight: 77,
    maxCol: 22,
    fillHeader(ws, top, opts) {
      const state = top.recipientState || extractState(top.recipientAddress);
      const b = (r: number, v: string | number) => {
        if (v !== "" && v !== null && v !== undefined) ws.getCell(r, 2).value = v;
      };
      const d = (r: number, v: string | number) => {
        if (v !== "" && v !== null && v !== undefined) ws.getCell(r, 4).value = v;
      };
      b(1, opts.customerOrderNo);            // 客户原单号
      b(2, opts.station || "福永仓");        // 交货仓库
      d(2, top.insurance || "否");           // 是否购买保险
      b(3, "HW01870");                       // 交货客户代号
      d(3, huaweiClearance(top.clearance));  // 清关方式
      b(4, opts.channelName);                // 走货线路
      b(5, opts.customsOverride || customsFor("华威尔", top.customs)); // 报关方式
      b(6, huaweiDestination(top.recipientCountry)); // 目的地/国家
      b(7, "商业地址");                       // 地址类型
      b(8, top.recipientCompany);            // 收货地址名称
      b(9, top.recipientCompany);            // 收件公司
      b(10, top.recipientName);              // 收件人
      b(11, phoneFirstLine(top.recipientPhone)); // 联系电话
      b(12, top.recipientZip);               // 邮编
      b(13, state);                          // 省份/州
      b(14, top.recipientCity);              // 城市
      b(15, top.recipientAddress);           // 收件地址
      b(16, top.totalBoxes);                 // 总件数
    },
    mapItem(item, top, ctx) {
      const n = item.boxCount >= 1 ? Math.round(item.boxCount) : 1;
      const perBox = perBoxQtyOf(item);
      const netPer = perBox ? clean(item.netWeight / perBox) : 0;
      const rows: CellValue[][] = [];
      for (let k = 0; k < n; k++) {
        if (item.boxCount >= 1) ctx.boxSeq++;
        const boxMark = `${item.boxId}U${String(ctx.boxSeq).padStart(2, "0")}`;
        rows.push([
          boxMark,                    // A 箱唛号
          item.refId,                 // B Reference ID
          item.nameCh,                // C 中文品名
          item.nameEn,                // D 英文品名
          item.material,              // E 材质
          item.use,                   // F 用途
          hsCodeFor(item.hsCode),     // G 海关编码
          "普货",                     // H 产品类型/属性
          perBox,                     // I 单箱产品数量
          1,                          // J 箱数(explode 后每箱一行)
          currencyCn(item.currency),  // K 币种
          item.unitPrice,             // L 单个产品申报单价
          netPer,                     // M 单个产品净重
          null,                       // N 产品图片(DISPIMG)
          item.brand,                 // O 品牌
          brandTypeFor(item.brandType), // P 品牌类型
          item.model,                 // Q 型号
          item.link || 0,             // R 销售链接(空→0)
          "",                         // S 备注
          perBox,                     // T 总产品数量
          clean(perBox * item.unitPrice), // U 总申报金额
          item.grossWeight,           // V 毛重
        ]);
      }
      return rows;
    },
    styleDataRow(row, maxCol) {
      for (let c = 1; c <= maxCol; c++) {
        const st = HUAWEI_COL_STYLES[c - 1];
        if (!st) continue;
        const cell = row.getCell(c);
        // 整体替换 style 以打断 ExcelJS 从模板读入的共享 style 引用(逐个赋值会串列)
        cell.style = {
          font: { name: st.font, size: st.size, bold: false },
          alignment: { horizontal: st.horiz, vertical: "middle", wrapText: !st.border },
          numFmt: st.numFmt,
          border: st.border
            ? { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } }
            : {},
        };
      }
    },
  },

  // ── 美琦(排序号累计区间) ─────────────────────────────────
  美琦: {
    key: "美琦",
    templateFile: "美琦-美线模板-2026.xlsx",
    sheetName: "发票",
    headerStartRow: 19,
    imageCol: 16, // P 列 产品图片(源无图片则留空)
    rowHeight: 18,
    maxCol: 19,
    fillHeader(ws, top, opts) {
      const state = top.recipientState || extractState(top.recipientAddress);
      const b = (r: number, v: string | number) => {
        if (v !== "" && v !== null && v !== undefined) ws.getCell(r, 2).value = v;
      };
      const set = (ref: string, v: string | number | null) => {
        if (v !== "" && v !== null && v !== undefined) ws.getCell(ref).value = v;
      };
      b(1, opts.customerOrderNo);   // 客户订单号
      b(2, opts.customerOrderNo);   // 客户参考号
      b(3, opts.channelName);       // 服务
      b(5, top.recipientName);      // 收件人姓名
      b(6, addressLine1(top.recipientAddress)); // 收件人地址一
      b(7, top.recipientCity);      // 城市
      b(8, state);                  // 省/洲
      b(9, top.recipientZip);       // 邮编
      b(10, countryCode(top.recipientCountry)); // 国家代码
      b(11, phoneFirstLine(top.recipientPhone)); // 电话
      b(12, top.recipientEmail);    // 邮箱
      b(13, top.recipientCompany);  // 公司
      b(14, top.hasBattery || "否");  // 带电(默认否)
      b(15, top.hasMagnet || "否");   // 带磁(默认否)
      b(16, opts.customsOverride || customsFor("美琦", top.customs)); // 报关方式
      b(17, top.totalBoxes);        // 箱数
      set("F1", opts.station || "中山仓"); // 预计交货站点
      set("F3", "自送货");          // 提货方式
    },
    mapItem(item, top, ctx) {
      const perBox = perBoxQtyOf(item);
      let sortNo: string;
      if (item.boxCount >= 1) {
        const start = ctx.boxSeq + 1;
        const end = ctx.boxSeq + Math.round(item.boxCount);
        sortNo = Math.round(item.boxCount) === 1 ? String(start) : `${start}-${end}`;
        ctx.boxSeq = end;
      } else {
        sortNo = String(ctx.boxSeq); // 混箱共享当前累计箱号
      }
      const { l, w, h } = dimOf(item, ctx);
      return [[
        sortNo,                         // A 货箱编号(排序号)
        item.nameEn,                    // B 英文品名
        item.nameCh,                    // C 中文品名
        "普货",                         // D 产品性质
        perBox,                         // E 单箱数量
        item.unitPrice,                 // F 单价(美金)
        item.grossWeight,               // G 货箱重量
        clean(l - 1),                   // H 长 -1(美琦特有)
        clean(w - 1),                   // I 宽 -1
        clean(h - 1),                   // J 高 -1
        hsCodeFor(item.hsCode),         // K 海关编码
        item.brand,                     // L 品牌
        item.material,                  // M 材质
        item.model,                     // N 型号
        item.use,                       // O 用途
        null,                           // P 产品图片(DISPIMG)
        item.refId,                     // Q PO Number
        item.boxCount >= 1 ? Math.round(item.boxCount) : 0, // R 物品箱号(源箱数原样透传, 混箱→0)
        item.boxId,                     // S 物品FBA ID
      ]];
    },
  },

  // ── 英美(排序号累计区间, 双模板) ─────────────────────────
  英美: {
    key: "英美",
    templateFile: "英美-模板.xlsx",
    sheetName: "模板",
    headerStartRow: 26,
    imageCol: 0,
    maxCol: 26,
    fillHeader(ws, top, opts) {
      fillYingmeiAddress(ws, top);
      const b = (r: number, v: string | number) => {
        if (v !== "" && v !== null && v !== undefined) ws.getCell(r, 2).value = v;
      };
      b(1, opts.channelName);       // 服务
      ws.getCell(14, 2).value = opts.customerOrderNo || null; // 客户订单号(清空示例)
      b(15, top.hasBattery || "否"); // 带电
      b(16, top.hasMagnet || "否");  // 带磁
      b(17, opts.customsOverride || customsFor("英美", top.customs)); // 报关方式
      b(18, yingmeiClearance(top.clearance)); // 清关方式
      b(19, top.vatNo);             // VAT号
      b(20, top.eoriNo);            // EORI
      b(21, top.currency);          // 币种
      b(22, opts.station || "");    // 所在货站
      b(24, top.totalBoxes);        // 箱数
    },
    mapItem(item, top, ctx) {
      const perBox = perBoxQtyOf(item);
      let sortNo: number | string;
      if (item.boxCount >= 1) {
        const start = ctx.boxSeq + 1;
        const end = ctx.boxSeq + Math.round(item.boxCount);
        // 英美单箱/混箱排序号用数字, 区间用字符串(与答案样本一致)
        sortNo = Math.round(item.boxCount) === 1 ? start : `${start}-${end}`;
        ctx.boxSeq = end;
      } else {
        sortNo = ctx.boxSeq;
      }
      const { l, w, h } = dimOf(item, ctx);
      return [[
        sortNo,                     // A 货箱编号(排序号)
        item.refId || 0,            // B PO Number(空→0)
        item.grossWeight,           // C 货箱重量
        l,                          // D 长
        w,                          // E 宽
        h,                          // F 高
        item.model,                 // G 产品SKU
        item.nameCh,                // H 中文品名
        item.nameEn,                // I 英文品名
        item.unitPrice,             // J 单价
        perBox,                     // K 数量
        item.material,              // L 材质
        hsCodeFor(item.hsCode),     // M 海关编码
        item.use,                   // N 用途
        item.brand,                 // O 品牌
        item.model,                 // P 型号
        item.link,                  // Q 销售链接
        null,                       // R 销售价格
        null,                       // S 图片链接
        null,                       // T 重量
        null,                       // U ASIN
        null,                       // V FNSKU
        null,                       // W 承运商
        null,                       // X 跟踪号
        null,                       // Y 产品尺寸
        null,                       // Z 产地
      ]];
    },
    // 美线(美国)模板: 无 EORI, R20 为币种, 数据区 24 列(无 Y/Z)
    countryVariants: [
      {
        match: (c) => countryCode(c) === "US",
        templateFile: "英美-美国模板.xlsx",
        maxCol: 24,
        fillHeader(ws, top, opts) {
          fillYingmeiAddress(ws, top);
          const b = (r: number, v: string | number) => {
            if (v !== "" && v !== null && v !== undefined) ws.getCell(r, 2).value = v;
          };
          b(1, opts.channelName);
          ws.getCell(14, 2).value = opts.customerOrderNo || null;
          b(15, top.hasBattery || "否");
          b(16, top.hasMagnet || "否");
          b(17, opts.customsOverride || customsFor("英美", top.customs));
          b(18, yingmeiClearance(top.clearance));
          b(19, top.vatNo);
          b(20, top.currency);      // 美线无 EORI, R20 即币种
          b(21, opts.station || "");
          b(24, top.totalBoxes);
        },
      },
    ],
  },

  // ── 航乐(非FBA 一箱一行 explode) ─────────────────────────
  航乐: {
    key: "航乐",
    templateFile: "航乐-模板.xlsx",
    sheetName: "Packing list装箱单发票",
    headerStartRow: 12,
    imageCol: 23, // W 列 产品图片
    rowHeight: 30,
    maxCol: 24,
    dataEndRow: 23, // 数据区 R12~R23(R24 为合计行)
    sumCols: ["K", "M", "O", "S", "T"],
    fillHeader(ws, top, opts) {
      const set = (ref: string, v: string | number | null) => {
        if (v !== "" && v !== null && v !== undefined) ws.getCell(ref).value = v;
      };
      set("B3", top.recipientAddress);            // ADD 收件地址
      set("G3", countryCode(top.recipientCountry)); // COUNTRY
      // H3/K3 为模板标签「SHIPMENT ID/AMAZON REFERENCE ID」, 非FBA 无值, 保持原标签
      set("B4", Number(top.recipientZip) || top.recipientZip); // ZIP
      set("D4", top.recipientCompany);            // COMPANY
      set("G4", phoneFirstLine(top.recipientPhone)); // TEL
      set("B5", top.recipientCity);               // CITY
      set("D5", top.recipientName);               // ATTN
      set("G5", top.recipientEmail);              // EMAIL
      set("I5", opts.customsOverride || customsFor("航乐", top.customs)); // 报关类型
      set("K5", opts.station || "中山航乐仓");    // 交货仓库
      set("B7", top.vatCompany || top.recipientCompany); // 公司名称
      set("I7", opts.channelName);                // 渠道
      set("L7", hangyueClearance(top.clearance)); // 是否包税
      set("B8", top.vatNo);                       // VAT号
      set("B9", top.eoriNo);                      // EORI号
      set("B10", top.vatAddress.replace(/\n/g, " ")); // VAT公司注册地址
    },
    mapItem(item, top, ctx) {
      const n = item.boxCount >= 1 ? Math.round(item.boxCount) : 1;
      const perBox = perBoxQtyOf(item);
      const netPer = perBox ? clean(item.netWeight / perBox) : 0;
      const { l, w, h } = dimOf(item, ctx);
      const dimVol = l * w * h;
      const rows: CellValue[][] = [];
      for (let k = 0; k < n; k++) {
        if (item.boxCount >= 1) ctx.boxSeq++;
        rows.push([
          ctx.boxSeq,                 // A Box No
          item.nameCh,                // B 品名(中文)
          item.nameEn,                // C 品名(英文)
          item.material,              // D 材质(中文)
          item.material,              // E 材质(英文)
          hsCodeFor(item.hsCode),     // F 海关编码
          item.model,                 // G 型号
          item.brand,                 // H 品牌
          item.use,                   // I 用途(英文)
          perBox,                     // J 单箱数量
          perBox,                     // K 总数量
          netPer,                     // L 净重(单个产品)
          item.grossWeight,           // M 实重
          item.unitPrice,             // N 单价(EUR)
          clean(perBox * item.unitPrice), // O 总价(EUR)
          l,                          // P 长
          w,                          // Q 宽
          h,                          // R 高
          clean(dimVol / 6000),       // S 材重
          clean(dimVol / 1000000),    // T CBM
          top.hasBattery,             // U 是否带电
          top.hasMagnet,              // V 是否带磁
          null,                       // W 产品图片(DISPIMG)
          item.link,                  // X 销售链接
        ]);
      }
      return rows;
    },
  },
};

// ============================================================
// 渠道下拉数据(从供应商模板读取, 单一数据源)
// ============================================================

interface ChannelSheetSpec {
  sheet: string;
  col: number;
  startRow: number;
  templateFile?: string;
}

const CHANNEL_SHEETS: Record<SiruiSupplierKey, ChannelSheetSpec[]> = {
  华威尔: [{ sheet: "走货线路名和申报币别要求-不能删", col: 1, startRow: 2 }],
  美琦: [{ sheet: "服务渠道", col: 2, startRow: 2 }],
  // 英美双模板: 英欧(默认, 渠道列表 B 列 R2 起) + 美线(渠道组 A 列 R2 起)
  英美: [
    { sheet: "渠道列表", col: 2, startRow: 2 },
    { templateFile: "英美-美国模板.xlsx", sheet: "渠道组", col: 1, startRow: 2 },
  ],
  航乐: [
    { sheet: "Sheet2", col: 1, startRow: 1 },
    { sheet: "Sheet1", col: 1, startRow: 1 },
  ],
};

const channelsCache = new Map<SiruiSupplierKey, string[]>();

export async function getSupplierChannels(supplier: SiruiSupplierKey): Promise<string[]> {
  const cached = channelsCache.get(supplier);
  if (cached) return cached;

  const spec = SUPPLIERS[supplier];
  const sheets = CHANNEL_SHEETS[supplier];
  if (!spec || !sheets) throw new Error(`未知供应商: ${supplier}`);

  const seen = new Set<string>();
  const channels: string[] = [];

  const byFile = new Map<string, ChannelSheetSpec[]>();
  for (const cs of sheets) {
    const file = cs.templateFile ?? spec.templateFile;
    if (!byFile.has(file)) byFile.set(file, []);
    byFile.get(file)!.push(cs);
  }

  for (const [file, css] of byFile) {
    const tplPath = path.join(process.cwd(), "public", "templates", "sirui-suppliers", file);
    if (!fs.existsSync(tplPath)) throw new Error(`供应商模板不存在: ${tplPath}`);

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(tplPath);

    for (const cs of css) {
      const ws = wb.getWorksheet(cs.sheet);
      if (!ws) continue;
      for (let r = cs.startRow; r <= ws.rowCount; r++) {
        const name = cellText(ws.getCell(r, cs.col));
        if (name && !seen.has(name)) {
          seen.add(name);
          channels.push(name);
        }
      }
    }
  }

  channelsCache.set(supplier, channels);
  return channels;
}

// ============================================================
// 送货货站/交货仓库下拉数据
// ============================================================

const HUAWEI_STATIONS = ["福永仓", "广州仓", "坂田仓", "义乌仓"];
const MEIQI_STATIONS = ["中山仓", "清溪仓", "义乌仓", "广州仓", "厦门仓", "汕头仓", "青岛仓"];
const HANGYUE_STATIONS = ["广州白云仓", "深圳平湖仓", "深圳福永仓", "浙江义乌仓", "中山航乐仓"];

const stationsCache = new Map<SiruiSupplierKey, string[]>();

export async function getSupplierStations(supplier: SiruiSupplierKey): Promise<string[]> {
  const cached = stationsCache.get(supplier);
  if (cached) return cached;

  let stations: string[] = [];
  if (supplier === "华威尔") {
    stations = [...HUAWEI_STATIONS];
  } else if (supplier === "美琦") {
    stations = [...MEIQI_STATIONS];
  } else if (supplier === "航乐") {
    stations = [...HANGYUE_STATIONS];
  } else if (supplier === "英美") {
    const spec = SUPPLIERS[supplier];
    const tplPath = path.join(process.cwd(), "public", "templates", "sirui-suppliers", spec.templateFile);
    if (!fs.existsSync(tplPath)) throw new Error(`供应商模板不存在: ${tplPath}`);
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(tplPath);
    const ws = wb.getWorksheet("站点");
    if (ws) {
      const seen = new Set<string>();
      for (let r = 2; r <= ws.rowCount; r++) { // R1 为表头「站点」, 从 R2 起读
        const name = cellText(ws.getCell(r, 1));
        if (name && !seen.has(name)) {
          seen.add(name);
          stations.push(name);
        }
      }
    }
  }

  stationsCache.set(supplier, stations);
  return stations;
}

// ============================================================
// DISPIMG 注入
// ============================================================

const DISPIMG_NS = {
  xdr: "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  etc: "http://www.wps.cn/officeDocument/2017/etCustomData",
};

const CELLIMAGE_REL_TYPE = "http://www.wps.cn/officeDocument/2020/cellImage";
const CELLIMAGE_CONTENT_TYPE = "application/vnd.wps-officedocument.cellimage+xml";

function newImageId(): string {
  return "ID_" + crypto.randomBytes(16).toString("hex").toUpperCase();
}

/** 图片缓冲 → 扩展名(按魔数) */
function imageExt(buf: Buffer): string {
  if (buf.length >= 3 && buf[0] === 0xff && buf[1] === 0xd8) return "jpeg";
  if (buf.length >= 8 && buf[0] === 0x89 && buf[1] === 0x50 && buf[2] === 0x4e && buf[3] === 0x47) return "png";
  if (buf.length >= 3 && buf[0] === 0x47 && buf[1] === 0x49 && buf[2] === 0x46) return "gif";
  return "png";
}

function buildCellImagesXml(entries: { id: string; rId: string; itemIndex: number }[]): string {
  const pics = entries
    .map(
      (e, i) =>
        `<etc:cellImage><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${i + 1}" name="${e.id}" descr="Picture"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="${e.rId}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="26471880" y="${5499735 + e.itemIndex * 956945}"/><a:ext cx="744220" cy="956945"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic></etc:cellImage>`
    )
    .join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<etc:cellImages xmlns:xdr="${DISPIMG_NS.xdr}" xmlns:r="${DISPIMG_NS.r}" xmlns:a="${DISPIMG_NS.a}" xmlns:etc="${DISPIMG_NS.etc}">${pics}</etc:cellImages>`;
}

function buildCellImagesRels(entries: { rId: string; target: string }[]): string {
  const rels = entries
    .map(
      (e) =>
        `<Relationship Id="${e.rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${e.target}"/>`
    )
    .join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`;
}

/**
 * 向 xlsx buffer 注入 DISPIMG 图片部件(media + cellimages.xml + 关系)。
 * images/imageIds 按「输出行」索引(explode 后一行一 ID)；相同图片字节去重共享同一 media 文件。
 */
async function injectDispimg(
  buffer: Buffer,
  images: (Buffer | null)[],
  imageIds: (string | null)[]
): Promise<Buffer> {
  const rows: { img: Buffer; id: string; index: number }[] = [];
  images.forEach((img, i) => {
    if (img && imageIds[i]) rows.push({ img, id: imageIds[i]!, index: i });
  });
  if (rows.length === 0) return buffer;

  const zip = await JSZip.loadAsync(buffer);

  // 按图片字节去重 media
  const hashOf = (b: Buffer) => crypto.createHash("sha1").update(b).digest("hex");
  const mediaByHash = new Map<string, string>(); // hash -> mediaName
  const relByMedia = new Map<string, string>();  // mediaName -> rId

  const entries: { id: string; rId: string; itemIndex: number }[] = [];
  const relEntries: { rId: string; target: string }[] = [];

  rows.forEach((row, n) => {
    const hash = hashOf(row.img);
    let mediaName = mediaByHash.get(hash);
    if (!mediaName) {
      const ext = imageExt(row.img);
      mediaName = `image${mediaByHash.size + 1}.${ext}`;
      mediaByHash.set(hash, mediaName);
      zip.file(`xl/media/${mediaName}`, row.img);
    }
    let rId = relByMedia.get(mediaName);
    if (!rId) {
      rId = `rId${relByMedia.size + 1}`;
      relByMedia.set(mediaName, rId);
      relEntries.push({ rId, target: `media/${mediaName}` });
    }
    entries.push({ id: row.id, rId, itemIndex: n });
  });

  zip.file("xl/cellimages.xml", buildCellImagesXml(entries));
  zip.file("xl/_rels/cellimages.xml.rels", buildCellImagesRels(relEntries));

  // workbook.xml.rels 加 cellImage 关系
  const wbRelsFile = zip.file("xl/_rels/workbook.xml.rels");
  if (wbRelsFile) {
    let rels = await wbRelsFile.async("string");
    if (!rels.includes("cellImage")) {
      const ids = Array.from(rels.matchAll(/Id="rId(\d+)"/g)).map((m) => parseInt(m[1], 10));
      const next = (ids.length ? Math.max(...ids) : 0) + 1;
      rels = rels.replace(
        "</Relationships>",
        `<Relationship Id="rId${next}" Type="${CELLIMAGE_REL_TYPE}" Target="cellimages.xml"/></Relationships>`
      );
      zip.file("xl/_rels/workbook.xml.rels", rels);
    }
  }

  // [Content_Types].xml 加 Override
  const ctFile = zip.file("[Content_Types].xml");
  if (ctFile) {
    let ct = await ctFile.async("string");
    if (!ct.includes("cellimages.xml")) {
      ct = ct.replace(
        "</Types>",
        `<Override PartName="/xl/cellimages.xml" ContentType="${CELLIMAGE_CONTENT_TYPE}"/></Types>`
      );
      zip.file("[Content_Types].xml", ct);
    }
  }

  const out = await zip.generateAsync({ type: "nodebuffer" });
  return Buffer.isBuffer(out) ? out : Buffer.from(out);
}

/**
 * ExcelJS writeBuffer 会把模板默认字体(宋体 charset=134)丢失成 Calibri。
 * 从模板 styles.xml 提取首个 <font>(默认字体), 替换输出 styles.xml 的首个 <font>。
 */
async function applyTemplateDefaultFont(buffer: Buffer, tplPath: string): Promise<Buffer> {
  const tplBuf = await fs.promises.readFile(tplPath);
  const tplZip = await JSZip.loadAsync(tplBuf);
  const tplStyles = await tplZip.file("xl/styles.xml")?.async("string");
  const tplFont = tplStyles?.match(/<font>[\s\S]*?<\/font>/)?.[0];
  if (!tplFont) return buffer;

  const zip = await JSZip.loadAsync(buffer);
  const styles = await zip.file("xl/styles.xml")?.async("string");
  if (!styles) return buffer;
  const replaced = styles.replace(/<font>[\s\S]*?<\/font>/, tplFont);
  zip.file("xl/styles.xml", replaced);
  return Buffer.from(await zip.generateAsync({ type: "nodebuffer" }));
}

// ============================================================
// 数据区覆盖式填充辅助
// ============================================================

function findTemplateDataLastRow(ws: ExcelJS.Worksheet, dataStart: number, maxCol: number): number {
  let last = dataStart - 1;
  for (let r = dataStart; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    let hasFormat = false;
    for (let c = 1; c <= maxCol; c++) {
      const cell = row.getCell(c);
      if (cell.value != null && String(cell.value).trim() !== "") {
        hasFormat = true;
        break;
      }
      const b = cell.border || {};
      if (b.left?.style || b.right?.style || b.top?.style || b.bottom?.style) {
        hasFormat = true;
        break;
      }
    }
    if (hasFormat) last = r;
    else if (r > dataStart) break;
  }
  return last;
}

// ============================================================
// 主入口
// ============================================================

export async function convertSiruiToSupplier(
  filePath: string,
  sourceFileName: string,
  supplier: SiruiSupplierKey,
  customerOrderNo: string,
  channelName: string,
  customsOverride?: string,
  station?: string
): Promise<SiruiConvertResult> {
  const baseSpec = SUPPLIERS[supplier];
  if (!baseSpec) throw new Error(`未知供应商: ${supplier}`);

  const orderNo = (customerOrderNo || "").trim();

  // ── 1. 解析思锐源 ──
  const srcWb = new ExcelJS.Workbook();
  await srcWb.xlsx.readFile(filePath);
  const srcSheet = srcWb.getWorksheet("ETTON电商物流 下单模板");
  if (!srcSheet) throw new Error("思锐发票中找不到「ETTON电商物流 下单模板」工作表");

  const top = parseTopInfo(srcSheet);
  const items = parseItems(srcSheet);
  if (items.length === 0) throw new Error("思锐发票中未找到有效箱单数据");
  top.currency = items[0]?.currency || "";

  // 按目的地国家切换模板变体(英美: 美国→美线模板)
  const variant = baseSpec.countryVariants?.find((v) => v.match(top.recipientCountry));
  const spec: SupplierSpec = variant
    ? {
        ...baseSpec,
        templateFile: variant.templateFile,
        maxCol: variant.maxCol ?? baseSpec.maxCol,
        fillHeader: variant.fillHeader ?? baseSpec.fillHeader,
      }
    : baseSpec;

  // 提取源图片(仅 DISPIMG 供应商需要)
  if (spec.imageCol > 0) {
    if (supplier === "航乐") {
      // 航乐源为浮动图片(drawing1.xml)
      const images = await extractSourceImages(filePath, items.length);
      items.forEach((it, i) => {
        it.image = images[i];
      });
    } else {
      // 华威尔/美琦源为 WPS DISPIMG(cellimages.xml)
      const dispimgMap = await extractDispimgImages(filePath);
      items.forEach((it) => {
        if (it.dispimgId && dispimgMap.has(it.dispimgId)) {
          it.image = dispimgMap.get(it.dispimgId)!;
        }
      });
    }
  }

  // ── 2. 加载供应商模板 ──
  const tplPath = path.join(process.cwd(), "public", "templates", "sirui-suppliers", spec.templateFile);
  if (!fs.existsSync(tplPath)) throw new Error(`供应商模板不存在: ${tplPath}`);
  const outWb = new ExcelJS.Workbook();
  await outWb.xlsx.readFile(tplPath);
  const outSheet = spec.sheetName
    ? outWb.getWorksheet(spec.sheetName)
    : outWb.worksheets[0];
  if (!outSheet) throw new Error(`模板中找不到「${spec.sheetName || "(首个)"}」工作表`);

  // ── 3. 填充订单头 ──
  spec.fillHeader(outSheet, top, { customerOrderNo: orderNo, channelName, customsOverride, station }, items);

  // ── 4. 展开明细行(explode/排序号) ──
  const ctx: MapCtx = { boxSeq: 0, curDim: { l: 0, w: 0, h: 0 } };
  const outputRows: { item: SiruiItem; values: CellValue[] }[] = [];
  for (const item of items) {
    for (const values of spec.mapItem(item, top, ctx)) {
      outputRows.push({ item, values });
    }
  }

  const dataStart = spec.headerStartRow;
  // 模板数据行有显式行高则沿用, 否则用 spec.rowHeight; 两者皆无则保持自动行高(不写死 15)
  const dataRowHeight = outSheet.getRow(dataStart).height ?? spec.rowHeight;

  let templateDataRows: number;
  if (spec.dataEndRow) {
    templateDataRows = spec.dataEndRow - dataStart + 1;
  } else {
    const templateLast = findTemplateDataLastRow(outSheet, dataStart, spec.maxCol);
    templateDataRows = templateLast >= dataStart ? templateLast - dataStart + 1 : 0;
  }

  // 清空模板数据区值(保留格式/行高)
  const clearEnd = dataStart + templateDataRows - 1;
  for (let r = dataStart; r <= clearEnd; r++) {
    const row = outSheet.getRow(r);
    for (let c = 1; c <= spec.maxCol; c++) row.getCell(c).value = null;
  }

  // 数据超出模板预留行 → 扩行(航乐: 在最后数据行后插行, 合计行随之顺延)
  if (outputRows.length > templateDataRows) {
    if (spec.dataEndRow) {
      outSheet.duplicateRow(spec.dataEndRow, outputRows.length - templateDataRows, true);
    } else {
      outSheet.duplicateRow(dataStart, outputRows.length - templateDataRows, true);
    }
  }

  // 每输出行分配唯一图片 ID(DISPIMG 一行一 ID)
  const outputImages: (Buffer | null)[] = outputRows.map(() => null);
  const outputImageIds: (string | null)[] = outputRows.map(() => null);

  // 写入数据
  for (let i = 0; i < outputRows.length; i++) {
    const { item, values } = outputRows[i];
    const row = outSheet.getRow(dataStart + i);
    if (dataRowHeight != null) row.height = dataRowHeight;

    for (let c = 1; c <= spec.maxCol; c++) {
      const v = values[c - 1];
      if (v !== null && v !== undefined && v !== "") row.getCell(c).value = v;
    }

    if (spec.imageCol > 0 && item.image) {
      outputImages[i] = item.image;
      outputImageIds[i] = newImageId();
      row.getCell(spec.imageCol).value = {
        formula: `_xlfn.DISPIMG("${outputImageIds[i]}",1)`,
        result: `=DISPIMG("${outputImageIds[i]}",1)`,
      };
    }

    // 模板数据区本身无格式时, 按供应商逐列补全样式(华威尔)
    spec.styleDataRow?.(row, spec.maxCol);
  }

  // 更新合计行 SUM 公式(航乐插行后)
  if (spec.sumCols && spec.dataEndRow) {
    const sumRow = dataStart + outputRows.length; // 合计行 = 数据区末尾下一行
    const sumEnd = dataStart + outputRows.length - 1;
    for (const col of spec.sumCols) {
      outSheet.getCell(`${col}${sumRow}`).value = {
        formula: `SUM(${col}${dataStart}:${col}${sumEnd})`,
      };
    }
  }

  // ── 5. 写 buffer ──
  let buffer: Buffer = Buffer.from((await outWb.xlsx.writeBuffer()) as unknown as Uint8Array);

  // 修复 ExcelJS 丢失模板默认字体(宋体 charset=134 → Calibri)
  buffer = await applyTemplateDefaultFont(buffer, tplPath);

  // ── 6. DISPIMG 图片部件注入 ──
  if (spec.imageCol > 0) {
    buffer = await injectDispimg(buffer, outputImages, outputImageIds);
  }

  const baseName = orderNo || path.basename(sourceFileName, path.extname(sourceFileName));
  const fileName = `${sanitizeFileName(baseName)}.xlsx`;

  const warnings: string[] = [];
  if (supplier === "华威尔" && countryCode(top.recipientCountry) === "US") {
    warnings.push("目的地默认为「美东」，请核对收货州属东/中/西岸后按需修改");
  }

  return {
    sourceFile: sourceFileName,
    fileName,
    supplier,
    customerOrderNo: orderNo,
    totalBoxes: top.totalBoxes,
    dataRows: outputRows.length,
    warnings,
    buffer,
  };
}
