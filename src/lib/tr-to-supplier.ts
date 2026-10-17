/**
 * TR 客户发票 → 供应商下单模板 转换
 *
 * 数据流：
 *   TR客户发票(Page1)
 *     ├─ 订单头(R2-21 两列: A/B 左区 + E/F 右区): 带电/带磁/报关方式/清关方式/VAT/EORI/总箱数/国家/仓库代码
 *     └─ 箱单明细(R24 起, A-Y): 货箱编号/追踪编码/品名/材质/用途/箱数/总数量/单价/币种/海关编码/品牌/型号/净重/毛重/长宽高/链接/ASIN
 *       ↓
 *   供应商模板(public/templates/tr-suppliers/): 凯鑫 / 英美 / 星链 / 航乐
 *     ├─ 订单头字段(渠道/仓库代码/客户订单号/带电带磁/报关/VAT/箱数...)
 *     └─ 数据区(逐列映射 + 特殊规则)
 *
 * 关键规则(以答案文件反推为准)：
 *   1. 货箱编号 = 源货箱编号 + "U000001-" + 箱数(凯鑫/英美/航乐)；星链用累计区间 "起-止"
 *   2. 单箱数量 = 总数量 / 箱数
 *   3. 凯鑫：货箱重量 = 源毛重 - 0.5；长/宽/高 = 源值 - 1(凯鑫特有)
 *   4. 带电/带磁/币种 原值透传
 *   5. 报关方式：有单证 → 报关退税(凯鑫/英美/航乐) / 合并报关(星链)
 *   6. 渠道名、客户订单号 由用户手动输入
 *   7. 图片：凯鑫/航乐用 WPS DISPIMG 嵌入(zip 后处理注入)；英美/星链图片留空
 *   8. 收件地址字段为模板内 VLOOKUP 公式，写入仓库代码后自动带出，无需手动处理
 */

import ExcelJS from "exceljs";
import JSZip from "jszip";
import path from "path";
import fs from "fs";
import crypto from "crypto";

// ============================================================
// Types
// ============================================================

export type SupplierKey = "凯鑫" | "英美" | "星链" | "航乐";

/** 一行 TR 箱单明细 */
export interface TrItem {
  boxId: string;         // A 货箱编号
  refId: string;         // B 货件追踪编码(Reference ID)
  nameEn: string;        // C 英文品名
  nameCh: string;        // D 中文品名
  material: string;      // E 材质
  use: string;           // F 用途
  boxCount: number;      // G 箱数/件数
  totalQty: number;      // H 总申报数量
  unitPrice: number;     // I 申报单价
  totalPrice: number;    // J 申报总价
  currency: string;      // K 申报币种
  hsCode: string;        // L 海关编码
  brand: string;         // M 品牌
  model: string;         // N 型号
  brandType: string;     // O 品牌类型
  netWeight: number;     // P 单箱货物净重
  grossWeight: number;   // Q 货箱重量(毛重)
  lengthCm: number;      // R 长
  widthCm: number;       // S 宽
  heightCm: number;      // T 高
  link: string;          // U 销售链接
  asin: string;          // X ASIN
  image: Buffer | null;  // 源浮动图片(V列主图, 用于 DISPIMG)
}

/** TR 发票订单头 */
export interface TrTopInfo {
  hasBattery: string;    // F2 带电
  hasMagnet: string;     // F3 带磁
  bizType: string;       // B6 业务类型
  channel: string;       // B7 服务渠道(仅参考)
  customs: string;       // B8 报关方式
  clearance: string;     // B10 清关方式
  vatCompany: string;    // B11 VAT公司名
  vatAddress: string;    // B12 VAT注册地址
  vatNo: string;         // B13 VAT号
  eoriNo: string;        // B14 EORI号
  totalBoxes: number;    // B17 总箱数
  totalWeight: number;   // B18 预计总重量
  totalVolume: number;   // B19 预计总体积
  country: string;       // F8 收件人国家
  warehouseCode: string; // F9 亚马逊仓库代码
  warehouseAddress: string; // F10 亚马逊仓库地址(航乐 CITY 用)
  zipCode: string;          // F11 邮编(航乐 ZIP 用)
  currency: string;      // 申报币种(取首个明细 K 列, 整票统一)
}

/** 顶层转换结果 */
export interface TrConvertResult {
  sourceFile: string;
  fileName: string;        // 客户订单号.xlsx
  supplier: SupplierKey;
  customerOrderNo: string;
  warehouseCode: string;
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

/** 四舍五入到 2 位小数，避免浮点噪声 */
function r2(n: number): number {
  return Math.round(n * 100) / 100;
}

function sanitizeFileName(name: string): string {
  const cleaned = name.replace(/[\\/:*?"<>|]/g, "_").trim();
  return cleaned || "未命名";
}

// ============================================================
// TR 源解析
// ============================================================

function parseTopInfo(ws: ExcelJS.Worksheet): TrTopInfo {
  const s = (ref: string) => cellText(ws.getCell(ref));
  return {
    hasBattery: s("F2"),
    hasMagnet: s("F3"),
    bizType: s("B6"),
    channel: s("B7"),
    customs: s("B8"),
    clearance: s("B10"),
    vatCompany: s("B11"),
    vatAddress: s("B12"),
    vatNo: s("B13"),
    eoriNo: s("B14"),
    totalBoxes: cellNum(ws.getCell("B17")),
    totalWeight: cellNum(ws.getCell("B18")),
    totalVolume: cellNum(ws.getCell("B19")),
    country: s("F8"),
    warehouseCode: s("F9"),
    warehouseAddress: s("F10"),
    zipCode: s("F11"),
    currency: "",
  };
}

function parseItems(ws: ExcelJS.Worksheet): TrItem[] {
  const items: TrItem[] = [];
  for (let r = 24; r <= ws.rowCount; r++) {
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
      asin: cellText(ws.getCell(r, 24)),
      image: null,
    });
  }
  return items;
}

/**
 * 从源文件浮动图片提取产品图(按数据行)。
 * 源 TR 发票图片锚定在 V 列(col 21, 主图)/W 列(col 22)，行号对应数据行(R24=锚点行23)。
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

    // rId -> media 路径(drawing 关系是 ../media/imageN.png)
    const relMap: Record<string, string> = {};
    for (const m of relsXml.matchAll(/<Relationship[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"/g)) {
      relMap[m[1]] = "xl/" + m[2].replace(/^\.\.\//, "");
    }

    // 每个 twoCellAnchor: from.row/col + blip r:embed
    const anchors: { row: number; col: number; rId: string }[] = [];
    for (const m of drawingXml.matchAll(/<xdr:twoCellAnchor>([\s\S]*?)<\/xdr:twoCellAnchor>/g)) {
      const block = m[1];
      const fromCol = /<xdr:from>[\s\S]*?<xdr:col>(\d+)<\/xdr:col>/.exec(block);
      const fromRow = /<xdr:from>[\s\S]*?<xdr:row>(\d+)<\/xdr:row>/.exec(block);
      const embed = /<a:blip[^>]*r:embed="([^"]+)"/.exec(block);
      if (fromCol && fromRow && embed) {
        anchors.push({
          col: parseInt(fromCol[1], 10),
          row: parseInt(fromRow[1], 10),
          rId: embed[1],
        });
      }
    }

    // 按数据行取 V 列(col 21)主图，退回 W 列(col 22)
    for (let i = 0; i < itemCount; i++) {
      const anchorRow = 23 + i; // 数据行 R24 对应 0 基行 23
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
    // 图片提取失败不致命(图片列留空)
  }
  return result;
}

// ============================================================
// 供应商配置
// ============================================================

type CellValue = string | number | boolean | null | ExcelJS.CellFormulaValue;

interface SupplierSpec {
  key: SupplierKey;
  templateFile: string;   // public/templates/tr-suppliers/ 下文件名
  sheetName: string;      // 数据所在 sheet
  headerStartRow: number; // 数据区起始行
  imageCol: number;       // DISPIMG 图片列(0 = 不用 DISPIMG)
  rowHeight?: number;     // 数据行高(图片供应商需要)
  maxCol: number;         // 数据区最大列(用于边框)
  /**
   * 数据区结束行(含)。设置后表示数据区之后仍有需保留的内容(合计行/下拉数据区),
   * 填充采用「覆盖式」: 不删除数据区行, 只覆盖前 N 行值, 保留其后内容与全部格式。
   * 不设置则数据区延续到 sheet 末尾, 采用「重建式」(删除示例行 + 复制首行格式重写)。
   */
  dataEndRow?: number;
  fillHeader: (
    ws: ExcelJS.Worksheet,
    top: TrTopInfo,
    opts: { customerOrderNo: string; channelName: string },
    items: TrItem[]
  ) => void;
  mapItem: (item: TrItem, top: TrTopInfo, cumulativeStart: number) => CellValue[];
}

function customsFor(supplier: SupplierKey, customs: string): string {
  if (customs === "有单证") {
    return supplier === "星链" ? "合并报关（两票及以上需要合并单独报关的）" : "报关退税";
  }
  if (customs === "无单证") return "买单报关";
  return customs;
}

/** 航乐「物品属性」: 按品名/用途关键词粗分类(命中不到留空, 人工补填) */
function categoryFor(item: TrItem): string {
  const hay = `${item.use} ${item.nameCh} ${item.nameEn} ${item.material}`;
  if (/照明|投影|夜灯|灯|LED|Lighting|Light/i.test(hay)) return "灯类";
  if (/玩具|积木|拼图|Toy/i.test(hay)) return "玩具类";
  if (/服装|衣服|T恤|鞋|Cloth|Shoes/i.test(hay)) return "服装类";
  if (/家具|家居|Home|Furniture/i.test(hay)) return "家居类";
  return "";
}

const SUPPLIERS: Record<SupplierKey, SupplierSpec> = {
  // ── 凯鑫 ──────────────────────────────────────────────
  凯鑫: {
    key: "凯鑫",
    templateFile: "凯鑫-模板.xlsx",
    sheetName: "发票",
    headerStartRow: 28,
    imageCol: 16, // P 列 产品图片
    rowHeight: 35,
    maxCol: 19,
    fillHeader(ws, top, opts) {
      const b = (r: number, v: string | number) => {
        if (v !== "" && v !== null && v !== undefined) ws.getCell(r, 2).value = v;
      };
      b(1, opts.channelName);                 // 服务(渠道)
      b(2, top.warehouseCode);                // 地址库编码(触发 VLOOKUP 自动带出地址)
      ws.getCell(14, 2).value = opts.customerOrderNo || null; // 客户订单号(留空则清空模板示例)
      b(15, top.hasBattery);                  // 带电
      b(16, top.hasMagnet);                   // 带磁
      b(17, customsFor("凯鑫", top.customs)); // 报关方式
      // 交税方式: 清关方式 自税→自税递延
      b(18, top.clearance === "自税" ? "自税递延" : top.clearance === "不包税" ? "不包税" : "");
      b(19, top.vatNo);                       // VAT号
      b(20, top.eoriNo);                      // EORI号
      b(26, top.totalBoxes);                  // 箱数
    },
    mapItem(item) {
      const perBox = item.boxCount ? r2(item.totalQty / item.boxCount) : 0;
      return [
        `${item.boxId}U000001-${item.boxCount}`, // A 货箱编号
        item.nameEn,                            // B 英文品名
        item.nameCh,                            // C 中文品名
        perBox,                                 // D 单箱数量
        item.unitPrice,                         // E 申报单价
        r2(item.grossWeight - 0.5),             // F 货箱重量 -0.5
        r2(item.lengthCm - 1),                  // G 长 -1
        r2(item.widthCm - 1),                   // H 宽 -1
        r2(item.heightCm - 1),                  // I 高 -1
        item.hsCode,                            // J 海关编码
        item.brand,                             // K 品牌
        item.material,                          // L 材质
        item.model,                             // M 型号
        item.use,                               // N 用途
        item.link,                              // O 销售链接
        null,                                   // P 产品图片(DISPIMG 后处理)
        item.model,                             // Q 产品SKU(=型号)
        item.refId,                             // R PO Number
        item.asin,                              // S 产品ASIN
      ];
    },
  },

  // ── 英美 ──────────────────────────────────────────────
  英美: {
    key: "英美",
    templateFile: "英美-模板.xlsx",
    sheetName: "模板",
    headerStartRow: 26,
    imageCol: 0,
    maxCol: 26,
    fillHeader(ws, top, opts) {
      const b = (r: number, v: string | number) => {
        if (v !== "" && v !== null && v !== undefined) ws.getCell(r, 2).value = v;
      };
      b(1, opts.channelName);                 // 服务
      b(2, top.warehouseCode);                // 地址库编码
      ws.getCell(14, 2).value = opts.customerOrderNo || null; // 客户订单号(留空则清空模板示例)
      b(15, top.hasBattery);                  // 带电
      b(16, top.hasMagnet);                   // 带磁
      b(17, customsFor("英美", top.customs)); // 报关方式
      b(19, top.vatNo);                       // VAT号
      b(20, top.eoriNo);                      // EORI
      b(21, top.currency);                    // 申报币种(透传)
      b(24, top.totalBoxes);                  // 箱数
    },
    mapItem(item) {
      const perBox = item.boxCount ? r2(item.totalQty / item.boxCount) : 0;
      return [
        `${item.boxId}U000001-${item.boxCount}`, // A 货箱编号
        item.refId,                             // B PO Number
        item.grossWeight,                       // C 客户货箱重量
        item.lengthCm,                          // D 客户货箱长度
        item.widthCm,                           // E 客户货箱宽度
        item.heightCm,                          // F 客户货箱高度
        item.model,                             // G 产品SKU(=型号)
        item.nameCh,                            // H 中文品名
        item.nameEn,                            // I 英文品名
        item.unitPrice,                         // J 申报单价
        perBox,                                 // K 申报数量
        item.material,                          // L 材质
        item.hsCode,                            // M 海关编码
        item.use,                               // N 用途
        item.brand,                             // O 品牌
        item.model,                             // P 型号
        item.link,                              // Q 销售链接
        null,                                   // R 产品销售价格
        null,                                   // S 产品图片链接
        null,                                   // T 产品重量
        item.asin,                              // U 产品ASIN
      ];
    },
  },

  // ── 星链 ──────────────────────────────────────────────
  星链: {
    key: "星链",
    templateFile: "星链-模板.xlsx",
    sheetName: "仓库1",
    headerStartRow: 18,
    imageCol: 0,
    rowHeight: 31.1,
    maxCol: 22,
    fillHeader(ws, top, opts) {
      ws.getCell(3, 2).value = opts.customerOrderNo || null; // B3 客户参考号2(留空则清空模板示例)
      ws.getCell(6, 2).value = opts.channelName;       // B6 收货渠道
      ws.getCell(7, 2).value = top.country;            // B7 目的地(用于 INDIRECT 选地址表)
      ws.getCell(8, 2).value = top.totalBoxes;         // B8 总件数
      ws.getCell(9, 2).value = customsFor("星链", top.customs); // B9 报关类型
      ws.getCell(10, 2).value = "否";                   // B10 是否投保
      ws.getCell(11, 2).value = "不投保";               // B11 投保类型
      ws.getCell(12, 6).value = top.warehouseCode;      // F12 FBA仓库码
    },
    mapItem(item, top, cumulativeStart) {
      const perBox = item.boxCount ? r2(item.totalQty / item.boxCount) : 0;
      const start = cumulativeStart + 1;
      const end = cumulativeStart + item.boxCount;
      return [
        `${start}-${end}`, // A 箱序号(累计区间)
        item.boxId,        // B FBA ID
        item.refId,        // C POAnumber
        item.nameCh,       // D 中文名
        item.nameEn,       // E 英文名
        item.netWeight,    // F (单个产品)净重
        perBox,            // G 单箱数量
        item.unitPrice,    // H 申报单价
        item.currency,     // I 申报币别
        item.brand,        // J 品牌
        item.model,        // K 产品型号
        item.material,     // L 材质
        item.use,          // M 用途
        item.hsCode,       // N 海关编码
        item.grossWeight,  // O 单件实重
        item.lengthCm,     // P 长
        item.widthCm,      // Q 宽
        item.heightCm,     // R 高
        null,              // S 产品图片
        top.hasBattery,    // T 是否带电
        top.hasMagnet,     // U 是否带磁
        item.link,         // V 销售链接
      ];
    },
  },

  // ── 航乐 ──────────────────────────────────────────────
  航乐: {
    key: "航乐",
    templateFile: "航乐-模板.xlsx",
    sheetName: "Packing list装箱单发票",
    headerStartRow: 12,
    imageCol: 23, // W 列 产品图片
    rowHeight: 35,
    maxCol: 24,
    dataEndRow: 23, // 数据区 R12~R23(R24 起为合计行/渠道下拉区, 须保留), 覆盖式填充
    fillHeader(ws, top, opts, items) {
      const boxIds = items.map((i) => i.boxId).join(",");
      const refIds = items.map((i) => i.refId).join(",");

      ws.getCell(3, 2).value = top.warehouseCode;                 // B3 收件地址 = 仓库代码
      ws.getCell(3, 7).value = top.country;                       // G3 国家
      ws.getCell(3, 8).value = `SHIPMENT ID（FBA货物编码）：\n${boxIds}`; // H3
      ws.getCell(3, 11).value = `AMAZON REFERENCE ID（亚马逊内部编码）：\n${refIds}`; // K3
      ws.getCell(4, 2).value = Number(top.zipCode) || top.zipCode; // B4 ZIP
      ws.getCell(5, 2).value = top.warehouseAddress;              // B5 CITY = 仓库地址
      ws.getCell(5, 4).value = "亚马逊";                          // D5 ATTN 收件人
      ws.getCell(5, 9).value = customsFor("航乐", top.customs);   // I5 报关类型
      ws.getCell(5, 11).value = "广州白云仓";                     // K5 交货仓库
      ws.getCell(7, 2).value = top.vatCompany;                    // B7 公司名称
      ws.getCell(7, 9).value = opts.channelName;                  // I7 渠道
      ws.getCell(7, 12).value =                                   // L7 是否包税
        top.clearance === "自税" ? "不包税递延" : top.clearance === "不包税" ? "不包税" : "";
      ws.getCell(7, 13).value = categoryFor(items[0]); // M7 物品属性
      ws.getCell(8, 2).value = top.vatNo;                         // B8 VAT号
      ws.getCell(9, 2).value = top.eoriNo;                        // B9 EORI号
      ws.getCell(10, 2).value = top.vatAddress.replace(/\n/g, " "); // B10 VAT注册地址(换行转空格)
    },
    mapItem(item, top) {
      const perBox = item.boxCount ? item.totalQty / item.boxCount : 0;
      const dimVol = item.lengthCm * item.widthCm * item.heightCm; // 长宽高乘积(cm³)
      const hs = /^\d+$/.test(item.hsCode) ? Number(item.hsCode) : item.hsCode;
      return [
        `${item.boxId}U000001-${item.boxCount}`, // A 货箱编号(FBA箱号)
        item.nameCh,        // B 品名(中文)
        item.nameEn,        // C 品名(英文)
        item.material,      // D 材质(中文)
        item.material,      // E 材质(英文) 同源
        hs,                 // F 海关编码(十位数, 数值)
        item.model,         // G 型号
        item.brand,         // H 品牌
        item.use,           // I 产品用途(英文)
        perBox,             // J 单箱数量 = 总数量/箱数
        item.totalQty,      // K 总数量
        perBox ? item.netWeight / perBox : 0, // L 净重(单个产品)
        item.grossWeight,   // M 实重(毛重)
        item.unitPrice,     // N 单价(EUR)
        item.totalPrice,    // O 总价(EUR)
        item.lengthCm,      // P 长
        item.widthCm,       // Q 宽
        item.heightCm,      // R 高
        dimVol / 6000,      // S 材重
        dimVol / 1000000,   // T CBM 方数
        top.hasBattery,     // U 是否带电
        top.hasMagnet,      // V 是否带磁
        null,               // W 产品图片(DISPIMG 后处理)
        item.link,          // X 产品销售链接
      ];
    },
  },
};

// ============================================================
// 渠道下拉数据(从供应商模板读取, 单一数据源)
// ============================================================

interface ChannelSheetSpec {
  sheet: string;
  col: number;      // 1-based 列号
  startRow: number; // 数据起始行(1-based)
}

const CHANNEL_SHEETS: Record<SupplierKey, ChannelSheetSpec[]> = {
  凯鑫: [{ sheet: "服务名称", col: 2, startRow: 2 }],
  英美: [{ sheet: "渠道列表", col: 2, startRow: 2 }],
  星链: [
    { sheet: "US忽略", col: 1, startRow: 2 },
    { sheet: "AU忽略", col: 1, startRow: 2 },
  ],
  航乐: [
    { sheet: "Sheet2", col: 1, startRow: 1 },
    { sheet: "Sheet1", col: 1, startRow: 1 },
  ],
};

const channelsCache = new Map<SupplierKey, string[]>();

/** 读取某供应商的渠道下拉列表(去重, 保持模板顺序) */
export async function getSupplierChannels(supplier: SupplierKey): Promise<string[]> {
  const cached = channelsCache.get(supplier);
  if (cached) return cached;

  const spec = SUPPLIERS[supplier];
  const sheets = CHANNEL_SHEETS[supplier];
  if (!spec || !sheets) throw new Error(`未知供应商: ${supplier}`);

  const tplPath = path.join(process.cwd(), "public", "templates", "tr-suppliers", spec.templateFile);
  if (!fs.existsSync(tplPath)) throw new Error(`供应商模板不存在: ${tplPath}`);

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(tplPath);

  const seen = new Set<string>();
  const channels: string[] = [];
  for (const cs of sheets) {
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

  channelsCache.set(supplier, channels);
  return channels;
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

/** 生成 cellimages.xml(每张图一个 etc:cellImage)。y = 5499735 + itemIndex * 956945 */
function buildCellImagesXml(entries: { id: string; rId: string; itemIndex: number }[]): string {
  const pics = entries
    .map(
      (e, i) =>
        `<etc:cellImage><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${i + 1}" name="${e.id}" descr="Picture"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="${e.rId}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="26471880" y="${5499735 + e.itemIndex * 956945}"/><a:ext cx="744220" cy="956945"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic></etc:cellImage>`
    )
    .join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<etc:cellImages xmlns:xdr="${DISPIMG_NS.xdr}" xmlns:r="${DISPIMG_NS.r}" xmlns:a="${DISPIMG_NS.a}" xmlns:etc="${DISPIMG_NS.etc}">${pics}</etc:cellImages>`;
}

/** 生成 cellimages.xml.rels */
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
 * 向已生成的 xlsx buffer 注入 DISPIMG 图片部件(media + cellimages.xml + 关系)。
 * 单元格公式已由 exceljs 在写入时设置, 此处只补图片实体与关系。
 */
async function injectDispimg(
  buffer: Buffer,
  images: (Buffer | null)[],
  imageIds: (string | null)[]
): Promise<Buffer> {
  const idxs: number[] = [];
  images.forEach((img, i) => {
    if (img && imageIds[i]) idxs.push(i);
  });
  if (idxs.length === 0) return buffer;

  const zip = await JSZip.loadAsync(buffer);

  const entries: { id: string; rId: string; itemIndex: number }[] = [];
  const relEntries: { rId: string; target: string }[] = [];
  idxs.forEach((itemIdx, n) => {
    const mediaName = `image${n + 1}.png`;
    zip.file(`xl/media/${mediaName}`, images[itemIdx]!);
    const rId = `rId${n + 1}`;
    entries.push({ id: imageIds[itemIdx]!, rId, itemIndex: itemIdx });
    relEntries.push({ rId, target: `media/${mediaName}` });
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

// ============================================================
// 数据区覆盖式填充辅助(保留模板数据区全部格式: 字体/颜色/背景/边框/对齐/数字格式/行高)
// ============================================================

/**
 * 从数据区起始行向下扫描, 返回最后一行「有值或有边框」的行号(1-based)。
 * 用于确定模板数据区预留了多少格式行; 数据超出时在末尾复制首行格式扩行。
 * 返回 dataStart-1 表示数据区为空(异常模板)。
 */
function findTemplateDataLastRow(
  ws: ExcelJS.Worksheet,
  dataStart: number,
  maxCol: number
): number {
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
    else if (r > dataStart) break; // 数据区之后连续空行, 停止扫描
  }
  return last;
}

// ============================================================
// 主入口
// ============================================================

export async function convertTrToSupplier(
  filePath: string,
  sourceFileName: string,
  supplier: SupplierKey,
  customerOrderNo: string,
  channelName: string
): Promise<TrConvertResult> {
  const spec = SUPPLIERS[supplier];
  if (!spec) throw new Error(`未知供应商: ${supplier}`);

  const orderNo = (customerOrderNo || "").trim();

  // ── 1. 解析 TR 源 ──
  const srcWb = new ExcelJS.Workbook();
  await srcWb.xlsx.readFile(filePath);
  const srcSheet = srcWb.getWorksheet("Page1");
  if (!srcSheet) throw new Error("TR 发票中找不到「Page1」工作表");

  const top = parseTopInfo(srcSheet);
  const items = parseItems(srcSheet);
  if (items.length === 0) throw new Error("TR 发票中未找到有效箱单数据");
  top.currency = items[0]?.currency || "";

  // 提取源图片(仅 DISPIMG 供应商需要) + 预生成图片 ID
  const imageIds: (string | null)[] = items.map(() => null);
  if (spec.imageCol > 0) {
    const images = await extractSourceImages(filePath, items.length);
    items.forEach((it, i) => {
      it.image = images[i];
      if (it.image) imageIds[i] = newImageId();
    });
  }

  // ── 2. 加载供应商模板 ──
  const tplPath = path.join(process.cwd(), "public", "templates", "tr-suppliers", spec.templateFile);
  if (!fs.existsSync(tplPath)) throw new Error(`供应商模板不存在: ${tplPath}`);
  const outWb = new ExcelJS.Workbook();
  await outWb.xlsx.readFile(tplPath);
  const outSheet = outWb.getWorksheet(spec.sheetName);
  if (!outSheet) throw new Error(`模板中找不到「${spec.sheetName}」工作表`);

  // ── 3. 填充订单头 ──
  spec.fillHeader(outSheet, top, { customerOrderNo: orderNo, channelName }, items);

  // ── 4. 填充数据区(覆盖式: 保留模板数据区全部格式, 仅覆盖值) ──
  const dataStart = spec.headerStartRow;
  const dataRowHeight = outSheet.getRow(dataStart).height ?? spec.rowHeight ?? 15;

  // 模板数据区预留的格式行数
  let templateDataRows: number;
  if (spec.dataEndRow) {
    // 航乐: 数据区固定行数, 之后有保留内容(合计行/下拉数据区), 超出报错
    templateDataRows = spec.dataEndRow - dataStart + 1;
    if (items.length > templateDataRows) {
      throw new Error(
        `「${spec.key}」数据行数(${items.length})超出模板预留行数(${templateDataRows})，请拆分发票`
      );
    }
  } else {
    // 凯鑫/英美/星链: 扫描模板数据区实际预留行数(示例行 + 空格式行)
    const templateLast = findTemplateDataLastRow(outSheet, dataStart, spec.maxCol);
    templateDataRows = templateLast >= dataStart ? templateLast - dataStart + 1 : 0;
  }

  // 清空模板数据区值(保留格式/行高)
  const clearEnd = dataStart + templateDataRows - 1;
  for (let r = dataStart; r <= clearEnd; r++) {
    const row = outSheet.getRow(r);
    for (let c = 1; c <= spec.maxCol; c++) row.getCell(c).value = null;
  }

  // 数据超出模板预留行 → 复制数据区首行格式扩行(保留边框/底色/字体/行高)
  if (!spec.dataEndRow && items.length > templateDataRows) {
    outSheet.duplicateRow(dataStart, items.length - templateDataRows, true);
  }

  let cumulative = 0;
  for (let i = 0; i < items.length; i++) {
    const item = items[i];
    const row = outSheet.getRow(dataStart + i);
    row.height = dataRowHeight;

    const values = spec.mapItem(item, top, cumulative);
    for (let c = 1; c <= spec.maxCol; c++) {
      const v = values[c - 1];
      // 覆盖式: 模板单元格已有正确样式, 仅写值(空值跳过, 保留模板原样式)
      if (v !== null && v !== undefined && v !== "") row.getCell(c).value = v;
    }

    // DISPIMG 公式(图片列)
    if (spec.imageCol > 0 && item.image && imageIds[i]) {
      const id = imageIds[i]!;
      row.getCell(spec.imageCol).value = {
        formula: `_xlfn.DISPIMG("${id}",1)`,
        result: `=DISPIMG("${id}",1)`,
      };
    }
    cumulative += item.boxCount;
  }

  // ── 5. 写 buffer ──
  let buffer: Buffer = Buffer.from(await outWb.xlsx.writeBuffer() as unknown as Uint8Array);

  // ── 6. DISPIMG 图片部件注入 ──
  if (spec.imageCol > 0) {
    buffer = await injectDispimg(buffer, items.map((it) => it.image), imageIds);
  }

  const baseName = orderNo || path.basename(sourceFileName, path.extname(sourceFileName));
  const fileName = `${sanitizeFileName(baseName)}.xlsx`;
  return {
    sourceFile: sourceFileName,
    fileName,
    supplier,
    customerOrderNo: orderNo,
    warehouseCode: top.warehouseCode,
    totalBoxes: top.totalBoxes,
    dataRows: items.length,
    warnings: [],
    buffer,
  };
}

// ============================================================
// 易通系统上传发票: 源发票第 8 行插入空行(第 8 行起全部下移一行), 保留图片/合并单元格/格式
// ============================================================

const ETON_INSERT_ROW = 8; // 1-based 行号: 在第 8 行插入空行

function resolveWorksheetTarget(wbXml: string, relsXml: string, sheetName: string): string | null {
  for (const m of wbXml.matchAll(/<sheet\b([^>]*)>/g)) {
    const attrs = m[1];
    const name = /name="([^"]*)"/.exec(attrs);
    if (!name || name[1] !== sheetName) continue;
    const rid = /r:id="([^"]+)"/.exec(attrs);
    if (!rid) return null;
    const rel = new RegExp(`<Relationship\\b[^>]*Id="${rid[1]}"[^>]*Target="([^"]+)"`).exec(relsXml);
    if (!rel) return null;
    const target = rel[1].replace(/^\/+/, "").replace(/^(\.\.\/)+/, "");
    return "xl/" + target;
  }
  return null;
}

/** sheet XML: 将 fromRow(1-based) 起的行/单元格/合并单元格/尺寸范围全部下移一行 */
function shiftRowRefsInSheetXml(xml: string, fromRow: number): string {
  // 1. dimension 范围
  xml = xml.replace(/<dimension ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"/g, (m, c1, r1s, c2, r2s) => {
    const r1 = parseInt(r1s, 10);
    const r2 = parseInt(r2s, 10);
    return `<dimension ref="${c1}${r1 >= fromRow ? r1 + 1 : r1}:${c2}${r2 >= fromRow ? r2 + 1 : r2}"`;
  });
  // 2. 行元素 <row r="N">
  xml = xml.replace(/<row r="(\d+)"/g, (m, n) => {
    const r = parseInt(n, 10);
    return r >= fromRow ? `<row r="${r + 1}"` : m;
  });
  // 3. 单元格 <c r="ColN">
  xml = xml.replace(/<c r="([A-Z]+)(\d+)"/g, (m, col, n) => {
    const r = parseInt(n, 10);
    return r >= fromRow ? `<c r="${col}${r + 1}"` : m;
  });
  // 4. 合并单元格 <mergeCell ref="ColN:ColN">
  xml = xml.replace(/<mergeCell ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"/g, (m, c1, r1s, c2, r2s) => {
    const r1 = parseInt(r1s, 10);
    const r2 = parseInt(r2s, 10);
    return `<mergeCell ref="${c1}${r1 >= fromRow ? r1 + 1 : r1}:${c2}${r2 >= fromRow ? r2 + 1 : r2}"`;
  });
  return xml;
}

/** drawing XML: 浮动图片锚点(0-based row)下移一行 */
function shiftRowsInDrawingXml(xml: string, fromRow: number): string {
  const zeroBasedFrom = fromRow - 1;
  return xml.replace(/<xdr:row>(\d+)<\/xdr:row>/g, (m, n) => {
    const r = parseInt(n, 10);
    return r >= zeroBasedFrom ? `<xdr:row>${r + 1}</xdr:row>` : m;
  });
}

/** 生成「易通系统上传」发票 = 源发票第 8 行插入空行(保留原格式/图片/合并单元格) */
export async function buildEttonInvoice(filePath: string): Promise<Buffer> {
  const raw = fs.readFileSync(filePath);
  const zip = await JSZip.loadAsync(raw);

  const wbXml = (await zip.file("xl/workbook.xml")?.async("string")) || "";
  const relsXml = (await zip.file("xl/_rels/workbook.xml.rels")?.async("string")) || "";
  if (!wbXml || !relsXml) throw new Error("易通发票 xlsx 结构不完整");

  const sheetPath = resolveWorksheetTarget(wbXml, relsXml, "Page1");
  if (!sheetPath) throw new Error("易通发票中找不到「Page1」工作表");

  const sheetFile = zip.file(sheetPath);
  if (!sheetFile) throw new Error(`易通发票中找不到工作表文件: ${sheetPath}`);
  zip.file(sheetPath, shiftRowRefsInSheetXml(await sheetFile.async("string"), ETON_INSERT_ROW));

  // 浮动图片锚点同步下移
  const drawingFile = zip.file("xl/drawings/drawing1.xml");
  if (drawingFile) {
    zip.file("xl/drawings/drawing1.xml", shiftRowsInDrawingXml(await drawingFile.async("string"), ETON_INSERT_ROW));
  }

  const out = await zip.generateAsync({ type: "nodebuffer" });
  return Buffer.isBuffer(out) ? out : Buffer.from(out);
}
