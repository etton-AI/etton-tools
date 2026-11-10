/**
 * 皮皮熊账单拆分核心逻辑
 *
 * 从皮皮熊合并账单（4 个数据 sheet）读取未支付记录，按货代识别号拆分为四类账单：
 *   国内账单（深圳抬头 / RMB，拖车费+港杂费）
 *   报关费账单（深圳抬头 / RMB，客户自行付报关行）
 *   国外账单（香港抬头 / USD）
 *   INVOICE（香港抬头 / 客户发票）
 *
 * 数据 sheet：
 *   1.大货出货数据费用统计        → 海运费/拖车费/报关费/港杂费（RMB + USD 混合）
 *   2.港皮海外仓转仓费用明细      → USD（香港公司支付）
 *   3.大货以及其他货物快递退运费用 → USD（香港公司支付）
 *   4.空派快递费用明细           → USD（香港公司支付）
 */

import ExcelJS from "exceljs";
import JSZip from "jszip";
import { readFileSync, existsSync } from "fs";
import { join } from "path";

// ============================================================
// Types
// ============================================================

export interface BillRow {
  seq: string;
  forwarderId: string; // 完整多行
  fwdFirst: string; // 首行（货代识别号）
  customsNo: string;
  time: string;
  platform: string;
  country: string;
  warehouseCode: string;
  qty: string | number;
  truckFee: number;
  customsFee: number;
  portFee: number;
  oceanTaxUSD: number;
  totalRMB: number;
  note1: string;
  note2: string;
  paid: string;
  jobNo?: string; // INVOICE JOB NO（去重后缀）
}

export interface UsdRow {
  seq: string;
  forwarderId: string; // 首行（货代识别号）
  time: string;
  qty: string | number;
  fee: number;
  currency: string;
  detail: string;
  paid: string;
  itemLabel: string; // INVOICE 项目标签
  jobNo?: string;
}

export interface SplitGroup {
  forwarderId: string;
  customsNo: string;
  seq: string;
  bill: BillRow;
  hasDomestic: boolean;
  hasCustoms: boolean;
  hasInternational: boolean;
  hasInvoice: boolean;
}

export interface PipixiongSplitResult {
  sourceFile: string;
  groups: SplitGroup[];
  domesticBuffer: Buffer;
  customsBuffer: Buffer;
  internationalBuffer: Buffer;
  invoiceBuffer: Buffer;
}

export interface BrokerInput {
  name: string;
  bankInfo: string;
}

// ============================================================
// Style constants — 与模板完全一致
// ============================================================

// 颜色
const TEAL = "FF18605A"; // 深青色 — 表头背景
const WHITE = "FFFFFFFF"; // 白色字
const BLACK = "FF000000";

// 字体（charset 134 = GB2312 中文字符集，与模板一致）
const fontDengXian = (size: number, bold = false, color?: string) =>
  ({ name: "等线", size, bold, charset: 134, color: color ? { argb: color } : undefined } as Partial<ExcelJS.Font>);

const fontYaHei = (size: number, bold = false, color?: string) =>
  ({ name: "微软雅黑", size, bold, charset: 134, color: color ? { argb: color } : undefined } as Partial<ExcelJS.Font>);

const fontArial = (size: number, bold = false, color?: string) =>
  ({ name: "Arial", size, bold, charset: 134, color: color ? { argb: color } : undefined } as Partial<ExcelJS.Font>);

const fontSong = (size: number, bold = false) =>
  ({ name: "宋体", size, bold, charset: 134 } as Partial<ExcelJS.Font>);

const fontTimesNR = (size: number, bold = false) =>
  ({ name: "Times New Roman", size, bold, charset: 134 } as Partial<ExcelJS.Font>);

// 页面设置：与模板一致（Letter 纸、fitToPage、fitToHeight=0；
// 显式置 undefined 以省略 ExcelJS 默认写入的 fitToWidth/scale/DPI）
function pageSetup(orientation: "landscape" | "portrait"): Partial<ExcelJS.PageSetup> {
  return {
    orientation,
    paperSize: 1 as unknown as ExcelJS.PageSetup["paperSize"], // 1 = Letter
    horizontalDpi: undefined,
    verticalDpi: undefined,
    scale: undefined,
    fitToPage: true,
    fitToWidth: undefined,
    fitToHeight: 0,
  };
}

// 边框
const thinBorder: Partial<ExcelJS.Borders> = {
  top: { style: "thin" },
  bottom: { style: "thin" },
  left: { style: "thin" },
  right: { style: "thin" },
};

// 填充
const headerFill: ExcelJS.Fill = {
  type: "pattern", pattern: "solid",
  fgColor: { argb: TEAL },
};

// 对齐
const centerWrap: Partial<ExcelJS.Alignment> = {
  horizontal: "center", vertical: "middle", wrapText: true,
};
const centerNoWrap: Partial<ExcelJS.Alignment> = {
  horizontal: "center", vertical: "middle",
};
const leftWrap: Partial<ExcelJS.Alignment> = {
  horizontal: "left", vertical: "middle", wrapText: true,
};
const leftCenter: Partial<ExcelJS.Alignment> = {
  horizontal: "left", vertical: "middle",
};

// 数字格式
const NUM_FMT = "#,##0.00";

// ============================================================
// Helpers
// ============================================================

function cellStr(cell: ExcelJS.Cell): string {
  if (isCoveredMergedCell(cell)) return "";
  const v = cell.value;
  if (v === null || v === undefined) return "";
  if (typeof v === "string") return v.trim();
  if (typeof v === "number") {
    if (v > 1e15) return String(Math.floor(v));
    return String(v);
  }
  if (typeof v === "object") {
    const obj = v as unknown as Record<string, unknown>;
    if ("richText" in obj && Array.isArray(obj.richText)) {
      return (obj.richText as Array<{ text: string }>).map((r) => r.text).join("").trim();
    }
    if ("result" in obj) return String((obj.result as string | number | undefined) ?? "").trim();
  }
  return String(v ?? "").trim();
}

function cellNum(cell: ExcelJS.Cell): number {
  if (isCoveredMergedCell(cell)) return 0;
  const v = cell.value;
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return v;
  if (typeof v === "string") {
    const cleaned = v.replace(/[￥¥$€£,，\s]/g, "");
    const n = parseFloat(cleaned);
    return isNaN(n) ? 0 : n;
  }
  if (typeof v === "object" && "result" in v && v.result !== undefined) {
    if (typeof v.result === "number") return v.result;
    const n = parseFloat(String(v.result).replace(/[￥¥$€£,，\s]/g, ""));
    return isNaN(n) ? 0 : n;
  }
  return 0;
}

function firstLine(s: string): string {
  const lines = s.split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
  return lines[0] ?? "";
}

/** 归一化「年.月」时间："2026.08" → "2026.8"，"2026.8" 保持不变 */
function fmtTime(v: string): string {
  const s = (v || "").trim();
  const m = s.match(/^(\d{4})\.(\d{1,2})$/);
  if (!m) return s;
  return `${m[1]}.${Number(m[2])}`;
}

/** 判断是否为「被合并覆盖」的单元格（非合并主单元格），其值应视为空 */
function isCoveredMergedCell(cell: ExcelJS.Cell): boolean {
  if (!cell.isMerged) return false;
  const master = cell.master as ExcelJS.Cell | undefined;
  return !!master && master.address !== cell.address;
}

function isPaid(paid: string): boolean {
  const p = (paid || "").trim().toLowerCase();
  return p === "y" || p === "yes" || p === "是" || p === "已支付";
}

// ============================================================
// INVOICE 项目标签
// ============================================================

const INVOICE_ITEM_OCEAN = "OCEAN FREIGHT & DESTINATION PORT TAXES";
const INVOICE_ITEM_COMPENSATION = "COMPENSATION COSTS";
const INVOICE_ITEM_WAREHOUSE_TRANSFER = "WAREHOUSE TRANSFER COST";
const INVOICE_ITEM_EXPRESS_RETURN = "EXPRESS RETURN COST";
const INVOICE_ITEM_AIR_EXPRESS = "INTERNATIONAL SHIPPING COST";

// ============================================================
// Parsing
// ============================================================

function parseDataSheet(ws: ExcelJS.Worksheet): BillRow[] {
  const rows: BillRow[] = [];
  for (let r = 4; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const seq = cellStr(row.getCell(1));
    if (!seq) continue;
    const forwarderId = cellStr(row.getCell(10));
    rows.push({
      seq,
      forwarderId,
      fwdFirst: firstLine(forwarderId),
      customsNo: cellStr(row.getCell(11)),
      time: cellStr(row.getCell(2)),
      platform: cellStr(row.getCell(3)),
      country: cellStr(row.getCell(4)),
      warehouseCode: cellStr(row.getCell(5)),
      qty: cellStr(row.getCell(9)),
      truckFee: cellNum(row.getCell(13)),
      customsFee: cellNum(row.getCell(14)),
      portFee: cellNum(row.getCell(15)),
      oceanTaxUSD: cellNum(row.getCell(16)),
      totalRMB: cellNum(row.getCell(12)),
      note1: cellStr(row.getCell(17)),
      note2: cellStr(row.getCell(18)),
      paid: cellStr(row.getCell(19)),
    });
  }
  return rows;
}

interface UsdSheetOpts {
  qtyCol: number;
  fwdCol: number;
  feeCol: number;
  currencyCol: number | null; // null → 固定 USD
  detailCol: number;
  paidCol: number;
  itemLabel: string;
}

function parseUsdSheet(ws: ExcelJS.Worksheet, opts: UsdSheetOpts): UsdRow[] {
  const rows: UsdRow[] = [];
  for (let r = 4; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const seq = cellStr(row.getCell(1));
    if (!seq) continue;
    const fwdFull = cellStr(row.getCell(opts.fwdCol));
    const currency = opts.currencyCol
      ? cellStr(row.getCell(opts.currencyCol)) || "USD"
      : "USD";
    rows.push({
      seq,
      forwarderId: firstLine(fwdFull),
      time: fmtTime(cellStr(row.getCell(2))),
      qty: cellNum(row.getCell(opts.qtyCol)),
      fee: cellNum(row.getCell(opts.feeCol)),
      currency,
      detail: cellStr(row.getCell(opts.detailCol)),
      paid: cellStr(row.getCell(opts.paidCol)),
      itemLabel: opts.itemLabel,
    });
  }
  return rows;
}

// ============================================================
// 国内账单 Builder（10 列：拖车费 + 港杂费，无报关费）
// ============================================================

function buildDomesticSheet(wb: ExcelJS.Workbook, name: string, bill: BillRow, logoId: number): void {
  const ws = wb.addWorksheet(name, {
    pageSetup: pageSetup("landscape"),
  });

  const colWidths = [7, 9.5278, 4.6111, 7, 9.6759, 10.2685, 26.6759, 9.5, 8.5, 13.6389];
  colWidths.forEach((w, i) => { ws.getColumn(i + 1).width = w; });

  // ── LOGO ──
  ws.addImage(logoId, {
    tl: { col: 0.1, row: 0.1 },
    ext: { width: 135, height: 105 },
    editAs: "oneCell",
  });

  // ── R1 (h=84): 公司抬头 ──
  ws.getRow(1).height = 84;
  ws.mergeCells("A1:C1");
  ws.mergeCells("D1:J1");
  const c1 = ws.getCell("D1");
  c1.value = "深圳市易通科技供应链有限公司\nSHENZHEN ETTON TECHNOLOGY SUPPLY CHAIN LTD\n地址：深圳市龙华区民治街道民强社区宝山时代大厦26楼2615";
  c1.font = fontYaHei(16, true);
  c1.alignment = { horizontal: "left", vertical: "middle", wrapText: true };

  // ── R2 (h=27): 副标题 ──
  ws.getRow(2).height = 27;
  ws.mergeCells("A2:J2");
  const c2 = ws.getCell("A2");
  c2.value = "深圳芝麻果实科技有限公司-出港皮历史出货数据";
  c2.font = fontDengXian(12, true, BLACK);
  c2.alignment = centerWrap;
  c2.border = thinBorder;

  // ── R3 (h=27): 基础信息 / 国内税票 ──
  ws.getRow(3).height = 27;
  ws.mergeCells("A3:F3");
  ws.mergeCells("H3:I3");
  const c3a = ws.getCell("A3");
  c3a.value = "基础信息";
  c3a.font = fontDengXian(10.5, true, WHITE);
  c3a.alignment = centerNoWrap;
  c3a.border = thinBorder;
  c3a.fill = headerFill;
  const c3b = ws.getCell("H3");
  c3b.value = "国内税票";
  c3b.font = fontDengXian(10.5, true, WHITE);
  c3b.alignment = centerNoWrap;
  c3b.border = thinBorder;
  c3b.fill = headerFill;

  // ── R4 (h=36): 列标题 ──
  ws.getRow(4).height = 36;
  const colHeaders = ["序列", "时间", "平台", "国家", "入库地址\n（代码）", "发货数量", "货代对应识别号", "拖车费RMB", "港杂费RMB", "总费用RMB"];
  colHeaders.forEach((h, i) => {
    const cell = ws.getCell(4, i + 1);
    cell.value = h;
    cell.font = fontDengXian(10.5, true, WHITE);
    cell.alignment = centerWrap;
    cell.border = thinBorder;
    cell.fill = headerFill;
  });

  // ── R5 (h=67): 数据行 ──
  ws.getRow(5).height = 67;
  const truck = +bill.truckFee.toFixed(2);
  const port = +bill.portFee.toFixed(2);
  const totalFee = +(truck + port).toFixed(2);

  const dataVals: Array<string | number | null> = [
    1, bill.time, bill.platform || null, bill.country,
    bill.warehouseCode || null, bill.qty, bill.fwdFirst,
    truck, port, null, // J5 = formula
  ];
  dataVals.forEach((v, i) => {
    if (v === null) return;
    const cell = ws.getCell(5, i + 1);
    cell.value = v;
    cell.font = fontDengXian(10.5, false);
    cell.alignment = centerWrap;
    cell.border = thinBorder;
    if (i >= 7 && i <= 8) cell.numFmt = NUM_FMT;
  });
  const j5 = ws.getCell("J5");
  j5.value = { formula: "SUM(H5:I5)", result: totalFee };
  j5.font = fontDengXian(10.5, true);
  j5.alignment = centerNoWrap;
  j5.border = thinBorder;
  j5.numFmt = NUM_FMT;

  // ── R7 (h=40): 温馨提示 ──
  ws.getRow(7).height = 40;
  ws.mergeCells("A7:J7");
  const c7 = ws.getCell("A7");
  c7.value = "温馨提示:感谢贵司对我司业务支持，请收到此账单三个工作日内确认回复，否则我司默认贵司已确认金额，如有问题请及时联系我司客服。";
  c7.font = fontYaHei(12, true);
  c7.alignment = leftWrap;

  // ── R8-R11 (h=25): 深圳银行信息 ──
  const bankLines = [
    "开户人:深圳市易通科技供应链有限公司",
    "开户行:中国银行深圳彩虹支行",
    "人民币账号:7536 7452 1237",
    "美金账号:7744 7515 5905",
  ];
  bankLines.forEach((text, idx) => {
    const rr = 8 + idx;
    ws.getRow(rr).height = 25;
    ws.mergeCells(`A${rr}:J${rr}`);
    const cell = ws.getCell(`A${rr}`);
    cell.value = text;
    cell.font = fontYaHei(14, true);
    cell.alignment = leftWrap;
  });
}

// ============================================================
// 报关费账单 Builder（9 列，报关费 RMB，客户自行付报关行）
// ============================================================

function buildCustomsSheet(
  wb: ExcelJS.Workbook, name: string, bill: BillRow, broker: BrokerInput | undefined, logoId: number
): void {
  const ws = wb.addWorksheet(name, {
    pageSetup: pageSetup("landscape"),
  });

  const colWidths = [7, 9.5278, 4.6111, 7, 9.6759, 10.2685, 26.6759, 9.3889, 25.7778];
  colWidths.forEach((w, i) => { ws.getColumn(i + 1).width = w; });

  ws.addImage(logoId, {
    tl: { col: 0.1, row: 0.1 },
    ext: { width: 135, height: 105 },
    editAs: "oneCell",
  });

  // ── R1 (h=84): 公司抬头 ──
  ws.getRow(1).height = 84;
  ws.mergeCells("A1:C1");
  ws.mergeCells("D1:I1");
  const c1 = ws.getCell("D1");
  c1.value = "深圳市易通科技供应链有限公司\nSHENZHEN ETTON TECHNOLOGY SUPPLY CHAIN LTD\n地址：深圳市龙华区民治街道民强社区宝山时代大厦26楼2615";
  c1.font = fontYaHei(16, true);
  c1.alignment = { horizontal: "left", vertical: "middle", wrapText: true };

  // ── R2 (h=27): 副标题 ──
  ws.getRow(2).height = 27;
  ws.mergeCells("A2:I2");
  const c2 = ws.getCell("A2");
  c2.value = "深圳芝麻果实科技有限公司-出港皮历史出货数据";
  c2.font = fontDengXian(12, true, BLACK);
  c2.alignment = centerWrap;
  c2.border = thinBorder;

  // ── R3 (h=27): 基础信息（无「国内税票」标签）──
  ws.getRow(3).height = 27;
  ws.mergeCells("A3:F3");
  const c3a = ws.getCell("A3");
  c3a.value = "基础信息";
  c3a.font = fontDengXian(10.5, true, WHITE);
  c3a.alignment = centerNoWrap;
  c3a.border = thinBorder;
  c3a.fill = headerFill;

  // ── R4 (h=36): 列标题 ──
  ws.getRow(4).height = 36;
  const colHeaders = ["序列", "时间", "平台", "国家", "入库地址\n（代码）", "发货数量", "货代对应识别号", "报关费RMB", "总费用RMB"];
  colHeaders.forEach((h, i) => {
    const cell = ws.getCell(4, i + 1);
    cell.value = h;
    cell.font = fontDengXian(10.5, true, WHITE);
    cell.alignment = centerWrap;
    cell.border = thinBorder;
    cell.fill = headerFill;
  });

  // ── R5 (h=67): 数据行 ──
  ws.getRow(5).height = 67;
  const customs = +bill.customsFee.toFixed(2);
  const dataVals: Array<string | number | null> = [
    1, bill.time, bill.platform || null, bill.country,
    bill.warehouseCode || null, bill.qty, bill.fwdFirst,
    customs, null, // I5 = formula
  ];
  dataVals.forEach((v, i) => {
    if (v === null) return;
    const cell = ws.getCell(5, i + 1);
    cell.value = v;
    cell.font = fontDengXian(10.5, false);
    cell.alignment = centerWrap;
    cell.border = thinBorder;
    if (i === 7) cell.numFmt = NUM_FMT;
  });
  const i5 = ws.getCell("I5");
  i5.value = { formula: "SUM(H5:H5)", result: customs };
  i5.font = fontDengXian(10.5, true);
  i5.alignment = centerNoWrap;
  i5.border = thinBorder;
  i5.numFmt = NUM_FMT;

  // ── R7 (h=40): 温馨提示 ──
  ws.getRow(7).height = 40;
  ws.mergeCells("A7:I7");
  const c7 = ws.getCell("A7");
  c7.value = "温馨提示:感谢贵司对我司业务支持，请收到此账单三个工作日内确认回复，否则我司默认贵司已确认金额，如有问题请及时联系我司客服。";
  c7.font = fontYaHei(12, true);
  c7.alignment = leftWrap;

  // ── R8 (h=130): 报关行收款信息 ──
  ws.getRow(8).height = 130;
  ws.mergeCells("A8:I8");
  const bankText = broker?.bankInfo?.trim() || (broker?.name ? `户名：${broker.name}` : "");
  const c8 = ws.getCell("A8");
  c8.value = bankText;
  c8.font = fontYaHei(14, true);
  c8.alignment = leftWrap;
}

// ============================================================
// 国外账单 Builder（7 列，香港抬头 / USD）
// ============================================================

interface InternationalRow {
  time: string;
  qty: string | number;
  forwarderId: string;
  fee: number;
  currency: string;
}

function buildInternationalSheet(wb: ExcelJS.Workbook, name: string, row: InternationalRow, logoId: number): void {
  const ws = wb.addWorksheet(name, {
    pageSetup: pageSetup("landscape"),
  });

  const colWidths = [4.6091, 10.3545, 8.6, 19.0818, 10, 12, 26];
  colWidths.forEach((w, i) => { ws.getColumn(i + 1).width = w; });

  ws.addImage(logoId, {
    tl: { col: 0.1, row: 0.1 },
    ext: { width: 135, height: 105 },
    editAs: "oneCell",
  });

  // ── R1 (h=84): 香港公司抬头 ──
  ws.getRow(1).height = 84;
  ws.mergeCells("A1:C1");
  ws.mergeCells("D1:G1");
  const c1 = ws.getCell("D1");
  c1.value = "易通科技供应链(香港)有限公司\nETTON TECHNOLOGY SUPPLY CHAIN (HK) CO.LIMITED\n地址：RM A13, 6/F HUNG TO CENTRE, 94-96 HOW MING\nSTREET KWUN TONG KLN,HONG KONG";
  c1.font = fontYaHei(12, true);
  c1.alignment = { horizontal: "left", vertical: "middle", wrapText: true };

  // ── R2 (h=15): 副标题 ──
  ws.getRow(2).height = 15;
  ws.mergeCells("A2:G2");
  const c2 = ws.getCell("A2");
  c2.value = "PAISEEC TECHNOLOGY LIMITED--出港皮历史出货数据";
  c2.font = fontDengXian(12, true, BLACK);
  c2.alignment = centerNoWrap;

  // ── R3 (h=13.1): 基础信息标签 ──
  ws.getRow(3).height = 13.1;
  ws.mergeCells("A3:C3");
  const c3 = ws.getCell("A3");
  c3.value = "基础信息";
  c3.font = fontDengXian(10.5, true, WHITE);
  c3.alignment = centerNoWrap;
  c3.border = thinBorder;
  c3.fill = headerFill;
  for (let c = 4; c <= 7; c++) {
    const cell = ws.getCell(3, c);
    cell.border = thinBorder;
    cell.fill = headerFill;
  }

  // ── R4 (h=26.25): 列标题 ──
  ws.getRow(4).height = 26.25;
  const colHeaders = ["序列", "时间", "发货数量", "货代对应识别号", "费用", "货币", "费用详情"];
  colHeaders.forEach((h, i) => {
    const cell = ws.getCell(4, i + 1);
    cell.value = h;
    cell.font = fontDengXian(10.5, true, WHITE);
    cell.alignment = centerWrap;
    cell.border = thinBorder;
    cell.fill = headerFill;
  });

  // ── R5 (h=31): 数据行 ──
  ws.getRow(5).height = 31;
  const feeVal = +row.fee.toFixed(2);
  const dataVals: Array<string | number> = [
    1, row.time, row.qty, row.forwarderId, feeVal, row.currency || "USD", "",
  ];
  dataVals.forEach((v, i) => {
    const cell = ws.getCell(5, i + 1);
    cell.value = v;
    cell.font = fontDengXian(i === 4 ? 9.75 : 10.5, i === 4);
    cell.alignment = centerWrap;
    cell.border = thinBorder;
    if (i === 4) cell.numFmt = NUM_FMT;
  });

  // ── R8 (h=44): 温馨提示 ──
  ws.getRow(8).height = 44;
  ws.mergeCells("A8:G8");
  const c8 = ws.getCell("A8");
  c8.value = "温馨提示:感谢贵司对我司业务支持，请收到此账单三个工作日内确认回复，否则我司默认贵司已确认金额，如有问题请及时联系我司客服。";
  c8.font = fontYaHei(12, true);
  c8.alignment = leftWrap;

  // ── R9-R16 (h=19): 香港银行信息 ──
  const bankLinesHK = [
    "收款人名稱中文：易通科技供应链(香港)有限公司",
    "收款人名稱英文：ETTON TECHNOLOGY SUPPLY CHAIN (HK) CO.LIMITED",
    "收 款 人 地 址 ： RM A13, 6/F HUNG TO CENTRE, 94-96 HOW MING\nSTREET KWUN TONG KLN,HONG KONG",
    "銀行名稱：The Hongkong and Shanghai Banking Corporation Limited",
    "帳號：747-086304-838",
    "銀行地址：香港皇后大道中 1 號",
    "銀行代碼：004（適用於香港本地付款）",
    "SWIFT CODE: HSBCHKHHHKH（適用於電匯）",
  ];
  const bankHeights = [19, 19, 36, 19, 19, 19, 19, 19];
  bankLinesHK.forEach((text, idx) => {
    const rr = 9 + idx;
    ws.getRow(rr).height = bankHeights[idx];
    ws.mergeCells(`A${rr}:G${rr}`);
    const cell = ws.getCell(`A${rr}`);
    cell.value = text;
    cell.font = fontYaHei(12, true);
    cell.alignment = leftWrap;
  });
}

// ============================================================
// INVOICE Builder（5 列，香港抬头 / 客户发票，无船务信息）
// ============================================================

interface InvoiceInput {
  jobNo: string;
  itemLabel: string;
  amount: number;
}

function buildInvoiceSheet(wb: ExcelJS.Workbook, name: string, input: InvoiceInput, logoId: number): void {
  const ws = wb.addWorksheet(name, {
    pageSetup: pageSetup("portrait"),
  });

  const colWidths = [30.0818, 12.3364, 16.2636, 12.7636, 20.0818];
  colWidths.forEach((w, i) => { ws.getColumn(i + 1).width = w; });

  ws.addImage(logoId, {
    tl: { col: 0.05, row: 0.05 },
    ext: { width: 135, height: 105 },
    editAs: "oneCell",
  });

  // ── 全表边框：R1-R16 整体加 thin 边框 ──
  for (let r = 1; r <= 16; r++) {
    for (let c = 1; c <= 5; c++) {
      ws.getCell(r, c).border = thinBorder;
    }
  }

  // ── R1 (h=30): 香港公司中文 ──
  ws.getRow(1).height = 30;
  ws.mergeCells("A1:E1");
  const c1 = ws.getCell("A1");
  c1.value = " 易通科技供应链(香港)有限公司";
  c1.font = fontSong(18, true);
  c1.alignment = { horizontal: "center" };

  // ── R2 (h=13): 香港公司英文 ──
  ws.getRow(2).height = 13;
  ws.mergeCells("A2:E2");
  const c2 = ws.getCell("A2");
  c2.value = "ETTON TECHNOLOGY SUPPLY CHAIN (HK) CO.LIMITED";
  c2.font = fontTimesNR(8.5, true);
  c2.alignment = centerNoWrap;

  // ── R3 (h=12): 地址 ──
  ws.getRow(3).height = 12;
  ws.mergeCells("A3:E3");
  const c3 = ws.getCell("A3");
  c3.value = "RM A13, 6/F HUNG TO CENTRE, 94-96 HOW MING STREET KWUN TONG KLN,HONG KONG";
  c3.font = fontTimesNR(6, true);
  c3.alignment = centerNoWrap;

  // ── R4: 空行（保留边框）──

  // ── R5-R6 (h=29): INVOICE 标题 ──
  ws.getRow(5).height = 29;
  ws.mergeCells("A5:E6");
  const c5 = ws.getCell("A5");
  c5.value = "INVOICE";
  c5.font = fontArial(18, true);
  c5.alignment = centerNoWrap;

  // ── R7 (h=21.75): TO ──
  ws.getRow(7).height = 21.75;
  ws.mergeCells("A7:E7");
  const c7 = ws.getCell("A7");
  c7.value = "TO:PAISEEC TECHNOLOGY LIMITED";
  c7.font = fontArial(12, true);
  c7.alignment = leftCenter;

  // ── R8 (h=33): TO 地址 ──
  ws.getRow(8).height = 33;
  ws.mergeCells("A8:E8");
  const c8 = ws.getCell("A8");
  c8.value = "Room 511, 5th Floor, Ming Sang Industrial Building, 19-21 Hing Yip Street, Kwun Tong, Hong Kong";
  c8.font = fontArial(12, true);
  c8.alignment = leftWrap;

  // ── R9 (h=24): JOB NO + DATE ──
  ws.getRow(9).height = 24;
  ws.mergeCells("C9:E9");
  const today = new Date().toISOString().slice(0, 10);
  const c9a = ws.getCell("A9");
  c9a.value = `JOB NO:${input.jobNo}`;
  c9a.font = fontArial(12, true);
  c9a.alignment = leftWrap;
  const c9b = ws.getCell("C9");
  c9b.value = `DATE:${today}`;
  c9b.font = fontArial(12, true);
  c9b.alignment = leftWrap;

  // ── R10 (h=24): 表格列标题 ──
  ws.getRow(10).height = 24;
  const invHeaders = ["ITEM", "QUANTITY", "CURRENCY", "PRICE", "AMOUNT"];
  invHeaders.forEach((h, i) => {
    const cell = ws.getCell(10, i + 1);
    cell.value = h;
    cell.font = fontArial(12, true);
    cell.alignment = { horizontal: "center" };
  });

  // ── R11 (h=50): 数据行 ──
  ws.getRow(11).height = 50;
  const price = +input.amount.toFixed(2);
  const dataVals: Array<string | number> = [input.itemLabel, 1, "USD", price];
  dataVals.forEach((val, i) => {
    const cell = ws.getCell(11, i + 1);
    cell.value = val;
    cell.font = fontArial(12, false);
    cell.alignment = { horizontal: i === 0 ? undefined : "center", wrapText: true, vertical: "middle" };
    if (i === 3) cell.numFmt = NUM_FMT;
  });
  const e11 = ws.getCell(11, 5);
  e11.value = { formula: "B11*D11", result: price };
  e11.font = fontArial(12, true);
  e11.alignment = { horizontal: "center", wrapText: true, vertical: "middle" };
  e11.numFmt = NUM_FMT;

  // ── R12-R15: 空行（保留边框）──

  // ── R16 (h=24): Total 行 ──
  ws.getRow(16).height = 24;
  ws.mergeCells("C16:E16");
  const ca16 = ws.getCell("A16");
  ca16.value = "                                                                                    Total:";
  ca16.font = fontArial(14, true);
  const cb16 = ws.getCell("B16");
  cb16.value = "USD";
  cb16.font = fontArial(14, true);
  cb16.alignment = { horizontal: "right" };
  const cc16 = ws.getCell("C16");
  cc16.value = { formula: "SUM(E11:E15)", result: price };
  cc16.font = fontArial(14, true);
  cc16.alignment = { horizontal: "right" };
  cc16.numFmt = NUM_FMT;

  // ── R18 (h=16.85): USD ACCOUNT DETAIL ──
  ws.getRow(18).height = 16.85;
  ws.mergeCells("A18:E18");
  const c18 = ws.getCell("A18");
  c18.value = "USD ACCOUNT DETAIL：";
  c18.font = fontYaHei(12, true);

  // ── R19-R23 (h=21): 银行详情 ──
  const accLines = [
    "ACCOUNT NAME: ETTON TECHNOLOGY SUPPLY CHAIN (HK) CO.LIMITED",
    "ACCOUNT NUMBER FOR USD: 747-086304-838",
    "BANK NAME: The Hongkong and Shanghai Banking Corporation Limited",
    "BANK ADDRESS: 1 Queen's Road Central,Hong Kong",
    "SWIFT CODE: HSBCHKHHHKH",
  ];
  accLines.forEach((text, idx) => {
    const rr = 19 + idx;
    ws.getRow(rr).height = 21;
    ws.mergeCells(`A${rr}:E${rr}`);
    const cell = ws.getCell(`A${rr}`);
    cell.value = text;
    cell.font = fontYaHei(12, true, BLACK);
    cell.alignment = idx === 4 ? { horizontal: "left", vertical: "top" } : leftCenter;
  });
}

// ============================================================
// Buffer helper
// ============================================================

function writeBuffer(wb: ExcelJS.Workbook): Promise<Buffer> {
  return wb.xlsx.writeBuffer().then((ab) => Buffer.from(ab as unknown as ArrayBuffer));
}

// ExcelJS 默认字体硬编码为 Calibri（styles-xform.js，无公开 API），
// 模板默认字体为 宋体 charset=134。用 JSZip 对 styles.xml 做字符串手术，
// 将首个 <font>（默认字体）替换为 宋体。
const DEFAULT_FONT_REF = '<font><sz val="11"/><color theme="1"/><name val="宋体"/><charset val="134"/><scheme val="minor"/></font>';

async function applyChineseDefaultFont(buffer: Buffer): Promise<Buffer> {
  const zip = await JSZip.loadAsync(buffer);
  const stylesPath = "xl/styles.xml";
  const entry = zip.file(stylesPath);
  if (!entry) return buffer;
  const styles = await entry.async("string");
  const replaced = styles.replace(/<font>[\s\S]*?<\/font>/, DEFAULT_FONT_REF);
  zip.file(stylesPath, replaced);
  return Buffer.from(await zip.generateAsync({ type: "nodebuffer" }));
}

// ============================================================
// Main entry
// ============================================================

const DATA_SHEET_NAMES = [
  "1.大货出货数据费用统计",
  "2.港皮海外仓转仓费用明细",
  "3.大货以及其他货物快递退运费用",
  "4.空派快递费用明细",
];

export async function processPipixiongSplit(
  filePath: string,
  fileName: string,
  broker?: BrokerInput
): Promise<PipixiongSplitResult> {
  console.log(`[pipi] 开始处理: ${fileName}`);

  const srcBuffer = readFileSync(filePath);
  const srcWb = new ExcelJS.Workbook();
  // @ts-expect-error — exceljs buffer type
  await srcWb.xlsx.load(srcBuffer);

  const domesticWb = new ExcelJS.Workbook();
  const customsWb = new ExcelJS.Workbook();
  const internationalWb = new ExcelJS.Workbook();
  const invoiceWb = new ExcelJS.Workbook();

  // ── 加载 LOGO ──
  let logoIdD = 0, logoIdC = 0, logoIdI = 0, logoIdV = 0;
  const logoPaths = [
    join(process.cwd(), "public", "etton-logo.png"),
    join(process.cwd(), "测试文档", "ETTON LOGO.png"),
    join(process.cwd(), "..", "对帐工具", "测试文档", "ETTON LOGO.png"),
  ];
  let logoPath = "";
  for (const p of logoPaths) {
    if (existsSync(p)) { logoPath = p; break; }
  }
  if (logoPath) {
    try {
      const logoBuf = readFileSync(logoPath);
      // @ts-expect-error Node Buffer vs ExcelJS Buffer type mismatch
      logoIdD = domesticWb.addImage({ buffer: logoBuf, extension: "png" });
      // @ts-expect-error Node Buffer vs ExcelJS Buffer type mismatch
      logoIdC = customsWb.addImage({ buffer: logoBuf, extension: "png" });
      // @ts-expect-error Node Buffer vs ExcelJS Buffer type mismatch
      logoIdI = internationalWb.addImage({ buffer: logoBuf, extension: "png" });
      // @ts-expect-error Node Buffer vs ExcelJS Buffer type mismatch
      logoIdV = invoiceWb.addImage({ buffer: logoBuf, extension: "png" });
      console.log(`[pipi] LOGO loaded: ${logoPath}`);
    } catch { /* skip logo */ }
  }

  // ── 解析所有数据 sheet（含已支付，用于 JOB NO 去重）──
  const dataWs = srcWb.getWorksheet(DATA_SHEET_NAMES[0]);
  if (!dataWs) throw new Error("找不到 '1.大货出货数据费用统计' sheet");

  const billRows = parseDataSheet(dataWs);

  const usdSheets: { rows: UsdRow[]; ws: ExcelJS.Worksheet | undefined }[] = [];
  const usdOptsList: (UsdSheetOpts | null)[] = [
    null, // sheet 1 已单独解析
    { qtyCol: 4, fwdCol: 5, feeCol: 6, currencyCol: null, detailCol: 7, paidCol: 9, itemLabel: INVOICE_ITEM_WAREHOUSE_TRANSFER },
    { qtyCol: 4, fwdCol: 5, feeCol: 6, currencyCol: 7, detailCol: 8, paidCol: 11, itemLabel: INVOICE_ITEM_EXPRESS_RETURN },
    { qtyCol: 4, fwdCol: 5, feeCol: 6, currencyCol: 7, detailCol: 8, paidCol: 11, itemLabel: INVOICE_ITEM_AIR_EXPRESS },
  ];

  for (let i = 1; i < DATA_SHEET_NAMES.length; i++) {
    const ws = srcWb.getWorksheet(DATA_SHEET_NAMES[i]);
    const opts = usdOptsList[i];
    if (ws && opts) {
      usdSheets.push({ rows: parseUsdSheet(ws, opts), ws });
    } else if (!ws) {
      console.warn(`[pipi] 未找到 sheet: ${DATA_SHEET_NAMES[i]}`);
    }
  }

  const allUsdRows = usdSheets.flatMap((s) => s.rows);

  // ── JOB NO 去重：统计所有记录（含已支付）中货代识别号的出现次数 ──
  const ordered = [...billRows, ...allUsdRows];
  const fwdTotal = new Map<string, number>();
  for (const r of ordered) {
    if (r.forwarderId) fwdTotal.set(r.forwarderId, (fwdTotal.get(r.forwarderId) || 0) + 1);
  }
  const fwdSeen = new Map<string, number>();
  for (const r of ordered) {
    const f = r.forwarderId;
    if (!f) { r.jobNo = f; continue; }
    const n = (fwdSeen.get(f) || 0) + 1;
    fwdSeen.set(f, n);
    r.jobNo = (fwdTotal.get(f) || 1) > 1 ? `${f}-${n}` : f;
  }

  // ── 生成账单 ──
  const groups: SplitGroup[] = [];
  const nextIdx = (wb: ExcelJS.Workbook) => String(wb.worksheets.length + 1);

  for (const bill of billRows) {
    if (isPaid(bill.paid)) continue;

    const hasDomestic = bill.truckFee > 0 || bill.portFee > 0;
    const hasCustoms = bill.customsFee > 0;
    const hasInternational = bill.oceanTaxUSD > 0;
    const hasInvoice = bill.oceanTaxUSD !== 0;

    if (!hasDomestic && !hasCustoms && !hasInternational && !hasInvoice) continue;

    const fwdId = bill.fwdFirst || `SEQ_${bill.seq}`;

    if (hasDomestic) buildDomesticSheet(domesticWb, nextIdx(domesticWb), bill, logoIdD);
    if (hasCustoms) buildCustomsSheet(customsWb, nextIdx(customsWb), bill, broker, logoIdC);
    if (hasInternational) {
      buildInternationalSheet(internationalWb, nextIdx(internationalWb), {
        time: bill.time, qty: bill.qty, forwarderId: bill.fwdFirst,
        fee: bill.oceanTaxUSD, currency: "USD",
      }, logoIdI);
    }
    if (hasInvoice) {
      buildInvoiceSheet(invoiceWb, nextIdx(invoiceWb), {
        jobNo: bill.jobNo || fwdId,
        itemLabel: bill.oceanTaxUSD > 0 ? INVOICE_ITEM_OCEAN : INVOICE_ITEM_COMPENSATION,
        amount: bill.oceanTaxUSD,
      }, logoIdV);
    }

    groups.push({ forwarderId: fwdId, customsNo: bill.customsNo, seq: bill.seq, bill, hasDomestic, hasCustoms, hasInternational, hasInvoice });
    console.log(`[pipi] ${groups.length}. ${fwdId} | 国内=${hasDomestic} 报关=${hasCustoms} 国外=${hasInternational} INVOICE=${hasInvoice}`);
  }

  for (const u of allUsdRows) {
    if (isPaid(u.paid)) continue;
    if (u.fee <= 0) continue;

    const fwdId = u.forwarderId || `SEQ_${u.seq}`;
    buildInternationalSheet(internationalWb, nextIdx(internationalWb), {
      time: u.time, qty: u.qty, forwarderId: u.forwarderId,
      fee: u.fee, currency: u.currency,
    }, logoIdI);
    buildInvoiceSheet(invoiceWb, nextIdx(invoiceWb), {
      jobNo: u.jobNo || fwdId,
      itemLabel: u.itemLabel,
      amount: u.fee,
    }, logoIdV);

    // 构造一个占位 BillRow 用于 groups 展示
    const placeholderBill: BillRow = {
      seq: u.seq, forwarderId: u.forwarderId, fwdFirst: u.forwarderId, customsNo: "",
      time: u.time, platform: "", country: "", warehouseCode: "", qty: u.qty,
      truckFee: 0, customsFee: 0, portFee: 0, oceanTaxUSD: u.fee, totalRMB: 0,
      note1: "", note2: "", paid: u.paid,
    };
    groups.push({ forwarderId: fwdId, customsNo: "", seq: u.seq, bill: placeholderBill, hasDomestic: false, hasCustoms: false, hasInternational: true, hasInvoice: true });
    console.log(`[pipi] ${groups.length}. ${fwdId} | USD(${u.itemLabel})`);
  }

  const domesticBuffer = await applyChineseDefaultFont(await writeBuffer(domesticWb));
  const customsBuffer = await applyChineseDefaultFont(await writeBuffer(customsWb));
  const internationalBuffer = await applyChineseDefaultFont(await writeBuffer(internationalWb));
  const invoiceBuffer = await applyChineseDefaultFont(await writeBuffer(invoiceWb));

  console.log(`[pipi] 完成: ${groups.length} 组 | 国内${domesticWb.worksheets.length} 报关${customsWb.worksheets.length} 国外${internationalWb.worksheets.length} INVOICE${invoiceWb.worksheets.length}`);

  return { sourceFile: fileName, groups, domesticBuffer, customsBuffer, internationalBuffer, invoiceBuffer };
}
