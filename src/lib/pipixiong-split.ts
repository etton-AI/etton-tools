/**
 * 皮皮熊账单拆分核心逻辑
 *
 * 从 "1.大货出货数据费用统计" sheet 读取合并账单 →
 * 按报关单号拆分为三种账单 workbooks，严格遵循模板格式：
 *   国内账单（深圳抬头 / RMB）
 *   国外账单（香港抬头 / USD）
 *   INVOICE（香港抬头 / 客户发票）
 */

import ExcelJS from "exceljs";
import { readFileSync, existsSync } from "fs";
import { join } from "path";

// ============================================================
// Types
// ============================================================

export interface BillRow {
  seq: string;
  forwarderId: string;
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
}

export interface SplitGroup {
  forwarderId: string;
  customsNo: string;
  seq: string;
  bill: BillRow;
  hasDomestic: boolean;
  hasInternational: boolean;
}

export interface PipixiongSplitResult {
  sourceFile: string;
  groups: SplitGroup[];
  domesticBuffer: Buffer;
  internationalBuffer: Buffer;
  invoiceBuffer: Buffer;
}

// ============================================================
// Style constants — 与模板完全一致
// ============================================================

// 颜色
const TEAL = "FF18605A"; // 深青色 — 表头背景
const WHITE = "FFFFFFFF"; // 白色字
const BLACK = "FF000000";

// 字体
const fontDengXian = (size: number, bold = false, color?: string) =>
  ({ name: "等线", size, bold, color: color ? { argb: color } : undefined } as Partial<ExcelJS.Font>);

const fontYaHei = (size: number, bold = false, color?: string) =>
  ({ name: "微软雅黑", size, bold, color: color ? { argb: color } : undefined } as Partial<ExcelJS.Font>);

const fontArial = (size: number, bold = false, color?: string) =>
  ({ name: "Arial", size, bold, color: color ? { argb: color } : undefined } as Partial<ExcelJS.Font>);

const fontSong = (size: number, bold = false) =>
  ({ name: "宋体", size, bold } as Partial<ExcelJS.Font>);

const fontTimesNR = (size: number, bold = false) =>
  ({ name: "Times New Roman", size, bold } as Partial<ExcelJS.Font>);

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
    const customsNo = cellStr(row.getCell(11));
    const paid = cellStr(row.getCell(19));
    if (paid === "Y" || paid === "y") continue;

    rows.push({
      seq,
      forwarderId,
      customsNo,
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
  console.log(`[pipi] 解析到 ${rows.length} 条未支付记录`);
  return rows;
}

// ============================================================
// 国内账单 Builder（严格匹配模板格式）
// ============================================================

function buildDomesticSheet(wb: ExcelJS.Workbook, name: string, bill: BillRow, logoId: number): void {
  const ws = wb.addWorksheet(name, {
    pageSetup: { orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
  });

  // 列宽 — 精确匹配模板
  const colWidths = [7, 9.53, 4.61, 7, 9.68, 10.27, 26.68, 9.5, 9.39, 8.5, 13.64];
  colWidths.forEach((w, i) => { ws.getColumn(i + 1).width = w; });

  // ── LOGO: 左上角 A1 位置 ──
  ws.addImage(logoId, {
    tl: { col: 0.1, row: 0.1 },
    ext: { width: 135, height: 105 },
    editAs: "oneCell",
  });

  // ── R1 (h=84): 公司抬头（3行合一格）──
  const r1 = ws.getRow(1);
  r1.height = 84;
  ws.mergeCells("A1:C1");
  ws.mergeCells("D1:K1");
  const c1 = ws.getCell("D1");
  c1.value = "深圳市易通科技供应链有限公司\nSHENZHEN ETTON TECHNOLOGY SUPPLY CHAIN LTD\n地址：深圳市龙华区民治街道民强社区宝山时代大厦26楼2615";
  c1.font = fontYaHei(16, true);
  c1.alignment = { horizontal: "left", vertical: "middle", wrapText: true };

  // ── R2 (h=27): 副标题 ──
  const r2 = ws.getRow(2);
  r2.height = 27;
  ws.mergeCells("A2:K2");
  const c2 = ws.getCell("A2");
  c2.value = "深圳芝麻果实科技有限公司-出港皮历史出货数据";
  c2.font = fontDengXian(12, true, BLACK);
  c2.alignment = centerWrap;
  c2.border = thinBorder;
  // A2 fill is transparent (no fill)

  // ── R3 (h=27): 基础信息 / 国内税票 ──
  const r3 = ws.getRow(3);
  r3.height = 27;
  ws.mergeCells("A3:F3");
  ws.mergeCells("H3:J3");
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
  const r4 = ws.getRow(4);
  r4.height = 36;
  const colHeaders = ["序列", "时间", "平台", "国家", "入库地址\n（代码）", "发货数量", "货代对应识别号", "拖车费RMB", "报关费RMB", "港杂费RMB", "总费用RMB"];
  colHeaders.forEach((h, i) => {
    const cell = ws.getCell(4, i + 1);
    cell.value = h;
    cell.font = fontDengXian(10.5, true, WHITE);
    cell.alignment = centerWrap;
    cell.border = thinBorder;
    cell.fill = headerFill;
  });

  // ── R5 (h=67): 数据行 ──
  const r5 = ws.getRow(5);
  r5.height = 67;
  const truck = +bill.truckFee.toFixed(2);
  const customs = +bill.customsFee.toFixed(2);
  const port = +bill.portFee.toFixed(2);
  const totalFee = +(truck + customs + port).toFixed(2);

  const dataVals: Array<string | number | null> = [
    1, bill.time, bill.platform || null, bill.country,
    bill.warehouseCode || null, bill.qty, bill.forwarderId,
    truck, customs, port, null, // K5 = formula
  ];
  dataVals.forEach((v, i) => {
    if (v === null) return;
    const cell = ws.getCell(5, i + 1);
    cell.value = v;
    cell.font = fontDengXian(10.5, false, i === 0 ? undefined : undefined);
    cell.alignment = centerWrap;
    // 数据行各列边框：A5无上边框，B5-K5 thin
    cell.border = thinBorder;
    if (i >= 7 && i <= 9) cell.numFmt = NUM_FMT;
  });
  // K5: SUM formula with bold
  const k5 = ws.getCell("K5");
  k5.value = { formula: "SUM(H5:J5)", result: totalFee };
  k5.font = fontDengXian(10.5, true);
  k5.alignment = centerNoWrap;
  k5.border = thinBorder;
  k5.numFmt = NUM_FMT;

  // ── R7 (h=40): 温馨提示 ──
  const r7 = ws.getRow(7);
  r7.height = 40;
  ws.mergeCells("A7:K7");
  const c7 = ws.getCell("A7");
  c7.value = "温馨提示:感谢贵司对我司业务支持，请收到此账单三个工作日内确认回复，否则我司默认贵司已确认金额，如有问题请及时联系我司客服。";
  c7.font = fontYaHei(12, true, undefined); // 默认颜色
  c7.alignment = leftWrap;

  // ── R8-R11 (h=25): 银行信息 ──
  const bankLines = [
    "开户人:深圳市易通科技供应链有限公司",
    "开户行:中国银行深圳彩虹支行",
    "人民币账号:7536 7452 1237",
    "美金账号:7744 7515 5905",
  ];
  bankLines.forEach((text, idx) => {
    const rr = 8 + idx;
    const row = ws.getRow(rr);
    row.height = 25;
    ws.mergeCells(`A${rr}:K${rr}`);
    const cell = ws.getCell(`A${rr}`);
    cell.value = text;
    cell.font = fontYaHei(14, true);
    cell.alignment = leftWrap;
  });
}

// ============================================================
// 国外账单 Builder（严格匹配模板格式）
// ============================================================

function buildInternationalSheet(wb: ExcelJS.Workbook, name: string, bill: BillRow, logoId: number): void {
  const ws = wb.addWorksheet(name, {
    pageSetup: { orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
  });

  const colWidths = [4.61, 10.35, 8.6, 19.08, 10, 12, 26];
  colWidths.forEach((w, i) => { ws.getColumn(i + 1).width = w; });

  // ── LOGO: 左上角 ──
  ws.addImage(logoId, {
    tl: { col: 0.1, row: 0.1 },
    ext: { width: 135, height: 105 },
    editAs: "oneCell",
  });

  // ── R1 (h=84): 香港公司抬头 ──
  colWidths.forEach((w, i) => { ws.getColumn(i + 1).width = w; });

  // ── R1 (h=84): 香港公司抬头（2行合一格）──
  const r1 = ws.getRow(1);
  r1.height = 84;
  ws.mergeCells("A1:C1");
  ws.mergeCells("D1:G1");
  const c1 = ws.getCell("D1");
  c1.value = "易通科技供应链(香港)有限公司\nETTON TECHNOLOGY SUPPLY CHAIN (HK) CO.LIMITED\n地址：RM A13, 6/F HUNG TO CENTRE, 94-96 HOW MING\nSTREET KWUN TONG KLN,HONG KONG";
  c1.font = fontYaHei(12, true);
  c1.alignment = { horizontal: "left", vertical: "middle", wrapText: true };

  // ── R2 (h=15): 副标题 ──
  const r2 = ws.getRow(2);
  r2.height = 15;
  ws.mergeCells("A2:G2");
  const c2 = ws.getCell("A2");
  c2.value = "PAISEEC TECHNOLOGY LIMITED--出港皮历史出货数据";
  c2.font = fontDengXian(12, true, BLACK);
  c2.alignment = centerNoWrap;

  // ── R3 (h=13.1): 基础信息标签 ──
  const r3 = ws.getRow(3);
  r3.height = 13.1;
  ws.mergeCells("A3:C3");
  const c3 = ws.getCell("A3");
  c3.value = "基础信息";
  c3.font = fontDengXian(10.5, true, WHITE);
  c3.alignment = centerNoWrap;
  c3.border = thinBorder;
  c3.fill = headerFill;
  // D3:G3 also have thin border + teal fill
  for (let c = 4; c <= 7; c++) {
    const cell = ws.getCell(3, c);
    cell.border = thinBorder;
    cell.fill = headerFill;
  }

  // ── R4 (h=26.25): 列标题 ──
  const r4 = ws.getRow(4);
  r4.height = 26.25;
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
  const r5 = ws.getRow(5);
  r5.height = 31;
  const oceanVal = +bill.oceanTaxUSD.toFixed(2);
  const dataVals: Array<string | number> = [
    1, bill.time, bill.qty, bill.forwarderId, oceanVal, "USD", "",
  ];
  dataVals.forEach((v, i) => {
    const cell = ws.getCell(5, i + 1);
    cell.value = v;
    cell.font = fontDengXian(i === 4 ? 9.75 : 10.5, i === 4, i === 4 ? undefined : undefined);
    cell.alignment = centerWrap;
    cell.border = thinBorder;
    if (i === 4) cell.numFmt = NUM_FMT;
  });
  // 数据行部分样式微调:
  // A5 font 9.75 not bold (sequence), E5 bold 9.75 (fee)

  // ── R8 (h=44): 温馨提示 ──
  const r8 = ws.getRow(8);
  r8.height = 44;
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
    const row = ws.getRow(rr);
    row.height = bankHeights[idx];
    ws.mergeCells(`A${rr}:G${rr}`);
    const cell = ws.getCell(`A${rr}`);
    cell.value = text;
    cell.font = fontYaHei(12, true);
    cell.alignment = leftWrap;
  });
}

// ============================================================
// INVOICE Builder（严格匹配模板格式）
// ============================================================

function buildInvoiceSheet(wb: ExcelJS.Workbook, name: string, bill: BillRow, logoId: number): void {
  const ws = wb.addWorksheet(name, {
    pageSetup: { orientation: "portrait", fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
  });

  const colWidths = [30.08, 12.33, 16.26, 12.76, 20.08];
  colWidths.forEach((w, i) => { ws.getColumn(i + 1).width = w; });

  // ── LOGO: 左上角 ──
  ws.addImage(logoId, {
    tl: { col: 0.05, row: 0.05 },
    ext: { width: 135, height: 105 },
    editAs: "oneCell",
  });

  // ── R1 (h=30): 香港公司中文 ──
  const r1 = ws.getRow(1);
  r1.height = 30;
  ws.mergeCells("A1:E1");
  const c1 = ws.getCell("A1");
  c1.value = " 易通科技供应链(香港)有限公司";
  c1.font = fontSong(18, true);
  c1.alignment = { horizontal: "center" };

  // ── R2 (h=13): 香港公司英文 ──
  const r2 = ws.getRow(2);
  r2.height = 13;
  ws.mergeCells("A2:E2");
  const c2 = ws.getCell("A2");
  c2.value = "ETTON TECHNOLOGY SUPPLY CHAIN (HK) CO.LIMITED";
  c2.font = fontTimesNR(8.5, true);
  c2.alignment = centerNoWrap;

  // ── R3 (h=12): 地址 ──
  const r3 = ws.getRow(3);
  r3.height = 12;
  ws.mergeCells("A3:E3");
  const c3 = ws.getCell("A3");
  c3.value = "RM A13, 6/F HUNG TO CENTRE, 94-96 HOW MING STREET KWUN TONG KLN,HONG KONG";
  c3.font = fontTimesNR(6, true);
  c3.alignment = centerNoWrap;

  // ── R5-R6 (h=29): INVOICE 标题 ──
  ws.getRow(5).height = 29;
  ws.mergeCells("A5:E6");
  const c5 = ws.getCell("A5");
  c5.value = "INVOICE";
  c5.font = fontArial(18, true);
  c5.alignment = centerNoWrap;

  // ── R7 (h=21.75): TO ──
  const r7 = ws.getRow(7);
  r7.height = 21.75;
  ws.mergeCells("A7:E7");
  const c7 = ws.getCell("A7");
  c7.value = "TO:PAISEEC TECHNOLOGY LIMITED";
  c7.font = fontArial(12, true);
  c7.alignment = leftCenter;

  // ── R8 (h=33): TO 地址 ──
  const r8 = ws.getRow(8);
  r8.height = 33;
  ws.mergeCells("A8:E8");
  const c8 = ws.getCell("A8");
  c8.value = "Room 511, 5th Floor, Ming Sang Industrial Building, 19-21 Hing Yip Street, Kwun Tong, Hong Kong";
  c8.font = fontArial(12, true);
  c8.alignment = leftWrap;

  // ── R9 (h=24): JOB NO + DATE ──
  const r9 = ws.getRow(9);
  r9.height = 24;
  ws.mergeCells("C9:E9");
  const today = new Date().toISOString().slice(0, 10);
  const c9a = ws.getCell("A9");
  c9a.value = `JOB NO:${bill.forwarderId}`;
  c9a.font = fontArial(12, true);
  c9a.alignment = leftWrap;
  const c9b = ws.getCell("C9");
  c9b.value = `DATE:${today}`;
  c9b.font = fontArial(12, true);
  c9b.alignment = leftWrap;

  // ── R10-R13 (h=24): 船务信息 ──
  const shipRows: Array<[string, string | null]> = [
    ["POL:NANSHA,CHIAN", "POD:RIYADH,SAUDI ARABIA"],
    ["FLIGHT DETAILS: MEX COSCO GUANGZHOU", "VOLUME : 106W"],
    ["ETD:12/6", "ETA:2/7"],
    ["CNTR NO.:CSNU7920253/CSNU7843835/TGBU4959745/OOCU7045142", null],
  ];
  shipRows.forEach(([left, right], idx) => {
    const rr = 10 + idx;
    const row = ws.getRow(rr);
    row.height = 24;
    const cl = ws.getCell(rr, 1);
    cl.value = left;
    cl.font = fontArial(12, true);
    cl.alignment = idx === 2 ? { horizontal: "left", vertical: "middle", wrapText: true } : leftCenter;
    if (right) {
      ws.mergeCells(`C${rr}:E${rr}`);
      const cr = ws.getCell(`C${rr}`);
      cr.value = right;
      cr.font = fontArial(12, true);
      cr.alignment = leftCenter;
    } else {
      ws.mergeCells(`A${rr}:E${rr}`);
    }
  });

  // ── R15 (h=24): 表格列标题 ──
  const r15 = ws.getRow(15);
  r15.height = 24;
  const invHeaders = ["ITEM", "QUANTITY", "CURRENCY", "PRICE", "AMOUNT"];
  invHeaders.forEach((h, i) => {
    const cell = ws.getCell(15, i + 1);
    cell.value = h;
    cell.font = fontArial(12, true);
    cell.alignment = { horizontal: "center" };
  });

  // ── R16 (h=50): 数据行 ──
  const r16 = ws.getRow(16);
  r16.height = 50;
  const price = +bill.oceanTaxUSD.toFixed(2);
  const itemLabels = ["OCEAN FREIGHT & DESTINATION PORT TAXES", 1, "USD", price];
  const itemAligns: Partial<ExcelJS.Alignment>[] = [
    {}, { horizontal: "center" }, { horizontal: "center" }, { horizontal: "center" },
  ];
  itemLabels.forEach((val, i) => {
    const cell = ws.getCell(16, i + 1);
    cell.value = val;
    cell.font = fontArial(12, false);
    cell.alignment = { ...itemAligns[i], wrapText: true, vertical: "middle" };
    if (i === 3) cell.numFmt = NUM_FMT;
  });
  // AMOUNT formula in E16
  const e16 = ws.getCell(16, 5);
  e16.value = { formula: "B16*D16", result: +price.toFixed(2) } as unknown as ExcelJS.CellValue;
  e16.font = fontArial(12, true);
  e16.alignment = { horizontal: "center", wrapText: true, vertical: "middle" };
  e16.numFmt = NUM_FMT;

  // ── R17-R20: 空行（保留边框）──
  for (let rr = 17; rr <= 20; rr++) {
    for (let c = 1; c <= 5; c++) {
      ws.getCell(rr, c).border = thinBorder;
    }
  }

  // ── R21 (h=24): Total 行 ──
  const r21 = ws.getRow(21);
  r21.height = 24;
  ws.mergeCells("C21:E21");
  // A21 has spaces for right alignment effect
  const ca21 = ws.getCell("A21");
  ca21.value = "                                                                                    Total:";
  ca21.font = fontArial(14, true);
  const cb21 = ws.getCell("B21");
  cb21.value = "USD";
  cb21.font = fontArial(14, true);
  cb21.alignment = { horizontal: "right" };
  const cc21 = ws.getCell("C21");
  cc21.value = { formula: "SUM(E16:E20)", result: +price.toFixed(2) };
  cc21.font = fontArial(14, true);
  cc21.alignment = { horizontal: "right" };
  cc21.numFmt = NUM_FMT;

  // ── R23 (h=16.85): USD ACCOUNT DETAIL ──
  const r23 = ws.getRow(23);
  r23.height = 16.85;
  ws.mergeCells("A23:E23");
  const c23 = ws.getCell("A23");
  c23.value = "USD ACCOUNT DETAIL：";
  c23.font = fontYaHei(12, true);

  // ── R24-R28 (h=21): 银行详情 ──
  const accLines = [
    "ACCOUNT NAME: ETTON TECHNOLOGY SUPPLY CHAIN (HK) CO.LIMITED",
    "ACCOUNT NUMBER FOR USD: 747-086304-838",
    "BANK NAME: The Hongkong and Shanghai Banking Corporation Limited",
    "BANK ADDRESS: 1 Queen's Road Central,Hong Kong",
    "SWIFT CODE: HSBCHKHHHKH",
  ];
  accLines.forEach((text, idx) => {
    const rr = 24 + idx;
    const row = ws.getRow(rr);
    row.height = 21;
    ws.mergeCells(`A${rr}:E${rr}`);
    const cell = ws.getCell(`A${rr}`);
    cell.value = text;
    cell.font = fontYaHei(12, true, BLACK);
    cell.alignment = idx === 4 ? { horizontal: "left", vertical: "top" } : leftCenter;
  });
}

// ============================================================
// Sheet copy helper
// ============================================================

function copySheet(src: ExcelJS.Worksheet, dst: ExcelJS.Worksheet): void {
  for (let c = 1; c <= (src.columnCount || 30); c++) {
    const srcCol = src.getColumn(c);
    const dstCol = dst.getColumn(c);
    if (srcCol.width) dstCol.width = srcCol.width;
  }
  for (let r = 1; r <= src.rowCount; r++) {
    const srcRow = src.getRow(r);
    const dstRow = dst.getRow(r);
    if (srcRow.height && srcRow.height > 0) dstRow.height = srcRow.height;
    srcRow.eachCell({ includeEmpty: true }, (srcCell, col) => {
      const dstCell = dstRow.getCell(col);
      dstCell.value = srcCell.value;
      if (srcCell.style) dstCell.style = srcCell.style;
      if (srcCell.font) dstCell.font = srcCell.font;
      if (srcCell.alignment) dstCell.alignment = srcCell.alignment;
      if (srcCell.border) dstCell.border = srcCell.border;
      if (srcCell.fill) dstCell.fill = srcCell.fill;
      if (srcCell.numFmt) dstCell.numFmt = srcCell.numFmt;
    });
  }
  if (src.model.merges) {
    for (const merge of src.model.merges) {
      try { dst.mergeCells(merge); } catch { /* skip conflicts */ }
    }
  }
  if (src.views) dst.views = src.views;
}

function writeBuffer(wb: ExcelJS.Workbook): Promise<Buffer> {
  return wb.xlsx.writeBuffer().then((ab) => Buffer.from(ab as unknown as ArrayBuffer));
}

// ============================================================
// Main entry
// ============================================================

export async function processPipixiongSplit(
  filePath: string,
  fileName: string
): Promise<PipixiongSplitResult> {
  console.log(`[pipi] 开始处理: ${fileName}`);

  const srcBuffer = readFileSync(filePath);
  const srcWb = new ExcelJS.Workbook();
  // @ts-expect-error — exceljs buffer type
  await srcWb.xlsx.load(srcBuffer);

  const domesticWb = new ExcelJS.Workbook();
  const internationalWb = new ExcelJS.Workbook();
  const invoiceWb = new ExcelJS.Workbook();

  // ── 加载 LOGO ──
  let logoIdD = 0, logoIdI = 0, logoIdV = 0;
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
      logoIdI = internationalWb.addImage({ buffer: logoBuf, extension: "png" });
      // @ts-expect-error Node Buffer vs ExcelJS Buffer type mismatch
      logoIdV = invoiceWb.addImage({ buffer: logoBuf, extension: "png" });
      console.log(`[pipi] LOGO loaded: ${logoPath}`);
    } catch { /* skip logo */ }
  }

  // 复制参考 sheets
  const refSheets = ["出货统计表", "货代 & 海外仓代码", "国家以及对应代码"];
  for (const name of refSheets) {
    const srcWs = srcWb.getWorksheet(name);
    if (srcWs) {
      for (const wb of [domesticWb, internationalWb, invoiceWb]) {
        copySheet(srcWs, wb.addWorksheet(name));
      }
    }
  }

  const dataWs = srcWb.getWorksheet("1.大货出货数据费用统计");
  if (!dataWs) throw new Error("找不到 '1.大货出货数据费用统计' sheet");

  const allRows = parseDataSheet(dataWs);
  if (allRows.length === 0) throw new Error("未找到未支付的记录");

  const groups: SplitGroup[] = [];

  for (const bill of allRows) {
    const hasDomestic = bill.truckFee > 0 || bill.customsFee > 0 || bill.portFee > 0;
    const hasInternational = bill.oceanTaxUSD > 0;
    if (!hasDomestic && !hasInternational) continue;

    const fwdId = bill.forwarderId || `SEQ_${bill.seq}`;

    if (hasDomestic) {
      buildDomesticSheet(domesticWb, String(domesticWb.worksheets.length - 2), bill, logoIdD);
    }
    if (hasInternational) {
      buildInternationalSheet(internationalWb, String(internationalWb.worksheets.length - 2), bill, logoIdI);
      buildInvoiceSheet(invoiceWb, String(invoiceWb.worksheets.length - 2), bill, logoIdV);
    }

    groups.push({ forwarderId: fwdId, customsNo: bill.customsNo, seq: bill.seq, bill, hasDomestic, hasInternational });
    console.log(`[pipi] ${groups.length}. ${fwdId} / ${bill.customsNo} | 国内=${hasDomestic} 国际=${hasInternational}`);
  }

  const domesticBuffer = await writeBuffer(domesticWb);
  const internationalBuffer = await writeBuffer(internationalWb);
  const invoiceBuffer = await writeBuffer(invoiceWb);

  console.log(`[pipi] 完成: ${groups.length} 组`);

  return { sourceFile: fileName, groups, domesticBuffer, internationalBuffer, invoiceBuffer };
}
