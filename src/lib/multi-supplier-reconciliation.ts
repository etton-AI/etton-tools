/**
 * 多供应商对账引擎 — TypeScript 版本
 *
 * 配置文件驱动：上传账单 → 自动识别供应商 → 加载预配置 →
 * 定位 SO 列 + 金额列 → 比对 → 输出
 *
 * 支持 16 家供应商的开箱即用配置
 */

import ExcelJS from "exceljs";
import { readFileSync } from "fs";
import path from "path";

// ============================================================
// Types
// ============================================================

export type ReconStatus = "一致" | "金额差异" | "供应商缺失" | "请款缺失";

export interface MultiReconRow {
  soNumber: string;
  billAmount: number | null;
  paymentAmount: number | null;
  difference: number | null;
  diffRate: number | null;
  status: ReconStatus;
}

export interface MultiReconSummary {
  totalSO: number;
  matchCount: number;
  diffCount: number;
  billOnlyCount: number;
  paymentOnlyCount: number;
}

export interface MultiReconResult {
  supplier: string;
  sourceFiles: { bill: string; payment: string };
  rows: MultiReconRow[];
  summary: MultiReconSummary;
  buffer: Buffer;
}

/** 供应商配置 */
interface SupplierConfig {
  file_pattern: string[];
  sheet_name: string | null;
  header_row: number;
  data_start_row: number;
  so_column: string;
  fallback_so_column?: string | null;
  amount_column: string;
  fallback_amount_column?: string | null;
  currency_column?: string | null;
  skip_keywords: string[];
}

/** 供应商配置集 */
interface SupplierConfigMap {
  [name: string]: SupplierConfig;
}

interface PaymentConfig {
  so_column_keywords: string[];
  amount_column_keywords: string[];
}

// ============================================================
// 供应商配置（内嵌，与 supplier_config.yaml 保持一致）
// ============================================================

const SUPPLIER_CONFIGS: SupplierConfigMap = {
  "天图通逊": {
    file_pattern: ["*天图*", "*通逊*", "*Tiantu*", "*TIANTU*"],
    sheet_name: null,
    header_row: 10,
    data_start_row: 11,
    so_column: "客户运单号",
    fallback_so_column: "运单号",
    amount_column: "应收金额",
    fallback_amount_column: "金额",
    currency_column: "币种",
    skip_keywords: ["合计", "小计", "总计", "应收款对账单"],
  },
  "星链/易通": {
    file_pattern: ["*星链*", "*易通科技物流*", "*ETTON*", "*易通-*", "*易通0709*"],
    sheet_name: null,
    header_row: 5,
    data_start_row: 6,
    so_column: "客户参考号",
    fallback_so_column: "运单号",
    amount_column: "应收金额",
    fallback_amount_column: "应收",
    skip_keywords: ["合计", "小计", "总计", "本期账单", "应收款明细表"],
  },
  "航乐": {
    file_pattern: ["*航乐*"],
    sheet_name: "Sheet1",
    header_row: 4,
    data_start_row: 5,
    so_column: "运单号",
    fallback_so_column: "客户单号",
    amount_column: "合计应收",
    fallback_amount_column: "金额",
    skip_keywords: ["合计", "小计", "总计", "客户对账明细单"],
  },
  "跨境堡/英美": {
    file_pattern: ["*英美*", "*跨境堡*"],
    sheet_name: null,
    header_row: 2,
    data_start_row: 3,
    so_column: "客户运单号",
    fallback_so_column: "运单号",
    amount_column: "金额",
    skip_keywords: ["合计", "人民币总计", "小计"],
  },
  "美琦/皓辉": {
    file_pattern: ["*美琦*", "*皓辉*", "*zsetton*"],
    sheet_name: null,
    header_row: 5,
    data_start_row: 6,
    so_column: "客户运单号",
    fallback_so_column: "订单号",
    amount_column: "合计金额",
    fallback_amount_column: "金额",
    skip_keywords: ["合计", "总计", "小计"],
  },
  "心一": {
    file_pattern: ["*心一*"],
    sheet_name: null,
    header_row: 8,
    data_start_row: 9,
    so_column: "客户运单号",
    fallback_so_column: "运单号",
    amount_column: "人民币应收金额",
    fallback_amount_column: "应收金额",
    skip_keywords: ["应收总计", "小计", "合计", "总计"],
  },
  "凯鑫": {
    file_pattern: ["*凯鑫*"],
    sheet_name: null,
    header_row: 4,
    data_start_row: 5,
    so_column: "客户运单号",
    fallback_so_column: "运单号",
    amount_column: "金额",
    skip_keywords: ["合计", "人民币总计", "小计", "总计"],
  },
  "华威尔": {
    file_pattern: ["*华威尔*"],
    sheet_name: null,
    header_row: 4,
    data_start_row: 5,
    so_column: "客户运单号",
    fallback_so_column: "运单号",
    amount_column: "金额",
    skip_keywords: ["合计", "总计", "小计"],
  },
  "天龙": {
    file_pattern: ["*天龙*"],
    sheet_name: null,
    header_row: 2,
    data_start_row: 3,
    so_column: "客户单号",
    amount_column: "金额",
    skip_keywords: ["合计", "本期账单", "小计", "总计"],
  },
  "松杰": {
    file_pattern: ["*松杰*"],
    sheet_name: null,
    header_row: 3,
    data_start_row: 4,
    so_column: "客户参考号",
    fallback_so_column: "运单号",
    amount_column: "应收金额",
    fallback_amount_column: "费用合计",
    skip_keywords: ["合计", "本期账单", "小计", "总计"],
  },
  "安时达": {
    file_pattern: ["*安时达*"],
    sheet_name: null,
    header_row: 5,
    data_start_row: 6,
    so_column: "单号",
    amount_column: "总价",
    fallback_amount_column: "金额",
    skip_keywords: ["合计Total", "总计", "小计"],
  },
  "鸿珉": {
    file_pattern: ["*鸿珉*"],
    sheet_name: null,
    header_row: 1,
    data_start_row: 2,
    so_column: "原单号",
    amount_column: "保费(RMB)",
    fallback_amount_column: "保费",
    skip_keywords: ["小计", "合计", "总计"],
  },
  "太平洋": {
    file_pattern: ["*太平洋*"],
    sheet_name: null,
    header_row: 2,
    data_start_row: 3,
    so_column: "客户单号",
    amount_column: "金额",
    skip_keywords: ["合计", "本期账单", "小计", "总计"],
  },
  "一腾": {
    file_pattern: ["*一腾*"],
    sheet_name: null,
    header_row: 2,
    data_start_row: 3,
    so_column: "原单号",
    amount_column: "费用合计",
    fallback_amount_column: "金额",
    skip_keywords: ["合计", "小计", "总计"],
  },
  "乐丰": {
    file_pattern: ["*乐丰*"],
    sheet_name: null,
    header_row: 1,
    data_start_row: 2,
    so_column: "运单号",
    fallback_so_column: "客户单号",
    amount_column: "金额",
    fallback_amount_column: "合计",
    skip_keywords: ["合计", "小计", "总计"],
  },
};

const PAYMENT_CONFIG: PaymentConfig = {
  so_column_keywords: ["SO号", "SO", "系统SO号", "运单号", "SO号码"],
  amount_column_keywords: ["应收金额", "人民币金额", "应付金额", "金额", "总金额"],
};

// ============================================================
// 供应商自动识别
// ============================================================

export function detectSupplier(filename: string): string | null {
  const basename = path.basename(filename).toLowerCase();
  const fullname = filename.toLowerCase();

  for (const [name, config] of Object.entries(SUPPLIER_CONFIGS)) {
    for (const pattern of config.file_pattern) {
      const regex = new RegExp(
        "^" + pattern.replace(/\*/g, ".*").replace(/\?/g, ".") + "$",
        "i"
      );
      if (regex.test(basename) || regex.test(fullname)) {
        return name;
      }
    }
  }
  return null;
}

export function getAvailableSuppliers(): string[] {
  return Object.keys(SUPPLIER_CONFIGS);
}

// ============================================================
// 工具函数
// ============================================================

function cleanAmount(val: unknown): number {
  if (val === null || val === undefined || val === "") return 0;
  if (typeof val === "number") return val;

  let s = String(val).trim();
  s = s.replace(/[￥¥$€£,，\s]/g, "");
  const negMatch = s.match(/^(-?[\d.]+)-?$/);
  if (negMatch) s = negMatch[1];
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

function cellStrVal(cell: ExcelJS.Cell): string {
  const v = cell.value;
  if (v === null || v === undefined) return "";
  if (typeof v === "string") return v.trim();
  if (typeof v === "number") return String(v);
  if (typeof v === "boolean") return String(v);
  if (typeof v === "object") {
    if ("richText" in v && Array.isArray((v as { richText: unknown }).richText)) {
      return (v as { richText: Array<{ text: string }> }).richText
        .map((r) => r.text).join("").trim();
    }
    if ("result" in v) return String((v as { result: unknown }).result ?? "").trim();
  }
  return String(v ?? "").trim();
}

function cellNumVal(cell: ExcelJS.Cell): number {
  const v = cell.value;
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return v;
  if (typeof v === "string") return cleanAmount(v);
  if (typeof v === "object") {
    if ("result" in v && v.result !== undefined) return cleanAmount(v.result);
    if ("richText" in v && Array.isArray((v as { richText: unknown }).richText)) {
      const text = (v as { richText: Array<{ text: string }> }).richText
        .map((r) => r.text).join("");
      return cleanAmount(text);
    }
  }
  return 0;
}

function isValidSONumber(val: string): boolean {
  if (!val || val.length > 50) return false;
  if (!/[a-zA-Z0-9]/.test(val)) return false;
  const chineseChars = (val.match(/[一-鿿]/g) || []).length;
  if (chineseChars > 4) return false;
  if (!/\d/.test(val)) return false;
  return true;
}

function isSkipRow(rowText: string, skipKeywords: string[]): boolean {
  return skipKeywords.some((kw) => rowText.includes(kw));
}

// ============================================================
// 列名模糊匹配
// ============================================================

interface HeaderInfo {
  col: number; // 1-based Excel column
  header: string;
}

function fuzzyFindColumn(
  headers: HeaderInfo[],
  target: string,
  fallback?: string | null
): number {
  if (!target) {
    if (fallback) return fuzzyFindColumn(headers, fallback, null);
    return -1;
  }

  const targetClean = target.toLowerCase().replace(/\s+/g, "").replace(/[()（）]/g, "");

  // Level 1: 精确匹配
  for (const h of headers) {
    const hClean = h.header.toLowerCase().replace(/\s+/g, "").replace(/[()（）]/g, "");
    if (hClean === targetClean) return h.col;
  }

  // Level 2: 包含匹配
  for (const h of headers) {
    const hClean = h.header.toLowerCase().replace(/\s+/g, "").replace(/[()（）]/g, "");
    if (hClean.includes(targetClean)) return h.col;
  }

  // Level 3: 关键词拆分匹配
  const keywords = extractKeywords(targetClean);
  if (keywords.length > 0) {
    for (const h of headers) {
      const hClean = h.header.toLowerCase().replace(/\s+/g, "").replace(/[()（）]/g, "");
      if (!hClean) continue;
      const matchCount = keywords.filter((kw) => hClean.includes(kw)).length;
      if (matchCount >= keywords.length * 0.6) return h.col;
    }
  }

  // Level 4: 备选列
  if (fallback) return fuzzyFindColumn(headers, fallback, null);

  return -1;
}

function extractKeywords(text: string): string[] {
  const cleaned = text.replace(/\(.*?\)|（.*?）/g, "");
  const parts = cleaned.match(/[一-鿿]{1,4}|[a-zA-Z]{2,}|\d+/g) || [];
  return parts.filter((p) => p.length >= 2);
}

// ============================================================
// 解析供应商账单
// ============================================================

async function parseSupplierBill(
  filePath: string,
  config: SupplierConfig
): Promise<Map<string, number>> {
  const buffer = readFileSync(filePath);
  const wb = new ExcelJS.Workbook();
  // @ts-expect-error — exceljs Buffer type vs @types/node Buffer
  await wb.xlsx.load(buffer);

  const ws = selectSheet(wb, config.sheet_name);
  if (!ws) throw new Error("Excel 文件中没有找到有效工作表");

  // 找表头行
  const headerRowNum = config.header_row;
  const headerRow = ws.getRow(headerRowNum);
  const headerCells: HeaderInfo[] = [];
  headerRow.eachCell((cell, col) => {
    const val = cellStrVal(cell);
    if (val) headerCells.push({ col, header: val });
  });

  // 定位列
  const soCol = fuzzyFindColumn(headerCells, config.so_column, config.fallback_so_column);
  const amountCol = fuzzyFindColumn(headerCells, config.amount_column, config.fallback_amount_column);

  if (soCol < 0) {
    const hdrList = headerCells.map((h) => h.header).join(", ");
    throw new Error(`无法定位 SO 号列 "${config.so_column}"。表头: ${hdrList}`);
  }
  if (amountCol < 0) {
    const hdrList = headerCells.map((h) => h.header).join(", ");
    throw new Error(`无法定位金额列 "${config.amount_column}"。表头: ${hdrList}`);
  }

  const soHeader = headerCells.find((h) => h.col === soCol)?.header || `col ${soCol}`;
  const amtHeader = headerCells.find((h) => h.col === amountCol)?.header || `col ${amountCol}`;
  console.log(`[multi-recon] 表头行=${headerRowNum}, SO列="${soHeader}"(col ${soCol}), 金额列="${amtHeader}"(col ${amountCol})`);

  // 解析数据行
  const soAmountMap = new Map<string, number>();
  const skipKeywords = config.skip_keywords || ["合计", "小计", "总计"];

  for (let r = config.data_start_row; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const soVal = cellStrVal(row.getCell(soCol));
    if (!soVal || !isValidSONumber(soVal)) continue;

    // 检查是否跳过行
    const rowText = Array.from({ length: Math.min(row.cellCount, 20) }, (_, c) =>
      cellStrVal(row.getCell(c + 1))
    ).join(" ");
    if (isSkipRow(rowText, skipKeywords)) continue;

    const amount = cellNumVal(row.getCell(amountCol));
    soAmountMap.set(soVal, (soAmountMap.get(soVal) || 0) + amount);
  }

  console.log(`[multi-recon] 供应商账单解析完成: ${soAmountMap.size} 个唯一 SO 号`);
  return soAmountMap;
}

function selectSheet(
  wb: ExcelJS.Workbook,
  sheetName: string | null
): ExcelJS.Worksheet | undefined {
  if (sheetName) {
    const ws = wb.getWorksheet(sheetName);
    if (ws) return ws;
  }
  // 自动选择数据最密集的 Sheet
  let best: ExcelJS.Worksheet | undefined;
  let bestScore = 0;
  for (const ws of wb.worksheets) {
    const score = ws.rowCount * Math.min(ws.columnCount || 1, 30);
    if (score > bestScore) {
      bestScore = score;
      best = ws;
    }
  }
  return best || wb.worksheets[0];
}

// ============================================================
// 解析请款明细
// ============================================================

async function parsePaymentFile(filePath: string): Promise<Map<string, number>> {
  const buffer = readFileSync(filePath);
  const wb = new ExcelJS.Workbook();
  // @ts-expect-error — exceljs Buffer type vs @types/node Buffer
  await wb.xlsx.load(buffer);

  const ws = selectSheet(wb, null);
  if (!ws) throw new Error("请款明细文件中没有找到有效工作表");

  // 自动找表头行
  let headerRowNum = 0;
  let headers: HeaderInfo[] = [];
  for (let r = 1; r <= Math.min(ws.rowCount, 15); r++) {
    const row = ws.getRow(r);
    const rowHeaders: HeaderInfo[] = [];
    row.eachCell((cell, col) => {
      const val = cellStrVal(cell);
      if (val) rowHeaders.push({ col, header: val });
    });
    const rowText = rowHeaders.map((h) => h.header).join(" ");
    if (PAYMENT_CONFIG.so_column_keywords.some((kw) => rowText.includes(kw))) {
      headerRowNum = r;
      headers = rowHeaders;
      break;
    }
  }

  if (headers.length === 0) {
    // fallback: 第一行
    headerRowNum = 1;
    const row = ws.getRow(1);
    row.eachCell((cell, col) => {
      const val = cellStrVal(cell);
      if (val) headers.push({ col, header: val });
    });
  }

  // 定位列
  let soCol = -1;
  let amountCol = -1;

  for (const h of headers) {
    const hClean = h.header.toLowerCase().replace(/\s+/g, "");
    if (soCol < 0 && PAYMENT_CONFIG.so_column_keywords.some(
      (kw) => hClean.includes(kw.toLowerCase().replace(/\s+/g, ""))
    )) {
      soCol = h.col;
    }
    if (amountCol < 0 && PAYMENT_CONFIG.amount_column_keywords.some(
      (kw) => hClean.includes(kw.toLowerCase().replace(/\s+/g, ""))
    )) {
      amountCol = h.col;
    }
  }

  // 自动检测
  if (soCol < 0) {
    for (let c = 1; c <= Math.min(10, headers.length || 10); c++) {
      let validCount = 0;
      for (let r = headerRowNum + 1; r <= Math.min(headerRowNum + 5, ws.rowCount); r++) {
        const val = cellStrVal(ws.getRow(r).getCell(c));
        if (isValidSONumber(val)) validCount++;
      }
      if (validCount >= 2) { soCol = c; break; }
    }
  }
  if (amountCol < 0) {
    let bestCol = -1;
    let bestScore = 0;
    for (let c = 1; c <= (headers.length || 20); c++) {
      const hdrText = headers.find((h) => h.col === c)?.header?.toLowerCase().replace(/\s+/g, "") || "";
      // 跳过明显是文本/类型的列
      if (/名称|类型|备注|日期|编号|渠道/.test(hdrText)) continue;

      let score = 0;
      if (/金额|应收|应付|总价|合计|费用/.test(hdrText)) score = 100;
      else if (/单价|数量|重量|体积/.test(hdrText)) score = 20;

      let numCount = 0;
      for (let r = headerRowNum + 1; r <= Math.min(headerRowNum + 5, ws.rowCount); r++) {
        const val = cellNumVal(ws.getRow(r).getCell(c));
        if (val !== 0) numCount++;
      }
      if (numCount >= 2) {
        score += 30;
        if (/rmb|cny|金额/.test(hdrText)) score += 20;
      }

      if (score > bestScore) { bestScore = score; bestCol = c; }
    }
    if (bestCol >= 0) amountCol = bestCol;
  }

  if (soCol < 0) throw new Error("无法识别请款明细中的 SO 号列");
  if (amountCol < 0) throw new Error("无法识别请款明细中的金额列");

  const soHeader = headers.find((h) => h.col === soCol)?.header || `col ${soCol}`;
  const amtHeader = headers.find((h) => h.col === amountCol)?.header || `col ${amountCol}`;
  console.log(`[multi-recon] 请款明细: SO列="${soHeader}"(col ${soCol}), 金额列="${amtHeader}"(col ${amountCol})`);

  // 解析
  const soAmountMap = new Map<string, number>();
  for (let r = headerRowNum + 1; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const soVal = cellStrVal(row.getCell(soCol));
    if (!soVal || !isValidSONumber(soVal)) continue;
    const amount = cellNumVal(row.getCell(amountCol));
    soAmountMap.set(soVal, (soAmountMap.get(soVal) || 0) + amount);
  }

  console.log(`[multi-recon] 请款明细解析完成: ${soAmountMap.size} 个唯一 SO 号`);
  return soAmountMap;
}

// ============================================================
// FULL OUTER JOIN
// ============================================================

function fullOuterJoin(
  billMap: Map<string, number>,
  paymentMap: Map<string, number>
): MultiReconRow[] {
  const allSO = new Set([...billMap.keys(), ...paymentMap.keys()]);
  const rows: MultiReconRow[] = [];

  for (const so of allSO) {
    const billAmount = billMap.get(so) ?? null;
    const paymentAmount = paymentMap.get(so) ?? null;

    let status: ReconStatus;
    let difference: number | null = null;
    let diffRate: number | null = null;

    if (billAmount !== null && paymentAmount !== null) {
      const diff = billAmount - paymentAmount;
      if (Math.abs(diff) < 0.02) {
        status = "一致";
      } else {
        status = "金额差异";
        difference = Math.round(diff * 100) / 100;
        diffRate = paymentAmount !== 0
          ? Math.round((diff / paymentAmount) * 10000) / 100
          : null;
      }
    } else if (billAmount !== null) {
      status = "请款缺失";
      difference = Math.round(billAmount * 100) / 100;
    } else {
      status = "供应商缺失";
      difference = paymentAmount !== null ? Math.round(-paymentAmount * 100) / 100 : null;
    }

    rows.push({
      soNumber: so,
      billAmount: billAmount !== null ? Math.round(billAmount * 100) / 100 : null,
      paymentAmount: paymentAmount !== null ? Math.round(paymentAmount * 100) / 100 : null,
      difference,
      diffRate,
      status,
    });
  }

  // 排序：差异优先
  const statusOrder: Record<string, number> = {
    "金额差异": 0, "请款缺失": 1, "供应商缺失": 2, "一致": 3,
  };
  rows.sort((a, b) => {
    const orderDiff = (statusOrder[a.status] ?? 99) - (statusOrder[b.status] ?? 99);
    if (orderDiff !== 0) return orderDiff;
    return a.soNumber.localeCompare(b.soNumber);
  });

  return rows;
}

// ============================================================
// 生成输出 Excel
// ============================================================

async function buildOutputWorkbook(
  rows: MultiReconRow[],
  summary: MultiReconSummary,
  supplier: string
): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("多供应商对账结果");

  const columns = [
    { header: "SO号", key: "soNumber", width: 22 },
    { header: "供应商金额", key: "billAmount", width: 16 },
    { header: "请款金额", key: "paymentAmount", width: 16 },
    { header: "差异", key: "difference", width: 14 },
    { header: "差异率", key: "diffRate", width: 12 },
    { header: "状态", key: "status", width: 16 },
  ];

  // 表头行
  const headerRow = ws.getRow(1);
  columns.forEach((col, i) => {
    const cell = headerRow.getCell(i + 1);
    cell.value = col.header;
    cell.font = { bold: true, size: 11, color: { argb: "FFFFFFFF" } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF374151" } };
    cell.alignment = { horizontal: "center", vertical: "middle" };
    cell.border = {
      top: { style: "thin" }, left: { style: "thin" },
      bottom: { style: "thin" }, right: { style: "thin" },
    };
  });
  headerRow.height = 28;
  columns.forEach((col, i) => { ws.getColumn(i + 1).width = col.width; });

  // 状态标签
  const statusLabels: Record<string, string> = {
    "一致": "✅ 一致", "金额差异": "⚠️ 金额差异",
    "供应商缺失": "❌ 供应商缺失", "请款缺失": "❌ 请款缺失",
  };

  // 数据行
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const excelRow = ws.getRow(i + 2);

    excelRow.getCell(1).value = row.soNumber;
    excelRow.getCell(2).value = row.billAmount;
    excelRow.getCell(3).value = row.paymentAmount;
    excelRow.getCell(4).value = row.difference;
    excelRow.getCell(5).value = row.diffRate !== null ? row.diffRate / 100 : null;
    excelRow.getCell(6).value = statusLabels[row.status] || row.status;

    const isIssue = row.status !== "一致";
    const bgColor = isIssue ? "FFFFB3B3" : "FFFFFFFF";

    for (let c = 1; c <= 6; c++) {
      const cell = excelRow.getCell(c);
      cell.border = {
        top: { style: "thin" }, left: { style: "thin" },
        bottom: { style: "thin" }, right: { style: "thin" },
      };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: bgColor } };
      cell.font = {
        size: 10, bold: isIssue,
        color: isIssue ? { argb: "FF991B1B" } : { argb: "FF1F2937" },
      };
    }

    excelRow.getCell(2).numFmt = "#,##0.00";
    excelRow.getCell(3).numFmt = "#,##0.00";
    excelRow.getCell(4).numFmt = "#,##0.00";
    excelRow.getCell(5).numFmt = "0.00%";
    excelRow.getCell(6).alignment = { horizontal: "center" };
    excelRow.height = 22;
  }

  // 汇总行
  const sr = rows.length + 3;
  const summaryRow = ws.getRow(sr);
  summaryRow.getCell(1).value = "📊 汇总";
  summaryRow.getCell(1).font = { bold: true, size: 11 };
  summaryRow.getCell(1).alignment = { horizontal: "center" };
  summaryRow.getCell(2).value =
    `供应商: ${supplier} | 总SO: ${summary.totalSO} | 一致: ${summary.matchCount} | 金额差异: ${summary.diffCount} | 供应商缺失: ${summary.paymentOnlyCount} | 请款缺失: ${summary.billOnlyCount}`;
  summaryRow.getCell(2).font = { bold: true, size: 10 };
  ws.mergeCells(sr, 2, sr, 6);

  for (let c = 1; c <= 6; c++) {
    const cell = summaryRow.getCell(c);
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF3F4F6" } };
    cell.border = {
      top: { style: "medium" }, left: { style: "thin" },
      bottom: { style: "medium" }, right: { style: "thin" },
    };
  }
  summaryRow.height = 28;

  ws.views = [{ state: "frozen", ySplit: 1 }];

  const arrayBuffer = await wb.xlsx.writeBuffer();
  return Buffer.from(arrayBuffer as unknown as ArrayBuffer);
}

// ============================================================
// 主入口
// ============================================================

export async function processMultiSupplierReconciliation(
  billPath: string,
  paymentPath: string,
  billName: string,
  paymentName: string,
  supplier?: string
): Promise<MultiReconResult> {
  console.log(`[multi-recon] 开始多供应商对账: bill=${billName}, payment=${paymentName}`);

  // 1. 识别供应商
  const detectedSupplier = supplier || detectSupplier(billName);
  if (!detectedSupplier) {
    throw new Error(
      `无法自动识别供应商。文件名 "${billName}" 不匹配任何已知供应商。` +
      `可用供应商: ${getAvailableSuppliers().join(", ")}`
    );
  }

  const supplierConfig = SUPPLIER_CONFIGS[detectedSupplier];
  if (!supplierConfig) {
    throw new Error(`未找到供应商 "${detectedSupplier}" 的配置。`);
  }

  console.log(`[multi-recon] 识别供应商: ${detectedSupplier}`);

  // 2. 解析账单
  const billMap = await parseSupplierBill(billPath, supplierConfig);

  // 3. 解析请款明细
  const paymentMap = await parsePaymentFile(paymentPath);

  console.log(`[multi-recon] 供应商账单: ${billMap.size} SO, 请款明细: ${paymentMap.size} SO`);

  // 4. FULL OUTER JOIN
  const rows = fullOuterJoin(billMap, paymentMap);

  // 5. 统计
  const summary: MultiReconSummary = {
    totalSO: rows.length,
    matchCount: rows.filter((r) => r.status === "一致").length,
    diffCount: rows.filter((r) => r.status === "金额差异").length,
    billOnlyCount: rows.filter((r) => r.status === "请款缺失").length,
    paymentOnlyCount: rows.filter((r) => r.status === "供应商缺失").length,
  };

  console.log(
    `[multi-recon] 对账完成: 总${summary.totalSO} | 一致${summary.matchCount} | 差异${summary.diffCount} | 供应商缺失${summary.paymentOnlyCount} | 请款缺失${summary.billOnlyCount}`
  );

  // 6. 生成输出 Excel
  const buffer = await buildOutputWorkbook(rows, summary, detectedSupplier);

  return {
    supplier: detectedSupplier,
    sourceFiles: { bill: billName, payment: paymentName },
    rows,
    summary,
    buffer,
  };
}
