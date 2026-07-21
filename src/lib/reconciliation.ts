/**
 * 请款核对（天图对账）核心逻辑
 *
 * 对比供应商（天图通逊）账单与内部请款系统，按 SO 号做 FULL OUTER JOIN，
 * 找出金额差异并生成标红 Excel。
 *
 * 使用 exceljs 生成输出，jszip 打包（供未来扩展）。
 */

import ExcelJS from "exceljs";
import { readFileSync } from "fs";

// ============================================================
// Types
// ============================================================

/** 对账结果行 */
export interface ReconRow {
  soNumber: string;
  tiantuAmount: number | null;
  paymentAmount: number | null;
  difference: number | null;
  diffRate: number | null;
  status: "一致" | "金额差异" | "天图缺失" | "请款缺失";
}

/** 汇总统计 */
export interface ReconSummary {
  totalSO: number;
  matchCount: number;
  diffCount: number;
  tiantuOnlyCount: number;
  paymentOnlyCount: number;
}

/** 顶层结果 */
export interface ReconciliationResult {
  sourceFiles: { tiantu: string; payment: string };
  rows: ReconRow[];
  summary: ReconSummary;
  buffer: Buffer;
}

// ============================================================
// Column detection
// ============================================================

interface HeaderCell {
  col: number;   // 1-based Excel column number
  header: string;
}

interface ColumnDetectResult {
  soCol: number;       // 1-based Excel column number
  amountCol: number;   // 1-based Excel column number
}

/**
 * 自动检测 SO 列和金额列
 * 使用评分制，优先选择最佳匹配列（如 "客户运单号" 优于 "运单号"）
 * headerCells: 附带实际 Excel 列号的表头数组
 */
function detectColumns(headerCells: HeaderCell[]): ColumnDetectResult {
  let bestSoCol = -1;
  let bestSoScore = 0;
  let bestAmountCol = -1;
  let bestAmountScore = 0;

  for (const { col, header } of headerCells) {
    const h = header.toLowerCase().trim();
    if (!h) continue;

    // ── SO 列评分（越高越优先）──
    let soScore = 0;
    if (h === "系统so号" || h === "客户运单号") {
      soScore = 100;
    } else if (h.includes("客户运单")) {
      soScore = 95;
    } else if (h.includes("系统so")) {
      soScore = 90;
    } else if (h.includes("so号") || h === "so") {
      soScore = 85;
    } else if (h.includes("运单号")) {
      soScore = 70;
    } else if (h.includes("运单") || h.includes("单号")) {
      soScore = 50;
    } else if (h.includes("so")) {
      soScore = 40;
    }

    if (soScore > bestSoScore) {
      bestSoScore = soScore;
      bestSoCol = col;
    }

    // ── 金额列评分（越高越优先）──
    let amountScore = 0;
    if (h.includes("人民币") && (h.includes("金额") || h.includes("应收"))) {
      amountScore = 100;
    } else if (h.includes("本位币") && h.includes("金额")) {
      amountScore = 95;
    } else if (h.includes("金额")) {
      amountScore = 80;
    } else if (h.includes("应收金额")) {
      amountScore = 85;
    } else if (h.includes("应收")) {
      amountScore = 70;
    } else if (h.includes("应付金额")) {
      amountScore = 85;
    } else if (h.includes("应付")) {
      amountScore = 65;
    } else if (h.includes("费用") || h.includes("合计") || h.includes("总价") || h.includes("总金额")) {
      amountScore = 60;
    } else if (h === "amount" || h.includes("amount")) {
      amountScore = 50;
    }

    if (amountScore > bestAmountScore) {
      bestAmountScore = amountScore;
      bestAmountCol = col;
    }
  }

  if (bestSoCol === -1) {
    throw new Error(
      `无法识别 SO 号列。表头: ${headerCells.map(c => c.header).join(", ")}。请确认文件包含 "客户运单号" 或 "系统SO号" 列。`
    );
  }
  if (bestAmountCol === -1) {
    throw new Error(
      `无法识别金额列。表头: ${headerCells.map(c => c.header).join(", ")}。请确认文件包含 "金额" 或 "费用" 列。`
    );
  }

  return { soCol: bestSoCol, amountCol: bestAmountCol };
}

// ============================================================
// Data validation
// ============================================================

/**
 * 判断一个值是否像合法的 SO 号
 * 过滤掉中文备注、公司名、表头文字等非数据行
 */
function isValidSONumber(val: string): boolean {
  // 空值
  if (!val) return false;
  // 太长的文本（超过 50 字符）通常是备注/说明
  if (val.length > 50) return false;
  // 纯中文字符（不含数字或字母）通常不是 SO 号
  if (!/[a-zA-Z0-9]/.test(val)) return false;
  // 纯中文 + 少量标点的文字行（如 "户名："、"公司名称"）
  const chineseChars = (val.match(/[一-鿿]/g) || []).length;
  if (chineseChars > 4) return false;
  // 包含至少一个数字（SO 号通常含数字）
  if (!/\d/.test(val)) return false;

  return true;
}

// ============================================================
// Amount cleaning
// ============================================================

/**
 * 清洗金额字符串: 去除 ￥、¥、千分位逗号、空格
 */
function cleanAmount(val: unknown): number {
  if (val === null || val === undefined || val === "") return 0;
  if (typeof val === "number") return val;

  let s = String(val).trim();
  // 移除货币符号和千分位
  s = s.replace(/[￥¥$€£,，\s]/g, "");
  // 处理负号在后面的情况（如 "100-")
  const negMatch = s.match(/^(-?[\d.]+)-?$/);
  if (negMatch) {
    s = negMatch[1];
  }
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

/**
 * 从 ExcelJS CellValue 中提取数值
 */
function cellNumVal(cell: ExcelJS.Cell): number {
  const v = cell.value;
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return v;
  if (typeof v === "string") return cleanAmount(v);
  if (typeof v === "object") {
    if ("result" in v && v.result !== undefined) {
      return cleanAmount(v.result);
    }
    if ("richText" in v && Array.isArray((v as { richText: unknown }).richText)) {
      const text = (v as { richText: Array<{ text: string }> }).richText
        .map((r) => r.text)
        .join("");
      return cleanAmount(text);
    }
  }
  return 0;
}

function cellStrVal(cell: ExcelJS.Cell): string {
  const v = cell.value;
  if (v === null || v === undefined) return "";
  if (typeof v === "string") return v.trim();
  if (typeof v === "number") return String(v);
  if (typeof v === "object") {
    if ("richText" in v && Array.isArray((v as { richText: unknown }).richText)) {
      return (v as { richText: Array<{ text: string }> }).richText
        .map((r) => r.text)
        .join("")
        .trim();
    }
    if ("result" in v) return String((v as { result: unknown }).result ?? "").trim();
  }
  return String(v ?? "").trim();
}

// ============================================================
// File parsing
// ============================================================

/**
 * 解析 Excel 文件，提取 SO 号和金额
 * 返回按 SO 聚合后的 Map (soNumber → totalAmount)
 */
async function parseFile(filePath: string): Promise<Map<string, number>> {
  const buffer = readFileSync(filePath);
  const wb = new ExcelJS.Workbook();
  // @ts-expect-error — exceljs Buffer type vs @types/node Buffer
  await wb.xlsx.load(buffer);

  const ws = wb.worksheets[0];
  if (!ws) {
    throw new Error("Excel 文件中没有找到工作表");
  }

  // 找表头行（扫描前 10 行）
  let headerRow: ExcelJS.Row | null = null;
  let headerRowNum = 0;
  const headerCells: HeaderCell[] = [];

  for (let r = 1; r <= Math.min(ws.rowCount, 10); r++) {
    const row = ws.getRow(r);
    const rowHeaders: string[] = [];
    const rowCells: HeaderCell[] = [];
    row.eachCell((cell, col) => {
      const val = cellStrVal(cell);
      rowHeaders.push(val);
      rowCells.push({ col, header: val });
    });

    // 检测是否包含 SO 关键字
    const hasSO = rowHeaders.some(
      (h) => h.toLowerCase().includes("so") || h.includes("运单") || h.includes("单号")
    );
    if (hasSO) {
      headerRow = row;
      headerRowNum = r;
      headerCells.push(...rowCells);
      break;
    }
  }

  if (!headerRow) {
    throw new Error("找不到表头行。请确认文件包含 SO 号 / 运单号 等列标题。");
  }

  const { soCol, amountCol } = detectColumns(headerCells);
  console.log(`[recon] 表头行 ${headerRowNum}, SO列=${soCol}(Excel), 金额列=${amountCol}(Excel)`);

  // 遍历数据行，提取 SO 和金额
  const soAmountMap = new Map<string, number>();

  for (let r = headerRowNum + 1; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const soCell = row.getCell(soCol);
    const soVal = cellStrVal(soCell);

    if (!soVal) continue; // 空行跳过

    // 跳过明显不是 SO 号的行（中文长篇文字、备注、公司名等）
    if (!isValidSONumber(soVal)) continue;

    const amountVal = cellNumVal(row.getCell(amountCol));

    // 聚合：同一 SO 号累加金额
    const existing = soAmountMap.get(soVal) || 0;
    soAmountMap.set(soVal, existing + amountVal);
  }

  console.log(`[recon] 解析完成: ${soAmountMap.size} 个唯一 SO 号`);
  return soAmountMap;
}

// ============================================================
// FULL OUTER JOIN
// ============================================================

function fullOuterJoin(
  tiantuMap: Map<string, number>,
  paymentMap: Map<string, number>
): ReconRow[] {
  const allSONumbers = new Set([...tiantuMap.keys(), ...paymentMap.keys()]);
  const rows: ReconRow[] = [];

  for (const so of allSONumbers) {
    const tiantuAmount = tiantuMap.get(so) ?? null;
    const paymentAmount = paymentMap.get(so) ?? null;

    let status: ReconRow["status"];
    let difference: number | null = null;
    let diffRate: number | null = null;

    if (tiantuAmount !== null && paymentAmount !== null) {
      // 两边都有 — 比对金额
      const diff = tiantuAmount - paymentAmount;
      const absDiff = Math.abs(diff);
      // 容忍 0.01 的浮点误差
      if (absDiff < 0.02) {
        status = "一致";
      } else {
        status = "金额差异";
        difference = Math.round(diff * 100) / 100; // 保留两位
        diffRate = paymentAmount !== 0
          ? Math.round((diff / paymentAmount) * 10000) / 100
          : null;
      }
    } else if (tiantuAmount !== null) {
      status = "请款缺失";
      difference = tiantuAmount;
    } else {
      status = "天图缺失";
      difference = paymentAmount !== null ? -paymentAmount : null;
    }

    rows.push({
      soNumber: so,
      tiantuAmount: tiantuAmount !== null ? Math.round(tiantuAmount * 100) / 100 : null,
      paymentAmount: paymentAmount !== null ? Math.round(paymentAmount * 100) / 100 : null,
      difference,
      diffRate,
      status,
    });
  }

  // 排序：差异优先 → 缺失其次 → 一致最后
  const statusOrder: Record<string, number> = {
    "金额差异": 0,
    "请款缺失": 1,
    "天图缺失": 2,
    "一致": 3,
  };

  rows.sort((a, b) => {
    const orderDiff = (statusOrder[a.status] ?? 99) - (statusOrder[b.status] ?? 99);
    if (orderDiff !== 0) return orderDiff;
    return a.soNumber.localeCompare(b.soNumber);
  });

  return rows;
}

// ============================================================
// Build output workbook
// ============================================================

async function buildOutputWorkbook(rows: ReconRow[], summary: ReconSummary): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("天图请款对账结果");

  // ── 列定义 ──
  const columns = [
    { header: "SO号", key: "soNumber", width: 22 },
    { header: "天图金额", key: "tiantuAmount", width: 16 },
    { header: "请款金额", key: "paymentAmount", width: 16 },
    { header: "差异", key: "difference", width: 14 },
    { header: "差异率", key: "diffRate", width: 12 },
    { header: "状态", key: "status", width: 16 },
  ];

  // ── 表头行 ──
  const headerRow = ws.getRow(1);
  columns.forEach((col, i) => {
    const cell = headerRow.getCell(i + 1);
    cell.value = col.header;
    cell.font = { bold: true, size: 11, color: { argb: "FFFFFFFF" } };
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FF374151" }, // zinc-700
    };
    cell.alignment = { horizontal: "center", vertical: "middle" };
    cell.border = {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "thin" },
    };
  });
  headerRow.height = 28;

  // 列宽
  columns.forEach((col, i) => {
    ws.getColumn(i + 1).width = col.width;
  });

  // ── 数据行 ──
  const statusLabels: Record<string, string> = {
    "一致": "✅ 一致",
    "金额差异": "⚠️ 金额差异",
    "天图缺失": "❌ 天图缺失",
    "请款缺失": "❌ 请款缺失",
  };

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const excelRow = ws.getRow(i + 2); // 1-indexed, header is row 1

    excelRow.getCell(1).value = row.soNumber;
    excelRow.getCell(2).value = row.tiantuAmount;
    excelRow.getCell(3).value = row.paymentAmount;
    excelRow.getCell(4).value = row.difference;
    excelRow.getCell(5).value = row.diffRate !== null ? row.diffRate / 100 : null; // 转为小数给百分比格式
    excelRow.getCell(6).value = statusLabels[row.status] || row.status;

    // ── 样式 ──
    const isIssue = row.status !== "一致";
    const bgColor = isIssue ? "FFFFB3B3" : "FFFFFFFF"; // 浅红 vs 白色

    for (let c = 1; c <= 6; c++) {
      const cell = excelRow.getCell(c);
      cell.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: bgColor },
      };
      cell.font = {
        size: 10,
        bold: isIssue,
        color: isIssue ? { argb: "FF991B1B" } : { argb: "FF1F2937" }, // 深红 vs zinc-800
      };
    }

    // 金额列格式
    excelRow.getCell(2).numFmt = "#,##0.00";
    excelRow.getCell(3).numFmt = "#,##0.00";
    excelRow.getCell(4).numFmt = "#,##0.00";
    // 差异率列格式 — 百分比
    excelRow.getCell(5).numFmt = "0.00%";
    // 状态列居中
    excelRow.getCell(6).alignment = { horizontal: "center" };

    excelRow.height = 22;
  }

  // ── 空行 + 汇总行 ──
  const summaryRowNum = rows.length + 3; // 空一行
  const summaryRow = ws.getRow(summaryRowNum);

  // 合并前两列作为"汇总"标签
  summaryRow.getCell(1).value = "📊 汇总";
  summaryRow.getCell(1).font = { bold: true, size: 11 };
  summaryRow.getCell(1).alignment = { horizontal: "center" };

  // 汇总内容放在第2列开始
  summaryRow.getCell(2).value = `总SO: ${summary.totalSO} | 一致: ${summary.matchCount} | 金额差异: ${summary.diffCount} | 天图缺失: ${summary.paymentOnlyCount} | 请款缺失: ${summary.tiantuOnlyCount}`;
  summaryRow.getCell(2).font = { bold: true, size: 10 };
  // 合并2-6列
  ws.mergeCells(summaryRowNum, 2, summaryRowNum, 6);

  for (let c = 1; c <= 6; c++) {
    const cell = summaryRow.getCell(c);
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFF3F4F6" }, // zinc-100
    };
    cell.border = {
      top: { style: "medium" },
      left: { style: "thin" },
      bottom: { style: "medium" },
      right: { style: "thin" },
    };
  }
  summaryRow.height = 28;

  // ── 冻结首行 ──
  ws.views = [{ state: "frozen", ySplit: 1 }];

  const arrayBuffer = await wb.xlsx.writeBuffer();
  return Buffer.from(arrayBuffer as unknown as ArrayBuffer);
}

// ============================================================
// Main entry
// ============================================================

export async function processReconciliation(
  tiantuPath: string,
  paymentPath: string,
  tiantuName: string,
  paymentName: string
): Promise<ReconciliationResult> {
  console.log(`[recon] 开始对账: 天图=${tiantuName}, 请款=${paymentName}`);

  // 1. 解析两个文件
  const tiantuMap = await parseFile(tiantuPath);
  const paymentMap = await parseFile(paymentPath);

  console.log(`[recon] 天图: ${tiantuMap.size} SO, 请款: ${paymentMap.size} SO`);

  // 2. FULL OUTER JOIN
  const rows = fullOuterJoin(tiantuMap, paymentMap);

  // 3. 统计
  const summary: ReconSummary = {
    totalSO: rows.length,
    matchCount: rows.filter((r) => r.status === "一致").length,
    diffCount: rows.filter((r) => r.status === "金额差异").length,
    tiantuOnlyCount: rows.filter((r) => r.status === "请款缺失").length,
    paymentOnlyCount: rows.filter((r) => r.status === "天图缺失").length,
  };

  console.log(
    `[recon] 对账完成: 总${summary.totalSO} | 一致${summary.matchCount} | 差异${summary.diffCount} | 天图缺失${summary.paymentOnlyCount} | 请款缺失${summary.tiantuOnlyCount}`
  );

  // 4. 生成输出 Excel
  const buffer = await buildOutputWorkbook(rows, summary);

  return {
    sourceFiles: { tiantu: tiantuName, payment: paymentName },
    rows,
    summary,
    buffer,
  };
}
