"use client";

import { useState, useCallback, useRef } from "react";

// ============================================================
// Types
// ============================================================

interface ReconIssue {
  soNumber: string;
  tiantuAmount: number | null;
  paymentAmount: number | null;
  difference: number | null;
  diffRate: number | null;
  status: string;
}

interface ReconSummary {
  totalSO: number;
  matchCount: number;
  diffCount: number;
  tiantuOnlyCount: number;
  paymentOnlyCount: number;
}

interface ReconResponse {
  sessionId: string;
  sourceFiles: { tiantu: string; payment: string };
  summary: ReconSummary;
  issueRows: ReconIssue[];
  totalRows: number;
  downloadUrl: string;
  error?: string;
}

// ============================================================
// Helpers
// ============================================================

function fmtMoney(n: number | null): string {
  if (n === null) return "-";
  return n.toLocaleString("zh-CN", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

function fmtDiff(n: number | null): string {
  if (n === null) return "-";
  const prefix = n > 0 ? "+" : "";
  return prefix + n.toLocaleString("zh-CN", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

function fmtRate(n: number | null): string {
  if (n === null) return "-";
  return n.toFixed(2) + "%";
}

const STATUS_BADGES: Record<string, string> = {
  "一致": "bg-emerald-100 text-emerald-700",
  "金额差异": "bg-red-100 text-red-700",
  "天图缺失": "bg-amber-100 text-amber-700",
  "请款缺失": "bg-orange-100 text-orange-700",
};

// ============================================================
// Page
// ============================================================

export default function ReconciliationPage() {
  const [tiantuFile, setTiantuFile] = useState<File | null>(null);
  const [paymentFile, setPaymentFile] = useState<File | null>(null);
  const [isDraggingTiantu, setIsDraggingTiantu] = useState(false);
  const [isDraggingPayment, setIsDraggingPayment] = useState(false);
  const [processing, setProcessing] = useState(false);
  const [result, setResult] = useState<ReconResponse | null>(null);
  const [error, setError] = useState<string | null>(null);
  const tiantuInputRef = useRef<HTMLInputElement>(null);
  const paymentInputRef = useRef<HTMLInputElement>(null);

  const handleSubmit = useCallback(async () => {
    if (!tiantuFile || !paymentFile) return;

    setProcessing(true);
    setError(null);
    setResult(null);

    try {
      const formData = new FormData();
      formData.append("tiantuFile", tiantuFile);
      formData.append("paymentFile", paymentFile);

      const res = await fetch("/api/reconciliation", {
        method: "POST",
        body: formData,
      });

      const data: ReconResponse = await res.json();

      if (!res.ok) {
        setError(data.error || "对账处理失败");
        return;
      }

      setResult(data);
    } catch (err) {
      const msg = err instanceof Error ? err.message : "网络错误，请重试";
      console.error("对账请求失败:", err);
      setError(`请求失败: ${msg}`);
    } finally {
      setProcessing(false);
    }
  }, [tiantuFile, paymentFile]);

  const reset = useCallback(() => {
    setResult(null);
    setError(null);
    setTiantuFile(null);
    setPaymentFile(null);
  }, []);

  const canSubmit = tiantuFile && paymentFile && !processing;

  // Drag handlers
  const makeDragHandlers = (
    setter: (v: boolean) => void,
    fileSetter: (f: File) => void
  ) => ({
    onDragOver: (e: React.DragEvent) => {
      e.preventDefault();
      setter(true);
    },
    onDragLeave: (e: React.DragEvent) => {
      e.preventDefault();
      setter(false);
    },
    onDrop: (e: React.DragEvent) => {
      e.preventDefault();
      setter(false);
      const file = e.dataTransfer.files[0];
      if (file) fileSetter(file);
    },
  });

  const handleFileSelect = (
    e: React.ChangeEvent<HTMLInputElement>,
    setter: (f: File) => void
  ) => {
    const file = e.target.files?.[0];
    if (file) setter(file);
  };

  return (
    <div className="space-y-8">
      {/* Header */}
      <div>
        <h1 className="text-2xl font-bold text-zinc-800">
          📊 天图请款对账
        </h1>
        <p className="mt-1 text-sm text-zinc-500">
          上传天图供应商账单和内部请款明细，自动按 SO 号比对金额差异
        </p>
      </div>

      {/* Upload Area */}
      {!result && (
        <div className="space-y-6">
          {/* Two upload boxes side by side */}
          <div className="grid grid-cols-1 gap-6 md:grid-cols-2">
            {/* 天图账单 */}
            <UploadBox
              label="天图账单"
              description="供应商（天图通逊）发来的账单"
              icon={
                <svg className="mb-3 h-10 w-10 text-emerald-400" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.5} d="M9 17v-2m3 2v-4m3 4v-6m2 10H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                </svg>
              }
              file={tiantuFile}
              isDragging={isDraggingTiantu}
              dragHandlers={makeDragHandlers(setIsDraggingTiantu, setTiantuFile)}
              onClear={() => setTiantuFile(null)}
              borderColor="border-emerald-300"
              activeBorder="border-emerald-400 bg-emerald-50"
              inputRef={tiantuInputRef}
              onFileSelect={(e) => handleFileSelect(e, setTiantuFile)}
            />

            {/* 请款明细 */}
            <UploadBox
              label="请款明细"
              description="内部请款系统导出的明细"
              icon={
                <svg className="mb-3 h-10 w-10 text-blue-400" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.5} d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                </svg>
              }
              file={paymentFile}
              isDragging={isDraggingPayment}
              dragHandlers={makeDragHandlers(setIsDraggingPayment, setPaymentFile)}
              onClear={() => setPaymentFile(null)}
              borderColor="border-blue-300"
              activeBorder="border-blue-400 bg-blue-50"
              inputRef={paymentInputRef}
              onFileSelect={(e) => handleFileSelect(e, setPaymentFile)}
            />
          </div>

          {/* Submit button */}
          <div className="flex justify-center">
            <button
              onClick={handleSubmit}
              disabled={!canSubmit}
              className={`inline-flex items-center gap-2 rounded-lg px-6 py-3 text-sm font-medium shadow-sm transition-colors ${
                canSubmit
                  ? "bg-emerald-600 text-white hover:bg-emerald-700"
                  : "cursor-not-allowed bg-zinc-200 text-zinc-400"
              }`}
            >
              {processing ? (
                <>
                  <div className="h-4 w-4 animate-spin rounded-full border-2 border-white border-t-transparent" />
                  正在对账中...
                </>
              ) : (
                <>
                  <svg className="h-4 w-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2m-6 9l2 2 4-4" />
                  </svg>
                  开始对账
                </>
              )}
            </button>
          </div>

          {processing && (
            <p className="text-center text-sm text-zinc-400">
              正在解析文件、匹配 SO 号、比对金额...
            </p>
          )}
        </div>
      )}

      {/* Error */}
      {error && (
        <div className="flex items-center gap-3 rounded-lg border border-red-300 bg-red-50 p-4">
          <svg className="h-5 w-5 shrink-0 text-red-500" fill="none" viewBox="0 0 24 24" stroke="currentColor">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 8v4m0 4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
          </svg>
          <p className="text-sm text-red-700">{error}</p>
          <button
            onClick={reset}
            className="ml-auto shrink-0 text-sm font-medium text-red-600 hover:text-red-800"
          >
            重新上传
          </button>
        </div>
      )}

      {/* Results */}
      {result && (
        <>
          {/* Action bar */}
          <div className="flex flex-wrap items-center justify-between gap-4 rounded-lg border border-zinc-200 bg-white p-4 shadow-sm">
            <div>
              <p className="text-sm text-zinc-500">
                天图: <span className="font-medium text-zinc-700">{result.sourceFiles.tiantu}</span>
                {" · "}
                请款: <span className="font-medium text-zinc-700">{result.sourceFiles.payment}</span>
              </p>
              <p className="text-sm text-zinc-500">
                共 {result.totalRows} 个 SO 号
                {" · "}
                一致 {result.summary.matchCount}
                {" · "}
                差异 {result.summary.diffCount}
                {" · "}
                缺失 {result.summary.tiantuOnlyCount + result.summary.paymentOnlyCount}
              </p>
            </div>
            <div className="flex gap-3">
              <a
                href={result.downloadUrl}
                className="inline-flex items-center gap-2 rounded-lg bg-emerald-600 px-4 py-2 text-sm font-medium text-white shadow-sm hover:bg-emerald-700 transition-colors"
                download
              >
                <svg className="h-4 w-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                </svg>
                下载对账结果 Excel
              </a>
              <button
                onClick={reset}
                className="inline-flex items-center gap-2 rounded-lg border border-zinc-300 bg-white px-4 py-2 text-sm font-medium text-zinc-600 hover:bg-zinc-50 transition-colors"
              >
                重新对账
              </button>
            </div>
          </div>

          {/* Summary cards */}
          <div className="grid grid-cols-2 gap-4 sm:grid-cols-5">
            <StatCard label="总 SO 数" value={result.summary.totalSO} color="text-zinc-700" bg="bg-zinc-50" />
            <StatCard label="✅ 一致" value={result.summary.matchCount} color="text-emerald-700" bg="bg-emerald-50" />
            <StatCard label="⚠️ 金额差异" value={result.summary.diffCount} color="text-red-700" bg="bg-red-50" />
            <StatCard label="❌ 天图缺失" value={result.summary.paymentOnlyCount} color="text-amber-700" bg="bg-amber-50" />
            <StatCard label="❌ 请款缺失" value={result.summary.tiantuOnlyCount} color="text-orange-700" bg="bg-orange-50" />
          </div>

          {/* Issue rows table — only show if there are issues */}
          {result.issueRows.length > 0 && (
            <div className="rounded-xl border border-zinc-200 bg-white overflow-hidden">
              <div className="border-b border-zinc-200 px-5 py-3">
                <h2 className="text-sm font-semibold text-zinc-700">
                  差异明细（{result.issueRows.length} 条）
                </h2>
              </div>
              <div className="overflow-x-auto">
                <table className="w-full text-sm">
                  <thead>
                    <tr className="border-b border-zinc-200 bg-zinc-50 text-left text-xs font-semibold uppercase tracking-wider text-zinc-500">
                      <th className="px-5 py-2">SO 号</th>
                      <th className="px-5 py-2 text-right">天图金额</th>
                      <th className="px-5 py-2 text-right">请款金额</th>
                      <th className="px-5 py-2 text-right">差异</th>
                      <th className="px-5 py-2 text-right">差异率</th>
                      <th className="px-5 py-2 text-center">状态</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-zinc-100">
                    {result.issueRows.map((row) => (
                      <tr key={row.soNumber} className="hover:bg-zinc-50">
                        <td className="px-5 py-2.5 font-mono text-xs font-medium text-zinc-700">
                          {row.soNumber}
                        </td>
                        <td className="px-5 py-2.5 text-right font-mono text-xs">
                          {fmtMoney(row.tiantuAmount)}
                        </td>
                        <td className="px-5 py-2.5 text-right font-mono text-xs">
                          {fmtMoney(row.paymentAmount)}
                        </td>
                        <td className={`px-5 py-2.5 text-right font-mono text-xs font-semibold ${
                          row.difference !== null && row.difference !== 0
                            ? row.difference > 0 ? "text-red-600" : "text-blue-600"
                            : "text-zinc-400"
                        }`}>
                          {fmtDiff(row.difference)}
                        </td>
                        <td className={`px-5 py-2.5 text-right font-mono text-xs ${
                          row.diffRate !== null && Math.abs(row.diffRate) > 0.1
                            ? "text-red-600 font-semibold"
                            : "text-zinc-500"
                        }`}>
                          {fmtRate(row.diffRate)}
                        </td>
                        <td className="px-5 py-2.5 text-center">
                          <span className={`inline-flex rounded-full px-2 py-0.5 text-xs font-semibold ${STATUS_BADGES[row.status] || "bg-zinc-100 text-zinc-600"}`}>
                            {row.status}
                          </span>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* All good message */}
          {result.issueRows.length === 0 && (
            <div className="flex flex-col items-center justify-center rounded-xl border-2 border-emerald-200 bg-emerald-50 py-12">
              <svg className="mb-4 h-12 w-12 text-emerald-400" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
              </svg>
              <p className="text-lg font-medium text-emerald-700">全部一致 🎉</p>
              <p className="mt-1 text-sm text-emerald-500">所有 SO 号金额完全匹配，无差异</p>
            </div>
          )}
        </>
      )}

      {/* Tip */}
      {!result && !error && (
        <div className="rounded-lg border border-zinc-200 bg-white p-4 text-sm text-zinc-500">
          <p>
            💡 提示：系统自动识别文件中的 SO 号列和金额列。
            如果同一 SO 号有多条记录，会自动汇总金额后再比对。
            对账结果 Excel 中差异行会标红，方便财务快速定位问题。
          </p>
        </div>
      )}
    </div>
  );
}

// ============================================================
// Sub-components
// ============================================================

function UploadBox({
  label,
  description,
  icon,
  file,
  isDragging,
  dragHandlers,
  onClear,
  borderColor,
  activeBorder,
  inputRef,
  onFileSelect,
}: {
  label: string;
  description: string;
  icon: React.ReactNode;
  file: File | null;
  isDragging: boolean;
  dragHandlers: {
    onDragOver: (e: React.DragEvent) => void;
    onDragLeave: (e: React.DragEvent) => void;
    onDrop: (e: React.DragEvent) => void;
  };
  onClear: () => void;
  borderColor: string;
  activeBorder: string;
  inputRef: React.RefObject<HTMLInputElement | null>;
  onFileSelect: (e: React.ChangeEvent<HTMLInputElement>) => void;
}) {
  return (
    <div
      {...dragHandlers}
      onClick={() => !file && inputRef.current?.click()}
      className={`
        flex cursor-pointer flex-col items-center justify-center rounded-xl border-2 border-dashed p-8 transition-colors
        ${file
          ? `${borderColor} bg-white`
          : isDragging
            ? activeBorder
            : "border-zinc-300 bg-zinc-50 hover:border-zinc-400 hover:bg-zinc-100"
        }
      `}
    >
      {file ? (
        <div className="flex flex-col items-center gap-3 text-center">
          <svg className="h-10 w-10 text-emerald-500" fill="none" viewBox="0 0 24 24" stroke="currentColor">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.5} d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
          </svg>
          <div>
            <p className="text-sm font-semibold text-zinc-700">{label}</p>
            <p className="mt-0.5 text-xs text-zinc-500">{file.name}</p>
            <p className="text-xs text-zinc-400">
              {(file.size / 1024).toFixed(1)} KB
            </p>
          </div>
          <button
            onClick={(e) => {
              e.stopPropagation();
              onClear();
            }}
            className="text-xs font-medium text-red-500 hover:text-red-700"
          >
            移除
          </button>
        </div>
      ) : (
        <>
          {icon}
          <p className="text-sm font-medium text-zinc-600">{label}</p>
          <p className="mt-1 text-xs text-zinc-400">{description}</p>
          <p className="mt-2 text-xs text-zinc-400">
            拖拽 .xlsx 文件到此处，或点击上传
          </p>
        </>
      )}
      <input
        ref={inputRef}
        type="file"
        accept=".xlsx,.xls"
        onChange={onFileSelect}
        className="hidden"
      />
    </div>
  );
}

function StatCard({
  label,
  value,
  color,
  bg,
}: {
  label: string;
  value: number;
  color: string;
  bg: string;
}) {
  return (
    <div className={`rounded-lg border border-zinc-200 ${bg} p-4 text-center`}>
      <p className={`text-2xl font-bold ${color}`}>{value}</p>
      <p className="mt-1 text-xs text-zinc-500">{label}</p>
    </div>
  );
}
