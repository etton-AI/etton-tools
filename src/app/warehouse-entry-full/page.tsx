"use client";

import { useState, useCallback, useRef } from "react";
import type { SuggestionRow } from "@/lib/warehouse-entry-full";

// ============================================================
// Types
// ============================================================

interface UploadResponse {
  rows: SuggestionRow[];
  supplierTotal: number;
  summary: {
    customerCount: number;
    boxCount: number;
    matchedCount: number;
    unmatchedCount: number;
  };
  error?: string;
}

// ============================================================
// Helpers
// ============================================================

function fmt(n: number, digits = 2): string {
  return n.toLocaleString("zh-CN", {
    minimumFractionDigits: digits,
    maximumFractionDigits: digits,
  });
}

// ============================================================
// 报警说明（点击「报警说明」查看的解读文案）
// ============================================================

type AlarmTone = "amber" | "red" | "blue" | "purple" | "gray";

const ALARM_TONE: Record<AlarmTone, string> = {
  amber: "bg-amber-100 text-amber-700",
  red: "bg-red-600 text-white",
  blue: "bg-blue-100 text-blue-700",
  purple: "bg-purple-600 text-white",
  gray: "bg-zinc-100 text-zinc-600",
};

const ALARM_HELP: {
  text: string;
  tone: AlarmTone;
  condition: string;
  example: string;
  meaning: string;
}[] = [
  {
    text: "⚠需人工复核",
    tone: "amber",
    condition: "该 FBA 号在供应商货箱清单里找不到对应箱子，系统无法自动选数。",
    example: "客户有 FBA19XXXX，但供应商清单里没有该 FBA 号的箱 → 出给客户沿用客户原值并标注。",
    meaning: "提醒人工介入，手动核对并填写出给客户的箱规。",
  },
  {
    text: "三边和差异超限，请核查过机图",
    tone: "red",
    condition: "材积主导（材积重 ≥ 实重）时，|出给客户三边和 − 客户三边和| ≥ 6 cm。",
    example: "客户 40+30+25=95，出给客户 45+35+30=110，差 15 ≥ 6 → 报警。",
    meaning: "建议箱子的总边长明显偏离客户申报，可能体积数据不符，需查过机图核实。",
  },
  {
    text: "材积重差异超限，请核查过机图",
    tone: "red",
    condition: "材积主导时，|出给客户材积重 − 客户材积重| ≥ 2 kg。",
    example: "客户材积重 5.0kg，出给客户 7.5kg，差 2.5 ≥ 2 → 报警。",
    meaning: "建议材积重比客户申报重太多，会推高计费重，需核实过机数据。",
  },
  {
    text: "实重差异超限，请核查过机图",
    tone: "red",
    condition: "实重主导（材积重 < 实重）时，|出给客户实重 − 客户实重| ≥ 0.5 kg。",
    example: "客户实重 3.0kg，出给客户 3.6kg，差 0.6 ≥ 0.5 → 报警。",
    meaning: "重货的实重比客户申报重了 0.5kg 以上，需核实过机重量。",
  },
  {
    text: "供应商存在过大箱，请核查过机图",
    tone: "red",
    condition: "供应商该 FBA 号下最大箱的材积重 − 出给客户材积重 ≥ 2 kg。",
    example: "某 FBA 号 10 箱，供应商最大箱材积重 15kg，出给客户 12kg，差 3 ≥ 2 → 报警。",
    meaning: "供应商货里混进了明显偏大的异常箱，即使建议值取自其它箱，也提示查供应商过机图核实。",
  },
  {
    text: "建议数据计费重小于客户，请确认",
    tone: "amber",
    condition: "出给客户计费重 < 客户申报计费重。",
    example: "客户计费重 8kg，出给客户建议 7kg → 报警。",
    meaning: "建议值比客户自己报的还小，可能少计费，需确认是否采纳。",
  },
  {
    text: "成本重过大，请和供应商申请",
    tone: "purple",
    condition: "同一 FBA 号，出给客户总计费重 − 供应商该 FBA 号总计费重 < 0（负差）。",
    example: "某 FBA 号出给客户总重 500kg，供应商实际 560kg，差 −60 < 0 → 报警。",
    meaning: "给客户报的计费重低于供应商实际收我们的计费重，这票可能亏损，需找供应商申请。",
  },
  {
    text: "[全局调整]（导出时标记）",
    tone: "gray",
    condition: "整单出给客户总计费重 ≤ 供应商总计费重时，导出前等比例放大（受差异约束限制）。",
    example: "整单出给客户 1000kg ≤ 供应商 1100kg → 触发，导出备注标 [全局调整]。",
    meaning: "防止整单给客户的计费重低于供应商成本；若放大后仍不足，导出时弹窗提示可能亏损。",
  },
];

// ============================================================
// Page
// ============================================================

export default function WarehouseEntryPage() {
  const [customerFile, setCustomerFile] = useState<File | null>(null);
  const [supplierFile, setSupplierFile] = useState<File | null>(null);
  const [isDraggingCustomer, setIsDraggingCustomer] = useState(false);
  const [isDraggingSupplier, setIsDraggingSupplier] = useState(false);
  const [processing, setProcessing] = useState(false);
  const [exporting, setExporting] = useState(false);
  const [rows, setRows] = useState<SuggestionRow[]>([]);
  const [supplierTotal, setSupplierTotal] = useState(0);
  const [summary, setSummary] = useState<UploadResponse["summary"] | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [notice, setNotice] = useState<string | null>(null);
  const [showAlarmHelp, setShowAlarmHelp] = useState(false);
  const customerInputRef = useRef<HTMLInputElement>(null);
  const supplierInputRef = useRef<HTMLInputElement>(null);

  const handleSubmit = useCallback(async () => {
    if (!customerFile || !supplierFile) return;
    setProcessing(true);
    setError(null);
    setRows([]);
    setSummary(null);

    try {
      const formData = new FormData();
      formData.append("customerFile", customerFile);
      formData.append("supplierFile", supplierFile);

      const res = await fetch("/api/warehouse-entry-full", {
        method: "POST",
        body: formData,
      });
      const data: UploadResponse = await res.json();

      if (!res.ok) {
        setError(data.error || "处理失败");
        return;
      }

      setRows(data.rows);
      setSupplierTotal(data.supplierTotal);
      setSummary(data.summary);
    } catch (err) {
      const msg = err instanceof Error ? err.message : "网络错误，请重试";
      console.error("TR入仓请求失败:", err);
      setError(`请求失败: ${msg}`);
    } finally {
      setProcessing(false);
    }
  }, [customerFile, supplierFile]);

  const reset = useCallback(() => {
    setRows([]);
    setSummary(null);
    setSupplierTotal(0);
    setError(null);
    setNotice(null);
    setCustomerFile(null);
    setSupplierFile(null);
  }, []);

  // 导出
  const handleExport = useCallback(async () => {
    if (rows.length === 0) return;
    setExporting(true);
    setError(null);
    setNotice(null);

    try {
      const res = await fetch("/api/warehouse-entry-full/export", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ rows, supplierTotal }),
      });

      if (!res.ok) {
        const data = await res.json().catch(() => null);
        setError(data?.error || "导出失败");
        return;
      }

      const warningHeader = res.headers.get("X-Warning");

      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      const totalBoxes = rows.reduce((s, r) => s + (r.totalBoxes || 0), 0);
      a.download = `内部三类数据_${new Date().toISOString().slice(0, 10)}_合计总箱数${totalBoxes}.xlsx`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);

      if (warningHeader) {
        // 全局兜底放大后总重仍不足：弹出可能亏损提示
        window.alert(decodeURIComponent(warningHeader));
      }

      setNotice("已导出三类数据的 Excel");
    } catch (err) {
      console.error("导出失败:", err);
      setError("导出失败，请重试");
    } finally {
      setExporting(false);
    }
  }, [rows, supplierTotal]);

  const canSubmit = customerFile && supplierFile && !processing;

  // 全局校验：出给客户总计费重 vs 供应商总计费重
  const customerTotal = rows.reduce(
    (s, r) => s + r.suggestion.chargeableWeight * r.totalBoxes,
    0
  );
  const globalDiff = customerTotal - supplierTotal;
  const globalWarn = rows.length > 0 && globalDiff <= 0;

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

  const hasResult = rows.length > 0;

  return (
    <div className="space-y-8">
      {/* Header */}
      <div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
        <div>
          <h1 className="text-2xl font-bold text-deep">📦 TR全量入仓数据整理</h1>
          <p className="mt-1 text-sm text-zinc-500">
            上传客户数据 + 供应商数据，逐箱匹配、自动放大尺寸、校验报警，生成「出给客户」建议箱规
          </p>
        </div>
      </div>

      {/* Upload Area */}
      {!hasResult && (
        <div className="space-y-6">
          <div className="grid grid-cols-1 gap-6 md:grid-cols-2">
            <UploadBox
              label="客户数据"
              description="自动识别 标准 / 易通发票 / 货箱清单 三种格式"
              file={customerFile}
              isDragging={isDraggingCustomer}
              dragHandlers={makeDragHandlers(setIsDraggingCustomer, setCustomerFile)}
              onClear={() => setCustomerFile(null)}
              inputRef={customerInputRef}
              onFileSelect={(e) => handleFileSelect(e, setCustomerFile)}
            />
            <UploadBox
              label="供应商数据"
              description="逐箱数据（自动识别 天图 / 英美入仓 / 给总部 三种格式）"
              file={supplierFile}
              isDragging={isDraggingSupplier}
              dragHandlers={makeDragHandlers(setIsDraggingSupplier, setSupplierFile)}
              onClear={() => setSupplierFile(null)}
              inputRef={supplierInputRef}
              onFileSelect={(e) => handleFileSelect(e, setSupplierFile)}
            />
          </div>

          <div className="flex justify-center">
            <button
              onClick={handleSubmit}
              disabled={!canSubmit}
              className={`inline-flex items-center gap-2 rounded-lg px-6 py-3 text-sm font-medium shadow-sm transition-colors ${
                canSubmit
                  ? "bg-primary text-white hover:bg-primary-dark"
                  : "cursor-not-allowed bg-zinc-200 text-zinc-400"
              }`}
            >
              {processing ? (
                <>
                  <div className="h-4 w-4 animate-spin rounded-full border-2 border-white border-t-transparent" />
                  正在整理中...
                </>
              ) : (
                <>开始整理</>
              )}
            </button>
          </div>

          <div className="rounded-lg border border-zinc-200 bg-white p-4 text-sm text-zinc-500">
            <p>💡 提示：系统按「FBA ID（供应商货箱编号前 12 位）」匹配箱组，逐箱（完整 FBA+U0000 编码）输出建议值并校验计费重。客户材积重 &gt; 实重时按 11 级修正规则放大供应商尺寸，否则取该箱原值。整理完成后直接导出三类数据 Excel。</p>
            <button
              onClick={() => setShowAlarmHelp(true)}
              className="mt-2 inline-flex items-center gap-1.5 text-sm font-medium text-primary hover:underline"
            >
              📖 报警说明
            </button>
          </div>
        </div>
      )}

      {/* Error */}
      {error && (
        <div className="flex items-center gap-3 rounded-lg border border-red-300 bg-red-50 p-4">
          <svg className="h-5 w-5 shrink-0 text-red-500" fill="none" viewBox="0 0 24 24" stroke="currentColor">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 8v4m0 4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
          </svg>
          <p className="text-sm text-red-700">{error}</p>
          <button onClick={reset} className="ml-auto shrink-0 text-sm font-medium text-red-600 hover:text-red-800">
            重新上传
          </button>
        </div>
      )}

      {/* Notice */}
      {notice && (
        <div className="flex items-center gap-3 rounded-lg border border-emerald-300 bg-emerald-50 p-4">
          <p className="text-sm text-emerald-700">{notice}</p>
        </div>
      )}

      {/* Results */}
      {hasResult && (
        <>
          {/* Action bar */}
          <div className="flex flex-wrap items-center justify-between gap-4 rounded-lg border border-zinc-200 bg-white p-4 shadow-sm">
            <div>
              <p className="text-sm text-zinc-500">
                客户 <span className="font-medium text-zinc-700">{summary?.customerCount}</span> 个产品
                {" · "}
                供应商 <span className="font-medium text-zinc-700">{summary?.boxCount}</span> 箱
                {" · "}
                匹配 <span className="font-medium text-zinc-700">{summary?.matchedCount}</span>
                {" · "}
                需复核 <span className="font-medium text-amber-600">{summary?.unmatchedCount}</span>
              </p>
            </div>
            <div className="flex flex-wrap items-center gap-3">
              <button
                onClick={reset}
                className="inline-flex items-center gap-2 rounded-lg border border-zinc-300 bg-white px-4 py-2 text-sm font-medium text-zinc-600 hover:bg-zinc-50 transition-colors"
              >
                重新整理
              </button>
            </div>
          </div>

          {/* 报警说明入口 */}
          <div className="flex justify-end">
            <button
              onClick={() => setShowAlarmHelp(true)}
              className="inline-flex items-center gap-1.5 rounded-lg border border-zinc-200 bg-white px-3 py-1.5 text-xs font-medium text-zinc-600 hover:border-primary hover:text-primary transition-colors"
            >
              <span className="text-sm">📖</span> 报警说明
            </button>
          </div>

          {/* 全局校验条 */}
          <div
            className={`flex items-center justify-between gap-4 rounded-lg border p-4 ${
              globalWarn
                ? "border-red-300 bg-red-50"
                : "border-emerald-300 bg-emerald-50"
            }`}
          >
            <div className="flex items-center gap-3">
              <span className={`text-lg font-bold ${globalWarn ? "text-red-700" : "text-emerald-700"}`}>
                全局校验
              </span>
              <p className="text-sm">
                <span className="font-medium">出给客户总计费重 {fmt(customerTotal)}</span>
                {" − "}
                <span className="font-medium">供应商总计费重 {fmt(supplierTotal)}</span>
                {" = "}
                <span className={`font-bold ${globalWarn ? "text-red-700" : "text-emerald-700"}`}>
                  {globalWarn ? "≤ 0" : fmt(globalDiff)}
                </span>
              </p>
            </div>
            {globalWarn && (
              <span className="rounded-full bg-red-100 px-3 py-1 text-xs font-semibold text-red-700">
                ⚠ 出给客户总重偏小，导出时将尝试放大（受差异约束），仍不足会提示可能亏损
              </span>
            )}
          </div>

          {/* 导出按钮 */}
          <div className="flex justify-center">
            <button
              onClick={handleExport}
              disabled={exporting}
              className="inline-flex items-center gap-2 rounded-lg bg-primary px-6 py-3 text-sm font-medium text-white shadow-sm hover:bg-primary-dark transition-colors disabled:cursor-not-allowed disabled:bg-zinc-300"
            >
              {exporting ? (
                <>
                  <div className="h-4 w-4 animate-spin rounded-full border-2 border-white border-t-transparent" />
                  正在导出...
                </>
              ) : (
                <>确认导出三类数据的 Excel</>
              )}
            </button>
          </div>
        </>
      )}

      {/* 报警说明弹窗 */}
      {showAlarmHelp && (
        <div
          className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 p-4"
          onClick={() => setShowAlarmHelp(false)}
        >
          <div
            className="flex max-h-[85vh] w-full max-w-3xl flex-col overflow-hidden rounded-xl bg-white shadow-2xl"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="flex items-center justify-between border-b border-zinc-200 px-6 py-4">
              <h2 className="text-lg font-bold text-deep">报警说明</h2>
              <button
                onClick={() => setShowAlarmHelp(false)}
                className="inline-flex h-8 w-8 items-center justify-center rounded-lg text-zinc-400 hover:bg-zinc-100 hover:text-zinc-600 transition-colors"
                aria-label="关闭"
              >
                ✕
              </button>
            </div>
            <div className="space-y-4 overflow-y-auto px-6 py-5">
              <div className="rounded-lg border border-zinc-200 bg-zinc-50 p-4 text-xs leading-relaxed text-zinc-600">
                <p className="mb-1 font-semibold text-zinc-700">口径说明</p>
                <p>三边和 = 长 + 宽 + 高；材积重 = 长 × 宽 × 高 ÷ 6000；计费重 = max(实重, 材积重)。</p>
                <p>客户申报 = 客户上传数据；出给客户 = 系统建议值；供应商 = 供应商货箱清单实际数据。</p>
              </div>
              {ALARM_HELP.map((a) => (
                <div key={a.text} className="rounded-lg border border-zinc-200 p-4">
                  <div className="mb-2">
                    <span className={`inline-flex rounded-full px-2.5 py-0.5 text-xs font-semibold ${ALARM_TONE[a.tone]}`}>
                      {a.text}
                    </span>
                  </div>
                  <dl className="space-y-1.5 text-sm">
                    <div>
                      <dt className="inline font-medium text-zinc-700">表达的意思：</dt>
                      <dd className="inline text-zinc-600">{a.condition}</dd>
                    </div>
                    <div>
                      <dt className="inline font-medium text-zinc-700">举例：</dt>
                      <dd className="inline text-zinc-600">{a.example}</dd>
                    </div>
                    <div>
                      <dt className="inline font-medium text-zinc-700">含义说明：</dt>
                      <dd className="inline text-zinc-600">{a.meaning}</dd>
                    </div>
                  </dl>
                </div>
              ))}
            </div>
          </div>
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
  file,
  isDragging,
  dragHandlers,
  onClear,
  inputRef,
  onFileSelect,
}: {
  label: string;
  description: string;
  file: File | null;
  isDragging: boolean;
  dragHandlers: {
    onDragOver: (e: React.DragEvent) => void;
    onDragLeave: (e: React.DragEvent) => void;
    onDrop: (e: React.DragEvent) => void;
  };
  onClear: () => void;
  inputRef: React.RefObject<HTMLInputElement | null>;
  onFileSelect: (e: React.ChangeEvent<HTMLInputElement>) => void;
}) {
  return (
    <div
      {...dragHandlers}
      onClick={() => !file && inputRef.current?.click()}
      className={`
        flex cursor-pointer flex-col items-center justify-center rounded-xl border-2 border-dashed p-12 transition-colors
        ${
          file
            ? "border-zinc-300 bg-white"
            : isDragging
              ? "border-primary bg-primary/5"
              : "border-[#C7E0F0] bg-white/60 hover:border-primary hover:bg-primary/5"
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
            <p className="text-xs text-zinc-400">{(file.size / 1024).toFixed(1)} KB</p>
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
          <svg
            className="mb-4 h-12 w-12 text-zinc-400"
            fill="none"
            viewBox="0 0 24 24"
            stroke="currentColor"
          >
            <path
              strokeLinecap="round"
              strokeLinejoin="round"
              strokeWidth={1.5}
              d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12"
            />
          </svg>
          <p className="text-lg font-medium text-zinc-600">{label}</p>
          <p className="mt-1 text-sm text-zinc-400">{description}</p>
          <p className="mt-2 text-sm text-zinc-400">拖拽 .xlsx 文件到此处，或点击上传</p>
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
