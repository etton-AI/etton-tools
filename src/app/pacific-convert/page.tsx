"use client";

import { useState, useCallback, useRef } from "react";

// ============================================================
// Types
// ============================================================

interface IntervalSummary {
  name: string;
  fileName: string;
  rowCount: number;
}

interface ConvertResponse {
  sessionId: string;
  sourceFile: string;
  totalRows: number;
  skippedRows: number;
  intervals: IntervalSummary[];
  downloads: {
    allZip: string;
    files: string[];
  };
  error?: string;
}

// ============================================================
// Color palette
// ============================================================

const INTERVAL_COLORS = [
  { bg: "bg-blue-50", border: "border-blue-300", text: "text-blue-700", badge: "bg-blue-100 text-blue-700" },
  { bg: "bg-emerald-50", border: "border-emerald-300", text: "text-emerald-700", badge: "bg-emerald-100 text-emerald-700" },
  { bg: "bg-amber-50", border: "border-amber-300", text: "text-amber-700", badge: "bg-amber-100 text-amber-700" },
  { bg: "bg-orange-50", border: "border-orange-300", text: "text-orange-700", badge: "bg-orange-100 text-orange-700" },
  { bg: "bg-rose-50", border: "border-rose-300", text: "text-rose-700", badge: "bg-rose-100 text-rose-700" },
  { bg: "bg-violet-50", border: "border-violet-300", text: "text-violet-700", badge: "bg-violet-100 text-violet-700" },
  { bg: "bg-cyan-50", border: "border-cyan-300", text: "text-cyan-700", badge: "bg-cyan-100 text-cyan-700" },
  { bg: "bg-lime-50", border: "border-lime-300", text: "text-lime-700", badge: "bg-lime-100 text-lime-700" },
  { bg: "bg-pink-50", border: "border-pink-300", text: "text-pink-700", badge: "bg-pink-100 text-pink-700" },
  { bg: "bg-teal-50", border: "border-teal-300", text: "text-teal-700", badge: "bg-teal-100 text-teal-700" },
];

function getIntervalColor(idx: number) {
  return INTERVAL_COLORS[idx % INTERVAL_COLORS.length];
}


// ============================================================
// Page
// ============================================================

export default function PacificConvertPage() {
  const [rates, setRates] = useState({ USD: 7, EUR: 8, GBP: 9, JPY: 0.04 });
  const [isDragging, setIsDragging] = useState(false);
  const [processing, setProcessing] = useState(false);
  const [result, setResult] = useState<ConvertResponse | null>(null);
  const [error, setError] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFile = useCallback(
    async (file: File) => {
      if (!file.name.endsWith(".xlsx") && !file.name.endsWith(".xls")) {
        setError("只支持 .xlsx 或 .xls 格式的 Excel 文件");
        return;
      }

      setProcessing(true);
      setError(null);
      setResult(null);

      try {
        const formData = new FormData();
        formData.append("file", file);
        formData.append("rateUSD", String(rates.USD));
        formData.append("rateEUR", String(rates.EUR));
        formData.append("rateGBP", String(rates.GBP));
        formData.append("rateJPY", String(rates.JPY));

        const res = await fetch("/api/convert-pacific", {
          method: "POST",
          body: formData,
        });

        const data: ConvertResponse = await res.json();

        if (!res.ok) {
          setError(data.error || "转换失败");
          return;
        }

        setResult(data);
      } catch {
        setError("网络错误，请重试");
      } finally {
        setProcessing(false);
      }
    },
    [rates]
  );

  const onDragOver = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  }, []);

  const onDragLeave = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
  }, []);

  const onDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault();
      setIsDragging(false);
      const file = e.dataTransfer.files[0];
      if (file) handleFile(file);
    },
    [handleFile]
  );

  const onFileChange = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      const file = e.target.files?.[0];
      if (file) handleFile(file);
    },
    [handleFile]
  );

  const reset = useCallback(() => {
    setResult(null);
    setError(null);
  }, []);

  // Use <a> tags directly — avoids browser async-click popup blocker for downloads

  return (
    <div className="space-y-8">
      {/* Header */}
      <div>
        <h1 className="text-2xl font-bold text-deep">
          🚢 太平洋货箱清单转换
        </h1>
        <p className="mt-1 text-sm text-zinc-500">
          上传易通货箱清单 Excel，自动转换为太平洋投保清单并按单箱货值(RMB)拆分
        </p>
      </div>

      {/* Exchange Rate Inputs */}
      <div className="rounded-lg border border-zinc-200 bg-white p-4">
        <p className="text-xs font-semibold text-zinc-500 uppercase tracking-wide mb-3">
          外币汇率设置
        </p>
        <div className="flex flex-wrap gap-4">
          <RateInput
            label="USD 美元"
            value={rates.USD}
            onChange={(v) => setRates((r) => ({ ...r, USD: v }))}
          />
          <RateInput
            label="EUR 欧元"
            value={rates.EUR}
            onChange={(v) => setRates((r) => ({ ...r, EUR: v }))}
          />
          <RateInput
            label="GBP 英镑"
            value={rates.GBP}
            onChange={(v) => setRates((r) => ({ ...r, GBP: v }))}
          />
          <RateInput
            label="JPY 日元"
            value={rates.JPY}
            onChange={(v) => setRates((r) => ({ ...r, JPY: v }))}
          />
        </div>
        <p className="mt-2 text-xs text-zinc-400">
          用于计算每箱人民币货值：总申报货值 ÷ 总箱数 × 汇率。未列出的币种默认按 1 计算。
        </p>
      </div>

      {/* Upload Area */}
      {!result && (
        <div
          onDragOver={onDragOver}
          onDragLeave={onDragLeave}
          onDrop={onDrop}
          onClick={() => !processing && fileInputRef.current?.click()}
          className={`
            flex cursor-pointer flex-col items-center justify-center rounded-xl border-2 border-dashed p-12 transition-colors
            ${
              isDragging
                ? "border-primary bg-primary/5"
                : "border-[#C7E0F0] bg-white/60 hover:border-primary hover:bg-primary/5"
            }
            ${processing ? "pointer-events-none opacity-60" : ""}
          `}
        >
          {processing ? (
            <>
              <div className="mb-4 h-10 w-10 animate-spin rounded-full border-4 border-zinc-300 border-t-primary" />
              <p className="text-lg font-medium text-zinc-600">正在处理中...</p>
              <p className="mt-1 text-sm text-zinc-400">读取数据、计算每箱货值、生成拆分文件</p>
            </>
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
              <p className="text-lg font-medium text-zinc-600">
                拖拽货箱清单 Excel 到此处，或点击上传
              </p>
              <p className="mt-1 text-sm text-zinc-400">
                支持 .xlsx / .xls 格式 · 自动按单箱货值区间拆分
              </p>
            </>
          )}
          <input
            ref={fileInputRef}
            type="file"
            accept=".xlsx,.xls"
            onChange={onFileChange}
            className="hidden"
          />
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
                源文件:{" "}
                <span className="font-medium text-zinc-700">{result.sourceFile}</span>
              </p>
              <p className="text-sm text-zinc-500">
                共 {result.totalRows} 行{" "}
                {result.skippedRows > 0 && `(跳过 ${result.skippedRows} 行无效数据) `}·{" "}
                {result.intervals.filter((i) => i.rowCount > 0).length} 个非空区间
              </p>
              <p className="text-xs text-zinc-400 mt-0.5">
                当前汇率: USD={rates.USD} EUR={rates.EUR} GBP={rates.GBP} JPY={rates.JPY}
              </p>
            </div>
            <div className="flex gap-3">
              <a
                href={result.downloads.allZip}
                className="inline-flex items-center gap-2 rounded-lg bg-primary px-4 py-2 text-sm font-medium text-white shadow-sm hover:bg-primary-dark transition-colors"
              >
                <svg className="h-4 w-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                </svg>
                下载全部 (ZIP)
              </a>
              <button
                onClick={reset}
                className="inline-flex items-center gap-2 rounded-lg border border-zinc-300 bg-white px-4 py-2 text-sm font-medium text-zinc-600 hover:bg-zinc-50 transition-colors"
              >
                重新上传
              </button>
            </div>
          </div>

          {/* Interval cards */}
          <div className="space-y-4">
            {result.intervals.map((iv, idx) => {
              const colors = getIntervalColor(idx);
              const isEmpty = iv.rowCount === 0;

              return (
                <div
                  key={iv.name}
                  className={`rounded-xl border-2 ${colors.border} ${colors.bg} ${
                    isEmpty ? "opacity-50" : ""
                  } overflow-hidden`}
                >
                  {/* Card header */}
                  <div className="flex flex-wrap items-center justify-between gap-3 px-5 py-3">
                    <div className="flex items-center gap-3">
                      <span
                        className={`inline-flex items-center rounded-full px-2.5 py-0.5 text-xs font-semibold ${colors.badge}`}
                      >
                        区间 {idx + 1}
                      </span>
                      <h2 className={`text-base font-bold ${colors.text}`}>
                        {iv.name}
                      </h2>
                      {isEmpty ? (
                        <span className="text-xs text-zinc-400">（无数据）</span>
                      ) : (
                        <span className="text-sm text-zinc-600">
                          {iv.rowCount} 行
                        </span>
                      )}
                    </div>
                    <div className="flex items-center gap-2">
                      {!isEmpty && (
                        <a
                          href={result.downloads.files[idx]}
                          className="inline-flex items-center gap-1.5 rounded-md bg-white/80 px-3 py-1.5 text-sm font-medium text-zinc-700 shadow-sm hover:bg-white transition-colors"
                        >
                          <svg className="h-4 w-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                          </svg>
                          下载
                        </a>
                      )}
                    </div>
                  </div>
                </div>
              );
            })}
          </div>
        </>
      )}

      {/* Tip */}
      {!result && (
        <div className="rounded-lg border border-zinc-200 bg-white p-4 text-sm text-zinc-500">
          <p>
            💡 提示：上传文件后将自动按单箱人民币货值拆分为多个区间文件。
            提单号、船名航次、柜号等字段需手动补充。
          </p>
        </div>
      )}
    </div>
  );
}

// ============================================================
// Mini components
// ============================================================

function RateInput({
  label,
  value,
  onChange,
}: {
  label: string;
  value: number;
  onChange: (v: number) => void;
}) {
  return (
    <div className="flex items-center gap-2">
      <label className="text-sm font-medium text-zinc-600">{label}</label>
      <input
        type="number"
        step="0.1"
        min="0.1"
        value={value}
        onChange={(e) => {
          const v = parseFloat(e.target.value);
          if (!isNaN(v) && v > 0) onChange(v);
        }}
        className="w-20 rounded-md border border-zinc-300 px-2.5 py-1.5 text-sm text-zinc-700 focus:border-primary focus:outline-none focus:ring-1 focus:ring-primary"
      />
    </div>
  );
}
