"use client";

import { useState, useCallback, useRef } from "react";

// ============================================================
// Types
// ============================================================

interface BoxGroup {
  rows: number[];
  boxes: number;
  totalPrice: number;
  currency: string;
  perBoxOrig: number;
  perBoxRMB: number;
}

interface IntervalSummary {
  name: string;
  fileName: string;
  totalBoxes: number;
  totalWeight: number;
  totalVolume: number;
  groupCount: number;
  groups: BoxGroup[];
}

interface SplitResponse {
  sessionId: string;
  sourceFile: string;
  totalBoxes: number;
  totalGroups: number;
  intervals: IntervalSummary[];
  downloads: {
    allZip: string;
    files: string[];
  };
  error?: string;
}

// ============================================================
// Helpers
// ============================================================

const INTERVAL_COLORS = [
  { bg: "bg-blue-50", border: "border-blue-300", text: "text-blue-700", badge: "bg-blue-100 text-blue-700", dot: "bg-blue-500" },
  { bg: "bg-emerald-50", border: "border-emerald-300", text: "text-emerald-700", badge: "bg-emerald-100 text-emerald-700", dot: "bg-emerald-500" },
  { bg: "bg-amber-50", border: "border-amber-300", text: "text-amber-700", badge: "bg-amber-100 text-amber-700", dot: "bg-amber-500" },
  { bg: "bg-orange-50", border: "border-orange-300", text: "text-orange-700", badge: "bg-orange-100 text-orange-700", dot: "bg-orange-500" },
  { bg: "bg-rose-50", border: "border-rose-300", text: "text-rose-700", badge: "bg-rose-100 text-rose-700", dot: "bg-rose-500" },
];

function fmt(n: number, d = 2): string {
  return n.toLocaleString("zh-CN", {
    minimumFractionDigits: d,
    maximumFractionDigits: d,
  });
}

function fmtRMB(n: number): string {
  return `¥${n.toLocaleString("zh-CN", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
}

// ============================================================
// Page
// ============================================================

export default function InsuranceSplitPage() {
  const [isDragging, setIsDragging] = useState(false);
  const [processing, setProcessing] = useState(false);
  const [result, setResult] = useState<SplitResponse | null>(null);
  const [error, setError] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFile = useCallback(async (file: File) => {
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

      const res = await fetch("/api/split-insurance", {
        method: "POST",
        body: formData,
      });

      const data: SplitResponse = await res.json();

      if (!res.ok) {
        setError(data.error || "处理失败");
        return;
      }

      setResult(data);
    } catch {
      setError("网络错误，请重试");
    } finally {
      setProcessing(false);
    }
  }, []);

  // Drag & drop handlers
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

  return (
    <div className="space-y-8">
      {/* Header */}
      <div>
        <h1 className="text-2xl font-bold text-deep">
          📦 保单投保区间拆分
        </h1>
        <p className="mt-1 text-sm text-zinc-500">
          上传易通下单 Excel 模板，自动按每箱申报价值(RMB)拆分为 5 个投保区间文件
        </p>
      </div>

      {/* Upload Area */}
      {!result && (
        <div
          onDragOver={onDragOver}
          onDragLeave={onDragLeave}
          onDrop={onDrop}
          onClick={() => fileInputRef.current?.click()}
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
              <p className="mt-1 text-sm text-zinc-400">读取文件、分组计算、生成拆分文件</p>
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
                拖拽 Excel 文件到此处，或点击上传
              </p>
              <p className="mt-1 text-sm text-zinc-400">
                支持 .xlsx / .xls 格式 · 自动保留原模板格式
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
            onClick={() => {
              setError(null);
              setResult(null);
            }}
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
                源文件: <span className="font-medium text-zinc-700">{result.sourceFile}</span>
              </p>
              <p className="text-sm text-zinc-500">
                总 {result.totalBoxes} 箱 · {result.totalGroups} 个箱组 ·{" "}
                {result.intervals.filter((i) => i.totalBoxes > 0).length} 个非空区间
              </p>
            </div>
            <div className="flex gap-3">
              <a
                href={result.downloads.allZip}
                className="inline-flex items-center gap-2 rounded-lg bg-primary px-4 py-2 text-sm font-medium text-white shadow-sm hover:bg-primary-dark transition-colors"
                download
              >
                <svg className="h-4 w-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                </svg>
                下载全部 (ZIP)
              </a>
              <button
                onClick={() => {
                  setResult(null);
                  setError(null);
                }}
                className="inline-flex items-center gap-2 rounded-lg border border-zinc-300 bg-white px-4 py-2 text-sm font-medium text-zinc-600 hover:bg-zinc-50 transition-colors"
              >
                重新上传
              </button>
            </div>
          </div>

          {/* Interval cards */}
          <div className="space-y-5">
            {result.intervals.map((iv, idx) => {
              const colors = INTERVAL_COLORS[idx];
              const isEmpty = iv.totalBoxes === 0;

              return (
                <div
                  key={iv.name}
                  className={`rounded-xl border-2 ${colors.border} ${colors.bg} ${
                    isEmpty ? "opacity-55" : ""
                  } overflow-hidden`}
                >
                  {/* Card header */}
                  <div className="flex flex-wrap items-center justify-between gap-3 px-5 py-4">
                    <div className="flex items-center gap-3">
                      <span
                        className={`inline-flex items-center rounded-full px-2.5 py-0.5 text-xs font-semibold ${colors.badge}`}
                      >
                        区间 {idx + 1}
                      </span>
                      <h2 className={`text-lg font-bold ${colors.text}`}>
                        {iv.name}
                      </h2>
                      {isEmpty && (
                        <span className="text-xs text-zinc-400">（无数据）</span>
                      )}
                    </div>
                    <div className="flex flex-wrap items-center gap-2">
                      {!isEmpty && (
                        <a
                          href={result.downloads.files[idx]}
                          download
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

                  {/* Stats row */}
                  <div className="flex flex-wrap gap-4 px-5 pb-3 text-sm">
                    <MiniStat label="箱数" value={`${iv.totalBoxes} 箱`} highlight />
                    <MiniStat label="重量" value={`${fmt(iv.totalWeight)} KG`} />
                    <MiniStat label="体积" value={`${fmt(iv.totalVolume, 4)} CBM`} />
                    <MiniStat label="箱组" value={`${iv.groupCount} 组`} />
                  </div>

                  {/* Groups table */}
                  {iv.groups.length > 0 && (
                    <div className="border-t border-white/50 px-5 py-4">
                      <div className="overflow-x-auto">
                        <table className="w-full text-sm">
                          <thead>
                            <tr className="border-b border-zinc-300/70 text-left text-xs font-semibold uppercase tracking-wider text-zinc-500">
                              <th className="pb-2 pr-3">源数据行</th>
                              <th className="pb-2 pr-3">箱数</th>
                              <th className="pb-2 pr-3">总价(原币)</th>
                              <th className="pb-2 pr-3">币种</th>
                              <th className="pb-2 pr-3">每箱原币</th>
                              <th className="pb-2">每箱RMB</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-zinc-200/70">
                            {iv.groups.map((g, gi) => (
                              <tr key={gi} className="text-zinc-700">
                                <td className="py-2 pr-3 font-mono text-xs">
                                  {g.rows.length === 1
                                    ? `Row ${g.rows[0]}`
                                    : g.rows.join(", ") + " (混箱)"}
                                </td>
                                <td className="py-2 pr-3 font-semibold">{g.boxes}</td>
                                <td className="py-2 pr-3">{fmt(g.totalPrice)}</td>
                                <td className="py-2 pr-3">{g.currency}</td>
                                <td className="py-2 pr-3">{fmt(g.perBoxOrig)}</td>
                                <td className="py-2 font-semibold text-zinc-900">
                                  {fmtRMB(g.perBoxRMB)}
                                </td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        </>
      )}
    </div>
  );
}

function MiniStat({
  label,
  value,
  highlight = false,
}: {
  label: string;
  value: string;
  highlight?: boolean;
}) {
  return (
    <span className="inline-flex items-baseline gap-1">
      <span className="text-zinc-400">{label}:</span>
      <span className={highlight ? "font-bold text-zinc-800" : "text-zinc-600"}>
        {value}
      </span>
    </span>
  );
}
