"use client";

import { useState, useCallback, useRef } from "react";

// ============================================================
// Types
// ============================================================

interface GroupInfo {
  seq: string;
  forwarderId: string;
  customsNo: string;
  country: string;
  qty: string | number;
  truckFee: number;
  customsFee: number;
  portFee: number;
  oceanTaxUSD: number;
  hasDomestic: boolean;
  hasInternational: boolean;
}

interface DownloadUrls {
  domestic: string;
  international: string;
  invoice: string;
}

interface SplitResponse {
  sessionId: string;
  sourceFile: string;
  totalGroups: number;
  domesticCount: number;
  internationalCount: number;
  invoiceCount: number;
  groups: GroupInfo[];
  downloadUrls: DownloadUrls;
  error?: string;
}

// ============================================================
// Helpers
// ============================================================

function fmtMoney(n: number): string {
  return n.toLocaleString("zh-CN", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

// ============================================================
// Page
// ============================================================

export default function PipixiongSplitPage() {
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
      const res = await fetch("/api/pipixiong-split", { method: "POST", body: formData });
      const data: SplitResponse = await res.json();
      if (!res.ok) { setError(data.error || "拆分处理失败"); return; }
      setResult(data);
    } catch (err) {
      setError(`请求失败: ${err instanceof Error ? err.message : "网络错误"}`);
    } finally {
      setProcessing(false);
    }
  }, []);

  const onDragOver = useCallback((e: React.DragEvent) => { e.preventDefault(); setIsDragging(true); }, []);
  const onDragLeave = useCallback((e: React.DragEvent) => { e.preventDefault(); setIsDragging(false); }, []);
  const onDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault(); setIsDragging(false);
    const file = e.dataTransfer.files[0];
    if (file) handleFile(file);
  }, [handleFile]);
  const onFileChange = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) handleFile(file);
  }, [handleFile]);
  const reset = useCallback(() => { setResult(null); setError(null); }, []);

  return (
    <div className="space-y-8">
      {/* Header */}
      <div>
        <h1 className="text-2xl font-bold text-zinc-800">🧾 皮皮熊账单拆分</h1>
        <p className="mt-1 text-sm text-zinc-500">
          上传皮皮熊合并账单 Excel，自动拆分为国内账单、国外账单和 INVOICE
        </p>
      </div>

      {/* Upload Area */}
      {!result && (
        <div
          onDragOver={onDragOver} onDragLeave={onDragLeave} onDrop={onDrop}
          onClick={() => !processing && fileInputRef.current?.click()}
          className={`flex cursor-pointer flex-col items-center justify-center rounded-xl border-2 border-dashed p-12 transition-colors
            ${isDragging ? "border-violet-400 bg-violet-50" : "border-zinc-300 bg-zinc-50 hover:border-zinc-400 hover:bg-zinc-100"}
            ${processing ? "pointer-events-none opacity-60" : ""}`}
        >
          {processing ? (
            <>
              <div className="mb-4 h-10 w-10 animate-spin rounded-full border-4 border-zinc-300 border-t-violet-600" />
              <p className="text-lg font-medium text-zinc-600">正在处理中...</p>
              <p className="mt-1 text-sm text-zinc-400">读取账单、拆分国内/国外费用、生成 INVOICE</p>
            </>
          ) : (
            <>
              <svg className="mb-4 h-12 w-12 text-zinc-400" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.5}
                  d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12" />
              </svg>
              <p className="text-lg font-medium text-zinc-600">拖拽皮皮熊账单 Excel 到此处，或点击上传</p>
              <p className="mt-1 text-sm text-zinc-400">支持 .xlsx / .xls · 按报关单号拆分 · 生成三种账单</p>
            </>
          )}
          <input ref={fileInputRef} type="file" accept=".xlsx,.xls" onChange={onFileChange} className="hidden" />
        </div>
      )}

      {/* Error */}
      {error && (
        <div className="flex items-center gap-3 rounded-lg border border-red-300 bg-red-50 p-4">
          <svg className="h-5 w-5 shrink-0 text-red-500" fill="none" viewBox="0 0 24 24" stroke="currentColor">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 8v4m0 4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
          </svg>
          <p className="text-sm text-red-700">{error}</p>
          <button onClick={reset} className="ml-auto shrink-0 text-sm font-medium text-red-600 hover:text-red-800">重新上传</button>
        </div>
      )}

      {/* Results */}
      {result && (
        <>
          {/* Download buttons */}
          <div className="rounded-xl border border-zinc-200 bg-white p-6 shadow-sm">
            <div className="mb-4 flex flex-wrap items-center justify-between gap-3">
              <div>
                <p className="text-sm text-zinc-500">
                  源文件: <span className="font-medium text-zinc-700">{result.sourceFile}</span>
                </p>
                <p className="text-sm text-zinc-500">
                  {result.totalGroups} 组账单 · 国内 {result.domesticCount} · 国外 {result.internationalCount} · INVOICE {result.invoiceCount}
                </p>
              </div>
              <button onClick={reset}
                className="rounded-lg border border-zinc-300 px-4 py-2 text-sm font-medium text-zinc-600 hover:bg-zinc-50 transition-colors">
                重新拆分
              </button>
            </div>

            <div className="grid grid-cols-1 gap-3 sm:grid-cols-3">
              {/* 国内账单 */}
              <a href={result.downloadUrls.domestic} download
                className="flex items-center gap-3 rounded-lg border-2 border-red-200 bg-red-50 p-4 hover:bg-red-100 transition-colors">
                <span className="flex h-10 w-10 items-center justify-center rounded-lg bg-red-100 text-xl">🇨🇳</span>
                <div>
                  <p className="font-semibold text-red-700">国内账单</p>
                  <p className="text-xs text-red-500">深圳抬头 · RMB 税票</p>
                  <p className="text-xs text-red-400 mt-0.5">{result.domesticCount} 个 sheet</p>
                </div>
                <svg className="ml-auto h-5 w-5 text-red-400" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                </svg>
              </a>

              {/* 国外账单 */}
              <a href={result.downloadUrls.international} download
                className="flex items-center gap-3 rounded-lg border-2 border-blue-200 bg-blue-50 p-4 hover:bg-blue-100 transition-colors">
                <span className="flex h-10 w-10 items-center justify-center rounded-lg bg-blue-100 text-xl">🇭🇰</span>
                <div>
                  <p className="font-semibold text-blue-700">国外账单</p>
                  <p className="text-xs text-blue-500">香港抬头 · USD 费用</p>
                  <p className="text-xs text-blue-400 mt-0.5">{result.internationalCount} 个 sheet</p>
                </div>
                <svg className="ml-auto h-5 w-5 text-blue-400" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                </svg>
              </a>

              {/* INVOICE */}
              <a href={result.downloadUrls.invoice} download
                className="flex items-center gap-3 rounded-lg border-2 border-emerald-200 bg-emerald-50 p-4 hover:bg-emerald-100 transition-colors">
                <span className="flex h-10 w-10 items-center justify-center rounded-lg bg-emerald-100 text-xl">📋</span>
                <div>
                  <p className="font-semibold text-emerald-700">INVOICE</p>
                  <p className="text-xs text-emerald-500">香港抬头 · 客户发票</p>
                  <p className="text-xs text-emerald-400 mt-0.5">{result.invoiceCount} 个 sheet</p>
                </div>
                <svg className="ml-auto h-5 w-5 text-emerald-400" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                </svg>
              </a>
            </div>
          </div>

          {/* Detail table */}
          <div className="rounded-xl border border-zinc-200 bg-white overflow-hidden shadow-sm">
            <div className="px-5 py-3 border-b border-zinc-100 bg-zinc-50">
              <h3 className="font-semibold text-zinc-700">账单明细</h3>
            </div>
            <div className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead>
                  <tr className="border-b border-zinc-200 text-left text-xs font-semibold uppercase text-zinc-500">
                    <th className="px-4 py-2">#</th>
                    <th className="px-4 py-2">货代识别号</th>
                    <th className="px-4 py-2">报关单号</th>
                    <th className="px-4 py-2">国家</th>
                    <th className="px-4 py-2">数量</th>
                    <th className="px-4 py-2 text-right">拖车费</th>
                    <th className="px-4 py-2 text-right">报关费</th>
                    <th className="px-4 py-2 text-right">港杂费</th>
                    <th className="px-4 py-2 text-right">海运费(USD)</th>
                    <th className="px-4 py-2 text-center">国内</th>
                    <th className="px-4 py-2 text-center">国际</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-zinc-100">
                  {result.groups.map((g, i) => (
                    <tr key={i} className="hover:bg-zinc-50">
                      <td className="px-4 py-2 font-mono text-xs text-zinc-400">{g.seq}</td>
                      <td className="px-4 py-2 font-mono text-xs font-medium text-zinc-700">{g.forwarderId}</td>
                      <td className="px-4 py-2 font-mono text-xs text-zinc-500">
                        {g.customsNo === "/" ? <span className="text-amber-500">不报关</span> : g.customsNo}
                      </td>
                      <td className="px-4 py-2 text-xs">{g.country}</td>
                      <td className="px-4 py-2 text-xs">{g.qty}</td>
                      <td className="px-4 py-2 text-right font-mono text-xs">{g.truckFee > 0 ? `¥${fmtMoney(g.truckFee)}` : "-"}</td>
                      <td className="px-4 py-2 text-right font-mono text-xs">{g.customsFee > 0 ? `¥${fmtMoney(g.customsFee)}` : "-"}</td>
                      <td className="px-4 py-2 text-right font-mono text-xs">{g.portFee > 0 ? `¥${fmtMoney(g.portFee)}` : "-"}</td>
                      <td className="px-4 py-2 text-right font-mono text-xs font-medium">
                        {g.oceanTaxUSD > 0 ? `$${fmtMoney(g.oceanTaxUSD)}` : "-"}
                      </td>
                      <td className="px-4 py-2 text-center">{g.hasDomestic ? "✅" : "—"}</td>
                      <td className="px-4 py-2 text-center">{g.hasInternational ? "✅" : "—"}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>

          {/* Tip */}
          <div className="rounded-lg border border-zinc-200 bg-white p-4 text-sm text-zinc-500">
            <p className="font-medium text-zinc-700 mb-1">📌 说明</p>
            <ul className="list-disc pl-5 space-y-1">
              <li><strong>国内账单</strong>：深圳抬头，RMB 税票（拖车费+报关费+港杂费），含深圳银行账户信息</li>
              <li><strong>国外账单</strong>：香港抬头，USD 海运费+税金，含汇丰银行账户信息</li>
              <li><strong>INVOICE</strong>：香港抬头，客户发票格式，含 JOB NO / 船务信息 / Total / 银行信息</li>
              <li>报关单号为 &quot;/&quot; 的记录视为不报关，仅生成国外账单和 INVOICE</li>
              <li>默认只处理&quot;是否支付&quot; ≠ Y 的记录</li>
            </ul>
          </div>
        </>
      )}

      {/* Initial tip */}
      {!result && !error && (
        <div className="rounded-lg border border-zinc-200 bg-white p-4 text-sm text-zinc-500">
          <p>
            💡 提示：系统自动从 &quot;1.大货出货数据费用统计&quot; sheet 读取未支付记录，
            按报关单号拆分，保留参考数据 sheet，分别生成国内账单（深圳/RMB）、
            国外账单（香港/USD）和 INVOICE（香港/客户发票）三个 Excel 文件。
          </p>
        </div>
      )}
    </div>
  );
}
