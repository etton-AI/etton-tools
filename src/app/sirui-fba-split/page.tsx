"use client";

import { useState, useCallback, useRef } from "react";

// ============================================================
// Types
// ============================================================

interface ResultFile {
  fba: string;
  warehouse: string;
  fileName: string;
  rowCount: number;
}

interface SplitResponse {
  sessionId: string;
  result: {
    zipFileName: string;
    total: number;
    files: ResultFile[];
    warnings: string[];
    downloadUrl: string;
  };
  error?: string;
}

// ============================================================
// Page
// ============================================================

export default function SiruiFbaSplitPage() {
  const [isDragging, setIsDragging] = useState(false);
  const [processing, setProcessing] = useState(false);
  const [quoteFile, setQuoteFile] = useState<File | null>(null);
  const [invoiceFile, setInvoiceFile] = useState<File | null>(null);
  const [result, setResult] = useState<SplitResponse | null>(null);
  const [error, setError] = useState<string | null>(null);
  const invoiceInputRef = useRef<HTMLInputElement>(null);
  const quoteInputRef = useRef<HTMLInputElement>(null);

  const isSupported = (name: string) => /\.(xlsx|xls)$/i.test(name);

  const convert = useCallback(async () => {
    if (!quoteFile) {
      setError("请先上传内部报价表");
      return;
    }
    if (!invoiceFile) {
      setError("请先上传下单发票");
      return;
    }

    setProcessing(true);
    setError(null);
    setResult(null);

    try {
      const formData = new FormData();
      formData.append("quote", quoteFile);
      formData.append("file", invoiceFile);

      const res = await fetch("/api/sirui-fba-split", {
        method: "POST",
        body: formData,
      });

      const data: SplitResponse = await res.json();

      if (!res.ok) {
        setError(data.error || "拆分失败");
        return;
      }

      setResult(data);
    } catch {
      setError("网络错误，请重试");
    } finally {
      setProcessing(false);
    }
  }, [quoteFile, invoiceFile]);

  const handleInvoice = useCallback((incoming: File) => {
    if (!isSupported(incoming.name)) {
      setError("请选择 .xlsx 格式的下单发票");
      return;
    }
    setInvoiceFile(incoming);
    setError(null);
    setResult(null);
  }, []);

  const handleQuote = useCallback((incoming: File) => {
    if (!isSupported(incoming.name)) {
      setError("请选择 .xlsx 格式的内部报价表");
      return;
    }
    setQuoteFile(incoming);
    setError(null);
    setResult(null);
  }, []);

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
      const files = Array.from(e.dataTransfer.files);
      if (files.length) handleInvoice(files[0]);
    },
    [handleInvoice]
  );

  const onInvoiceChange = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      const files = Array.from(e.target.files || []);
      if (files.length) handleInvoice(files[0]);
      e.target.value = "";
    },
    [handleInvoice]
  );

  const onQuoteChange = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      const files = Array.from(e.target.files || []);
      if (files.length) handleQuote(files[0]);
      e.target.value = "";
    },
    [handleQuote]
  );

  const reset = useCallback(() => {
    setResult(null);
    setError(null);
    setInvoiceFile(null);
    setQuoteFile(null);
  }, []);

  return (
    <div className="space-y-8">
      {/* Header */}
      <div>
        <h1 className="text-2xl font-bold text-deep">🧾 思锐FBA下单发票拆分</h1>
        <p className="mt-1 text-sm text-zinc-500">
          上传内部报价表 + 下单发票，按 FBA 码自动拆分为多张发票（每仓一张）
        </p>
      </div>

      {/* Form */}
      {!result && (
        <div className="space-y-5">
          {/* 下单发票上传 */}
          <div
            onDragOver={onDragOver}
            onDragLeave={onDragLeave}
            onDrop={onDrop}
            onClick={() => !processing && invoiceInputRef.current?.click()}
            className={`
              flex cursor-pointer flex-col items-center justify-center rounded-xl border-2 border-dashed p-8 transition-colors
              ${
                isDragging
                  ? "border-primary bg-primary/5"
                  : "border-[#C7E0F0] bg-white/60 hover:border-primary hover:bg-primary/5"
              }
              ${processing ? "pointer-events-none opacity-60" : ""}
            `}
          >
            {invoiceFile ? (
              <>
                <svg className="mb-3 h-10 w-10 text-emerald-500" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.5} d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
                </svg>
                <p className="max-w-md truncate text-lg font-medium text-zinc-700">{invoiceFile.name}</p>
                <p className="mt-1 text-sm text-zinc-400">点击可重新选择下单发票</p>
              </>
            ) : (
              <>
                <svg className="mb-4 h-12 w-12 text-zinc-400" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.5} d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12" />
                </svg>
                <p className="text-lg font-medium text-zinc-600">拖拽下单发票到此处，或点击选择</p>
                <p className="mt-1 text-sm text-zinc-400">支持 .xlsx 格式</p>
              </>
            )}
            <input
              ref={invoiceInputRef}
              type="file"
              accept=".xlsx,.xls"
              onChange={onInvoiceChange}
              className="hidden"
            />
          </div>

          {/* 内部报价表上传 */}
          <div
            onClick={() => !processing && quoteInputRef.current?.click()}
            className={`
              flex cursor-pointer items-center justify-between rounded-xl border-2 border-dashed p-4 transition-colors
              ${
                quoteFile
                  ? "border-emerald-300 bg-emerald-50/60"
                  : "border-[#C7E0F0] bg-white/60 hover:border-primary hover:bg-primary/5"
              }
              ${processing ? "pointer-events-none opacity-60" : ""}
            `}
          >
            <div className="flex items-center gap-3">
              <svg className="h-8 w-8 text-zinc-400" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.5} d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
              </svg>
              <div>
                <p className="font-medium text-zinc-700">
                  {quoteFile ? quoteFile.name : "内部报价表"}
                  <span className="ml-1 text-red-500">*</span>
                </p>
                <p className="text-xs text-zinc-400">提供「内部」工作表里的 FBA 码 / 仓点 / 渠道映射</p>
              </div>
            </div>
            {quoteFile && (
              <span className="inline-flex items-center gap-1 rounded bg-emerald-100 px-2 py-0.5 text-xs font-medium text-emerald-700">
                ✓ 已选择
              </span>
            )}
            <input
              ref={quoteInputRef}
              type="file"
              accept=".xlsx,.xls"
              onChange={onQuoteChange}
              className="hidden"
            />
          </div>

          {/* Split button */}
          <button
            onClick={convert}
            disabled={processing}
            className="inline-flex items-center gap-2 rounded-lg bg-primary px-5 py-2.5 text-sm font-medium text-white shadow-sm hover:bg-primary-dark transition-colors disabled:pointer-events-none disabled:opacity-60"
          >
            {processing ? (
              <>
                <div className="h-4 w-4 animate-spin rounded-full border-2 border-white/40 border-t-white" />
                正在拆分...
              </>
            ) : (
              "开始拆分"
            )}
          </button>
        </div>
      )}

      {/* Error */}
      {error && (
        <div className="flex items-center gap-3 rounded-lg border border-red-300 bg-red-50 p-4">
          <svg className="h-5 w-5 shrink-0 text-red-500" fill="none" viewBox="0 0 24 24" stroke="currentColor">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 8v4m0 4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
          </svg>
          <p className="text-sm text-red-700">{error}</p>
        </div>
      )}

      {/* Result */}
      {result && (
        <>
          <div className="flex flex-wrap items-center justify-between gap-4 rounded-lg border border-zinc-200 bg-white p-4 shadow-sm">
            <div>
              <p className="text-sm text-zinc-600">
                拆分完成：<span className="font-semibold text-emerald-700">{result.result.zipFileName}</span>
              </p>
              <p className="mt-0.5 text-xs text-zinc-400">共 {result.result.total} 张发票</p>
            </div>
            <div className="flex gap-3">
              <a
                href={result.result.downloadUrl}
                className="inline-flex items-center gap-2 rounded-lg bg-primary px-4 py-2 text-sm font-medium text-white shadow-sm hover:bg-primary-dark transition-colors"
              >
                <svg className="h-4 w-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                </svg>
                下载 ZIP
              </a>
              <button
                onClick={reset}
                className="inline-flex items-center gap-2 rounded-lg border border-zinc-300 bg-white px-4 py-2 text-sm font-medium text-zinc-600 hover:bg-zinc-50 transition-colors"
              >
                重新拆分
              </button>
            </div>
          </div>

          {result.result.files.length > 0 && (
            <div className="rounded-lg border border-zinc-200 bg-white p-4">
              <p className="mb-2 text-sm font-medium text-zinc-700">生成的发票清单</p>
              <div className="max-h-64 overflow-y-auto">
                <table className="w-full text-sm">
                  <thead>
                    <tr className="border-b border-zinc-100 text-left text-xs text-zinc-400">
                      <th className="py-1.5 pr-3 font-medium">FBA 码</th>
                      <th className="py-1.5 pr-3 font-medium">仓点</th>
                      <th className="py-1.5 pr-3 font-medium">行数</th>
                      <th className="py-1.5 font-medium">文件名</th>
                    </tr>
                  </thead>
                  <tbody>
                    {result.result.files.map((f, i) => (
                      <tr key={i} className="border-b border-zinc-50 last:border-0">
                        <td className="py-1.5 pr-3 font-mono text-zinc-700">{f.fba}</td>
                        <td className="py-1.5 pr-3 text-zinc-600">{f.warehouse}</td>
                        <td className="py-1.5 pr-3 text-zinc-600">{f.rowCount}</td>
                        <td className="py-1.5 text-zinc-500">{f.fileName}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {result.result.warnings?.length > 0 && (
            <div className="flex flex-wrap gap-1.5">
              {result.result.warnings.map((w, i) => (
                <span key={i} className="inline-flex items-center gap-1 rounded bg-amber-50 px-2 py-0.5 text-xs font-medium text-amber-700">
                  ⚠ {w}
                </span>
              ))}
            </div>
          )}
        </>
      )}

      {/* Tip */}
      {!result && (
        <div className="rounded-lg border border-zinc-200 bg-white p-4 text-sm text-zinc-500">
          <p>
            💡 提示：系统读取内部报价表「内部」工作表的 FBA 码（D 列）、仓点（C 列）、渠道（H 列），
            把一张下单发票按 FBA 码拆分。每张发票自动填收件人国家（美国）、仓库代码（仓点）、备注（渠道）、
            申报币种（USD），并在 W 列标注「FBA-仓点」，同时保留产品图片。
          </p>
        </div>
      )}
    </div>
  );
}
