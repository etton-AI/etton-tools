"use client";

import { useState, useCallback, useRef } from "react";

// ============================================================
// Types
// ============================================================

interface OverseasAddress {
  name: string;
  company: string;
  address: string;
  phone: string;
  city: string;
  state: string;
  zip: string;
}

interface YanxunTopInfo {
  transportMode: string;
  customsRaw: string;
  customsMapped: string;
  hasBattery: string;
  country: string;
  warehouseCode: string;
  warehouseType: string;
  channel: string;
  fbaId: string;
  transferOrderNo: string;
  overseas: OverseasAddress;
  referenceId: string;
  totalBoxes: number;
}

interface ResultItem {
  sourceFile: string;
  fileName: string;
  fbaId: string;
  totalBoxes: number;
  dataRows: number;
  mixedBoxGroups: number;
  topInfo: YanxunTopInfo;
  warnings: string[];
  downloadUrl: string;
}

interface FailedItem {
  sourceFile: string;
  error: string;
}

interface ConvertResponse {
  sessionId: string;
  summary: { total: number; success: number; failed: number };
  results: ResultItem[];
  failed: FailedItem[];
  downloadAllUrl: string;
  error?: string;
}

// ============================================================
// Page
// ============================================================

export default function YanxunConvertPage() {
  const [isDragging, setIsDragging] = useState(false);
  const [processing, setProcessing] = useState(false);
  const [fileNames, setFileNames] = useState<string[]>([]);
  const [result, setResult] = useState<ConvertResponse | null>(null);
  const [error, setError] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const isSupported = (name: string) => /\.(xlsx|xls|zip)$/i.test(name);

  const convert = useCallback(async (fileList: File[]) => {
    setProcessing(true);
    setError(null);
    setResult(null);
    setFileNames(fileList.map((f) => f.name));

    try {
      const formData = new FormData();
      fileList.forEach((f) => formData.append("files", f));

      const res = await fetch("/api/yanxun-convert", {
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
  }, []);

  const handleFiles = useCallback(
    (incoming: File[]) => {
      const valid = incoming.filter((f) => isSupported(f.name));
      if (valid.length === 0) {
        setError("请选择 .xlsx / .xls / .zip 文件");
        return;
      }
      void convert(valid);
    },
    [convert]
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
      const files = Array.from(e.dataTransfer.files);
      if (files.length) handleFiles(files);
    },
    [handleFiles]
  );

  const onFileChange = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      const files = Array.from(e.target.files || []);
      if (files.length) handleFiles(files);
      e.target.value = "";
    },
    [handleFiles]
  );

  const reset = useCallback(() => {
    setResult(null);
    setError(null);
    setFileNames([]);
  }, []);

  return (
    <div className="space-y-8">
      {/* Header */}
      <div>
        <h1 className="text-2xl font-bold text-deep">
          📄 延讯下单优化
        </h1>
        <p className="mt-1 text-sm text-zinc-500">
          批量上传延讯下单发票（多选或 ZIP 包），自动转换为易通下单模版并按 ETTON_FBA号 / ETTON_调拨单号 命名
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
              <p className="text-lg font-medium text-zinc-600">
                正在处理 {fileNames.length} 个文件...
              </p>
              <p className="mt-1 max-w-md truncate text-sm text-zinc-400">
                {fileNames.join("、")}
              </p>
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
                拖拽多个延讯发票到此处，或点击选择
              </p>
              <p className="mt-1 text-sm text-zinc-400">
                支持批量多选 .xlsx / .xls，或上传打包好的 .zip 文件
              </p>
            </>
          )}
          <input
            ref={fileInputRef}
            type="file"
            accept=".xlsx,.xls,.zip"
            multiple
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
          {/* Summary + actions */}
          <div className="flex flex-wrap items-center justify-between gap-4 rounded-lg border border-zinc-200 bg-white p-4 shadow-sm">
            <div>
              <p className="text-sm text-zinc-600">
                转换完成：{" "}
                <span className="font-semibold text-emerald-700">{result.summary.success} 票成功</span>
                {result.summary.failed > 0 && (
                  <span className="font-semibold text-red-600"> · {result.summary.failed} 票失败</span>
                )}
              </p>
              <p className="mt-0.5 text-xs text-zinc-400">
                共 {result.summary.total} 个文件 · 输出文件按 ETTON_FBA号 / ETTON_调拨单号 命名
              </p>
            </div>
            <div className="flex gap-3">
              <a
                href={result.downloadAllUrl}
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

          {/* Success list */}
          {result.results.length > 0 && (
            <div className="overflow-hidden rounded-lg border border-zinc-200 bg-white">
              <div className="border-b border-zinc-200 bg-zinc-50 px-4 py-2 text-xs font-semibold uppercase tracking-wide text-zinc-500">
                成功转换（{result.results.length}）
              </div>
              <ul className="divide-y divide-zinc-100">
                {result.results.map((r) => (
                  <li
                    key={r.fileName}
                    className="flex flex-wrap items-center justify-between gap-3 px-4 py-3"
                  >
                    <div className="min-w-0">
                      <p className="font-medium text-zinc-800">
                        {r.fileName}
                        {(r.topInfo.warehouseType === "FBA" ? r.fbaId : r.topInfo.transferOrderNo) && (
                          <span className="ml-2 text-xs text-emerald-600">
                            {r.topInfo.warehouseType === "FBA" ? r.fbaId : r.topInfo.transferOrderNo}
                          </span>
                        )}
                        <span className="ml-2 text-xs text-zinc-400">{r.sourceFile}</span>
                      </p>
                      <p className="mt-0.5 text-xs text-zinc-500">
                        总箱数 {r.totalBoxes} · 数据 {r.dataRows} 行
                        {r.mixedBoxGroups > 0 && (
                          <span className="text-amber-600"> · 混箱 {r.mixedBoxGroups} 组</span>
                        )}
                        <span className="text-zinc-400">
                          {" · "}
                          {r.topInfo.transportMode} / {r.topInfo.customsMapped} / {r.topInfo.warehouseType} / {r.topInfo.country}{" "}
                          {r.topInfo.warehouseType === "FBA"
                            ? r.topInfo.warehouseCode
                            : `${r.topInfo.overseas?.city ?? ""} ${r.topInfo.overseas?.state ?? ""} ${r.topInfo.overseas?.zip ?? ""}`}
                        </span>
                      </p>
                      {r.warnings?.length > 0 && (
                        <p className="mt-1 flex flex-wrap gap-1.5">
                          {r.warnings.map((w, i) => (
                            <span
                              key={i}
                              className="inline-flex items-center gap-1 rounded bg-amber-50 px-2 py-0.5 text-xs font-medium text-amber-700"
                            >
                              ⚠ {w}
                            </span>
                          ))}
                        </p>
                      )}
                    </div>
                    <a
                      href={r.downloadUrl}
                      className="inline-flex shrink-0 items-center gap-1.5 rounded-md bg-emerald-50 px-3 py-1.5 text-sm font-medium text-emerald-700 hover:bg-emerald-100 transition-colors"
                    >
                      <svg className="h-4 w-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                      </svg>
                      下载
                    </a>
                  </li>
                ))}
              </ul>
            </div>
          )}

          {/* Failed list */}
          {result.failed.length > 0 && (
            <div className="overflow-hidden rounded-lg border border-red-200 bg-white">
              <div className="border-b border-red-200 bg-red-50 px-4 py-2 text-xs font-semibold uppercase tracking-wide text-red-500">
                转换失败（{result.failed.length}）
              </div>
              <ul className="divide-y divide-red-50">
                {result.failed.map((f, idx) => (
                  <li key={idx} className="flex flex-wrap items-start justify-between gap-3 px-4 py-3">
                    <div className="min-w-0">
                      <p className="text-sm font-medium text-zinc-700">{f.sourceFile}</p>
                      <p className="mt-0.5 text-xs text-red-600">{f.error}</p>
                    </div>
                  </li>
                ))}
              </ul>
            </div>
          )}
        </>
      )}

      {/* Tip */}
      {!result && (
        <div className="rounded-lg border border-zinc-200 bg-white p-4 text-sm text-zinc-500">
          <p>
            💡 提示：可一次拖入多个延讯发票，或上传打包好的 ZIP 包，系统会逐个转换并打包下载。
            自动完成：运输方式→业务类型、正式报关→报关方式（公司自报/否=普通报关，永德吉报关=报关退税）、
            目的地→收件人国家/仓库代码（FBA 场景）或「私人地址/海外仓」收件人信息（海外仓场景）、
            渠道→备注、FBA号/调拨单号→文件名（ETTON_FBA号 / ETTON_调拨单号）。混箱时仅首行保留重量尺寸。
          </p>
        </div>
      )}
    </div>
  );
}
