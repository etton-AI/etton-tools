"use client";

import { useState, useCallback, useEffect, useRef } from "react";

// ============================================================
// Types
// ============================================================

interface SingleResult {
  sourceFile: string;
  fileName: string;
  supplier: string;
  customerOrderNo: string;
  warehouseCode: string;
  totalBoxes: number;
  dataRows: number;
  warnings: string[];
  downloadUrl: string;
}

interface BatchFileItem {
  sourceFile: string;
  fileName: string;
  country: string;
  warehouseCode: string;
  totalBoxes: number;
  dataRows: number;
}

interface BatchResult {
  zipFileName: string;
  supplier: string;
  country: string;
  files: BatchFileItem[];
  warnings: string[];
  downloadUrl: string;
}

interface ConvertResponse {
  sessionId: string;
  kind: "single" | "batch";
  result: SingleResult | BatchResult;
  error?: string;
}

const SUPPLIERS = [
  { key: "天图", label: "天图", ready: true },
  { key: "星链", label: "星链", ready: true },
  { key: "皓鹏", label: "皓鹏", ready: true },
  { key: "美琦", label: "美琦", ready: true },
  { key: "安邦", label: "安邦（待提供模板）", ready: false },
] as const;

// ============================================================
// Page
// ============================================================

export default function YxToTiantuPage() {
  const [isDragging, setIsDragging] = useState(false);
  const [processing, setProcessing] = useState(false);
  const [files, setFiles] = useState<File[]>([]);
  const [supplier, setSupplier] = useState<string>("天图");
  const [channels, setChannels] = useState<string[]>([]);
  const [channelsLoading, setChannelsLoading] = useState(false);
  const [channelName, setChannelName] = useState("");
  const [channelOpen, setChannelOpen] = useState(false);
  const [customerOrderNo, setCustomerOrderNo] = useState("");
  const [result, setResult] = useState<ConvertResponse | null>(null);
  const [error, setError] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const channelWrapRef = useRef<HTMLDivElement>(null);

  const isSupported = (name: string) => /\.(xlsx|xls)$/i.test(name);

  // 供应商变化时读取对应渠道下拉列表
  useEffect(() => {
    let cancelled = false;
    setChannelsLoading(true);
    setChannelName("");
    setError(null);

    fetch(`/api/yx-to-tiantu?supplier=${encodeURIComponent(supplier)}`)
      .then((r) => r.json())
      .then((data) => {
        if (cancelled) return;
        if (data.error) setError(data.error);
        else setChannels(data.channels || []);
      })
      .catch(() => {
        if (!cancelled) setError("渠道加载失败，请重试");
      })
      .finally(() => {
        if (!cancelled) setChannelsLoading(false);
      });

    return () => {
      cancelled = true;
    };
  }, [supplier]);

  // 点击渠道输入框外部时收起下拉
  useEffect(() => {
    const onClickOutside = (e: MouseEvent) => {
      if (channelWrapRef.current && !channelWrapRef.current.contains(e.target as Node)) {
        setChannelOpen(false);
      }
    };
    document.addEventListener("mousedown", onClickOutside);
    return () => document.removeEventListener("mousedown", onClickOutside);
  }, []);

  const selectChannel = useCallback((c: string) => {
    setChannelName(c);
    setChannelOpen(false);
  }, []);

  const convert = useCallback(async () => {
    if (files.length === 0) {
      setError("请先上传延讯易通标准模板");
      return;
    }
    if (!supplier) {
      setError("请选择供应商");
      return;
    }
    if (!channelName) {
      setError("请选择渠道");
      return;
    }

    setProcessing(true);
    setError(null);
    setResult(null);

    try {
      const formData = new FormData();
      for (const f of files) formData.append("files", f);
      formData.append("supplier", supplier);
      formData.append("channelName", channelName);
      if (files.length === 1) formData.append("customerOrderNo", customerOrderNo.trim());

      const res = await fetch("/api/yx-to-tiantu", {
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
  }, [files, supplier, channelName, customerOrderNo]);

  const addFiles = useCallback((incoming: File[]) => {
    const valid = incoming.filter((f) => isSupported(f.name));
    if (valid.length !== incoming.length) {
      setError("存在非 .xlsx/.xls 文件，已忽略");
    } else {
      setError(null);
    }
    if (valid.length === 0) return;

    setFiles((prev) => {
      const map = new Map(prev.map((f) => [`${f.name}:${f.size}`, f]));
      for (const f of valid) map.set(`${f.name}:${f.size}`, f);
      return Array.from(map.values());
    });
    setResult(null);
  }, []);

  const removeFile = useCallback((index: number) => {
    setFiles((prev) => prev.filter((_, i) => i !== index));
    setResult(null);
  }, []);

  const clearFiles = useCallback(() => {
    setFiles([]);
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
      addFiles(Array.from(e.dataTransfer.files));
    },
    [addFiles]
  );

  const onFileChange = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      addFiles(Array.from(e.target.files || []));
      e.target.value = "";
    },
    [addFiles]
  );

  const reset = useCallback(() => {
    setResult(null);
    setError(null);
    setFiles([]);
    setSupplier("天图");
    setChannelName("");
    setCustomerOrderNo("");
  }, []);

  const single = result?.kind === "single" ? (result.result as SingleResult) : null;
  const batch = result?.kind === "batch" ? (result.result as BatchResult) : null;

  const channelQuery = channelName.trim().toLowerCase();
  const filteredChannels = channels.filter((c) => c.toLowerCase().includes(channelQuery));

  return (
    <div className="space-y-8">
      {/* Header */}
      <div>
        <h1 className="text-2xl font-bold text-deep">🏷️ 延讯标准模板转供应商发票</h1>
        <p className="mt-1 text-sm text-zinc-500">
          批量上传延讯易通标准模板，选择供应商与渠道，自动输出对应供应商发票（单张或 ZIP）
        </p>
      </div>

      {/* Form */}
      {!result && (
        <div className="space-y-5">
          {/* Upload */}
          <div
            onDragOver={onDragOver}
            onDragLeave={onDragLeave}
            onDrop={onDrop}
            onClick={() => !processing && fileInputRef.current?.click()}
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
            <svg className="mb-4 h-12 w-12 text-zinc-400" fill="none" viewBox="0 0 24 24" stroke="currentColor">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.5} d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12" />
            </svg>
            <p className="text-lg font-medium text-zinc-600">
              拖拽多份延讯易通标准模板到此处，或点击选择
            </p>
            <p className="mt-1 text-sm text-zinc-400">
              支持 .xlsx 格式，可一次选择多份（同一国家 + 同一渠道）
            </p>
            <input
              ref={fileInputRef}
              type="file"
              accept=".xlsx,.xls"
              multiple
              onChange={onFileChange}
              className="hidden"
            />
          </div>

          {/* File list */}
          {files.length > 0 && (
            <div className="rounded-lg border border-zinc-200 bg-white p-4">
              <div className="mb-2 flex items-center justify-between">
                <p className="text-sm font-medium text-zinc-700">
                  已选 {files.length} 份发票
                </p>
                <button
                  onClick={clearFiles}
                  className="text-xs text-zinc-400 hover:text-red-500 transition-colors"
                >
                  清空
                </button>
              </div>
              <ul className="max-h-48 space-y-1.5 overflow-y-auto">
                {files.map((f, i) => (
                  <li
                    key={`${f.name}:${f.size}`}
                    className="flex items-center justify-between gap-3 rounded bg-zinc-50 px-3 py-1.5 text-sm"
                  >
                    <span className="truncate text-zinc-700">{f.name}</span>
                    <button
                      onClick={() => removeFile(i)}
                      className="shrink-0 text-zinc-400 hover:text-red-500 transition-colors"
                      title="移除"
                    >
                      ✕
                    </button>
                  </li>
                ))}
              </ul>
            </div>
          )}

          {/* Options */}
          <div className="grid gap-4 sm:grid-cols-3">
            <div>
              <label className="mb-1.5 block text-sm font-medium text-zinc-700">
                供应商 <span className="text-red-500">*</span>
              </label>
              <select
                value={supplier}
                onChange={(e) => setSupplier(e.target.value)}
                className="w-full rounded-lg border border-zinc-300 bg-white px-3 py-2 text-sm focus:border-primary focus:outline-none focus:ring-1 focus:ring-primary"
              >
                {SUPPLIERS.map((s) => (
                  <option key={s.key} value={s.key} disabled={!s.ready}>
                    {s.label}
                  </option>
                ))}
              </select>
            </div>
            <div ref={channelWrapRef} className="relative">
              <label className="mb-1.5 block text-sm font-medium text-zinc-700">
                渠道 <span className="text-red-500">*</span>
              </label>
              <input
                type="text"
                value={channelName}
                onChange={(e) => {
                  setChannelName(e.target.value);
                  setChannelOpen(true);
                }}
                onFocus={() => setChannelOpen(true)}
                placeholder={channelsLoading ? "渠道加载中..." : "输入关键字筛选渠道（如：美转加 / 19日）"}
                disabled={channelsLoading}
                className="w-full rounded-lg border border-zinc-300 bg-white px-3 py-2 text-sm focus:border-primary focus:outline-none focus:ring-1 focus:ring-primary disabled:opacity-60"
              />
              {channelOpen && !channelsLoading && filteredChannels.length > 0 && (
                <ul className="absolute left-0 right-0 top-full z-20 mt-1 max-h-56 overflow-y-auto rounded-lg border border-zinc-200 bg-white py-1 shadow-lg">
                  {filteredChannels.map((c) => (
                    <li key={c}>
                      <button
                        type="button"
                        onMouseDown={(e) => {
                          e.preventDefault();
                          selectChannel(c);
                        }}
                        className="block w-full px-3 py-1.5 text-left text-sm text-zinc-700 hover:bg-primary/10"
                      >
                        {c}
                      </button>
                    </li>
                  ))}
                </ul>
              )}
              {channelOpen && !channelsLoading && filteredChannels.length === 0 && (
                <div className="absolute left-0 right-0 top-full z-20 mt-1 rounded-lg border border-zinc-200 bg-white px-3 py-2 text-sm text-zinc-400 shadow-lg">
                  无匹配渠道
                </div>
              )}
            </div>
            {files.length === 1 && (
              <div>
                <label className="mb-1.5 block text-sm font-medium text-zinc-700">
                  客户订单号
                </label>
                <input
                  type="text"
                  value={customerOrderNo}
                  onChange={(e) => setCustomerOrderNo(e.target.value)}
                  placeholder="留空则按 FBA 编号命名"
                  className="w-full rounded-lg border border-zinc-300 bg-white px-3 py-2 text-sm focus:border-primary focus:outline-none focus:ring-1 focus:ring-primary"
                />
              </div>
            )}
          </div>

          {/* Convert button */}
          <button
            onClick={convert}
            disabled={processing}
            className="inline-flex items-center gap-2 rounded-lg bg-primary px-5 py-2.5 text-sm font-medium text-white shadow-sm hover:bg-primary-dark transition-colors disabled:pointer-events-none disabled:opacity-60"
          >
            {processing ? (
              <>
                <div className="h-4 w-4 animate-spin rounded-full border-2 border-white/40 border-t-white" />
                正在转换...
              </>
            ) : (
              "开始转换"
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

      {/* Result — single */}
      {single && (
        <>
          <div className="flex flex-wrap items-center justify-between gap-4 rounded-lg border border-zinc-200 bg-white p-4 shadow-sm">
            <div>
              <p className="text-sm text-zinc-600">
                转换完成：<span className="font-semibold text-emerald-700">{single.fileName}</span>
              </p>
              <p className="mt-0.5 text-xs text-zinc-400">
                {single.supplier} · 仓库 {single.warehouseCode} · 总箱数 {single.totalBoxes} · 数据 {single.dataRows} 行
              </p>
            </div>
            <div className="flex gap-3">
              <a
                href={single.downloadUrl}
                className="inline-flex items-center gap-2 rounded-lg bg-primary px-4 py-2 text-sm font-medium text-white shadow-sm hover:bg-primary-dark transition-colors"
              >
                <svg className="h-4 w-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                </svg>
                下载
              </a>
              <button
                onClick={reset}
                className="inline-flex items-center gap-2 rounded-lg border border-zinc-300 bg-white px-4 py-2 text-sm font-medium text-zinc-600 hover:bg-zinc-50 transition-colors"
              >
                重新转换
              </button>
            </div>
          </div>

          {single.warnings?.length > 0 && (
            <div className="flex flex-wrap gap-1.5">
              {single.warnings.map((w, i) => (
                <span key={i} className="inline-flex items-center gap-1 rounded bg-amber-50 px-2 py-0.5 text-xs font-medium text-amber-700">
                  ⚠ {w}
                </span>
              ))}
            </div>
          )}
        </>
      )}

      {/* Result — batch */}
      {batch && (
        <>
          <div className="flex flex-wrap items-center justify-between gap-4 rounded-lg border border-zinc-200 bg-white p-4 shadow-sm">
            <div>
              <p className="text-sm text-zinc-600">
                转换完成：<span className="font-semibold text-emerald-700">{batch.zipFileName}</span>
              </p>
              <p className="mt-0.5 text-xs text-zinc-400">
                {batch.supplier} · 国家「{batch.country}」 · 共 {batch.files.length} 份供应商发票
              </p>
            </div>
            <div className="flex gap-3">
              <a
                href={batch.downloadUrl}
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
                重新转换
              </button>
            </div>
          </div>

          {batch.files.length > 0 && (
            <div className="rounded-lg border border-zinc-200 bg-white p-4">
              <p className="mb-2 text-sm font-medium text-zinc-700">生成的发票清单</p>
              <div className="max-h-64 overflow-y-auto">
                <table className="w-full text-sm">
                  <thead>
                    <tr className="border-b border-zinc-100 text-left text-xs text-zinc-400">
                      <th className="py-1.5 pr-3 font-medium">源发票</th>
                      <th className="py-1.5 pr-3 font-medium">供应商发票</th>
                      <th className="py-1.5 pr-3 font-medium">仓库</th>
                      <th className="py-1.5 pr-3 font-medium">箱数</th>
                      <th className="py-1.5 font-medium">数据行</th>
                    </tr>
                  </thead>
                  <tbody>
                    {batch.files.map((f, i) => (
                      <tr key={i} className="border-b border-zinc-50 last:border-0">
                        <td className="py-1.5 pr-3 text-zinc-700">{f.sourceFile}</td>
                        <td className="py-1.5 pr-3 text-zinc-600">{f.fileName}</td>
                        <td className="py-1.5 pr-3 font-mono text-zinc-600">{f.warehouseCode}</td>
                        <td className="py-1.5 pr-3 text-zinc-600">{f.totalBoxes}</td>
                        <td className="py-1.5 text-zinc-600">{f.dataRows}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {batch.warnings?.length > 0 && (
            <div className="flex flex-wrap gap-1.5">
              {batch.warnings.map((w, i) => (
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
            💡 提示：支持天图 / 星链 / 皓鹏 / 美琦 四家供应商（安邦待提供模板）。渠道名按供应商从模板自动解析，可在输入框输入关键字（如「美转加」「19日」）实时筛选匹配项；
            批量上传时须保证所有发票为同一国家、同一渠道，客户订单号可留空（按 FBA 编号命名）。
            收件地址由模板内 VLOOKUP 公式根据仓库代码自动带出（皓鹏为静态示例值，需人工核对）。
          </p>
        </div>
      )}
    </div>
  );
}
