"use client";

import { useState, useCallback, useEffect, useRef } from "react";

// ============================================================
// Types
// ============================================================

interface ResultFile {
  sourceFile: string;
  outputFile: string;
  warehouseCode: string;
  totalBoxes: number;
  dataRows: number;
  warnings: string[];
}

interface ResultItem {
  zipFileName: string;
  mode: string;
  supplier: string;
  channel: string;
  total: number;
  files: ResultFile[];
  downloadUrl: string;
}

interface ConvertResponse {
  sessionId: string;
  result: ResultItem;
  error?: string;
}

/** 入口2 的每文件 SO(客户订单号) */
interface FileItem {
  file: File;
  so: string;
}

const SUPPLIERS = [
  { key: "凯鑫", label: "凯鑫", ready: true },
  { key: "英美", label: "英美", ready: true },
  { key: "星链", label: "星链", ready: true },
  { key: "航乐", label: "航乐", ready: true },
] as const;

const isSupported = (name: string) => /\.(xlsx|xls)$/i.test(name);

// ============================================================
// 公共小部件
// ============================================================

function ErrorBox({ error }: { error: string | null }) {
  if (!error) return null;
  return (
    <div className="flex items-center gap-3 rounded-lg border border-red-300 bg-red-50 p-4">
      <svg className="h-5 w-5 shrink-0 text-red-500" fill="none" viewBox="0 0 24 24" stroke="currentColor">
        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 8v4m0 4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
      </svg>
      <p className="text-sm text-red-700">{error}</p>
    </div>
  );
}

function DownloadBar({ result, onReset, subtitle }: { result: ResultItem; onReset: () => void; subtitle: string }) {
  return (
    <div className="flex flex-wrap items-center justify-between gap-4 rounded-lg border border-zinc-200 bg-white p-4 shadow-sm">
      <div>
        <p className="text-sm text-zinc-600">
          转换完成：<span className="font-semibold text-emerald-700">{result.zipFileName}</span>
        </p>
        <p className="mt-0.5 text-xs text-zinc-400">{subtitle}</p>
      </div>
      <div className="flex gap-3">
        <a
          href={result.downloadUrl}
          className="inline-flex items-center gap-2 rounded-lg bg-primary px-4 py-2 text-sm font-medium text-white shadow-sm hover:bg-primary-dark transition-colors"
        >
          <svg className="h-4 w-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
          </svg>
          下载 ZIP
        </a>
        <button
          onClick={onReset}
          className="inline-flex items-center gap-2 rounded-lg border border-zinc-300 bg-white px-4 py-2 text-sm font-medium text-zinc-600 hover:bg-zinc-50 transition-colors"
        >
          重新转换
        </button>
      </div>
    </div>
  );
}

// ============================================================
// 入口1: 易通系统上传
// ============================================================

function EttongEntry() {
  const [isDragging, setIsDragging] = useState(false);
  const [processing, setProcessing] = useState(false);
  const [files, setFiles] = useState<File[]>([]);
  const [result, setResult] = useState<ConvertResponse | null>(null);
  const [error, setError] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const addFiles = useCallback((incoming: File[]) => {
    const valid = incoming.filter((f) => isSupported(f.name));
    if (valid.length !== incoming.length) setError("存在非 .xlsx 文件，已忽略");
    else setError(null);
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

  const convert = useCallback(async () => {
    if (files.length === 0) {
      setError("请先上传 TR 客户发票文件");
      return;
    }
    setProcessing(true);
    setError(null);
    setResult(null);
    try {
      const formData = new FormData();
      for (const f of files) formData.append("file", f);
      formData.append("mode", "ettong");

      const res = await fetch("/api/tr-to-supplier", { method: "POST", body: formData });
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
  }, [files]);

  const reset = useCallback(() => {
    setResult(null);
    setError(null);
    setFiles([]);
  }, []);

  const onDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  };
  const onDragLeave = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
  };
  const onDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    addFiles(Array.from(e.dataTransfer.files));
  };
  const onFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    addFiles(Array.from(e.target.files || []));
    e.target.value = "";
  };

  return (
    <div className="space-y-5">
      {!result && (
        <>
          <div
            onDragOver={onDragOver}
            onDragLeave={onDragLeave}
            onDrop={onDrop}
            onClick={() => !processing && fileInputRef.current?.click()}
            className={`flex cursor-pointer flex-col items-center justify-center rounded-xl border-2 border-dashed p-8 transition-colors ${
              isDragging ? "border-primary bg-primary/5" : "border-[#C7E0F0] bg-white/60 hover:border-primary hover:bg-primary/5"
            } ${processing ? "pointer-events-none opacity-60" : ""}`}
          >
            <svg className="mb-4 h-12 w-12 text-zinc-400" fill="none" viewBox="0 0 24 24" stroke="currentColor">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.5} d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12" />
            </svg>
            <p className="text-lg font-medium text-zinc-600">拖拽多份 TR 客户发票到此处，或点击选择</p>
            <p className="mt-1 text-sm text-zinc-400">支持 .xlsx 格式，可一次选择多份</p>
            <input ref={fileInputRef} type="file" accept=".xlsx,.xls" multiple onChange={onFileChange} className="hidden" />
          </div>

          {files.length > 0 && (
            <div className="rounded-lg border border-zinc-200 bg-white p-4">
              <div className="mb-2 flex items-center justify-between">
                <p className="text-sm font-medium text-zinc-700">已选 {files.length} 份发票</p>
                <button onClick={clearFiles} className="text-xs text-zinc-400 hover:text-red-500 transition-colors">清空</button>
              </div>
              <ul className="max-h-48 space-y-1.5 overflow-y-auto">
                {files.map((f, i) => (
                  <li key={`${f.name}:${f.size}`} className="flex items-center justify-between gap-3 rounded bg-zinc-50 px-3 py-1.5 text-sm">
                    <span className="truncate text-zinc-700">{f.name}</span>
                    <button onClick={() => removeFile(i)} className="shrink-0 text-zinc-400 hover:text-red-500 transition-colors" title="移除">✕</button>
                  </li>
                ))}
              </ul>
            </div>
          )}

          <button
            onClick={convert}
            disabled={processing}
            className="inline-flex items-center gap-2 rounded-lg bg-primary px-5 py-2.5 text-sm font-medium text-white shadow-sm hover:bg-primary-dark transition-colors disabled:pointer-events-none disabled:opacity-60"
          >
            {processing ? (
              <>
                <div className="h-4 w-4 animate-spin rounded-full border-2 border-white/40 border-t-white" />
                正在生成...
              </>
            ) : (
              "生成易通系统上传发票"
            )}
          </button>

          <div className="rounded-lg border border-zinc-200 bg-white p-4 text-sm text-zinc-500">
            <p>💡 每份发票在原第 8 行插入空行（第 8 行及以下全部下移一行），对齐易通上传模板的偏移；保留源文件图片、合并单元格与格式。</p>
          </div>
        </>
      )}

      <ErrorBox error={error} />

      {result && (
        <>
          <DownloadBar
            result={result.result}
            onReset={reset}
            subtitle={`共 ${result.result.total} 份易通系统上传发票`}
          />
          {result.result.files.length > 0 && (
            <div className="rounded-lg border border-zinc-200 bg-white p-4">
              <p className="mb-2 text-sm font-medium text-zinc-700">生成的发票清单</p>
              <div className="max-h-64 overflow-y-auto">
                <table className="w-full text-sm">
                  <thead>
                    <tr className="border-b border-zinc-100 text-left text-xs text-zinc-400">
                      <th className="py-1.5 pr-3 font-medium">源发票</th>
                      <th className="py-1.5 font-medium">易通系统上传</th>
                    </tr>
                  </thead>
                  <tbody>
                    {result.result.files.map((f, i) => (
                      <tr key={i} className="border-b border-zinc-50 last:border-0">
                        <td className="py-1.5 pr-3 text-zinc-700">{f.sourceFile}</td>
                        <td className="py-1.5 text-zinc-600">{f.outputFile}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}
        </>
      )}
    </div>
  );
}

// ============================================================
// 入口2: 转供应商发票
// ============================================================

function SupplierEntry() {
  const [isDragging, setIsDragging] = useState(false);
  const [processing, setProcessing] = useState(false);
  const [fileItems, setFileItems] = useState<FileItem[]>([]);
  const [supplier, setSupplier] = useState<string>("凯鑫");
  const [channels, setChannels] = useState<string[]>([]);
  const [channelsLoading, setChannelsLoading] = useState(false);
  const [channelName, setChannelName] = useState("");
  const [channelOpen, setChannelOpen] = useState(false);
  const [customsOptions, setCustomsOptions] = useState<string[]>([]);
  const [stations, setStations] = useState<string[]>([]);
  const [customsOverride, setCustomsOverride] = useState(""); // "" = 自动按发票判断
  const [station, setStation] = useState(""); // "" = 模板默认
  const [result, setResult] = useState<ConvertResponse | null>(null);
  const [error, setError] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const channelWrapRef = useRef<HTMLDivElement>(null);

  // 供应商变化时读取对应渠道下拉
  useEffect(() => {
    let cancelled = false;
    setChannelsLoading(true);
    setChannelName("");
    setError(null);
    fetch(`/api/tr-to-supplier/channels?supplier=${encodeURIComponent(supplier)}`)
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

  // 供应商变化时读取报关方式选项 + 送货货站列表
  useEffect(() => {
    let cancelled = false;
    setCustomsOptions([]);
    setStations([]);
    setCustomsOverride("");
    setStation("");
    fetch(`/api/tr-to-supplier/options?supplier=${encodeURIComponent(supplier)}`)
      .then((r) => r.json())
      .then((data) => {
        if (cancelled) return;
        if (data.error) setError(data.error);
        else {
          setCustomsOptions(data.customs || []);
          setStations(data.stations || []);
        }
      })
      .catch(() => {
        if (!cancelled) setError("选项加载失败，请重试");
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

  const addFiles = useCallback((incoming: File[]) => {
    const valid = incoming.filter((f) => isSupported(f.name));
    if (valid.length !== incoming.length) setError("存在非 .xlsx 文件，已忽略");
    else setError(null);
    if (valid.length === 0) return;
    setFileItems((prev) => {
      const map = new Map(prev.map((it) => [`${it.file.name}:${it.file.size}`, it]));
      for (const f of valid) {
        const key = `${f.name}:${f.size}`;
        if (!map.has(key)) map.set(key, { file: f, so: "" });
      }
      return Array.from(map.values());
    });
    setResult(null);
  }, []);

  const removeFile = useCallback((index: number) => {
    setFileItems((prev) => prev.filter((_, i) => i !== index));
    setResult(null);
  }, []);

  const setSo = useCallback((index: number, value: string) => {
    setFileItems((prev) => prev.map((it, i) => (i === index ? { ...it, so: value } : it)));
  }, []);

  const clearFiles = useCallback(() => {
    setFileItems([]);
    setResult(null);
  }, []);

  const convert = useCallback(async () => {
    if (fileItems.length === 0) {
      setError("请先上传 TR 客户发票文件");
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
      for (const it of fileItems) formData.append("file", it.file);
      for (const it of fileItems) formData.append("customerOrderNo", it.so);
      formData.append("mode", "supplier");
      formData.append("supplier", supplier);
      formData.append("channelName", channelName);
      if (customsOverride) formData.append("customsOverride", customsOverride);
      if (station) formData.append("station", station);

      const res = await fetch("/api/tr-to-supplier", { method: "POST", body: formData });
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
  }, [fileItems, supplier, channelName, customsOverride, station]);

  const reset = useCallback(() => {
    setResult(null);
    setError(null);
    setFileItems([]);
    setSupplier("凯鑫");
    setChannelName("");
    setCustomsOverride("");
    setStation("");
  }, []);

  const onDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  };
  const onDragLeave = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
  };
  const onDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    addFiles(Array.from(e.dataTransfer.files));
  };
  const onFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    addFiles(Array.from(e.target.files || []));
    e.target.value = "";
  };

  const channelQuery = channelName.trim().toLowerCase();
  const filteredChannels = channels.filter((c) => c.toLowerCase().includes(channelQuery));

  return (
    <div className="space-y-5">
      {!result && (
        <>
          <div
            onDragOver={onDragOver}
            onDragLeave={onDragLeave}
            onDrop={onDrop}
            onClick={() => !processing && fileInputRef.current?.click()}
            className={`flex cursor-pointer flex-col items-center justify-center rounded-xl border-2 border-dashed p-8 transition-colors ${
              isDragging ? "border-primary bg-primary/5" : "border-[#C7E0F0] bg-white/60 hover:border-primary hover:bg-primary/5"
            } ${processing ? "pointer-events-none opacity-60" : ""}`}
          >
            <svg className="mb-4 h-12 w-12 text-zinc-400" fill="none" viewBox="0 0 24 24" stroke="currentColor">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.5} d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12" />
            </svg>
            <p className="text-lg font-medium text-zinc-600">拖拽多份 TR 客户发票到此处，或点击选择</p>
            <p className="mt-1 text-sm text-zinc-400">支持 .xlsx 格式，可一次选择多份（每份可单独填写 SO）</p>
            <input ref={fileInputRef} type="file" accept=".xlsx,.xls" multiple onChange={onFileChange} className="hidden" />
          </div>

          {fileItems.length > 0 && (
            <div className="rounded-lg border border-zinc-200 bg-white p-4">
              <div className="mb-2 flex items-center justify-between">
                <p className="text-sm font-medium text-zinc-700">已选 {fileItems.length} 份发票（每份填写对应 SO）</p>
                <button onClick={clearFiles} className="text-xs text-zinc-400 hover:text-red-500 transition-colors">清空</button>
              </div>
              <ul className="max-h-64 space-y-2 overflow-y-auto">
                {fileItems.map((it, i) => (
                  <li key={`${it.file.name}:${it.file.size}`} className="rounded bg-zinc-50 px-3 py-2">
                    <div className="flex items-center justify-between gap-3">
                      <span className="truncate text-sm text-zinc-700">{it.file.name}</span>
                      <button onClick={() => removeFile(i)} className="shrink-0 text-zinc-400 hover:text-red-500 transition-colors" title="移除">✕</button>
                    </div>
                    <input
                      type="text"
                      value={it.so}
                      onChange={(e) => setSo(i, e.target.value)}
                      placeholder="SO 号（留空按文件名命名）"
                      className="mt-1.5 w-full rounded-md border border-zinc-300 bg-white px-2.5 py-1.5 text-sm focus:border-primary focus:outline-none focus:ring-1 focus:ring-primary"
                    />
                  </li>
                ))}
              </ul>
            </div>
          )}

          {/* Options */}
          <div className="grid gap-4 sm:grid-cols-2 lg:grid-cols-4">
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
                placeholder={channelsLoading ? "渠道加载中..." : "输入关键字筛选渠道"}
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
            <div>
              <label className="mb-1.5 block text-sm font-medium text-zinc-700">报关方式</label>
              <select
                value={customsOverride}
                onChange={(e) => setCustomsOverride(e.target.value)}
                className="w-full rounded-lg border border-zinc-300 bg-white px-3 py-2 text-sm focus:border-primary focus:outline-none focus:ring-1 focus:ring-primary"
              >
                <option value="">自动（按发票判断）</option>
                {customsOptions.map((c) => (
                  <option key={c} value={c}>{c}</option>
                ))}
              </select>
            </div>
            {stations.length > 0 && (
              <div>
                <label className="mb-1.5 block text-sm font-medium text-zinc-700">
                  送货货站
                </label>
                <select
                  value={station}
                  onChange={(e) => setStation(e.target.value)}
                  className="w-full rounded-lg border border-zinc-300 bg-white px-3 py-2 text-sm focus:border-primary focus:outline-none focus:ring-1 focus:ring-primary"
                >
                  <option value="">自动（模板默认）</option>
                  {stations.map((s) => (
                    <option key={s} value={s}>{s}</option>
                  ))}
                </select>
              </div>
            )}
          </div>

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
              "生成供应商发票"
            )}
          </button>

          <div className="rounded-lg border border-zinc-200 bg-white p-4 text-sm text-zinc-500">
            <p>
              💡 系统自动解析 TR 发票的带电/带磁、报关方式、清关方式、VAT/EORI、总箱数、国家、仓库代码并按供应商模板逐列映射。
              报关方式默认按发票 B8 自动判断，可手动覆盖；送货货站仅英美（所在货站）与航乐（交货仓库）可改，凯鑫/星链无此字段。
              凯鑫/航乐会自动嵌入产品图片（WPS DISPIMG），收件地址由模板内 VLOOKUP 公式自动带出。
            </p>
          </div>
        </>
      )}

      <ErrorBox error={error} />

      {result && (
        <>
          <DownloadBar
            result={result.result}
            onReset={reset}
            subtitle={`${result.result.supplier} · 渠道「${result.result.channel}」 · 共 ${result.result.total} 份供应商发票`}
          />
          {result.result.files.length > 0 && (
            <div className="rounded-lg border border-zinc-200 bg-white p-4">
              <p className="mb-2 text-sm font-medium text-zinc-700">生成的发票清单</p>
              <div className="max-h-64 overflow-y-auto">
                <table className="w-full text-sm">
                  <thead>
                    <tr className="border-b border-zinc-100 text-left text-xs text-zinc-400">
                      <th className="py-1.5 pr-3 font-medium">源发票</th>
                      <th className="py-1.5 pr-3 font-medium">供应商发票</th>
                      <th className="py-1.5 pr-3 font-medium">仓库</th>
                      <th className="py-1.5 font-medium">箱数</th>
                    </tr>
                  </thead>
                  <tbody>
                    {result.result.files.map((f, i) => (
                      <tr key={i} className="border-b border-zinc-50 last:border-0">
                        <td className="py-1.5 pr-3 text-zinc-700">{f.sourceFile}</td>
                        <td className="py-1.5 pr-3 text-zinc-600">{f.outputFile}</td>
                        <td className="py-1.5 pr-3 font-mono text-zinc-600">{f.warehouseCode}</td>
                        <td className="py-1.5 text-zinc-600">{f.totalBoxes}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}
        </>
      )}
    </div>
  );
}

// ============================================================
// Page
// ============================================================

export default function TrToSupplierPage() {
  const [mode, setMode] = useState<"ettong" | "supplier">("supplier");

  const tabBase = "flex-1 rounded-lg px-4 py-2 text-sm font-medium transition-colors";
  const tabActive = "bg-primary text-white shadow-sm";
  const tabInactive = "bg-zinc-100 text-zinc-600 hover:bg-zinc-200";

  return (
    <div className="space-y-8">
      {/* Header */}
      <div>
        <h1 className="text-2xl font-bold text-deep">🏷️ TR 客户发票转换</h1>
        <p className="mt-1 text-sm text-zinc-500">
          批量上传 TR 客户发票，生成易通系统上传发票或转供应商发票（ZIP）
        </p>
      </div>

      {/* Tab switch */}
      <div className="flex gap-2 rounded-xl border border-zinc-200 bg-white p-1.5">
        <button
          onClick={() => setMode("ettong")}
          className={`${tabBase} ${mode === "ettong" ? tabActive : tabInactive}`}
        >
          易通系统上传
        </button>
        <button
          onClick={() => setMode("supplier")}
          className={`${tabBase} ${mode === "supplier" ? tabActive : tabInactive}`}
        >
          转供应商发票
        </button>
      </div>

      {mode === "ettong" ? <EttongEntry /> : <SupplierEntry />}
    </div>
  );
}
