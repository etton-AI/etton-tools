"use client";

import { useState, useCallback, useRef } from "react";

// ============================================================
// Types
// ============================================================

interface ResultItem {
  sourceFile: string;
  fileName: string;
  customerOrderNo: string;
  warehouseCode: string;
  totalBoxes: number;
  dataRows: number;
  warnings: string[];
  downloadUrl: string;
}

interface ConvertResponse {
  sessionId: string;
  result: ResultItem;
  error?: string;
}

const CHANNELS = [
  "加拿大超快线-卡派",
  "加拿大快线卡派-包税（ZIM/合德统配）",
  "加拿大快线UPS派包税（ZIM/合德统配）",
  "加拿大普船包税-海派",
  "美转加30日达卡派包税",
  "美转加28日达卡派包税",
  "易E达卡派包税",
  "易速达卡派包税",
];

// ============================================================
// Page
// ============================================================

export default function YxToTiantuPage() {
  const [isDragging, setIsDragging] = useState(false);
  const [processing, setProcessing] = useState(false);
  const [file, setFile] = useState<File | null>(null);
  const [customerOrderNo, setCustomerOrderNo] = useState("");
  const [channelName, setChannelName] = useState("");
  const [result, setResult] = useState<ConvertResponse | null>(null);
  const [error, setError] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const isSupported = (name: string) => /\.(xlsx|xls)$/i.test(name);

  const convert = useCallback(async () => {
    if (!file) {
      setError("请先上传延讯易通标准模板");
      return;
    }
    if (!customerOrderNo.trim()) {
      setError("请填写客户订单号");
      return;
    }
    if (!channelName.trim()) {
      setError("请填写服务渠道");
      return;
    }

    setProcessing(true);
    setError(null);
    setResult(null);

    try {
      const formData = new FormData();
      formData.append("file", file);
      formData.append("customerOrderNo", customerOrderNo.trim());
      formData.append("channelName", channelName.trim());

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
  }, [file, customerOrderNo, channelName]);

  const handleFile = useCallback((incoming: File) => {
    if (!isSupported(incoming.name)) {
      setError("请选择 .xlsx 格式的易通标准模板");
      return;
    }
    setFile(incoming);
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
      if (files.length) handleFile(files[0]);
    },
    [handleFile]
  );

  const onFileChange = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      const files = Array.from(e.target.files || []);
      if (files.length) handleFile(files[0]);
      e.target.value = "";
    },
    [handleFile]
  );

  const reset = useCallback(() => {
    setResult(null);
    setError(null);
    setFile(null);
    setCustomerOrderNo("");
    setChannelName("");
  }, []);

  return (
    <div className="space-y-8">
      {/* Header */}
      <div>
        <h1 className="text-2xl font-bold text-deep">🏷️ 延讯标准模板转天图发票</h1>
        <p className="mt-1 text-sm text-zinc-500">
          上传延讯易通标准模板，填写客户订单号与服务渠道，自动转换为天图供应商发票
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
            {file ? (
              <>
                <svg className="mb-3 h-10 w-10 text-emerald-500" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.5} d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
                </svg>
                <p className="max-w-md truncate text-lg font-medium text-zinc-700">{file.name}</p>
                <p className="mt-1 text-sm text-zinc-400">点击可重新选择易通标准模板</p>
              </>
            ) : (
              <>
                <svg className="mb-4 h-12 w-12 text-zinc-400" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.5} d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12" />
                </svg>
                <p className="text-lg font-medium text-zinc-600">拖拽延讯易通标准模板到此处，或点击选择</p>
                <p className="mt-1 text-sm text-zinc-400">支持 .xlsx 格式</p>
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

          {/* Options */}
          <div className="grid gap-4 sm:grid-cols-2">
            <div>
              <label className="mb-1.5 block text-sm font-medium text-zinc-700">
                客户订单号 <span className="text-red-500">*</span>
              </label>
              <input
                type="text"
                value={customerOrderNo}
                onChange={(e) => setCustomerOrderNo(e.target.value)}
                placeholder="如 YXMY26090028"
                className="w-full rounded-lg border border-zinc-300 bg-white px-3 py-2 text-sm focus:border-primary focus:outline-none focus:ring-1 focus:ring-primary"
              />
            </div>
            <div>
              <label className="mb-1.5 block text-sm font-medium text-zinc-700">
                服务渠道 <span className="text-red-500">*</span>
              </label>
              <input
                type="text"
                list="yx-channels"
                value={channelName}
                onChange={(e) => setChannelName(e.target.value)}
                placeholder="如 加拿大超快线-卡派"
                className="w-full rounded-lg border border-zinc-300 bg-white px-3 py-2 text-sm focus:border-primary focus:outline-none focus:ring-1 focus:ring-primary"
              />
              <datalist id="yx-channels">
                {CHANNELS.map((c) => (
                  <option key={c} value={c} />
                ))}
              </datalist>
            </div>
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

      {/* Result */}
      {result && (
        <>
          <div className="flex flex-wrap items-center justify-between gap-4 rounded-lg border border-zinc-200 bg-white p-4 shadow-sm">
            <div>
              <p className="text-sm text-zinc-600">
                转换完成：<span className="font-semibold text-emerald-700">{result.result.fileName}</span>
              </p>
              <p className="mt-0.5 text-xs text-zinc-400">
                仓库 {result.result.warehouseCode} · 总箱数 {result.result.totalBoxes} · 数据 {result.result.dataRows} 行
              </p>
            </div>
            <div className="flex gap-3">
              <a
                href={result.result.downloadUrl}
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
            💡 提示：系统自动解析易通标准模板的带电/带磁、报关方式、清关方式、总箱数、国家、仓库代码，
            并按天图模板逐列映射货箱明细。收件地址由模板内 VLOOKUP 公式根据仓库代码自动带出，
            交税方式（清关方式「无」→ 包税）与报关方式（普通报关 → 托管报关）自动映射。
          </p>
        </div>
      )}
    </div>
  );
}
