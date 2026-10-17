"use client";

import { useState, useCallback, useRef } from "react";

// ============================================================
// Types
// ============================================================

interface ResultFile {
  so: string;
  sku: string;
  fileName: string;
}

interface ResultItem {
  sourceFile: string;
  zipFileName: string;
  total: number;
  files: ResultFile[];
  skipped: { sku: string; reason: string }[];
  warnings: string[];
  downloadUrl: string;
}

interface ConvertResponse {
  sessionId: string;
  result: ResultItem;
  error?: string;
}

const CHANNELS = [
  "OA-拆送",
  "OA-奥克兰卡派",
  "Match12-海派",
  "Match12-卡派",
  "Match15-海派",
  "Match15-卡派",
  "萨瓦纳专线-卡派",
  "休斯顿专线-卡派",
  "芝加哥海铁快车-卡派",
  "KC-卡车派送",
  "KD-快递派送",
];

const PICKUP_METHODS = ["自送货", "上门提货"];

const DELIVERY_SITES = [
  "清溪仓",
  "松山湖仓",
  "深圳仓",
  "中山仓",
  "广州仓",
  "揭阳仓",
  "汕头仓",
  "青岛仓",
];

// ============================================================
// Page
// ============================================================

export default function LsToSupplierPage() {
  const [isDragging, setIsDragging] = useState(false);
  const [processing, setProcessing] = useState(false);
  const [file, setFile] = useState<File | null>(null);
  const [soSkuFile, setSoSkuFile] = useState<File | null>(null);
  const [channelName, setChannelName] = useState("OA-拆送");
  const [pickupMethod, setPickupMethod] = useState("自送货");
  const [deliverySite, setDeliverySite] = useState("清溪仓");
  const [result, setResult] = useState<ConvertResponse | null>(null);
  const [error, setError] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const soSkuInputRef = useRef<HTMLInputElement>(null);

  const isSupported = (name: string) => /\.(xlsx|xls)$/i.test(name);

  const convert = useCallback(async () => {
    if (!file) {
      setError("请先上传朗胜易通下单发票");
      return;
    }
    if (!soSkuFile) {
      setError("请上传 SO-SKU 对应表");
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
      formData.append("soSkuFile", soSkuFile);
      formData.append("channelName", channelName.trim());
      formData.append("pickupMethod", pickupMethod);
      formData.append("deliverySite", deliverySite);

      const res = await fetch("/api/ls-to-supplier", {
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
  }, [file, soSkuFile, channelName, pickupMethod, deliverySite]);

  const handleFile = useCallback((incoming: File) => {
    if (!isSupported(incoming.name)) {
      setError("请选择 .xlsx 格式的易通下单发票");
      return;
    }
    setFile(incoming);
    setError(null);
    setResult(null);
  }, []);

  const handleSoSkuFile = useCallback((incoming: File) => {
    if (!isSupported(incoming.name)) {
      setError("请选择 .xlsx 格式的 SO-SKU 对应表");
      return;
    }
    setSoSkuFile(incoming);
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

  const onSoSkuChange = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      const files = Array.from(e.target.files || []);
      if (files.length) handleSoSkuFile(files[0]);
      e.target.value = "";
    },
    [handleSoSkuFile]
  );

  const reset = useCallback(() => {
    setResult(null);
    setError(null);
    setFile(null);
    setSoSkuFile(null);
    setChannelName("OA-拆送");
    setPickupMethod("自送货");
    setDeliverySite("清溪仓");
  }, []);

  return (
    <div className="space-y-8">
      {/* Header */}
      <div>
        <h1 className="text-2xl font-bold text-deep">🏷️ 朗胜发票转美线下单映射</h1>
        <p className="mt-1 text-sm text-zinc-500">
          上传朗胜易通下单发票 + SO-SKU 对应表，按 SKU 自动拆分为多张「美线发票下单映射」
        </p>
      </div>

      {/* Form */}
      {!result && (
        <div className="space-y-5">
          {/* 源发票上传 */}
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
                <p className="mt-1 text-sm text-zinc-400">点击可重新选择易通下单发票</p>
              </>
            ) : (
              <>
                <svg className="mb-4 h-12 w-12 text-zinc-400" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.5} d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12" />
                </svg>
                <p className="text-lg font-medium text-zinc-600">拖拽朗胜易通下单发票到此处，或点击选择</p>
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

          {/* SO-SKU 对应表上传 */}
          <div
            onClick={() => !processing && soSkuInputRef.current?.click()}
            className={`
              flex cursor-pointer items-center justify-between rounded-xl border-2 border-dashed p-4 transition-colors
              ${
                soSkuFile
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
                  {soSkuFile ? soSkuFile.name : "SO-SKU 对应表"}
                  <span className="ml-1 text-red-500">*</span>
                </p>
                <p className="text-xs text-zinc-400">提供「发票文件名 / SO / SKU」三列映射，点击选择</p>
              </div>
            </div>
            {soSkuFile && (
              <span className="inline-flex items-center gap-1 rounded bg-emerald-100 px-2 py-0.5 text-xs font-medium text-emerald-700">
                ✓ 已选择
              </span>
            )}
            <input
              ref={soSkuInputRef}
              type="file"
              accept=".xlsx,.xls"
              onChange={onSoSkuChange}
              className="hidden"
            />
          </div>

          {/* Options */}
          <div className="grid gap-4 sm:grid-cols-3">
            <div>
              <label className="mb-1.5 block text-sm font-medium text-zinc-700">
                服务渠道 <span className="text-red-500">*</span>
              </label>
              <input
                type="text"
                list="ls-channels"
                value={channelName}
                onChange={(e) => setChannelName(e.target.value)}
                placeholder="如 OA-拆送"
                className="w-full rounded-lg border border-zinc-300 bg-white px-3 py-2 text-sm focus:border-primary focus:outline-none focus:ring-1 focus:ring-primary"
              />
              <datalist id="ls-channels">
                {CHANNELS.map((c) => (
                  <option key={c} value={c} />
                ))}
              </datalist>
            </div>
            <div>
              <label className="mb-1.5 block text-sm font-medium text-zinc-700">提货方式</label>
              <select
                value={pickupMethod}
                onChange={(e) => setPickupMethod(e.target.value)}
                className="w-full rounded-lg border border-zinc-300 bg-white px-3 py-2 text-sm focus:border-primary focus:outline-none focus:ring-1 focus:ring-primary"
              >
                {PICKUP_METHODS.map((m) => (
                  <option key={m} value={m}>
                    {m}
                  </option>
                ))}
              </select>
            </div>
            <div>
              <label className="mb-1.5 block text-sm font-medium text-zinc-700">交货站点</label>
              <select
                value={deliverySite}
                onChange={(e) => setDeliverySite(e.target.value)}
                className="w-full rounded-lg border border-zinc-300 bg-white px-3 py-2 text-sm focus:border-primary focus:outline-none focus:ring-1 focus:ring-primary"
              >
                {DELIVERY_SITES.map((s) => (
                  <option key={s} value={s}>
                    {s}
                  </option>
                ))}
              </select>
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
                转换完成：<span className="font-semibold text-emerald-700">{result.result.zipFileName}</span>
              </p>
              <p className="mt-0.5 text-xs text-zinc-400">
                共 {result.result.total} 张美线发票
                {result.result.skipped.length > 0 && ` · ${result.result.skipped.length} 个 SKU 跳过`}
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

          {result.result.files.length > 0 && (
            <div className="rounded-lg border border-zinc-200 bg-white p-4">
              <p className="mb-2 text-sm font-medium text-zinc-700">生成的发票清单</p>
              <div className="max-h-64 overflow-y-auto">
                <table className="w-full text-sm">
                  <thead>
                    <tr className="border-b border-zinc-100 text-left text-xs text-zinc-400">
                      <th className="py-1.5 pr-3 font-medium">SO 号</th>
                      <th className="py-1.5 pr-3 font-medium">SKU</th>
                      <th className="py-1.5 font-medium">文件名</th>
                    </tr>
                  </thead>
                  <tbody>
                    {result.result.files.map((f, i) => (
                      <tr key={i} className="border-b border-zinc-50 last:border-0">
                        <td className="py-1.5 pr-3 font-mono text-zinc-700">{f.so}</td>
                        <td className="py-1.5 pr-3 text-zinc-600">{f.sku}</td>
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

          {result.result.skipped?.length > 0 && (
            <div className="rounded-lg border border-amber-200 bg-amber-50/60 p-4 text-sm text-amber-700">
              <p className="mb-1 font-medium">以下 SKU 未能匹配到 SO 号，已跳过：</p>
              <ul className="list-inside list-disc space-y-0.5">
                {result.result.skipped.map((s, i) => (
                  <li key={i}>
                    {s.sku || "(空 SKU)"} — {s.reason}
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
            💡 提示：朗胜按 SKU 拆分下单（一个 SKU 一个 SO）。系统自动解析易通发票的带电/带磁、报关方式、
            仓库代码，并按 SO-SKU 对应表补客户订单号，逐 SKU 生成美线发票。
            收件地址由模板内 VLOOKUP 公式根据仓库代码自动带出，预计提货时间自动填今天。
          </p>
        </div>
      )}
    </div>
  );
}
