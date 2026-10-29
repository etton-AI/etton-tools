"use client";

import { useState, useCallback, useRef } from "react";

// ============================================================
// Types
// ============================================================

interface GroupItem {
  so: string;
  orderNo: string;
  country: string;
  warehouseCode: string;
  destType: string;
  goodsDescription: string;
  goodsCategory: string;
  isFragile: boolean;
  currency: string;
  currencyCode: string;
  totalValue: number;
  totalBoxes: number;
  totalWeight: number;
  remark: string;
}

interface CurrencyWarning {
  so: string;
  country: string;
  declared: string;
  expected: string;
}

interface NonFbaWarning {
  so: string;
  fbaCodes: string[];
  destType: string;
}

interface SplitInfo {
  so: string;
}

interface ConvertResponse {
  sessionId: string;
  sourceFile: string;
  outputName: string;
  totalGroups: number;
  lockWarehouseCount: number;
  currencyWarnings: CurrencyWarning[];
  nonFbaWarnings: NonFbaWarning[];
  splitInfos: SplitInfo[];
  groups: GroupItem[];
  downloadUrl: string;
  error?: string;
}

// ============================================================
// 表单默认值
// ============================================================

const TRANSPORT_OPTIONS = ["海运", "空运", "铁路", "公路", "全程快递", "多式联运"];
const DELIVERY_OPTIONS = ["卡车派", "快递派", "待定"];
const PACKAGE_OPTIONS = ["纸箱", "木箱", "木箱+纸箱", "托盘", "航空箱", "集装箱"];

// ============================================================
// Page
// ============================================================

export default function KuajingbaoInsurancePage() {
  const [isDragging, setIsDragging] = useState(false);
  const [processing, setProcessing] = useState(false);
  const [fileName, setFileName] = useState<string>("");
  const [result, setResult] = useState<ConvertResponse | null>(null);
  const [error, setError] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // 配置项
  const [shipDate, setShipDate] = useState<string>("");
  const [transportMode, setTransportMode] = useState("海运");
  const [deliveryMethod, setDeliveryMethod] = useState("卡车派");
  const [originPlace, setOriginPlace] = useState("深圳");
  const [packageType, setPackageType] = useState("纸箱");

  const isSupported = (name: string) => /\.(xlsx|xls)$/i.test(name);

  const convert = useCallback(
    async (f: File) => {
      setProcessing(true);
      setError(null);
      setResult(null);
      setFileName(f.name);

      try {
        const formData = new FormData();
        formData.append("file", f);
        formData.append("shipDate", shipDate);
        formData.append("transportMode", transportMode);
        formData.append("deliveryMethod", deliveryMethod);
        formData.append("originPlace", originPlace);
        formData.append("packageType", packageType);

        const res = await fetch("/api/kuajingbao-insurance", {
          method: "POST",
          body: formData,
        });
        const data: ConvertResponse = await res.json();

        if (!res.ok) {
          setError(data.error || "生成失败");
          return;
        }
        setResult(data);
      } catch {
        setError("网络错误，请重试");
      } finally {
        setProcessing(false);
      }
    },
    [shipDate, transportMode, deliveryMethod, originPlace, packageType]
  );

  const handleFile = useCallback(
    (incoming: File) => {
      if (!isSupported(incoming.name)) {
        setError("请选择 .xlsx / .xls 文件");
        return;
      }
      void convert(incoming);
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
    setFileName("");
  }, []);

  const inputCls =
    "rounded-lg border border-zinc-300 bg-white px-3 py-2 text-sm text-zinc-800 focus:border-primary focus:outline-none focus:ring-1 focus:ring-primary";

  return (
    <div className="space-y-8">
      {/* Header */}
      <div>
        <h1 className="text-2xl font-bold text-deep">📋 跨境堡批量投保</h1>
        <p className="mt-1 text-sm text-zinc-500">
          上传系统导出的货箱清单，按 SO 号自动生成货物描述、货值、币种、备注，填充跨境堡批量投保箱单
        </p>
      </div>

      {/* 配置项 */}
      {!result && (
        <div className="rounded-xl border border-zinc-200 bg-white p-5 shadow-sm">
          <p className="mb-3 text-xs font-semibold uppercase tracking-wide text-zinc-500">
            投保基础信息（生成后按模板示例常量填充）
          </p>
          <div className="grid grid-cols-2 gap-4 sm:grid-cols-3">
            <label className="flex flex-col gap-1">
              <span className="text-xs text-zinc-500">
                起运日期 <span className="text-red-500">*</span>
              </span>
              <input
                type="date"
                value={shipDate}
                onChange={(e) => setShipDate(e.target.value)}
                className={inputCls}
              />
            </label>
            <label className="flex flex-col gap-1">
              <span className="text-xs text-zinc-500">干线运输方式</span>
              <select value={transportMode} onChange={(e) => setTransportMode(e.target.value)} className={inputCls}>
                {TRANSPORT_OPTIONS.map((o) => (
                  <option key={o} value={o}>{o}</option>
                ))}
              </select>
            </label>
            <label className="flex flex-col gap-1">
              <span className="text-xs text-zinc-500">派送方式</span>
              <select value={deliveryMethod} onChange={(e) => setDeliveryMethod(e.target.value)} className={inputCls}>
                {DELIVERY_OPTIONS.map((o) => (
                  <option key={o} value={o}>{o}</option>
                ))}
              </select>
            </label>
            <label className="flex flex-col gap-1">
              <span className="text-xs text-zinc-500">起运地</span>
              <input value={originPlace} onChange={(e) => setOriginPlace(e.target.value)} className={inputCls} />
            </label>
            <label className="flex flex-col gap-1">
              <span className="text-xs text-zinc-500">包装类型</span>
              <select value={packageType} onChange={(e) => setPackageType(e.target.value)} className={inputCls}>
                {PACKAGE_OPTIONS.map((o) => (
                  <option key={o} value={o}>{o}</option>
                ))}
              </select>
            </label>
          </div>
        </div>
      )}

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
              <p className="text-lg font-medium text-zinc-600">正在生成投保箱单...</p>
              <p className="mt-1 max-w-md truncate text-sm text-zinc-400">{fileName}</p>
            </>
          ) : (
            <>
              <svg className="mb-4 h-12 w-12 text-zinc-400" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                <path
                  strokeLinecap="round"
                  strokeLinejoin="round"
                  strokeWidth={1.5}
                  d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12"
                />
              </svg>
              <p className="text-lg font-medium text-zinc-600">拖拽系统货箱清单到此处，或点击选择</p>
              <p className="mt-1 text-sm text-zinc-400">支持 .xlsx / .xls，按 SO 号逐票生成</p>
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
          <button onClick={reset} className="ml-auto shrink-0 text-sm font-medium text-red-600 hover:text-red-800">
            重新上传
          </button>
        </div>
      )}

      {/* Results */}
      {result && (
        <>
          {/* 审核结果 */}
          {(result.currencyWarnings.length > 0 ||
            result.nonFbaWarnings.length > 0 ||
            result.splitInfos.length > 0) && (
            <div className="space-y-3">
              <p className="text-xs font-semibold uppercase tracking-wide text-zinc-500">
                审核结果（请人工核对确认，以下告警不阻断下载）
              </p>

              {result.splitInfos.length > 0 && (
                <div className="rounded-lg border border-sky-300 bg-sky-50 p-4">
                  <p className="text-sm font-semibold text-sky-800">
                    🔀 易碎品拆分（{result.splitInfos.length} 票）
                  </p>
                  <ul className="mt-2 space-y-1 text-sm text-sky-800">
                    {result.splitInfos.map((s) => (
                      <li key={s.so}>
                        <span className="font-mono text-xs">{s.so}</span>
                        ：同时含易碎品与非易碎品，已拆分为「{s.so}-1」（易碎品）「{s.so}-2」（非易碎品）两票
                      </li>
                    ))}
                  </ul>
                </div>
              )}

              {result.currencyWarnings.length > 0 && (
                <div className="rounded-lg border border-amber-300 bg-amber-50 p-4">
                  <p className="text-sm font-semibold text-amber-800">
                    ⚠️ 币种异常（{result.currencyWarnings.length} 票，请确认是否客户填错币种）
                  </p>
                  <ul className="mt-2 space-y-1 text-sm text-amber-800">
                    {result.currencyWarnings.map((w) => (
                      <li key={w.so}>
                        <span className="font-mono text-xs">{w.so}</span>
                        （{w.country}）：清单申报币种为 <span className="font-semibold">{w.declared}</span>，
                        该国正常应为 <span className="font-semibold">{w.expected}</span>
                      </li>
                    ))}
                  </ul>
                </div>
              )}

              {result.nonFbaWarnings.length > 0 && (
                <div className="rounded-lg border border-amber-300 bg-amber-50 p-4">
                  <p className="text-sm font-semibold text-amber-800">
                    ⚠️ 非 FBA 仓（{result.nonFbaWarnings.length} 票，目的地类型已自动填入，请人工确认）
                  </p>
                  <ul className="mt-2 space-y-1 text-sm text-amber-800">
                    {result.nonFbaWarnings.map((w) => (
                      <li key={w.so}>
                        <span className="font-mono text-xs">{w.so}</span>
                        ：入仓编号 <span className="font-mono text-xs">{w.fbaCodes.join(" ")}</span> 不含「FBA」字眼，
                        {w.destType !== "FBA" ? (
                          <>已识别为「<span className="font-semibold">{w.destType}</span>」</>
                        ) : (
                          <>无法识别仓类型，暂按「FBA」填写</>
                        )}
                      </li>
                    ))}
                  </ul>
                </div>
              )}
            </div>
          )}

          {/* Summary + actions */}
          <div className="flex flex-wrap items-center justify-between gap-4 rounded-lg border border-zinc-200 bg-white p-4 shadow-sm">
            <div>
              <p className="text-sm text-zinc-600">
                生成完成：<span className="font-semibold text-emerald-700">{result.totalGroups} 票</span>
                {result.splitInfos.length > 0 && (
                  <span className="font-semibold text-sky-700"> · 易碎拆分 {result.splitInfos.length} 票</span>
                )}
                {result.lockWarehouseCount > 0 && (
                  <span className="font-semibold text-emerald-700"> · 全美锁仓（带备注）{result.lockWarehouseCount} 票</span>
                )}
                {result.currencyWarnings.length === 0 &&
                  result.nonFbaWarnings.length === 0 &&
                  result.splitInfos.length === 0 && (
                    <span className="font-semibold text-emerald-700"> · 审核通过，未发现异常</span>
                  )}
              </p>
              <p className="mt-0.5 text-xs text-zinc-400">
                {result.sourceFile} → {result.outputName}
              </p>
            </div>
            <div className="flex gap-3">
              <a
                href={result.downloadUrl}
                className="inline-flex items-center gap-2 rounded-lg bg-primary px-4 py-2 text-sm font-medium text-white shadow-sm hover:bg-primary-dark transition-colors"
              >
                <svg className="h-4 w-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                </svg>
                下载投保箱单
              </a>
              <button
                onClick={reset}
                className="inline-flex items-center gap-2 rounded-lg border border-zinc-300 bg-white px-4 py-2 text-sm font-medium text-zinc-600 hover:bg-zinc-50 transition-colors"
              >
                重新生成
              </button>
            </div>
          </div>

          {/* Preview table */}
          <div className="overflow-hidden rounded-lg border border-zinc-200 bg-white">
            <div className="border-b border-zinc-200 bg-zinc-50 px-4 py-2 text-xs font-semibold uppercase tracking-wide text-zinc-500">
              生成结果预览（{result.groups.length} 票）
            </div>
            <div className="overflow-x-auto">
              <table className="w-full text-left text-sm">
                <thead>
                  <tr className="border-b border-zinc-100 text-xs text-zinc-400">
                    <th className="px-4 py-2 font-medium">原单号</th>
                    <th className="px-4 py-2 font-medium">货物类别</th>
                    <th className="px-4 py-2 font-medium">仓库</th>
                    <th className="px-4 py-2 font-medium">目的地类型</th>
                    <th className="px-4 py-2 font-medium">货物描述</th>
                    <th className="px-4 py-2 font-medium">币种</th>
                    <th className="px-4 py-2 font-medium">货值</th>
                    <th className="px-4 py-2 font-medium">箱数</th>
                    <th className="px-4 py-2 font-medium">公斤数</th>
                    <th className="px-4 py-2 font-medium">备注</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-zinc-50">
                  {result.groups.map((g) => (
                    <tr key={g.orderNo} className="text-zinc-700">
                      <td className="whitespace-nowrap px-4 py-2 font-mono text-xs">{g.orderNo}</td>
                      <td className="whitespace-nowrap px-4 py-2">
                        {g.isFragile ? (
                          <span className="rounded bg-orange-100 px-1.5 py-0.5 text-xs font-medium text-orange-700">是易碎品</span>
                        ) : (
                          <span className="rounded bg-zinc-100 px-1.5 py-0.5 text-xs font-medium text-zinc-500">无易碎品</span>
                        )}
                      </td>
                      <td className="whitespace-nowrap px-4 py-2 font-mono text-xs">{g.warehouseCode}</td>
                      <td className="whitespace-nowrap px-4 py-2">
                        {g.destType === "FBA" ? (
                          <span className="text-xs text-zinc-600">{g.destType}</span>
                        ) : (
                          <span className="rounded bg-amber-100 px-1.5 py-0.5 text-xs font-medium text-amber-700">
                            {g.destType}
                          </span>
                        )}
                      </td>
                      <td className="px-4 py-2">{g.goodsDescription}</td>
                      <td className="whitespace-nowrap px-4 py-2">{g.currency}</td>
                      <td className="whitespace-nowrap px-4 py-2 text-right">{g.totalValue.toFixed(2)}</td>
                      <td className="whitespace-nowrap px-4 py-2 text-right">{g.totalBoxes}</td>
                      <td className="whitespace-nowrap px-4 py-2 text-right">{g.totalWeight}</td>
                      <td className="px-4 py-2">
                        {g.remark ? (
                          <span className="text-xs text-emerald-600">已生成</span>
                        ) : (
                          <span className="text-xs text-zinc-300">—</span>
                        )}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </>
      )}

      {/* Tip */}
      {!result && (
        <div className="rounded-lg border border-zinc-200 bg-white p-4 text-sm text-zinc-500">
          <p>
            💡 自动完成：货物描述（品名+数量PCS）、货值（总申报货值求和）、币种（申报币种转中文）、
            货物类别（按材质自动判定易碎/非易碎）、备注（「全美锁仓」渠道按仓库代码生成 AGL 备注）。
            审核：按材质拆分易碎/非易碎、按目的国校验币种、按入仓编号/仓库代码自动识别目的地类型（FBA/沃尔玛/卫星/海外仓）。
          </p>
        </div>
      )}
    </div>
  );
}
