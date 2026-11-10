"use client";

import { useState, useCallback, useRef, useEffect } from "react";

// ============================================================
// Types
// ============================================================

interface Broker {
  id: string;
  name: string;
  bankInfo: string;
}

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
  hasCustoms: boolean;
  hasInternational: boolean;
  hasInvoice: boolean;
}

interface DownloadUrls {
  domestic: string;
  customs: string;
  international: string;
  invoice: string;
}

interface SplitResponse {
  sessionId: string;
  sourceFile: string;
  totalGroups: number;
  domesticCount: number;
  customsCount: number;
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

  // 报关行状态
  const [brokers, setBrokers] = useState<Broker[]>([]);
  const [selectedBrokerId, setSelectedBrokerId] = useState<string>("");
  const [brokersLoading, setBrokersLoading] = useState(true);
  const [showAddBroker, setShowAddBroker] = useState(false);
  const [newBrokerName, setNewBrokerName] = useState("");
  const [newBrokerBankInfo, setNewBrokerBankInfo] = useState("");
  const [savingBroker, setSavingBroker] = useState(false);
  const [brokerError, setBrokerError] = useState<string | null>(null);

  const loadBrokers = useCallback(async () => {
    try {
      setBrokersLoading(true);
      const res = await fetch("/api/brokers");
      const data = await res.json();
      if (res.ok && Array.isArray(data.brokers)) {
        setBrokers(data.brokers);
        setSelectedBrokerId((prev) => prev || data.brokers[0]?.id || "");
      }
    } catch {
      // 忽略加载失败，允许手动新增
    } finally {
      setBrokersLoading(false);
    }
  }, []);

  useEffect(() => { loadBrokers(); }, [loadBrokers]);

  const selectedBroker = brokers.find((b) => b.id === selectedBrokerId);

  const saveNewBroker = useCallback(async () => {
    if (!newBrokerName.trim()) { setBrokerError("请填写报关行名称"); return; }
    if (!newBrokerBankInfo.trim()) { setBrokerError("请填写报关行收款信息"); return; }
    setBrokerError(null);
    setSavingBroker(true);
    try {
      const res = await fetch("/api/brokers", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ name: newBrokerName.trim(), bankInfo: newBrokerBankInfo.trim() }),
      });
      const data = await res.json();
      if (!res.ok) { setBrokerError(data.error || "保存失败"); return; }
      if (Array.isArray(data.brokers)) {
        setBrokers(data.brokers);
        const added = data.brokers[data.brokers.length - 1];
        setSelectedBrokerId(added.id);
      }
      setNewBrokerName("");
      setNewBrokerBankInfo("");
      setShowAddBroker(false);
    } catch {
      setBrokerError("网络错误，保存失败");
    } finally {
      setSavingBroker(false);
    }
  }, [newBrokerName, newBrokerBankInfo]);

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
      if (selectedBroker) {
        formData.append("brokerName", selectedBroker.name);
        formData.append("brokerBankInfo", selectedBroker.bankInfo);
      }
      const res = await fetch("/api/pipixiong-split", { method: "POST", body: formData });
      const data: SplitResponse = await res.json();
      if (!res.ok) { setError(data.error || "拆分处理失败"); return; }
      setResult(data);
    } catch (err) {
      setError(`请求失败: ${err instanceof Error ? err.message : "网络错误"}`);
    } finally {
      setProcessing(false);
    }
  }, [selectedBroker]);

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
        <h1 className="text-2xl font-bold text-deep">🧾 皮皮熊账单拆分</h1>
        <p className="mt-1 text-sm text-zinc-500">
          上传皮皮熊合并账单 Excel，自动拆分为国内账单、报关费账单、国外账单和 INVOICE
        </p>
      </div>

      {/* Broker selector + Upload Area */}
      {!result && (
        <>
          {/* 报关行选择 */}
          <div className="rounded-xl border border-zinc-200 bg-white p-5 shadow-sm">
            <div className="mb-3 flex items-center justify-between">
              <div>
                <h2 className="text-sm font-semibold text-zinc-700">选择报关行</h2>
                <p className="text-xs text-zinc-400">报关费账单底部将打印此报关行的收款信息（客户自行付款给报关行）</p>
              </div>
              <button
                onClick={() => setShowAddBroker((v) => !v)}
                className="shrink-0 rounded-lg border border-zinc-300 px-3 py-1.5 text-xs font-medium text-zinc-600 hover:bg-zinc-50 transition-colors"
              >
                {showAddBroker ? "取消新增" : "+ 新增报关行"}
              </button>
            </div>

            {brokersLoading ? (
              <p className="text-sm text-zinc-400">加载报关行列表…</p>
            ) : (
              <select
                value={selectedBrokerId}
                onChange={(e) => setSelectedBrokerId(e.target.value)}
                className="w-full rounded-lg border border-zinc-300 bg-white px-3 py-2 text-sm text-zinc-700 focus:border-primary focus:outline-none"
              >
                {brokers.length === 0 && <option value="">暂无报关行，请先新增</option>}
                {brokers.map((b) => (
                  <option key={b.id} value={b.id}>{b.name}</option>
                ))}
              </select>
            )}

            {selectedBroker && (
              <pre className="mt-3 max-h-32 overflow-auto whitespace-pre-wrap rounded-lg bg-zinc-50 p-3 text-xs leading-relaxed text-zinc-500">
                {selectedBroker.bankInfo}
              </pre>
            )}

            {showAddBroker && (
              <div className="mt-4 space-y-3 border-t border-zinc-100 pt-4">
                <div>
                  <label className="mb-1 block text-xs font-medium text-zinc-600">报关行名称</label>
                  <input
                    value={newBrokerName}
                    onChange={(e) => setNewBrokerName(e.target.value)}
                    placeholder="如：上海某报关有限公司"
                    className="w-full rounded-lg border border-zinc-300 px-3 py-2 text-sm focus:border-primary focus:outline-none"
                  />
                </div>
                <div>
                  <label className="mb-1 block text-xs font-medium text-zinc-600">收款信息（多行，含户名/账号/开户行等）</label>
                  <textarea
                    value={newBrokerBankInfo}
                    onChange={(e) => setNewBrokerBankInfo(e.target.value)}
                    rows={6}
                    placeholder={"户名：xxx\n开户行：xxx\n账号：xxx"}
                    className="w-full rounded-lg border border-zinc-300 px-3 py-2 text-sm font-mono focus:border-primary focus:outline-none"
                  />
                </div>
                {brokerError && <p className="text-xs text-red-600">{brokerError}</p>}
                <button
                  onClick={saveNewBroker}
                  disabled={savingBroker}
                  className="rounded-lg bg-primary px-4 py-2 text-sm font-medium text-white hover:bg-primary/90 disabled:opacity-60 transition-colors"
                >
                  {savingBroker ? "保存中…" : "保存并选择"}
                </button>
              </div>
            )}
          </div>

          {/* Upload Area */}
          <div
            onDragOver={onDragOver} onDragLeave={onDragLeave} onDrop={onDrop}
            onClick={() => !processing && fileInputRef.current?.click()}
            className={`flex cursor-pointer flex-col items-center justify-center rounded-xl border-2 border-dashed p-12 transition-colors
              ${isDragging ? "border-primary bg-primary/5" : "border-[#C7E0F0] bg-white/60 hover:border-primary hover:bg-primary/5"}
              ${processing ? "pointer-events-none opacity-60" : ""}`}
          >
            {processing ? (
              <>
                <div className="mb-4 h-10 w-10 animate-spin rounded-full border-4 border-zinc-300 border-t-primary" />
                <p className="text-lg font-medium text-zinc-600">正在处理中...</p>
                <p className="mt-1 text-sm text-zinc-400">读取四个数据页、拆分国内/报关/国外费用、生成 INVOICE</p>
              </>
            ) : (
              <>
                <svg className="mb-4 h-12 w-12 text-zinc-400" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.5}
                    d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12" />
                </svg>
                <p className="text-lg font-medium text-zinc-600">拖拽皮皮熊账单 Excel 到此处，或点击上传</p>
                <p className="mt-1 text-sm text-zinc-400">支持 .xlsx / .xls · 自动识别四个数据页 · 生成四份账单</p>
              </>
            )}
            <input ref={fileInputRef} type="file" accept=".xlsx,.xls" onChange={onFileChange} className="hidden" />
          </div>
        </>
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
                  {result.totalGroups} 组账单 · 国内 {result.domesticCount} · 报关 {result.customsCount} · 国外 {result.internationalCount} · INVOICE {result.invoiceCount}
                </p>
              </div>
              <button onClick={reset}
                className="rounded-lg border border-zinc-300 px-4 py-2 text-sm font-medium text-zinc-600 hover:bg-zinc-50 transition-colors">
                重新拆分
              </button>
            </div>

            <div className="grid grid-cols-1 gap-3 sm:grid-cols-2 xl:grid-cols-4">
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

              {/* 报关费账单 */}
              <a href={result.downloadUrls.customs} download
                className="flex items-center gap-3 rounded-lg border-2 border-amber-200 bg-amber-50 p-4 hover:bg-amber-100 transition-colors">
                <span className="flex h-10 w-10 items-center justify-center rounded-lg bg-amber-100 text-xl">🛃</span>
                <div>
                  <p className="font-semibold text-amber-700">报关费账单</p>
                  <p className="text-xs text-amber-500">客户自行付报关行</p>
                  <p className="text-xs text-amber-400 mt-0.5">{result.customsCount} 个 sheet</p>
                </div>
                <svg className="ml-auto h-5 w-5 text-amber-400" fill="none" viewBox="0 0 24 24" stroke="currentColor">
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
                    <th className="px-4 py-2">国家</th>
                    <th className="px-4 py-2">数量</th>
                    <th className="px-4 py-2 text-right">拖车费</th>
                    <th className="px-4 py-2 text-right">报关费</th>
                    <th className="px-4 py-2 text-right">港杂费</th>
                    <th className="px-4 py-2 text-right">费用(USD)</th>
                    <th className="px-4 py-2 text-center">国内</th>
                    <th className="px-4 py-2 text-center">报关</th>
                    <th className="px-4 py-2 text-center">国外</th>
                    <th className="px-4 py-2 text-center">INVOICE</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-zinc-100">
                  {result.groups.map((g, i) => (
                    <tr key={i} className="hover:bg-zinc-50">
                      <td className="px-4 py-2 font-mono text-xs text-zinc-400">{g.seq}</td>
                      <td className="px-4 py-2 font-mono text-xs font-medium text-zinc-700">{g.forwarderId}</td>
                      <td className="px-4 py-2 text-xs">{g.country || "—"}</td>
                      <td className="px-4 py-2 text-xs">{g.qty}</td>
                      <td className="px-4 py-2 text-right font-mono text-xs">{g.truckFee > 0 ? `¥${fmtMoney(g.truckFee)}` : "-"}</td>
                      <td className="px-4 py-2 text-right font-mono text-xs">{g.customsFee > 0 ? `¥${fmtMoney(g.customsFee)}` : "-"}</td>
                      <td className="px-4 py-2 text-right font-mono text-xs">{g.portFee > 0 ? `¥${fmtMoney(g.portFee)}` : "-"}</td>
                      <td className="px-4 py-2 text-right font-mono text-xs font-medium">
                        {g.oceanTaxUSD !== 0 ? `$${fmtMoney(g.oceanTaxUSD)}` : "-"}
                      </td>
                      <td className="px-4 py-2 text-center">{g.hasDomestic ? "✅" : "—"}</td>
                      <td className="px-4 py-2 text-center">{g.hasCustoms ? "✅" : "—"}</td>
                      <td className="px-4 py-2 text-center">{g.hasInternational ? "✅" : "—"}</td>
                      <td className="px-4 py-2 text-center">{g.hasInvoice ? "✅" : "—"}</td>
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
              <li><strong>国内账单</strong>：深圳抬头，RMB 税票（拖车费+港杂费），含深圳银行账户信息</li>
              <li><strong>报关费账单</strong>：深圳抬头，RMB 报关费，底部打印所选报关行的收款信息（客户自行付款）</li>
              <li><strong>国外账单</strong>：香港抬头，USD 费用，含汇丰银行账户信息</li>
              <li><strong>INVOICE</strong>：香港抬头，客户发票格式，含 JOB NO / 项目 / Total / 银行信息</li>
              <li>自动读取第 1/2/3/4 页数据，只处理&quot;是否支付&quot; ≠ Y 的记录</li>
            </ul>
          </div>
        </>
      )}

      {/* Initial tip */}
      {!result && !error && (
        <div className="rounded-lg border border-zinc-200 bg-white p-4 text-sm text-zinc-500">
          <p>
            💡 提示：系统自动从第 1/2/3/4 四个数据页读取未支付记录，按货代识别号拆分，
            分别生成国内账单（深圳/RMB）、报关费账单（客户付报关行）、国外账单（香港/USD）
            和 INVOICE（香港/客户发票）四个 Excel 文件。上传前请先选择或新增报关行。
          </p>
        </div>
      )}
    </div>
  );
}
