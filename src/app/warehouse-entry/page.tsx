"use client";

import { useState, useCallback, useRef } from "react";
import type { SuggestionRow } from "@/lib/warehouse-entry";

// ============================================================
// Types
// ============================================================

interface UploadResponse {
  rows: SuggestionRow[];
  supplierTotal: number;
  summary: {
    customerCount: number;
    boxCount: number;
    matchedCount: number;
    unmatchedCount: number;
    historyCount: number;
  };
  error?: string;
}

// ============================================================
// Helpers
// ============================================================

function round2(n: number): number {
  return Math.round(n * 100) / 100;
}

function round3(n: number): number {
  return Math.round(n * 1000) / 1000;
}

function fmt(n: number, digits = 2): string {
  return n.toLocaleString("zh-CN", {
    minimumFractionDigits: digits,
    maximumFractionDigits: digits,
  });
}

/** 依据长宽高实重重算 材积重 / 计费重 / 三边和 */
function recompute(s: SuggestionRow["suggestion"]): SuggestionRow["suggestion"] {
  const volumeWeight = round3((s.lengthCm * s.widthCm * s.heightCm) / 6000);
  const chargeableWeight = Math.max(s.actualWeight, volumeWeight);
  const sumSides = round2(s.lengthCm + s.widthCm + s.heightCm);
  return { ...s, volumeWeight, chargeableWeight, sumSides };
}

/** 依据当前建议值实时重算报警（阈值与服务端 buildSuggestions 一致，供应商原始箱规 vs 客户） */
function recomputeAlarms(r: SuggestionRow): string[] {
  // 未匹配行（pickedRank 为 null）：保留人工复核提示
  if (r.pickedRank === null) return ["⚠需人工复核"];
  const alarms: string[] = [];
  const supplierSumSides = round2(r.supplier.lengthCm + r.supplier.widthCm + r.supplier.heightCm);
  if (r.supplier.volumeWeight >= r.supplier.actualWeight) {
    // 材积主导：查三边和差、材积重差
    if (Math.abs(supplierSumSides - r.customer.sumSides) >= 6) {
      alarms.push("三边和差异超限，请核查过机图");
    }
    if (Math.abs(r.supplier.volumeWeight - r.customer.volumeWeight) >= 2) {
      alarms.push("材积重差异超限，请核查过机图");
    }
  } else {
    // 实重主导：查实重差
    if (Math.abs(r.supplier.actualWeight - r.customer.actualWeight) >= 0.5) {
      alarms.push("实重差异超限，请核查过机图");
    }
  }
  if (r.supplierMaxVolumeWeight - r.suggestion.volumeWeight >= 2) {
    alarms.push("供应商存在过大箱，请核查过机图");
  }
  if (r.historyMax && r.historyMax.chargeableWeight > r.suggestion.chargeableWeight) {
    alarms.push("建议参考历史最大值放大");
  }
  return alarms;
}

// ============================================================
// Page
// ============================================================

export default function WarehouseEntryPage() {
  const [customerFile, setCustomerFile] = useState<File | null>(null);
  const [supplierFile, setSupplierFile] = useState<File | null>(null);
  const [isDraggingCustomer, setIsDraggingCustomer] = useState(false);
  const [isDraggingSupplier, setIsDraggingSupplier] = useState(false);
  const [processing, setProcessing] = useState(false);
  const [exporting, setExporting] = useState(false);
  const [rows, setRows] = useState<SuggestionRow[]>([]);
  const [supplierTotal, setSupplierTotal] = useState(0);
  const [summary, setSummary] = useState<UploadResponse["summary"] | null>(null);
  const [historyCount, setHistoryCount] = useState(0);
  const [error, setError] = useState<string | null>(null);
  const [notice, setNotice] = useState<string | null>(null);
  const customerInputRef = useRef<HTMLInputElement>(null);
  const supplierInputRef = useRef<HTMLInputElement>(null);
  const historyInputRef = useRef<HTMLInputElement>(null);

  const handleSubmit = useCallback(async () => {
    if (!customerFile || !supplierFile) return;
    setProcessing(true);
    setError(null);
    setRows([]);
    setSummary(null);

    try {
      const formData = new FormData();
      formData.append("customerFile", customerFile);
      formData.append("supplierFile", supplierFile);

      const res = await fetch("/api/warehouse-entry", {
        method: "POST",
        body: formData,
      });
      const data: UploadResponse = await res.json();

      if (!res.ok) {
        setError(data.error || "处理失败");
        return;
      }

      setRows(data.rows);
      setSupplierTotal(data.supplierTotal);
      setSummary(data.summary);
      setHistoryCount(data.summary.historyCount);
    } catch (err) {
      const msg = err instanceof Error ? err.message : "网络错误，请重试";
      console.error("TR入仓请求失败:", err);
      setError(`请求失败: ${msg}`);
    } finally {
      setProcessing(false);
    }
  }, [customerFile, supplierFile]);

  const reset = useCallback(() => {
    setRows([]);
    setSummary(null);
    setSupplierTotal(0);
    setError(null);
    setNotice(null);
    setCustomerFile(null);
    setSupplierFile(null);
  }, []);

  // 编辑建议长/宽/高/实重（重算材积重/计费重/三边和，并实时重算报警）
  const editSuggestion = useCallback(
    (index: number, field: keyof SuggestionRow["suggestion"], value: number) => {
      setRows((prev) =>
        prev.map((r, i) => {
          if (i !== index) return r;
          const next = { ...r, suggestion: recompute({ ...r.suggestion, [field]: value }) };
          return { ...next, alarms: recomputeAlarms(next) };
        })
      );
    },
    []
  );

  // 导出
  const handleExport = useCallback(async () => {
    if (rows.length === 0) return;
    setExporting(true);
    setError(null);
    setNotice(null);

    try {
      const res = await fetch("/api/warehouse-entry/export", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ rows, supplierTotal }),
      });

      if (!res.ok) {
        const data = await res.json().catch(() => null);
        setError(data?.error || "导出失败");
        return;
      }

      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      const totalBoxes = rows.reduce((s, r) => s + (r.totalBoxes || 0), 0);
      a.download = `内部三类数据_${new Date().toISOString().slice(0, 10)}_合计总箱数${totalBoxes}.xlsx`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);

      setNotice("已导出，并已累积到历史库");
    } catch (err) {
      console.error("导出失败:", err);
      setError("导出失败，请重试");
    } finally {
      setExporting(false);
    }
  }, [rows, supplierTotal]);

  // 导入历史库
  const handleHistoryImport = useCallback(
    async (e: React.ChangeEvent<HTMLInputElement>) => {
      const file = e.target.files?.[0];
      if (!file) return;
      setNotice(null);
      setError(null);

      try {
        const formData = new FormData();
        formData.append("historyFile", file);
        const res = await fetch("/api/warehouse-entry/history", {
          method: "POST",
          body: formData,
        });
        const data = await res.json();
        if (!res.ok) {
          setError(data.error || "历史库导入失败");
          return;
        }
        setHistoryCount(data.totalCount);
        setNotice(`历史库已导入 ${data.importedCount} 条，当前共 ${data.totalCount} 条`);
      } catch {
        setError("历史库导入失败，请重试");
      } finally {
        if (historyInputRef.current) historyInputRef.current.value = "";
      }
    },
    []
  );

  // 导出历史库
  const handleHistoryExport = useCallback(async () => {
    try {
      const res = await fetch("/api/warehouse-entry/history");
      if (!res.ok) {
        setError("历史库导出失败");
        return;
      }
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `历史库_${new Date().toISOString().slice(0, 10)}.xlsx`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    } catch {
      setError("历史库导出失败，请重试");
    }
  }, []);

  const canSubmit = customerFile && supplierFile && !processing;

  // 全局校验：出给客户总计费重 vs 供应商总计费重
  const customerTotal = rows.reduce(
    (s, r) => s + r.suggestion.chargeableWeight * r.totalBoxes,
    0
  );
  const globalDiff = customerTotal - supplierTotal;
  const globalWarn = rows.length > 0 && globalDiff <= 0;

  const makeDragHandlers = (
    setter: (v: boolean) => void,
    fileSetter: (f: File) => void
  ) => ({
    onDragOver: (e: React.DragEvent) => {
      e.preventDefault();
      setter(true);
    },
    onDragLeave: (e: React.DragEvent) => {
      e.preventDefault();
      setter(false);
    },
    onDrop: (e: React.DragEvent) => {
      e.preventDefault();
      setter(false);
      const file = e.dataTransfer.files[0];
      if (file) fileSetter(file);
    },
  });

  const handleFileSelect = (
    e: React.ChangeEvent<HTMLInputElement>,
    setter: (f: File) => void
  ) => {
    const file = e.target.files?.[0];
    if (file) setter(file);
  };

  const hasResult = rows.length > 0;

  return (
    <div className="space-y-8">
      {/* Header */}
      <div>
        <h1 className="text-2xl font-bold text-zinc-800">📦 TR入仓数据整理</h1>
        <p className="mt-1 text-sm text-zinc-500">
          上传客户数据 + 供应商数据，自动匹配选数、校验报警，生成「出给客户」建议箱规
        </p>
      </div>

      {/* Upload Area */}
      {!hasResult && (
        <div className="space-y-6">
          <div className="grid grid-cols-1 gap-6 md:grid-cols-2">
            <UploadBox
              label="客户数据"
              description="自动识别 标准 / 易通发票 / 货箱清单 三种格式"
              file={customerFile}
              isDragging={isDraggingCustomer}
              dragHandlers={makeDragHandlers(setIsDraggingCustomer, setCustomerFile)}
              onClear={() => setCustomerFile(null)}
              inputRef={customerInputRef}
              onFileSelect={(e) => handleFileSelect(e, setCustomerFile)}
            />
            <UploadBox
              label="供应商数据"
              description="逐箱数据（自动识别 天图 / 英美入仓 / 给总部 三种格式）"
              file={supplierFile}
              isDragging={isDraggingSupplier}
              dragHandlers={makeDragHandlers(setIsDraggingSupplier, setSupplierFile)}
              onClear={() => setSupplierFile(null)}
              inputRef={supplierInputRef}
              onFileSelect={(e) => handleFileSelect(e, setSupplierFile)}
            />
          </div>

          <div className="flex justify-center">
            <button
              onClick={handleSubmit}
              disabled={!canSubmit}
              className={`inline-flex items-center gap-2 rounded-lg px-6 py-3 text-sm font-medium shadow-sm transition-colors ${
                canSubmit
                  ? "bg-emerald-600 text-white hover:bg-emerald-700"
                  : "cursor-not-allowed bg-zinc-200 text-zinc-400"
              }`}
            >
              {processing ? (
                <>
                  <div className="h-4 w-4 animate-spin rounded-full border-2 border-white border-t-transparent" />
                  正在整理中...
                </>
              ) : (
                <>开始整理</>
              )}
            </button>
          </div>

          <div className="rounded-lg border border-zinc-200 bg-white p-4 text-sm text-zinc-500">
            <p>💡 提示：系统按「FBA ID（供应商货箱编号前 12 位）」匹配箱组，自动选数并校验计费重，生成建议箱规。结果可手动微调后再导出。</p>
          </div>
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

      {/* Notice */}
      {notice && (
        <div className="flex items-center gap-3 rounded-lg border border-emerald-300 bg-emerald-50 p-4">
          <p className="text-sm text-emerald-700">{notice}</p>
        </div>
      )}

      {/* Results */}
      {hasResult && (
        <>
          {/* Action bar */}
          <div className="flex flex-wrap items-center justify-between gap-4 rounded-lg border border-zinc-200 bg-white p-4 shadow-sm">
            <div>
              <p className="text-sm text-zinc-500">
                客户 <span className="font-medium text-zinc-700">{summary?.customerCount}</span> 个产品
                {" · "}
                供应商 <span className="font-medium text-zinc-700">{summary?.boxCount}</span> 箱
                {" · "}
                匹配 <span className="font-medium text-zinc-700">{summary?.matchedCount}</span>
                {" · "}
                需复核 <span className="font-medium text-amber-600">{summary?.unmatchedCount}</span>
                {" · "}
                历史库 <span className="font-medium text-zinc-700">{historyCount}</span> 条
              </p>
            </div>
            <div className="flex flex-wrap gap-3">
              <button
                onClick={() => historyInputRef.current?.click()}
                className="inline-flex items-center gap-2 rounded-lg border border-zinc-300 bg-white px-4 py-2 text-sm font-medium text-zinc-600 hover:bg-zinc-50 transition-colors"
              >
                导入历史库
              </button>
              <button
                onClick={handleHistoryExport}
                className="inline-flex items-center gap-2 rounded-lg border border-zinc-300 bg-white px-4 py-2 text-sm font-medium text-zinc-600 hover:bg-zinc-50 transition-colors"
              >
                导出历史库
              </button>
              <button
                onClick={reset}
                className="inline-flex items-center gap-2 rounded-lg border border-zinc-300 bg-white px-4 py-2 text-sm font-medium text-zinc-600 hover:bg-zinc-50 transition-colors"
              >
                重新整理
              </button>
            </div>
          </div>

          {/* 全局校验条 */}
          <div
            className={`flex items-center justify-between gap-4 rounded-lg border p-4 ${
              globalWarn
                ? "border-red-300 bg-red-50"
                : "border-emerald-300 bg-emerald-50"
            }`}
          >
            <div className="flex items-center gap-3">
              <span className={`text-lg font-bold ${globalWarn ? "text-red-700" : "text-emerald-700"}`}>
                全局校验
              </span>
              <p className="text-sm">
                <span className="font-medium">出给客户总计费重 {fmt(customerTotal)}</span>
                {" − "}
                <span className="font-medium">供应商总计费重 {fmt(supplierTotal)}</span>
                {" = "}
                <span className={`font-bold ${globalWarn ? "text-red-700" : "text-emerald-700"}`}>
                  {globalWarn ? "≤ 0" : fmt(globalDiff)}
                </span>
              </p>
            </div>
            {globalWarn && (
              <span className="rounded-full bg-red-100 px-3 py-1 text-xs font-semibold text-red-700">
                ⚠ 导出时将自动等比例放大并标注 [全局调整]
              </span>
            )}
          </div>

          {/* 可编辑建议表格 */}
          <div className="rounded-xl border border-zinc-200 bg-white overflow-hidden">
            <div className="border-b border-zinc-200 px-5 py-3">
              <h2 className="text-sm font-semibold text-zinc-700">
                建议箱规（{rows.length} 个产品，可直接编辑长/宽/高/实重）
              </h2>
            </div>
            <div className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead>
                  <tr className="border-b border-zinc-200 bg-zinc-50 text-left text-xs font-semibold uppercase tracking-wider text-zinc-500">
                    <th className="px-3 py-2">FBA ID</th>
                    <th className="px-3 py-2">中文品名</th>
                    <th className="px-3 py-2 text-right">箱数</th>
                    <th className="px-3 py-2 text-center">客户箱规</th>
                    <th className="px-3 py-2 text-center">建议长(CM)</th>
                    <th className="px-3 py-2 text-center">建议宽(CM)</th>
                    <th className="px-3 py-2 text-center">建议高(CM)</th>
                    <th className="px-3 py-2 text-center">建议实重</th>
                    <th className="px-3 py-2 text-right">材积重</th>
                    <th className="px-3 py-2 text-right">计费重</th>
                    <th className="px-3 py-2 text-right">供应商计费重</th>
                    <th className="px-3 py-2 text-center">三边和差异</th>
                    <th className="px-3 py-2 text-center">选中名次</th>
                    <th className="px-3 py-2 text-right">历史最大</th>
                    <th className="px-3 py-2">报警</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-zinc-100">
                  {rows.map((r, i) => {
                    const supplierValid =
                      r.supplier.lengthCm > 0 ||
                      r.supplier.widthCm > 0 ||
                      r.supplier.heightCm > 0;
                    const supplierSumSides = round2(
                      r.supplier.lengthCm + r.supplier.widthCm + r.supplier.heightCm,
                    );
                    const sumSidesDiff = supplierValid
                      ? round2(supplierSumSides - r.customer.sumSides)
                      : null;
                    return (
                    <tr key={`${r.fbaId}-${r.productName}-${i}`} className="hover:bg-zinc-50">
                      <td className="px-3 py-2 font-mono text-xs text-zinc-600">{r.fbaId}</td>
                      <td className="px-3 py-2 text-zinc-700">{r.productName}</td>
                      <td className="px-3 py-2 text-right text-zinc-600">{r.totalBoxes}</td>
                      <td className="px-3 py-2 text-center text-xs text-zinc-500">
                        {fmt(r.customer.lengthCm, 1)}×{fmt(r.customer.widthCm, 1)}×{fmt(r.customer.heightCm, 1)}
                        <span className="text-zinc-400"> / {fmt(r.customer.chargeableWeight)}</span>
                      </td>
                      <td className="px-1 py-1 text-center">
                        <input
                          type="number"
                          step="1"
                          value={r.suggestion.lengthCm}
                          onChange={(e) => editSuggestion(i, "lengthCm", parseFloat(e.target.value) || 0)}
                          className="w-16 rounded border border-zinc-300 px-1 py-0.5 text-center text-sm focus:border-emerald-500 focus:outline-none"
                        />
                      </td>
                      <td className="px-1 py-1 text-center">
                        <input
                          type="number"
                          step="1"
                          value={r.suggestion.widthCm}
                          onChange={(e) => editSuggestion(i, "widthCm", parseFloat(e.target.value) || 0)}
                          className="w-16 rounded border border-zinc-300 px-1 py-0.5 text-center text-sm focus:border-emerald-500 focus:outline-none"
                        />
                      </td>
                      <td className="px-1 py-1 text-center">
                        <input
                          type="number"
                          step="1"
                          value={r.suggestion.heightCm}
                          onChange={(e) => editSuggestion(i, "heightCm", parseFloat(e.target.value) || 0)}
                          className="w-16 rounded border border-zinc-300 px-1 py-0.5 text-center text-sm focus:border-emerald-500 focus:outline-none"
                        />
                      </td>
                      <td className="px-1 py-1 text-center">
                        <input
                          type="number"
                          step="0.1"
                          value={r.suggestion.actualWeight}
                          onChange={(e) => editSuggestion(i, "actualWeight", parseFloat(e.target.value) || 0)}
                          className="w-16 rounded border border-zinc-300 px-1 py-0.5 text-center text-sm focus:border-emerald-500 focus:outline-none"
                        />
                      </td>
                      <td className="px-3 py-2 text-right font-mono text-xs text-zinc-600">
                        {fmt(r.suggestion.volumeWeight, 3)}
                      </td>
                      <td className="px-3 py-2 text-right font-mono text-xs font-bold text-zinc-800">
                        {fmt(r.suggestion.chargeableWeight)}
                      </td>
                      <td className="px-3 py-2 text-right font-mono text-xs text-zinc-500">
                        {fmt(r.supplierChargeable)}
                      </td>
                      <td className="px-3 py-2 text-center font-mono text-xs">
                        {sumSidesDiff === null ? (
                          <span className="text-zinc-400">—</span>
                        ) : (
                          <span
                            className={
                              Math.abs(sumSidesDiff) >= 6
                                ? "font-semibold text-red-600"
                                : "text-zinc-600"
                            }
                          >
                            {sumSidesDiff > 0 ? "+" : ""}
                            {fmt(sumSidesDiff)}
                          </span>
                        )}
                        {sumSidesDiff !== null && (
                          <div className="text-[10px] text-zinc-400">
                            供{fmt(supplierSumSides)} / 客{fmt(r.customer.sumSides)}
                          </div>
                        )}
                      </td>
                      <td className="px-3 py-2 text-center font-mono text-xs">
                        {r.pickedRank === null ? (
                          <span className="text-zinc-400">—</span>
                        ) : (
                          <span className={r.pickedRank === 1 ? "text-zinc-500" : "font-semibold text-amber-600"}>
                            第{r.pickedRank}大
                          </span>
                        )}
                      </td>
                      <td className="px-3 py-2 text-right font-mono text-xs text-zinc-500">
                        {r.historyMax ? fmt(r.historyMax.chargeableWeight) : "—"}
                      </td>
                      <td className="px-3 py-2">
                        {r.alarms.length === 0 ? (
                          <span className="text-xs text-zinc-400">正常</span>
                        ) : (
                          <div className="flex flex-wrap gap-1">
                            {r.alarms.map((a, ai) => (
                              <span
                                key={ai}
                                className={`inline-flex rounded-full px-2 py-0.5 text-xs font-semibold ${
                                  a.includes("需人工复核") || a.includes("需人工确认")
                                    ? "bg-amber-100 text-amber-700"
                                    : a.includes("建议参考历史")
                                      ? "bg-blue-100 text-blue-700"
                                      : a.includes("核查过机图")
                                        ? "bg-red-600 text-white"
                                        : "bg-red-100 text-red-700"
                                }`}
                              >
                                {a}
                              </span>
                            ))}
                          </div>
                        )}
                      </td>
                    </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>

          {/* 导出按钮 */}
          <div className="flex justify-center">
            <button
              onClick={handleExport}
              disabled={exporting}
              className="inline-flex items-center gap-2 rounded-lg bg-emerald-600 px-6 py-3 text-sm font-medium text-white shadow-sm hover:bg-emerald-700 transition-colors disabled:cursor-not-allowed disabled:bg-zinc-300"
            >
              {exporting ? (
                <>
                  <div className="h-4 w-4 animate-spin rounded-full border-2 border-white border-t-transparent" />
                  正在导出...
                </>
              ) : (
                <>确认导出三类数据的 Excel</>
              )}
            </button>
          </div>
        </>
      )}

      {/* 历史库隐藏文件输入 */}
      <input
        ref={historyInputRef}
        type="file"
        accept=".xlsx,.xls"
        onChange={handleHistoryImport}
        className="hidden"
      />
    </div>
  );
}

// ============================================================
// Sub-components
// ============================================================

function UploadBox({
  label,
  description,
  file,
  isDragging,
  dragHandlers,
  onClear,
  inputRef,
  onFileSelect,
}: {
  label: string;
  description: string;
  file: File | null;
  isDragging: boolean;
  dragHandlers: {
    onDragOver: (e: React.DragEvent) => void;
    onDragLeave: (e: React.DragEvent) => void;
    onDrop: (e: React.DragEvent) => void;
  };
  onClear: () => void;
  inputRef: React.RefObject<HTMLInputElement | null>;
  onFileSelect: (e: React.ChangeEvent<HTMLInputElement>) => void;
}) {
  return (
    <div
      {...dragHandlers}
      onClick={() => !file && inputRef.current?.click()}
      className={`
        flex cursor-pointer flex-col items-center justify-center rounded-xl border-2 border-dashed p-12 transition-colors
        ${
          file
            ? "border-zinc-300 bg-white"
            : isDragging
              ? "border-emerald-400 bg-emerald-50"
              : "border-zinc-300 bg-zinc-50 hover:border-zinc-400 hover:bg-zinc-100"
        }
      `}
    >
      {file ? (
        <div className="flex flex-col items-center gap-3 text-center">
          <svg className="h-10 w-10 text-emerald-500" fill="none" viewBox="0 0 24 24" stroke="currentColor">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.5} d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
          </svg>
          <div>
            <p className="text-sm font-semibold text-zinc-700">{label}</p>
            <p className="mt-0.5 text-xs text-zinc-500">{file.name}</p>
            <p className="text-xs text-zinc-400">{(file.size / 1024).toFixed(1)} KB</p>
          </div>
          <button
            onClick={(e) => {
              e.stopPropagation();
              onClear();
            }}
            className="text-xs font-medium text-red-500 hover:text-red-700"
          >
            移除
          </button>
        </div>
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
          <p className="text-lg font-medium text-zinc-600">{label}</p>
          <p className="mt-1 text-sm text-zinc-400">{description}</p>
          <p className="mt-2 text-sm text-zinc-400">拖拽 .xlsx 文件到此处，或点击上传</p>
        </>
      )}
      <input
        ref={inputRef}
        type="file"
        accept=".xlsx,.xls"
        onChange={onFileSelect}
        className="hidden"
      />
    </div>
  );
}
