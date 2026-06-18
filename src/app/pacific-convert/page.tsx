"use client";

import { useState, useCallback, useRef } from "react";

export default function PacificConvertPage() {
  const [isDragging, setIsDragging] = useState(false);
  const [processing, setProcessing] = useState(false);
  const [fileName, setFileName] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFile = useCallback(async (file: File) => {
    if (!file.name.endsWith(".xlsx") && !file.name.endsWith(".xls")) {
      setError("只支持 .xlsx 或 .xls 格式的 Excel 文件");
      return;
    }

    setProcessing(true);
    setError(null);
    setFileName(file.name);

    try {
      const formData = new FormData();
      formData.append("file", file);

      const res = await fetch("/api/convert-pacific", {
        method: "POST",
        body: formData,
      });

      if (!res.ok) {
        const data = await res.json();
        setError(data.error || "转换失败");
        return;
      }

      // Trigger file download
      const blob = await res.blob();
      const disposition = res.headers.get("Content-Disposition");
      let downloadName = file.name.replace(/\.xlsx?$/i, "") + "_太平洋投保清单.xlsx";
      if (disposition) {
        const match = disposition.match(/filename\*?=(?:UTF-8'')?(.+)/i);
        if (match) downloadName = decodeURIComponent(match[1]);
      }

      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = downloadName;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    } catch {
      setError("网络错误，请重试");
    } finally {
      setProcessing(false);
    }
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
      const file = e.dataTransfer.files[0];
      if (file) handleFile(file);
    },
    [handleFile]
  );

  const onFileChange = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      const file = e.target.files?.[0];
      if (file) handleFile(file);
    },
    [handleFile]
  );

  const reset = useCallback(() => {
    setFileName(null);
    setError(null);
  }, []);

  return (
    <div className="space-y-8">
      {/* Header */}
      <div>
        <h1 className="text-2xl font-bold text-zinc-800">
          🚢 太平洋货箱清单转换
        </h1>
        <p className="mt-1 text-sm text-zinc-500">
          上传易通货箱清单 Excel，自动按模板公式转换为太平洋投保清单并下载
        </p>
      </div>

      {/* Mapping info */}
      <div className="rounded-lg border border-zinc-200 bg-zinc-50/50 p-4">
        <p className="text-xs font-semibold text-zinc-500 uppercase tracking-wide mb-2">
          转换映射
        </p>
        <div className="grid grid-cols-2 sm:grid-cols-3 gap-x-6 gap-y-1 text-sm text-zinc-600">
          <span>FBA ID → 入仓编号</span>
          <span>中文+英文品名 → DESCRIPTION</span>
          <span>申报总数量 → QTY PCS</span>
          <span>单个产品申报货值 → UNIT VALUE</span>
          <span>总申报货值 → TOTAL VALUE</span>
          <span>申报币种 → 币种</span>
          <span>总箱数 → CTNS</span>
          <span>单箱毛重 → G.W.(KG)</span>
          <span>长×宽×高/10⁶ → MEASUREMENT</span>
        </div>
      </div>

      {/* Upload Area */}
      <div
        onDragOver={onDragOver}
        onDragLeave={onDragLeave}
        onDrop={onDrop}
        onClick={() => !processing && fileInputRef.current?.click()}
        className={`
          flex cursor-pointer flex-col items-center justify-center rounded-xl border-2 border-dashed p-12 transition-colors
          ${
            isDragging
              ? "border-blue-400 bg-blue-50"
              : "border-zinc-300 bg-zinc-50 hover:border-zinc-400 hover:bg-zinc-100"
          }
          ${processing ? "pointer-events-none opacity-60" : ""}
        `}
      >
        {processing ? (
          <>
            <div className="mb-4 h-10 w-10 animate-spin rounded-full border-4 border-zinc-300 border-t-blue-600" />
            <p className="text-lg font-medium text-zinc-600">正在转换中...</p>
            <p className="mt-1 text-sm text-zinc-400">
              {fileName} — 读取数据、应用映射公式
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
              拖拽货箱清单 Excel 到此处，或点击上传
            </p>
            <p className="mt-1 text-sm text-zinc-400">
              支持 .xlsx / .xls 格式 · 自动匹配表头列
            </p>
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

      {/* Tip */}
      <div className="rounded-lg border border-zinc-200 bg-white p-4 text-sm text-zinc-500">
        <p>
          💡 提示：转换后的文件包含所有自动映射的字段（入仓编号、品名、数量、货值等），
          提单号、船名航次、柜号等字段需手动补充。
        </p>
      </div>
    </div>
  );
}
