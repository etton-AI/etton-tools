"use client";

import { useEffect, useRef, useState } from "react";

// ============================================================
// Types
// ============================================================

interface Ticket {
  folder: string;
  record: Record<string, string>;
  warning?: string;
}

interface ExtractBatchResponse {
  ok: boolean;
  tickets: Ticket[];
  bl_fields: string[];
  bl_header: Record<string, string>;
  telex_fields: string[];
  error?: string;
}

interface GenerateResponse {
  ok: boolean;
  zip: string;
  previews: { folder: string; bl: string; telex: string }[];
  error?: string;
}

interface FolderGroup {
  folder: string;
  pdfs: File[];
  packing: File | null;
}

interface CustomerOption {
  key: string;
  label: string;
}

// 解析响应 JSON；遇到非 JSON 响应（如网关/服务重启返回的 HTML "Internal Server Error"）时抛出友好错误，
// 避免 res.json() 抛「Unexpected token」这种难以理解的报错。
async function readJson<T>(res: Response): Promise<T> {
  const text = await res.text();
  try {
    return JSON.parse(text) as T;
  } catch {
    const hint =
      res.status === 502 || res.status === 504
        ? "后端服务暂时不可用（可能正在重启），请稍后重试"
        : `服务返回异常（HTTP ${res.status}），请稍后重试`;
    throw new Error(hint);
  }
}

// ============================================================
// UploadBox（复刻 warehouse-entry 的写法）
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
              ? "border-primary bg-primary/5"
              : "border-[#C7E0F0] bg-white/60 hover:border-primary hover:bg-primary/5"
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
        <div className="flex flex-col items-center gap-3 text-center">
          <svg className="h-10 w-10 text-primary" fill="none" viewBox="0 0 24 24" stroke="currentColor">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.5} d="M7 16a4 4 0 01-.88-7.9A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12" />
          </svg>
          <div>
            <p className="text-sm font-semibold text-zinc-700">{label}</p>
            <p className="mt-0.5 text-xs text-zinc-500">{description}</p>
          </div>
        </div>
      )}
    </div>
  );
}

// ============================================================
// Page
// ============================================================

export default function BlReviewPage() {
  const [folderGroups, setFolderGroups] = useState<FolderGroup[]>([]);
  const [rootFolder, setRootFolder] = useState<string>("");
  const [trackingFile, setTrackingFile] = useState<File | null>(null);
  const [trackingDragging, setTrackingDragging] = useState(false);
  const [weeklyPackingFile, setWeeklyPackingFile] = useState<File | null>(null);
  const [weeklyDragging, setWeeklyDragging] = useState(false);
  const [customer, setCustomer] = useState<string>("拓锐");
  const [customers, setCustomers] = useState<CustomerOption[]>([
    { key: "拓锐", label: "拓锐（广州拓锐科技有限公司）" },
  ]);
  const [tickets, setTickets] = useState<Ticket[]>([]);
  const [blFields, setBlFields] = useState<string[]>([]);
  const [blHeader, setBlHeader] = useState<Record<string, string>>({});
  const [telexFields, setTelexFields] = useState<string[]>([]);
  const [extracting, setExtracting] = useState(false);
  const [generating, setGenerating] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [notice, setNotice] = useState<string | null>(null);
  const [previews, setPreviews] = useState<GenerateResponse["previews"]>([]);
  const [zipUrl, setZipUrl] = useState<string | null>(null);
  const [previewFile, setPreviewFile] = useState<{ title: string; url: string } | null>(null);
  const folderInputRef = useRef<HTMLInputElement>(null);
  const trackingInputRef = useRef<HTMLInputElement>(null);
  const weeklyInputRef = useRef<HTMLInputElement>(null);

  // 页面加载时拉取客户列表（加客户只改后端 CUSTOMERS，前端下拉自动更新）
  useEffect(() => {
    fetch("/api/bl/customers")
      .then((r) => r.json())
      .then((d) => {
        if (d.ok && Array.isArray(d.customers) && d.customers.length) {
          setCustomers(d.customers);
          if (d.default) setCustomer(d.default);
        }
      })
      .catch(() => {});
  }, []);

  const makeDragHandlers = (
    setter: (v: boolean) => void,
    fileSetter: (f: File) => void,
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
      const f = e.dataTransfer.files[0];
      if (f) fileSetter(f);
    },
  });

  const handleTrackingSelect = (e: React.ChangeEvent<HTMLInputElement>) => {
    const f = e.target.files?.[0];
    if (f) setTrackingFile(f);
  };

  const handleWeeklySelect = (e: React.ChangeEvent<HTMLInputElement>) => {
    const f = e.target.files?.[0];
    if (f) setWeeklyPackingFile(f);
  };

  // 选择文件夹：按「直接父目录」分组，每个子文件夹 = 一票（底单 PDF 可多份 + 箱货清单 xlsx）
  const handleFolderSelect = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = Array.from(e.target.files || []);
    if (!files.length) return;
    // 大文件夹名 = webkitRelativePath 第一段（用户点开选择的顶层文件夹，如「7月第2周」），用于 ZIP 命名
    const firstRel = files[0].webkitRelativePath || "";
    setRootFolder(firstRel ? firstRel.split("/")[0] : "");
    const groups = new Map<string, FolderGroup>();
    for (const f of files) {
      const parts = f.webkitRelativePath ? f.webkitRelativePath.split("/") : [f.name];
      const folder = parts.length >= 2 ? parts[parts.length - 2] : f.name;
      if (!groups.has(folder)) groups.set(folder, { folder, pdfs: [], packing: null });
      const g = groups.get(folder)!;
      const name = f.name.toLowerCase();
      if (name.endsWith(".pdf")) {
        g.pdfs.push(f);
      } else if (name.endsWith(".xlsx") || name.endsWith(".xls")) {
        // 优先把文件名含「箱/货/清单/packing」的 xlsx 识别为箱货清单
        if (!g.packing || /箱|货|清单|packing/i.test(f.name)) g.packing = f;
      }
    }
    const list = Array.from(groups.values());
    setFolderGroups(list);
    setError(null);
    setNotice(
      `已识别 ${list.length} 票（文件夹），共 ${list.reduce((n, g) => n + g.pdfs.length, 0)} 份底单 PDF`,
    );
  };

  const reset = () => {
    setFolderGroups([]);
    setRootFolder("");
    setTrackingFile(null);
    setWeeklyPackingFile(null);
    setTickets([]);
    setBlFields([]);
    setBlHeader({});
    setTelexFields([]);
    setPreviews([]);
    setZipUrl(null);
    setPreviewFile(null);
    setError(null);
    setNotice(null);
  };

  // 批量提取字段
  const handleExtract = async () => {
    if (!folderGroups.length) return;
    setExtracting(true);
    setError(null);
    setNotice(null);
    try {
      const form = new FormData();
      folderGroups.forEach((g, i) => {
        form.append(`folder_${i}`, g.folder);
        g.pdfs.forEach((pdf, j) => form.append(`pdf_${i}_${j}`, pdf));
        if (g.packing) form.append(`packing_${i}`, g.packing);
      });
      if (trackingFile) form.append("tracking", trackingFile);
      if (weeklyPackingFile) form.append("weekly_packing", weeklyPackingFile);
      form.append("customer", customer);

      const res = await fetch("/api/bl/extract-batch", { method: "POST", body: form });
      const data: ExtractBatchResponse = await readJson(res);
      if (!data.ok) {
        setError(data.error || "提取失败");
        return;
      }
      setTickets(data.tickets);
      setBlFields(data.bl_fields);
      setBlHeader(data.bl_header);
      setTelexFields(data.telex_fields);
      const warnCount = data.tickets.filter((t) => t.warning).length;
      setNotice(
        `提取完成：${data.tickets.length} 票${
          warnCount ? `，其中 ${warnCount} 票有提醒（见各票下方）` : ""
        }，请核对并编辑字段`,
      );
    } catch (e) {
      setError(`请求失败：${e instanceof Error ? e.message : String(e)}`);
    } finally {
      setExtracting(false);
    }
  };

  // 更新某个字段
  const editField = (ri: number, key: string, value: string) => {
    setTickets((prev) => {
      const next = prev.map((t) => ({ ...t, record: { ...t.record } }));
      if (next[ri]) next[ri].record[key] = value;
      return next;
    });
  };

  // 标红：空白字段标红；含中文字段（申请日期除外）标红
  const highlight = (key: string, value: string | undefined) => {
    if (value == null || String(value).trim() === "") return true;
    if (key === "申请日期") return false;
    return /[一-鿿]/.test(String(value));
  };

  // 票/文件夹列只显示「BG」开始的部分（去掉「N月第N周-(易通报关资料）-」前缀），全名放 title 悬浮
  const shortFolder = (folder: string) => {
    const i = folder.indexOf("BG");
    return i >= 0 ? folder.slice(i) : folder;
  };

  // 生成并返回预览 + ZIP 下载地址
  const handleGenerate = async () => {
    if (!tickets.length) return;
    setGenerating(true);
    setError(null);
    setNotice(null);
    try {
      const res = await fetch("/api/bl/generate", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          customer,
          root_folder: rootFolder,
          tickets: tickets.map((t) => ({ folder: t.folder, record: t.record })),
        }),
      });
      const data: GenerateResponse = await readJson(res);
      if (!data.ok) {
        setError(data.error || "生成失败");
        return;
      }
      setPreviews(data.previews);
      setZipUrl(`/api/bl/file/${data.zip}`);
      setNotice(`生成成功：${data.previews.length} 票，请预览确认后下载 ZIP`);
    } catch (e) {
      setError(`请求失败：${e instanceof Error ? e.message : String(e)}`);
    } finally {
      setGenerating(false);
    }
  };

  const columns = [...blFields, ...telexFields].filter((c) => c !== "文件命名");
  const hasResult = tickets.length > 0;
  const generated = zipUrl != null;

  return (
    <div className="space-y-8">
      {/* Header */}
      <div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
        <div>
          <h1 className="text-2xl font-bold text-deep">📦 提单 + 电放保函（批量）</h1>
          <p className="mt-1 text-sm text-zinc-500">
            按文件夹批量上传 → 自动提取 → <span className="font-semibold text-primary">人工审核</span> → 生成提单 & 电放保函
          </p>
        </div>
        {hasResult && (
          <button
            onClick={reset}
            className="inline-flex items-center gap-2 rounded-lg border border-zinc-300 bg-white px-4 py-2 text-sm font-medium text-zinc-600 hover:bg-zinc-50 transition-colors"
          >
            重新上传
          </button>
        )}
      </div>

      {/* Upload Area */}
      {!hasResult && (
        <div className="space-y-6">
          {/* 客户选择 */}
          <div className="flex flex-wrap items-center gap-3 rounded-xl border border-zinc-200 bg-white px-5 py-4">
            <label htmlFor="customer" className="shrink-0 text-sm font-semibold text-zinc-700">
              客户
            </label>
            <select
              id="customer"
              value={customer}
              onChange={(e) => setCustomer(e.target.value)}
              className="rounded-lg border border-zinc-300 bg-white px-3 py-2 text-sm text-zinc-700 focus:border-primary focus:outline-none"
            >
              {customers.map((c) => (
                <option key={c.key} value={c.key}>
                  {c.label}
                </option>
              ))}
            </select>
            <p className="text-xs text-zinc-400">
              不同客户的提单/保函格式可能有差异，选择后按对应客户生成
            </p>
          </div>

          {/* 文件夹上传 */}
          <div className="rounded-xl border-2 border-dashed border-[#C7E0F0] bg-white/60 p-12">
            <div className="flex flex-col items-center gap-4 text-center">
              <svg className="h-12 w-12 text-primary" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.5} d="M3 7a2 2 0 012-2h4l2 2h8a2 2 0 012 2v8a2 2 0 01-2 2H5a2 2 0 01-2-2V7z" />
              </svg>
              <div>
                <p className="text-sm font-semibold text-zinc-700">选择文件夹（批量）</p>
                <p className="mt-1 text-xs text-zinc-500">
                  每个子文件夹 = 一票：同一票的报关底单 PDF（可多份拆分）+ 箱货清单 xlsx 放同一文件夹
                </p>
              </div>
              <button
                onClick={() => folderInputRef.current?.click()}
                className="inline-flex items-center gap-2 rounded-lg bg-primary px-6 py-3 text-sm font-medium text-white shadow-sm hover:bg-primary-dark transition-colors"
              >
                📁 点此选择文件夹
              </button>
              <input
                ref={folderInputRef}
                type="file"
                multiple
                {...({ webkitdirectory: "", directory: "" } as React.InputHTMLAttributes<HTMLInputElement>)}
                className="hidden"
                onChange={handleFolderSelect}
              />
            </div>
          </div>

          {/* 已识别票列表 */}
          {folderGroups.length > 0 && (
            <div className="rounded-xl border border-zinc-200 bg-white">
              <div className="border-b border-zinc-200 px-5 py-3">
                <h2 className="text-sm font-semibold text-zinc-700">
                  🗂️ 已识别 {folderGroups.length} 票
                </h2>
              </div>
              <ul className="divide-y divide-zinc-100">
                {folderGroups.map((g, i) => {
                  const noPdf = g.pdfs.length === 0;
                  return (
                    <li
                      key={i}
                      className={`flex items-center justify-between gap-4 px-5 py-3 ${noPdf ? "bg-red-50" : ""}`}
                    >
                      <span
                        className={`truncate text-sm ${noPdf ? "font-semibold text-red-600" : "text-zinc-700"}`}
                        title={g.folder}
                      >
                        {g.folder}
                      </span>
                      <span
                        className={`shrink-0 text-xs ${noPdf ? "font-semibold text-red-600" : "text-zinc-500"}`}
                      >
                        {noPdf ? "⚠ 0 份底单" : `${g.pdfs.length} 份底单`}
                        {g.packing ? " · 有箱货清单" : " · 无箱货清单"}
                      </span>
                    </li>
                  );
                })}
              </ul>
            </div>
          )}

          {/* 周汇总箱货清单（整周一份，可选） */}
          <UploadBox
            label="周汇总箱货清单 xlsx（可选，整周一份）"
            description="按周导出的整份箱货清单，自动匹配每票底单；匹配不到回退文件夹内单票清单"
            file={weeklyPackingFile}
            isDragging={weeklyDragging}
            dragHandlers={makeDragHandlers(setWeeklyDragging, setWeeklyPackingFile)}
            onClear={() => setWeeklyPackingFile(null)}
            inputRef={weeklyInputRef}
            onFileSelect={handleWeeklySelect}
          />
          <input
            ref={weeklyInputRef}
            type="file"
            accept=".xlsx"
            className="hidden"
            onChange={handleWeeklySelect}
          />

          {/* 物流追踪表（单独上传） */}
          <UploadBox
            label="物流追踪表 xlsx（可选，单独上传）"
            description="用 FBA ID 匹配 ETD 开船日，自动填「起运日期」"
            file={trackingFile}
            isDragging={trackingDragging}
            dragHandlers={makeDragHandlers(setTrackingDragging, setTrackingFile)}
            onClear={() => setTrackingFile(null)}
            inputRef={trackingInputRef}
            onFileSelect={handleTrackingSelect}
          />
          <input
            ref={trackingInputRef}
            type="file"
            accept=".xlsx"
            className="hidden"
            onChange={handleTrackingSelect}
          />

          <div className="flex justify-center">
            <button
              onClick={handleExtract}
              disabled={!folderGroups.length || extracting}
              className={`inline-flex items-center gap-2 rounded-lg px-6 py-3 text-sm font-medium shadow-sm transition-colors ${
                folderGroups.length && !extracting
                  ? "bg-primary text-white hover:bg-primary-dark"
                  : "cursor-not-allowed bg-zinc-200 text-zinc-400"
              }`}
            >
              {extracting ? (
                <>
                  <div className="h-4 w-4 animate-spin rounded-full border-2 border-white border-t-transparent" />
                  正在批量提取字段...
                </>
              ) : (
                <>提取字段</>
              )}
            </button>
          </div>

          <div className="rounded-lg border border-zinc-200 bg-white p-4 text-sm text-zinc-500">
            <p>
              💡 提示：① 同一票的底单 + 箱货清单放同一文件夹；② 底单被拆成多份（报/放/委托）会自动合并；③ 空白或含中文的字段会标红提醒（申请日期除外）。
            </p>
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
        </div>
      )}

      {/* Notice */}
      {notice && (
        <div className="flex items-center gap-3 rounded-lg border border-emerald-300 bg-emerald-50 p-4">
          <p className="text-sm text-emerald-700">{notice}</p>
        </div>
      )}

      {/* Review Table */}
      {hasResult && (
        <>
          <div className="rounded-xl border border-zinc-200 bg-white">
            <div className="flex items-center justify-between border-b border-zinc-200 px-5 py-3">
              <h2 className="text-sm font-semibold text-zinc-700">✏️ 请核对并编辑提取结果</h2>
              <span className="flex items-center gap-2 text-xs text-zinc-400">
                <span className="inline-block h-3 w-3 rounded border border-red-400 bg-red-50" />
                空白 / 含中文需人工改成英文
              </span>
            </div>
            <div className="overflow-x-auto">
              <table className="w-full min-w-max text-sm">
                <thead>
                  <tr className="border-b border-zinc-200 bg-zinc-50 text-left text-xs font-semibold uppercase tracking-wider text-zinc-500">
                    <th className="sticky left-0 z-10 w-[180px] bg-zinc-50 px-3 py-2 whitespace-nowrap">票 / 文件夹</th>
                    {columns.map((c) => (
                      <th
                        key={c}
                        className={`whitespace-nowrap px-3 py-2 ${
                          c === "系统SO" ? "sticky left-[180px] z-10 bg-zinc-50" : ""
                        }`}
                      >
                        {blHeader[c] || c}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody className="divide-y divide-zinc-100">
                  {tickets.map((t, ri) => (
                    <tr key={ri} className="hover:bg-zinc-50">
                      <td className="sticky left-0 z-10 w-[180px] bg-white px-3 py-1 align-top">
                        <div className="w-full">
                          <p className="whitespace-normal break-all text-xs font-medium leading-snug text-zinc-700" title={t.folder}>
                            {shortFolder(t.folder)}
                          </p>
                          {t.warning && (
                            <p className="mt-1 whitespace-normal break-all text-xs leading-snug text-amber-600" title={t.warning}>
                              ⚠️ {t.warning}
                            </p>
                          )}
                        </div>
                      </td>
                      {columns.map((c) => {
                        const val = t.record[c] ?? "";
                        const hl = highlight(c, val);
                        return (
                          <td
                            key={c}
                            className={`px-2 py-1 ${
                              c === "系统SO" ? "sticky left-[180px] z-10 bg-white" : ""
                            }`}
                          >
                            {c === "系统SO" ? (
                              <div className="whitespace-pre-wrap px-2 py-1 text-sm text-zinc-700">
                                {val}
                              </div>
                            ) : c === "品名" ? (
                              <textarea
                                value={val}
                                onChange={(e) => editField(ri, c, e.target.value)}
                                rows={3}
                                className={`w-full min-w-[200px] rounded border px-2 py-1 text-sm focus:border-primary focus:outline-none ${
                                  hl ? "border-red-400 bg-red-50" : "border-zinc-300"
                                }`}
                              />
                            ) : (
                              <input
                                type="text"
                                value={val}
                                onChange={(e) => editField(ri, c, e.target.value)}
                                className={`w-full min-w-[120px] rounded border px-2 py-1 text-sm focus:border-primary focus:outline-none ${
                                  hl ? "border-red-400 bg-red-50" : "border-zinc-300"
                                }`}
                              />
                            )}
                          </td>
                        );
                      })}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>

          <div className="flex justify-end gap-3">
            <button
              onClick={reset}
              className="inline-flex items-center gap-2 rounded-lg border border-zinc-300 bg-white px-5 py-2.5 text-sm font-medium text-zinc-600 hover:bg-zinc-50 transition-colors"
            >
              重新上传
            </button>
            <button
              onClick={handleGenerate}
              disabled={generating}
              className={`inline-flex items-center gap-2 rounded-lg px-6 py-2.5 text-sm font-medium text-white transition-colors ${
                generating ? "cursor-not-allowed bg-zinc-300" : "bg-primary hover:bg-primary-dark"
              }`}
            >
              {generating ? (
                <>
                  <div className="h-4 w-4 animate-spin rounded-full border-2 border-white border-t-transparent" />
                  正在生成...
                </>
              ) : (
                <>✓ 确认生成 提单 + 保函</>
              )}
            </button>
          </div>
        </>
      )}

      {/* Preview + Download */}
      {generated && (
        <div className="rounded-xl border border-zinc-200 bg-white">
          <div className="flex items-center justify-between border-b border-zinc-200 px-5 py-3">
            <h2 className="text-sm font-semibold text-zinc-700">👁️ 预览 & 下载</h2>
            <a
              href={zipUrl || "#"}
              download
              className="inline-flex items-center gap-2 rounded-lg bg-primary px-5 py-2 text-sm font-medium text-white shadow-sm hover:bg-primary-dark transition-colors"
            >
              ⬇️ 下载 ZIP（底单 + 提单 + 保函）
            </a>
          </div>
          <ul className="divide-y divide-zinc-100">
            {previews.map((p, i) => (
              <li key={i} className="flex flex-wrap items-center justify-between gap-3 px-5 py-3">
                <span className="min-w-0 flex-1 truncate text-sm text-zinc-700" title={p.folder}>
                  {p.folder}
                </span>
                <div className="flex shrink-0 items-center gap-2">
                  <button
                    onClick={() =>
                      setPreviewFile({
                        title: `提单 - ${p.folder}`,
                        url: `/api/bl/file/${p.bl}`,
                      })
                    }
                    className="rounded-lg border border-zinc-300 bg-white px-3 py-1.5 text-xs font-medium text-zinc-600 hover:bg-zinc-50 transition-colors"
                  >
                    👁️ 提单
                  </button>
                  <button
                    onClick={() =>
                      setPreviewFile({
                        title: `电放保函 - ${p.folder}`,
                        url: `/api/bl/file/${p.telex}`,
                      })
                    }
                    className="rounded-lg border border-zinc-300 bg-white px-3 py-1.5 text-xs font-medium text-zinc-600 hover:bg-zinc-50 transition-colors"
                  >
                    👁️ 保函
                  </button>
                </div>
              </li>
            ))}
          </ul>
        </div>
      )}

      {/* Preview Modal */}
      {previewFile && (
        <div
          className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 p-4"
          onClick={() => setPreviewFile(null)}
        >
          <div
            className="flex h-[90vh] w-[90vw] flex-col rounded-xl bg-white"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="flex items-center justify-between border-b border-zinc-200 px-5 py-3">
              <h3 className="truncate text-sm font-semibold text-zinc-700">{previewFile.title}</h3>
              <button
                onClick={() => setPreviewFile(null)}
                className="rounded-lg border border-zinc-300 px-3 py-1.5 text-xs font-medium text-zinc-600 hover:bg-zinc-50 transition-colors"
              >
                关闭
              </button>
            </div>
            <iframe src={previewFile.url} className="min-h-0 flex-1 w-full" title={previewFile.title} />
          </div>
        </div>
      )}
    </div>
  );
}
