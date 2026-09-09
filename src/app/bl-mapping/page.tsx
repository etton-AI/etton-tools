"use client";

import { useEffect, useState } from "react";

// ============================================================
// 港口 / 渠道映射编辑（写回 bl-service 的 port_map.json / channel_map.json）
// ============================================================

type Row = { key: string; value: string };
type Map = Record<string, string>;

interface MappingsResponse {
  ok: boolean;
  port_map: { origin: Map; destination: Map };
  channel_map: Map;
  customs_office_map: Map;
  error?: string;
}

function toRows(o: Map | undefined): Row[] {
  return Object.entries(o ?? {}).map(([key, value]) => ({ key, value }));
}

function toObj(rows: Row[]): Map {
  const out: Map = {};
  for (const r of rows) {
    const k = r.key.trim();
    if (k) out[k] = r.value.trim();
  }
  return out;
}

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
// MapTable —— 可增删改的键值对编辑器
// ============================================================

function MapTable({
  title,
  hint,
  rows,
  onChange,
  keyPlaceholder,
  valuePlaceholder,
  keyHeader = "中文港口 / 渠道",
  valueHeader = "提单英文（港口 / 国家）",
}: {
  title: string;
  hint: string;
  rows: Row[];
  onChange: (rows: Row[]) => void;
  keyPlaceholder: string;
  valuePlaceholder: string;
  keyHeader?: string;
  valueHeader?: string;
}) {
  const update = (i: number, field: "key" | "value", v: string) => {
    const next = rows.map((r, idx) => (idx === i ? { ...r, [field]: v } : r));
    onChange(next);
  };
  const remove = (i: number) => onChange(rows.filter((_, idx) => idx !== i));
  const add = () => onChange([...rows, { key: "", value: "" }]);

  return (
    <div className="rounded-xl border border-zinc-200 bg-white">
      <div className="flex items-center justify-between border-b border-zinc-200 px-5 py-3">
        <div>
          <h2 className="text-sm font-semibold text-zinc-700">{title}</h2>
          <p className="mt-0.5 text-xs text-zinc-400">{hint}</p>
        </div>
        <button
          onClick={add}
          className="inline-flex items-center gap-1 rounded-lg border border-zinc-300 bg-white px-3 py-1.5 text-xs font-medium text-zinc-600 transition-colors hover:bg-zinc-50"
        >
          + 添加一行
        </button>
      </div>
      <div className="overflow-x-auto p-5">
        <table className="w-full min-w-[480px] text-sm">
          <thead>
            <tr className="border-b border-zinc-200 text-left text-xs text-zinc-400">
              <th className="pb-2 pr-3 font-medium">{keyHeader}</th>
              <th className="pb-2 pr-3 font-medium">{valueHeader}</th>
              <th className="pb-2 w-16 font-medium"></th>
            </tr>
          </thead>
          <tbody>
            {rows.length === 0 && (
              <tr>
                <td colSpan={3} className="py-4 text-center text-xs text-zinc-400">
                  暂无映射，点「添加一行」新增
                </td>
              </tr>
            )}
            {rows.map((r, i) => (
              <tr key={i} className="border-b border-zinc-100 last:border-0">
                <td className="py-1.5 pr-3">
                  <input
                    value={r.key}
                    onChange={(e) => update(i, "key", e.target.value)}
                    placeholder={keyPlaceholder}
                    className="w-full rounded-lg border border-zinc-300 px-3 py-1.5 text-sm text-zinc-700 outline-none transition-colors focus:border-primary"
                  />
                </td>
                <td className="py-1.5 pr-3">
                  <input
                    value={r.value}
                    onChange={(e) => update(i, "value", e.target.value)}
                    placeholder={valuePlaceholder}
                    className="w-full rounded-lg border border-zinc-300 px-3 py-1.5 text-sm text-zinc-700 outline-none transition-colors focus:border-primary"
                  />
                </td>
                <td className="py-1.5 text-center">
                  <button
                    onClick={() => remove(i)}
                    title="删除此行"
                    className="rounded-md px-2 py-1 text-xs text-red-500 transition-colors hover:bg-red-50 hover:text-red-700"
                  >
                    删除
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ============================================================
// Page
// ============================================================

export default function BlMappingPage() {
  const [originRows, setOriginRows] = useState<Row[]>([]);
  const [destRows, setDestRows] = useState<Row[]>([]);
  const [channelRows, setChannelRows] = useState<Row[]>([]);
  const [customsRows, setCustomsRows] = useState<Row[]>([]);
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");
  const [notice, setNotice] = useState("");

  useEffect(() => {
    (async () => {
      setLoading(true);
      setError("");
      try {
        const res = await fetch("/api/bl/mappings");
        const data = await readJson<MappingsResponse>(res);
        if (!res.ok || !data.ok) {
          throw new Error(data.error || "加载失败");
        }
        setOriginRows(toRows(data.port_map?.origin));
        setDestRows(toRows(data.port_map?.destination));
        setChannelRows(toRows(data.channel_map));
        setCustomsRows(toRows(data.customs_office_map));
      } catch (e) {
        setError(e instanceof Error ? e.message : "加载失败");
      } finally {
        setLoading(false);
      }
    })();
  }, []);

  const save = async () => {
    setSaving(true);
    setError("");
    setNotice("");
    try {
      const res = await fetch("/api/bl/mappings", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          port_map: { origin: toObj(originRows), destination: toObj(destRows) },
          channel_map: toObj(channelRows),
          customs_office_map: toObj(customsRows),
        }),
      });
      const data = await readJson<{ ok: boolean; error?: string }>(res);
      if (!res.ok || !data.ok) {
        throw new Error(data.error || "保存失败");
      }
      setNotice("已保存，下次提单提取即生效");
    } catch (e) {
      setError(e instanceof Error ? e.message : "保存失败");
    } finally {
      setSaving(false);
    }
  };

  return (
    <div className="space-y-8">
      {/* Header */}
      <div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
        <div>
          <h1 className="text-2xl font-bold text-deep">🗺️ 港口 / 渠道映射设置</h1>
          <p className="mt-1 text-sm text-zinc-500">
            编辑报关底单中文港口 → 提单英文的对照关系，保存后提单提取立即生效
          </p>
        </div>
        <button
          onClick={save}
          disabled={saving || loading}
          className="inline-flex items-center gap-2 rounded-lg bg-primary px-5 py-2 text-sm font-semibold text-white transition-colors hover:opacity-90 disabled:opacity-50"
        >
          {saving && <span className="h-4 w-4 animate-spin rounded-full border-2 border-white border-t-transparent" />}
          {saving ? "保存中…" : "保存全部"}
        </button>
      </div>

      {loading ? (
        <div className="flex items-center gap-3 rounded-xl border border-zinc-200 bg-white px-5 py-8 text-sm text-zinc-500">
          <span className="h-4 w-4 animate-spin rounded-full border-2 border-primary border-t-transparent" />
          正在加载映射…
        </div>
      ) : (
        <>
          {error && (
            <div className="flex items-center gap-3 rounded-lg border border-red-300 bg-red-50 p-4">
              <p className="text-sm text-red-700">{error}</p>
            </div>
          )}
          {notice && (
            <div className="flex items-center gap-3 rounded-lg border border-emerald-300 bg-emerald-50 p-4">
              <p className="text-sm text-emerald-700">{notice}</p>
            </div>
          )}

          <MapTable
            title="起运港映射（离境口岸 → 提单起运港）"
            hint="报关底单「离境口岸」中文值 → 提单起运港英文。如：南沙港一期码头 → GUANGZHOU"
            rows={originRows}
            onChange={setOriginRows}
            keyPlaceholder="如：南沙港一期码头"
            valuePlaceholder="如：GUANGZHOU"
          />

          <MapTable
            title="目的港映射（运抵国 → 提单目的港）"
            hint="报关底单「运抵国」中文值 → 提单目的港英文。如：英国 → MALASZEWICZE、德国 → DUISBURG"
            rows={destRows}
            onChange={setDestRows}
            keyPlaceholder="如：英国"
            valuePlaceholder="如：MALASZEWICZE"
          />

          <MapTable
            title="渠道映射（客户渠道 → 提单目的港）"
            hint="箱货清单「客户渠道」→ 提单目的港英文（比运抵国映射更精确，命中时优先）。如：欧洲卡航包税-卡派 → GERMANY"
            rows={channelRows}
            onChange={setChannelRows}
            keyPlaceholder="如：欧洲卡航包税-卡派"
            valuePlaceholder="如：GERMANY"
          />

          <MapTable
            title="关区名映射（非海运 → 提单起运港）"
            hint="非海运（铁路/公路/航空）报关底单「海关编号」后的关区名 → 提单起运港英文。如：增城海关 → GUANGZHOU、蓉青关 → CHENGDU"
            rows={customsRows}
            onChange={setCustomsRows}
            keyPlaceholder="如：增城海关"
            valuePlaceholder="如：GUANGZHOU"
            keyHeader="关区名"
            valueHeader="提单起运港英文"
          />
        </>
      )}
    </div>
  );
}
