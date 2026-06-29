"use client";

import { useState, useEffect } from "react";
import Link from "next/link";

interface PriceEntry {
  supplier: string;
  channel_name: string;
  vessel_config: string;
  delivery_method: string;
  destination_code: string;
  destination_region: string;
  origin_region: string;
  origin_cities: string[];
  billing_type: string;
  min_quantity: string;
  unit_price: number;
  price_unit: string;
  transit_time_min: number | null;
  transit_time_max: number | null;
  transit_time_desc: string;
  claim_rule: string;
  effective_date: string;
  source_file: string;
}

interface QueryResult {
  success: boolean;
  results: PriceEntry[];
  total: number;
  best: PriceEntry | null;
  stats: { total: number; generated_at: string } | null;
  error?: string;
}

// 常用仓库列表
const POPULAR_WAREHOUSES = [
  "ONT8", "LGB8", "LAX9", "SBD1", "SMF3", "SCK4", "LAS1",
  "FTW1", "DFW6", "IAH3", "MEM1", "MDW2", "IND9",
  "ABE8", "TEB9", "AVP1", "RDU2", "CLT2",
];

// 常用城市
const POPULAR_CITIES = ["深圳", "东莞", "广州", "义乌", "上海", "宁波", "厦门", "泉州", "武汉"];

// 常用船司
const POPULAR_VESSELS = ["美森", "CLX", "EXX", "以星", "COSCO", "OA", "普船", "合德"];

// 送仓方式
const DELIVERY_METHODS = ["", "卡派", "海派", "整柜直送", "自提"];

// 供应商
const SUPPLIERS = ["", "易通ETTON", "天图通逊", "英美跨境"];

// 供应商颜色标签
function supplierBadge(s: string) {
  if (s.includes("易通")) return { bg: "bg-blue-100", text: "text-blue-700", border: "border-blue-300" };
  if (s.includes("天图")) return { bg: "bg-green-100", text: "text-green-700", border: "border-green-300" };
  if (s.includes("英美")) return { bg: "bg-purple-100", text: "text-purple-700", border: "border-purple-300" };
  return { bg: "bg-gray-100", text: "text-gray-700", border: "border-gray-300" };
}

export default function PriceQueryPage() {
  const [dest, setDest] = useState("");
  const [origin, setOrigin] = useState("");
  const [weight, setWeight] = useState("");
  const [vessel, setVessel] = useState("");
  const [method, setMethod] = useState("");
  const [supplier, setSupplier] = useState("");
  const [limit, setLimit] = useState("30");

  const [loading, setLoading] = useState(false);
  const [results, setResults] = useState<PriceEntry[]>([]);
  const [best, setBest] = useState<PriceEntry | null>(null);
  const [total, setTotal] = useState(0);
  const [error, setError] = useState("");
  const [stats, setStats] = useState<{ total: number; generated_at: string } | null>(null);
  const [showExport, setShowExport] = useState(false);
  const [exportData, setExportData] = useState("");

  // 首次加载时获取数据统计
  useEffect(() => {
    fetch("/api/price-query?top=1")
      .then((r) => r.json())
      .then((d: QueryResult) => {
        if (d.stats) setStats(d.stats);
      })
      .catch(() => {});
  }, []);

  const handleSearch = async () => {
    if (!dest.trim() && !origin.trim() && !vessel.trim()) {
      setError("请至少输入目的仓、发货城市或船司关键词");
      return;
    }
    setError("");
    setLoading(true);
    try {
      const params = new URLSearchParams();
      if (dest.trim()) params.set("dest", dest.trim());
      if (origin.trim()) params.set("origin", origin.trim());
      if (weight.trim()) params.set("weight", weight.trim());
      if (vessel.trim()) params.set("vessel", vessel.trim());
      if (method) params.set("method", method);
      if (supplier) params.set("supplier", supplier);
      params.set("top", limit);

      const resp = await fetch(`/api/price-query?${params.toString()}`);
      const data: QueryResult = await resp.json();
      if (data.success) {
        setResults(data.results);
        setBest(data.best);
        setTotal(data.total);
        if (data.stats) setStats(data.stats);
      } else {
        setError(data.error || "查询失败");
        setResults([]);
        setBest(null);
        setTotal(0);
      }
    } catch (e: unknown) {
      setError("网络错误: " + (e instanceof Error ? e.message : ""));
    }
    setLoading(false);
  };

  const handleExportCSV = () => {
    if (results.length === 0) return;
    const headers = ["供应商", "渠道名", "船配置", "送仓方式", "目的仓", "发货仓", "计费方式", "起收量", "单价", "价格单位", "时效", "赔付规则", "生效日期"];
    const keys: (keyof PriceEntry)[] = ["supplier", "channel_name", "vessel_config", "delivery_method", "destination_code", "origin_region", "billing_type", "min_quantity", "unit_price", "price_unit", "transit_time_desc", "claim_rule", "effective_date"];
    const lines = [headers.join(",")];
    for (const r of results) {
      lines.push(keys.map((k) => `"${String(r[k] ?? "").replace(/"/g, '""')}"`).join(","));
    }
    const csv = lines.join("\n");
    const blob = new Blob(["﻿" + csv], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `比价查询_${dest || "all"}_${new Date().toISOString().slice(0, 10)}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const handleExportJSON = () => {
    const json = JSON.stringify({ query: { dest, origin, weight, vessel, method, supplier }, results, best, total }, null, 2);
    setExportData(json);
    setShowExport(true);
  };

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === "Enter") handleSearch();
  };

  return (
    <div className="min-h-screen bg-gray-50">
      {/* 顶部横幅 */}
      <div className="bg-gradient-to-r from-blue-700 to-blue-900 text-white">
        <div className="max-w-7xl mx-auto px-4 py-6">
          <div className="flex items-center justify-between">
            <div>
              <h1 className="text-2xl font-bold">🚢 美线FBA海运比价查询</h1>
              <p className="text-blue-200 text-sm mt-1">
                覆盖 ETTON易通 · 天图通逊 · 英美跨境 三家供应商
                {stats && (
                  <span className="ml-2 text-blue-300">
                    | {stats.total.toLocaleString()} 条价格数据 | 更新于 {stats.generated_at?.slice(0, 10)}
                  </span>
                )}
              </p>
            </div>
            <Link href="/" className="text-blue-200 hover:text-white text-sm underline">← 返回首页</Link>
          </div>
        </div>
      </div>

      <div className="max-w-7xl mx-auto px-4 py-6">
        {/* 查询表单 */}
        <div className="bg-white rounded-xl shadow-sm border border-gray-200 p-6 mb-6">
          <div className="grid grid-cols-2 md:grid-cols-4 lg:grid-cols-8 gap-3">
            {/* 目的仓 */}
            <div className="col-span-2 lg:col-span-1">
              <label className="block text-xs font-medium text-gray-600 mb-1">目的仓 *</label>
              <input
                type="text"
                value={dest}
                onChange={(e) => setDest(e.target.value.toUpperCase())}
                onKeyDown={handleKeyDown}
                placeholder="ONT8"
                list="warehouse-list"
                className="w-full px-3 py-2 border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none"
              />
              <datalist id="warehouse-list">
                {POPULAR_WAREHOUSES.map((w) => (<option key={w} value={w} />))}
              </datalist>
            </div>

            {/* 发货城市 */}
            <div className="col-span-2 lg:col-span-1">
              <label className="block text-xs font-medium text-gray-600 mb-1">发货城市</label>
              <input
                type="text"
                value={origin}
                onChange={(e) => setOrigin(e.target.value)}
                onKeyDown={handleKeyDown}
                placeholder="深圳"
                list="city-list"
                className="w-full px-3 py-2 border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none"
              />
              <datalist id="city-list">
                {POPULAR_CITIES.map((c) => (<option key={c} value={c} />))}
              </datalist>
            </div>

            {/* 重量 */}
            <div className="lg:col-span-1">
              <label className="block text-xs font-medium text-gray-600 mb-1">重量(KG)</label>
              <input
                type="number"
                value={weight}
                onChange={(e) => setWeight(e.target.value)}
                onKeyDown={handleKeyDown}
                placeholder="100"
                className="w-full px-3 py-2 border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none"
              />
            </div>

            {/* 船司 */}
            <div className="col-span-2 lg:col-span-1">
              <label className="block text-xs font-medium text-gray-600 mb-1">船司</label>
              <input
                type="text"
                value={vessel}
                onChange={(e) => setVessel(e.target.value)}
                onKeyDown={handleKeyDown}
                placeholder="EXX/美森"
                list="vessel-list"
                className="w-full px-3 py-2 border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none"
              />
              <datalist id="vessel-list">
                {POPULAR_VESSELS.map((v) => (<option key={v} value={v} />))}
              </datalist>
            </div>

            {/* 送仓方式 */}
            <div className="lg:col-span-1">
              <label className="block text-xs font-medium text-gray-600 mb-1">送仓方式</label>
              <select
                value={method}
                onChange={(e) => setMethod(e.target.value)}
                className="w-full px-3 py-2 border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none bg-white"
              >
                {DELIVERY_METHODS.map((m) => (<option key={m} value={m}>{m || "全部"}</option>))}
              </select>
            </div>

            {/* 供应商 */}
            <div className="lg:col-span-1">
              <label className="block text-xs font-medium text-gray-600 mb-1">供应商</label>
              <select
                value={supplier}
                onChange={(e) => setSupplier(e.target.value)}
                className="w-full px-3 py-2 border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none bg-white"
              >
                {SUPPLIERS.map((s) => (<option key={s} value={s}>{s || "全部"}</option>))}
              </select>
            </div>

            {/* 显示条数 + 搜索按钮 */}
            <div className="lg:col-span-1">
              <label className="block text-xs font-medium text-gray-600 mb-1">显示条数</label>
              <select
                value={limit}
                onChange={(e) => setLimit(e.target.value)}
                className="w-full px-3 py-2 border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none bg-white"
              >
                <option value="10">10条</option>
                <option value="30">30条</option>
                <option value="50">50条</option>
                <option value="100">100条</option>
                <option value="0">全部</option>
              </select>
            </div>
          </div>

          {/* 操作按钮行 */}
          <div className="flex items-center gap-3 mt-4">
            <button
              onClick={handleSearch}
              disabled={loading}
              className="px-6 py-2.5 bg-blue-600 text-white rounded-lg font-medium hover:bg-blue-700 disabled:opacity-50 disabled:cursor-not-allowed text-sm transition-colors"
            >
              {loading ? "🔍 查询中..." : "🔍 查询"}
            </button>
            {results.length > 0 && (
              <>
                <button onClick={handleExportCSV} className="px-4 py-2 text-sm text-gray-600 bg-white border border-gray-300 rounded-lg hover:bg-gray-50 transition-colors">
                  📥 导出CSV
                </button>
                <button onClick={handleExportJSON} className="px-4 py-2 text-sm text-gray-600 bg-white border border-gray-300 rounded-lg hover:bg-gray-50 transition-colors">
                  📋 导出JSON
                </button>
              </>
            )}
            {error && <span className="text-red-500 text-sm">{error}</span>}
          </div>

          {/* 快捷预设 */}
          <div className="mt-3 flex flex-wrap gap-2">
            <span className="text-xs text-gray-400 pt-1">快捷:</span>
            <button onClick={() => { setDest("ONT8"); setOrigin("深圳"); setWeight("100"); setVessel("EXX"); setMethod("卡派"); }} className="text-xs px-2 py-1 bg-gray-100 hover:bg-gray-200 rounded text-gray-600 transition-colors">ONT8+深圳+EXX</button>
            <button onClick={() => { setDest("ONT8"); setOrigin("深圳"); setVessel("美森"); setMethod("卡派"); setWeight(""); }} className="text-xs px-2 py-1 bg-gray-100 hover:bg-gray-200 rounded text-gray-600 transition-colors">ONT8+美森卡派</button>
            <button onClick={() => { setDest("LAX9"); setOrigin("义乌"); setVessel("美森"); setMethod("卡派"); }} className="text-xs px-2 py-1 bg-gray-100 hover:bg-gray-200 rounded text-gray-600 transition-colors">LAX9+义乌+美森</button>
            <button onClick={() => { setDest("ONT8"); setOrigin("深圳"); setMethod(""); setVessel(""); setWeight(""); }} className="text-xs px-2 py-1 bg-gray-100 hover:bg-gray-200 rounded text-gray-600 transition-colors">ONT8所有渠道</button>
          </div>
        </div>

        {/* 最优推荐 */}
        {best && (
          <div className="bg-gradient-to-r from-yellow-50 to-orange-50 border border-yellow-200 rounded-xl p-5 mb-6">
            <div className="flex items-center gap-3">
              <span className="text-3xl">🏆</span>
              <div className="flex-1">
                <span className="text-sm text-yellow-700 font-medium">最优选择</span>
                <h3 className="text-lg font-bold text-gray-900">
                  {best.supplier} — {best.channel_name}
                </h3>
                <div className="flex flex-wrap gap-3 mt-1 text-sm text-gray-600">
                  <span>💰 <strong>{best.unit_price} {best.price_unit}</strong></span>
                  <span>⏱ {best.transit_time_desc || `${best.transit_time_min}-${best.transit_time_max}天`}</span>
                  <span>🚢 {best.vessel_config}</span>
                  <span>📦 {best.origin_region}</span>
                  {best.claim_rule && <span>📋 {best.claim_rule}</span>}
                </div>
              </div>
              <div className={`px-3 py-1 rounded-full text-xs font-medium ${supplierBadge(best.supplier).bg} ${supplierBadge(best.supplier).text} border ${supplierBadge(best.supplier).border}`}>
                最低价
              </div>
            </div>
          </div>
        )}

        {/* 统计行 */}
        {total > 0 && (
          <div className="text-sm text-gray-500 mb-3">
            共找到 <strong className="text-gray-700">{total}</strong> 条匹配记录
            {results.length < total && (
              <span>（当前显示前 {results.length} 条，可调整「显示条数」查看更多）</span>
            )}
          </div>
        )}

        {/* 结果表格 */}
        {results.length > 0 && (
          <div className="bg-white rounded-xl shadow-sm border border-gray-200 overflow-hidden">
            <div className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead>
                  <tr className="bg-gray-50 border-b border-gray-200">
                    <th className="text-left px-4 py-3 font-medium text-gray-600 w-10">#</th>
                    <th className="text-left px-4 py-3 font-medium text-gray-600">供应商</th>
                    <th className="text-left px-4 py-3 font-medium text-gray-600 min-w-[180px]">渠道名</th>
                    <th className="text-left px-4 py-3 font-medium text-gray-600 min-w-[160px]">船配置</th>
                    <th className="text-right px-4 py-3 font-medium text-gray-600">单价</th>
                    <th className="text-center px-4 py-3 font-medium text-gray-600">时效</th>
                    <th className="text-left px-4 py-3 font-medium text-gray-600 min-w-[160px]">赔付规则</th>
                    <th className="text-left px-4 py-3 font-medium text-gray-600">发货仓</th>
                    <th className="text-center px-4 py-3 font-medium text-gray-600">生效日期</th>
                  </tr>
                </thead>
                <tbody>
                  {results.map((r, i) => {
                    const badge = supplierBadge(r.supplier);
                    const isBest = best && r.supplier === best.supplier && r.channel_name === best.channel_name;
                    return (
                      <tr key={i} className={`border-b border-gray-100 hover:bg-gray-50 transition-colors ${isBest ? "bg-yellow-50/50" : ""}`}>
                        <td className="px-4 py-3 text-gray-400">{i + 1}</td>
                        <td className="px-4 py-3">
                          <span className={`inline-block px-2 py-0.5 rounded text-xs font-medium ${badge.bg} ${badge.text} border ${badge.border}`}>
                            {r.supplier}
                          </span>
                        </td>
                        <td className="px-4 py-3 font-medium text-gray-800">{r.channel_name}</td>
                        <td className="px-4 py-3 text-gray-600">{r.vessel_config || "-"}</td>
                        <td className="px-4 py-3 text-right">
                          <span className={`font-bold ${isBest ? "text-green-600" : "text-gray-900"}`}>
                            {r.unit_price}
                          </span>
                          <span className="text-gray-400 ml-1">{r.price_unit}</span>
                        </td>
                        <td className="px-4 py-3 text-center text-gray-600">
                          {r.transit_time_min ? `${r.transit_time_min}-${r.transit_time_max}天` : r.transit_time_desc || "-"}
                        </td>
                        <td className="px-4 py-3 text-gray-500 text-xs leading-relaxed max-w-[200px] truncate" title={r.claim_rule}>
                          {r.claim_rule || "-"}
                        </td>
                        <td className="px-4 py-3 text-gray-600 text-xs">{r.origin_region}</td>
                        <td className="px-4 py-3 text-center text-gray-400 text-xs">{r.effective_date || "-"}</td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* 无结果提示 */}
        {!loading && total === 0 && !error && (
          <div className="bg-white rounded-xl shadow-sm border border-gray-200 p-12 text-center">
            <p className="text-gray-400 text-lg">🔍 输入条件后点击查询</p>
            <p className="text-gray-300 text-sm mt-2">例如：ONT8 + 深圳 + 100KG + EXX</p>
          </div>
        )}

        {/* 加载状态 */}
        {loading && (
          <div className="bg-white rounded-xl shadow-sm border border-gray-200 p-12 text-center">
            <p className="text-gray-400">⏳ 查询中...</p>
          </div>
        )}

        {/* JSON导出弹窗 */}
        {showExport && (
          <div className="fixed inset-0 bg-black/50 z-50 flex items-center justify-center" onClick={() => setShowExport(false)}>
            <div className="bg-white rounded-xl shadow-2xl max-w-3xl w-full mx-4 max-h-[80vh] flex flex-col" onClick={(e) => e.stopPropagation()}>
              <div className="flex items-center justify-between px-6 py-4 border-b border-gray-200">
                <h3 className="font-semibold text-gray-800">JSON 导出 ({results.length} 条)</h3>
                <button onClick={() => setShowExport(false)} className="text-gray-400 hover:text-gray-600 text-xl leading-none">&times;</button>
              </div>
              <div className="p-6 overflow-auto flex-1">
                <textarea
                  readOnly
                  value={exportData}
                  className="w-full h-96 font-mono text-xs bg-gray-50 border border-gray-200 rounded-lg p-3 outline-none"
                  onFocus={(e) => e.target.select()}
                />
              </div>
              <div className="px-6 py-3 border-t border-gray-200 flex gap-3">
                <button
                  onClick={() => { navigator.clipboard.writeText(exportData); }}
                  className="px-4 py-2 text-sm bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors"
                >
                  📋 复制到剪贴板
                </button>
                <button onClick={() => setShowExport(false)} className="px-4 py-2 text-sm text-gray-600 bg-gray-100 rounded-lg hover:bg-gray-200 transition-colors">关闭</button>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
