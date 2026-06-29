#!/usr/bin/env node
/**
 * query.js — 比价查询 CLI
 *
 * 日常查价工具，支持按目的仓/发货城市/重量/船司/送仓方式过滤，
 * 按单价排序输出，标记最优选择。
 *
 * 用法:
 *   node query.js -d ONT8 -o 深圳 -w 100 -v EXX -m 卡派
 *   node query.js -d ONT8 -o 深圳
 *   node query.js -d ONT8 --best
 */

const path = require("path");
const fs = require("fs");

// ── 命令行参数解析 ──
function parseArgs() {
  const args = process.argv.slice(2);
  const params = {
    dest: null,
    origin: null,
    weight: null,
    vessel: null,
    method: null,
    supplier: null,
    top: 0,
    best: false,
    export: null,
    help: false,
  };

  for (let i = 0; i < args.length; i++) {
    const a = args[i];
    switch (a) {
      case "-d":
      case "--dest":
        params.dest = args[++i]?.toUpperCase();
        break;
      case "-o":
      case "--origin":
        params.origin = args[++i];
        break;
      case "-w":
      case "--weight":
        params.weight = parseFloat(args[++i]);
        break;
      case "-v":
      case "--vessel":
        params.vessel = args[++i];
        break;
      case "-m":
      case "--method":
        params.method = args[++i];
        break;
      case "-s":
      case "--supplier":
        params.supplier = args[++i];
        break;
      case "-t":
      case "--top":
        params.top = parseInt(args[++i]);
        break;
      case "--best":
        params.best = true;
        break;
      case "--export":
        params.export = args[++i] || "csv";
        break;
      case "-h":
      case "--help":
        params.help = true;
        break;
    }
  }
  return params;
}

// ── 帮助信息 ──
function showHelp() {
  console.log(`
  ╔══════════════════════════════════════════════════════════╗
  ║        ETTON 跨境物流比价查询工具                       ║
  ╚══════════════════════════════════════════════════════════╝

  用法: node query.js [选项]

  查询选项:
    -d, --dest      <仓库代码>  目的仓 (如 ONT8, LAX9, FTW1)
    -o, --origin    <城市名>    发货城市 (如 深圳, 义乌, 东莞)
    -w, --weight    <重量KG>    货物重量 (如 100)
    -v, --vessel    <关键词>    船司 (如 EXX, 美森, CLX, COSCO, 以星)
    -m, --method    <方式>      送仓方式 (卡派, 海派, 整柜直送, 自提)
    -s, --supplier  <供应商>    指定供应商 (ETTON, 天图, 英美)

  输出选项:
    -t, --top       <N>         只显示前N条
    --best                      只显示最优选择(最低价)
    --export        <format>    导出格式 (csv, json)

  示例:
    node query.js -d ONT8 -o 深圳 -w 100 -v EXX -m 卡派
    node query.js -d ONT8 -o 深圳                     # 查所有渠道
    node query.js -d LAX9 -v 美森 --best               # 只看最优
    node query.js -d ONT8 --export csv                 # 导出CSV
  `);
}

// ── 城市到供应商区域的映射 ──
const CITY_TO_ORIGIN = {
  // ETTON 区域映射
  etton: {
    深圳: ["东莞", "中山", "广州"],
    东莞: ["东莞", "中山", "广州"],
    广州: ["东莞", "中山", "广州"],
    中山: ["东莞", "中山", "广州"],
    惠州: ["东莞", "中山", "广州"],
    嘉兴: ["嘉兴", "义乌"],
    义乌: ["嘉兴", "义乌"],
    杭州: ["嘉兴", "义乌"],
    宁波: ["嘉兴", "义乌"],
    上海: ["嘉兴", "义乌"],
    苏州: ["嘉兴", "义乌"],
    汕头: ["汕头", "厦门", "泉州"],
    厦门: ["汕头", "厦门", "泉州"],
    泉州: ["汕头", "厦门", "泉州"],
    福州: ["汕头", "厦门", "泉州"],
    武汉: ["武汉", "长沙"],
    长沙: ["武汉", "长沙"],
  },
  // 天图城市映射（天图按城市单独定价，用模糊匹配）
  tiantu: {
    深圳: ["深圳", "广州", "中山", "东莞南城", "惠州"],
    东莞: ["深圳", "广州", "中山", "东莞南城", "惠州"],
    广州: ["深圳", "广州", "中山", "东莞南城", "惠州"],
    中山: ["深圳", "广州", "中山", "东莞南城", "惠州"],
    惠州: ["深圳", "广州", "中山", "东莞南城", "惠州"],
    义乌: ["义乌", "上海", "宁波", "苏州", "杭州", "绍兴"],
    上海: ["义乌", "上海", "宁波", "苏州", "杭州", "绍兴"],
    杭州: ["义乌", "上海", "宁波", "苏州", "杭州", "绍兴"],
    宁波: ["义乌", "上海", "宁波", "苏州", "杭州", "绍兴"],
    厦门: ["厦门", "泉州", "福州"],
    泉州: ["厦门", "泉州", "福州"],
    福州: ["厦门", "泉州", "福州"],
    汕头: ["汕头"],
    重庆: ["重庆"],
    武汉: ["武汉", "长沙", "成都"],
    长沙: ["武汉", "长沙", "成都"],
    青岛: ["青岛", "郑州", "温州", "台州", "连云港", "南京", "合肥"],
    济南: ["济南", "潍坊"],
    天津: ["天津", "南昌", "石家庄"],
    西安: ["西安", "沧州", "保定"],
  },
  // 英美区域映射
  yingmei: {
    深圳: ["东莞", "宝安", "中山", "广州", "南城", "汕头", "深圳"],
    东莞: ["东莞", "宝安", "中山", "广州", "南城", "汕头", "深圳"],
    宝安: ["东莞", "宝安", "中山", "广州", "南城", "汕头", "深圳"],
    广州: ["东莞", "宝安", "中山", "广州", "南城", "汕头", "深圳"],
    中山: ["东莞", "宝安", "中山", "广州", "南城", "汕头", "深圳"],
    汕头: ["东莞", "宝安", "中山", "广州", "南城", "汕头", "深圳"],
    义乌: ["义乌", "上海", "宁波", "杭州", "温州"],
    上海: ["义乌", "上海", "宁波", "杭州", "温州"],
    杭州: ["义乌", "上海", "宁波", "杭州", "温州"],
    宁波: ["义乌", "上海", "宁波", "杭州", "温州"],
    温州: ["义乌", "上海", "宁波", "杭州", "温州"],
    福州: ["福州", "厦门", "泉州", "合肥", "青岛", "温州", "汕头"],
    厦门: ["福州", "厦门", "泉州", "合肥", "青岛", "温州", "汕头"],
    泉州: ["福州", "厦门", "泉州", "合肥", "青岛", "温州", "汕头"],
    合肥: ["福州", "厦门", "泉州", "合肥", "青岛", "温州", "汕头"],
    青岛: ["福州", "厦门", "泉州", "合肥", "青岛", "温州", "汕头"],
  },
};

// ── 供应商简称映射 ──
const SUPPLIER_ALIAS = {
  etton: "易通",
  ettong: "易通",
  易通: "易通",
  tiantu: "天图",
  天图: "天图",
  yingmei: "英美",
  英美: "英美",
};

// ── 数据加载 ──
function loadPrices() {
  const dataPath = path.resolve(__dirname, "data", "prices.json");
  if (!fs.existsSync(dataPath)) {
    console.error("❌ 未找到数据文件，请先运行: node parsers/build_db.js");
    process.exit(1);
  }
  const raw = JSON.parse(fs.readFileSync(dataPath, "utf-8"));
  console.log(`📅 数据日期: ${raw.generated_at.split("T")[0]} | ${raw.total_records} 条记录\n`);
  return raw.data;
}

// ── 主查询逻辑 ──
function query(prices, params) {
  let results = [...prices];

  // 1. 目的仓过滤
  if (params.dest) {
    results = results.filter((r) => {
      // 排除不区分目的仓的渠道（如海外仓自提）
      if (r.destination_type === "none" || r.destination_code === "*") return false;
      return r.destination_code.toUpperCase() === params.dest.toUpperCase();
    });
  }

  // 2. 发货城市过滤
  if (params.origin) {
    const origin = params.origin;
    // 用 CITY_TO_ORIGIN 映射表将用户输入的城市转为各供应商的匹配城市列表
    const ettonCities = (CITY_TO_ORIGIN.etton[origin] || []).map(c => c.toLowerCase());
    const tiantuCities = (CITY_TO_ORIGIN.tiantu[origin] || []).map(c => c.toLowerCase());
    const yingmeiCities = (CITY_TO_ORIGIN.yingmei[origin] || []).map(c => c.toLowerCase());
    const searchLower = origin.toLowerCase();

    results = results.filter((r) => {
      if (!r.origin_cities || r.origin_cities.length === 0) return true;
      const cities = r.origin_cities.map((c) => c.toLowerCase());
      const region = r.origin_region.toLowerCase();
      const supplier = (r.supplier || "").toLowerCase();

      // 根据供应商选择对应的城市映射
      let targetCities;
      if (supplier.includes("易通") || supplier.includes("etton")) {
        targetCities = ettonCities;
      } else if (supplier.includes("天图") || supplier.includes("tiantu")) {
        targetCities = tiantuCities;
      } else if (supplier.includes("英美") || supplier.includes("yingmei")) {
        targetCities = yingmeiCities;
      } else {
        targetCities = [];
      }

      // 城市名直接匹配 (搜索词在origin_cities中，或origin_cities在映射表中)
      if (cities.some((c) => c.includes(searchLower) || searchLower.includes(c))) return true;
      // 通过映射表匹配
      if (targetCities.length > 0 && targetCities.some((tc) => cities.some((c) => c.includes(tc) || tc.includes(c)))) return true;
      // 区域名中包含搜索词
      if (region.includes(searchLower)) return true;

      return false;
    });
  }

  // 3. 重量过滤
  if (params.weight) {
    const w = params.weight;
    results = results.filter((r) => {
      // CBM 计费的跳过（按方不计重）
      if (r.price_unit === "元/CBM") return false;
      return r.min_quantity_value <= w;
    });
    // 对同一渠道+区域，选匹配最近的重量段
    const grouped = {};
    for (const r of results) {
      const key = `${r.supplier}|${r.channel_name}|${r.destination_code}|${r.origin_region}`;
      if (!grouped[key] || r.min_quantity_value > grouped[key].min_quantity_value) {
        grouped[key] = r;
      }
    }
    results = Object.values(grouped);
  }

  // 4. 船司过滤
  if (params.vessel) {
    const v = params.vessel.toLowerCase();
    results = results.filter((r) => {
      const vesselTags = (r.vessel_tags || []).map((t) => t.toLowerCase());
      const vesselConfig = (r.vessel_config || "").toLowerCase();
      const channelName = (r.channel_name || "").toLowerCase();
      return (
        vesselTags.some((t) => t.includes(v) || v.includes(t)) ||
        vesselConfig.includes(v) ||
        channelName.includes(v)
      );
    });
  }

  // 5. 送仓方式过滤
  if (params.method) {
    const m = params.method;
    results = results.filter((r) => {
      const dm = (r.delivery_method || "").toLowerCase();
      const methodLower = m.toLowerCase();
      if (methodLower.includes("卡派")) return dm.includes("卡派") || dm.includes("拆派");
      if (methodLower.includes("海派")) return dm.includes("海派") || dm.includes("快递派");
      if (methodLower.includes("整柜") || methodLower.includes("直送")) return dm.includes("整柜") || dm.includes("直送");
      if (methodLower.includes("自提")) return dm.includes("自提");
      return dm.includes(methodLower);
    });
  }

  // 6. 供应商过滤
  if (params.supplier) {
    const sAlias = SUPPLIER_ALIAS[params.supplier.toLowerCase()] || params.supplier;
    results = results.filter((r) => {
      const supName = (r.supplier || "").toLowerCase();
      return supName.includes(sAlias) || supName.includes(params.supplier.toLowerCase());
    });
  }

  // 7. 排序: 单价升序 → 时效升序
  results.sort((a, b) => {
    if (a.unit_price !== b.unit_price) return a.unit_price - b.unit_price;
    return (a.transit_time_min || 999) - (b.transit_time_min || 999);
  });

  // 8. Top N 或 best
  if (params.best && results.length > 0) {
    results = [results[0]];
  } else if (params.top > 0) {
    results = results.slice(0, params.top);
  }

  return results;
}

// ── 格式化输出 ──
function formatOutput(results, params) {
  // 构建查询条件描述
  const conditions = [];
  if (params.dest) conditions.push(params.dest);
  if (params.origin) conditions.push(params.origin + "仓");
  if (params.weight) conditions.push(params.weight + "KG+");
  if (params.vessel) conditions.push(params.vessel + "船");
  if (params.method) conditions.push(params.method);

  console.log(`🔍 查询条件: ${conditions.join(" | ")}`);
  console.log("");

  if (results.length === 0) {
    console.log("❌ 未找到匹配的价格记录");
    console.log("💡 提示: 尝试放宽条件（如去掉船司/重量限制），或检查仓库代码是否正确");
    return;
  }

  // 表头
  const header =
    "┌──────┬────────────┬──────────────────────────┬──────────────────────┬──────────┬──────────┬────────────────────────────┬────────────────────┐";
  const title =
    "│  #   │ 供应商     │ 渠道名                   │ 船配置               │ 单价     │ 时效     │ 赔付规则                   │ 发货仓             │";
  const sep =
    "├──────┼────────────┼──────────────────────────┼──────────────────────┼──────────┼──────────┼────────────────────────────┼────────────────────┤";

  console.log(header);
  console.log(title);
  console.log(sep);

  // 数据行
  for (let i = 0; i < results.length; i++) {
    const r = results[i];
    const num = String(i + 1).padStart(3);
    const supplier = (r.supplier || "").padEnd(10).substring(0, 10);
    const channel = (r.channel_name || "").padEnd(24).substring(0, 24);
    const vessel = (r.vessel_config || "-").padEnd(20).substring(0, 20);
    const price = `${r.unit_price} ${r.price_unit || ""}`.padEnd(8);
    const transit = r.transit_time_min
      ? `${r.transit_time_min}-${r.transit_time_max || r.transit_time_min}天`.padEnd(8)
      : "-".padEnd(8);
    const claim = (r.claim_rule || "-").padEnd(26).substring(0, 26);
    const origin = (r.origin_region || "").padEnd(18).substring(0, 18);

    const row = `│ ${num}  │ ${supplier} │ ${channel} │ ${vessel} │ ${price} │ ${transit} │ ${claim} │ ${origin} │`;
    console.log(row);
  }

  console.log(
    "└──────┴────────────┴──────────────────────────┴──────────────────────┴──────────┴──────────┴────────────────────────────┴────────────────────┘"
  );

  // 最优选择
  console.log("");
  const best = results[0];
  console.log(`🏆 最优选择: ${best.supplier} ${best.channel_name}`);
  console.log(`   💰 ${best.unit_price} ${best.price_unit} | ⏱ ${best.transit_time_desc || best.transit_time_min + "-" + best.transit_time_max + "天"} | 🚢 ${best.vessel_config}`);
  if (best.claim_rule) console.log(`   📋 赔付: ${best.claim_rule}`);
  console.log(`   📦 发货: ${best.origin_region}`);
  console.log(`   📄 来源: ${best.source_file}`);

  // 如果有多个结果，给个简短对比
  if (results.length > 1) {
    const worst = results[results.length - 1];
    const diff = worst.unit_price - best.unit_price;
    const pct = best.unit_price > 0 ? ((diff / best.unit_price) * 100).toFixed(0) : 0;
    console.log(`\n💡 价差: 最高 ${worst.unit_price} vs 最低 ${best.unit_price}，差额 ${diff.toFixed(2)} (${pct}%)`);
  }
}

// ── 导出函数 ──
function exportResults(results, format) {
  switch (format.toLowerCase()) {
    case "json":
      console.log(JSON.stringify(results, null, 2));
      break;
    case "csv":
      if (results.length === 0) {
        console.log("(empty)");
        return;
      }
      const headers = Object.keys(results[0]);
      console.log(headers.join(","));
      for (const r of results) {
        console.log(headers.map((h) => `"${String(r[h] || "").replace(/"/g, '""')}"`).join(","));
      }
      break;
    default:
      console.error("不支持的导出格式: " + format);
  }
}

// ── 入口 ──
function main() {
  const params = parseArgs();

  if (params.help || (!params.dest && !params.origin && !params.vessel)) {
    showHelp();
    process.exit(0);
  }

  const prices = loadPrices();
  const results = query(prices, params);

  if (params.export) {
    exportResults(results, params.export);
  } else {
    formatOutput(results, params);
  }
}

main();
