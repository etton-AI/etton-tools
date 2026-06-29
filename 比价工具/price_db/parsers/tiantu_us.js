/**
 * 天图通逊 — 美线海运价格解析器
 *
 * 天图结构：仓库为行 × 城市为列，每个 Sheet 含 1-2 个渠道。
 * 城市列固定 11 个：华南/重庆/汕头/厦泉福/华东/青岛等/天津等/济南等/西安等/武汉等
 * 左侧：包税-50KG+ 若干城市列，右侧：不包税-0.5CBM+ 若干城市列
 */

const XLSX = require("xlsx");

const SUPPLIER = "天图通逊";
const COUNTRY = "美国";

// 天图城市列 → 统一城市名映射
const CITY_MAP = {
  华南: ["深圳", "广州", "中山", "东莞南城", "惠州"],
  重庆: ["重庆"],
  汕头: ["汕头"],
  "厦门/泉州/福州": ["厦门", "泉州", "福州"],
  华东: ["义乌", "上海", "宁波", "苏州", "杭州", "绍兴"],
  "青岛/郑州/温州/台州/连云港/南京/合肥": [
    "青岛", "郑州", "温州", "台州", "连云港", "南京", "合肥",
  ],
  "天津/南昌/石家庄": ["天津", "南昌", "石家庄"],
  "济南/潍坊": ["济南", "潍坊"],
  "西安/沧州/保定": ["西安", "沧州", "保定"],
  "武汉/长沙/成都": ["武汉", "长沙", "成都"],
};

// 天图城市名标准化（处理换行符等）
function normalizeCityName(name) {
  return String(name)
    .replace(/\r?\n/g, "")
    .replace(/\s+/g, "")
    .trim();
}

// 在映射表中匹配城市
function matchCity(rawName) {
  const name = normalizeCityName(rawName);
  for (const [key, cities] of Object.entries(CITY_MAP)) {
    const normKey = normalizeCityName(key);
    if (name.includes(normKey) || normKey.includes(name)) {
      return cities;
    }
  }
  // 尝试部分匹配
  for (const [key, cities] of Object.entries(CITY_MAP)) {
    if (key.includes(name) || name.includes(key)) {
      return cities;
    }
  }
  return [name]; // fallback
}

// ── 渠道 → 统一维度映射 ──
function inferTiantuMeta(channelName) {
  const meta = { speed_tier: "", vessel_config: "", delivery_method: "" };
  const n = channelName;

  if (n.includes("直送")) {
    meta.delivery_method = "整柜直送";
  } else {
    meta.delivery_method = "卡派(拆派)";
  }

  if (n.includes("美森极速12日达")) {
    meta.speed_tier = "12-15日达";
    meta.vessel_config = "美森正班CLX";
  } else if (n.includes("CLX13日达") || n.includes("美森正班13日达")) {
    meta.speed_tier = "12-15日达";
    meta.vessel_config = "美森正班CLX";
  } else if (n.includes("MAX14日达") || n.includes("美森15日达")) {
    meta.speed_tier = "12-15日达";
    meta.vessel_config = "美森MAX";
  } else if (n.includes("EXX16日达") || n.includes("OA/以星17日达")) {
    meta.speed_tier = "17-20日达";
    meta.vessel_config = "EXX/以星带托架";
  } else if (n.includes("17日达") && !n.includes("EXX") && !n.includes("OA")) {
    meta.speed_tier = "17-20日达";
    meta.vessel_config = "以星不带托架/合德普提";
  } else if (n.includes("20日达")) {
    meta.speed_tier = "17-20日达";
    meta.vessel_config = "以星/合德/SEA3/CEN";
  } else if (n.includes("22日达") || n.includes("盐田23日达")) {
    meta.speed_tier = "22-28日达";
    meta.vessel_config = "COSCO/OA";
  } else if (n.includes("26日达") || n.includes("盐田26日达")) {
    meta.speed_tier = "22-28日达";
    meta.vessel_config = "OA/MSC/WHL统配";
  } else if (n.includes("28日达直送")) {
    meta.speed_tier = "22-28日达";
    meta.vessel_config = "普船统配";
  } else if (n.includes("美森统配17日达")) {
    meta.speed_tier = "17-20日达";
    meta.vessel_config = "美森统配";
  } else if (n.includes("OA26日达直送")) {
    meta.speed_tier = "22-28日达";
    meta.vessel_config = "OA";
  } else if (n.includes("纽约直航42日达")) {
    meta.speed_tier = "40-45日达";
    meta.vessel_config = "MSC/ZIM/WHL/HMM直航";
  } else if (n.includes("纽约直航38日达")) {
    meta.speed_tier = "40-45日达";
    meta.vessel_config = "直航普船";
  } else if (n.includes("萨凡纳直航")) {
    meta.speed_tier = "40-45日达";
    meta.vessel_config = "OA/ZIM/MSC/WHL统配(SAV)";
  } else if (n.includes("休斯顿直航") || n.includes("休斯顿普船")) {
    meta.speed_tier = "40-45日达";
    meta.vessel_config = "OA/ZIM/MSC统配(HOU)";
  } else if (n.includes("芝加哥海铁")) {
    meta.speed_tier = "22-28日达";
    meta.vessel_config = "海铁统配";
  } else if (n.includes("芝加哥普船")) {
    meta.speed_tier = "22-28日达";
    meta.vessel_config = "OA/ZIM/MSC统配";
  } else if (n.includes("OAK专线") || n.includes("奥克兰")) {
    meta.speed_tier = "22-28日达";
    meta.vessel_config = "OA";
  } else if (n.includes("西雅图专线")) {
    meta.speed_tier = "22-28日达";
    meta.vessel_config = "普船统配";
  } else if (n.includes("纽约美森快线")) {
    meta.speed_tier = "22-28日达";
    meta.vessel_config = "美森CLX";
  } else if (n.includes("纽约海陆")) {
    meta.speed_tier = "22-28日达";
    meta.vessel_config = "海铁/海陆";
  } else if (n.includes("以星/合德20日达")) {
    meta.speed_tier = "17-20日达";
    meta.vessel_config = "以星/合德";
  }

  return meta;
}

function extractVesselTags(vesselConfig) {
  if (!vesselConfig) return [];
  const lower = vesselConfig.toLowerCase();
  const tags = new Set();
  const keywords = {
    matson: ["matson"],
    clx: ["clx"],
    max: ["max"],
    exx: ["exx", "zim"],
    zim: ["zim"],
    以星: ["zim", "合德"],
    合德: ["合德"],
    cosco: ["cosco", "oa"],
    oa: ["oa"],
    msc: ["msc"],
    whl: ["whl"],
    hmm: ["hmm"],
    one: ["one"],
    tsl: ["tsl"],
    oocl: ["oocl"],
    海领: ["海领"],
    海铁: ["海铁"],
  };
  for (const [key, vals] of Object.entries(keywords)) {
    if (lower.includes(key.toLowerCase())) {
      vals.forEach((v) => tags.add(v));
    }
  }
  return [...tags];
}

function inferRegion(sheetName) {
  const s = sheetName.toLowerCase();
  if (s.includes("美西") || s.includes("美森")) return "美西";
  if (s.includes("美西北")) return "美西北";
  if (s.includes("美东")) return "美东北";
  if (s.includes("美东南")) return "美东南";
  if (s.includes("休斯顿")) return "美东南";
  if (s.includes("芝加哥")) return "美中";
  return "美西";
}

// ── 解析天图仓库级 Sheet ──
// 每个 sheet 结构: Row3=标题, Row4=仓库代码头+渠道名, Row5=价格类型(包税/不包税), Row6=城市, Row7+=数据
function parseTiantuWarehouseSheet(ws, sheetName) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 7) return [];

  const results = [];

  // Row 4 (index 3) 包含渠道名 —— 在列 1/列 n 中
  const titleRow = data[3] || [];
  // Row 5 (index 4) 包含价格类型
  const typeRow = data[4] || [];
  // Row 6 (index 5) 是城市列
  const cityRow = data[5] || [];

  // 找到所有渠道的起始位置
  // 渠道名出现在 titleRow 中，格式如 "美森极速12日达-卡派（CLX）"
  const channels = [];
  let i = 1;
  while (i < titleRow.length) {
    const cell = String(titleRow[i] || "").trim();
    if (cell && !cell.includes("仓库代码") && cell.length > 3) {
      // 这是一个渠道名
      channels.push({ name: cell, startCol: i });
    }
    i++;
  }

  // 如果没找到渠道名，使用sheet名推断
  if (channels.length === 0) {
    channels.push({ name: sheetName, startCol: 1 });
  }

  // 为每个渠道确定城市列
  for (let ci = 0; ci < channels.length; ci++) {
    const ch = channels[ci];
    const nextStart = ci + 1 < channels.length ? channels[ci + 1].startCol : cityRow.length;

    // 在这个渠道的列范围内，识别包税和不包税城市列
    // 包税城市列在 typeRow 中包含 "包税" 或 "50kg+"
    // 不包税列包含 "不包税" 或 "CBM" 或 "0.5CBM+"
    const taxCols = []; // 包税城市列 {col, cities}
    const noTaxCols = []; // 不包税城市列

    for (let col = ch.startCol; col < nextStart; col++) {
      const typeCell = String(typeRow[col] || "").toLowerCase();
      const cityCell = String(cityRow[col] || "").trim();

      if (!cityCell) continue;

      const cities = matchCity(cityCell);
      const isTax = typeCell.includes("包税") || typeCell.includes("50kg+") || typeCell.includes("kg+");
      const isNoTax = typeCell.includes("不包税") || typeCell.includes("cbm") || typeCell.includes("0.5cbm+");

      if (isTax || (!isTax && !isNoTax)) {
        // 默认为包税（如果无法判断）
        taxCols.push({ col, cities, region: cityCell });
      }
      if (isNoTax) {
        noTaxCols.push({ col, cities, region: cityCell });
      }
    }

    // 时效和赔付列在渠道块末尾
    // 在 titleRow 和 typeRow 中查找
    let transitCol = -1;
    let claimCol = -1;
    const searchEnd = Math.min(nextStart + 3, Math.max(titleRow.length, typeRow.length));
    for (let col = ch.startCol; col < searchEnd; col++) {
      const t = String(titleRow[col] || "").toLowerCase();
      const tp = String(typeRow[col] || "").toLowerCase();
      if (t.includes("时效") || t.includes("签收") || tp.includes("时效") || tp.includes("签收")) transitCol = col;
      if (t.includes("赔付") || t.includes("延时") || tp.includes("赔付") || tp.includes("延时")) claimCol = col;
    }

    // 解析数据行
    for (let ri = 6; ri < data.length; ri++) {
      const row = data[ri];
      const wareCell = String(row[0] || "").trim();
      if (!wareCell || wareCell === "-" || wareCell.includes("仓库代码")) continue;
      // WM 仓库跳过（Walmart特殊仓）
      if (wareCell.startsWith("WM-")) continue;

      const destCodes = wareCell.split("/").map((s) => s.trim()).filter(Boolean);

      const transitText = transitCol > 0 ? String(row[transitCol] || "").trim() : "";
      const claimText = claimCol > 0 ? String(row[claimCol] || "").trim() : "";
      const transitMin = transitText.match(/(\d+)/) ? parseInt(transitText.match(/(\d+)/)[1]) : null;
      const transitMax = transitText.match(/(\d+)[-–](\d+)/)
        ? parseInt(transitText.match(/(\d+)[-–](\d+)/)[2])
        : transitMin;

      const meta = inferTiantuMeta(ch.name);

      // 处理包税列
      for (const tc of taxCols) {
        const price = parseFloat(row[tc.col]);
        if (isNaN(price) || price <= 0) continue;

        for (const dc of destCodes) {
          results.push({
            supplier: SUPPLIER,
            country: COUNTRY,
            channel_name: ch.name,
            speed_tier: meta.speed_tier,
            vessel_config: meta.vessel_config,
            vessel_tags: extractVesselTags(meta.vessel_config),
            delivery_method: meta.delivery_method,
            destination_type: "warehouse",
            destination_code: dc,
            destination_region: inferRegion(sheetName),
            origin_region: tc.region,
            origin_cities: tc.cities,
            billing_type: "含税KG",
            min_quantity: "50KG+",
            min_quantity_value: 50,
            unit_price: price,
            price_unit: "元/KG",
            transit_time_min: transitMin,
            transit_time_max: transitMax,
            transit_time_desc: transitText,
            claim_rule: claimText.replace(/\r?\n/g, ""),
            effective_date: "",
            source_sheet: sheetName,
          });
        }
      }

      // 处理不包税列
      for (const ntc of noTaxCols) {
        const price = parseFloat(row[ntc.col]);
        if (isNaN(price) || price <= 0) continue;

        for (const dc of destCodes) {
          results.push({
            supplier: SUPPLIER,
            country: COUNTRY,
            channel_name: ch.name,
            speed_tier: meta.speed_tier,
            vessel_config: meta.vessel_config,
            vessel_tags: extractVesselTags(meta.vessel_config),
            delivery_method: meta.delivery_method,
            destination_type: "warehouse",
            destination_code: dc,
            destination_region: inferRegion(sheetName),
            origin_region: ntc.region,
            origin_cities: ntc.cities,
            billing_type: "不含税CBM",
            min_quantity: "0.5CBM+",
            min_quantity_value: 0.5,
            unit_price: price,
            price_unit: "元/CBM",
            transit_time_min: transitMin,
            transit_time_max: transitMax,
            transit_time_desc: transitText,
            claim_rule: claimText.replace(/\r?\n/g, ""),
            effective_date: "",
            source_sheet: sheetName,
          });
        }
      }
    }
  }

  return results;
}

// ── 主解析入口 ──────────────────────────────────
function parseTiantu(filePath) {
  console.log("[天图] 开始解析:", filePath);
  const wb = XLSX.readFile(filePath);
  const allResults = [];

  // 美线海运相关 Sheets
  const usSheets = [
    "美西直送系列",
    "美西-美森(12-17)",
    "美西-以星&普船(17-26)",
    "美西北",
    "美东直航系列",
    "美东-海铁系列",
    "美东南直航系列",
    "美东南海陆系列",
    "美中-休斯顿",
    "美中-芝加哥",
  ];

  for (const sn of usSheets) {
    if (wb.SheetNames.includes(sn)) {
      const results = parseTiantuWarehouseSheet(wb.Sheets[sn], sn);
      console.log(`  [${sn}] ${results.length} 条`);
      allResults.push(...results);
    } else {
      console.log(`  [${sn}] Sheet不存在，跳过`);
    }
  }

  console.log(`[天图] 总计解析 ${allResults.length} 条价格记录`);
  return allResults;
}

module.exports = { parseTiantu };
