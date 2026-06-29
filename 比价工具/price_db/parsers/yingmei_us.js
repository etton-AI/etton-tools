/**
 * 英美跨境 — 美线价格解析器
 *
 * 结构特点：
 * - 海派: 渠道代码 A1/A2/A4/A5，按邮编分区，3 国内仓区域，2 重量段(12-100KG/101KG+)
 * - 海卡: 渠道代码 B1/B2/B4/B5/B7/B10/B12/B14/B15，仓库分组(中文逗号分隔)，4 重量段
 * - 海外仓自提: B8 渠道，不区分目的仓
 * - 渠道代码体系: A=海派, B=卡派/自提
 */

const XLSX = require("xlsx");

const SUPPLIER = "英美跨境";
const COUNTRY = "美国";

// 国内仓区域 (英美有3个区域，与ETTON/天图不同!)
const YM_REGIONS = [
  { name: "义乌/上海/宁波/杭州", cities: ["义乌", "上海", "宁波", "杭州", "温州"] },
  { name: "东莞/宝安/中山/广州/南城", cities: ["东莞", "宝安", "中山", "广州", "南城", "汕头", "深圳"] },
  { name: "福州/厦门/泉州/合肥/青岛/温州/汕头", cities: ["福州", "厦门", "泉州", "合肥", "青岛", "温州", "汕头"] },
];

// 渠道代码 → 统一维度映射
const CHANNEL_META = {
  A1: { speed_tier: "12-15日达", vessel_config: "美森CLX", delivery_method: "海派(快递派)" },
  A2: { speed_tier: "12-15日达", vessel_config: "美森MAX", delivery_method: "海派(快递派)" },
  A4_18: { speed_tier: "17-20日达", vessel_config: "合德快提/ZIM带车架", delivery_method: "海派(快递派)" },
  A4_21: { speed_tier: "17-20日达", vessel_config: "合德/ZIM", delivery_method: "海派(快递派)" },
  A5: { speed_tier: "22-28日达", vessel_config: "OA/MSC/WHL/HPL统配", delivery_method: "海派(快递派)" },
  B1: { speed_tier: "12-15日达", vessel_config: "美森CLX", delivery_method: "卡派(拆派)" },
  B2: { speed_tier: "12-15日达", vessel_config: "美森MAX", delivery_method: "卡派(拆派)" },
  B4: { speed_tier: "17-20日达", vessel_config: "合德/ZIM", delivery_method: "卡派(拆派)" },
  B5: { speed_tier: "22-28日达", vessel_config: "普船统配(LA)", delivery_method: "卡派(拆派)" },
  B7: { speed_tier: "40-45日达", vessel_config: "普船统配(NY)", delivery_method: "卡派(拆派)" },
  B8_CLX: { speed_tier: "12-15日达", vessel_config: "美森CLX", delivery_method: "海外仓自提" },
  B8_MAX: { speed_tier: "12-15日达", vessel_config: "美森MAX", delivery_method: "海外仓自提" },
  B8_HDZIM: { speed_tier: "17-20日达", vessel_config: "合德/ZIM", delivery_method: "海外仓自提" },
  B8_LA: { speed_tier: "22-28日达", vessel_config: "普船统配(LA)", delivery_method: "海外仓自提" },
  B8_NY: { speed_tier: "40-45日达", vessel_config: "普船统配(NY)", delivery_method: "海外仓自提" },
  B8_OAK: { speed_tier: "22-28日达", vessel_config: "普船统配(OAK)", delivery_method: "海外仓自提" },
  B8_HOU: { speed_tier: "40-45日达", vessel_config: "普船统配(HOU)", delivery_method: "海外仓自提" },
  B8_SAV: { speed_tier: "40-45日达", vessel_config: "普船统配(SAV)", delivery_method: "海外仓自提" },
  B8_CHI: { speed_tier: "22-28日达", vessel_config: "普船统配(CHI)", delivery_method: "海外仓自提" },
  B10: { speed_tier: "40-45日达", vessel_config: "普船统配(OAK)", delivery_method: "卡派(拆派)" },
  B12: { speed_tier: "22-28日达", vessel_config: "COSCO/OOCL", delivery_method: "卡派(拆派)" },
  B14: { speed_tier: "40-45日达", vessel_config: "WHL/HMM/海领/HPL", delivery_method: "卡派(拆派)" },
  B15: { speed_tier: "22-28日达", vessel_config: "WHL/HMM/海领/HPL", delivery_method: "卡派(拆派)" },
};

function getChannelMeta(channelCode, channelFullName) {
  // 特殊处理 A4（18日/21日两个变体）
  if (channelCode === "A4") {
    if (channelFullName.includes("18日")) return CHANNEL_META.A4_18;
    if (channelFullName.includes("21日")) return CHANNEL_META.A4_21;
    return CHANNEL_META.A4_18; // default
  }
  // 特殊处理 B8（多个子渠道）
  if (channelCode === "B8") {
    const n = channelFullName;
    if (n.includes("美森正班")) return CHANNEL_META.B8_CLX;
    if (n.includes("美森加班")) return CHANNEL_META.B8_MAX;
    if (n.includes("合德以星") || n.includes("合德")) return CHANNEL_META.B8_HDZIM;
    if (n.includes("普船洛杉矶")) return CHANNEL_META.B8_LA;
    if (n.includes("普船纽约")) return CHANNEL_META.B8_NY;
    if (n.includes("普船奥克兰")) return CHANNEL_META.B8_OAK;
    if (n.includes("普船休斯顿")) return CHANNEL_META.B8_HOU;
    if (n.includes("普船萨凡纳")) return CHANNEL_META.B8_SAV;
    if (n.includes("普船芝加哥")) return CHANNEL_META.B8_CHI;
    return CHANNEL_META.B8_CLX;
  }
  return CHANNEL_META[channelCode] || {};
}

function extractVesselTags(vesselConfig) {
  if (!vesselConfig) return [];
  const lower = vesselConfig.toLowerCase();
  const tags = new Set();
  const keywords = [
    "matson", "clx", "max", "exx", "zim", "合德", "cosco", "oa", "msc",
    "whl", "hpl", "hmm", "one", "tsl", "oocl", "海领",
  ];
  keywords.forEach((kw) => {
    if (lower.includes(kw.toLowerCase())) tags.add(kw.toLowerCase());
  });
  return [...tags];
}

// ── 从渠道名中提取渠道代码 ──
function extractChannelCode(channelName) {
  const m = channelName.match(/\b([AB]\d+)\b/);
  return m ? m[1] : null;
}

// ── 解析仓库代码分组（中文逗号分隔） ──
function splitWarehouses(cell) {
  if (!cell) return [];
  return String(cell)
    .split(/[，,]/)
    .map((s) => s.trim())
    .filter((s) => s && s.length >= 3 && !s.includes("特价") && !s.includes("特惠"));
}

// ── 解析时效文本 ──
function parseTransitTime(text) {
  if (!text) return { min: null, max: null, desc: "" };
  const str = String(text).trim();
  const match = str.match(/(\d+)\s*[-–]\s*(\d+)/);
  if (match) return { min: parseInt(match[1]), max: parseInt(match[2]), desc: str };
  const single = str.match(/(\d+)/);
  if (single) return { min: parseInt(single[1]), max: parseInt(single[1]), desc: str };
  return { min: null, max: null, desc: str };
}

// ── 判断邮编分区属于哪个区域 ──
function inferZipRegion(zoneText) {
  const t = String(zoneText);
  if (t.includes("美西") || t.includes("8") && t.includes("9")) return "美西";
  if (t.includes("美中") || t.includes("4") || t.includes("5") || t.includes("6") || t.includes("7")) return "美中";
  if (t.includes("美东") || t.includes("0") || t.includes("1") || t.includes("2") || t.includes("3")) return "美东";
  return "";
}

// ── 解析 美国海派 Sheet ──
function parseSeaParcel(ws) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 6) return [];
  const results = [];

  let currentChannelCode = null;
  let currentChannelName = null;

  // Row 4-5 是表头，Row 6+ 是数据
  for (let ri = 5; ri < data.length; ri++) {
    const row = data[ri];
    const col0 = String(row[0] || "").trim();
    const col1 = String(row[1] || "").trim();

    // 检测渠道名称行
    if (col0 === "渠道名称" || col0 === "渠道代码") continue;

    // 新渠道开始
    if (col0 && col0.match(/^[AB]\d/)) {
      currentChannelName = col0;
      currentChannelCode = extractChannelCode(col0);
      continue;
    }

    // 邮编分区行
    if (col1 && (col1.includes("美西") || col1.includes("美中") || col1.includes("美东"))) {
      if (!currentChannelCode) continue;

      const zoneText = col1;
      const meta = getChannelMeta(currentChannelCode, currentChannelName);

      // 3 个区域 × 2 重量段 (12-100KG, 101KG+) = 6 个价格列
      // 列布局: col2-3=区域A(12-100KG,101KG+), col4-5=区域B, col6-7=区域C
      const regionCols = [
        { region: YM_REGIONS[0], kg100Col: 2, kg101Col: 3 },
        { region: YM_REGIONS[1], kg100Col: 4, kg101Col: 5 },
        { region: YM_REGIONS[2], kg100Col: 6, kg101Col: 7 },
      ];

      const transitText = String(row[9] || "").trim();
      const claimText = String(row[10] || "").trim();
      const transit = parseTransitTime(transitText);

      for (const rc of regionCols) {
        // 12-100KG
        const price100 = parseFloat(row[rc.kg100Col]);
        if (!isNaN(price100) && price100 > 0) {
          results.push({
            supplier: SUPPLIER,
            country: COUNTRY,
            channel_name: `${currentChannelCode} ${currentChannelName.replace(currentChannelCode, "").trim()}`,
            speed_tier: meta.speed_tier || "",
            vessel_config: meta.vessel_config || "",
            vessel_tags: extractVesselTags(meta.vessel_config),
            delivery_method: meta.delivery_method || "海派(快递派)",
            destination_type: "zip_zone",
            destination_code: zoneText,
            destination_region: inferZipRegion(zoneText),
            origin_region: rc.region.name,
            origin_cities: rc.region.cities,
            billing_type: "海派KG",
            min_quantity: "12-100KG",
            min_quantity_value: 12,
            unit_price: price100,
            price_unit: "元/KG",
            transit_time_min: transit.min,
            transit_time_max: transit.max,
            transit_time_desc: transit.desc,
            claim_rule: claimText.replace(/\r?\n/g, ""),
            effective_date: "",
            source_sheet: "美国海派",
          });
        }

        // 101KG+
        const price101 = parseFloat(row[rc.kg101Col]);
        if (!isNaN(price101) && price101 > 0) {
          results.push({
            supplier: SUPPLIER,
            country: COUNTRY,
            channel_name: `${currentChannelCode} ${currentChannelName.replace(currentChannelCode, "").trim()}`,
            speed_tier: meta.speed_tier || "",
            vessel_config: meta.vessel_config || "",
            vessel_tags: extractVesselTags(meta.vessel_config),
            delivery_method: meta.delivery_method || "海派(快递派)",
            destination_type: "zip_zone",
            destination_code: zoneText,
            destination_region: inferZipRegion(zoneText),
            origin_region: rc.region.name,
            origin_cities: rc.region.cities,
            billing_type: "海派KG",
            min_quantity: "101KG+",
            min_quantity_value: 101,
            unit_price: price101,
            price_unit: "元/KG",
            transit_time_min: transit.min,
            transit_time_max: transit.max,
            transit_time_desc: transit.desc,
            claim_rule: claimText.replace(/\r?\n/g, ""),
            effective_date: "",
            source_sheet: "美国海派",
          });
        }
      }
    }
  }

  return results;
}

// ── 解析 海卡类 Sheet (美森海卡, 合德以星海卡, 普船海卡等) ──
function parseSeaCard(ws, sheetName) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 5) return [];
  const results = [];

  // Row 3: 渠道名称, Row 4: 仓库代码 + 重量段头
  let currentChannelName = null;

  for (let ri = 4; ri < data.length; ri++) {
    const row = data[ri];
    const col0 = String(row[0] || "").trim();
    const col1 = String(row[1] || "").trim();

    // 跳过空行、说明行
    if (!col0 && !col1) continue;
    if (col0.includes("海关查验") || col0.includes("理赔注意") || col0.includes("产品附加") ||
        col0.includes("拒收产品") || col0.includes("服务小贴士") || col0.includes("其它费用") ||
        col0.includes("特别提示") || col0.includes("重货优惠")) continue;

    // 渠道名称行 (第一列含渠道代码如 B1/B2/B4等)
    if (col0 && col0.match(/^B\d+/)) {
      currentChannelName = col0;
      continue;
    }

    // 仓库分组行 (渠道名只在第一行col0, 后续行col0为空、col1为仓库分组)
    const warehouseCell = col0 || col1;
    if (warehouseCell && currentChannelName) {
      // 同时检查是否有 "按方包税" 变体
      const fullName = currentChannelName.replace(/\r?\n/g, " ");
      const channelCode = extractChannelCode(currentChannelName);

      if (!channelCode) continue;

      const warehouses = splitWarehouses(warehouseCell);
      if (warehouses.length === 0) continue;

      const meta = getChannelMeta(channelCode, fullName);

      // 列布局: col2=空白, col3-6=区域A(12KG+,51KG+,350KG+,1CBM+), col7-10=区域B, col11-14=区域C
      const regionCols = [
        { region: YM_REGIONS[0], kg12Col: 3, kg51Col: 4, kg350Col: 5, cbm1Col: 6 },
        { region: YM_REGIONS[1], kg12Col: 7, kg51Col: 8, kg350Col: 9, cbm1Col: 10 },
        { region: YM_REGIONS[2], kg12Col: 11, kg51Col: 12, kg350Col: 13, cbm1Col: 14 },
      ];

      const transitText = String(row[15] || "").trim();
      const claimText = String(row[16] || "").trim();
      const transit = parseTransitTime(transitText);

      for (const rc of regionCols) {
        // 12KG+
        const p12 = parseFloat(row[rc.kg12Col]);
        if (!isNaN(p12) && p12 > 0) {
          for (const wh of warehouses) {
            results.push(buildRecord(meta, channelCode, fullName, wh, rc.region, "含税KG", "12KG+", 12, p12, "元/KG", transit, claimText, sheetName));
          }
        }
        // 51KG+ (或 51KG+)
        const p51 = parseFloat(row[rc.kg51Col]);
        if (!isNaN(p51) && p51 > 0) {
          for (const wh of warehouses) {
            results.push(buildRecord(meta, channelCode, fullName, wh, rc.region, "含税KG", "51KG+", 51, p51, "元/KG", transit, claimText, sheetName));
          }
        }
        // 350KG+ (或 300KG+)
        const p350 = parseFloat(row[rc.kg350Col]);
        if (!isNaN(p350) && p350 > 0) {
          for (const wh of warehouses) {
            results.push(buildRecord(meta, channelCode, fullName, wh, rc.region, "含税KG", "350KG+", 350, p350, "元/KG", transit, claimText, sheetName));
          }
        }
        // 1CBM+
        const pCbm = parseFloat(row[rc.cbm1Col]);
        if (!isNaN(pCbm) && pCbm > 0) {
          for (const wh of warehouses) {
            results.push(buildRecord(meta, channelCode, fullName, wh, rc.region, "不含税CBM", "1CBM+", 1, pCbm, "元/CBM", transit, claimText, sheetName));
          }
        }
      }
    }
  }

  return results;
}

function buildRecord(meta, channelCode, fullName, destCode, region, billingType, minQty, minQtyVal, price, priceUnit, transit, claimText, sheetName) {
  const destType = (sheetName === "海外仓自提") ? "none" : "warehouse";
  return {
    supplier: SUPPLIER,
    country: COUNTRY,
    channel_name: `${channelCode} ${fullName.replace(channelCode, "").trim()}`,
    speed_tier: meta.speed_tier || "",
    vessel_config: meta.vessel_config || "",
    vessel_tags: extractVesselTags(meta.vessel_config),
    delivery_method: meta.delivery_method || "卡派(拆派)",
    destination_type: destType,
    destination_code: destType === "none" ? "*" : destCode,
    destination_region: "",
    origin_region: region.name,
    origin_cities: region.cities,
    billing_type: billingType,
    min_quantity: minQty,
    min_quantity_value: minQtyVal,
    unit_price: price,
    price_unit: priceUnit,
    transit_time_min: transit.min,
    transit_time_max: transit.max,
    transit_time_desc: transit.desc,
    claim_rule: String(claimText).replace(/\r?\n/g, ""),
    effective_date: "",
    source_sheet: sheetName,
  };
}

// ── 解析 海外仓自提 Sheet ──
function parseSelfPickup(ws) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 5) return [];
  const results = [];

  for (let ri = 4; ri < data.length; ri++) {
    const row = data[ri];
    const col0 = String(row[0] || "").trim();

    // 跳过说明行
    if (!col0 || col0.includes("产品附加") || col0.includes("拒收产品") ||
        col0.includes("注意事项") || col0.includes("特别提示") || col0.includes("所有自提")) continue;

    // 渠道行 (B8...)
    if (!col0.startsWith("B8")) continue;

    const fullName = col0.replace(/\r?\n/g, " ");
    const channelCode = "B8";
    const meta = getChannelMeta(channelCode, fullName);

    // 列布局: col2=空白, col3-4=区域A(350kg+,2CBM+), col5-6=区域B, col7-8=区域C
    const regionCols = [
      { region: YM_REGIONS[0], kgCol: 3, cbmCol: 4 },
      { region: YM_REGIONS[1], kgCol: 5, cbmCol: 6 },
      { region: YM_REGIONS[2], kgCol: 7, cbmCol: 8 },
    ];

    const transitText = String(row[9] || "").trim();
    const transit = parseTransitTime(transitText);

    for (const rc of regionCols) {
      // 350KG+
      const pKg = parseFloat(row[rc.kgCol]);
      if (!isNaN(pKg) && pKg > 0) {
        results.push(buildRecord(meta, channelCode, fullName, "*", rc.region, "含税KG", "350KG+", 350, pKg, "元/KG", transit, "", "海外仓自提"));
      }
      // 2CBM+
      const pCbm = parseFloat(row[rc.cbmCol]);
      if (!isNaN(pCbm) && pCbm > 0) {
        results.push(buildRecord(meta, channelCode, fullName, "*", rc.region, "不含税CBM", "2CBM+", 2, pCbm, "元/CBM", transit, "", "海外仓自提"));
      }
    }
  }

  return results;
}

// ── 解析 美西/美中/美东普船海卡渠道 Sheets ──
// 结构与美森海卡类似但可能有细微差异
function parseGeneralSeaCard(ws, sheetName) {
  return parseSeaCard(ws, sheetName);
}

// ── 主解析入口 ──────────────────────────────────
function parseYingmei(filePath) {
  console.log("[英美] 开始解析:", filePath);
  const wb = XLSX.readFile(filePath);
  const allResults = [];

  // 美国海派
  if (wb.SheetNames.includes("美国海派")) {
    const r = parseSeaParcel(wb.Sheets["美国海派"]);
    console.log(`  [美国海派] ${r.length} 条`);
    allResults.push(...r);
  }

  // 美森海卡
  if (wb.SheetNames.includes("美森海卡")) {
    const r = parseSeaCard(wb.Sheets["美森海卡"], "美森海卡");
    console.log(`  [美森海卡] ${r.length} 条`);
    allResults.push(...r);
  }

  // 合德以星海卡
  if (wb.SheetNames.includes("合德以星海卡")) {
    const r = parseSeaCard(wb.Sheets["合德以星海卡"], "合德以星海卡");
    console.log(`  [合德以星海卡] ${r.length} 条`);
    allResults.push(...r);
  }

  // 美西普船海卡渠道
  if (wb.SheetNames.includes("美西普船海卡渠道")) {
    const r = parseGeneralSeaCard(wb.Sheets["美西普船海卡渠道"], "美西普船海卡渠道");
    console.log(`  [美西普船海卡渠道] ${r.length} 条`);
    allResults.push(...r);
  }

  // 美中普船海卡渠道
  if (wb.SheetNames.includes("美中普船海卡渠道")) {
    const r = parseGeneralSeaCard(wb.Sheets["美中普船海卡渠道"], "美中普船海卡渠道");
    console.log(`  [美中普船海卡渠道] ${r.length} 条`);
    allResults.push(...r);
  }

  // 美东普船海卡渠道
  if (wb.SheetNames.includes("美东普船海卡渠道")) {
    const r = parseGeneralSeaCard(wb.Sheets["美东普船海卡渠道"], "美东普船海卡渠道");
    console.log(`  [美东普船海卡渠道] ${r.length} 条`);
    allResults.push(...r);
  }

  // 美东纽约快线
  if (wb.SheetNames.includes("美东纽约快线")) {
    const r = parseGeneralSeaCard(wb.Sheets["美东纽约快线"], "美东纽约快线");
    console.log(`  [美东纽约快线] ${r.length} 条`);
    allResults.push(...r);
  }

  // 海外仓自提
  if (wb.SheetNames.includes("海外仓自提")) {
    const r = parseSelfPickup(wb.Sheets["海外仓自提"]);
    console.log(`  [海外仓自提] ${r.length} 条`);
    allResults.push(...r);
  }

  console.log(`[英美] 总计解析 ${allResults.length} 条价格记录`);
  return allResults;
}

module.exports = { parseYingmei };
