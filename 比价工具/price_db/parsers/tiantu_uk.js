/**
 * 天图通逊 — 英国线价格解析器
 *
 * 英国线 Sheet 布局与美线不同:
 * - 城市为行(Row 6+), 渠道为列组
 * - 每个渠道含: 包税价 / 不包税价 / 时效 / 赔付
 * - 仓库代码（如 BHX4）在渠道标题行中
 */

const XLSX = require("xlsx");

const SUPPLIER = "天图通逊";
const COUNTRY = "英国";

// ── 城市名映射 ──
function parseCities(cell) {
  const text = String(cell).replace(/\r?\n/g, " ").replace(/\s+/g, " ").trim();
  // "深圳/广州/中山/东莞南城/汕头 (华南)" → ["深圳","广州","中山","东莞南城","汕头"]
  const cleaned = text.replace(/\s*[（(].*?[）)]/g, "").trim();
  return cleaned.split(/[\/,，]/).map(s => s.trim()).filter(s => s && s.length > 0);
}

// ── 解析城市行 Sheet ──
// 适用: 英国铁运专线, 英国中英专车, 苏新号, 英国海运-海派
// 结构: R4=渠道名, R5=子头(包税/不包税/时效/赔付), R6+=城市行
function parseCityRowSheet(ws, sheetName, config) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 6) return [];

  const titleRow = data[config.titleRow || 3] || []; // 渠道名行
  const subRow = data[config.subRow || 4] || [];      // 包税/不包税/时效行
  const results = [];

  // 在 titleRow 中找渠道名
  const channels = [];
  for (let col = 2; col < titleRow.length; col++) {
    const cell = String(titleRow[col] || "").trim();
    if (!cell || cell.length < 3) continue;
    channels.push({ name: cell.replace(/\r?\n.*$/, "").trim(), startCol: col });
  }

  // 为每个渠道确定列: 包税col, 不包税col, 时效col, 赔付col
  for (const ch of channels) {
    // 在 ch.startCol 开始找子列
    let taxCol = -1, noTaxCol = -1, transitCol = -1, claimCol = -1;
    for (let c = ch.startCol; c < Math.min(ch.startCol + 6, subRow.length); c++) {
      const cell = String(subRow[c] || "").toLowerCase();
      if (cell.includes("包税") && taxCol < 0) taxCol = c;
      else if (cell.includes("不包税") && noTaxCol < 0) noTaxCol = c;
      else if ((cell.includes("时效") || cell.includes("签收")) && transitCol < 0) transitCol = c;
      else if ((cell.includes("赔付") || cell.includes("延时")) && claimCol < 0) claimCol = c;
    }

    // 解析城市行
    const startRow = config.dataStartRow || 5;
    for (let ri = startRow; ri < data.length; ri++) {
      const row = data[ri];
      const cityCell = String(row[0] || "").trim();
      const weightCell = String(row[1] || "").trim();

      if (!cityCell || cityCell.includes("赔偿标准") || cityCell.includes("线路图")) continue;
      if (cityCell.includes("仓库代码")) continue;

      const cities = parseCities(cityCell);
      if (cities.length === 0) continue;

      const transitText = transitCol >= 0 ? String(row[transitCol] || "").trim() : "";
      const claimText = claimCol >= 0 ? String(row[claimCol] || "").trim() : "";
      const transitMatch = transitText.match(/(\d+)[-–](\d+)/);
      const transitMin = transitMatch ? parseInt(transitMatch[1]) : null;
      const transitMax = transitMatch ? parseInt(transitMatch[2]) : transitMin;

      // 从渠道名推断运输方式
      const tm = inferTransportMode(ch.name);

      // 包税价
      if (taxCol >= 0) {
        const price = parseFloat(row[taxCol]);
        if (!isNaN(price) && price > 0) {
          const whCodes = inferWarehouses(ch.name, sheetName);
          for (const wh of whCodes) {
            results.push(makeRecord(SUPPLIER, ch.name, tm, wh, "warehouse", cityCell, cities, "包税", weightCell, price, "元/KG", transitMin, transitMax, transitText, claimText, sheetName));
          }
        }
      }
      // 不包税价
      if (noTaxCol >= 0) {
        const price = parseFloat(row[noTaxCol]);
        if (!isNaN(price) && price > 0) {
          const whCodes = inferWarehouses(ch.name, sheetName);
          for (const wh of whCodes) {
            results.push(makeRecord(SUPPLIER, ch.name, tm, wh, "warehouse", cityCell, cities, "VAT自税", weightCell, price, "元/KG", transitMin, transitMax, transitText, claimText, sheetName));
          }
        }
      }
    }
  }

  return results;
}

// ── 解析仓库行 Sheet (海运-卡派) ──
function parseWarehouseRowSheet(ws, sheetName) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 7) return [];

  // R4 (idx 3): 渠道名
  // R5 (idx 4): 价格类型
  // R6 (idx 5): 城市列头
  const titleRow = data[3] || [];
  const typeRow = data[4] || [];
  const cityRow = data[5] || [];

  const channelName = String(titleRow[0] || titleRow[1] || titleRow[7] || "").trim().replace(/\r?\n.*/, "");
  const results = [];

  // 找包税城市列和不包税城市列
  const taxCities = [], noTaxCities = [];
  for (let col = 1; col < cityRow.length; col++) {
    const city = String(cityRow[col] || "").trim();
    const type = String(typeRow[col] || "").toLowerCase();
    if (!city) continue;
    const cities = parseCities(city);
    if (type.includes("包税") || (!type.includes("不包税") && col < 14)) {
      taxCities.push({ col, cities, region: city });
    }
    if (type.includes("不包税")) {
      noTaxCities.push({ col, cities, region: city });
    }
  }

  const tm = inferTransportMode(channelName);

  for (let ri = 6; ri < data.length; ri++) {
    const row = data[ri];
    const wh = String(row[0] || "").trim();
    if (!wh || wh.length < 3) continue;

    for (const tc of taxCities) {
      const price = parseFloat(row[tc.col]);
      if (!isNaN(price) && price > 0) {
        results.push(makeRecord(SUPPLIER, channelName, tm, wh, "warehouse", tc.region, tc.cities, "包税", "100KG+", price, "元/KG", null, null, "", "", sheetName));
      }
    }
    for (const ntc of noTaxCities) {
      const price = parseFloat(row[ntc.col]);
      if (!isNaN(price) && price > 0) {
        results.push(makeRecord(SUPPLIER, channelName, tm, wh, "warehouse", ntc.region, ntc.cities, "VAT自税", "100KG+", price, "元/KG", null, null, "", "", sheetName));
      }
    }
  }

  return results;
}

// ── 辅助函数 ──
function inferTransportMode(channelName) {
  const n = channelName;
  if (n.includes("空运") || n.includes("提") && (n.includes("普货") || n.includes("带电"))) return "空运";
  if (n.includes("卡航") || n.includes("苏新号")) return "卡航";
  if (n.includes("铁") || n.includes("班列")) return "铁路";
  if (n.includes("专车")) return "中英专车";
  if (n.includes("海卡") || n.includes("海派") || n.includes("海运")) return "海运";
  return "海运";
}

function inferWarehouses(channelName, sheetName) {
  const s = sheetName + channelName;
  // 英国FBA核心仓
  const ukWh = ["BHX4", "LBA4", "BHX8", "BHX7", "LBA2", "LBA8", "MAN4", "MAN8", "LTN7", "LPL2"];
  // 海运卡派 sheet 包含更多仓
  if (s.includes("海运") && s.includes("卡派")) return ukWh;
  // 铁运/卡航/专车 通常覆盖 3 个主仓
  return ["BHX4", "LBA4", "BHX8"];
}

function makeRecord(supplier, channelName, transportMode, destCode, destType, originRegion, originCities, taxMode, minQty, price, priceUnit, transitMin, transitMax, transitText, claimText, sheetName) {
  const mqv = parseInt(String(minQty).match(/(\d+)/)?.[1] || "100");
  return {
    supplier,
    country: COUNTRY,
    channel_name: channelName.replace(/\r?\n/g, " ").trim(),
    transport_mode: transportMode,
    vessel_config: transportMode,
    vessel_tags: [],
    delivery_method: channelName.includes("DPD") ? "DPD派" : "卡派",
    destination_type: destType,
    destination_code: destCode,
    destination_region: "",
    origin_region: originRegion.replace(/\r?\n/g, " ").trim(),
    origin_cities: originCities,
    billing_type: taxMode,
    tax_mode: taxMode,
    min_quantity: minQty.replace(/\r?\n/g, " ").trim(),
    min_quantity_value: mqv,
    unit_price: price,
    price_unit: priceUnit,
    transit_time_min: transitMin,
    transit_time_max: transitMax,
    transit_time_desc: transitText.replace(/\r?\n/g, " ").trim(),
    claim_rule: claimText.replace(/\r?\n/g, " ").trim(),
    effective_date: "",
    source_sheet: sheetName,
  };
}

// ── 主入口 ──
function parseTiantuUK(filePath) {
  console.log("[天图UK] 开始解析:", filePath);
  const wb = XLSX.readFile(filePath);
  const allResults = [];

  // 英国铁运专线 (city-row, 4 channels)
  if (wb.SheetNames.includes("英国铁运专线(18-23)")) {
    const r = parseCityRowSheet(wb.Sheets["英国铁运专线(18-23)"], "英国铁运专线", { titleRow: 3, subRow: 4, dataStartRow: 5 });
    console.log(`  [铁运专线] ${r.length} 条`);
    allResults.push(...r);
  }

  // 英国中英专车 (city-row, 4 channels, mostly paused)
  if (wb.SheetNames.includes("英国中英专车(20-25)")) {
    const r = parseCityRowSheet(wb.Sheets["英国中英专车(20-25)"], "中英专车", { titleRow: 3, subRow: 4, dataStartRow: 5 });
    console.log(`  [中英专车] ${r.length} 条`);
    allResults.push(...r);
  }

  // 苏新号-英国卡航 (city-row, 2 channels)
  if (wb.SheetNames.includes("苏新号-英国卡航20日达")) {
    const r = parseCityRowSheet(wb.Sheets["苏新号-英国卡航20日达"], "苏新号卡航", { titleRow: 3, subRow: 4, dataStartRow: 5 });
    console.log(`  [苏新号卡航] ${r.length} 条`);
    allResults.push(...r);
  }

  // 英国海运-海派 (city-row, 3 channels)
  if (wb.SheetNames.includes("英国海运-海派(23-40)")) {
    const r = parseCityRowSheet(wb.Sheets["英国海运-海派(23-40)"], "海运海派", { titleRow: 3, subRow: 4, dataStartRow: 5 });
    console.log(`  [海运海派] ${r.length} 条`);
    allResults.push(...r);
  }

  // 英国海运-卡派 (warehouse-row, special structure)
  if (wb.SheetNames.includes("英国海运-卡派(23-40) ")) {
    const r = parseWarehouseRowSheet(wb.Sheets["英国海运-卡派(23-40) "], "海运卡派");
    console.log(`  [海运卡派] ${r.length} 条`);
    allResults.push(...r);
  }

  // 英国商私地址
  if (wb.SheetNames.includes("英国商私地址")) {
    const r = parseCityRowSheet(wb.Sheets["英国商私地址"], "商私地址", { titleRow: 3, subRow: 4, dataStartRow: 6 });
    console.log(`  [商私地址] ${r.length} 条`);
    allResults.push(...r);
  }

  console.log(`[天图UK] 总计 ${allResults.length} 条`);
  return allResults;
}

module.exports = { parseTiantuUK };
