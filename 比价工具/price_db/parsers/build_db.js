#!/usr/bin/env node
/**
 * build_db.js — 统一构建入口
 *
 * 自动识别比价工具文件夹中所有 Excel 文件，调用对应解析器，
 * 输出统一结构的 prices.json 到 data/ 目录。
 *
 * 用法: node price_db/parsers/build_db.js
 */

const path = require("path");
const fs = require("fs");
const { parseETTON } = require("./etton_us");
const { parseTiantu } = require("./tiantu_us");
const { parseYingmei } = require("./yingmei_us");

// ── 供应商识别规则 ──
function identifySupplier(fileName) {
  const n = fileName.toLowerCase();
  // 跳过非价格表文件（船期表、出运计划等）
  if (n.includes("出运计划") || n.includes("船期") || n.includes("schedule")) return "skip";
  if (n.includes("etton") || n.includes("易通")) return "etton";
  if (n.includes("天图") || n.includes("tiantu")) return "tiantu";
  if (n.includes("英美") || n.includes("yingmei")) return "yingmei";
  return null;
}

// ── 主流程 ──
function main() {
  // 比价工具目录（parsers 的上两级）
  const baseDir = path.resolve(__dirname, "..", "..");
  console.log("📂 扫描目录:", baseDir);

  const files = fs.readdirSync(baseDir).filter((f) => f.endsWith(".xlsx"));
  console.log(`📋 发现 ${files.length} 个 Excel 文件\n`);

  const allPrices = [];
  const stats = {};

  for (const file of files) {
    const supplier = identifySupplier(file);
    if (!supplier) {
      console.log(`⏭ 跳过 (无法识别供应商): ${file}`);
      continue;
    }
    if (supplier === "skip") {
      console.log(`⏭ 跳过 (非价格表文件): ${file}`);
      continue;
    }

    const filePath = path.join(baseDir, file);
    console.log(`\n🔍 解析: ${file} → 供应商: ${supplier}`);

    let results = [];
    try {
      switch (supplier) {
        case "etton":
          results = parseETTON(filePath);
          break;
        case "tiantu":
          results = parseTiantu(filePath);
          break;
        case "yingmei":
          results = parseYingmei(filePath);
          break;
      }
    } catch (err) {
      console.error(`  ❌ 解析失败: ${err.message}`);
      console.error(err.stack);
      continue;
    }

    // 计算生效日期（从文件名提取）
    // 支持格式: 2026-06-23, 2026年6月23日, 6.25, 6月25日
    let effectiveDate = "";
    const dateMatch1 = file.match(/(\d{4})[年.-]?(\d{1,2})[月.-]?(\d{1,2})/);
    if (dateMatch1) {
      effectiveDate = `${dateMatch1[1]}-${String(dateMatch1[2]).padStart(2, "0")}-${String(dateMatch1[3]).padStart(2, "0")}`;
    } else {
      const dateMatch2 = file.match(/(\d{1,2})[.·](\d{1,2})/);
      if (dateMatch2) {
        const now = new Date();
        effectiveDate = `${now.getFullYear()}-${String(parseInt(dateMatch2[1])).padStart(2, "0")}-${String(parseInt(dateMatch2[2])).padStart(2, "0")}`;
      }
    }

    // 设置生效日期和来源文件
    for (const r of results) {
      r.effective_date = effectiveDate;
      r.source_file = file;
    }

    stats[supplier] = results.length;
    allPrices.push(...results);
    console.log(`  ✅ 导入 ${results.length} 条记录 (生效日期: ${effectiveDate})`);
  }

  // ── 输出去重 ──
  // 唯一键: supplier + channel_name + destination_code + origin_region + billing_type + min_quantity
  const seen = new Set();
  const deduped = [];
  for (const r of allPrices) {
    const key = `${r.supplier}|${r.channel_name}|${r.destination_code}|${r.origin_region}|${r.billing_type}|${r.min_quantity}`;
    if (!seen.has(key)) {
      seen.add(key);
      deduped.push(r);
    }
  }
  const dupCount = allPrices.length - deduped.length;
  if (dupCount > 0) {
    console.log(`\n⚠ 去重: 移除 ${dupCount} 条重复记录`);
  }

  // ── 输出 JSON (压缩格式减小体积) ──
  const outDir = path.resolve(__dirname, "..", "data");
  if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });

  const outPath = path.join(outDir, "prices.json");
  const output = {
    generated_at: new Date().toISOString(),
    total_records: deduped.length,
    stats,
    data: deduped,
  };

  // 压缩格式写入 (无缩进，节省 ~40% 空间)
  const jsonStr = JSON.stringify(output);
  fs.writeFileSync(outPath, jsonStr, "utf-8");
  console.log(`\n💾 已保存: ${outPath} (${(Buffer.byteLength(jsonStr)/1024/1024).toFixed(1)} MB)`);

  // 同时复制到 public/data/ (供 Web 部署)
  const publicDir = path.resolve(__dirname, "..", "..", "..", "public", "data");
  if (fs.existsSync(publicDir) || fs.existsSync(path.resolve(publicDir, ".."))) {
    if (!fs.existsSync(publicDir)) fs.mkdirSync(publicDir, { recursive: true });
    const publicPath = path.join(publicDir, "prices.json");
    fs.writeFileSync(publicPath, jsonStr, "utf-8");
    console.log(`🌐 已复制到: ${publicPath}`);
  }
  console.log(`📊 总计: ${deduped.length} 条价格记录`);
  console.log(`   易通ETTON: ${stats.etton || 0} 条`);
  console.log(`   天图通逊: ${stats.tiantu || 0} 条`);
  console.log(`   英美跨境: ${stats.yingmei || 0} 条`);

  // ── 简要数据质量报告 ──
  const uniqueWarehouses = new Set(deduped.map((r) => r.destination_code));
  const uniqueChannels = new Set(deduped.map((r) => r.channel_name));
  console.log(`\n📈 数据覆盖: ${uniqueWarehouses.size} 个目的仓, ${uniqueChannels.size} 个渠道`);
}

main();
