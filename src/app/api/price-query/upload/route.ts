/**
 * 上传供应商最新价格表 Excel
 * POST /api/price-query/upload
 *
 * 接收 Excel 文件 → 自动识别供应商 → 解析 → 合并入库 → 刷新缓存
 */

import { NextRequest, NextResponse } from "next/server";
import { refreshCache, getDataPath, type PriceEntry } from "@/lib/price-store";
import fs from "fs";
import path from "path";
import os from "os";
// eslint-disable-next-line @typescript-eslint/no-require-imports
import { createRequire } from "module";

const nodeRequire = createRequire(import.meta.url);

function parseWithNode(filePath: string, supplier: string): PriceEntry[] {
  const fileName = path.basename(filePath);

  // 通过文件名自动识别供应商
  const patterns: Record<string, RegExp> = {
    etton: /etton|易通/i,
    tiantu: /天图|tiantu/i,
    yingmei: /英美|yingmei/i,
  };

  if (!supplier) {
    for (const [key, pattern] of Object.entries(patterns)) {
      if (pattern.test(fileName)) {
        supplier = key;
        break;
      }
    }
  }

  if (!supplier) {
    throw new Error(`无法识别供应商，文件名需包含 "ETTON/易通"、"天图/tiantu" 或 "英美/yingmei"`);
  }

  // 加载对应解析器（纯 JS CommonJS 模块）
  const parsersDir = path.join(process.cwd(), "比价工具", "price_db", "parsers");
  let results: PriceEntry[] = [];

  switch (supplier) {
    case "etton": {
      const mod = nodeRequire(path.join(parsersDir, "etton_us.js"));
      results = mod.parseETTON(filePath);
      break;
    }
    case "tiantu": {
      const mod = nodeRequire(path.join(parsersDir, "tiantu_us.js"));
      results = mod.parseTiantu(filePath);
      break;
    }
    case "yingmei": {
      const mod = nodeRequire(path.join(parsersDir, "yingmei_us.js"));
      results = mod.parseYingmei(filePath);
      break;
    }
    default:
      throw new Error(`未知供应商: ${supplier}`);
  }

  // 提取生效日期
  let effectiveDate = "";
  const dm1 = fileName.match(/(\d{4})[年.-]?(\d{1,2})[月.-]?(\d{1,2})/);
  if (dm1) {
    effectiveDate = `${dm1[1]}-${String(dm1[2]).padStart(2, "0")}-${String(dm1[3]).padStart(2, "0")}`;
  } else {
    const dm2 = fileName.match(/(\d{1,2})[.·](\d{1,2})/);
    if (dm2) {
      effectiveDate = `${new Date().getFullYear()}-${String(parseInt(dm2[1])).padStart(2, "0")}-${String(parseInt(dm2[2])).padStart(2, "0")}`;
    }
  }

  // 设置源文件和生效日期
  for (const r of results) {
    r.source_file = fileName;
    r.effective_date = effectiveDate;
  }

  return results;
}

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    const files = formData.getAll("files") as File[];

    if (files.length === 0) {
      return NextResponse.json({ success: false, error: "请上传至少一个 Excel 文件" }, { status: 400 });
    }

    const results: { file: string; supplier: string; count: number; effectiveDate: string }[] = [];
    const allNewRecords: PriceEntry[] = [];
    const suppliersUpdated = new Set<string>();

    // 逐个处理上传的文件
    for (const file of files) {
      if (!file.name.endsWith(".xlsx") && !file.name.endsWith(".xls")) {
        continue;
      }

      console.log(`[upload] 处理: ${file.name} (${(file.size / 1024).toFixed(0)} KB)`);

      // 保存到临时文件
      const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "price-upload-"));
      const tmpFile = path.join(tmpDir, file.name);
      const bytes = await file.arrayBuffer();
      fs.writeFileSync(tmpFile, Buffer.from(bytes));

      try {
        const records = parseWithNode(tmpFile, "");
        allNewRecords.push(...records);

        if (records.length > 0) {
          suppliersUpdated.add(records[0].supplier);
        }

        const effectiveDate = records.length > 0 ? records[0].effective_date : "";
        results.push({
          file: file.name,
          supplier: records.length > 0 ? records[0].supplier : "未知",
          count: records.length,
          effectiveDate,
        });

        console.log(`[upload] ✅ ${file.name}: ${records.length} 条 (${records[0]?.supplier})`);
      } finally {
        // 清理临时文件
        try { fs.unlinkSync(tmpFile); fs.rmdirSync(tmpDir); } catch {}
      }
    }

    if (allNewRecords.length === 0) {
      return NextResponse.json({ success: false, error: "未能从上传文件中解析到任何价格数据" }, { status: 400 });
    }

    // ── 合并到现有数据库 ──
    // 读取现有数据
    const dataPath = getDataPath();
    let existingData: PriceEntry[] = [];
    if (fs.existsSync(dataPath)) {
      const raw = JSON.parse(fs.readFileSync(dataPath, "utf-8"));
      existingData = raw.data || [];
    }

    // 去掉被更新供应商的旧数据
    const updatedSuppliers = Array.from(suppliersUpdated);
    const preserved = existingData.filter(
      (r: PriceEntry) => !updatedSuppliers.some((s) => r.supplier.includes(s) || s.includes(r.supplier))
    );

    console.log(`[upload] 保留其他供应商: ${preserved.length} 条, 新增: ${allNewRecords.length} 条`);

    // 合并且去重
    const merged = [...preserved, ...allNewRecords];
    const seen = new Set<string>();
    const deduped: PriceEntry[] = [];
    for (const r of merged) {
      const key = `${r.supplier}|${r.channel_name}|${r.destination_code}|${r.origin_region}|${r.billing_type}|${r.min_quantity}`;
      if (!seen.has(key)) {
        seen.add(key);
        deduped.push(r);
      }
    }

    // 统计
    const stats: Record<string, number> = {};
    for (const r of deduped) {
      const key = r.supplier.includes("易通") ? "etton" : r.supplier.includes("天图") ? "tiantu" : r.supplier.includes("英美") ? "yingmei" : "other";
      stats[key] = (stats[key] || 0) + 1;
    }

    // 写入文件
    const output = {
      generated_at: new Date().toISOString(),
      total_records: deduped.length,
      stats,
      data: deduped,
    };

    // 确保目录存在
    const dir = path.dirname(dataPath);
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });

    fs.writeFileSync(dataPath, JSON.stringify(output), "utf-8");

    // 刷新内存缓存
    refreshCache();

    const dupRemoved = allNewRecords.length + preserved.length - deduped.length;

    return NextResponse.json({
      success: true,
      message: `导入完成：${results.map((r) => `${r.file} (${r.supplier} ${r.count}条)`).join(", ")}`,
      files: results,
      totals: {
        new: allNewRecords.length,
        preserved,
        deduped,
        dupRemoved,
      },
      stats,
    });
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : "上传处理失败";
    console.error("[upload] 错误:", message);
    return NextResponse.json({ success: false, error: message }, { status: 500 });
  }
}
