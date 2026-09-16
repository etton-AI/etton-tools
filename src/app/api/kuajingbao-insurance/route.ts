/**
 * POST /api/kuajingbao-insurance — 上传系统货箱清单，生成跨境堡批量投保箱单
 * GET  /api/kuajingbao-insurance?session=xxx — 下载生成的投保箱单
 */

import { NextRequest, NextResponse } from "next/server";
import {
  processKuajingbaoInsurance,
  kuajingbaoOutputName,
  type KuajingbaoConfig,
  type KuajingbaoResult,
} from "@/lib/kuajingbao-insurance";
import { tmpdir } from "os";
import { join } from "path";
import { writeFileSync, unlinkSync } from "fs";

interface SessionEntry {
  result: KuajingbaoResult;
  outputName: string;
  createdAt: number;
}

const store = new Map<string, SessionEntry>();

function cleanup() {
  const cutoff = Date.now() - 30 * 60 * 1000;
  for (const [k, v] of store) {
    if (v.createdAt < cutoff) store.delete(k);
  }
}

function makeSessionId(): string {
  return Math.random().toString(36).slice(2, 10) + Date.now().toString(36);
}

/** 从表单取配置，缺省用示例默认值 */
function parseConfig(formData: FormData): KuajingbaoConfig {
  const str = (key: string) => String(formData.get(key) ?? "").trim();
  return {
    shipDate: str("shipDate"),
    transportMode: str("transportMode") || "海运",
    deliveryMethod: str("deliveryMethod") || "卡车派",
    originPlace: str("originPlace") || "深圳",
    goodsCategory: str("goodsCategory") || "无易碎品",
    packageType: str("packageType") || "纸箱",
  };
}

export async function POST(request: NextRequest) {
  try {
    cleanup();

    const formData = await request.formData();
    const file = formData.get("file");

    if (!file || !(file instanceof File)) {
      return NextResponse.json({ error: "请上传系统导出的货箱清单 Excel 文件" }, { status: 400 });
    }
    if (!file.name.endsWith(".xlsx") && !file.name.endsWith(".xls")) {
      return NextResponse.json({ error: "只支持 .xlsx 或 .xls 格式的文件" }, { status: 400 });
    }

    const config = parseConfig(formData);
    if (!config.shipDate) {
      return NextResponse.json({ error: "请填写起运日期" }, { status: 400 });
    }

    console.log(`📖 跨境堡投保: ${file.name} (${(file.size / 1024).toFixed(1)} KB)`);

    const arrayBuffer = await file.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);
    const tmpPath = join(tmpdir(), `kuajingbao_${Date.now()}.xlsx`);
    writeFileSync(tmpPath, buffer);

    let result: KuajingbaoResult;
    try {
      result = await processKuajingbaoInsurance(tmpPath, file.name, config);
    } finally {
      try { unlinkSync(tmpPath); } catch {}
    }

    const sessionId = makeSessionId();
    const outputName = kuajingbaoOutputName(file.name);
    store.set(sessionId, { result, outputName, createdAt: Date.now() });

    console.log(`✅ 跨境堡投保: ${result.groups.length} 票完成`);

    return NextResponse.json({
      sessionId,
      sourceFile: result.sourceFile,
      outputName,
      totalGroups: result.groups.length,
      lockWarehouseCount: result.lockWarehouseCount,
      groups: result.groups.map((g) => ({
        so: g.so,
        country: g.country,
        warehouseCode: g.warehouseCode,
        goodsDescription: g.goodsDescription,
        currency: g.currency,
        totalValue: g.totalValue,
        totalBoxes: g.totalBoxes,
        totalWeight: g.totalWeight,
        remark: g.remark,
      })),
      downloadUrl: `/api/kuajingbao-insurance?session=${sessionId}`,
    });
  } catch (error) {
    console.error("跨境堡投保 error:", error);
    return NextResponse.json(
      { error: error instanceof Error ? error.message : "生成投保箱单失败，请确认上传的是系统货箱清单。" },
      { status: 500 }
    );
  }
}

export async function GET(request: NextRequest) {
  cleanup();

  const { searchParams } = new URL(request.url);
  const sessionId = searchParams.get("session");

  if (!sessionId) {
    return NextResponse.json({ error: "Missing session parameter" }, { status: 400 });
  }

  const entry = store.get(sessionId);
  if (!entry) {
    return NextResponse.json({ error: "会话已过期，请重新上传文件" }, { status: 404 });
  }

  const encoded = encodeURIComponent(entry.outputName);
  const asciiName = entry.outputName.replace(/[^\x00-\x7F]/g, "_");

  return new NextResponse(new Uint8Array(entry.result.buffer), {
    headers: {
      "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "Content-Disposition": `attachment; filename="${asciiName}"; filename*=UTF-8''${encoded}`,
    },
  });
}
