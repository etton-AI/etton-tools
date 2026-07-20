/**
 * POST /api/pipixiong-split — 上传皮皮熊账单，返回拆分结果摘要
 * GET  /api/pipixiong-split?session=xxx&type=domestic|international|invoice — 下载对应账单
 */

import { NextRequest, NextResponse } from "next/server";
import { processPipixiongSplit, type PipixiongSplitResult } from "@/lib/pipixiong-split";
import { tmpdir } from "os";
import { join } from "path";
import { writeFileSync, unlinkSync } from "fs";

interface SessionEntry {
  result: PipixiongSplitResult;
  fileName: string;
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

export async function POST(request: NextRequest) {
  try {
    cleanup();

    const formData = await request.formData();
    const file = formData.get("file");

    if (!file || !(file instanceof File)) {
      return NextResponse.json({ error: "请上传皮皮熊账单 Excel 文件" }, { status: 400 });
    }
    if (!file.name.endsWith(".xlsx") && !file.name.endsWith(".xls")) {
      return NextResponse.json({ error: "只支持 .xlsx 或 .xls 格式的文件" }, { status: 400 });
    }

    console.log(`📖 Pipi split: ${file.name} (${(file.size / 1024).toFixed(1)} KB)`);

    const arrayBuffer = await file.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);
    const tmpPath = join(tmpdir(), `pipixiong_${Date.now()}.xlsx`);
    writeFileSync(tmpPath, buffer);

    let result: PipixiongSplitResult;
    try {
      result = await processPipixiongSplit(tmpPath, file.name);
    } finally {
      try { unlinkSync(tmpPath); } catch {}
    }

    const sessionId = makeSessionId();
    store.set(sessionId, { result, fileName: file.name, createdAt: Date.now() });

    const domesticCount = result.groups.filter(g => g.hasDomestic).length;
    const intlCount = result.groups.filter(g => g.hasInternational).length;

    console.log(`✅ Pipi split: ${result.groups.length} groups → 国内${domesticCount} / 国外${intlCount} / INVOICE${intlCount}`);

    return NextResponse.json({
      sessionId,
      sourceFile: result.sourceFile,
      totalGroups: result.groups.length,
      domesticCount,
      internationalCount: intlCount,
      invoiceCount: intlCount,
      groups: result.groups.map((g) => ({
        seq: g.seq,
        forwarderId: g.forwarderId,
        customsNo: g.customsNo,
        country: g.bill.country,
        qty: g.bill.qty,
        truckFee: g.bill.truckFee,
        customsFee: g.bill.customsFee,
        portFee: g.bill.portFee,
        oceanTaxUSD: g.bill.oceanTaxUSD,
        hasDomestic: g.hasDomestic,
        hasInternational: g.hasInternational,
      })),
      downloadUrls: {
        domestic: `/api/pipixiong-split?session=${sessionId}&type=domestic`,
        international: `/api/pipixiong-split?session=${sessionId}&type=international`,
        invoice: `/api/pipixiong-split?session=${sessionId}&type=invoice`,
      },
    });
  } catch (error) {
    console.error("Pipi split error:", error);
    return NextResponse.json(
      { error: error instanceof Error ? error.message : "拆分处理失败，请确认上传的是皮皮熊账单文件。" },
      { status: 500 }
    );
  }
}

export async function GET(request: NextRequest) {
  cleanup();

  const { searchParams } = new URL(request.url);
  const sessionId = searchParams.get("session");
  const type = searchParams.get("type") || "domestic";

  if (!sessionId) {
    return NextResponse.json({ error: "Missing session parameter" }, { status: 400 });
  }

  const entry = store.get(sessionId);
  if (!entry) {
    return NextResponse.json({ error: "会话已过期，请重新上传文件" }, { status: 404 });
  }

  let buffer: Buffer;
  let typeLabel: string;

  switch (type) {
    case "international":
      buffer = entry.result.internationalBuffer;
      typeLabel = "国外账单";
      break;
    case "invoice":
      buffer = entry.result.invoiceBuffer;
      typeLabel = "INVOICE";
      break;
    default:
      buffer = entry.result.domesticBuffer;
      typeLabel = "国内账单";
  }

  const baseName = entry.fileName.replace(/\.(xlsx?|xls)$/i, "");
  const outName = `${baseName}_${typeLabel}.xlsx`;
  const asciiName = outName.replace(/[^\x00-\x7F]/g, "_");
  const encoded = encodeURIComponent(outName);

  return new NextResponse(new Uint8Array(buffer), {
    headers: {
      "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "Content-Disposition": `attachment; filename="${asciiName}"; filename*=UTF-8''${encoded}`,
    },
  });
}
