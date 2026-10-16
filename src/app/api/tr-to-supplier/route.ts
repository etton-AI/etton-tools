/**
 * POST /api/tr-to-supplier — 上传 TR 客户发票，按所选供应商模板转换为供应商下单发票
 * GET  /api/tr-to-supplier?session=xxx — 下载转换结果
 */

import { NextRequest, NextResponse } from "next/server";
import { convertTrToSupplier, SupplierKey, TrConvertResult } from "@/lib/tr-to-supplier";
import { tmpdir } from "os";
import { join } from "path";
import { writeFileSync, unlinkSync } from "fs";

// ============================================================
// Types & store
// ============================================================

interface SessionEntry {
  result: TrConvertResult;
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

function errMsg(e: unknown): string {
  return e instanceof Error ? e.message : "处理文件时出错";
}

const VALID_SUPPLIERS: SupplierKey[] = ["凯鑫", "英美", "星链", "航乐"];

// ============================================================
// POST — 转换
// ============================================================

export async function POST(request: NextRequest) {
  try {
    cleanup();

    const formData = await request.formData();
    const rawFile = formData.get("file");
    const supplier = (formData.get("supplier") as string) || "";
    const customerOrderNo = ((formData.get("customerOrderNo") as string) || "").trim();
    const channelName = ((formData.get("channelName") as string) || "").trim();

    if (!(rawFile instanceof File)) {
      return NextResponse.json({ error: "请上传 TR 客户发票文件（.xlsx）" }, { status: 400 });
    }
    if (!VALID_SUPPLIERS.includes(supplier as SupplierKey)) {
      return NextResponse.json({ error: "请选择有效的供应商" }, { status: 400 });
    }
    if (!customerOrderNo) {
      return NextResponse.json({ error: "请填写客户订单号" }, { status: 400 });
    }

    const lower = rawFile.name.toLowerCase();
    if (!lower.endsWith(".xlsx") && !lower.endsWith(".xls")) {
      return NextResponse.json({ error: "仅支持 .xlsx 格式的 TR 客户发票" }, { status: 400 });
    }

    const buffer = Buffer.from(await rawFile.arrayBuffer());
    const tmpPath = join(
      tmpdir(),
      `tr_upload_${Date.now()}_${Math.random().toString(36).slice(2, 8)}.xlsx`
    );
    writeFileSync(tmpPath, buffer);

    let result: TrConvertResult;
    try {
      result = await convertTrToSupplier(
        tmpPath,
        rawFile.name,
        supplier as SupplierKey,
        customerOrderNo,
        channelName
      );
    } finally {
      try { unlinkSync(tmpPath); } catch {}
    }

    const sessionId = makeSessionId();
    store.set(sessionId, { result, createdAt: Date.now() });

    console.log(`✅ TR convert: ${rawFile.name} → ${supplier} / ${result.fileName}`);

    return NextResponse.json({
      sessionId,
      result: {
        sourceFile: result.sourceFile,
        fileName: result.fileName,
        supplier: result.supplier,
        customerOrderNo: result.customerOrderNo,
        warehouseCode: result.warehouseCode,
        totalBoxes: result.totalBoxes,
        dataRows: result.dataRows,
        warnings: result.warnings,
        downloadUrl: `/api/tr-to-supplier?session=${sessionId}`,
      },
    });
  } catch (error) {
    console.error("TR convert error:", error);
    return NextResponse.json({ error: errMsg(error) }, { status: 500 });
  }
}

// ============================================================
// GET — 下载
// ============================================================

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

  const { result } = entry;
  const encodedName = encodeURIComponent(result.fileName);
  const asciiName = result.fileName.replace(/[^\x00-\x7F]/g, "_");
  return new NextResponse(new Uint8Array(result.buffer), {
    headers: {
      "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "Content-Disposition": `attachment; filename="${asciiName}"; filename*=UTF-8''${encodedName}`,
    },
  });
}

// Increase body size limit for large Excel uploads
export const config = {
  api: {
    bodyParser: {
      sizeLimit: "50mb",
    },
  },
};
