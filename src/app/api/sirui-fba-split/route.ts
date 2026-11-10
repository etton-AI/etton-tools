/**
 * POST /api/sirui-fba-split — 上传内部报价表 + 下单发票，按 FBA 码拆分为多张发票(ZIP)
 * GET  /api/sirui-fba-split?session=xxx — 下载拆分结果 ZIP
 */

import { NextRequest, NextResponse } from "next/server";
import { splitInvoice, SiruiSplitResult } from "@/lib/sirui-fba-split";
import { tmpdir } from "os";
import { join } from "path";
import { writeFileSync, unlinkSync } from "fs";

// ============================================================
// Types & store
// ============================================================

interface SessionEntry {
  result: SiruiSplitResult;
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

// ============================================================
// POST — 转换
// ============================================================

export async function POST(request: NextRequest) {
  let quotePath = "";
  let invoicePath = "";
  try {
    cleanup();

    const formData = await request.formData();
    const quoteFile = formData.get("quote");
    const invoiceFile = formData.get("file");

    if (!(quoteFile instanceof File)) {
      return NextResponse.json({ error: "请上传内部报价表（.xlsx）" }, { status: 400 });
    }
    if (!(invoiceFile instanceof File)) {
      return NextResponse.json({ error: "请上传下单发票（.xlsx）" }, { status: 400 });
    }

    quotePath = join(
      tmpdir(),
      `sirui_quote_${Date.now()}_${Math.random().toString(36).slice(2, 8)}.xlsx`
    );
    invoicePath = join(
      tmpdir(),
      `sirui_invoice_${Date.now()}_${Math.random().toString(36).slice(2, 8)}.xlsx`
    );
    writeFileSync(quotePath, Buffer.from(await quoteFile.arrayBuffer()));
    writeFileSync(invoicePath, Buffer.from(await invoiceFile.arrayBuffer()));

    const result = await splitInvoice(quotePath, invoicePath, invoiceFile.name);

    const sessionId = makeSessionId();
    store.set(sessionId, { result, createdAt: Date.now() });

    console.log(
      `✅ 思锐FBA拆分: ${invoiceFile.name} → ${result.total} 张发票`
    );

    return NextResponse.json({
      sessionId,
      result: {
        zipFileName: result.zipFileName,
        total: result.total,
        files: result.files.map((f) => ({
          fba: f.fba,
          warehouse: f.warehouse,
          fileName: f.fileName,
          rowCount: f.rowCount,
        })),
        warnings: result.warnings,
        downloadUrl: `/api/sirui-fba-split?session=${sessionId}`,
      },
    });
  } catch (error) {
    console.error("Sirui FBA split error:", error);
    return NextResponse.json({ error: errMsg(error) }, { status: 500 });
  } finally {
    try { if (quotePath) unlinkSync(quotePath); } catch {}
    try { if (invoicePath) unlinkSync(invoicePath); } catch {}
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
  const encodedName = encodeURIComponent(result.zipFileName);
  const asciiName = result.zipFileName.replace(/[^\x00-\x7F]/g, "_");
  return new NextResponse(new Uint8Array(result.buffer), {
    headers: {
      "Content-Type": "application/zip",
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
