/**
 * POST /api/ls-to-supplier — 上传朗胜易通下单发票 + SO-SKU 对应表，按 SKU 拆分为美线发票下单映射(ZIP)
 * GET  /api/ls-to-supplier?session=xxx — 下载转换结果 ZIP
 */

import { NextRequest, NextResponse } from "next/server";
import { convertLsToSupplier, LsConvertResult } from "@/lib/ettong-to-supplier";
import { tmpdir } from "os";
import { join } from "path";
import { writeFileSync, unlinkSync } from "fs";

// ============================================================
// Types & store
// ============================================================

interface SessionEntry {
  result: LsConvertResult;
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
  let srcPath = "";
  let soSkuPath = "";
  try {
    cleanup();

    const formData = await request.formData();
    const rawFile = formData.get("file");
    const soSkuFile = formData.get("soSkuFile");
    const channelName = ((formData.get("channelName") as string) || "").trim();
    const pickupMethod = ((formData.get("pickupMethod") as string) || "").trim();
    const deliverySite = ((formData.get("deliverySite") as string) || "").trim();

    if (!(rawFile instanceof File)) {
      return NextResponse.json({ error: "请上传朗胜易通下单发票（.xlsx）" }, { status: 400 });
    }
    if (!(soSkuFile instanceof File)) {
      return NextResponse.json({ error: "请上传 SO-SKU 对应表（.xlsx）" }, { status: 400 });
    }
    if (!channelName) {
      return NextResponse.json({ error: "请填写服务渠道" }, { status: 400 });
    }

    // 写出两个临时文件
    srcPath = join(
      tmpdir(),
      `ls_src_${Date.now()}_${Math.random().toString(36).slice(2, 8)}.xlsx`
    );
    soSkuPath = join(
      tmpdir(),
      `ls_sosku_${Date.now()}_${Math.random().toString(36).slice(2, 8)}.xlsx`
    );
    writeFileSync(srcPath, Buffer.from(await rawFile.arrayBuffer()));
    writeFileSync(soSkuPath, Buffer.from(await soSkuFile.arrayBuffer()));

    const result = await convertLsToSupplier(
      srcPath,
      rawFile.name,
      soSkuPath,
      {
        channelName,
        pickupMethod: pickupMethod || "自送货",
        deliverySite: deliverySite || "清溪仓",
      }
    );

    const sessionId = makeSessionId();
    store.set(sessionId, { result, createdAt: Date.now() });

    console.log(
      `✅ LS convert: ${rawFile.name} → ${result.files.length} 张美线发票 (${result.skipped.length} 跳过)`
    );

    return NextResponse.json({
      sessionId,
      result: {
        sourceFile: result.sourceFile,
        zipFileName: result.zipFileName,
        total: result.files.length,
        files: result.files.map((f) => ({ so: f.so, sku: f.sku, fileName: f.fileName })),
        skipped: result.skipped,
        warnings: result.warnings,
        downloadUrl: `/api/ls-to-supplier?session=${sessionId}`,
      },
    });
  } catch (error) {
    console.error("LS convert error:", error);
    return NextResponse.json({ error: errMsg(error) }, { status: 500 });
  } finally {
    try { if (srcPath) unlinkSync(srcPath); } catch {}
    try { if (soSkuPath) unlinkSync(soSkuPath); } catch {}
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
