/**
 * POST /api/yx-to-tiantu — 上传延讯易通标准模板（单/多文件），转换为指定供应商发票
 * GET  /api/yx-to-tiantu?session=xxx — 下载转换结果（单文件 xlsx / 批量 zip）
 * GET  /api/yx-to-tiantu?supplier=天图 — 获取该供应商渠道下拉列表
 */

import { NextRequest, NextResponse } from "next/server";
import {
  convertYxToSupplier,
  convertYxBatch,
  parseYxChannels,
  YX_SUPPLIER_KEYS,
  YxValidationError,
  YxConvertResult,
  YxBatchResult,
  YxSupplierKey,
} from "@/lib/ettong-to-supplier";
import { tmpdir } from "os";
import { join } from "path";
import { writeFileSync, unlinkSync } from "fs";

// ============================================================
// Types & store
// ============================================================

type SessionEntry =
  | { kind: "single"; result: YxConvertResult; createdAt: number }
  | { kind: "batch"; result: YxBatchResult; createdAt: number };

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
  const tmpPaths: string[] = [];
  try {
    cleanup();

    const formData = await request.formData();
    const rawFiles = formData.getAll("files");
    const supplier = ((formData.get("supplier") as string) || "").trim();
    const channelName = ((formData.get("channelName") as string) || "").trim();
    const customerOrderNo = ((formData.get("customerOrderNo") as string) || "").trim();

    const files: File[] = rawFiles.filter((f): f is File => f instanceof File);
    if (files.length === 0) {
      return NextResponse.json({ error: "请上传延讯易通标准模板（.xlsx）" }, { status: 400 });
    }
    if (!supplier || !YX_SUPPLIER_KEYS.includes(supplier as YxSupplierKey)) {
      return NextResponse.json({ error: "请选择有效的供应商" }, { status: 400 });
    }
    if (!channelName) {
      return NextResponse.json({ error: "请选择服务渠道" }, { status: 400 });
    }

    for (const f of files) {
      const lower = f.name.toLowerCase();
      if (!lower.endsWith(".xlsx") && !lower.endsWith(".xls")) {
        return NextResponse.json({ error: `「${f.name}」不是支持的 .xlsx/.xls 格式` }, { status: 400 });
      }
    }

    // 写临时文件
    const parsed: { path: string; name: string }[] = [];
    for (const f of files) {
      const tmpPath = join(
        tmpdir(),
        `yx_upload_${Date.now()}_${Math.random().toString(36).slice(2, 8)}.xlsx`
      );
      writeFileSync(tmpPath, Buffer.from(await f.arrayBuffer()));
      tmpPaths.push(tmpPath);
      parsed.push({ path: tmpPath, name: f.name });
    }

    const opts = { customerOrderNo, channelName };
    const sessionId = makeSessionId();

    if (parsed.length === 1) {
      const result = await convertYxToSupplier(parsed[0].path, parsed[0].name, supplier as YxSupplierKey, opts);
      store.set(sessionId, { kind: "single", result, createdAt: Date.now() });
      console.log(`✅ YX convert: ${parsed[0].name} → ${result.fileName} (${supplier})`);
      return NextResponse.json({
        sessionId,
        kind: "single",
        result: {
          sourceFile: result.sourceFile,
          fileName: result.fileName,
          supplier: result.supplier,
          customerOrderNo: result.customerOrderNo,
          warehouseCode: result.warehouseCode,
          totalBoxes: result.totalBoxes,
          dataRows: result.dataRows,
          warnings: result.warnings,
          downloadUrl: `/api/yx-to-tiantu?session=${sessionId}`,
        },
      });
    }

    const result = await convertYxBatch(parsed, supplier as YxSupplierKey, opts);
    store.set(sessionId, { kind: "batch", result, createdAt: Date.now() });
    console.log(`✅ YX batch: ${parsed.length} files → ${result.zipFileName} (${supplier})`);
    return NextResponse.json({
      sessionId,
      kind: "batch",
      result: {
        zipFileName: result.zipFileName,
        supplier: result.supplier,
        country: result.country,
        files: result.files,
        warnings: result.warnings,
        downloadUrl: `/api/yx-to-tiantu?session=${sessionId}`,
      },
    });
  } catch (error) {
    console.error("YX convert error:", error);
    const status = error instanceof YxValidationError ? 400 : 500;
    return NextResponse.json({ error: errMsg(error) }, { status });
  } finally {
    for (const p of tmpPaths) {
      try { unlinkSync(p); } catch {}
    }
  }
}

// ============================================================
// GET — 下载 / 渠道列表
// ============================================================

export async function GET(request: NextRequest) {
  cleanup();

  const { searchParams } = new URL(request.url);

  // 渠道下拉
  const supplier = searchParams.get("supplier");
  if (supplier) {
    const channels = await parseYxChannels(supplier as YxSupplierKey);
    return NextResponse.json({ channels });
  }

  // 下载
  const sessionId = searchParams.get("session");
  if (!sessionId) {
    return NextResponse.json({ error: "Missing session parameter" }, { status: 400 });
  }

  const entry = store.get(sessionId);
  if (!entry) {
    return NextResponse.json({ error: "会话已过期，请重新上传文件" }, { status: 404 });
  }

  if (entry.kind === "batch") {
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
