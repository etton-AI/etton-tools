/**
 * POST /api/yanxun-convert — 批量上传延讯下单发票（多文件或 ZIP 包），转换为易通下单模版
 * GET  /api/yanxun-convert?session=xxx&file=yyy — 下载单个转换结果，或 file=all.zip 下载打包结果
 */

import { NextRequest, NextResponse } from "next/server";
import {
  convertYanxunToEtton,
  generateYanxunZip,
  YanxunConvertResult,
} from "@/lib/yanxun-convert";
import JSZip from "jszip";
import { tmpdir } from "os";
import { join } from "path";
import { writeFileSync, unlinkSync } from "fs";

// ============================================================
// Types
// ============================================================

interface FailedItem {
  sourceFile: string;
  error: string;
}

interface SessionEntry {
  results: YanxunConvertResult[];
  failed: FailedItem[];
  createdAt: number;
}

// In-memory session store (30-min expiry)
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
// Helpers
// ============================================================

/** 从 ZIP 包中提取所有 .xlsx/.xls 文件 */
async function extractXlsxFromZip(file: File): Promise<{ buffer: Buffer; name: string }[]> {
  const arrayBuffer = await file.arrayBuffer();
  const zip = await JSZip.loadAsync(arrayBuffer);
  const result: { buffer: Buffer; name: string }[] = [];

  for (const [entryPath, entry] of Object.entries(zip.files)) {
    if (entry.dir) continue;
    if (!/\.(xlsx|xls)$/i.test(entryPath)) continue;

    const name = entryPath.split("/").pop() || entryPath;
    // 跳过 macOS 元数据、临时文件和隐藏文件
    if (name.startsWith(".") || name.startsWith("~$")) continue;

    const buffer = Buffer.from(await entry.async("nodebuffer"));
    result.push({ buffer, name });
  }
  return result;
}

/** 转换单个文件 buffer */
async function convertOne(buffer: Buffer, name: string): Promise<YanxunConvertResult> {
  const tmpPath = join(
    tmpdir(),
    `yanxun_upload_${Date.now()}_${Math.random().toString(36).slice(2, 8)}.xlsx`
  );
  writeFileSync(tmpPath, buffer);
  try {
    return await convertYanxunToEtton(tmpPath, name);
  } finally {
    try { unlinkSync(tmpPath); } catch {}
  }
}

// ============================================================
// POST — 批量转换
// ============================================================

export async function POST(request: NextRequest) {
  try {
    cleanup();

    const formData = await request.formData();
    const rawFiles = formData.getAll("files");
    const files = rawFiles.filter((f): f is File => f instanceof File);

    if (files.length === 0) {
      return NextResponse.json(
        { error: "请上传至少一个延讯下单发票文件（支持多选或 ZIP 包）" },
        { status: 400 }
      );
    }

    console.log(`📖 Yanxun convert: 收到 ${files.length} 个文件`);

    const success: YanxunConvertResult[] = [];
    const failed: FailedItem[] = [];

    for (const file of files) {
      const lower = file.name.toLowerCase();

      try {
        if (lower.endsWith(".zip")) {
          // ZIP 输入：解压后逐个转换
          const innerFiles = await extractXlsxFromZip(file);
          if (innerFiles.length === 0) {
            failed.push({ sourceFile: file.name, error: "ZIP 包内没有找到 .xlsx/.xls 文件" });
            continue;
          }
          console.log(`📦 ${file.name}: 解压出 ${innerFiles.length} 个发票文件`);
          for (const inner of innerFiles) {
            try {
              const r = await convertOne(inner.buffer, inner.name);
              success.push(r);
              console.log(`  ✅ ${inner.name} → ${r.fileName}`);
            } catch (e) {
              failed.push({ sourceFile: inner.name, error: errMsg(e) });
              console.log(`  ❌ ${inner.name}: ${errMsg(e)}`);
            }
          }
        } else if (lower.endsWith(".xlsx") || lower.endsWith(".xls")) {
          const r = await convertOne(Buffer.from(await file.arrayBuffer()), file.name);
          success.push(r);
          console.log(`✅ ${file.name} → ${r.fileName}`);
        } else {
          failed.push({ sourceFile: file.name, error: "不支持的文件格式（仅 .xlsx/.xls/.zip）" });
        }
      } catch (e) {
        failed.push({ sourceFile: file.name, error: errMsg(e) });
        console.log(`❌ ${file.name}: ${errMsg(e)}`);
      }
    }

    if (success.length === 0) {
      const first = failed[0];
      return NextResponse.json(
        {
          error:
            first
              ? `全部 ${failed.length} 个文件转换失败。首个错误：${first.sourceFile} — ${first.error}`
              : "没有可转换的文件",
          failed,
        },
        { status: 500 }
      );
    }

    const sessionId = makeSessionId();
    store.set(sessionId, { results: success, failed, createdAt: Date.now() });

    console.log(
      `✅ Yanxun convert 完成: ${success.length} 成功 / ${failed.length} 失败`
    );

    return NextResponse.json({
      sessionId,
      summary: {
        total: success.length + failed.length,
        success: success.length,
        failed: failed.length,
      },
      results: success.map((r) => ({
        sourceFile: r.sourceFile,
        fileName: r.fileName,
        fbaId: r.fbaId,
        totalBoxes: r.totalBoxes,
        dataRows: r.dataRows,
        mixedBoxGroups: r.mixedBoxGroups,
        topInfo: r.topInfo,
        warnings: r.warnings,
        downloadUrl: `/api/yanxun-convert?session=${sessionId}&file=${encodeURIComponent(r.fileName)}`,
      })),
      failed,
      downloadAllUrl: `/api/yanxun-convert?session=${sessionId}&file=all.zip`,
    });
  } catch (error) {
    console.error("Yanxun convert error:", error);
    return NextResponse.json(
      { error: errMsg(error) },
      { status: 500 }
    );
  }
}

// ============================================================
// GET — 下载
// ============================================================

export async function GET(request: NextRequest) {
  cleanup();

  const { searchParams } = new URL(request.url);
  const sessionId = searchParams.get("session");
  const fileName = searchParams.get("file");

  if (!sessionId || !fileName) {
    return NextResponse.json({ error: "Missing session or file parameter" }, { status: 400 });
  }

  const entry = store.get(sessionId);
  if (!entry) {
    return NextResponse.json({ error: "会话已过期，请重新上传文件" }, { status: 404 });
  }

  // ZIP 打包下载
  if (fileName === "all.zip") {
    const zipBuffer = await generateYanxunZip(entry.results);
    const zipName = `延讯下单优化结果_${entry.results.length}票.zip`;
    const asciiName = `yanxun_convert_${entry.results.length}.zip`;
    const encoded = encodeURIComponent(zipName);
    return new NextResponse(new Uint8Array(zipBuffer), {
      headers: {
        "Content-Type": "application/zip",
        "Content-Disposition": `attachment; filename="${asciiName}"; filename*=UTF-8''${encoded}`,
      },
    });
  }

  // 单个文件下载
  const result = entry.results.find((r) => r.fileName === fileName);
  if (!result) {
    return NextResponse.json({ error: `文件 "${fileName}" 不存在` }, { status: 404 });
  }

  const encodedName = encodeURIComponent(result.fileName);
  const asciiName = result.fileName.replace(/[^\x00-\x7F]/g, "_");
  return new NextResponse(new Uint8Array(result.buffer), {
    headers: {
      "Content-Type":
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
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
