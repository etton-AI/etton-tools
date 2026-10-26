/**
 * POST /api/tr-to-supplier — 两个入口, 按 mode 区分:
 *   mode=ettong    入口1: 仅生成「易通系统上传」发票(源发票第 8 行插入空行), 批量 → ZIP
 *   mode=supplier  入口2: 转供应商发票(供应商 + 渠道 + 每文件独立 SO + 报关方式/送货货站可覆盖), 批量 → ZIP
 * GET  /api/tr-to-supplier?session=xxx — 下载转换结果 ZIP
 * (渠道下拉见 channels/route.ts; 报关方式/货站下拉见 options/route.ts)
 */

import { NextRequest, NextResponse } from "next/server";
import { convertTrToSupplier, buildEttonInvoice, SupplierKey } from "@/lib/tr-to-supplier";
import JSZip from "jszip";
import { tmpdir } from "os";
import { join, extname } from "path";
import { writeFileSync, unlinkSync } from "fs";

// ============================================================
// Types & store
// ============================================================

interface ResultFile {
  sourceFile: string;
  outputFile: string; // 生成的发票文件名(入口1=易通上传文件名, 入口2=供应商发票文件名)
  warehouseCode: string;
  totalBoxes: number;
  dataRows: number;
  warnings: string[];
}

interface SessionEntry {
  zipBuffer: Buffer;
  zipFileName: string;
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
const VALID_MODES = ["ettong", "supplier"];

/** ZIP 内去重文件名, 避免同名覆盖(如 "(2)" 副本) */
function uniqueName(name: string, used: Set<string>): string {
  const ext = extname(name);
  const base = name.slice(0, name.length - ext.length);
  let candidate = name;
  let i = 1;
  while (used.has(candidate)) {
    candidate = `${base}(${i})${ext}`;
    i++;
  }
  used.add(candidate);
  return candidate;
}

// ============================================================
// POST — 转换(批量 → ZIP)
// ============================================================

export async function POST(request: NextRequest) {
  const tmpPaths: string[] = [];
  try {
    cleanup();

    const formData = await request.formData();
    const rawFiles = formData.getAll("file").filter((f): f is File => f instanceof File);
    const mode = (formData.get("mode") as string) || "supplier";
    const supplier = (formData.get("supplier") as string) || "";
    const channelName = ((formData.get("channelName") as string) || "").trim();
    // 每文件独立 SO(入口2 supplier 模式, 与 file 顺序对齐; 可留空按源文件名命名)
    const customerOrderNos = formData.getAll("customerOrderNo").map((v) => String(v).trim());
    const customsOverride = ((formData.get("customsOverride") as string) || "").trim();
    const station = ((formData.get("station") as string) || "").trim();

    if (rawFiles.length === 0) {
      return NextResponse.json({ error: "请上传至少一份 TR 客户发票（.xlsx）" }, { status: 400 });
    }
    if (!VALID_MODES.includes(mode)) {
      return NextResponse.json({ error: "无效的转换模式" }, { status: 400 });
    }

    for (const f of rawFiles) {
      const lower = f.name.toLowerCase();
      if (!lower.endsWith(".xlsx") && !lower.endsWith(".xls")) {
        return NextResponse.json({ error: `仅支持 .xlsx 格式的 TR 客户发票（${f.name}）` }, { status: 400 });
      }
    }

    // 入口2 额外校验
    if (mode === "supplier") {
      if (!VALID_SUPPLIERS.includes(supplier as SupplierKey)) {
        return NextResponse.json({ error: "请选择有效的供应商" }, { status: 400 });
      }
      if (!channelName) {
        return NextResponse.json({ error: "请选择渠道" }, { status: 400 });
      }
    }

    const zip = new JSZip();
    const outputNames = new Set<string>();
    const files: ResultFile[] = [];

    for (let i = 0; i < rawFiles.length; i++) {
      const f = rawFiles[i];
      const buffer = Buffer.from(await f.arrayBuffer());
      const tmpPath = join(
        tmpdir(),
        `tr_upload_${Date.now()}_${Math.random().toString(36).slice(2, 8)}.xlsx`
      );
      writeFileSync(tmpPath, buffer);
      tmpPaths.push(tmpPath);

      if (mode === "ettong") {
        // 入口1: 仅生成易通系统上传发票(原第 8 行插入空行)
        const ettongBuffer = await buildEttonInvoice(tmpPath);
        const name = uniqueName(f.name, outputNames);
        zip.file(name, ettongBuffer);
        files.push({
          sourceFile: f.name,
          outputFile: name,
          warehouseCode: "",
          totalBoxes: 0,
          dataRows: 0,
          warnings: [],
        });
      } else {
        // 入口2: 转供应商发票(每文件独立 SO + 报关方式/货站可覆盖)
        const so = customerOrderNos[i] ?? "";
        const result = await convertTrToSupplier(
          tmpPath,
          f.name,
          supplier as SupplierKey,
          so,
          channelName,
          customsOverride || undefined,
          station || undefined
        );
        const name = uniqueName(result.fileName, outputNames);
        zip.file(name, result.buffer);
        files.push({
          sourceFile: f.name,
          outputFile: name,
          warehouseCode: result.warehouseCode,
          totalBoxes: result.totalBoxes,
          dataRows: result.dataRows,
          warnings: result.warnings,
        });
      }
    }

    const zipBuffer = Buffer.from(await zip.generateAsync({ type: "nodebuffer" }));
    const stamp = new Date().toISOString().replace(/[-:T]/g, "").slice(0, 14);
    const zipFileName = mode === "ettong" ? `易通系统上传_${stamp}.zip` : `TR转${supplier}_${stamp}.zip`;

    const sessionId = makeSessionId();
    store.set(sessionId, { zipBuffer, zipFileName, createdAt: Date.now() });

    console.log(`✅ TR convert [${mode}]: ${files.length} 份发票 → ${zipFileName}`);

    return NextResponse.json({
      sessionId,
      result: {
        zipFileName,
        mode,
        supplier: mode === "supplier" ? supplier : "",
        channel: mode === "supplier" ? channelName : "",
        total: files.length,
        files,
        downloadUrl: `/api/tr-to-supplier?session=${sessionId}`,
      },
    });
  } catch (error) {
    console.error("TR convert error:", error);
    return NextResponse.json({ error: errMsg(error) }, { status: 500 });
  } finally {
    for (const p of tmpPaths) {
      try { unlinkSync(p); } catch {}
    }
  }
}

// ============================================================
// GET — 下载 ZIP
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

  const encodedName = encodeURIComponent(entry.zipFileName);
  const asciiName = entry.zipFileName.replace(/[^\x00-\x7F]/g, "_");
  return new NextResponse(new Uint8Array(entry.zipBuffer), {
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
