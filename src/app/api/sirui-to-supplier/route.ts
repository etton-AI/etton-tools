/**
 * POST /api/sirui-to-supplier — 思锐客户下单发票 → 供应商下单模板, 批量 → ZIP
 *   供应商 + 渠道 + 每文件独立客户订单号 + 报关方式/送货货站可覆盖
 * GET  /api/sirui-to-supplier?session=xxx — 下载转换结果 ZIP
 * (渠道下拉见 channels/route.ts; 报关方式/货站下拉见 options/route.ts)
 */

import { NextRequest, NextResponse } from "next/server";
import { convertSiruiToSupplier, SiruiSupplierKey } from "@/lib/sirui-to-supplier";
import JSZip from "jszip";
import { tmpdir } from "os";
import { join, extname } from "path";
import { writeFileSync, unlinkSync } from "fs";

// ============================================================
// Types & store
// ============================================================

interface ResultFile {
  sourceFile: string;
  outputFile: string; // 生成的供应商发票文件名
  customerOrderNo: string;
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

const VALID_SUPPLIERS: SiruiSupplierKey[] = ["华威尔", "美琦", "英美", "航乐"];

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
    const supplier = (formData.get("supplier") as string) || "";
    const channelName = ((formData.get("channelName") as string) || "").trim();
    // 每文件独立客户订单号(与 file 顺序对齐; 可留空按源文件名命名)
    const customerOrderNos = formData.getAll("customerOrderNo").map((v) => String(v).trim());
    const customsOverride = ((formData.get("customsOverride") as string) || "").trim();
    const station = ((formData.get("station") as string) || "").trim();
    const pickupDate = ((formData.get("pickupDate") as string) || "").trim(); // 预计提货时间(美琦 F2)

    if (rawFiles.length === 0) {
      return NextResponse.json({ error: "请上传至少一份思锐客户下单发票（.xlsx）" }, { status: 400 });
    }
    if (!VALID_SUPPLIERS.includes(supplier as SiruiSupplierKey)) {
      return NextResponse.json({ error: "请选择有效的供应商" }, { status: 400 });
    }
    if (!channelName) {
      return NextResponse.json({ error: "请选择渠道" }, { status: 400 });
    }

    for (const f of rawFiles) {
      const lower = f.name.toLowerCase();
      if (!lower.endsWith(".xlsx") && !lower.endsWith(".xls")) {
        return NextResponse.json({ error: `仅支持 .xlsx 格式的思锐客户发票（${f.name}）` }, { status: 400 });
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
        `sirui_upload_${Date.now()}_${Math.random().toString(36).slice(2, 8)}.xlsx`
      );
      writeFileSync(tmpPath, buffer);
      tmpPaths.push(tmpPath);

      const customerOrderNo = customerOrderNos[i] ?? "";
      const result = await convertSiruiToSupplier(
        tmpPath,
        f.name,
        supplier as SiruiSupplierKey,
        customerOrderNo,
        channelName,
        customsOverride || undefined,
        station || undefined,
        pickupDate || undefined
      );
      const name = uniqueName(result.fileName, outputNames);
      zip.file(name, result.buffer);
      files.push({
        sourceFile: f.name,
        outputFile: name,
        customerOrderNo: result.customerOrderNo,
        totalBoxes: result.totalBoxes,
        dataRows: result.dataRows,
        warnings: result.warnings,
      });
    }

    const zipBuffer = Buffer.from(await zip.generateAsync({ type: "nodebuffer" }));
    const stamp = new Date().toISOString().replace(/[-:T]/g, "").slice(0, 14);
    const zipFileName = `思锐转${supplier}_${stamp}.zip`;

    const sessionId = makeSessionId();
    store.set(sessionId, { zipBuffer, zipFileName, createdAt: Date.now() });

    console.log(`✅ Sirui convert [${supplier}]: ${files.length} 份发票 → ${zipFileName}`);

    return NextResponse.json({
      sessionId,
      result: {
        zipFileName,
        supplier,
        channel: channelName,
        total: files.length,
        files,
        downloadUrl: `/api/sirui-to-supplier?session=${sessionId}`,
      },
    });
  } catch (error) {
    console.error("Sirui convert error:", error);
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
