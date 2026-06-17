/**
 * POST /api/split-insurance — 上传 Excel，返回拆分结果摘要 + 下载链接
 * GET  /api/split-insurance?session=xxx&file=xxx — 下载单个文件或 all.zip
 */

import { NextRequest, NextResponse } from "next/server";
import { processSplit, generateZip, SplitResult } from "@/lib/split-insurance";
import { tmpdir } from "os";
import { join } from "path";
import { writeFileSync, unlinkSync } from "fs";

// In-memory session store (LAN tool, bounded usage)
const store = new Map<string, { result: SplitResult; createdAt: number }>();

// Cleanup entries older than 30 minutes
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
      return NextResponse.json(
        { error: "请上传一个 .xlsx 文件" },
        { status: 400 }
      );
    }

    if (
      !file.name.endsWith(".xlsx") &&
      !file.name.endsWith(".xls")
    ) {
      return NextResponse.json(
        { error: "只支持 .xlsx 格式的文件" },
        { status: 400 }
      );
    }

    const arrayBuffer = await file.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);

    console.log(`📖 Processing: ${file.name} (${(buffer.length / 1024).toFixed(1)} KB)`);

    // Write to temp file first — exceljs load(buffer) can fail in Next.js edge
    // runtime, but readFile always works.
    const tmpDir = tmpdir();
    const tmpPath = join(tmpDir, `upload_${Date.now()}.xlsx`);
    writeFileSync(tmpPath, buffer);

    let result;
    try {
      result = await processSplit(tmpPath, file.name);
    } finally {
      try { unlinkSync(tmpPath); } catch {}
    }

    const sessionId = makeSessionId();
    store.set(sessionId, { result, createdAt: Date.now() });

    console.log(
      `✅ Split complete: ${result.totalBoxes} boxes, ${result.totalGroups} groups → ${result.intervals.filter((i) => i.totalBoxes > 0).length} non-empty intervals`
    );

    // Return summary WITHOUT buffer data (too large for JSON)
    return NextResponse.json({
      sessionId,
      sourceFile: result.sourceFile,
      totalBoxes: result.totalBoxes,
      totalGroups: result.totalGroups,
      intervals: result.intervals.map((iv) => ({
        name: iv.name,
        fileName: iv.fileName,
        totalBoxes: iv.totalBoxes,
        totalWeight: iv.totalWeight,
        totalVolume: iv.totalVolume,
        groupCount: iv.groupCount,
        groups: iv.groups,
      })),
      downloads: {
        allZip: `/api/split-insurance?session=${sessionId}&file=all.zip`,
        files: result.intervals.map(
          (iv) =>
            `/api/split-insurance?session=${sessionId}&file=${encodeURIComponent(iv.fileName)}`
        ),
      },
    });
  } catch (error) {
    console.error("Split error:", error);
    return NextResponse.json(
      {
        error:
          error instanceof Error
            ? error.message
            : "处理文件时出错，请确认上传的是易通下单模板文件。",
      },
      { status: 500 }
    );
  }
}

export async function GET(request: NextRequest) {
  cleanup();

  const { searchParams } = new URL(request.url);
  const sessionId = searchParams.get("session");
  const fileName = searchParams.get("file");

  if (!sessionId || !fileName) {
    return NextResponse.json(
      { error: "Missing session or file parameter" },
      { status: 400 }
    );
  }

  const entry = store.get(sessionId);
  if (!entry) {
    return NextResponse.json(
      { error: "会话已过期，请重新上传文件" },
      { status: 404 }
    );
  }

  if (fileName === "all.zip") {
    const zipBuffer = await generateZip(entry.result);
    return new NextResponse(new Uint8Array(zipBuffer), {
      headers: {
        "Content-Type": "application/zip",
        "Content-Disposition": `attachment; filename="保单拆分结果.zip"`,
      },
    });
  }

  const interval = entry.result.intervals.find(
    (iv) => iv.fileName === fileName
  );
  if (!interval) {
    return NextResponse.json(
      { error: `文件 "${fileName}" 不存在` },
      { status: 404 }
    );
  }

  return new NextResponse(new Uint8Array(interval.buffer), {
    headers: {
      "Content-Type":
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "Content-Disposition": `attachment; filename*=UTF-8''${encodeURIComponent(fileName)}`,
    },
  });
}

// Increase body size limit to 50MB for large Excel uploads
// (App Router Node.js runtime supports api.bodyParser.sizeLimit)
export const config = {
  api: {
    bodyParser: {
      sizeLimit: "50mb",
    },
  },
};