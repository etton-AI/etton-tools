/**
 * POST /api/convert-pacific — 上传货箱清单 Excel + 汇率，返回拆分结果摘要 + 下载链接
 * GET  /api/convert-pacific?session=xxx&file=xxx — 下载单个区间文件或 all.zip
 */

import { NextRequest, NextResponse } from "next/server";
import {
  convertAndSplitPacific,
  generatePacificZip,
  ExchangeRates,
  PacificSplitResult,
} from "@/lib/convert-pacific-insurance";
import { tmpdir } from "os";
import { join } from "path";
import { writeFileSync, unlinkSync } from "fs";

// In-memory session store
const store = new Map<string, { result: PacificSplitResult; createdAt: number }>();

function cleanup() {
  const cutoff = Date.now() - 30 * 60 * 1000;
  for (const [k, v] of store) {
    if (v.createdAt < cutoff) store.delete(k);
  }
}

function makeSessionId(): string {
  return Math.random().toString(36).slice(2, 10) + Date.now().toString(36);
}

function parseRates(formData: FormData): ExchangeRates {
  return {
    USD: parseFloat(formData.get("rateUSD") as string) || 7,
    EUR: parseFloat(formData.get("rateEUR") as string) || 8,
    GBP: parseFloat(formData.get("rateGBP") as string) || 9,
    JPY: parseFloat(formData.get("rateJPY") as string) || 0.04,
  };
}

export async function POST(request: NextRequest) {
  try {
    cleanup();

    const formData = await request.formData();
    const file = formData.get("file");
    const rates = parseRates(formData);

    if (!file || !(file instanceof File)) {
      return NextResponse.json(
        { error: "请上传一个 .xlsx 文件" },
        { status: 400 }
      );
    }

    if (!file.name.endsWith(".xlsx") && !file.name.endsWith(".xls")) {
      return NextResponse.json(
        { error: "只支持 .xlsx 格式的文件" },
        { status: 400 }
      );
    }

    const arrayBuffer = await file.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);

    console.log(
      `📖 Pacific convert: ${file.name} (${(buffer.length / 1024).toFixed(1)} KB), rates: USD=${rates.USD} EUR=${rates.EUR} GBP=${rates.GBP} JPY=${rates.JPY}`
    );

    const tmpPath = join(tmpdir(), `pacific_upload_${Date.now()}.xlsx`);
    writeFileSync(tmpPath, buffer);

    let result: PacificSplitResult;
    try {
      result = await convertAndSplitPacific(tmpPath, file.name, rates);
    } finally {
      try { unlinkSync(tmpPath); } catch {}
    }

    const sessionId = makeSessionId();
    store.set(sessionId, { result, createdAt: Date.now() });

    console.log(
      `✅ Pacific split: ${result.totalRows} rows → ${result.intervals.filter((i) => i.rowCount > 0).length} non-empty intervals`
    );

    // Return summary without buffer data
    return NextResponse.json({
      sessionId,
      sourceFile: result.sourceFile,
      totalRows: result.totalRows,
      skippedRows: result.skippedRows,
      intervals: result.intervals.map((iv) => ({
        name: iv.name,
        fileName: iv.fileName,
        rowCount: iv.rowCount,
      })),
      downloads: {
        allZip: `/api/convert-pacific?session=${sessionId}&file=all.zip`,
        files: result.intervals.map(
          (iv) =>
            `/api/convert-pacific?session=${sessionId}&file=${encodeURIComponent(iv.fileName)}`
        ),
      },
    });
  } catch (error) {
    console.error("Pacific convert error:", error);
    return NextResponse.json(
      {
        error:
          error instanceof Error
            ? error.message
            : "处理文件时出错，请确认上传的是易通货箱清单模板文件。",
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

  // ZIP download
  if (fileName === "all.zip") {
    const zipBuffer = await generatePacificZip(entry.result);
    const baseName = entry.result.sourceFile.replace(/\.(xlsx?|xls)$/i, "");
    const zipName = `${baseName} 太平洋拆分结果.zip`;
    // ASCII fallback must NOT contain non-ASCII characters (Chinese etc.)
    const asciiName = baseName.replace(/[^\x00-\x7F]/g, "_").replace(/\s+/g, "_") + "_pacific.zip";
    const encoded = encodeURIComponent(zipName);
    return new NextResponse(new Uint8Array(zipBuffer), {
      headers: {
        "Content-Type": "application/zip",
        "Content-Disposition": `attachment; filename="${asciiName}"; filename*=UTF-8''${encoded}`,
      },
    });
  }

  // Individual interval file
  const interval = entry.result.intervals.find((iv) => iv.fileName === fileName);
  if (!interval) {
    return NextResponse.json(
      { error: `文件 "${fileName}" 不存在` },
      { status: 404 }
    );
  }

  const encodedName = encodeURIComponent(fileName);
  // ASCII fallback for interval filenames (may contain Chinese)
  const asciiInterval = fileName.replace(/[^\x00-\x7F]/g, "_");
  return new NextResponse(new Uint8Array(interval.buffer), {
    headers: {
      "Content-Type":
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "Content-Disposition": `attachment; filename="${asciiInterval}"; filename*=UTF-8''${encodedName}`,
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
