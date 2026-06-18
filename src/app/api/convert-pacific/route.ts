/**
 * POST /api/convert-pacific — 上传货箱清单 Excel，返回太平洋投保清单
 */

import { NextRequest, NextResponse } from "next/server";
import { convertPacificInsurance } from "@/lib/convert-pacific-insurance";
import { tmpdir } from "os";
import { join } from "path";
import { writeFileSync, unlinkSync } from "fs";

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    const file = formData.get("file");

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

    console.log(`📖 Pacific convert: ${file.name} (${(buffer.length / 1024).toFixed(1)} KB)`);

    // Write to temp file — exceljs load(buffer) can fail in Node.js runtime
    const tmpPath = join(tmpdir(), `pacific_upload_${Date.now()}.xlsx`);
    writeFileSync(tmpPath, buffer);

    let result;
    try {
      result = await convertPacificInsurance(tmpPath, file.name);
    } finally {
      try { unlinkSync(tmpPath); } catch {}
    }

    console.log(`✅ Pacific convert: ${result.rowCount} rows mapped`);

    // Return the file directly for download
    const outFileName = file.name
      .replace(/\.xlsx?$/i, "")
      .replace(/[\\/:*?"<>|]/g, "_");

    return new NextResponse(new Uint8Array(result.buffer), {
      headers: {
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": `attachment; filename*=UTF-8''${encodeURIComponent(outFileName + "_太平洋投保清单.xlsx")}`,
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
