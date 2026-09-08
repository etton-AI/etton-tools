/**
 * POST /api/warehouse-entry-full/export — 导出建议结果为「出给客户」Excel（逐箱版）
 * body: { rows: SuggestionRow[], supplierTotal: number }
 */

import { NextRequest, NextResponse } from "next/server";
import {
  exportOutputBuffer,
  type SuggestionRow,
} from "@/lib/warehouse-entry-full";

export async function POST(request: NextRequest) {
  try {
    const body = await request.json();
    const rows = (body?.rows ?? []) as SuggestionRow[];
    const supplierTotal = Number(body?.supplierTotal) || 0;

    if (!Array.isArray(rows) || rows.length === 0) {
      return NextResponse.json({ error: "没有可导出的数据" }, { status: 400 });
    }

    // 先导出（exportOutputBuffer 会执行全局兜底校验，必要时放大并标注 [全局调整]）
    const { buffer, warning } = await exportOutputBuffer(rows, supplierTotal);

    const totalBoxes = rows.reduce((s, r) => s + (r.totalBoxes || 0), 0);
    const fileName = `内部三类数据_${new Date().toISOString().slice(0, 10)}_合计总箱数${totalBoxes}.xlsx`;
    const asciiName = fileName.replace(/[^\x00-\x7F]/g, "_");
    const encoded = encodeURIComponent(fileName);

    const response = new NextResponse(new Uint8Array(buffer), {
      headers: {
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": `attachment; filename="${asciiName}"; filename*=UTF-8''${encoded}`,
      },
    });
    if (warning) {
      response.headers.set("X-Warning", encodeURIComponent(warning));
    }
    return response;
  } catch (error) {
    console.error("warehouse-entry-full export error:", error);
    return NextResponse.json(
      {
        error:
          error instanceof Error ? error.message : "导出失败，请重试。",
      },
      { status: 500 }
    );
  }
}

export const config = {
  api: {
    bodyParser: {
      sizeLimit: "50mb",
    },
  },
};
