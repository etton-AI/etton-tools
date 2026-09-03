/**
 * POST /api/warehouse-entry/export — 导出编辑后的建议结果为「出给客户」Excel
 * body: { rows: SuggestionRow[], supplierTotal: number }
 */

import { NextRequest, NextResponse } from "next/server";
import {
  exportOutputBuffer,
  accumulateHistory,
  type SuggestionRow,
} from "@/lib/warehouse-entry";

export async function POST(request: NextRequest) {
  try {
    const body = await request.json();
    const rows = (body?.rows ?? []) as SuggestionRow[];
    const supplierTotal = Number(body?.supplierTotal) || 0;

    if (!Array.isArray(rows) || rows.length === 0) {
      return NextResponse.json({ error: "没有可导出的数据" }, { status: 400 });
    }

    // 先导出（exportOutputBuffer 会执行全局兜底校验，必要时放大并标注 [全局调整]）
    const buffer = await exportOutputBuffer(rows, supplierTotal);

    // 导出后累积历史库（以出给客户数据为准，同款取计费重更大者）
    accumulateHistory(rows);

    const totalBoxes = rows.reduce((s, r) => s + (r.totalBoxes || 0), 0);
    const fileName = `内部三类数据_${new Date().toISOString().slice(0, 10)}_合计总箱数${totalBoxes}.xlsx`;
    const asciiName = fileName.replace(/[^\x00-\x7F]/g, "_");
    const encoded = encodeURIComponent(fileName);

    return new NextResponse(new Uint8Array(buffer), {
      headers: {
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": `attachment; filename="${asciiName}"; filename*=UTF-8''${encoded}`,
      },
    });
  } catch (error) {
    console.error("warehouse-entry export error:", error);
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
