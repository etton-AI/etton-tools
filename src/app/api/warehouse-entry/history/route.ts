/**
 * POST /api/warehouse-entry/history — 导入历史 Excel（首次初始化 / 备份恢复）
 * GET  /api/warehouse-entry/history — 导出历史库为 Excel（备份）
 */

import { NextRequest, NextResponse } from "next/server";
import {
  loadHistory,
  saveHistory,
  importHistoryFromExcel,
  exportHistoryBuffer,
  upsertHistoryEntry,
} from "@/lib/warehouse-entry";
import { tmpdir } from "os";
import { join } from "path";
import { writeFileSync, unlinkSync } from "fs";

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    const historyFile = formData.get("historyFile");

    if (!historyFile || !(historyFile instanceof File)) {
      return NextResponse.json(
        { error: "请上传历史 Excel 文件" },
        { status: 400 }
      );
    }

    if (!historyFile.name.endsWith(".xlsx") && !historyFile.name.endsWith(".xls")) {
      return NextResponse.json(
        { error: `文件 "${historyFile.name}" 格式不支持，请上传 .xlsx 或 .xls 文件` },
        { status: 400 }
      );
    }

    const tmpPath = join(
      tmpdir(),
      `history_${Date.now()}_${Math.random().toString(36).slice(2, 6)}.xlsx`
    );
    writeFileSync(tmpPath, Buffer.from(await historyFile.arrayBuffer()));

    try {
      const imported = await importHistoryFromExcel(tmpPath);
      // 与现有历史库合并：同款（品名 + 客户箱规/实重相近）直接覆盖。
      // 导入的是「最终出给客户」数据，以我们最终提供的值为准（覆盖自动累积的建议值）。
      const existing = loadHistory();
      for (const entry of imported) {
        upsertHistoryEntry(existing, entry, { overwrite: true });
      }
      saveHistory(existing);

      return NextResponse.json({
        importedCount: imported.length,
        totalCount: existing.length,
      });
    } finally {
      try { unlinkSync(tmpPath); } catch {}
    }
  } catch (error) {
    console.error("warehouse-entry history import error:", error);
    return NextResponse.json(
      {
        error:
          error instanceof Error
            ? error.message
            : "历史库导入失败，请确认文件格式正确。",
      },
      { status: 500 }
    );
  }
}

export async function GET() {
  try {
    const history = loadHistory();
    const buffer = await exportHistoryBuffer(history);

    const fileName = `历史库_${new Date().toISOString().slice(0, 10)}.xlsx`;
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
    console.error("warehouse-entry history export error:", error);
    return NextResponse.json({ error: "历史库导出失败" }, { status: 500 });
  }
}

export const config = {
  api: {
    bodyParser: {
      sizeLimit: "50mb",
    },
  },
};
