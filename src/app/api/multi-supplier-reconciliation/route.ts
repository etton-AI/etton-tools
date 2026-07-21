/**
 * POST /api/multi-supplier-reconciliation — 上传供应商账单 + 请款明细，返回多供应商对账结果
 * GET  /api/multi-supplier-reconciliation?session=xxx — 下载对账结果 Excel
 */

import { NextRequest, NextResponse } from "next/server";
import {
  processMultiSupplierReconciliation,
  type MultiReconResult,
} from "@/lib/multi-supplier-reconciliation";
import { tmpdir } from "os";
import { join } from "path";
import { writeFileSync, unlinkSync } from "fs";

// In-memory session store
const store = new Map<string, { result: MultiReconResult; createdAt: number }>();

function cleanup() {
  const cutoff = Date.now() - 30 * 60 * 1000;
  for (const [k, v] of store) {
    if (v.createdAt < cutoff) store.delete(k);
  }
}

function makeSessionId(): string {
  return "msr_" + Math.random().toString(36).slice(2, 10) + Date.now().toString(36);
}

export async function POST(request: NextRequest) {
  try {
    cleanup();

    const formData = await request.formData();
    const billFile = formData.get("billFile");
    const paymentFile = formData.get("paymentFile");
    const supplier = formData.get("supplier") as string | null;

    if (!billFile || !paymentFile || !(billFile instanceof File) || !(paymentFile instanceof File)) {
      return NextResponse.json(
        { error: "请同时上传供应商账单和请款明细两个文件" },
        { status: 400 }
      );
    }

    // 格式校验
    for (const f of [billFile, paymentFile]) {
      if (!f.name.endsWith(".xlsx") && !f.name.endsWith(".xls")) {
        return NextResponse.json(
          { error: `文件 "${f.name}" 格式不支持，请上传 .xlsx 或 .xls 文件` },
          { status: 400 }
        );
      }
    }

    console.log(
      `📖 Multi-Supplier Recon: 账单=${billFile.name} (${(billFile.size / 1024).toFixed(1)} KB), 请款=${paymentFile.name} (${(paymentFile.size / 1024).toFixed(1)} KB), 供应商=${supplier || "自动识别"}`
    );

    // 写入临时文件
    const tmpDir = tmpdir();
    const billPath = join(tmpDir, `bill_${Date.now()}.xlsx`);
    const paymentPath = join(tmpDir, `payment_${Date.now()}.xlsx`);

    const billBuffer = Buffer.from(await billFile.arrayBuffer());
    const paymentBuffer = Buffer.from(await paymentFile.arrayBuffer());

    writeFileSync(billPath, billBuffer);
    writeFileSync(paymentPath, paymentBuffer);

    let result: MultiReconResult;
    try {
      result = await processMultiSupplierReconciliation(
        billPath,
        paymentPath,
        billFile.name,
        paymentFile.name,
        supplier || undefined
      );
    } finally {
      try { unlinkSync(billPath); } catch {}
      try { unlinkSync(paymentPath); } catch {}
    }

    const sessionId = makeSessionId();
    store.set(sessionId, { result, createdAt: Date.now() });

    // 返回结果（不含 buffer）
    return NextResponse.json({
      sessionId,
      supplier: result.supplier,
      sourceFiles: result.sourceFiles,
      summary: result.summary,
      // 仅返回有问题的行给前端展示
      issueRows: result.rows
        .filter((r) => r.status !== "一致")
        .map((r) => ({
          soNumber: r.soNumber,
          billAmount: r.billAmount,
          paymentAmount: r.paymentAmount,
          difference: r.difference,
          diffRate: r.diffRate,
          status: r.status,
        })),
      totalRows: result.rows.length,
      downloadUrl: `/api/multi-supplier-reconciliation?session=${sessionId}`,
    });
  } catch (error) {
    console.error("Multi-supplier reconciliation error:", error);
    return NextResponse.json(
      {
        error:
          error instanceof Error
            ? error.message
            : "对账处理失败，请确认上传的是正确的账单文件。",
      },
      { status: 500 }
    );
  }
}

export async function GET(request: NextRequest) {
  cleanup();

  const { searchParams } = new URL(request.url);
  const sessionId = searchParams.get("session");

  if (!sessionId) {
    return NextResponse.json({ error: "Missing session parameter" }, { status: 400 });
  }

  const entry = store.get(sessionId);
  if (!entry) {
    return NextResponse.json(
      { error: "会话已过期，请重新上传文件进行对账" },
      { status: 404 }
    );
  }

  const { result } = entry;
  const safeSupplier = result.supplier.replace(/[\/\\]/g, "_");
  const baseName = `多供应商对账_${safeSupplier}_${result.sourceFiles.bill.replace(/\.(xlsx?|xls)$/i, "")}`;
  const fileName = `${baseName}.xlsx`;
  const asciiName = baseName.replace(/[^\x00-\x7F]/g, "_") + ".xlsx";
  const encoded = encodeURIComponent(fileName);

  return new NextResponse(new Uint8Array(result.buffer), {
    headers: {
      "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "Content-Disposition": `attachment; filename="${asciiName}"; filename*=UTF-8''${encoded}`,
    },
  });
}

export const config = {
  api: {
    bodyParser: { sizeLimit: "50mb" },
  },
};
