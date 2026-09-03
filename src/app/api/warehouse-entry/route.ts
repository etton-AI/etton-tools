/**
 * POST /api/warehouse-entry — 上传客户数据 + 供应商数据，返回建议结果
 * GET  /api/warehouse-entry — 读取历史库（JSON）
 */

import { NextRequest, NextResponse } from "next/server";
import {
  parseCustomerFile,
  parseSupplierFile,
  buildSuggestions,
  loadHistory,
} from "@/lib/warehouse-entry";
import { tmpdir } from "os";
import { join } from "path";
import { writeFileSync, unlinkSync } from "fs";

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    const customerFile = formData.get("customerFile");
    const supplierFile = formData.get("supplierFile");

    if (
      !customerFile ||
      !supplierFile ||
      !(customerFile instanceof File) ||
      !(supplierFile instanceof File)
    ) {
      return NextResponse.json(
        { error: "请同时上传客户数据和供应商数据两个文件" },
        { status: 400 }
      );
    }

    for (const f of [customerFile, supplierFile]) {
      if (!f.name.endsWith(".xlsx") && !f.name.endsWith(".xls")) {
        return NextResponse.json(
          { error: `文件 "${f.name}" 格式不支持，请上传 .xlsx 或 .xls 文件` },
          { status: 400 }
        );
      }
    }

    console.log(
      `📦 TR入仓: 客户=${customerFile.name} (${(customerFile.size / 1024).toFixed(1)} KB), 供应商=${supplierFile.name} (${(supplierFile.size / 1024).toFixed(1)} KB)`
    );

    const tmpDir = tmpdir();
    const suffix = `${Date.now()}_${Math.random().toString(36).slice(2, 6)}`;
    const customerPath = join(tmpDir, `cust_${suffix}.xlsx`);
    const supplierPath = join(tmpDir, `supp_${suffix}.xlsx`);
    writeFileSync(customerPath, Buffer.from(await customerFile.arrayBuffer()));
    writeFileSync(supplierPath, Buffer.from(await supplierFile.arrayBuffer()));

    try {
      const customers = await parseCustomerFile(customerPath);
      const boxes = await parseSupplierFile(supplierPath);
      const history = loadHistory();
      const { rows, supplierTotal } = buildSuggestions(customers, boxes, history);

      const matchedCount = rows.filter(
        (r) => !r.alarms.includes("⚠需人工复核")
      ).length;

      return NextResponse.json({
        rows,
        supplierTotal,
        summary: {
          customerCount: customers.length,
          boxCount: boxes.length,
          matchedCount,
          unmatchedCount: rows.length - matchedCount,
          historyCount: history.length,
        },
      });
    } finally {
      try { unlinkSync(customerPath); } catch {}
      try { unlinkSync(supplierPath); } catch {}
    }
  } catch (error) {
    console.error("warehouse-entry error:", error);
    return NextResponse.json(
      {
        error:
          error instanceof Error
            ? error.message
            : "处理失败，请确认上传的是正确的文件。",
      },
      { status: 500 }
    );
  }
}

export async function GET() {
  try {
    const history = loadHistory();
    return NextResponse.json({ history, count: history.length });
  } catch (error) {
    console.error("warehouse-entry history error:", error);
    return NextResponse.json({ error: "读取历史库失败" }, { status: 500 });
  }
}

export const config = {
  api: {
    bodyParser: {
      sizeLimit: "50mb",
    },
  },
};
