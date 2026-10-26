/**
 * GET /api/tr-to-supplier/options?supplier=xxx — 获取某供应商的报关方式选项 + 送货货站/交货仓库列表
 */

import { NextRequest, NextResponse } from "next/server";
import { getSupplierCustomsOptions, getSupplierStations, SupplierKey } from "@/lib/tr-to-supplier";

const VALID_SUPPLIERS: SupplierKey[] = ["凯鑫", "英美", "星链", "航乐"];

export async function GET(request: NextRequest) {
  const { searchParams } = new URL(request.url);
  const supplier = searchParams.get("supplier") || "";

  if (!VALID_SUPPLIERS.includes(supplier as SupplierKey)) {
    return NextResponse.json({ error: "请选择有效的供应商" }, { status: 400 });
  }

  try {
    const key = supplier as SupplierKey;
    const customs = getSupplierCustomsOptions(key);
    const stations = await getSupplierStations(key);
    return NextResponse.json({ supplier, customs, stations });
  } catch (error) {
    console.error("TR options error:", error);
    return NextResponse.json({ error: error instanceof Error ? error.message : "读取选项失败" }, { status: 500 });
  }
}
