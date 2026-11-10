/**
 * GET /api/sirui-to-supplier/options?supplier=xxx — 获取某供应商的报关方式选项 + 送货货站/交货仓库列表
 */

import { NextRequest, NextResponse } from "next/server";
import { getSupplierCustomsOptions, getSupplierStations, SiruiSupplierKey } from "@/lib/sirui-to-supplier";

const VALID_SUPPLIERS: SiruiSupplierKey[] = ["华威尔", "美琦", "英美", "航乐"];

export async function GET(request: NextRequest) {
  const { searchParams } = new URL(request.url);
  const supplier = searchParams.get("supplier") || "";

  if (!VALID_SUPPLIERS.includes(supplier as SiruiSupplierKey)) {
    return NextResponse.json({ error: "请选择有效的供应商" }, { status: 400 });
  }

  try {
    const key = supplier as SiruiSupplierKey;
    const customs = getSupplierCustomsOptions(key);
    const stations = await getSupplierStations(key);
    return NextResponse.json({ supplier, customs, stations });
  } catch (error) {
    console.error("Sirui options error:", error);
    return NextResponse.json({ error: error instanceof Error ? error.message : "读取选项失败" }, { status: 500 });
  }
}
