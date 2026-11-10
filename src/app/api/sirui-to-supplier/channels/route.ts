/**
 * GET /api/sirui-to-supplier/channels?supplier=xxx — 获取某供应商的渠道下拉列表(从模板读取)
 */

import { NextRequest, NextResponse } from "next/server";
import { getSupplierChannels, SiruiSupplierKey } from "@/lib/sirui-to-supplier";

const VALID_SUPPLIERS: SiruiSupplierKey[] = ["华威尔", "美琦", "英美", "航乐"];

export async function GET(request: NextRequest) {
  const { searchParams } = new URL(request.url);
  const supplier = searchParams.get("supplier") || "";

  if (!VALID_SUPPLIERS.includes(supplier as SiruiSupplierKey)) {
    return NextResponse.json({ error: "请选择有效的供应商" }, { status: 400 });
  }

  try {
    const channels = await getSupplierChannels(supplier as SiruiSupplierKey);
    return NextResponse.json({ supplier, channels });
  } catch (error) {
    console.error("Sirui channels error:", error);
    return NextResponse.json({ error: error instanceof Error ? error.message : "读取渠道失败" }, { status: 500 });
  }
}
