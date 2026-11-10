/**
 * GET  /api/brokers — 获取已记录的报关行列表
 * POST /api/brokers — 新增报关行（name + bankInfo），返回更新后的列表
 */

import { NextRequest, NextResponse } from "next/server";
import { loadBrokers, addBroker } from "@/lib/pipixiong-brokers";

export async function GET() {
  try {
    const brokers = loadBrokers();
    return NextResponse.json({ brokers });
  } catch (error) {
    console.error("Brokers GET error:", error);
    return NextResponse.json({ error: "获取报关行列表失败" }, { status: 500 });
  }
}

export async function POST(request: NextRequest) {
  try {
    const body = await request.json();
    const name = (body?.name ?? "").toString().trim();
    const bankInfo = (body?.bankInfo ?? "").toString().trim();

    if (!name) {
      return NextResponse.json({ error: "请填写报关行名称" }, { status: 400 });
    }
    if (!bankInfo) {
      return NextResponse.json({ error: "请填写报关行收款信息" }, { status: 400 });
    }

    const brokers = addBroker(name, bankInfo);
    return NextResponse.json({ brokers });
  } catch (error) {
    console.error("Brokers POST error:", error);
    return NextResponse.json({ error: "保存报关行失败" }, { status: 500 });
  }
}
