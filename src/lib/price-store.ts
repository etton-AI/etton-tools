/**
 * 价格数据共享模块
 * 查询 API 和上传 API 共用，支持缓存和刷新
 */

import fs from "fs";
import path from "path";

export interface PriceEntry {
  supplier: string;
  channel_name: string;
  vessel_config: string;
  delivery_method: string;
  destination_code: string;
  destination_type: string;
  destination_region: string;
  origin_region: string;
  origin_cities: string[];
  billing_type: string;
  min_quantity: string;
  min_quantity_value: number;
  unit_price: number;
  price_unit: string;
  transit_time_min: number | null;
  transit_time_max: number | null;
  transit_time_desc: string;
  claim_rule: string;
  effective_date: string;
  source_file: string;
}

interface PriceData {
  generated_at: string;
  total_records: number;
  stats: Record<string, number>;
  data: PriceEntry[];
}

let cache: PriceData | null = null;
const DATA_PATH = path.join(process.cwd(), "public", "data", "prices.json");

export function getData(): PriceData {
  if (cache) return cache;

  if (!fs.existsSync(DATA_PATH)) {
    console.warn("[price-store] 数据文件不存在，返回空");
    return { generated_at: "", total_records: 0, stats: {}, data: [] };
  }

  const raw = fs.readFileSync(DATA_PATH, "utf-8");
  cache = JSON.parse(raw) as PriceData;
  console.log(`[price-store] 数据已加载: ${cache.total_records} 条`);
  return cache;
}

/** 刷新缓存（上传新数据后调用） */
export function refreshCache(): void {
  cache = null;
  getData(); // 立即重新加载
}

/** 获取数据文件路径 */
export function getDataPath(): string {
  return DATA_PATH;
}
