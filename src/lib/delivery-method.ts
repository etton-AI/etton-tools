/**
 * 尾程派送方式识别(纯函数, 无 Node 依赖, 客户端/服务端均可 import)。
 * 用于「渠道选错」提醒: 识别源发票(文件名/服务渠道 B7)与所选渠道的派送方式并比对。
 */

/** 尾程派送方式分类 */
export type DeliveryMethod = "海派" | "卡派" | "快递派";

/**
 * 从任意文本(文件名 / 服务渠道 B7 / 渠道名)识别尾程派送方式。
 * 优先级: 快递/UPS 类 > 海派 > 卡派(含海卡)。识别不到返回 null。
 * - 海派   = 海运 + 快递尾程
 * - 卡派   = 卡车派送(海卡=海运+卡车, 归入卡派)
 * - 快递派 = UPS / DPD / 红单 / 快递 等快递尾程
 */
export function detectDeliveryMethod(...texts: (string | null | undefined)[]): DeliveryMethod | null {
  const hay = texts
    .filter((t): t is string => !!t && String(t).trim() !== "")
    .join(" ");
  if (!hay) return null;
  if (/UPS|DPD|红单|快递/i.test(hay)) return "快递派";
  if (/海派/.test(hay)) return "海派";
  if (/卡派|海卡|卡车/.test(hay)) return "卡派";
  return null;
}
