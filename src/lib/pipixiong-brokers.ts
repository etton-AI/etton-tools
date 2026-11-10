/**
 * 报关行（customs broker）配置存储
 *
 * 报关费账单底部需要打印「客户自行付给报关行」的收款信息，
 * 报关行由用户在上传时从下拉框选择，可手动新增并记忆。
 *
 * 持久化：data/brokers.json（与 history.json 同目录）。线上 K8s 无持久卷，
 * pod 重启后新增的报关行会丢失，故内置两家种子报关行兜底。
 */

import { readFileSync, writeFileSync, existsSync, mkdirSync } from "fs";
import { join } from "path";

export interface Broker {
  id: string;
  name: string;
  bankInfo: string;
}

const DATA_DIR = () => join(process.cwd(), "data");
const BROKER_FILE = () => join(DATA_DIR(), "brokers.json");

export const BROKER_SEED: Broker[] = [
  {
    id: "broker-yibaoda",
    name: "上海易豹达报关有限公司",
    bankInfo:
      "户名：上海易豹达报关有限公司\n纳税识别号：91310109MAEK0M0E5K \n开户行：浙江泰隆商业银行上海闸北支行0579-85551024\n账号：31010140201000070935\n接受号码：15957931517\n邮箱：2050609026@qq.com",
  },
  {
    id: "broker-shengtai",
    name: "上海圣塔艾洛斯国际物流有限公司",
    bankInfo:
      "单位名称:上海圣塔艾洛斯国际物流有限公司\n统一社会信用代码/纳税人识别号:91310114MAE6HG696Q\n注册地址:上海市嘉定区宝塔路138弄1号1513室\n注册电话:021-65417186\n开户银行名称-CNY:中国银行上海市大名支行\n银行账号-CNY:452088278484\n电子邮箱:stelogistics@163.com\n实际经营地址:上海市虹口区东大名路908号金岸大厦223室\n联系人:孙意平\n联系人电话:021-55235701",
  },
];

export function loadBrokers(): Broker[] {
  try {
    const p = BROKER_FILE();
    if (!existsSync(p)) return BROKER_SEED.map((b) => ({ ...b }));
    const raw = readFileSync(p, "utf-8");
    const parsed = JSON.parse(raw);
    if (Array.isArray(parsed) && parsed.length > 0) return parsed as Broker[];
    return BROKER_SEED.map((b) => ({ ...b }));
  } catch (e) {
    console.error("loadBrokers failed:", e);
    return BROKER_SEED.map((b) => ({ ...b }));
  }
}

export function saveBrokers(list: Broker[]): void {
  try {
    mkdirSync(DATA_DIR(), { recursive: true });
    writeFileSync(BROKER_FILE(), JSON.stringify(list, null, 2), "utf-8");
  } catch (e) {
    console.error("saveBrokers failed:", e);
  }
}

export function addBroker(name: string, bankInfo: string): Broker[] {
  const list = loadBrokers();
  const id =
    "broker-" + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
  list.push({ id, name: name.trim(), bankInfo: bankInfo.trim() });
  saveBrokers(list);
  return list;
}
