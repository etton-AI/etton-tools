# 皮皮熊账单拆分 — 功能规格文档

> 最后更新: 2026-11-09 | 维护者: berry-bi

## 1. 功能概述

**皮皮熊账单拆分**（`/pipixiong-split`）把皮皮熊合并账单 Excel（4 个数据 sheet）中**未支付**记录，按「货代对应识别号」拆分为四类账单：

| 账单 | 抬头 | 币种 | 内容 |
|------|------|------|------|
| 国内账单 | 深圳易通 | RMB | 拖车费 + 港杂费 |
| 报关费账单 | 深圳易通 | RMB | 报关费（客户自行付报关行） |
| 国外账单 | 香港易通 | USD | 海运费 + 税金 / 各类 USD 费用 |
| INVOICE | 香港易通 | USD | 客户发票（一行明细 + Total） |

- **输入**: 皮皮熊合并账单 Excel（`.xlsx` / `.xls`），必含 4 个数据 sheet
- **输出**: 4 个独立 Excel（国内账单 / 报关费 / 国外账单 / INVOICE），分别下载，不打 ZIP

## 2. 数据来源（4 个数据 sheet）

源文件除前 3 个汇总 sheet 外，有 4 个数据 sheet：

| sheet | 费用类型 | 账单去向 |
|-------|---------|---------|
| `1.大货出货数据费用统计` | 海运费/拖车费/报关费/港杂费（RMB + USD 混合） | 国内 / 报关 / 国外 / INVOICE |
| `2.港皮海外仓转仓费用明细` | USD（香港公司支付） | 国外 / INVOICE |
| `3.大货以及其他货物快递退运费用` | USD（香港公司支付） | 国外 / INVOICE |
| `4.空派快递费用明细` | USD（香港公司支付） | 国外 / INVOICE |

### 2.1 sheet 1（大货出货数据费用统计）

列：1 序列、2 时间、3 平台、4 国家、5 入库地址（代码）、9 发货数量、10 货代对应识别号、11 报关单号、12 总费用RMB、13 拖车费RMB、14 报关费RMB、15 港杂费RMB、16 海运+税金USD、19 是否支付。

- `时间`/`平台`/`国家`/`数量`/`入库地址`/`报关单号`/`是否支付` 读作字符串；金额列读作数字
- 是否支付 `isPaid` 判定：`y` / `yes` / `是` / `已支付`（大小写不敏感）

### 2.2 sheet 2/3/4（USD 明细）

列：1 序列、2 时间、4 发货数量、5 货代对应识别号、6 费用、7（或 8）费用详情、币种列（sheet2 无币种列固定 USD；sheet3/4 为第 7 列）、是否支付列（sheet2 为第 9 列、sheet3/4 为第 11 列）。

- 货代对应识别号可能是**多行**（如 `ZMGS26080001\n875399167912`），只取首行（`firstLine`）
- `时间` 归一化「年.月」：`2026.08` → `2026.8`
- `发货数量` 读作数字（int）
- `是否支付` 已支付（`Y`）的行跳过

## 3. 输出变换

### 3.1 拆分规则

1. 只处理**未支付**记录；已支付（`Y`）用于 JOB NO 去重统计，但不生成账单
2. sheet 1 每行按费用判断生成哪些账单：
   - `hasDomestic` = 拖车费 > 0 或 港杂费 > 0
   - `hasCustoms` = 报关费 > 0
   - `hasInternational` = 海运+税金USD > 0
   - `hasInvoice` = 海运+税金USD ≠ 0（< 0 时 INVOICE 项目标签为 `COMPENSATION COSTS` 补偿款）
3. sheet 2/3/4 每行（未支付、费用 > 0）生成 国外账单 + INVOICE

### 3.2 JOB NO 去重

- 统计**全部记录**（含已支付）中每个货代识别号出现次数
- 出现 ≥ 2 次的，JOB NO 加 `-N` 后缀（按全表顺序第 N 次出现），如 `ZMGS26060001-2`

### 3.3 报关行（报关费账单抬头）

- 报关费账单 R8 显示报关行收款信息（`bankInfo` 多行）
- 报关行**不按关区代码映射**，由用户在**上传时下拉选择**已记录的报关行，或手动新增并保存（持久化到 `data/brokers.json`）
- 内置两家：上海易豹达报关有限公司、上海圣塔艾洛斯国际物流有限公司

## 4. 技术实现

- **ExcelJS** 读源 / 写输出（4 个 Workbook 分别生成）
- **格式与参考文件精确对齐**：字体统一 `charset: 134`；`applyChineseDefaultFont()` 用 JSZip 替换 `xl/styles.xml` 默认字体（Calibri → 宋体 11 charset 134）；列宽用 4 位小数精确值；`pageSetup()` 工厂省略 `fitToWidth/scale/DPI` 并设 `paperSize: 1`。已验证与参考文件在内容值/字体/边框/填充/列宽/行高/页面设置/默认字体上**零差异**。
- **被合并单元格覆盖判定**（`isCoveredMergedCell`）：源文件含纵向合并单元格（如 `P73:P74`），ExcelJS 对覆盖单元格会返回 master 的值；读取前须判断 `cell.isMerged && cell.master.address !== cell.address`，覆盖单元格按空处理
- **报关行持久化**：`src/lib/pipixiong-brokers.ts` 文件存储（`data/brokers.json` + seed 兜底，沿用 history.json 模式）
- **结果内存存储**：4 个 Buffer 存进程内 Map，30 分钟过期

## 5. API

- `POST /api/pipixiong-split`：formData `file`（账单）+ `brokerName` / `brokerBankInfo`（可选，报关行）→ `{ sessionId, groups[], downloadUrls{domestic, customs, international, invoice} }`
  - `groups[]`：`{ seq, forwarderId, customsNo, country, qty, truckFee, customsFee, portFee, oceanTaxUSD, hasDomestic, hasCustoms, hasInternational, hasInvoice }`
- `GET /api/pipixiong-split?session=xxx&type=domestic|customs|international|invoice`：下载对应账单（30 分钟 session 过期）
- `GET /api/brokers` → `{ brokers[] }`（`{ id, name, bankInfo }`）
- `POST /api/brokers`：`{ name, bankInfo }` → 新增并返回 `{ brokers[] }`

## 6. 非目标 / 已知限制

- 报关行**不自动**按关区代码 / 报关单号匹配，需用户上传时手动选择或新增
- 结果仅内存保存，进程重启丢失 session；报关行持久化文件 `data/brokers.json` 在 Sealos 无持久卷时 pod 重启会丢（回落到内置两家 seed）
- INVOICE 无船务信息（无 POL / CNTR NO 等 shipRows）
