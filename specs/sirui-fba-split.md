# 思锐FBA下单发票拆分 — 功能规格文档

> 最后更新: 2026-11-09 | 维护者: berry-bi

## 1. 功能概述

**思锐FBA下单发票拆分**（`/sirui-fba-split`）把客户提供的一张「下单发票」（含多个 Amazon FBA 仓库的全部明细）按内部报价表里 FBA 码对应的「仓点 / 渠道」，拆分为多张发票（每个 FBA 仓一张）。

- **输入**:
  - 下单发票（`.xlsx`，一张发票含所有 FBA 仓明细）
  - 内部报价表（`.xlsx`，`内部` sheet 记录 FBA 码 → 仓点 / 渠道）
- **输出**: ZIP，每张拆分发票文件名 = `{FBA}-{仓点}.xlsx`

## 2. 数据来源

### 2.1 内部报价表（`内部` sheet）

| 列 | 含义 |
|----|------|
| D | FBA 码（`/^FBA[0-9A-Z]{6,}$/i`）|
| C | 仓点（如 `FAT2`）|
| H | 渠道（如 `OA-拆送（OA联盟）`）|

### 2.2 下单发票

- F9：收件人国家（`USA` / `US` → `美国`）
- R25 起：明细行，A 列货箱编号、K 列申报币种（`美元` → `USD`）
- FBA 码从货箱编号用前缀匹配

## 3. 输出变换

每张拆分发票（以参考文件「拆分的发票」反推）：

1. 只保留该 FBA 码对应的明细行，行号从 R25 重新连续编号
2. F9 收件人国家：`USA`/`US` → `美国`
3. F10 亚马逊仓库代码：填「仓点」
4. B21 备注：填「渠道」
5. K 列申报币种：`美元` → `USD`（其余按 `CURRENCY_TO_ISO` 映射）
6. 每行 W 列填 `{FBA}-{仓点}`（无表头）
7. 删除「服务渠道」「FBA地址」两个 sheet，输出只保留 1 个 sheet
8. 保留 WPS DISPIMG 产品图片

## 4. 技术实现

- **ExcelJS** 读取报价表与发票数据（只读，不写回）
- **JSZip** 对源 xlsx 做字符串级手术，避免 ExcelJS 写回时丢失 DISPIMG 图片
- 新增单元格一律写**共享字符串**（`t="s"`）而非内联字符串（`inlineStr`）——源发票中文都在 sharedStrings，易通系统已验证能读；`inlineStr` 是易通解析器不支持的类型

## 5. 产品图片（WPS DISPIMG）处理

- 图片存于 `xl/cellimages.xml`（`<xdr:pic>` + `<a:off x/y>` 绝对 EMU 锚点）、媒体在 `xl/media/`、公式 `_xlfn.DISPIMG("ID_xxx",1)` 在 V 列
- 源 57 个明细行只有 50 张唯一图片（同一产品多箱/跨 FBA 组共用同一张图）
- 拆分时每组按图片 ID **去重只保留首次出现**的那条 `<xdr:pic>`；否则会产出同 ID 多条，图片数多于参考（LRKT9/HIA1 应为 5 张，错产出 8 张）
- 图片 y 偏移按整行重映射 `newY = offY + (targetNewRow - anchorRow) * 558800`（数据行统一 44pt = 558800 EMU），保证行重编号后图片仍对位

## 6. API

- `POST /api/sirui-fba-split`：formData `quote`（内部报价表）+ `file`（下单发票）→ `{ sessionId, result }`
  - `result` 含 `{ zipFileName, total, files[], warnings[], downloadUrl }`
  - `files[]`：`{ fba, warehouse, fileName, rowCount }`
- `GET /api/sirui-fba-split?session=xxx`：下载 ZIP（30 分钟 session 过期）
- `bodyParser.sizeLimit: "50mb"`

## 7. 非目标 / 已知限制

- **图片 y 偏移与参考文件的有意差异**：参考文件拆出后图片仍「浮」在原绝对位置（行重编号后图片不对位），本实现按整行偏移重映射 y 使图片对位到新行。图片「计数」与参考完全一致，仅 y 偏移更正确。
- **跨组借图**：同一张 DISPIMG 图被多个 FBA 组引用时（锚点行在别的组），本实现把它锚到本组首次引用该图的行，位置略偏但计数正确。
- **结果内存存储**：拆分结果仅在内存保存 30 分钟，不落盘。

## 8. 与参考文件的样式对齐

经 openpyxl「展开后视觉样式」全量对比（值/字体/填充/边框/对齐/数字格式/行高/列宽/合并/图片计数）验证 15 份输出零差异。关键对齐点：

- 新增单元格统一写**共享字符串**（`t="s"`）并**保留原 `s` 样式索引**，避免 `setCellSharedStr` 非贪婪正则吞样式/吞相邻单元格
- 跨组 DISPIMG 共享公式用 `expandSharedFormulas` 展开为独立 `<f>`，避免 master 被丢弃导致 V 列变 `=`
- K 列币种用 `dominantStyle` 统一为无填充主样式（源文件 s=102/104 混用）
- W 列宽设为 `25.2090909090909` + `customWidth="1"`（与参考一致）
