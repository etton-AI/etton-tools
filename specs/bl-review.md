# 提单 + 电放保函自动生成 — 功能规格

## 背景

用户日常根据「报关底单」手工填写提单（Bill of Lading）与电放保函（Telex Release Application）。参考实现（`ETTON_审核流程` 目录，Python + pdfplumber + docx-mailmerge）已验证可用，本功能把它落地为 ETTON 平台正式功能：上传报关底单 PDF → 自动提取字段 → 人工审核补填 → 一键生成提单 + 电放保函 → 打包 ZIP。

## 流程

```
文件夹(每子文件夹=一票) ──①批量提取──> 字段 ──②审核表(标红:空白/中文)──> 人工确认 ──③生成──> 提单.pdf + 电放保函.docx + 底单.pdf → ZIP
    箱货清单xlsx(同文件夹,可选) ──①补充──> 英文品名 + 总体积 + 国家/箱数(校验)
    周汇总箱货清单xlsx(整周一份,可选) ──①按工作号分组+匹配──> 每票品名/体积/渠道（匹配不到回退单票清单）
    物流追踪表xlsx(单独上传,可选) ──①匹配──> FBA ID→ETD/ETA/船名航次 → 起运日期(ETD) + 申请日期(ETA-2天) + 铁路船名航次
    拆分底单(报/放/委托) ──①合并──> 一份报关底单
```

- **前端**：`/bl-review`（Next.js 页面），批量文件夹上传（`webkitdirectory`）+ 物流追踪表单独上传；审核表空白/中文标红；生成后页面内预览提单/保函 + ZIP 下载。
- **后端**：`bl-service/`（Flask，原生 Python 跑 + 原生 LibreOffice 转 PDF），接口 `extract` / `extract-batch` / `generate` / `file`。

## 字段映射

| 报关单位置 | 提单字段 | 保函字段 |
|-----------|---------|---------|
| 境内发货人 | shipper（**拓锐固定英文常量**） | 申请单位（固定英文） |
| 境外收货人 | consignee（**拓锐固定英文常量**） | 收货人(保函)（固定英文） |
| 提运单号 | 提单号 | B/L NO |
| 运输工具名称及航次号 | 船名航次 | 运输工具 |
| 集装箱号 | 柜号 | 柜号 |
| 指运港 | 目的港 | 目的地 |
| 离境口岸 | 起运地 | — |
| 件数 | 箱数 | — |
| 毛重 | 总重量 | — |
| 申报日期 | 起运日期 | — |
| 商品明细 | 品名 | — |
| **箱货清单「系统SO」列** | **系统SO（审核表展示，供查找相关信息）** | — |
| **箱货清单「英文品名」列** | **品名（去版本号 + 去重后覆盖）** | — |
| **箱货清单「体积」列求和** | **总体积(CBM)（覆盖）** | — |
| **物流追踪表「ETD 开船日」（按 FBA ID 匹配）** | **起运日期（覆盖）** | — |
| **物流追踪表「ETA 到港日」（按 FBA ID 匹配）** | — | **申请日期（=到港时间：ETA−2天；ETA−ETD<2天取 ETA）** |
| **物流追踪表「船名航次/班列号」（按 FBA ID 匹配）** | **船名航次（底单为空时补入，铁路班列号）** | — |
| **箱货清单「客户渠道」** | **目的港（按 channel_map 覆盖）** | **目的地（同步）** |

> 起运日期落提单时转 `DD MMM YYYY`（如 `2026-05-14` → `14 MAY 2026`，带空格）；品名落提单时保留空格、全大写（如 `BLUE OCEAN DREAM GALAXY PROJECTOR`），并去掉末尾版本号（`3.0`/`2.0`）。

## 固定客户信息（拓锐）

- `SHIPPER_BL`（3 行英文）、`CONSIGNEE_BL`、`SHIPPER_TELEX`、`CONSIGNEE_TELEX` 硬编码在 `core.py` 顶部，覆盖底单提取的中文名。来源=正确提单/保函。

## 箱货清单汇总（品名英文 + 总体积）

- `parse_packing_list(xlsx)` 定位「系统SO/英文品名/中文品名/体积/货箱重量/FBA ID/客户渠道」表头（跳过表头行），系统SO 去重供审核表展示（方便查找），英文品名去版本号 + 去重（全大写）覆盖「品名」，体积列求和覆盖「总体积」，FBA ID 去重供物流追踪表匹配，客户渠道取首个非空值供目的港映射。
- 箱货清单**无「体积」列**时（如海运单 TRKJ26050005），回退用「长(CM)×宽(CM)×高(CM)×总箱数(CTN) / 1e6」算体积。
- 未上传箱货清单时，底单中文品名经 `PRODUCT_EN_MAP` 词典翻译兜底（同样去版本号 + 全大写）。

## 港口英文（起运港按渠道大类 + 运输方式分流；目的港按渠道大类 + 运抵国）

- **起运港分流**（`apply_port_map`，先按渠道大类、再按运输方式）：
  - 快递渠道（客户渠道含 `联邦`/`空派`，如「香港联邦IP」）→ 固定 `SHENZHEN`。
  - 海运（`水路运输`）→ 取「离境口岸」中文 → 查 `port_map.origin`（如 `南沙港一期码头`→`NANSHA`、`外高桥`→`SHANGHAI`、`盐田`→`YANTIAN`）。
  - 非海运（`铁路运输`/`公路运输`/`航空运输`）→ 取「海关编号」后面的**关区名** → 查 `customs_office_map`（如 `增城海关`→`GUANGZHOU`、`蓉青关`→`CHENGDU`、`渝州海关`→`CHONGQING`、`车站海关`→`XIAN`）。关区名缺失时保留离境口岸中文供人工填。
- **目的港分流**（按「运抵国」查，规则见 `_ROUTE_DEST_RULES`）：
  - 铁路渠道（客户渠道含 `铁路`/`快铁`/`铁派`）→ 英国→`MALASZEWICZE`、德国→`DUISBURG`。
  - 卡航渠道（客户渠道含 `卡航`）→ 英国→`THE UK`、德国→`GERMANY`。
  - 其他 → 查 `port_map.destination`（运抵国→英文）。
- `port_map.json`：**结构化** `{ origin: 离境口岸→英文, destination: 运抵国→英文 }`。未命中保留中文供人工填。
- `customs_office_map.json`：非海运票的**关区名 → 起运港英文**对照（报关关区 ≈ 发货城市，比边境离境口岸更贴近实际起运地）。
- `channel_map.json`：非规则渠道（如海运「美转加」系列）的**渠道名 → 目的港英文**精确映射，`apply_channel_map` 兜底；铁路/卡航/快递三类已规则化，不再走 channel_map。
- **自动记忆**：generate 阶段 `remember_ports`（海运起运地→origin、非海运关区名→customs_office_map、目的港→destination，快递起运地固定不记忆）+ `remember_channels`（非规则渠道→港口），把人工修正后的新映射回写 json。
- **在线编辑入口**：`/bl-mapping` 页面可增删改四类映射（起运港 / 目的港 / 渠道 / 关区名），经 `GET/POST /api/bl/mappings` 读写三个 json，保存后下次提取即生效；`/bl-review` 页面顶部有「🗺️ 港口 / 渠道映射设置」按钮跳转到 `/bl-mapping`。

## 接口契约

- `GET /api/bl/customers` —— 返回 `{ ok, customers:[{key,label}], default }`，前端下拉据此渲染（加客户只改后端 `CUSTOMERS` 一处）。
- 所有接口均接受 `customer`（客户标识，默认 `拓锐`），决定固定 shipper/consignee + 保函页眉/TO/FROM 等；未知客户回退拓锐。
- `POST /api/bl/extract`（`file=` PDF，可选 `packing=` xlsx、`tracking=` xlsx、`customer=`）—— 单票
  - 响应：`{ ok, record, bl_fields, bl_header, telex_fields, warning }`
- `POST /api/bl/extract-batch`（`folder_{i}` / `pdf_{i}_{j}` / `packing_{i}` / `weekly_packing` / `tracking` / `customer`）—— 批量
  - 响应：`{ ok, tickets:[{folder, record, warning}], bl_fields, bl_header, telex_fields }`
  - `warning`：箱货清单/物流追踪表解析失败的非阻断提示；保函「运输工具」为空（非海运单 `@` 占位符）提示手工填写（等于船名航次）；**箱数/国家不一致**（底单 vs 箱货清单 vs 文件夹名）提示核对原底单/箱货清单
- `POST /api/bl/generate`（`{ customer, tickets:[{folder, record}] }`）
  - 响应：`{ ok, zip, previews:[{folder, bl, telex}] }`（相对路径，经 `/api/bl/file/<rel>` 访问）
  - ZIP 含三份：`{命名}提单.pdf` + `{命名}电放保函.docx` + `{命名}底单.pdf`（命名 = 文件夹名把「易通报关资料」替换成类型词）
- `GET /api/bl/file/<path:rel>` —— 生成文件服务（ZIP 下载 / 提单·保函 PDF 预览）
- `GET /api/bl/mappings` —— 返回 `{ ok, port_map:{origin,destination}, channel_map, customs_office_map }`
- `POST /api/bl/mappings`（`{ port_map:{origin,destination}, channel_map, customs_office_map }`）—— 保存映射（键值均为字符串，非法返回 400）

## 关键实现

### 提单（邮件合并）

- 模板 `提单模板.docx` 含 12 个 MERGEFIELD 域（重复出现 32 次）。
- `clean_template.dedupe_template` 先去重重复域，`docx-mailmerge` 填域，保格式不漂移。
- **注意**：字段值落在文本框 `w:txbxContent` 内，python-docx 的 `paragraphs`/`tables` 读不到，验证需读原始 `document.xml`。

### 电放保函（程序重建，精确复刻参考模板）

- `fill_telex_docx` 用 python-docx 逐段生成，字体/字号/对齐/边距/标点严格对齐参考「拓锐6月电放保函」模板：中文宋体/等线、英文 Times New Roman；标题 17/15pt、正文 14pt、声明 12pt bold（对齐 Heading3 默认）；页边距上 2.61/下 0/左 3.15/右 2.29cm、header_distance 1.63cm。
- **页眉**：客户公司名（拓锐=`广州拓锐科技有限公司`，取自 `CUSTOMERS` 配置）等线 15.5pt bold + 底边框黑色分隔线（0.75pt、宽 415.35pt，右缩进 42.35pt 对齐标准答案 shape 宽度）。
- 字段含 `柜号 (CONTAINER NO.)` 行；TO / FROM / 申请单位 / 申请日期用全角冒号，字段标签用英文冒号+空格；申请日期无前导零（`2026 年 7 月 8 日`）。**申请日期 = 到港时间**（物流追踪表 ETA 提前两天，ETA−ETD<2 天取 ETA；无追踪表回退今天）。
- 收货人(保函) 与 consignee 独立：提单用大写无空格、保函用正常大小写带空格。

## 已验证（2026-09-04，拿 6 月真实底单比对）

- 6/9 文字层底单逐字段提取正确（提单号/柜号/箱数/总重量/consignee/shipper/目的港/起运地/品名/起运日期）。
- 3/9 扫描件底单无文字层 → 返回空字段，需人工填写。
- 提单 12 字段邮件合并落位正确；保函字段正确（含柜号 + 独立收货人）。
- **拓锐固定 shipper/consignee**：常量硬编码，与正确提单/保函一致（提单大写无空格、保函正常大小写带空格）。
- **箱货清单汇总**（`箱货清单-TRKJ26060003.xlsx` ↔ 提单 BG20260605002）：品名 5 个英文去重、总体积 4.307CBM、总重量 783.48kg、箱数 54 全部正确落位。
- ⚠️ 参考提单 BG20260605002 品名列仅 4 个（漏写 `ROSE PROJECTOR LIGHT` 玫瑰投影灯），但 54 箱/783.48kg/4.307CBM 都含这 10 箱 → 本功能按箱货清单全量汇总 5 个；用户已确认参考提单确属漏写，5 个为准。
- 船名航次在底单「运输工具名称及航次号」列：海运能取到船名航次（如 `CMA CGM MANTA RAY/0GVMWE`），非海运（铁路/卡航）该列为 `@` 占位符 → 空，需人工填船名航次。保函「运输工具」= 船名航次，手工填一个即可（不再回退提单号）。**过长时自动缩字适应、不换行**：`_fit_field_font` 估算文本框内宽度，超宽按 0.5pt 步进缩到下限 7pt（`CMA CGM MANTA RAY/0GVMWE` → 7pt 单行），短文本保持模板默认字号。
- 仍留人工的只有：起运地/目的港英文（底单是中文「阿拉山口铁路/英国」，提单要 `CHONGQING`/`MALASZEWICZE`）+ 实际起运日期（上传物流追踪表可按 FBA ID 匹配 ETD 开船日自动填，匹配不到时仍需人工）。
- **物流追踪表匹配 ETD**（`拓锐物流追踪表更新（易通科技）2026.8.31.xlsx` ↔ 海运单 BG20260508013）：从箱货清单 H 列取 FBA ID，在追踪表 5 个 sheet（含「最新物流动态」）按 Shipment ID 匹配，16/16 命中，ETD=2026-05-14 → 覆盖「起运日期」= `2026-05-14`；船名航次保持底单 `UN9949778/009E`（用户确认不用追踪表覆盖）。
- **PDF 生成（原生 LibreOffice）**：本机装 LibreOffice 26.8.0，`docx_to_pdf()` 经 `_find_soffice()` 定位 `soffice.exe` 转 PDF。全链路（`:3001` 代理 → Flask `:5000`）实测 ZIP 含 `提单.pdf`(397KB) + `电放保函.pdf`(69KB)，不再依赖 Docker。
- **提单号 = 底单「提运单号」标签取值**：`_value_by_label(lines, "提运单号")` 直接定位标签、取下方同列值，海运/铁路布局通用。实测铁路单取 `800620260613122207`、海运单（BG20260508013）取 `G2605115309`（原逻辑对海运单取空）。
- **海运单（BG20260508013 + 箱货清单 TRKJ26050005）全字段实测**：提单号 `G2605115309`、船名航次 `UN9949778/009E`、柜号 `FFAU6162017`、箱数 45、总重量 695.11、总体积 `3.8859`（无「体积」列，回退长宽高×箱数算出）、品名 4 个英文（含 `INS STYLE NORTHERN LIGHTS NIGHT LIGHT` 北极光灯）。
- **柜号 = 备注「集装箱标箱数及号码」上下文提取**（2026-09-09）：`_container_numbers` 从备注行 `集装箱标箱数及号码：N;XXXX1234567;` 里 findall 柜号（海运/铁路格式一致），多柜分号拼接去重。不再用全文本 `[A-Z]{4}\d{7}` 搜索——海运提单号（`ZIMUNGB1391012S`）含 `UNGB1391012` 会被误当柜号（多票复现）。
- **排版修复（2026-09-05）**：①提单运费恢复带空格「FREIGHT PREPAID」+ 文本框加宽 96→115pt，PDF 完整单行显示（对齐正确提单，正确提单是**带空格**非连写）；②品名恢复带空格（`CILP TABLE LAMP` 原样）+ 品名文本框 109→240pt + 左对齐，4 个品名全部单行、不换行、不溢出；③起运日期转 `DD MMM YYYY`（`2026-05-14` → `14 MAY 2026`，带空格）。
- **港口/渠道映射（2026-09-05）**：海运单渠道「易·22日达卡派包税」→ 目的港 `LONG BEACH,CA`、起运地 `YANTIAN`；提单 + 电放保函「目的地」均正确落位（`DESTINATION：LONG BEACH,CA`）。
- **7 项格式修复（2026-09-05，本地实测）**：①shipper/consignee/品名文本框段落加 `<w:jc w:val="left"/>` 左对齐，消除 LibreOffice 默认 justify 导致的中间大空格；②「SHIPPED ON BOARD」「FREIGHT PREPAID」恢复中间空格；③起运日期三处（SHIPPED ON BOARD / LADEN ON BOARD / Place and date of issue）统一 `14 MAY 2026`；④目的港左列 Port of Discharge（Text Box 10）LibreOffice 把 `margin:align left` 误渲染到右列（x0≈172），改 `column:posOffset=0` + VML `margin-left:0pt` 修复——现 Port of Discharge=x0≈43（左）、Place of Delivery=x0≈180（右），与正确提单一致。⑤物流追踪表上传入口已在前端 `/bl-review` 提供（用 FBA ID 匹配 ETD 覆盖「起运日期」）。

## 批量工作流（2026-09-05）

### 拆分底单合并
- `merge_declaration_pdfs(pdf_paths, output)`（pypdf）按文件名关键词排序 `报→放→委托` 后合并。参考 7 月单：一票底单常拆成「报.pdf / 放.pdf / 委托.pdf」三份。
- ⚠️ 上传保存必须保留原始文件名（`batch_{i}_{j}_{stamp}_{原文件名}.pdf`），排序 key 靠文件名里的「报/放/委托」——曾因重命名为 `batch_0_0_xxx.pdf` 丢失关键词，误把委托书当第一页，提取全空。

### 命名规则
- `derive_output_name(folder, kind)`：文件夹名里的「易通报关资料」替换成 `kind`（提单/电放保函/底单），其余严格照抄。
- 例：`7月第1周-(易通报关资料）-BG20260703003 4票 英国快铁不包税-卡派 39箱 是8` → `7月第1周-(提单）-BG20260703003 4票 英国快铁不包税-卡派 39箱 是8`

### 箱数/国家一致性校验
- `validate_ticket(folder, rec, packing)`：对比三方 —— 报关底单（`箱数`/`运抵国`）vs 箱货清单（`总箱数(CTN)` 求和/`国家` 列）vs 文件夹名（`N箱`/国家），有值且互相不同则 warning「原底单/箱货清单有问题」。
- `parse_packing_list` 新增返回 `total_boxes`（CTN 求和）与 `country`（「国家」列首值，兼容「目的国」）。
- `extract_customs_data` 新增 `运抵国`（指运港值行倒数第 3 token）供校验。
- ⚠️ 文件夹名「欧洲」是运输走廊（欧洲铁路）非目的国，已从国家表剔除避免误报（欧洲铁路单底单运抵国=德国、箱货清单=德国，文件夹名无国家 → 一致不提醒）。

### 周汇总箱货清单匹配（整周一份，可选）
- `parse_packing_list_weekly(xlsx)`：按「工作号」列分组（每组=国家/渠道/箱数/重量/品名/体积/系统SO），支持整周多票合并一份 xlsx。
- `match_weekly_packing(rec, groups)`：用「运抵国 + 件数(箱数) + 毛重」回溯匹配每票底单对应的一个或多个工作号（一票合并报关常对应多工作号，凑箱数 + 验重量容差 0.5kg），命中则覆盖品名/总体积/渠道/箱数/国家；匹配失败回退文件夹内单票清单。
- 前端 `/bl-review` 提供「周汇总箱货清单」上传入口（`weekly_packing=`），与每票 `packing_{i}` 兼容（周汇总优先、单票兜底）。

### 格式
- 提单只产 PDF、保函只产 DOCX、底单只产 PDF；ZIP 仅含这三份。
- 保函预览 PDF 由 docx 转（`preview/` 子目录），不进 ZIP。

### 已验证（7 月第 1 周 4 票实测）
- 拆分底单 3 份合并 4 页、提取提单号/箱数/柜号正确；箱数/国家三方一致无误报。
- 命名、ZIP 三份、预览 PDF、file 服务全链路 200。

### 已验证（6 月第 3/4 周，与标准答案比对）
- **6 月第 4 周 4 票**：命名 12 个文件名 100% 与标准答案一致；提单号/柜号/箱数/重量逐项一致；箱数/国家校验无误报。两票扫描件（BG20260625003 卡航、BG20260627019 联邦IP）无文字层需人工填；运输工具（=船名航次）需人工填。
- **6 月第 4 周 待确认项**（均非 bug）：①票1 品名 5 vs 标准答案 4——箱货清单 TRKJ26060019 确有 5 个英文品名（含 `MOONLIGHT BUNNY NIGHT LIGHT PROJECTOR`），标准答案提单漏写 1 个（同已知坑 #48 模式）；②总体积 8.5158 vs 8.5143——四舍五入差异（~0.018%），本功能按长×宽×高×箱数逐行精确求和。
- **6 月第 3 周 2 票**：票2（英国 40箱）周汇总匹配成功（2 工作号合并，品名 4/体积 3.2918/重量 605.41 正确），箱数/国家一致；票1（联邦IP 5箱）扫描件需人工填。发现两个非代码问题：①票2 源文件夹名「英国铁路」vs 标准答案「英国快铁」（命名混用，严格照抄）；②票2 是升级公路运输，目的港应为 The U.K. 但渠道映射为 MALASZEWICZE（见 dev-spec 已知坑 #59）。

### 已验证（7 月第 1 周 4 票实测，2026-09-07）
- 4 票（欧洲铁路整柜直送 90箱 / 英国快铁不包税 39箱 / 加拿大海运卡派 6箱 / 19箱）全部提取 + 周汇总匹配成功，箱数/毛重/品名/体积与底单对齐；英国票「委托/报/放」三份拆分底单合并正常。
- 新增渠道映射：`美转加美森-加东-卡派`、`美转加美森-加西-卡派` → `LOS ANGELES,CA`（美转加业务海运段到洛杉矶再转加拿大，提单目的港写洛杉矶）。
- `霍尔果斯 → GUANGZHOU` 确认正确（广州启运，霍尔果斯口岸出境）。
- 品名改为「去版本号 + 全大写」：`Blue Ocean Dream Galaxy Projector 3.0` → `BLUE OCEAN DREAM GALAXY PROJECTOR`。体积维持长×宽×高×总箱数逐行合计（用户确认算法正确）。
- 新增「系统SO」字段：从箱货清单 A 列（系统SO）提取，审核表展示在「票/文件夹」后第一列、只读不可编辑（一票多工作号时用换行合并多个 SO、分行显示），方便按 SO 查找相关信息。

## 已知坑

见 `docs/dev-spec.md` 已知坑 #44 起。
