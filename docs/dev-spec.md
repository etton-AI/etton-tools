# ETTON TOOLS — 开发者规格文档

> 最后更新: 2026-07-21 | 维护者: berry-bi

---

## 1. 项目概述与技术栈

**ETTON 效率提升助手** — 基于 Next.js 的 Web 工具集，面向易通科技内部物流操作，提供保单投保区间拆分、太平洋货箱清单转换、多供应商对账引擎、皮皮熊账单拆分、延讯下单优化五个核心功能。

### 技术栈

| 类别 | 技术 | 版本 |
|------|------|------|
| 框架 | Next.js (App Router) | `^15.3.3` |
| UI | React | `^19.1.0` |
| 语言 | TypeScript | `^5.8.3` |
| 样式 | Tailwind CSS 4 (PostCSS 插件) | `^4.1.8` |
| Excel 读写 | exceljs | `^4.4.0` |
| ZIP 打包 | jszip | `^3.10.1` |
| Lint | ESLint 9 flat config | `^9.27.0` |
| 运行时 | Node.js 22 (Alpine) | — |
| 部署 | Docker → GHCR → Sealos K8s | — |

### 仓库信息

- **GitHub**: `etton-AI/etton-tools`
- **容器镜像**: `ghcr.io/etton-ai/etton-tools:latest`
- **Sealos Ingress**: `vftnaopzqgqv.cloud.sealos.io`

---

## 2. 目录结构与各模块职责

```
ETTON 电商AI/
├── .github/workflows/
│   └── docker-build.yml              # CI: push main → 构建 Docker 镜像 → 推送 GHCR
├── k8s/
│   └── deploy.yaml                   # K8s Deployment + Service + Ingress
├── public/
│   ├── output/                       # 历史测试输出（仅供开发参考）
│   │   ├── 不足5000RMB.xlsx
│   │   ├── 5000-10000RMB.xlsx
│   │   ├── 10000-20000RMB.xlsx
│   │   ├── 20000-30000RMB.xlsx
│   │   └── 30000-40000RMB.xlsx
│   └── templates/
│       └── 易通下单模版.xlsx          # 延讯下单优化的输出模板
├── src/
│   ├── app/
│   │   ├── layout.tsx                # 根布局：html lang=zh-CN + body 全局样式
│   │   ├── page.tsx                  # 首页：四个工具入口卡片导航
│   │   ├── globals.css               # Tailwind CSS 4 入口 (@import "tailwindcss")
│   │   ├── insurance-split/
│   │   │   └── page.tsx              # 保单拆分页面（完整客户端组件）
│   │   ├── pacific-convert/
│   │   │   └── page.tsx              # 太平洋转换页面（含汇率配置）
│   │   ├── reconciliation/
│   │   │   └── page.tsx              # [旧] 单供应商对账页面（已合并入多供应商对账引擎）
│   │   ├── multi-supplier-reconciliation/
│   │   │   └── page.tsx              # 多供应商对账引擎页面（17 家供应商）
│   │   ├── pipixiong-split/
│   │   │   └── page.tsx              # 皮皮熊账单拆分页面
│   │   ├── yanxun-convert/
│   │   │   └── page.tsx              # 延讯下单优化页面
│   │   └── api/
│   │       ├── split-insurance/
│   │       │   └── route.ts          # POST 上传 + GET 下载（session 管理）
│   │       ├── convert-pacific/
│   │       │   └── route.ts          # POST 上传（含汇率）+ GET 下载
│   │       ├── reconciliation/
│   │       │   └── route.ts          # [旧] 单供应商对账 API
│   │       ├── multi-supplier-reconciliation/
│   │       │   └── route.ts          # POST 双文件上传 + GET 下载（多供应商）
│   │       ├── pipixiong-split/
│   │       │   └── route.ts          # POST 上传 + GET 下载 ZIP
│   │       └── yanxun-convert/
│   │           └── route.ts          # POST 批量上传（多文件/ZIP）+ GET 下载（单票/ZIP）
│   ├── components/
│   │   └── Header.tsx                # Header (sticky) + Footer（同文件导出）
│   └── lib/
│       ├── split-insurance.ts        # 保单拆分核心逻辑 (556 行)
│       ├── convert-pacific-insurance.ts  # 太平洋转换核心逻辑 (502 行)
│       ├── reconciliation.ts         # [旧] 单供应商对账逻辑 (272 行)
│       ├── multi-supplier-reconciliation.ts  # 多供应商对账引擎 (17家供应商配置，~870行)
│       ├── pipixiong-split.ts        # 皮皮熊拆分核心逻辑 (371 行)
│       └── yanxun-convert.ts         # 延讯下单优化核心逻辑（含 generateYanxunZip）
├── 保单拆分功能/                     # 原型/实验脚本（不参与构建）
│   ├── split_insurance.js            # 原始 Node.js 拆分脚本
│   ├── read_excel.js                 # Excel 读取调试脚本
│   ├── 易通投保区间拆分规则说明.md    # 业务规格文档
│   └── 太平洋货箱清单转换功能/       # 太平洋模板和分析文件
├── 比价工具/                         # 遗留价格表（仅供对比参考，已于 2026-06 拆分为 price-system 仓库）
├── Dockerfile                        # 双阶段构建 (builder + runner)
├── next.config.ts                    # output: "standalone"
├── postcss.config.mjs                # @tailwindcss/postcss 插件
├── eslint.config.mjs                 # ESLint 9 flat config
├── tsconfig.json                     # strict, path alias @/* → ./src/*
├── package.json
└── CLAUDE.md                         # Claude Code 项目指引
```

---

## 3. 已实现的模块

### 3.1 首页 — `/`

**文件**: `src/app/page.tsx`

纯展示型导航页，包含五个卡片链接：
- `/insurance-split` — 保单投保区间拆分
- `/pacific-convert` — 太平洋货箱清单转换
- `/multi-supplier-reconciliation` — 多供应商对账引擎
- `/pipixiong-split` — 皮皮熊账单拆分
- `/yanxun-convert` — 延讯下单优化

无服务端逻辑，纯客户端渲染。

---

### 3.2 保单拆分 — `/insurance-split`

**文件**: `src/app/insurance-split/page.tsx`

#### 功能
1. 拖拽或点击上传 `.xlsx` / `.xls` 文件
2. 调用 `POST /api/split-insurance` 处理
3. 展示拆分结果：每个区间的箱数、重量、体积、箱组数
4. 每个区间包含详细表格（源行、箱数、总价原币、币种、每箱原币、每箱RMB）
5. 支持下载单个区间 `.xlsx` 或全部 `.zip`

#### 状态管理
- `status`: `"idle"` → `"uploading"` → `"success"` / `"error"`
- 结果存储在 API 层内存 session（非客户端状态）

#### 关键细节
- 5 个预定义区间配色（蓝/翠绿/琥珀/橙/玫红）
- 数字格式化 `fmt()` 保留 2 位小数，`fmtRMB()` 保留 2 位小数
- 错误状态包含"重新上传"按钮

---

### 3.3 太平洋转换 — `/pacific-convert`

**文件**: `src/app/pacific-convert/page.tsx`

#### 功能
1. 与保单拆分类似的上传流程
2. 额外包含汇率设置区域：USD(默认7), EUR(默认8), GBP(默认9), JPY(默认0.04)
3. 调用 `POST /api/convert-pacific`，汇率随 FormData 一起发送
4. 结果按单箱货值区间拆分，展示区间名和行数
5. 提示框：提单号、船名航次、柜号需手动补充

#### 关键细节
- `RateInput` 子组件：数字输入框，step=0.01
- 10 个预定义区间颜色（循环）
- 结果卡片比保单拆分更简洁（仅显示区间名 + 行数 + 下载按钮）

---

### 3.4 组件: Header + Footer

**文件**: `src/components/Header.tsx`

- **Header**: `position: sticky; top: 0; z-index: 50`，白色半透明模糊背景，标题 "ETTON 效率提升助手"
- **Footer**: 上边框，版权年份动态计算 `new Date().getFullYear()`

---

### 3.5 库: 保单拆分核心逻辑

**文件**: `src/lib/split-insurance.ts` (556 行)

#### 导出类型

```typescript
interface BoxGroup {
  rows: number[];      // 1-indexed Excel 行号
  boxes: number;       // 箱数
  totalPrice: number;  // 总价（原币）
  currency: string;    // 币种
  rate: number;        // 汇率
  perBoxOrig: number;  // 每箱原币
  perBoxRMB: number;   // 每箱RMB
}

interface IntervalResult {
  name: string;        // 区间名 (如 "不足5000RMB")
  fileName: string;    // 文件名
  totalBoxes: number;
  totalWeight: number;
  totalVolume: number;
  groupCount: number;
  groups: BoxGroup[];
  buffer: Buffer;      // xlsx 二进制
}

interface SplitResult {
  sourceFile: string;
  totalBoxes: number;
  totalGroups: number;
  intervals: IntervalResult[];
}
```

#### 导出函数

| 函数 | 说明 |
|------|------|
| `processSplit(filePathOrBuffer, fileName)` | 主入口：加载 Excel → 解析分组 → 计算区间 → 生成文件 |
| `generateZip(result)` | 将 SplitResult 所有区间打包为 ZIP（含 `split_summary.json`） |

#### 核心逻辑

1. **数据源**: Sheet `"ETTON电商物流 下单模板"`，数据从 Row 25 开始
2. **汇率** (`RATES`): USD=7, EUR=8（硬编码）
3. **列映射** (0-indexed):
   - `COL_A(0)`: 订单号（空值 = 数据结束）
   - `COL_G(6)`: 箱数 (>0 = 新箱组, 0 = 混入当前组)
   - `COL_J(9)`: 总价
   - `COL_K(10)`: 币种
   - `COL_Q(16)`: 重量(KG)
   - `COL_R(17)`, `COL_S(18)`, `COL_T(19)`: 长/宽/高(CM)
4. **预定义区间**:
   - 不足5000RMB: [0, 5000)
   - 5000-10000RMB: [5000, 10000)
   - 10000-20000RMB: [10000, 20000)
   - 20000-30000RMB: [20000, 30000)
   - 30000-40000RMB: [30000, 40000]
5. **文件生成策略**: 原始 workbook 序列化一次 → 每个区间反序列化克隆 → 删除不属该区间的行 → 更新头部（标题/箱数/重量/体积） → 添加 W 列（每箱RMB）
6. **DISPIMG 保留**: exceljs 不支持 `=DISPIMG()` 公式的内嵌图片，通过 `injectCellImagesIntoOutput()` 从原始 xlsx zip 中提取 `cellImages.xml` + `media/` 注入输出文件

---

### 3.6 库: 太平洋转换核心逻辑

**文件**: `src/lib/convert-pacific-insurance.ts` (502 行)

#### 导出类型

```typescript
interface ExchangeRates {
  USD: number;  // 默认 7
  EUR: number;  // 默认 8
  GBP: number;  // 默认 9
  JPY: number;  // 默认 0.04
}

interface PacificDataRow {
  fbaId: string;
  description: string;    // 中文品名 + 英文品名
  qtyPcs: number;         // 申报总数量
  unitValue: number;      // 单个产品申报货值(USD)
  totalValue: number;     // 总申报货值
  currency: string;       // 申报币种
  ctns: number;           // 总箱数
  grossWeight: number;    // 单箱货物毛重(KG)
  measurement: number;    // 长×宽×高/1,000,000
  perBoxRMB: number;
  exchangeRate: number;
}

interface PacificSplitResult {
  sourceFile: string;
  totalRows: number;
  skippedRows: number;
  intervals: PacificIntervalResult[];
}
```

#### 导出函数

| 函数 | 说明 |
|------|------|
| `convertAndSplitPacific(filePath, sourceFileName, rates)` | 主入口：读取源 → 列映射 → 数据转换 → 按区间拆分 |
| `generatePacificZip(result)` | 将所有区间输出打包为 ZIP |

#### 核心逻辑

1. **列映射**: 基于表头自动检测（搜索 "FBA ID", "中文品名", "英文品名" 等关键字）
2. **字段映射表**（源列 → 太平洋模板列）:

   | 太平洋列 | 源列 |
   |----------|------|
   | 入仓编号 | FBA ID |
   | DESCRIPTION | 中文品名 + 英文品名 |
   | QTY PCS | 申报总数量 |
   | UNIT VALUE | 单个产品申报货值(USD) |
   | TOTAL VALUE | 总申报货值 |
   | 币种 | 申报币种 |
   | CTNS | 总箱数(CTN) |
   | G.W.(KG) | 单箱货物毛重(KG) |
   | MEASUREMENT(CBM) | 长×宽×高/1,000,000 |

3. **区间生成**: 动态区间，从 40000RMB 起每 10000 一个区间，最后一个为 catch-all (>max)
4. **混合箱处理**: ctns=0 的行与前一行合并计算 perBoxRMB，但各行独立输出
5. **输出格式**: 18 列表头，预设列宽，边框，冻结首行

---

### 3.7 API: 保单拆分

**文件**: `src/app/api/split-insurance/route.ts`

#### POST `/api/split-insurance`
- **输入**: `multipart/form-data`，字段 `file` (`.xlsx` / `.xls`)
- **处理**: 临时写入 os.tmpdir() → 调用 `processSplit()` → 删除临时文件 → 存入内存 session
- **输出**: JSON `{ sessionId, sourceFile, totalBoxes, totalGroups, intervals[], downloads: { allZip, files[] } }`
- **限制**: `bodyParser.sizeLimit = "50mb"`

#### GET `/api/split-insurance`
- **参数**: `session` (session ID), `file` (文件名或 `all.zip`)
- **输出**: 单个 `.xlsx` 文件或 `.zip` 包
- **过期**: session 30 分钟无访问自动清除

---

### 3.8 API: 太平洋转换

**文件**: `src/app/api/convert-pacific/route.ts`

#### POST `/api/convert-pacific`
- **输入**: `multipart/form-data`，字段 `file` + `rateUSD` + `rateEUR` + `rateGBP` + `rateJPY`
- **处理**: 解析汇率 → 临时写入 → 调用 `convertAndSplitPacific()` → session 存储
- **输出**: JSON `{ sessionId, sourceFile, totalRows, skippedRows, intervals[], downloads }`
- **限制**: `bodyParser.sizeLimit = "50mb"`

#### GET `/api/convert-pacific`
- **参数**: `session`, `file`
- **输出**: 单个 `.xlsx` 或 `.zip`（文件名使用 `filename*=UTF-8''` 编码支持中文）
- **ASCII fallback**: 中文文件名替换为 `_` 以保证兼容性

---

### 3.9 天图请款对账 — `/reconciliation`

**文件**: `src/app/reconciliation/page.tsx` + `src/lib/reconciliation.ts` + `src/app/api/reconciliation/route.ts`

#### 功能

1. 上传两个 Excel 文件：天图供应商账单 + 内部请款明细
2. 自动识别 SO 号列和金额列（关键词匹配）
3. 按 SO 号 FULL OUTER JOIN，比对金额差异
4. 生成对账结果 Excel（差异行标红），支持下载

#### 核心逻辑 (`src/lib/reconciliation.ts`)

**导出函数**:
- `processReconciliation(tiantuPath, paymentPath, tiantuName, paymentName)` → `ReconciliationResult`

**算法流程**:
1. `parseFile()` — 自动检测表头中的 SO 列和金额列
2. `cleanAmount()` — 清洗金额（去除 ￥、¥、千分位逗号）
3. SO 号 `groupBy + sum` 聚合 → 两个 `Map<string, number>`
4. `fullOuterJoin()` — FULL OUTER JOIN，分类: 一致/金额差异/天图缺失/请款缺失
5. `buildOutputWorkbook()` — exceljs 生成输出，差异行红底、汇总行加粗

**列检测规则**:
- SO 列: 含 "so" / "运单" / "单号"
- 金额列: 含 "金额" / "费用" / "合计" / "amount"

---

### 3.10 多供应商对账引擎 — `/multi-supplier-reconciliation`

**文件**: `src/app/multi-supplier-reconciliation/page.tsx` + `src/lib/multi-supplier-reconciliation.ts` + `src/app/api/multi-supplier-reconciliation/route.ts`

#### 功能

1. 上传两个 Excel 文件：供应商账单 + 内部请款明细
2. 根据文件名自动识别供应商（支持 17 家供应商）
3. 自动检测表头行位置（扫描前20行，不再写死固定行号）
4. 按 SO 号 FULL OUTER JOIN，比对金额差异
5. 生成对账结果 Excel（差异行红标 + 汇总行 + 冻结首行），支持下载

#### 支持的供应商（17 家）

| # | 供应商 | 文件名特征 | SO 列 | 金额列 | 表头行 |
|---|--------|-----------|-------|--------|--------|
| 1 | 天图通逊 | `*天图*` / `*通逊*` / `*Tiantu*` | 客户运单号 | 应收金额 | 10 |
| 2 | 星链/易通 | `*星链*` / `*易通*` / `*ETTON*` | 客户参考号 | 应收金额 | 5 |
| 3 | 航乐 | `*航乐*` | 运单号 | 合计应收 | 4 |
| 4 | 跨境堡/英美 | `*英美*` / `*跨境堡*` | 客户运单号 | 金额 | 2 |
| 5 | 美琦/皓辉 | `*美琦*` / `*皓辉*` / `*zsetton*` | 客户运单号 | 合计金额 | 5 |
| 6 | 心一 | `*心一*` | 客户运单号 | 人民币应收金额 | 8 |
| 7 | 凯鑫 | `*凯鑫*` | 客户运单号 | 金额 | 4 |
| 8 | 华威尔 | `*华威尔*` | 客户运单号 | 金额 | 4 |
| 9 | 天龙 | `*天龙*` | 客户单号 | 金额 | 2 |
| 10 | 松杰 | `*松杰*` | 客户参考号 | 应收金额 | 3 |
| 11 | 安时达 | `*安时达*` | 单号 | 总价 | 5 |
| 12 | 鸿珉 | `*鸿珉*` | 原单号 | 保费(RMB) | 1 |
| 13 | 太平洋 | `*太平洋*` | 客户单号 | 金额 | 2 |
| 14 | 一腾 | `*一腾*` | 原单号 | 费用合计 | 2 |
| 15 | 乐丰 | `*乐丰*` | 运单号 | 金额 | 1 |
| 16 | 深圳总部 | `*总部*` / `*散货*` / `*深圳*` | SO号码 | 总费用 | 3 |

#### 核心逻辑 (`src/lib/multi-supplier-reconciliation.ts`)

**导出函数**:
- `processMultiSupplierReconciliation(billPath, paymentPath, billName, paymentName, supplier?)` → `MultiReconResult`
- `detectSupplier(filename)` → `string | null`
- `getAvailableSuppliers()` → `string[]`

**算法流程**:
1. `detectSupplier()` — 根据文件名正则匹配供应商
2. `parseSupplierBill()` — 自动扫描前20行定位表头 → 模糊匹配 SO 列和金额列 → 解析数据
3. `parsePaymentFile()` — 自动扫描前15行定位表头 → 关键词评分匹配 SO 列和金额列（优先本位币列）→ 解析数据
4. `fullOuterJoin()` — 两个 Map 做 FULL OUTER JOIN，按状态分类排序
5. `buildOutputWorkbook()` — exceljs 生成输出 Excel

**列检测规则**:
- **供应商账单 SO 列**: `fuzzyFindColumn()` 四级匹配：精确匹配 → 包含匹配 → 关键词拆分匹配（≥60%命中）→ 备选列
- **供应商账单金额列**: 同四级匹配 + 自动按数值密度检测
- **请款明细 SO 列**: 关键词匹配（`系统SO号` > `SO号` > `SO` > `运单号` > `SO号码`）+ 数据验证（连续行含有效SO号≥2）
- **请款明细金额列**: 评分机制 — 更长的关键词（如 `金额(本位币)`）得分更高，优先匹配本位币列 → fallback 自动检测数值列

**表头自动检测** (2026-07-21 新增):
- 不再写死 `header_row`，改为扫描前 20 行
- 每行检查是否包含 SO 列关键词（`fuzzyFindColumn`）
- 找到匹配后立即使用该行作为表头
- 失败时回退到配置的固定行号，并输出详细调试信息（文件前几行预览）

**金额列评分机制** (2026-07-21 新增):
- 请款明细中可能有"金额"(原币)和"金额(本位币)"(RMB)两列
- 评分 = 关键词长度，`金额(本位币)` 长度 > `金额`，优先匹配本位币列
- 确保对账使用统一币种，避免原币和本位币混用导致差异

**跳过行检测** (2026-07-21 新增):
- 供应商账单末尾的"开户人"、"开户行"、"账号"等行内容不含有效 SO 号，`isValidSONumber()` 自动过滤
- 额外配置 `skip_keywords` 处理"费用确认单"等标题行

---

### 3.11 皮皮熊账单拆分 — `/pipixiong-split`

**文件**: `src/app/pipixiong-split/page.tsx` + `src/lib/pipixiong-split.ts` + `src/app/api/pipixiong-split/route.ts`

#### 功能

1. 上传皮皮熊合并账单 Excel
2. 自动识别货代识别号、报关单号和各费用列
3. 按报关单号拆分，每个报关单号生成 3 个账单：
   - 国内账单（报关费 + 港杂费 + 拖车费）
   - 国外账单（海运费 + 税金）
   - INVOICE（总费用明细）
4. 按货代识别号分目录打包 ZIP 下载

#### 核心逻辑 (`src/lib/pipixiong-split.ts`)

**导出函数**:
- `processPipixiongSplit(filePath, fileName)` → `PipixiongSplitResult`
- `generatePipixiongZip(result)` → `Buffer`

**列检测规则**（关键词模糊匹配）:
- 货代识别号: "货代" / "识别号" / "zmgs"
- 报关单号: "报关单号" / "海关编号" / 22位数字检测
- 报关费: "报关费"
- 港杂费: "港杂" / "港口费"
- 拖车费: "拖车" / "内陆费"
- 海运费+税金: "海运" / "税金" / "关税"
- 总费用: "总费用" / "合计"

**输出结构 (ZIP)**:
```
货代识别号1/
├── 国内账单_报关单号.xlsx
├── 国外账单_报关单号.xlsx
└── INVOICE_报关单号.xlsx
...
```

---

### 3.12 延讯下单优化 — `/yanxun-convert`

**文件**: `src/app/yanxun-convert/page.tsx` + `src/lib/yanxun-convert.ts` + `src/app/api/yanxun-convert/route.ts`

#### 功能

1. 上传延讯下单发票 Excel（发货单 sheet，即第一个 sheet）；支持**批量多选**或**上传 ZIP 包**（一次处理几十票）
2. 自动提取顶部信息：运输方式、正式报关、带电、目的地、渠道、FBA号/海外仓、调拨单号、ReferenceID、总箱数
3. 解析货箱清单（箱号/品名/英文/材质/用途/发货数量/申报货值/海关编码/毛重/长宽高/币种）
4. 映射填充到易通下单模版（`public/templates/易通下单模版.xlsx`）的顶部字段 + 数据区
5. 输出文件按「`ETTON_FBA号`（FBA 场景）/ `ETTON_调拨单号`（海外仓场景）」命名（如 `ETTON_FBA19MX7M8KR.xlsx` / `ETTON_TF2608270070.xlsx`），支持**单票下载**或**打包 ZIP 下载**（含 `转换结果汇总.json`，重名自动加 `_2` 后缀）

**批量处理流程**:
- 前端 `input multiple` / 拖拽多文件 → `formData.append("files", f)` 逐个追加
- 后端 POST 遍历 `formData.getAll("files")`；`.zip` 先 `JSZip` 解压提取内部 `.xlsx/.xls`（跳过 `.`/`~$` 临时文件），单个文件失败记录到 `failed[]` 不中断整体
- GET `?file=all.zip` 打包全部成功结果；`?file=<文件名>` 下载单票

#### 核心逻辑 (`src/lib/yanxun-convert.ts`)

**导出函数**:
- `convertYanxunToEtton(filePath, sourceFileName)` → `YanxunConvertResult`
- 读取前先经 `loadFirstSheetOnly()` 精简：用 JSZip 把 xlsx 精简成只含发货单 sheet（删除隐藏的 VLOOKUP 数据源表与外部链接），避免大文件全量解析 OOM（见「已知坑」#42）

**顶部字段映射**（延讯发货单 → 易通模版）:

| 延讯字段 | 易通字段 | 易通位置 |
|---------|---------|---------|
| 运输方式 | 业务类型 | B6 |
| 正式报关 | 报关方式 | B8 |
| 带电 | 带电 | F2（否则不填） |
| 目的地（第 1 个） | 收件人国家 | F9（仅 FBA 场景） |
| 目的地（第 2 个） | 仓库代码 | F10（仅 FBA 场景） |
| 目的地（地址文本） | 私人地址/海外仓 | F14~F22（仅海外仓场景，解析收件人姓名/公司/国家/城市/州/邮编/联系方式/地址） |
| 渠道 | 备注 | B21 |
| FBA号/海外仓（是否含「海外仓」） | 仓点类型 | F7/F8（FBA 场景填 FBA；海外仓场景清空该组走「私人地址/海外仓」） |
| FBA号 / 调拨单号 | 文件名 | —（FBA 场景 `ETTON_FBA号.xlsx`，海外仓场景 `ETTON_调拨单号.xlsx`） |
| ReferenceID | 货件追踪编码 | 数据区 c2（FBA 场景）；`/`（海外仓场景，无追踪编码） |
| 总箱数 | 总箱数 | B18 |

**报关方式映射**:
- `公司自报` → `普通报关`（买单）
- `永德吉报关` → `报关退税`
- `否` → `普通报关`（买单）

**货箱清单列映射**（延讯箱单 → 易通 24 列数据区）:
- `箱号` → Shipment ID；`英文` → Name(En)；`品名` → Name(Ch)；`材质` → Material；`用途` → Use
- `发货数量` → Quantity；`申报货值` → Unit Price；`数量×单价` → Total Price（保留完整精度，不四舍五入）；`币种` → currency（FBA 场景从「币种」列取；海外仓备货单无币种列，默认 `USD`）
- `海关编码` → HS Code；`brand`/`Model`/`Brand Type` → 无；`毛重` → 净重&毛重；`长/宽/高` → 尺寸
- `ReferenceID`（顶部） → Reference ID（FBA 场景取顶部 ReferenceID；海外仓场景填 `/`）；`链接` → 链接（无数据填 `0`）；`图片` → `0`；`是否申报`/`申报数量` → 留空

**表头自动检测**:
- 延讯箱单表头行：扫描前 60 行，定位同时含「箱号」+「发货数量」的行
- 列映射：基于表头关键词 `startsWith` 匹配（长/宽/高用「长（」「宽」「高」前缀避免误匹配「超围长」）

**混箱处理**:
- 同一箱号出现多行 = 混箱
- 仅首行保留毛重/长宽高（Number 列首行=1），其余行 Number/净重/毛重/长/宽/高 **补 `0`**（不留空白）
- 避免导入系统时把混箱当作新的一箱货；补 0 保证下游导入不因空值报错

**预计总重量/总体积**:
- 总重量 = 各箱毛重合计（混箱仅首行计入）；总体积 = 各箱「体积CBM」列合计（混箱仅首行计入），自动填充 B19/B20
- 体积用延讯「体积CBM」列（每箱已 ROUND 到 2 位），而非直接 长×宽×高 计算，避免与标准答案的取整差异

**必填项校验**:
- 转换前校验 运输方式/正式报关/渠道，缺失时抛错（批量场景记录到 `failed[]`）
- **FBA 场景**额外校验 目的地国家/仓库代码/FBA号（缺失报「无目的地国家/无仓库代码/无FBA号」）
- **海外仓场景**额外校验 收件人姓名/地址/调拨单号（缺失报「无收件人姓名/无收件人地址/无调拨单号」）
- 渠道从「物流商/渠道」标签右侧多格扫描，跳过含「发件人/地址/邮编/收件/电话」的地址类文本，缺失报「无渠道名」，不 fallback、不误填发件人地址
- 二选一必填组：`仓点类型+收件人国家+仓库代码`（FBA 地址库）与「私人地址/海外仓」二选一；场景判断依据为「FBA号/海外仓」字段值是否含「海外仓」——FBA 场景走第一组（仓点类型=FBA），海外仓场景清空第一组、走第二组（从「目的地」地址文本自动解析收件人信息并填入）
- 模板自带默认值：发货公司、是否合并报关，无需处理
- **服务渠道（B7）延讯无法自动映射** → 不做校正、不提示（延讯渠道名与易通服务渠道名非一一对应，由人工选定）

---

### 3.13 TR入仓数据整理 — `/warehouse-entry`

**文件**: `src/app/warehouse-entry/page.tsx` + `src/lib/warehouse-entry.ts` + `src/app/api/warehouse-entry/route.ts`（含 `/export`、`/history` 子路由）

#### 功能

1. 上传**客户数据**（一行一个产品）+ **供应商数据**（逐箱），自动匹配选数、校验报警，生成「出给客户」建议箱规
2. 前端**可编辑**建议长/宽/高/实重（实时重算材积重/计费重/三边和），**全局校验条**提示「出给客户总计费重 − 供应商总计费重」
3. 导出《出给客户.xlsx》，并**自动累积历史库**；支持历史库导入/导出备份
4. **导出格式对齐参考文件《拓锐…入仓数据（成本）.xlsx》**：表头两行（分组 + 列名）、三组并排对比（客户的 / 供应商 / 出给客户）、派生列用 Excel 公式（材积重/总重/计费重/差异）以便人工调整尺寸后自动重算；系统SO/客户渠道/国家/仓库代码/单证报关 从客户数据自动取数，出货日期/成本KG/渠道/总成本重 等无来源列留空供人工填写

#### 核心业务规则

- **产品唯一键** = 品名 + 长 + 宽 + 高 + 实重；**历史库同款判定** = 品名一致 且 客户长宽高**排序后**逐边差 ≤ 1cm 且 实重差 ≤ 1kg（仅品名相同但箱规/重量差异更大 → 视为不同产品，不参与历史对比；排序是为了忽略长宽高书写顺序差异）。**合并语义分两种**：自动累积（导出后）同款取计费重**更大者**；手动导入最终《出给客户的.xlsx》同款**直接覆盖**（以最终提供值为准）
- **FBA 匹配**：客户 FBA 12 位；供应商按格式取「U+流水号」前的值（`FBA19MYJ057TU000001` → `FBA19MYJ057T`，比固定取前 12 位更健壮）——
  - **天图格式**：`货箱编号`（`FBA15M8F4YZR` + `U000001`）
  - **英美入仓格式**：`扩展箱号`（`FBA19MYJ057T` + `U000001`）；箱规列 = `货箱重量(BI)/货箱长度(BJ)/货箱宽度(BK)/货箱高度(BL)/货箱材积重(BM)`，**忽略该表客户数据，只看 BI~BM**。自动检测：表头含「扩展箱号」即按此格式解析
- **客户数据三格式**：标准（一行一产品，`实重`列）、易通发票（表单头+明细表，`FBA货箱编号/中文品名/箱数件数/货箱重量/长宽高`）、货箱清单（表单头「货箱清单」+明细表，`FBA ID/中文品名/总箱数(CTN)/长宽高/单箱货物毛重`）。自动检测：表头含「FBA货箱编号」→ 发票格式（「货箱重量」→ 实重）；含「单箱货物毛重」→ 货箱清单格式（「单箱货物毛重」→ 实重）；否则标准格式
- **材积重** = 长×宽×高 ÷ 6000；**计费重** = `max(实重, 材积重)`，按计费重降序
- **箱规匹配**：长宽高向下取整后相等；**实重容差** ≤ 0.3kg 视为同产品
- **选数**（按历史参考值决定第 1/第 2 大）：有历史同款时 供应商第 1 大计费重 ≤ 历史最大计费重 → 取第 1 大，否则退取第 2 大；新品（无历史）取第 2 大（避免偶发偏大的异常箱）；仅 1 箱回退第 1 大
- **建议值** = 选中箱规长宽高 + 该箱规所有箱最大实重 + 公式材积重
- **出给客户 = 供应商选数箱规 + 最短边 +1**（放大最短边作安全余量，尽量不放大最大边）：尺寸 = 选中箱长宽高再放大最短边 +1；实重 = 该箱规所有箱最大实重；材积重 = 公式重算；计费重 = `max(实重, 材积重)`。放大约束：放大后材积重 − 客户材积重 `< 2`、计费重不超过历史最大计费重，任一不满足则不放大（退回供应商原尺寸）。例：供应商 `54×53×36`、客户 `51.5×51.5×36` → 最多 `54×53×37`
- **差异约束（对比客户申报）**：三边和差必须 `< 6`、材积重差必须 `< 2`；超限 → 报警并提示找供应商核查过机图，核实后再修改
- **参考优先级**：历史有数据用历史参考；有供应商数据用供应商数据参考（供应商实测为准）。供应商数据明显偏大（如 51/51/51 vs 客户 49/49/49，三边和差 ≥ 6）→ 仍采用供应商数据，但标红要求核查过机图
- **全局兜底校验**：导出前 Σ出给客户总计费重 > Σ供应商总计费重，否则等比例放大并标 `[全局调整]`（`AMPLIFY_RATIO = 1.02`）
- **报警阈值**（供应商**原始箱规** vs 客户，区分材积/实重主导）：**材积主导**（供应商体积重 ≥ 实重）→ 三边和差 ≥ 6、材积重差 ≥ 2；**实重主导**（供应商体积重 < 实重）→ 实重差 ≥ 0.5；另有 供应商过大箱（最大材积重 − 建议材积重 ≥ 2）、历史最大值 > 建议值（建议参考历史最大值放大）、未匹配 → `⚠需人工复核`。编辑建议值后前端 `recomputeAlarms` 实时重算（与后端口径一致）

#### 核心逻辑 (`src/lib/warehouse-entry.ts`)

**导出函数**: `parseCustomerFile` / `parseSupplierFile` / `buildSuggestions` / `exportOutputBuffer` / `loadHistory` / `saveHistory` / `accumulateHistory` / `importHistoryFromExcel` / `exportHistoryBuffer`

**导出列位**（40 列，与参考文件对齐）：A 出货日期 · B 系统SO · C 客户渠道 · D 国家 · E 仓库代码 · F 单证报关 · G FBA ID · H 中文品名 · I 总箱数 · J–N 客户长/宽/高/实重/材积重 · P 差异 · Q–W 供应商长/宽/高/实重/材积重/总实重/总材积重 · Y 差异 · Z–AH 出给客户长/宽/高/实重/材积重/总实重/总材积重/计费重/总计费重 · AI 成本KG · AJ 备注(报警) · AK 渠道 · AL 箱数 · AM 总计费重 · AN 总成本重。`SuggestionRow.supplier` 记录选数命中的供应商代表箱（未匹配时全 0），供「供应商」对比列输出。`SuggestionRow.so/channel/country/warehouse/customs` 五个客户元信息字段从客户数据自动取数（列名匹配：系统SO/客户渠道/国家/仓库代码/单证报关），写入 B–F 列。

**复用**: `cellText` / `cellNum` 安全读单元格逻辑（与 `yanxun-convert.ts` 同款实现）

#### API

- `POST /api/warehouse-entry`：上传两文件 → `{ rows, supplierTotal, summary }`
- `POST /api/warehouse-entry/export`：body `{ rows, supplierTotal }` → 下载 Excel（导出前全局兜底、导出后累积历史库）
- `GET /api/warehouse-entry`：读取历史库；`POST/GET /api/warehouse-entry/history`：导入/导出历史库

---

### 3.14 提单 + 电放保函 — `/bl-review`

**文件**: `src/app/bl-review/page.tsx` + `bl-service/`（Python Flask 后端；docx→PDF 用**原生 LibreOffice** 转换，不依赖 Docker）

#### 功能

1. **批量**：点按钮选文件夹，每个子文件夹 = 一票（同一票的底单 PDF + 箱货清单 xlsx 放同一文件夹）；「物流追踪表」单独上传。底单被拆成多份（报/放/委托）自动先合并
2. 审核表 = **提单 13 字段**（文件命名/shipper/consignee/提单号/柜号/船名航次/起运地/目的港/箱数/品名/总重量/总体积/起运日期）+ **保函 5 字段**（申请单位/运输工具/目的地/收货人(保函)/申请日期），逐格可编辑；**空白字段标红、含中文字段（申请日期除外）标红**提醒
3. 人工核对/补填后 → 生成**提单 PDF（仅 PDF）** + **电放保函 DOCX（仅 DOC）** + **底单 PDF**，页面内预览提单/保函，再一键 ZIP 下载三份文件
4. **星速(HNXS) 客户特殊流程**：Amazon FBA 直送，**无需电放保函**（generate 跳过 telex，ZIP 只含提单+底单）；无物流追踪表，改传「订单列表」xlsx（发往国家/开船时间/船名航次/业务类型）；shipper=境内发货人中文转拼音大写、consignee=`AMAZONFULFILMENTCENTER`、notify=`SAMEASCONSIGNEE`；目的港按「运输方式+发往国家」查 `xs_dest_map.json`（海运→实际港口、陆运→国家英文名），起运港海运`YANTIAN`/陆运`SHENZHEN`；前端每份 PDF 独立成一票（按文件名里 FBA 号分组，xlsx 忽略）。详见已知坑 #84

#### 核心业务规则

- **字段提取**（`core.extract_customs_data`，横向 842×595 报关单，标签定位）：shipper=境内发货人左栏、consignee=境外收货人左栏、提单号=「提运单号」标签正下方同列值（`_value_by_label`，海运/铁路通用）、柜号=备注「集装箱标箱数及号码」上下文提取（`_container_numbers`，前缀 CIMU/WNGU/TIIU/TLLU 等，非参考实现的硬编码 `MATU`）、船名航次=「运输工具名称及航次号」标签紧邻下方同列值（`_value_below_label`，卡航填车架号如 `/91440112M0J494` 去开头 `/`），取不到回退备注「运输工具名称：XXX」、箱数/总重量=件数行的整数/小数、起运地/目的港=指运港行、起运日期=申报 8 位日期、品名=明细行「序号+10位商品编号+中文名」去重拼接
- **提单模板** 12 个 MERGEFIELD 邮件合并域（`shipper/consignee/提单号/柜号/船名航次/起运地/目的港/箱数/品名/总重量/总体积/起运日期`），`clean_template.dedupe_template` 先去重重复域（如「起运日期」出现 6 次、「目的港」4 次），再用 docx-mailmerge 填域保格式不漂移
- **保函程序重建**（`fill_telex_docx`）：TO/FROM/提单号/运输工具/**柜号**/目的地/收货人/申请单位/申请日期，逐行 `python-docx` 生成
- **收货人(保函) 与 consignee 独立**：提单 consignee 用「大写无空格」（`HONGKONGLIXIANG...`），保函收货人用「正常大小写带空格」（`Hong Kong Lixiang...`），审核表两列独立可编辑，`build_review_excel`/`parse_review_excel` 完整回读
- **拓锐固定 shipper/consignee（硬编码）**：该客户发货人/收货人固定不变，直接常量覆盖——`SHIPPER_BL`（3 行英文：`GUANGZHOU TUORUI TECHNOLOGY CO., LTD` + 2 行地址）、`CONSIGNEE_BL`（`HONG KONG LIXIANG TRADING COMPANY LIMITED`）；保函对应 `SHIPPER_TELEX`/`CONSIGNEE_TELEX`（正常大小写带空格）。来源=正确提单/保函，取值见 `core.py` 顶部常量
- **箱货清单（可选上传，参考箱单发票）**：`parse_packing_list(xlsx)` 定位「英文品名/中文品名/体积/货箱重量/FBA ID/客户渠道」表头（跳过表头行，避免表头字被当成品名），英文品名去重后覆盖提单「品名」（每行一个品名）、「体积」列求和覆盖「总体积(CBM)」（**无「体积」列时回退「长×宽×高×总箱数/1e6」**）、FBA ID 去重供物流追踪表匹配、「客户渠道」列取首个非空值供渠道→目的港映射；`/api/bl/extract` 额外接受 `packing=` xlsx，响应加 `warning`（解析失败为非阻断提示，仍返回底单提取结果）
- **周汇总箱货清单（可选，整周一份）**：用户按周导出的整份箱货清单（多票合并一个 xlsx），`parse_packing_list_weekly(xlsx)` 按「工作号」列分组（每组=国家/渠道/箱数/重量/品名/体积），`match_weekly_packing(rec, groups)` 用「运抵国 + 件数(箱数) + 毛重」回溯匹配每票底单对应的**一个或多个工作号**（一票合并报关常对应多工作号，凑箱数 + 验重量容差 1.0kg），命中则覆盖品名/总体积/渠道/箱数/国家。匹配失败回退文件夹内单票清单（`packing_{i}`）。`extract-batch` 额外接受 `weekly_packing=` xlsx
- **物流追踪表（可选上传）**：`parse_tracking_list(xlsx)` 遍历所有 sheet（2023/2024/2025/最新物流动态/10月），按表头文字定位「Shipment ID（FBA号）」+「ETD」+「ETA」+「船名航次/班列」列，合并成 `FBA ID → {etd, eta, vessel}` 索引。用箱货清单的 FBA ID 匹配后：`etd` 覆盖提单「起运日期」（ON BOARD 日期）；`eta` 经 `_arrival_date`（ETA−2天，ETA−ETD<2天取 ETA）算到港时间 → 电放保函「申请日期」；`vessel` 在底单船名航次为空时补入（铁路班列号）。匹配不到回退底单申报日期/今天。`/api/bl/extract` 额外接受 `tracking=` xlsx
- **港口英文（起运港按渠道大类+运输方式分流；目的港按渠道大类+运抵国）**：①`port_map.json`（结构化 `{origin: 离境口岸→英文, destination: 运抵国→英文}`）；②`customs_office_map.json`（非海运票关区名→起运港英文，如「增城海关」→`GUANGZHOU`）；③`channel_map.json`（非规则渠道名→目的港英文）。`_channel_category` 按客户渠道关键词分大类：铁路（含`铁路`/`快铁`/`铁派`）、卡航（含`卡航`）、快递（含`联邦`/`空派`）。`apply_port_map(rec, channel)` 分流——起运港：快递→固定 `SHENZHEN`，海运→离境口岸→origin，非海运→关区名→customs_office_map；目的港（按「运抵国」）：铁路→{英国:`MALASZEWICZE`,德国:`DUISBURG`}、卡航→{英国:`THE UK`,德国:`GERMANY`}、其他→destination。`apply_channel_map` 仅对非规则渠道用 channel_map 兜底（避免旧 channel_map 覆盖铁路/卡航规则）。generate 里 `remember_ports`（快递起运地固定不记忆）+ `remember_channels`（非规则渠道）回写 json（自动积累）
- **品名翻译兜底**：未上传箱货清单时，底单中文品名经 `PRODUCT_EN_MAP`（灯具词典）翻译成英文并大写换行；上传箱货清单则直接用其英文品名（权威，含 `Cilp` 等源数据原始拼写）。提单品名**保留空格**（`CILP TABLE LAMP` 原样，不连写——正确提单本身是带空格的，早期「连写去空格」是误解已回退）
- **扫描件 OCR 兜底**：无文字层的扫描底单 `extract_words` 返回空 → 走 RapidOCR（`_get_ocr` + `_ocr_page_to_words`）识别文字后复用同一套「标签定位」提取；扫描报关单多为竖版 594×843 而内容横排，OCR 前先 `np.rot90(arr, k=1)` 逆时针转正；OCR 失败才回退 `_empty_record()` 全空字段供人工填写
- **拆分底单合并**（`merge_declaration_pdfs`，pypdf）：按文件名关键词排序 `报(报关单)→放(放行单)→委托(委托书)` 后合并。⚠️ 上传保存时必须**保留原始文件名**（`batch_{i}_{j}_{stamp}_{原文件名}.pdf`），否则排序 key 丢失会误把委托书当第一页
- **命名规则**（`derive_output_name`）：把文件夹名里的「易通报关资料」替换成「提单/保函/底单」，其余严格照抄。例 `7月第1周-(易通报关资料）-BG20260703003 4票 英国快铁不包税-卡派 39箱 是8` → `7月第1周-(提单）-BG20260703003 4票 英国快铁不包税-卡派 39箱 是8`
- **箱数/国家一致性校验**（`validate_ticket`）：对比 报关底单（箱数/运抵国）vs 箱货清单（`总箱数(CTN)` 求和 / `国家` 列）vs 文件夹名（`N箱` / 国家），三者有值且互相不同则 warning「原底单/箱货清单有问题」。⚠️ 文件夹名里的「欧洲」是运输走廊（欧洲铁路）非目的国，已从国家表剔除避免误报
- **格式**：提单只产 PDF、保函只产 DOCX、底单只产 PDF；ZIP 压缩包命名 =「大文件夹名 + 系统制作文件」（如 `7月第2周系统制作文件.zip`，取 `webkitRelativePath` 第一段）；ZIP 内按票建**子文件夹**（子文件夹名 = 原上传文件夹名），每票三份文件按「底单.pdf → 提单.pdf → 保函.docx」归档（保函预览 PDF 放 `preview/` 子目录不进 ZIP）

#### API

- `POST /api/bl/extract`（`file=` PDF，可选 `packing=` xlsx、`tracking=` xlsx）→ `{ ok, record, bl_fields, bl_header, telex_fields, warning }`（单票，保留）
- `POST /api/bl/extract-batch`（`folder_{i}` / `pdf_{i}_{j}` / `packing_{i}` / `weekly_packing` / `tracking`）→ `{ ok, tickets:[{folder, record, warning}], bl_fields, bl_header, telex_fields }`（批量）
- `POST /api/bl/generate`（`{ tickets:[{folder, record}] }`）→ `{ ok, zip, previews:[{folder, bl, telex}] }`（相对路径，经 `/api/bl/file/<rel>` 访问）
- `GET /api/bl/file/<path:rel>` → 生成文件（ZIP 下载 / 提单·保函 PDF 预览），仅允许 `workspace/output` 内文件

#### 部署

- 前端经 `next.config.ts` `rewrites()` 把 `/api/bl/:path*` 代理到 Flask（`BL_SERVICE_URL`，默认 `http://localhost:5000`）
- Flask 用 pm2 常驻（`ecosystem.config.cjs` 的 `etton-bl-review`，原生 Python + 原生 LibreOffice 转 PDF，不依赖 Docker）

---

### 3.15 TR全量入仓数据整理 — `/warehouse-entry-full`

**文件**: `src/app/warehouse-entry-full/page.tsx` + `src/lib/warehouse-entry-full.ts` + `src/app/api/warehouse-entry-full/route.ts`（含 `/export` 子路由）

#### 与「TR入仓数据整理」的区别

原 `/warehouse-entry`（3.13）是**代表箱选数**（每个 FBA 号选一个代表箱 + 历史库对比），本功能是**逐箱输出**：
- 一个 FBA 号下每个供应商箱各生成一行（`totalBoxes = 1`），不再做代表箱选数
- **无历史库**：不读条数、不累积、无历史库 UI、无「建议参考历史最大值放大」报警
- 出给客户尺寸：材积主导（材积重 > 实重）→ 11 级尺寸修正规则放大；实重主导 → 取该箱原值（长宽高降序）

#### 功能

1. 上传客户数据 + 供应商数据，逐箱匹配、自动放大尺寸、校验报警，生成「出给客户」建议箱规
2. 页面**不展示箱规表格**，整理完成后直接展示汇总 + 全局校验条 + 导出按钮
3. 导出《内部三类数据_<日期>_合计总箱数<N>.xlsx》（复用 `exportOutputBuffer`，含全局兜底校验）

#### 核心业务规则

- **11 级尺寸修正规则**（材积主导时）：三边降序为 `[最长, 次长, 短]`，依次尝试规则 1–10（三边各+1 / 短边+2 / 次长+短各+1 / 最长+短各+1 / 次长+2 / 最长+次长各+1 / 最长+2 / 短+1 / 次长+1 / 最长+1），命中第一条「不冲突」即返回；冲突 = 修正后三边和 vs 客户三边和 差 ≥ 6，或 修正后材积重 vs 客户材积重 差 ≥ 2；规则 11「原值」兜底（前面全冲突时不放大）
- **差异校验**（该供应商箱 vs 客户申报）：材积主导 → 三边和差 ≥ 6、材积重差 ≥ 2「核查过机图」；实重主导 → 实重差 ≥ 0.5「核查过机图」
- **供应商过大箱**：该 FBA 号最大材积重 − 当前箱材积重 ≥ 2 → 「核查过机图」
- **成本重校验**：Σ出给客户总计费重 < Σ供应商总计费重 → 「成本重过大，请和供应商申请」（导出前全局兜底等比例放大并标 `[全局调整]`）

#### API

- `POST /api/warehouse-entry-full`：上传两文件 → `{ rows, supplierTotal, summary }`（无历史库字段）
- `POST /api/warehouse-entry-full/export`：body `{ rows, supplierTotal }` → 下载 Excel（无历史库累积）

---

## 4. 非目标（明确没做的）

- ❌ **用户认证/登录**: LAN 工具，无权限控制
- ❌ **持久化存储**: 拆分结果仅在内存中保存 30 分钟
- ❌ **Claude AI 集成**: `@anthropic-ai/sdk` 已安装但未在页面/API 中使用（仅占位依赖）
- ❌ **多 Sheet 支持**: 保单拆分仅处理 `"ETTON电商物流 下单模板"` Sheet；多供应商对账自动选择数据最密集的 Sheet
- ❌ **多币种对账**: 当前对账引擎优先匹配本位币(RMB)列，但不自动做币种转换
- ❌ **Excel 输出自定义**: 太平洋转换输出为全新 workbook（不保留源格式），保单拆分保留克隆格式
- ❌ **数据库**: 无任何数据库依赖
- ❌ **i18n 国际化**: 仅中文
- ❌ **PWA / 离线支持**: 无 Service Worker
- ❌ **自动化测试**: 无单元测试 / E2E 测试
- ✅ **OCR 扫描件识别**（2026-09 已实现）：扫描底单（无文字层）用 RapidOCR（rapidocr_onnxruntime，中文识别准、比 tesseract 稳）识别，复用标签定位提取（提单号/船名航次/起运地/目的港/运抵国/箱数/品名/总重量/起运日期均可提取）
- ❌ **英文自动翻译**: 提单/保函的英文字段（shipper 英文名+地址、consignee 英文、起运地/目的港英文、品名翻译、总体积、实际起运日期）底单里没有，靠人工在审核表补填，不做自动翻译/港口映射
  - ⚠️ 已部分放开：拓锐客户 shipper/consignee 英文为**固定常量硬编码**；品名英文翻译 + 总体积改由**箱货清单 xlsx** 自动汇总（`parse_packing_list`）。仍留人工的只有：起运地/目的港**英文**（底单是中文「阿拉山口铁路/英国」，提单要英文 `CHONGQING`/`MALASZEWICZE`）+ 实际起运日期（提单要实际开船/发车日，底单只有申报日期）。船名航次本身**在底单有**（「运输工具名称及航次号」列，海运为船名航次、非海运该列为 `@` 占位符→空）

---

## 5. 已知坑 / 绕过的 Hack / 待重构项

### 已知坑

1. **exceljs 不支持 `=DISPIMG()` 公式**（内嵌图片）
   - **绕过**: `injectCellImagesIntoOutput()` 手动从原始 xlsx zip 中提取 `cellImages.xml` + `xl/media/`，重新注入到 exceljs 生成的 xlsx zip 中
   - **影响范围**: 保单拆分输出文件中的内嵌图片保留
   - **风险**: 依赖 xlsx 内部 zip 结构，Office 版本升级可能导致路径变化

2. **exceljs Buffer 类型兼容**
   - `wb.xlsx.load(buffer)` 的类型签名与 `@types/node` 的 `Buffer` 不完全兼容
   - **绕过**: 使用 `// @ts-expect-error` 注释抑制类型错误
   - **位置**: `split-insurance.ts` 和 `multi-supplier-reconciliation.ts`

3. **ESLint 导致 Docker 构建失败** (2026-07-21)
   - `next build` 包含 lint 检查，`prefer-const` 和 `no-unused-vars` 错误会阻断构建
   - **教训**: 每次修改后必须运行 `npm run build` 验证，不能仅依赖 `tsc --noEmit`
   - **影响**: 未检查的代码推送到 GitHub 后 Docker 构建会静默失败（42 秒内报错）

4. **供应商账单表头行位置不固定** (2026-07-21 修复)
   - 不同供应商甚至同一供应商不同时期的账单，表头行位置可能不同
   - **修复**: `parseSupplierBill()` 改为扫描前 20 行自动检测
   - **位置**: `multi-supplier-reconciliation.ts` line 395-440

5. **请款明细金额列币种混用** (2026-07-21 修复)
   - 请款明细中"金额"列可能是原币(USD)或本位币(RMB)，直接用"金额"会导致对账误差
   - **修复**: 评分机制优先匹配"金额(本位币)"（关键词更长=得分更高）
   - **位置**: `multi-supplier-reconciliation.ts` `parsePaymentFile()` 列检测逻辑

6. **共享公式 (Shared Formula) 展平**
   - exceljs 的 `spliceRows` 删除行时会破坏共享公式引用链
   - **绕过**: `flattenFormulas()` 在删除行之前将共享公式转为独立公式（保留 result）
   - **位置**: `split-insurance.ts` line 269-293

7. **内存 session 存储**
   - 拆分结果（含 Buffer）存储在 Node.js 进程内存 Map 中
   - **风险**: 大文件多用户并发可能 OOM；进程重启丢失所有 session
   - **缓解**: 30 分钟自动过期清除

8. **tailwindcss 版本锁定**
   - 项目使用 Tailwind CSS 4，语法与 v3 完全不同（无 `tailwind.config.ts`）
   - @tailwindcss/postcss 插件是必需的，缺失会导致样式完全丢失

9. **SEALOS Deployment 易丢失** (2026-07-21)
   - SEALOS 上 Deployment 可能因资源回收或平台升级被删除，但 Service 和 Ingress 会保留
   - **症状**: 网站不可访问，`kubectl get deploy` 找不到 etton-tools
   - **修复命令**: `kubectl apply -f k8s/deploy-sealos.yaml`
   - **恢复脚本**: 项目根目录 `redeploy.sh`

10. **延讯发票 sheet 名带尾随空格** (2026-08-28)
    - 延讯发货单 sheet 名为 `"发货单 "`（含尾随空格），直接按名称匹配会失败
    - **绕过**: 延讯转换固定用 `wb.worksheets[0]` 定位发货单，不依赖 sheet 名
    - **位置**: `yanxun-convert.ts` `convertYanxunToEtton()`

11. **易通模板数据区示例数据需先删除** (2026-08-28)
    - 易通模版 R25 起预置 5 行示例数据（Night light），需 `spliceRows` 删除后写入真实数据
    - **注意**: 删除后新数据行无边框，需统一补 `thin` 边框（24 列）；删除行数须用 `actualRowCount`（见坑 #22），不能用 `rowCount`
    - **位置**: `yanxun-convert.ts` 数据区填充逻辑

12. **延讯「带电」字段为否时不填** (2026-08-28)
    - 需求规则「带电（不带电就不填）」，易通模版 F2 下拉为「是,否」，默认留空
    - 仅当延讯带电值为「是」或「带电」时才写「是」
    - **报关方式映射**: 公司自报→普通报关、永德吉报关→报关退税（`CUSTOMS_MAP`，可扩展）

13. **批量 ZIP 打包重名处理** (2026-08-28)
    - 多票导入文件可能同名（如不同目录下同名文件），ZIP 内同名文件会覆盖
    - **绕过**: `generateYanxunZip()` 用 `nameCount` Map 计数，重名自动加 `_2`/`_3` 后缀（`原名_ETTON.xlsx` → `原名_ETTON_2.xlsx`）
    - **位置**: `yanxun-convert.ts` `generateYanxunZip()`

14. **延讯渠道缺失会误填发件人地址** (2026-08-28)
    - 延讯发票未写渠道时，「物流商/渠道」标签右侧为空，旧的 fallback/大范围 `valueRight` 会扫到「发件人地址」标签，把地址误填进易通「备注」
    - **修复**: 渠道只取「物流商/渠道」标签右侧 1~2 格；扫不到即报「无渠道名」，并用 `/发件人|地址|邮编/` 正则兜底过滤
    - **位置**: `yanxun-convert.ts` `parseTopInfo()` + `validateTopInfo()`

15. **混箱后续行须补 0 而非空白** (2026-08-28)
    - 易通标准答案中混箱后续行的 Number/净重/毛重/长/宽/高 均为 `0`，而非空值；空值会导致下游导入报错
    - **修复**: 数据区填充时混箱后续行统一写 `0`
    - **位置**: `yanxun-convert.ts` 数据区填充逻辑

16. **预计总体积须用「体积CBM」列** (2026-08-28)
    - 直接用 长×宽×高/1000000 会得到 0.4566，而标准答案为 0.45（各箱「体积CBM」列 ROUND 到 2 位后求和）
    - **修复**: 新增 `volumeCbm` 列映射（表头关键词「体积」），预计总体积 = 各箱「体积CBM」合计
    - **位置**: `yanxun-convert.ts` `YANXUN_HEADER_PATTERNS` + 预计总体积计算

17. **场景判断依据为「FBA号/海外仓」字段，海外仓须自动填「私人地址/海外仓」** (2026-08-28)
    - 易通「仓点类型」下拉为 `FBA,Walmart` 二选一，且「私人地址/海外仓」与 FBA 地址库组二选一
    - **规则**: 场景判断依据是「FBA号/海外仓」字段值是否含「海外仓」字眼（不是「目的地」字段）——含 → 海外仓场景（清空 FBA 地址库组 R7~R12，从「目的地」地址文本解析收件人姓名/公司/地址/电话/城市/州/邮编并填到「私人地址/海外仓」R14~R22），否则 → FBA 场景（填 FBA + 校验 FBA号）
    - **位置**: `yanxun-convert.ts` `parseTopInfo()`（`warehouseType` 字段 + `parseOverseasAddress()`） + `convertYanxunToEtton()`

18. **输出文件名改为「ETTON_FBA号 / ETTON_调拨单号」命名** (2026-08-28)
    - 海外仓场景没有 FBA号，须用「调拨单号」命名；FBA 场景用 FBA号 命名，两者统一加 `ETTON_` 前缀
    - **修复**: 文件名 = FBA 场景 `ETTON_FBA号.xlsx`（如 `ETTON_FBA19MX7M8KR.xlsx`），海外仓场景 `ETTON_调拨单号.xlsx`（如 `ETTON_TF2608270070.xlsx`）；`sanitizeFileName` 清洗非法字符，缺失时 fallback「未命名」
    - **位置**: `yanxun-convert.ts` `convertYanxunToEtton()` 文件名生成

19. **`.next` 编译缓存损坏导致核心 JS chunk 404、React 交互全失效** (2026-09-01)
    - 症状：页面能正常显示（SSR HTML），`<label>` 原生触发文件框也正常，但所有 React 合成事件（onChange/onDragOver/onDrop/onClick）都失效——文件名不显示、拖拽边框不变绿、状态不更新
    - 根因：`polyfills.js` / `main-app.js` / `app-pages-internals.js` 等核心 chunk 返回 404，React 在浏览器根本没加载
    - 修复：停止 dev server → `rm -rf .next` → 重新 `npm run dev`，重新编译后 chunk 全部 200
    - 教训：排查「原生 HTML 行为正常、React 交互全失效」时，先 curl SSR HTML 里的 script 引用确认 JS chunk 是否 404，不要只在业务代码里找 bug

20. **TR入仓历史库无持久卷，pod 重启即丢失** (2026-09-01)
    - `data/history.json` 写入容器文件系统（`runAsNonRoot: true`、无 PVC 挂载），K8s pod 重启/重建后历史库清零
    - **缓解**: 前端提供「导入/导出历史库」按钮，重要历史数据需手动导出备份（.xlsx），重启后重新导入；2026-09-03 起 `loadHistory()` 在 `history.json` 不存在时回退到内置 66 条种子数据（见 #41）
    - **位置**: `warehouse-entry.ts` `HISTORY_FILE()`（`process.cwd()/data/history.json`）；`data/` 已加入 `.gitignore`

21. **TR入仓历史对比方向按业务语义实现为「历史最大 > 建议值」才提示放大** (2026-09-01)
    - 原始规格文字写「历史最大计费重 < 建议值 → 报警放大」，与规则 6「选了最大还比历史小则放大」语义矛盾
    - **实现**: 采用「历史最大计费重 **>** 建议值 → 提示『建议参考历史最大值放大』」（建议值偏小时才有放大空间），与规则 6 一致
    - **位置**: `warehouse-entry.ts` `buildSuggestions()` 历史对比分支

22. **延讯转换：`spliceRows` 删模板示例行须用 `actualRowCount` 而非 `rowCount`** (2026-09-02 修复)
    - 症状：转换 FBA19NFKCR3S（源文件仅 3 箱）后，输出文件底部多出 2 行 `FBA19MXX2JCWU000004/005`（Night light）——是模板 R28/R29 的示例数据没删干净
    - 根因：exceljs 的 `rowCount` 把模板带格式的空行也算进去（本例 =53），而实际有数据的行是 `actualRowCount`=29。旧代码 `rowsToDelete = rowCount - 25 + 1 = 29`，使 `spliceRows` 内部 `nKeep = start + count = 54 > 实际行数`，删除循环一次都不执行（静默失败，不报错）
    - 修复：`rowsToDelete = outSheet.actualRowCount - dataStartRow + 1`（=5），只删 R25~R29 这 5 行示例数据
    - 位置：`yanxun-convert.ts` `convertYanxunToEtton()` 数据区删除逻辑

23. **延讯「币种」列常是公式且查不到值，`cellText` 会把错误对象转成 `[object Object]`** (2026-09-02 修复)
    - 症状：加拿大 FBA 发货的币种列 C33 是 `VLOOKUP(...,英欧链接价格!H:Q,9,0)`，查不到返回 `result: { error: "#N/A" }`；旧 `cellText` 对 `result` 直接 `String()` → `[object Object]`，导致输出币种列变成 `[object Object]`，`|| "USD"` 回退失效
    - 修复：`cellText` 对公式 `result` 分支只接受 string/number/boolean，错误对象（`{error:"#N/A"}` 等）返回空字符串 → 币种 `"" || "USD"` 正确回退 `USD`（美加等美元区默认）
    - 位置：`yanxun-convert.ts` `cellText()`

24. **TR入仓供应商「英美入仓」格式的 FBA 在「扩展箱号」列，不是「货箱编号」列** (2026-09-02)
    - 症状：若沿用天图格式的 `fbaId: ["货箱编号"]`，会误匹配到英美入仓格式里的「货箱编号」列（该列实为运单号 `10593316U001`，非 FBA），导致 FBA 提取错误
    - 根因：英美入仓格式（如 `TRKJ26080105-英美入仓数据.xlsx`）中「货箱编号」= 运单号，真正的逐箱 FBA 在「扩展箱号」列（`FBA19MYJ057TU000001` → 前 12 位）；箱规在 货箱重量/长度/宽度/高度/材积重（BI~BM）
    - 修复：新增 `SUPPLIER_ENTRY_PATTERNS`（`fbaId: ["扩展箱号"]` 等精确列名），`parseSupplierFile()` 通过「表头是否含『扩展箱号』」自动切换两套模式
    - 位置：`warehouse-entry.ts` `SUPPLIER_ENTRY_PATTERNS` + `parseSupplierFile()` 格式检测分支

25. **TR入仓客户「易通发票」格式的重量在「货箱重量」列，且「品名」会误命中「英文品名」** (2026-09-02)
    - 症状：若沿用标准 `CUSTOMER_PATTERNS`，易通发票（`8月第4周（易通发票）...xlsx`）的实重列读成 0（标准模式只认「实重」），且品名取到「英文品名」而非「中文品名」
    - 根因：易通发票是「表单头 + 明细表」结构，明细表中文表头为 `FBA货箱编号/英文品名/中文品名/箱数件数/货箱重量/长宽高`；「货箱重量」= 实重，「品名」兜底会先命中靠前的「英文品名」列
    - 修复：新增 `CUSTOMER_INVOICE_PATTERNS`（`fbaId:["FBA货箱编号"]`、`productName:["中文品名"]`、`actualWeight:["货箱重量"]` 等精确列名），`parseCustomerFile()` 通过「表头是否含『FBA货箱编号』」自动切换；无「材积重」列时按公式重算
    - 位置：`warehouse-entry.ts` `CUSTOMER_INVOICE_PATTERNS` + `parseCustomerFile()` 格式检测分支

26. **TR入仓导出格式对齐参考文件《拓锐…入仓数据（成本）》——三组并排 + 公式列** (2026-09-02)
    - 背景：用户要求导出「按照参考文件格式输出，方便人工参考调整」，参考文件为 40 列、表头两行（分组 `客户的(J1:N1)` / `供应商(Q1:W1)` / `出给客户(Z1:AD1)` + 列名）、三组数据并排
    - 关键点：派生列（材积重/总实重/总材积重/计费重/差异）写成 **Excel 公式**（如 `=J3*K3*L3/6000`、`=ROUND(MAX(AE3,AF3),0)`），这样人工改动尺寸后能自动重算；而非写死数值
    - 无来源列留空：出货日期/成本KG/渠道/总成本重 是人工业务字段，工具不产生，导出时空列占位供人工填写（2026-09-03 起 系统SO/客户渠道/国家/仓库代码/单证报关 已改为从客户数据自动取数，见 #33）；备注列(AJ，无表头)放报警文案
    - 新增字段：`SuggestionRow.supplier`（选数命中的供应商代表箱原始值，未匹配时全 0），供「供应商」对比列输出；`buildSuggestions()` 三处 push 均需填充该字段，漏填会导致导出「供应商」列空白
    - 位置：`warehouse-entry.ts` `exportOutputBuffer()`（列位常量 `C` + `colLetter` 公式拼接）+ `SupplierRepresentative` 接口

27. **TR入仓「实重主导」单行校验改用 0.4 上限，取代 ×1.02 放大** (2026-09-02)
    - 规则：当 `实重 > 材积重`（计费重由实重主导）时，建议实重 = `max(供应商最大实重, 历史最大实重)`，上限 `≤ 客户实重 + 0.4`（`ACTUAL_CAP_TOLERANCE`）；材积主导仍保留 `×1.02` 放大
    - 判定口径：`suggestion.actualWeight > suggestion.volumeWeight`（选数命中箱规的最大实重 > 公式材积重）即视为实重主导，与 `forceAmplify()` 内部 `volumeWeight >= actualWeight` 分支互补
    - 坑：客户实重为 0（缺失）时上限无意义 → 用 `Infinity` 跳过封顶；历史库该品名用 `actualWeight`（非 chargeableWeight）参与取 max，避免材积主导历史把实重虚高
    - 位置：`warehouse-entry.ts` `buildSuggestions()` 单行校验分支 + `ACTUAL_CAP_TOLERANCE` 常量

28. **TR入仓历史库「同款」按客户箱规/实重相近判定，不再按品名去重** (2026-09-02)
    - 规则：同名产品（如「水波纹灯」）可能对应多款不同规格，仅品名一致不足以判定同款。改为：品名一致 且 客户长宽高**排序后**逐边差 ≤ 1cm 且 实重差 ≤ 1kg → 同款；否则视为不同产品，不参与历史对比。排序是为了让「同一款箱的长宽高书写顺序不同」（如 `39.6×39.6×43.8` vs `43.8×39.6×39.6`）不被误判成两款
    - 数据：`HistoryEntry` 拆成「客户申报（`customerLengthCm/WidthCm/HeightCm/customerActualWeight`，用于同款判定）」+「出给客户建议（`lengthCm/.../chargeableWeight`，用于历史最大实重/计费重对比）」两组字段；历史库由 `Record<品名, ...>` 改为扁平数组 `HistoryLibrary = HistoryEntry[]`
    - 匹配/去重统一走 `isSameHistoryProduct()` + `upsertHistoryEntry()`（同款取计费重更大者）；`buildSuggestions()` 用 `history.filter(isSameHistoryProduct)` 取同款历史，分别取 max 实重（实重主导用）与 max 计费重（历史对比报警用）
    - 导入：`importHistoryFromExcel()` 用 `pickGroupCols()` 区分「客户组 / 出给客户组」列（参考文件两列同名「长(CM)」，客户取首次、出给客户取末次）；`loadHistory()` 兼容旧 `{品名:...}` 格式自动迁移为数组
    - 位置：`warehouse-entry.ts` `HistoryEntry`/`HistoryLibrary` + `isSameHistoryProduct`/`historyIdentity`/`upsertHistoryEntry` + `importHistoryFromExcel`/`loadHistory`

29. **TR入仓客户「货箱清单」格式：实重是「单箱货物毛重」而非「货箱重量」** (2026-09-02)
    - 症状：新客户文件 `客户数据-给英美-0824到0830.xlsx` 的「货箱重量」列（如 2442.45）是整张 SO 的**总重**，若误取为单箱实重会导致计费重虚高数十倍
    - 根因：此格式与「易通发票」不同——易通发票「货箱重量」= 单箱实重，而货箱清单「货箱重量」= SO 总重；真正的单箱实重在「单箱货物毛重(KG)」列（另有「单箱货物净重(KG)」列，毛重≠净重，须取毛重）。同表同时有「英文品名」「中文品名」两列，「品名」兜底会误命中「英文品名」
    - 修复：新增 `CUSTOMER_CARGO_PATTERNS`（`productName:["中文品名"]`、`actualWeight:["单箱货物毛重"]`、`totalBoxes:["总箱数(CTN)"]` 等），`parseCustomerFile()` 通过「表头是否含『单箱货物毛重』」自动切换（优先级在发票格式之后、标准格式之前）。无「材积重」列，按公式重算
    - 位置：`warehouse-entry.ts` `CUSTOMER_CARGO_PATTERNS` + `parseCustomerFile()` 格式检测分支

30. **TR入仓「出给客户」改为沿用供应商选数箱规，取消单行强制放大** (2026-09-02)
    - 规则：出给客户 = 供应商选数箱规（尺寸 + 该箱规最大实重 + 公式材积重），不再按「实重/材积主导」分支强制放大；同时撤销 #27 的「实重主导 0.4 上限」与「材积主导 ×1.02 放大」
    - 差异约束（对比客户申报）：三边和差必须 `< 6`（`>= 6` 报警）、材积重差必须 `< 2`（`>= 2` 报警）；超限报警文案改为「请核查过机图」（供应商过大箱同理），对应「供应商数据明显偏大 → 标红要求找供应商核查过机图，核实后再修改」
    - 参考优先级：历史有数据用历史参考、有供应商数据用供应商数据参考（供应商实测为准），历史仅作「历史最大计费重 > 建议值」报警用，不再参与建议实重取 max
    - 连带删除：`ACTUAL_CAP_TOLERANCE` 常量、`historyMaxActual` 局部变量（实重主导分支专用）；`forceAmplify()` 仅剩 `exportOutputBuffer()` 全局兜底引用
    - 位置：`warehouse-entry.ts` `buildSuggestions()`（删除单行校验分支）+ 报警文案

31. **TR入仓「出给客户」恢复最短边放大：供应商选数箱规 + 最短边 +1** (2026-09-03)
    - 背景：实测答案文件《正确答案-给英美-0824到0830.xlsx》里 53/105 例对供应商选数箱做了「单边 +1」放大（多为最短边，如 `54×53×36`→`54×53×37`），并约束「材积重差 < 2」
    - 规则：出给客户 = 供应商选数箱规 + 放大**最短边 +1**（尽量不放大最大边，取较短两边中更短那条）；实重 = 该箱规所有箱最大实重；材积重 = 公式重算
    - 放大约束：放大后材积重 − 客户材积重 必须 `< 2`、计费重不超过历史最大计费重（历史最大是上限）；任一不满足 → 不放大（退回供应商原尺寸）。例：供应商 `54×53×36`、客户 `51.5×51.5×36` → 最多 `54×53×37`，再大材积重差 ≥ 2
    - 实测匹配率：尺寸精确匹配 42/105（不放大）→ **48/105（最短边 +1）**；其余为人工手改噪声（同一供应商箱有时加宽、有时加高），无法确定性 100% 复现
    - 位置：`warehouse-entry.ts` 新增 `amplifyDims()`（最短边 +1 + 两条约束），`buildSuggestions()` 供应商代表箱保留原始值、`suggestion` 用放大值

32. **TR入仓「查过机图」报警改为按供应商原始箱规 + 区分材积/实重主导** (2026-09-03)
    - 背景：原报警用「出给客户建议值（已放大）」对比客户；现改为用「供应商**原始箱规**」（选数命中的代表箱，未放大值）对比客户
    - 规则：**材积主导**（供应商体积重 ≥ 实重）→ 三边和差 ≥ 6、材积重差 ≥ 2 均「核查过机图」；**实重主导**（供应商体积重 < 实重）→ 实重差 ≥ 0.5「核查过机图」
    - 新增常量 `ACTUAL_DIFF_THRESHOLD = 0.5`；材积重差文案由「材积重差异超限」补为「材积重差异超限，请核查过机图」
    - 前端 `recomputeAlarms()` 同步该口径（编辑建议值后实时重算报警）
    - 位置：`warehouse-entry.ts` `buildSuggestions()` 报警块 + `page.tsx` `recomputeAlarms()`

33. **TR入仓导出 B–F 列（系统SO/客户渠道/国家/仓库代码/单证报关）从客户数据自动取数** (2026-09-03)
    - 背景：用户反馈导出的「出给客户」Excel 里 系统SO/客户渠道/国家/仓库代码/单证报关 五列空白，实际数据源在客户数据文件里（货箱清单格式的 系统SO(第1列)/国家(第3列)/仓库代码(第4列)/单证报关(第5列)/客户渠道(第27列)）
    - 修复：`CustomerRow`/`SuggestionRow` 各新增 `so/channel/country/warehouse/customs` 五字段；三个 `CUSTOMER_*_PATTERNS` 增加对应列名匹配（`系统SO`/`客户渠道`/`国家`/`仓库代码`/`单证报关`）；`parseCustomerFile` 读取、`buildSuggestions` 透传（未匹配分支与正常分支两处 push 都需填，漏填导致对应行空白）、`exportOutputBuffer` 写入 B/C/D/E/F 列。出货日期(A) 客户数据无此列，仍留空
    - 前端无需改动：`editSuggestion` 用 `{...r, suggestion}` 整体保留行对象、`handleExport` 整体 `JSON.stringify({rows})` 回传，新字段自动透传
    - 位置：`warehouse-entry.ts` `CustomerRow`/`SuggestionRow`/`CUSTOMER_*_PATTERNS`/`parseCustomerFile`/`buildSuggestions`/`exportOutputBuffer`

34. **TR入仓选数改为按历史参考值决定第 1/第 2 大** (2026-09-03)
    - 规则：选数不再固定取计费重第 1 大——有历史同款时：供应商第 1 大计费重 ≤ 历史最大计费重 → 取第 1 大（并保留「历史最大 > 建议值 → 建议参考历史最大值放大」报警）；第 1 大超过历史最大 → 退取第 2 大；新品（无历史）→ 直接取第 2 大（避免取到偶发偏大的异常箱）；仅 1 箱回退第 1 大
    - 实现：`selectBox(sorted, historyMaxChargeable)` 增加历史参数；`buildSuggestions` 调用时传 `historyMax?.chargeableWeight ?? null`。`supplierChargeable` 仍保持 `sorted[0].chargeableWeight`（供应商最大计费重，供前端与「历史最大」列对比）；`pickedRank` 自动反映选中的是 1 还是 2
    - 位置：`warehouse-entry.ts` `selectBox()` + `buildSuggestions()` 选数处

35. **TR入仓导出格式精确对齐参考文件（字体/颜色/列宽/边框）** (2026-09-03)
    - 背景：用户要求导出 Excel 与参考文件《TRKJ26080099和TRKJ26080100全部数据.xlsx》格式一致（含字体、颜色）
    - 实现（`exportOutputBuffer`）：
      - 字体：全表宋体 11、黑色；第 2 行列名表头加粗，分组表头/数据/合计不加粗
      - 分组表头三色：客户的 `FFDEEBF7`（浅蓝）、供应商 `FFFBE5D6`（浅橙）、出给客户 `FFE2F0D9`（浅绿）
      - 列名表头分类：主体 `FFADB9CA`（灰蓝加粗）、分隔列 O/X `FFE2F0D9`、尾列 AK-AN `FF5B9BD5`（蓝不加粗）、备注 AJ 无填充无边框
      - 边框：数据区细黑边框（`FF000000`）；合计行与 AJ 列无边框；`showGridLines=false`（隐藏网格线）
      - 列宽 40 列精确复刻；numFmt：材积重/总材积重/差异 `0.00_);[Red](0.00)`（负数红）、总实重 `0.00_`、计费重 `0_`
      - 合计行无边框无填充、宋体 11，仅填总箱数/计费重/总计费重/箱数/总计费重 SUM 公式
    - 坑：统一字体循环若从 R2 起会覆盖表头加粗，须从 R3 起；分组表头合并后非左上角单元格也要逐个设 fill 才整片着色
    - 位置：`warehouse-entry.ts` `exportOutputBuffer()` 样式常量 + 分组/表头/数据/合计/列宽/边框块

36. **TR入仓导出按 SO 合并 + 渠道汇总行 + 文件名含合计箱数** (2026-09-03)
    - 背景：AH 总计费重/AI 成本KG 按 SO 合并；AK-AN 列按渠道汇总；文件名带合计总箱数
    - 实现（`exportOutputBuffer` + `export/route.ts` + `page.tsx`）：
      - 明细行先按 渠道→SO 稳定排序（`channelOrder`/`soOrder` Map 记录首次出现顺序），确保同 SO、同渠道相邻
      - SO 级合并：B(系统SO)/F(单证报关)/AH(总计费重)/AI(成本KG) 用 `mergeCells`；B/F 仅 SO 首行填值；AH 首行填 `SUM(AG范围)`（=该 SO 各产品计费重之和）；AI 留空供人工回填（可与供应商砍价）
      - 渠道汇总：AK-AN 列从第 3 行起连续填每渠道一行（AK=渠道名、AL=`SUM(I范围)`、AM=`SUM(AH范围)`、AN=`SUM(AI范围)`），末行「合计」`SUM` 各渠道汇总；A-AJ 列明细从 R3 起不受影响
      - 总合计行：I/AG/AH/AI 的 `SUM`（范围仅明细区）
      - 文件名：`内部三类数据_<日期>_合计总箱数<N>.xlsx`，N=Σ totalBoxes（前端 `a.download` 与服务端 Content-Disposition 都要同步改，浏览器以 a.download 为准）
    - 坑：参考文件 AH 是 SO 级 `SUM(AG)` 而非逐行 `=AG`（原实现错误）；明细行 AK-AN 应为空、只在渠道汇总行填（原实现把 AL/AM 抄到明细行，与参考文件不符）
    - 位置：`warehouse-entry.ts` `exportOutputBuffer()` 数据/合并/渠道汇总/合计块；`export/route.ts` 与 `page.tsx` 文件名

37. **TR入仓「供应商小于客户」时出给客户取历史最大值 + 单独提示** (2026-09-03)
    - 背景：供应商实测（选数命中箱）计费重 < 客户申报计费重时，若仍以供应商偏小值出给客户会亏运费，需改参考历史合理值
    - 规则：`supplierPickedChargeable = max(供应商代表箱实重, 材积重) < 客户计费重` → 报警「供应商小于客户，请确认」；若有历史同款（`historyMax`）→ 出给客户建议值整体取历史最大值（长/宽/高/实重/材积重/计费重/三边和），避免出给客户比客户还小；无历史（新品）→ 仅提示，保持供应商选数+放大值
    - 顺序：放在「供应商过大箱」报警之后、「历史对比」报警之前——取历史最大值后 `historyMax.chargeableWeight > suggestion.chargeableWeight` 恒为 false（相等），故「建议参考历史最大值放大」不会误触发
    - 位置：`warehouse-entry.ts` `buildSuggestions()` 供应商过大箱报警与历史对比之间

38. **TR入仓供应商新增「给总部」格式（逐箱，FBA号列 + 单件重量）** (2026-09-03)
    - 背景：新增「供应商数据-给总部-0824到0830 更新.xlsx」格式，与天图/英美入仓同构但列名不同，且是「多 SO 块堆叠」（每个 SO 块：块头 4 行 + 表头 + 逐箱数据 + TOTAL）
    - 格式：FBA 在「FBA号」列（`FBA19MTJH5NPU000014` → 取「U+流水号」前）；箱规在 长(CM)/宽(CM)/高(CM)；单箱实重=「单件重量（KGS)」、材积重=「单件材积(KGS)」
    - 实现：新增 `SUPPLIER_HEADQUARTERS_PATTERNS`；`parseSupplierFile` 的 `findHeaderRow` 增加 `["FBA号","单件重量"]` 兜底（原 `["货箱编号","货箱长"]` 匹配不到此格式）；格式检测用「单件重量」→ 给总部
    - 坑：多 SO 块堆叠，数据行之间夹着块头行（「目的港/仓库代码」等文本）和 TOTAL 行，`readRowText(map.fbaId)` 会读到这些非空文本；须在读取时过滤「FBA 提取后非 `FBA+字母数字`」的行（`/^FBA[A-Z0-9]+$/i`），否则会把表头「FBA号」/TOTAL「计费重:xxx」/块头「目的港:加拿大」误当成箱（曾误读 2 箱 FBA=「FBA号」）
    - 坑：文件带第二个 sheet「TRKJ26080107」（某 SO 的重复明细），仅读 `worksheets[0]` 避免重复计数
    - 位置：`warehouse-entry.ts` `SUPPLIER_HEADQUARTERS_PATTERNS` + `parseSupplierFile()` 表头检测/格式检测/数据过滤

39. **TR入仓历史库导入误取「总」汇总列（总实重/总材积重）** (2026-09-03)
    - 背景：历史库从 40 列「内部三类数据」格式导入（`importHistoryFromExcel`），该格式含「总实重/总材积重」汇总列（值 = 单箱 × 总箱数），而 `pickGroupCols` 用 `t.includes(keyword)` 匹配关键字「实重/材积重」，导致「总实重」「总材积重」也被命中
    - 坑：`pickGroupCols` 兜底取 `hits[hits.length - 1]`（最后一列）时，选中的是「总实重」(列 31) 而非单箱「实重」(列 29)，导致「出给客户」建议重量被读成总重量（曾出现 chargeableWeight = 590.85 / 751.67 等异常大值）
    - 修复：匹配关键字时增加 `!t.startsWith("总")` 过滤，排除「总实重/总材积重/总计费重」等汇总列，只取单箱值；修复后 chargeableWeight 恢复正常（如 15.15 / 17.43，无 > 100 的异常值）
    - 位置：`warehouse-entry.ts` `importHistoryFromExcel()` 的 `pickGroupCols`

40. **TR入仓历史库手动导入最终数据改为「覆盖」而非「取更大者」** (2026-09-03)
    - 背景：导出后自动累积的历史库用的是「出给客户建议值」（可能因放大而偏大），但最终提供给客户的值往往是我们手工调整后的值，两者未必一致
    - 做法：新增 `upsertHistoryEntry(lib, entry, { overwrite })` 参数——自动累积 `accumulateHistory` 仍「取计费重更大者」（不把历史最大值压小）；手动导入（`/api/warehouse-entry/history`）传 `{ overwrite: true }`，同款直接覆盖为最终值（即便更小也以最终为准）
    - 坑：若沿用「取更大者」，最终值比自动累积建议值小时会被忽略，导致历史库始终记住偏大的建议值
    - 同时支持导入**单组格式**《…出给客户的.xlsx》（仅一组 长/宽/高/实重/材积重，无「客户/出给客户」分组）——`pickGroupCols` 对单组列 `hits[0]` 与 `hits[last]` 指向同一列，客户标识回填为最终值
    - 位置：`warehouse-entry.ts` `upsertHistoryEntry()` + `importHistoryFromExcel()`；`api/warehouse-entry/history/route.ts`

41. **TR入仓历史库内置 66 条种子数据（K8s 无持久卷兜底）** (2026-09-03)
    - 背景：历史库 `data/history.json` 运行时写入且被 `.gitignore` 忽略，K8s 无持久卷，pod 重启即清零（见 #20），线上默认是空库
    - 做法：把历史库重建结果 66 条生成为 `src/lib/history-seed.ts`（`HISTORY_SEED: HistoryEntry[]`，打进镜像），`loadHistory()` 在 `history.json` 不存在时返回 `HISTORY_SEED.map(e => ({...e}))` 兜底；一旦运行时 `accumulateHistory` 累积后写文件，仍优先读文件
    - 数据来源：66 条 = 《拓锐入仓数据参考(1).xlsx》(58 条) + 《拓锐8.29出货1989件入仓数据（做为历史参考）.xlsx》(31 条) 按同款判定（品名+排序尺寸差≤1cm+实重差≤1kg）去重后全集
    - 位置：`src/lib/history-seed.ts`（种子）+ `warehouse-entry.ts` `loadHistory()` 兜底分支

42. **延讯下单优化大文件在低内存环境 OOM（线上 503），用 `loadFirstSheetOnly()` 精简加载修复** (2026-09-03)
    - 症状：Sealos 线上（256M/0.2 核）上传 3.7MB 延讯发票报「网络错误，请重试」，网关 503 `connection termination`；本地 192.168.3.16:3001 正常
    - 根因：该 3.7MB xlsx 解压后达 26MB——含隐藏的 VLOOKUP 数据源表（产品资料 `sheet4.xml` 12MB、`externalLink1.xml` 6MB、`sharedStrings.xml` 5.8MB，发货单 `sheet1.xml` 仅 368KB）。ExcelJS `readFile` 全量解析所有 sheet + 外部链接，低内存直接 OOM 被杀，前端 catch 又吞掉真实原因只显示「网络错误」
    - 修复：新增 `loadFirstSheetOnly()`，在交给 ExcelJS 前先用 JSZip 读取 xlsx 并精简成只含发货单 sheet——`xl/workbook.xml` 只保留第一个 `<sheet>` 并删 `<externalReferences>`/`<definedNames>`；`xl/_rels/workbook.xml.rels` 只留第一个 worksheet 关系、删 externalLink 关系；`[Content_Types].xml` 删多余 worksheet/externalLink 的 `<Override>`；再删多余 `sheetN.xml`/`_rels`/`externalLinks` 文件，`generateAsync` 出精简 buffer 再 `xlsx.load`。大文件内存从 1GB+ 降到一两百 MB
    - 连带坑：Node v24 `Buffer<ArrayBufferLike>` 与 ExcelJS 自声明的 `interface Buffer extends ArrayBuffer` 类型不兼容，`load` 入参用 `as unknown as Parameters<typeof srcWb.xlsx.load>[0]` 断言绕过（运行时内部走 jszip，Node Buffer 完全可用）
    - 位置：`yanxun-convert.ts` `loadFirstSheetOnly()` + 读取处（`loadFirstSheetOnly` 前于 `convertYanxunToEtton` 的读取逻辑）

43. **TR入仓「供应商小于客户」报警：口径统一 + 文案明确计费重维度** (2026-09-04)
    - 症状：北极光灯 `FBA19MRWWWR5` 页面「三边和」列显示 供135.00 / 客132.00（供 > 客），却报「供应商小于客户」，用户困惑
    - 根因（两层）：① 页面「供X/客Y」是**三边和**（cm），而报警比较的是**计费重**（kg）——供应商箱规 49×43×43（三边和 135、计费重 15.10），客户申报 42.2×42.4×47.4（三边和 132、实重 15.21、计费重 15.21），供应商箱子尺寸更大但实重更轻 → 计费重反而更小，报警是**正确**的，只是文案没写清「计费重」维度；② 报警判断原本取 `supplierPickedChargeable`（选数命中箱规 `picked`），而展示的「供应商计费重」列取 `supplierChargeable`（第 1 大箱 `sorted[0]`），`selectBox` 在「新品」或「第 1 大 > 历史最大」时退选第 2 大，两者不一致会误报
    - 修复：① 判断改用 `supplierChargeable`（第 1 大真实最大计费重），与展示一致；② 文案改为「供应商计费重小于客户，请确认」，明确是计费重维度，避免与三边和混淆
    - 位置：`warehouse-entry.ts` `buildSuggestions()` 报警分支（`supplierChargeable < c.chargeableWeight`）

44. **提单模板字段值在文本框 `w:txbxContent` 里，python-docx 读不到** (2026-09-04)
    - 症状：生成后的提单 docx 用 `Document().paragraphs` + `tables` 遍历，`含提单号/柜号` 等全部 False，但残留占位符却是 0 个——看似「合并没填值」，实则字段值落在文本框里
    - 根因：提单模板是带文本框/形状的复杂 Word，MERGEFIELD 的 field-result 文本嵌在 `<w:txbxContent>` 内，python-docx 的 `paragraph.text`/`cell.text` 不遍历文本框
    - 校验方法：直接解压 docx 读 `word/document.xml`，`re.sub(r"<[^>]+>","",xml)` 后判断字段值是否在纯文本里（实测 12 个字段全在）
    - 位置：`bl-service/core.py` `fill_bl_docx()`（生成本身正确，坑在验证方式）

45. **报关底单是横向 842×595，且 3/9 是扫描件无文字层** (2026-09-04)
    - 参考实现假设纵向坐标（y≈101/125/170/194），对真实横向报关单完全错位，已重写为「标签定位」（`_value_row(label)` 找标签行下方值行），并对 6 张文字层底单逐字段校准
    - 3/9 底单 `extract_words` 返回 0 字（纯扫描图片）→ 走 RapidOCR 识别（竖版先 `np.rot90` 转正），再复用标签定位提取；OCR 失败才回退 `_empty_record()`（见「扫描件 OCR 兜底」）
    - 柜号前缀不固定（CIMU/WNGU/TIIU/TLLU），正则用 `[A-Z]{4}\d{7}` 而非参考的硬编码 `MATU\d{7}`
    - 位置：`bl-service/core.py` `extract_customs_data()`

46. **docx-mailmerge 版本上限 0.5.0** (2026-09-04)
    - `requirements.txt` 若写 `docx-mailmerge>=0.6` 会安装失败（PyPI 最新仅 0.5.0），已改为 `docx-mailmerge>=0.5.0`
    - 位置：`bl-service/requirements.txt`

47. **docx→PDF 用原生 LibreOffice 转，不依赖 Docker** (2026-09-04)
    - `docx_to_pdf()` 经 `_find_soffice()` 定位可执行文件：优先 Windows 原生安装（`C:\Program Files\LibreOffice\program\soffice.exe` 及 x86 路径），再回退 `shutil.which("soffice"/"libreoffice")`（Docker/Linux）
    - 转换加独立 user profile（`-env:UserInstallation=...`）避免与用户手动打开的 LibreOffice GUI 实例「已运行/锁 profile」冲突；找不到 LibreOffice 时 `try/except` **静默跳过 PDF、只产 docx**，不阻断生成流程
    - 本机已装 LibreOffice 26.8.0（winget），Flask 原生跑（`python review_app.py`）即可产 PDF，无需 Docker
    - 位置：`bl-service/core.py` `_find_soffice()` + `docx_to_pdf()`

48. **箱货清单表头字会漏进品名；且「品名」数量可能与参考提单不一致** (2026-09-04)
    - `parse_packing_list` 首版把表头行「英文品名」四字当成了第一个品名（因为表头行该列值恰好等于「英文品名」）；已通过记录 `header_idx` 并 `i <= header_idx` 跳过表头行修复
    - 参考提单 BG20260605002 品名列 **4 个**（CILP/FOLDABLE/INS RIPPLE/MAGIC BALL），但对应箱货清单 `箱货清单-TRKJ26060003.xlsx` 实为 **5 个**（还含 `ROSE PROJECTOR LIGHT` 玫瑰投影灯，10 箱/131kg/0.827CBM）——且 54 箱、783.48kg、4.307CBM 三数都把这 10 箱算进去了，说明参考提单漏写了玫瑰投影灯（人为遗漏）。本功能按箱货清单**全量汇总 5 个**；用户已确认参考提单确属漏写，保留 5 个
    - 位置：`bl-service/core.py` `parse_packing_list()`

49. **本机 Docker Desktop 跑不起来（BIOS 虚拟化未开），故改用原生 LibreOffice** (2026-09-04)
    - 症状：Docker Desktop 首启报 `Virtualization support not detected`，`docker info` 返回 500，`wsl -l -v` 无 docker-desktop 发行版
    - 根因：Windows 11 Home + BIOS 里 AMD SVM（虚拟化）未开启，Docker Desktop 依赖 WSL2/Hyper-V 无法启动；开启需重启进 BIOS 改设置，风险高
    - 决策：本功能唯一依赖 Docker 的理由是 LibreOffice（docx→PDF），而 LibreOffice 可直接装 Windows 原生 → **放弃 Docker，改用原生 LibreOffice**（见 #47），Flask 用 `python review_app.py` 原生跑即可
    - 备注：本机已装但未启用的 Docker Desktop 保留未动，若日后要跑其他容器再处理 BIOS 虚拟化

50. **提单号应从「提运单号」标签取值，而非从收货人值行右侧猜** (2026-09-05)
    - 原逻辑 `bl_words = [w for w in crow if w["x0"] >= 500]` 把提单号绑定在「境外收货人」值行右侧（x≥500）：铁路底单里「提运单号」的值恰好与收货人同值行，碰巧取对
    - 海运底单（如 `BG20260508013`）布局不同：「提运单号」标签在 y=127，值 `G2605115309` 在 y=140 单独一行，与收货人值行（y=141）分开 → 原逻辑提单号取空
    - 已改 `_value_by_label(lines, "提运单号")`：定位「提运单号」标签的 x0，取下方同 x 列（x0±12）第一个非占位值，海运/铁路通用
    - 位置：`bl-service/core.py` `_value_by_label()` + `extract_customs_data()` 第 3 步

51. **箱货清单可能无「体积」列，需回退长×宽×高×箱数算体积** (2026-09-05)
    - 铁路单箱货清单（TRKJ26060003）有「体积」列 → 直接求和（4.307）
    - 海运单箱货清单（TRKJ26050005）**无「体积」列**，但有「长(CM)/宽(CM)/高(CM)/总箱数(CTN)」→ 原逻辑总体积取 0，提单体积字段错
    - 已改 `parse_packing_list`：有「体积」列则直接求和；无则回退 `长*宽*高*总箱数/1e6`（cm³→m³）。实测海运单 3.8859 CBM（45 箱），总重量 695.11 与底单一致
    - 位置：`bl-service/core.py` `parse_packing_list()`

52. **物流追踪表匹配 ETD 开船日（多 sheet 表头不一）** (2026-09-05)
    - 物流追踪表 5 个 sheet（2023/2024/2025/最新物流动态/10月），表头位置不一：「Shipment ID（FBA号）」在 A 列（新版）或 AA 列（2023/10月），ETD 在 D 列，2023/10月表头在第 1 行（第 0 行是标题）
    - 匹配键 = 货箱清单「FBA ID」↔ 物流追踪表「Shipment ID」，实测海运单 16/16 全命中；ETD 开船日覆盖提单「起运日期」
    - ETD 日期多是 Excel 序列号（如 45668）或 datetime，`_excel_date_str` 统一转 `YYYY-MM-DD`
    - 「船名航次」维持底单提取、**不用**物流追踪表（用户决策）
    - 位置：`bl-service/core.py` `parse_tracking_list()` + `_excel_date_str()` + `review_app.py` extract

53. **提单 FREIGHT 文本框 96pt 太窄，「FREIGHT PREPAID」换行被裁成只剩 FREIGHT** (2026-09-05)
    - 症状：模板改成「FREIGHT PREPAID」后，LibreOffice 渲染的 PDF 仍只显示「FREIGHT」（`PREPAID` 换到第二行，被 25pt 高的文本框裁掉）
    - 根因：运费文本框 96.35×25pt 太窄，带空格「FREIGHT PREPAID」16 字符在空格处换行。**正确提单就是带空格**「FREIGHT PREPAID」（早期误判为连写，已回退）
    - 修复：模板保持「FREIGHT PREPAID」带空格 + 文本框加宽 96.35→115pt（DrawingML `wp:extent`/`a:ext` cx 与 VML `width` 双分支同步改，LibreOffice 渲染 VML fallback 分支）
    - 位置：`提单模板.docx`（脚本改，备份 `.bak_freight2`）

54. **品名长行换行成多行、且早期误判「连写去空格」** (2026-09-05)
    - 症状：海运单品名「INS STYLE NORTHERN LIGHTS NIGHT LIGHT」在 109pt 宽文本框里换行成 4 行、溢出文本框
    - 根因：品名文本框 109×127.4pt 太窄（最长品名 34 字符需约 220pt）。**正确提单品名是带空格**（`CILP TABLE LAMP`），早期「连写去空格」是误解已回退
    - 修复：①`fill_bl_docx` 填品名时**保留空格**（去掉早期 `.replace(" ", "")`）；②品名文本框加宽 109→240pt（`wp:extent`/`a:ext` cx 1384300→3048000 + VML `width:109pt→240pt`）。实测 4 个品名全部单行、不重叠
    - 位置：`bl-service/core.py` `fill_bl_docx()` + `提单模板.docx`（备份 `.bak_product`）

55. **港口英文 = 中文港口名查表 + 箱货清单「客户渠道」查表，双映射自动记忆** (2026-09-05)
    - 港口英文不能只靠中文名硬译：同一目的国不同渠道对应不同港口（「易·22日达卡派包税」→ `LONG BEACH,CA`、「易·15日达卡派包税」→ `LOS ANGELES,CA`）
    - 实现：①`port_map.json`（中文港口→英文，如「盐田」→`YANTIAN`），`apply_port_map` 查表、`remember_ports` 人工修正后回写；②`channel_map.json`（客户渠道→目的港英文），`parse_packing_list` 定位「客户渠道」列，`apply_channel_map` 按渠道覆盖目的港（比中文港口名更精确）、`remember_channels` 带渠道的票人工改港后回写
    - 顺序：extract 里 `apply_port_map` 先跑、`apply_channel_map` 后跑（渠道命中则覆盖目的港）；generate 里 `remember_ports` + `remember_channels` 都做记忆回写
    - 位置：`bl-service/core.py`（`apply_channel_map`/`remember_channels`/`parse_packing_list` 加「客户渠道」列）+ `bl-service/review_app.py` + `bl-service/channel_map.json`

56. **提单文本框段落被 LibreOffice 默认渲染成左右对齐（justify），英文中间大空格** (2026-09-05)
    - 症状：shipper「GUANGZHOU TUORUI TECHNOLOGY CO., LTD」与品名第二/第三个，单词间距被拉成 justify（~14.6pt），看起来中间空太多
    - 根因：模板文本框段落 `w:txbxContent` 里的 `<w:p>` 未显式声明对齐，Word 默认 left，但 LibreOffice 渲染文本框段落时默认 justify
    - 修复：给 shipper / consignee / 品名文本框段落补 `<w:pPr><w:jc w:val="left"/></w:pPr>`，显式左对齐（间距回落到 ~2.1pt 正常值）
    - 位置：`提单模板.docx`（脚本改，备份 `.bak_align`）

57. **提单日期格式应为 `DD MMM YYYY`（带空格），且 SHIPPED ON BOARD / FREIGHT PREPAID 中间要空格** (2026-09-05)
    - 症状：早期把日期转成 `14MAY2026`、SHIPPED ON BOARD 连成 `SHIPPEDONBOARD`、FREIGHT PREPAID 连成 `FREIGHTPREPAID`，与正确提单不符
    - 根因：早期用 pdfplumber `x_tolerance=3` 提取正确提单时把空格合并，误判为「连写」；改用 `x_tolerance=1` 后确认正确提单全部**带空格**（`14 MAY 2026`、`SHIPPED ON BOARD:14 MAY 2026`、`FREIGHT PREPAID`）
    - 修复：①`_date_to_bl` 输出 `DD MMM YYYY`（`2026-05-14` → `14 MAY 2026`，月英文 + 两侧空格）；②模板「SHIPPEDONBOARD:」→「SHIPPED ON BOARD:」、「FREIGHTPREPAID」→「FREIGHT PREPAID」恢复空格
    - 位置：`bl-service/core.py` `_date_to_bl()` + `提单模板.docx`

58. **目的港左列「Port of Discharge」文本框被 LibreOffice 误渲染到右列** (2026-09-05)
    - 症状：左列 Port of Discharge（`Text Box 10`，DrawingML `positionH relativeFrom="margin" align="left"` + VML `mso-position-horizontal:left;relative:margin`）渲染到 x0≈172（右列），与右列 Place of Delivery（x0≈180）重叠，PDF 里两段「LONG BEACH,CA」叠成乱码
    - 根因：LibreOffice 对该文本框的 `margin:align left` 水平定位解析失效（同结构的「船名航次」`Text Box 7` 却正常渲染在 x0≈43，未定位到差异根因）
    - 修复：把 `Text Box 10` 的水平定位从「对齐 margin left」改成显式偏移——DrawingML `positionH relativeFrom="margin" align="left"` → `relativeFrom="column" posOffset="0"`、VML `mso-position-horizontal:left;relative:margin` → `margin-left:0pt`。修复后 Port of Discharge=x0≈43（左）、Place of Delivery=x0≈180（右），与正确提单一致
    - 位置：`提单模板.docx`（脚本改，备份 `.bak_port`）

59. **渠道名与实际运输方式不符 → 目的港歧义（公路 vs 铁路）** (2026-09-05)
    - 症状：同一渠道「英国快铁自税递延-卡派」，铁路单（6月第4周 BG20260626029，运输方式(3)=铁路）目的港=MALASZEWICZE；公路单（6月第3周 BG20260619015，运输方式(4)=公路，被升级公路运输）目的港=The U.K.。channel_map 按渠道固定映射「英国快铁自税递延-卡派」→MALASZEWICZE，对公路单会错
    - 根因：渠道名「快铁」是历史命名，个别票实际改走公路；底单「运输方式(N)」字段才准确，但当前不做运输方式感知
    - 处理：不自动区分运输方式（用户决策：升级公路的特殊票是少数个例），channel_map 保留渠道映射（对铁路单正确），公路单由人工在审核表改目的港
    - 位置：`bl-service/channel_map.json`（「英国快铁自税递延-卡派」→MALASZEWICZE）

60. **源文件夹命名「铁路/快铁」混用** (2026-09-05)
    - 症状：6月第3周 票2 源文件夹名「英国铁路不包税-卡派」，但底单/标准答案文件名是「英国快铁不包税-卡派」；命名严格照抄文件夹名 → 生成文件名跟着叫「铁路」
    - 处理：命名严格照抄源文件夹名（用户决策「不归一」，不猜），由用户保证源命名与标准一致（统一用「快铁」）

61. **品名去版本号（保留全大写）** (2026-09-07)
    - 症状：原逻辑把箱货清单「英文品名」列 `.upper()` 转全大写，但保留末尾版本号，如 `Blue Ocean Dream Galaxy Projector 3.0` → `BLUE OCEAN DREAM GALAXY PROJECTOR 3.0`（多了 `3.0`）
    - 处理：新增 `_strip_version`（正则去掉末尾版本号 `x.y`，含连写如 `Projector2.0`），品名仍全大写 → `BLUE OCEAN DREAM GALAXY PROJECTOR`；体积维持长×宽×高×总箱数逐行合计（用户确认算法正确）
    - 位置：`bl-service/core.py`（`_strip_version` + `parse_packing_list`/`parse_packing_list_weekly`/`_translate_products`）

62. **系统SO 展示（第一列 + 只读 + 多值分行）** (2026-09-07)
    - 症状：一票报关底单可能对应多个工作号（每个工作号一个系统SO），后端 `_merge_workgroups` 已把多个 SO 去重收集，用换行 `"\n".join` 写入 `rec["系统SO"]`；审核表原用单行 `<input>` 渲染，换行不显示、且字段位置靠后、可编辑易误改
    - 处理：系统SO 移到审核表第一列（紧跟「票/文件夹」后），只读展示（不允许编辑），用 `<div whitespace-pre-wrap>` 分行显示多个 SO，方便按 SO 查找
    - 位置：`bl-service/core.py`（`BL_FIELDS` 首位放 `"系统SO"`）、`src/app/bl-review/page.tsx`（单元格渲染 `c === "系统SO"` 走只读 div）

63. **运输工具 = 船名航次，手工填写即可** (2026-09-07)
    - 症状：非海运单（铁路/卡航）底单「运输工具」列为 `@` 占位符、提取不到船名航次 → 原警告「请找供应商核实班列号/车次」误导（运输工具本应等于船名航次，不是班列号/车次）
    - 处理：`core.py` 已让 `运输工具 = 船名航次`；为空时警告改为「运输工具为空，请手工填写（等于船名航次）」，用户在审核表手工填一个即可
    - 位置：`bl-service/review_app.py`（extract / extract-batch 两处警告文案）

64. **船名航次过长时自动缩字适应、不换行** (2026-09-07)
    - 症状：船名航次文本框约 108pt 宽，`CMA CGM MANTA RAY/0GVMWE`（24 字符）在默认 10.5pt 下换行，与正确提单「`CMACGMMANTARAY/0GVMWE` 8.45pt 单行」不符
    - 处理：`core.py` 新增 `_fit_field_font(text, box_pt=108, default_pt=10.5, min_pt=7.0)`——按字符估算宽度（大写≈0.78em、数字≈0.55em、空格≈0.25em），超出可用宽则 0.5pt 步进缩小（下限 7pt）；`_apply_run_font_size()` 在 `doc.merge()` 后按「船名航次」文本定位对应 `<w:t>` 的 run 并写 `w:sz`/`w:szCs`（半点单位）。短文本保持模板默认字号不变
    - 位置：`bl-service/core.py`（`_fit_field_font` / `_apply_run_font_size` / `fill_bl_docx`）

65. **电放保函复刻标准答案格式（含页眉 + 黑线）** (2026-09-07)
    - 症状：`fill_telex_docx` 程序重建的保函缺「广州拓锐科技有限公司」页眉、无黑色分隔线，字体/行距/间距/标点都与标准答案 docx 不符
    - 处理：解析标准答案 docx 的 35 个段落与页眉 shape，按段落级精确复刻——页眉「广州拓锐科技有限公司」等线 15.5pt bold + 底边框黑线（0.75pt、宽 415.35pt、右缩进 42.35pt 对齐 shape 宽度）；标题 17/15pt、正文 14pt、声明 12pt bold（对齐 Heading3 默认）；中文=宋体/等线、英文数字=Times New Roman（`add_mixed` 按中英文拆分 run）；页面边距 top2.61/bottom0/left3.15/right2.29cm、header_distance 1.63cm、行距/段前/左右缩进/全半角标点（字段用半角 `: `、TO/FROM/申请单位用全角 `：`）逐段对齐
    - 位置：`bl-service/core.py`（`fill_telex_docx`）

66. **客户配置化（提单/保函按客户分发）** (2026-09-07)
    - 背景：提单/保函最初写死为拓锐一家（shipper/consignee 常量 + 保函页眉「广州拓锐科技有限公司」硬编码）；用户要求入口加「客户」选项，拓锐为第一个，其他客户提单需求略有偏差后续补上
    - 处理：`core.py` 新增 `CUSTOMERS` 字典（key=客户，value=shipper_bl/consignee_bl/shipper_telex/consignee_telex/telex_header/telex_to/telex_from）+ `get_customer_config()`（未知客户回退默认拓锐）；`extract_customs_data`/`_empty_record`/`fill_telex_docx`/`generate_batch` 均加 `customer` 参数；`review_app.py` 三个接口从 form/payload 读 `customer`；前端 `/bl-review` 上传区顶部加「客户」下拉（默认拓锐），extract/generate 透传 `customer`
    - 注意：新增客户 = 后端 `CUSTOMERS` 加一条即可（前端下拉经 `GET /api/bl/customers` 动态拉取 `list_customers()`，无需改前端）；保函版式（35 段）暂统一拓锐版式，其他客户版式不同时需在 `fill_telex_docx` 按 customer 分发
    - 位置：`bl-service/core.py`（`CUSTOMERS`/`get_customer_config`/`list_customers`）、`bl-service/review_app.py`（`GET /api/bl/customers`）、`src/app/bl-review/page.tsx`

67. **电放保函生成 2 页 → 1 页（docDefaults 段后距/行距溢出）** (2026-09-07)
    - 症状：`fill_telex_docx` 生成的保函本该 1 页，实测溢出成 2 页（第 2 页只有页眉 + 盖章/日期两行）
    - 根因：python-docx 默认模板的 `docDefaults` 带 `<w:spacing w:after="200" w:line="276" w:lineRule="auto"/>`（每段段后距 10pt、行距 1.15），标准答案 docx 无 docDefaults 间距（段后 0、单倍行距 1.0）；35 段每段都继承 10pt 段后距 + 1.15 行距，累积溢出约 60pt
    - 处理：`fill_telex_docx` 的 `para()` 助手与页眉段落显式写 `space_after = Pt(0)` + `line_spacing = 1.0` 作为默认（显式传入的 `line`/`before` 仍覆盖），覆盖模板 docDefaults；校验 `len(pdf.pages) == 1`
    - 位置：`bl-service/core.py`（`fill_telex_docx` 的 `para()` 与页眉段落）

68. **运输工具取值修正：优先「运输工具名称及航次号」栏位而非备注** (2026-09-07)
    - 症状：欧洲卡航底单（扫描件 TRKJ26070007）「运输工具」被提取成备注里的船名 `C258T0797/71ADV02N/M`，而正确值是「运输工具名称及航次号」栏下的车架号 `/91440112M0J494`
    - 处理：`core.py` 新增 `_value_below_label()`（标签紧邻下方 dy 4–20 内、同 x 列的非占位值），「船名航次」优先取该栏值并 `lstrip("/")`，取不到才回退备注「运输工具名称：XXX」；dy 上限防止栏值为空/@ 时误抓下方「征免性质」标签（x0 相同但更远）
    - 位置：`bl-service/core.py`（`_value_below_label` / `extract_customs_data` 第 4 步）

69. **审核表「票/文件夹」换行完整显示 + 「系统SO」列横向滚动锁定** (2026-09-07)
    - 症状：票/文件夹名过长被 `truncate` 截断成省略号；横向滚动查看后面字段时「系统SO」列跟着滚走、无法对照哪一票
    - 处理：`src/app/bl-review/page.tsx` 审核表——「票/文件夹」列只显示「BG」开始后的部分（`shortFolder`，全名放 title），固定 `w-[180px]`（约 12 字）+ `break-all` 换行；「系统SO」列 `sticky left-[180px]`（配合票/文件夹列 180px 宽）一起锁定，横向滚动不动；`table` 加 `min-w-max` 防止各列被压缩成「一行 6 字」
    - 位置：`src/app/bl-review/page.tsx`（审核表 thead/tbody）

70. **「文件命名」列移除 + ZIP 命名 = 大文件夹名 + 系统制作文件** (2026-09-07)
    - 背景：审核表「文件命名」列原本显示提单号（误导，实际命名基于文件夹）；用户要求移除该列，并规定 ZIP 压缩包命名为「大文件夹名 + 系统制作文件」
    - 处理：前端 `columns` 过滤掉「文件命名」；上传时取 `webkitRelativePath` 第一段作为「大文件夹名」（rootFolder），generate 时随 `root_folder` 传后端；`generate_batch` 新增 `root_folder` 参数，ZIP 名 = `{大文件夹名}系统制作文件.zip`（空则退回 `ETTON提单_电放保函_时间戳.zip`）。保函文件名关键词「电放保函」→「保函」
    - 位置：`src/app/bl-review/page.tsx`（columns 过滤 / rootFolder / handleGenerate）、`bl-service/review_app.py`（api_generate 读 root_folder）、`bl-service/core.py`（`generate_batch` root_folder / `derive_output_name` 调用）

71. **周汇总箱货清单匹配容差 0.5→1.0kg** (2026-09-08)
    - 症状：7月第4周 BG20260724005 合并报关（TRKJ26070017+18+20，119箱）底单毛重 1949.89kg，箱货清单三工作号合计 1950.44kg，差 0.55kg 超过原容差 0.5kg → `match_weekly_packing` 匹配失败，系统SO 号码为空
    - 处理：`match_weekly_packing` / `_subset_sum` 的毛重容差从 `< 0.5` 放宽到 `< 1.0`（报关单毛重与箱货清单求和存在 ±0.5~1kg 舍入差异）
    - 位置：`bl-service/core.py`（`match_weekly_packing` / `_subset_sum`）

72. **提单批量上传大文件夹报 500「Request body exceeded 10MB」→ 后端 ClientDisconnected** (2026-09-08)
    - 症状：上传整个「拓锐7月份底单」目录（35 个 PDF，约 27MB）时，前端报「请求失败：Unexpected token 'I', "Internal S..." is not valid JSON」；Next.js 日志 `Request body exceeded 10MB for /api/bl/extract-batch. Only the first 10MB will be available...` + `Failed to proxy ... socket hang up { code: 'ECONNRESET' }`；Flask 端 `werkzeug.exceptions.ClientDisconnected: 400 Bad Request`
    - 根因：Next.js rewrite 代理默认对请求体做 clone（供 middleware/route 复用），上限 `DEFAULT_BODY_CLONE_SIZE_LIMIT = 10MB`（`body-streams.js`）。超过 10MB 时 Next.js 截断请求体（只传前 10MB）再转发到 Flask，Content-Length 与实际字节不符 → Flask 解析 multipart 读到一半流就结束 → 抛 `ClientDisconnected`；Next.js 侧报 `socket hang up`/`ECONNRESET` 回 HTML 500。单周目录（约 7MB）不超限所以之前能成功，整月目录（约 27MB）才触发
    - 处理：①`next.config.ts` 加 `experimental.middlewareClientMaxBodySize: "100mb"`（放宽 clone 上限，根治截断）；②前端 `readJson()` 兜底——`res.text()` 后 `JSON.parse`，失败按 HTTP 状态抛友好中文错误；③后端 `@app.errorhandler(Exception)` 兜底——未捕获异常返回 JSON（而非 HTML 500），traceback 打到 pm2 error log
    - 位置：`next.config.ts`（`middlewareClientMaxBodySize`）、`src/app/bl-review/page.tsx`（`readJson`）、`bl-service/review_app.py`（`handle_unexpected`）

73. **TR全量入仓（warehouse-entry-full）是 TR入仓（warehouse-entry）的逐箱分支，两套 lib 独立维护** (2026-09-08)
    - 背景：逐箱输出 + 11 级尺寸修正的新逻辑作为独立新功能「TR全量入仓数据整理」发布，原「TR入仓数据整理」（代表箱选数 + 历史库）保留不动
    - 实现：`warehouse-entry-full.ts` 从 `warehouse-entry.ts` 复制后剥离历史库（`HistoryEntry`/`loadHistory`/`accumulateHistory`/`importHistoryFromExcel`/`exportHistoryBuffer` 等全部删除，`fs`/`path` import 一并移除），`buildSuggestions(customers, boxes)` 无 history 参数，逐箱输出 + `applyDimensionRules`（11 级）+ `forceAmplify`（仅全局兜底用）
    - 坑：两套 lib 共享同一套 客户/供应商格式检测（`CUSTOMER_*_PATTERNS`/`SUPPLIER_*_PATTERNS`）与 `exportOutputBuffer` 列位逻辑，后续新增格式或改导出列需**两边同步**，否则新功能会漏掉新格式
    - 位置：`src/lib/warehouse-entry-full.ts` + `src/app/warehouse-entry-full/` + `src/app/api/warehouse-entry-full/`

74. **提单批量生成（generate）超 30 秒被 Next.js 代理提前断开 → socket hang up / ECONNRESET** (2026-09-08)
    - 症状：前端点「确认生成」批量生成（如 29 票）时，报「请求失败：服务返回异常（HTTP 500）」（前端 `readJson` 的兜底文案）；Next.js 日志 `Failed to proxy http://localhost:5000/api/bl/generate [Error: socket hang up] { code: 'ECONNRESET' }`，而 Flask 端无新 traceback（generate 后端仍在跑或已完成）——即**后端没崩，是代理层主动断的**
    - 根因：Next.js rewrite 代理对转发请求默认 **30 秒超时**（`proxy-request.js`: `proxyTimeout: proxyTimeout === null ? undefined : proxyTimeout || 30000`，单位毫秒）。`generate` 每票 2 次 LibreOffice 转换（提单 PDF + 保函预览 PDF），实测单次转换约 3.6s，29 票 ≈ 58 次转换 ≈ 209s（约 3.5 分钟），远超 30s → 代理 30s 后主动断开上游连接 → Next.js 侧 `socket hang up`/`ECONNRESET` 回 HTTP 500。与 #72（extract 请求体超 10MB 被截断）是**同一次「整周多票」故障的两个不同根因**：#72 是「请求体太大被截断」，本坑是「响应太慢被超时断开」，都发生在 Next.js rewrite 代理层
    - 处理：①`next.config.ts` 加 `experimental.proxyTimeout: 600000`（10 分钟，毫秒），让批量生成有充足时间；②顺手修 `core.py` 的 `docx_to_pdf`——原来只捕获 `(FileNotFoundError, OSError)`，未捕获 `subprocess.TimeoutExpired`（单次 LibreOffice 超 120s 会抛未捕获异常 → 整批 500），补上后单票转换超时降级为「仅产 docx」不阻断整批
    - 验证：经 `localhost:3001` 代理发 6 票 generate（12 次转换，实测 32.18s > 30s），返回 HTTP 200 `ok=true`、6 条 previews，确认超过原 30s 线不再断开
    - 位置：`next.config.ts`（`proxyTimeout`）、`bl-service/core.py`（`docx_to_pdf`）

75. **拆分报关资料靠文件名「报/放/委托」排序不可靠 → 改读首页表头识别类型排序合并** (2026-09-08)
    - 症状：`BG20260716005 2票 英国快铁不包税-卡派 7箱 是7` 文件夹「扫描不到内容」；文件夹内 3 个 PDF 命名不规范——`广州拓锐科技有限公司.pdf`（实为委托报关协议）、`广州拓锐科技有限公司7件 .pdf`（报关单）、`广州拓锐科技有限公司7件.pdf`（放行单），文件名都不含「报/放/委托」关键词
    - 根因：旧 `_declaration_sort_key` 靠文件名「报/放/委托」排序，这三个文件全归「其他=3」，按字典序 `广州拓锐科技有限公司.pdf`（`.` 排 `7` 前）落最前 → 合并后第一页是「委托报关协议」；`extract_customs_data` 只读第一页（`pdf.pages[0]`）→ 报关单标签全对不上 → 字段全空
    - 处理（按用户「三类要合并成一个 PDF，先报关单→放行单→委托协议，不管命名如何、读表头判断」的要求）：
      ① `merge_declaration_pdfs` 改读首页表头识别类型 `_declaration_kind`：报关单=0（「出口/进口货物报关单」标题）→ 放行单=1（「放行通知书」标题）→ 委托协议=2（「委托报关协议」标题）→ 其他=3；按此排序后**三类都合并不丢弃**（委托协议也是报关资料一部分）
      ② `extract_customs_data` 改遍历各页，选第一页含报关单关键标签的页兜底（即便排序后第一页非报关单也能提取）
    - 注意：判断关键词要用「出口货物报关单/放行通知书」这类**标题词**，不能用「通关无纸化」——报关单备注栏也常写「通关无纸化」，会把报关单误判成放行单（实测 `39件 报.pdf` 备注含「通关无纸化」）
    - 位置：`bl-service/core.py`（`_declaration_kind` / `merge_declaration_pdfs` / `extract_customs_data`）

76. **运输工具（船名航次）值含空格被拆成多个 word → 只提取到第一个词** (2026-09-08)
    - 症状：`BG20260724008 美转加美森--加东` 底单运输工具应为「HAWK I/15E」，但提取表格只写了「HAWK」；同理含斜杠的「UN9951147/2611E」之前也只取到「UN9951147」
    - 根因：`_value_below_label` 按「标签 x0 ± x_tol」匹配下方值行的**单个 word** 就返回；英文值含空格时 pdfplumber 拆成多个 word（HAWK x0=387、I/15E x0=409.5），第二个 word 距标签 x0 超 12pt → 只返回「HAWK」
    - 处理：`_value_below_label` 改为从标签列开始**收集同值行 x 连续（gap ≤ 15pt）的多个 word 拼接**，遇到大 gap（下一字段列，如「提运单号」与运输工具列 x0 差 87pt）即停
    - 位置：`bl-service/core.py`（`extract_customs_data` 内 `_value_below_label`）

77. **`_declaration_kind` 读不到扫描件首页文本 → 拆分扫描件被排到末尾** (2026-09-08)
    - 症状：扫描件报关资料（无文字层）拆成多份 PDF 上传时，`_declaration_kind` 读 `pdf.pages[0].extract_text()` 得到空串 → 全部归类 kind=3（其他），合并排序退化为按文件名，报关单可能排到放行单/委托协议之后
    - 根因：`_declaration_kind` 只读文本层，扫描件首页无文本层（如 `TRKJ26070007` / `TRKJ26070060` 广州拓锐底单，4 页纯图片）
    - 处理：首页文本为空时用 `_ocr_page_to_words(page)` OCR 首页再判断类型（复用 RapidOCR，模型有全局缓存不重复加载）；有文字层的正常 PDF 不受影响、不额外耗时
    - 位置：`bl-service/core.py`（`_declaration_kind`）

78. **RapidOCR 依赖 opencv，slim 镜像缺 libGL 等图形库 → 扫描件 OCR 静默失败** (2026-09-08)
    - 症状：线上 Sealos 部署后，扫描件 extract 返回 `ok:true` 但提单号/箱数等字段全空（本地正常）；后端日志无任何报错（`_ocr_page_to_words` 的 `except Exception: return []` 静默吞掉）
    - 根因：`rapidocr_onnxruntime` 依赖 `opencv-python`，其 `cv2` 在 `python:3.11-slim` 里 import 时报 `ImportError: libGL.so.1: cannot open shared object file`（缺 libgl1 等图形库）；onnxruntime 另需 `libgomp1`（OpenMP 运行时）
    - 处理：`bl-service/Dockerfile` 补 `libgomp1 libgl1 libglib2.0-0 libsm6 libxrender1 libxext6`；排查方法 `kubectl exec deploy/etton-bl-service -- python -c "from rapidocr_onnxruntime import RapidOCR"`
    - 位置：`bl-service/Dockerfile`

79. **bl-service 以 root 运行触发 PodSecurity `restricted` 警告；LibreOffice javaldx 警告无害** (2026-09-08)
    - 症状：`kubectl apply` 时告警 `would violate PodSecurity "restricted:v1.25"`（allowPrivilegeEscalation / capabilities / runAsNonRoot / seccompProfile），当前 namespace 为 warn 模式不阻断，pod 正常运行
    - 根因：LibreOffice headless 首次运行需写 `$HOME/.config` 建 profile，非 root 且无 HOME 时 soffice 会静默失败 → 特意用 root；`soffice` 报的 `Warning: failed to launch javaldx` 是 Java 不可用的无害警告，writer_pdf_Export 不依赖 Java
    - 处理：暂保持 root（功能优先）；如未来 namespace 升级 enforce，需 Dockerfile 设 `ENV HOME=/tmp` + deployment 加 `runAsNonRoot/runAsUser/capabilities.drop=["ALL"]/seccompProfile`
    - 位置：`k8s/bl-service-deploy.yaml`、`bl-service/Dockerfile`

80. **port_map.json 从扁平改为结构化 `{origin, destination}`；新增 `/bl-mapping` 在线编辑入口** (2026-09-09)
    - 动机：原 `port_map.json` 是扁平 `{中文: 英文}`，把「离境口岸」（起运港）和「运抵国」（目的港）混在一张表，语义不清、编辑时难区分；用户需要能直接编辑起运港/目的港映射的入口
    - 改动：`port_map.json` 改为 `{origin: {离境口岸→英文}, destination: {运抵国→英文}}`；`load_port_map` 兼容旧扁平格式（无 origin/destination 键时全量归 destination 兜底）；`apply_port_map` 起运地查 origin、目的港查 destination；`remember_ports` 分别回写 origin/destination
    - 入口：新增 `/bl-mapping` 页面（三个可增删改表：起运港/目的港/渠道）+ `GET/POST /api/bl/mappings`（读写两个 json，键值校验必须为字符串）
    - 位置：`bl-service/core.py`（load/apply/remember）、`bl-service/review_app.py`（mappings 接口）、`src/app/bl-mapping/page.tsx`

81. **起运港按运输方式分流 + 新增关区名映射 `customs_office_map.json`** (2026-09-09)
    - 动机：非海运票（铁路/公路卡航/航空）的「离境口岸」是边境口岸（如阿拉山口、霍尔果斯），不代表实际发货城市；真正的起运港应取报关单「海关编号」后面的**关区名**（报关关区 ≈ 发货城市，如增城海关→广州、蓉青关→成都）
    - 规则：`apply_port_map` 按「运输方式」分流起运港——海运（`水路运输`）→离境口岸→`port_map.origin`；非海运（`铁路运输`/`公路运输`/`航空运输`）→关区名→`customs_office_map.json`。目的港仍走 `port_map.destination`。`remember_ports` 对称回写（海运→origin、非海运→customs_office_map）
    - 关区名提取：`_customs_office_from_text` 正则 `海关编号[：:]?\s*(\d+)\s*[\(（]([^\)）]*)[\)）]`，兼容 `(增城海关)` 和 `(7901) 蓉青关` 两种格式；OCR 扫描件偶发括号丢失（如塔城海关）需人工兜底
    - 历史扫描：1-7月 48 票底单，海运 20 票（5 离境口岸：外高桥9/盐田5/宁波北仑3/南沙港一期2/烟台1）+ 非海运 28 票（7 关区名：增城10/渝州6/塔城4/蓉青3/车站2/深圳湾2/都拉塔1），无航空
    - 入口：`/bl-mapping` 在 #80 三表基础上新增第 4 表「关区名映射」；`review_app.py` 的 `GET/POST /api/bl/mappings` 加 `customs_office_map` 读写
    - 位置：`bl-service/core.py`（`_customs_office_from_text`/load/apply/remember）、`bl-service/customs_office_map.json`、`bl-service/review_app.py`、`src/app/bl-mapping/page.tsx`

82. **渠道大类规则 + 追踪表扩展解析 ETA/船名航次 + 电放保函申请日期=到港时间** (2026-09-09)
    - 动机：①目的港/起运地需按「渠道大类」（铁路/卡航/快递）区分，而非单一 destination 映射——铁路到德国要 `DUISBURG`（非 destination 里的 `MALASZEWICZE`）、卡航目的港写国家名（`THE UK`/`GERMANY`）、快递（香港联邦IP）起运地固定 `SHENZHEN`；②铁路底单无船名航次，需从追踪表取班列号；③电放保函「申请日期」= 到港时间（ETA 提前两天）
    - 渠道大类：`_channel_category` 按客户渠道关键词分（铁路含`铁路`/`快铁`/`铁派`，卡航含`卡航`，快递含`联邦`/`空派`），顺序匹配先命中先返回（避免「铁路卡派」误判为卡航）
    - `apply_port_map(rec, channel)` 分流：起运港 快递→`SHENZHEN`/海运→离境口岸→origin/非海运→关区名→customs_office_map；目的港（改按「运抵国」查，修正原用「指运港」查 destination 的错位）铁路→`{英国:MALASZEWICZE,德国:DUISBURG}`、卡航→`{英国:THE UK,德国:GERMANY}`、其他→destination
    - `apply_channel_map` 仅对非规则渠道用 channel_map 兜底（否则旧 channel_map「欧洲铁路包税--卡派→MALASZEWICZE」会把铁路德国票的 DUISBURG 错误覆盖回 MALASZEWICZE）；`remember_ports` 快递起运地固定不记忆、`remember_channels` 规则渠道不记忆
    - 追踪表：`parse_tracking_list` 返回 `{FBA ID: {etd, eta, vessel}}`（表头含 ETD/ETA/船名航次定位列）；`_arrival_date(etd, eta)`=ETA−2天、ETA−ETD<2天取 ETA；`_date_cn` 转「Y 年 M 月 D 日」填「申请日期」
    - 文件命名：`derive_output_name(folder, "保函")` → `"电放保函"`（对齐 specs 的 `{命名}电放保函.docx`）
    - 位置：`bl-service/core.py`、`bl-service/review_app.py`

83. **柜号误匹配提单号（海运多票复现）** (2026-09-09)
    - 症状：柜号应填 `GOSU1060349`，提单/电放保函却写成 `UNGB1391012`（实为提单号 `ZIMUNGB1391012S` 的中间段 `UNGB1391012`），多票海运都错
    - 根因：柜号原用全文本 `_find_any(lines, r"[A-Z]{4}\d{7}")` 取第一个匹配，但海运提单号（如 `ZIMUNGB1391012S`）里 `UNGB1391012` 恰是「4 字母 + 7 数字」，且提单号行 y 比备注行更靠前，被误当柜号
    - 修复：新增 `_container_numbers`，只在备注「集装箱标箱数及号码：N;XXXX1234567;」上下文里 `findall` 柜号（海运/铁路底单格式一致），多柜分号拼接去重；无该上下文返回空（宁缺勿错，不再全文本兜底）
    - 位置：`bl-service/core.py`（`_container_numbers` + `extract_customs_data` 第 5 步柜号）

84. **新增「星速(HNXS)」客户：Amazon FBA 直送，无物流追踪表、无需电放保函** (2026-09-11)
    - 背景：星速是 FBA 直送客户，B/L 极简——consignee 固定 `AMAZONFULFILMENTCENTER`、notify 固定 `SAMEASCONSIGNEE`、shipper 变量（境内发货人中文 → pinyin 大写，经 `pypinyin.lazy_pinyin` 拼接后 `upper()`，括号内备注先剔除）；**不出电放保函**（`CUSTOMERS["星速"]["no_telex"]=True`，generate 跳过 telex，ZIP 只含提单+底单）
    - 数据源差异：无「物流追踪表」，改由「订单列表」xlsx 提供 发往国家/开船时间/业务类型/船名航次；箱货清单的 FBA ID = 基础 12 位（`FBA`+9 位）+ `U` + 6 位序号 = 19 位，解析时用 `_base_fba` 截前 12 位分组
    - 目的港映射 `xs_dest_map.json`（`{海运/陆运: {国家→英文}}`）：海运→实际港口（英国`FELIXSTOWE`、德国`ROTTERDAM,NL`）、陆运→国家英文名（英国`BRITAIN`、德国`GERMANY`）；起运港按运输方式固定——海运`YANTIAN`、陆运`SHENZHEN`；起运日期=订单列表「开船时间」无条件覆盖
    - ⚠️ 坑1：`parse_order_list` 用 `openpyxl.load_workbook(read_only=True)` 时**只返回表头行、数据行全丢**（订单列表 xlsx 的 merged/格式导致）→ 必须 `load_workbook(path, data_only=True)`（去掉 read_only），实测 193 个 FBA 全部命中
    - ⚠️ 坑2：订单列表「开船时间」单元格是 `"2026-07-03 00:00:00"` 带时分秒的字符串 → `_excel_date_str` 用正则 `^(\d{4}-\d{2}-\d{2})` 剥离
    - FBA 匹配：`extract_customs_data` 里星速用 `_fba_from_text(full_text)` 从报关单全文取 FBA；`review_app` 里再优先 `_fba_from_text(folder)`（文件夹名）匹配订单列表/箱货清单，避免底单与文件夹名 FBA 不一致时错位
    - 位置：`bl-service/core.py`（`_cn_to_pinyin_upper`/`_base_fba`/`_fba_from_text`/`load_xs_dest_map`/`apply_xs_map`/`parse_order_list`/`parse_packing_list_xs`/`CUSTOMERS["星速"]`）、`bl-service/xs_dest_map.json`、`bl-service/review_app.py`（extract-batch/generate 星速分支）、`src/app/bl-review/page.tsx`
85. **星速提单 7 月真实数据比对后的一轮修复（6 处 + 发现的数据坑）** (2026-09-11)
    - 用 18 张人工提单 PDF 作标准答案，逐字段比对自动化输出，修复以下 6 处结构性问题（`core.py`）：
      1. **shipper 海运票为空**：报关底单「境内发货人」值行被 pdfplumber 拆成两行（公司名 y=101 / 关区名 y=100 基线差 1pt），`_value_row` 原来只取第一行（关区名 x0>250 被 `_leftmost` 丢弃）→ 改为合并标签下方 dy 范围内、与首个值行间距 ≤4pt 的相邻行（`_value_row` 加 `merge_gap` 参数）
      2. **起运地**：原来海运固定 `YANTIAN` 是错的（外高桥/南沙港出口的票应为 SHANGHAI/NANSHA）→ 海运改查 `port_map.json` 的 `origin`（离境口岸→英文），无对照时保留中文供人工；陆运仍固定 `SHENZHEN`
      3. **目的港**：① 加拿大美转加按「供应渠道」分流——含 `美森`/`CLX` → `LONGBEACH,CA`，否则 → `LOSANGELES,CA`（用户原话两处都写 LONGBEACH，实测非美森票人工写 LOSANGELES，故按数据修正）；直航加拿大（加东/加西普船）留空人工填。② 法国海运 `LE HAVRE` → `ROTTERDAM,NL`（`xs_dest_map.json` 同步改）
      4. **体积**：原来用箱货清单「长×宽×高×箱数」是错的 → 改用订单列表「总CBM」列（与人工提单体积精确吻合，如 FBA15LZ786KB=1.6762）
      5. **多 FBA 漏品名/体积**：底单文件名有逗号分隔（`FBA15LZW7KR4,FBA15M03GFHD`，同订单一行，总CBM 已聚合）和 `+6位后缀` 缩写（`FBA15M0WQDJZ+0S50DJ` = 两单合拼，第二个 FBA 与前一个共享前 6 位 `FBA15M`，总CBM 需按唯一订单行求和）→ 新增 `_fbas_from_text` 解析全部 FBA + `xs_apply_packing` 统一回填（品名跨 FBA 去重合并、体积按系统SO 去重求和）
      6. **品名连写**：箱货清单英文品名去空格（`SOY CANDLES` → `SOYCANDLES`，对齐人工提单）
    - ⚠️ 比对仍剩的差异均为**源数据/人工答案本身不一致**，非代码 bug：
      - 品名命名不统一（9 票）：箱货清单英文品名里「香薰蜡烛」有 `Scented candle` / `Aromatherapy candle` 两种，人工提单大多统一写 `SOYCANDLES`（仅 FBA15M0SYJZL 写 `SCENTEDCANDLE`）；**已定：保持箱货清单英文品名原样（不归一）**，人工的 SOYCANDLES 归一视为人工简化，代码忠实输出源数据英文品名
      - 体积精度（1 票）：FBA15M01RD74 订单列表总CBM=0.2849，人工提单四舍五入写 0.28
      - 起运地（1 票）：FBA19HC7SYNT 报关单离境口岸=外高桥(→SHANGHAI)，人工提单写 NANSHA（与同船同柜的 FBA19HBRLVV3=SHANGHAI 矛盾，疑人工笔误）
      - 目的港（1 票）：FBA15M0SYJZL 法国陆运人工写 `FRENCH`（语法错，应为 FRANCE）
      - 日期（1 票）：FBA15M0QLY0L 订单列表开船时间=2026-07-29，人工提单写 7-20（差 9 天，源数据差异）
    - 新增 `xs_apply_packing(rec, folder, xs_packing, order_map)` 取代 `review_app.py` 里散落的星速匹配逻辑（`_fba_from_text` + 单 FBA 查箱货清单 + 单 FBA 查订单列表）

86. **星速批量上传「18 份底单只解析出 1 份」+ 体积/开船时间提炼不到——根因是客户下拉停在「拓锐」** (2026-09-11 修复)
    - 症状：用户上传 7 月底单 18 份 PDF，结果表格只出现 1 票（文件夹名「7月底单」），且体积/开船时间/系统SO 全空
    - 根因：客户下拉（拓锐 vs 星速）**不持久化**，刷新页面即重置回默认「拓锐」；前端重启后用户未重新选「星速」就直接上传。拓锐模式下分组规则是「每个子文件夹 = 一票」，18 份 PDF 同在「7月底单」目录 → 被并成 1 票；且拓锐分支不解析订单列表/星速箱货清单，故体积（订单列表总CBM）、开船时间（订单列表开船时间）、系统SO（箱货清单）全部为空
    - 修复（`src/app/bl-review/page.tsx`）：
      1. 客户选择持久化到 `localStorage("bl_customer")`——`useState` 惰性初始化读 localStorage，下拉 onChange 写回；`useEffect` 仅在无持久化值时回退后端默认值
      2. 星速目录上传时文件名可能带相对路径前缀（部分浏览器 `webkitdirectory` 把 `7月底单/xxx.pdf` 塞进 `File.name`），`handleFolderSelect` 里先 `f.name.split("/").pop()` 取纯文件名再 `xsTicketBase`
    - 验证：直接跑 `core.parse_order_list` + `parse_packing_list_xs` + `xs_apply_packing` 遍历 18 份底单，18/18 都正确回填体积与开船时间（体积=订单列表总CBM，与人工提单逐票吻合）；故核心逻辑无误，纯属前端客户态问题

87. **星速提单改套用拓锐 ETTON 邮件合并模板（废弃「极简版」）** (2026-09-11 修复)
    - 症状：用户发现星速生成的提单是「带字段标签的极简版」（`B/L NO:`/`Shipper:`/`Consignee:` 等标签印在 PDF 上），与拓锐导出的真实 ETTON 提单模板完全不符；人工星速提单实为「无标签、标准 ETTON 版式」
    - 根因：`generate_batch` 里对 `is_xs` 单独走了 `fill_xs_bl_docx()`（从零用 `python-docx` 拼段落、每行带 `label:` 前缀），而不是与其他客户共用 `fill_bl_docx()`（`提单模板.docx` 邮件合并 MERGEFIELD）
    - 修复：删除 `generate_batch` 的 `if is_xs` 提单分支，所有客户统一走 `fill_bl_docx(bl_template, ...)`；删除废弃的 `fill_xs_bl_docx` / `_xs_date_short`（星速日期改用模板统一 `_date_to_bl` 输出 `DD MMM YYYY`，如 `03 JUL 2026`）
    - 效果：星速提单与拓锐共用同一模板——shipper 拼音大写、consignee `AMAZONFULFILMENTCENTER`、notify 由模板静态文字 `SAME AS CONSIGNEE` 自动显示；唛头 `N/M`、装卸 `CFS TO CFS`、单位 `CTNS`/`KGS`/`CBM`、`FREIGHT PREPAID`、`SHIPPED ON BOARD:` 均为模板静态文字；柜号海运有值、陆运留空。18 票实测全部生成，shipper 长文本（最长 48 字符）换行与人工提单一致
    - 位置：`bl-service/core.py`（`generate_batch`）

### 待重构项

- [ ] 将 session 存储从内存 Map 改为临时文件或 Redis
- [ ] 统一四个 API 的 session 管理逻辑（目前各自维护独立的 store — split-insurance, convert-pacific, reconciliation, pipixiong-split）
- [ ] `RATES` 常量（USD=7, EUR=8）硬编码在 `split-insurance.ts`，应与太平洋转换的汇率统一
- [ ] 移除未使用的 `@anthropic-ai/sdk` 依赖（或实现实际 AI 功能）
- [ ] 添加上传进度条（当前只有旋转动画，无百分比）
- [ ] `public/output/` 下的测试文件应清理或移到 `保单拆分功能/output/`
- [ ] 对账功能的列检测规则可进一步细化（如支持更多金额列别名）
- [ ] 多供应商对账引擎中 `reconciliation.ts`（旧版天图对账）可移除，统一使用 `multi-supplier-reconciliation.ts`
- [ ] 供应商配置（17家）可考虑外置为 YAML/JSON 配置文件，支持热加载
- [ ] SEALOS 部署增加 liveness/readiness 探针失败时的自动告警

---

## 6. 运行 & 测试命令

### 开发环境

```bash
# 安装依赖
npm install

# 启动开发服务器 (默认 http://localhost:3000)
npm run dev

# 启动并暴露给局域网 (同网段设备可访问)
npx next dev -H 0.0.0.0 -p 3002
```

### 生产构建

```bash
# 构建
npm run build

# 运行生产版本
npm run start
```

### 代码检查

```bash
npm run lint
```

### Docker 构建

```bash
docker build -t etton-tools .
docker run -p 3000:3000 etton-tools
```

### 手动功能测试

1. 访问 `http://localhost:3000`
2. 点击"保单投保区间拆分" → 上传 `保单拆分功能/易通下单05.19 101SO2605130039 -110箱.xlsx` → 验证生成 5 个区间文件
3. 点击"太平洋货箱清单转换" → 上传货箱清单 → 调整汇率 → 验证拆分结果
4. 点击"天图请款对账" → 上传天图账单 + 请款明细 → 验证差异检测和标红
5. 点击"皮皮熊账单拆分" → 上传皮皮熊账单 → 验证按报关单号拆分和 ZIP 下载
6. 下载各工具的输出文件，用 Excel/WPS 打开验证

---

## 7. 对外 API

### 无外部 API 依赖

本项目为纯离线工具，不调用任何外部 API。

### 环境变量

| 变量 | 必需 | 说明 |
|------|------|------|
| `ANTHROPIC_API_KEY` | 否 | Claude API Key（预留，当前未使用） |
| `NODE_ENV` | 否 | `production` 时启用优化 |
| `NEXT_TELEMETRY_DISABLED` | 否 | Dockerfile 设为 1 禁用遥测 |

### CI/CD 流程

```
Git push main
  → GitHub Actions: docker-build.yml
    → docker/build-push-action@v5
      → ghcr.io/etton-ai/etton-tools:latest + :sha
        → kubectl rollout restart deployment/etton-tools -n ns-wqw6rrmf
```

#### 部署维护命令

```bash
# 查看 Pod 状态
kubectl get pods -n ns-wqw6rrmf -l app=etton-tools

# 重启部署（拉取最新镜像）
kubectl rollout restart deployment/etton-tools -n ns-wqw6rrmf

# 查看日志
kubectl logs -n ns-wqw6rrmf -l app=etton-tools --tail=50

# 重新创建 Deployment（如果丢失）
kubectl apply -f k8s/deploy-sealos.yaml

# 本地构建验证
npm run build
```

#### 本机内网部署（192.168.3.16:3001，pm2 生产模式）

> 2026-09-02 由 `next dev -p 3001`（开发模式）切换为 pm2 托管的 `next start` 生产模式。配置见项目根目录 `ecosystem.config.cjs`。

```bash
cd "C:/Users/berry/Downloads/ETTON 电商AI"
npm run build                         # 先构建生产版本
pm2 restart etton-tools               # 重启（生产模式，端口 3001）
pm2 logs etton-tools                  # 查看日志
pm2 start ecosystem.config.cjs        # 首次启动
pm2 save                              # 保存进程列表（pm2 重启后自动恢复）

# 若需开机自启（需管理员权限，一次性）：
pm2 startup
pm2 save
```

> 注意：切生产模式前必须**彻底停掉所有 `next dev` 进程**（`netstat -ano | grep 3001` 找到 PID 后 `taskkill /F /T /PID <pid>`），否则 dev 会占用 3001 端口导致 `EADDRINUSE`，或覆盖 `.next` 生产构建导致 `Could not find a production build`。

#### Docker 构建失败排查

1. 在 GitHub Actions 页面检查构建日志
2. 常见原因: ESLint 错误（`npm run build` 包含 lint）
3. 本地验证: `npm run build` 必须通过
4. 修复后重新 push 触发构建

---

## 附录: 与 price-system 仓库的关系

本项目原包含 FBA 比价功能（`src/app/price-query/`、`比价工具/`），已于 2026-06-29 拆分为独立仓库 [price-system](https://github.com/etton-ai/price-system)。两个仓库独立构建、独立部署：

| 项目 | 端口 | Ingress | 用途 |
|------|------|---------|------|
| ETTON TOOLS | 3000 | `vftnaopzqgqv.cloud.sealos.io` | 保单拆分 + 太平洋转换 + 多供应商对账 + 皮皮熊拆分 |
| Price System | 3000 | `wlylcsujbziw.cloud.sealos.io` | FBA 多供应商比价查询 |
