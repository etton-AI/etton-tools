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
