# ETTON TOOLS — 开发者规格文档

> 最后更新: 2026-07-21 | 维护者: berry-bi

---

## 1. 项目概述与技术栈

**ETTON 效率提升助手** — 基于 Next.js 的 Web 工具集，面向易通科技内部物流操作，提供保单投保区间拆分、太平洋货箱清单转换、多供应商对账引擎和皮皮熊账单拆分四个核心功能。

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
│   └── output/                       # 历史测试输出（仅供开发参考）
│       ├── 不足5000RMB.xlsx
│       ├── 5000-10000RMB.xlsx
│       ├── 10000-20000RMB.xlsx
│       ├── 20000-30000RMB.xlsx
│       └── 30000-40000RMB.xlsx
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
│   │   └── api/
│   │       ├── split-insurance/
│   │       │   └── route.ts          # POST 上传 + GET 下载（session 管理）
│   │       ├── convert-pacific/
│   │       │   └── route.ts          # POST 上传（含汇率）+ GET 下载
│   │       ├── reconciliation/
│   │       │   └── route.ts          # [旧] 单供应商对账 API
│   │       ├── multi-supplier-reconciliation/
│   │       │   └── route.ts          # POST 双文件上传 + GET 下载（多供应商）
│   │       └── pipixiong-split/
│   │           └── route.ts          # POST 上传 + GET 下载 ZIP
│   ├── components/
│   │   └── Header.tsx                # Header (sticky) + Footer（同文件导出）
│   └── lib/
│       ├── split-insurance.ts        # 保单拆分核心逻辑 (556 行)
│       ├── convert-pacific-insurance.ts  # 太平洋转换核心逻辑 (502 行)
│       ├── reconciliation.ts         # [旧] 单供应商对账逻辑 (272 行)
│       ├── multi-supplier-reconciliation.ts  # 多供应商对账引擎 (17家供应商配置，~870行)
│       └── pipixiong-split.ts        # 皮皮熊拆分核心逻辑 (371 行)
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

纯展示型导航页，包含四个卡片链接：
- `/insurance-split` — 保单投保区间拆分
- `/pacific-convert` — 太平洋货箱清单转换
- `/multi-supplier-reconciliation` — 多供应商对账引擎
- `/pipixiong-split` — 皮皮熊账单拆分

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
