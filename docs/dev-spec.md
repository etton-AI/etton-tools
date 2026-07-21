# ETTON TOOLS — 开发者规格文档

> 最后更新: 2026-07-09 | 维护者: berry-bi

---

## 1. 项目概述与技术栈

**ETTON 效率提升助手** — 基于 Next.js 的 Web 工具集，面向易通科技内部物流操作，提供保单投保区间拆分和太平洋货箱清单转换两个核心功能。

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

- **GitHub**: `etton-ai/etton-ecommerce-ai`
- **容器镜像**: `ghcr.io/etton-ai/etton-ecommerce-ai:latest`
- **Sealos Ingress**: `deprysrenldz.cloud.sealos.io`

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
│   │   │   └── page.tsx              # 天图请款对账页面（双文件上传）
│   │   ├── pipixiong-split/
│   │   │   └── page.tsx              # 皮皮熊账单拆分页面
│   │   └── api/
│   │       ├── split-insurance/
│   │       │   └── route.ts          # POST 上传 + GET 下载（session 管理）
│   │       ├── convert-pacific/
│   │       │   └── route.ts          # POST 上传（含汇率）+ GET 下载
│   │       ├── reconciliation/
│   │       │   └── route.ts          # POST 双文件上传 + GET 下载
│   │       └── pipixiong-split/
│   │           └── route.ts          # POST 上传 + GET 下载 ZIP
│   ├── components/
│   │   └── Header.tsx                # Header (sticky) + Footer（同文件导出）
│   └── lib/
│       ├── split-insurance.ts        # 保单拆分核心逻辑 (556 行)
│       ├── convert-pacific-insurance.ts  # 太平洋转换核心逻辑 (502 行)
│       ├── reconciliation.ts         # 天图对账核心逻辑 (272 行)
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

纯展示型导航页，包含两个卡片链接：
- `/insurance-split` — 保单投保区间拆分
- `/pacific-convert` — 太平洋货箱清单转换

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

### 3.10 皮皮熊账单拆分 — `/pipixiong-split`

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
- ❌ **多 Sheet 支持**: 保单拆分仅处理 `"ETTON电商物流 下单模板"` Sheet
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
   - **位置**: `split-insurance.ts` 的 `generateIntervalFile()` 和 `processSplit()`

3. **共享公式 (Shared Formula) 展平**
   - exceljs 的 `spliceRows` 删除行时会破坏共享公式引用链
   - **绕过**: `flattenFormulas()` 在删除行之前将共享公式转为独立公式（保留 result）
   - **位置**: `split-insurance.ts` line 269-293

4. **内存 session 存储**
   - 拆分结果（含 Buffer）存储在 Node.js 进程内存 Map 中
   - **风险**: 大文件多用户并发可能 OOM；进程重启丢失所有 session
   - **缓解**: 30 分钟自动过期清除

5. **tailwindcss 版本锁定**
   - 项目使用 Tailwind CSS 4，语法与 v3 完全不同（无 `tailwind.config.ts`）
   - @tailwindcss/postcss 插件是必需的，缺失会导致样式完全丢失

### 待重构项

- [ ] 将 session 存储从内存 Map 改为临时文件或 Redis
- [ ] 统一四个 API 的 session 管理逻辑（目前各自维护独立的 store — split-insurance, convert-pacific, reconciliation, pipixiong-split）
- [ ] `RATES` 常量（USD=7, EUR=8）硬编码在 `split-insurance.ts`，应与太平洋转换的汇率统一
- [ ] 移除未使用的 `@anthropic-ai/sdk` 依赖（或实现实际 AI 功能）
- [ ] 添加上传进度条（当前只有旋转动画，无百分比）
- [ ] `public/output/` 下的测试文件应清理或移到 `保单拆分功能/output/`
- [ ] 对账功能的列检测规则可进一步细化（如支持更多金额列别名）

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
      → ghcr.io/etton-ai/etton-ecommerce-ai:latest + :sha
        → Sealos 控制台手动重新部署
```

---

## 附录: 与 price-system 仓库的关系

本项目原包含 FBA 比价功能（`src/app/price-query/`、`比价工具/`），已于 2026-06-29 拆分为独立仓库 [price-system](https://github.com/etton-ai/price-system)。两个仓库独立构建、独立部署：

| 项目 | 端口 | Ingress | 用途 |
|------|------|---------|------|
| ETTON TOOLS | 3000 | `deprysrenldz.cloud.sealos.io` | 保单拆分 + 太平洋转换 + 天图对账 + 皮皮熊拆分 |
| Price System | 3000 | `wlylcsujbziw.cloud.sealos.io` | FBA 多供应商比价查询 |
