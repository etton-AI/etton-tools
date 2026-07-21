# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 项目简介
ETTON 电商 AI 平台 — 基于 Next.js + Claude API 的智能电商演示项目。

## 技术栈
- **框架**: Next.js 15.3 (App Router) + React 19 + TypeScript 5.8
- **样式**: Tailwind CSS 4 (PostCSS `@tailwindcss/postcss` 插件)
- **AI**: Anthropic Claude API (`@anthropic-ai/sdk` v0.74)
- **Lint**: ESLint 9 flat config (`next/core-web-vitals` + `next/typescript`)
- **路径别名**: `@/*` → `./src/*`

## 快速开始

```powershell
# 1. 安装依赖
npm install

# 2. 配置 API Key
copy .env.example .env.local
# 编辑 .env.local，填入 ANTHROPIC_API_KEY

# 3. 启动开发服务器
npm run dev
```

浏览器访问 http://localhost:3000

## 常用命令

| 命令 | 说明 |
|------|------|
| `npm run dev` | 启动开发服务器 |
| `npm run build` | 构建生产版本 |
| `npm run start` | 运行生产版本 |
| `npm run lint` | ESLint 代码检查 |

## 项目结构

```
src/
├── app/
│   ├── layout.tsx         # 根布局（Header + Footer + 全局样式）
│   ├── page.tsx           # 首页（当前为 Hello World 占位）
│   └── globals.css        # Tailwind CSS 4 入口（@import "tailwindcss"）
├── components/
│   └── Header.tsx         # Header + Footer 组件（同文件导出）
├── lib/                   # 业务逻辑（待创建）
└── types/                 # TypeScript 类型（待创建）
```

路由遵循 Next.js App Router 约定。计划中的路由：`/products/`（商品）、`/assistant/`（AI 购物助手）、`/api/chat/`（Claude 对话 API）。

## 关键配置

- **Tailwind CSS 4**: 使用 `@import "tailwindcss"` 入口（无需 `tailwind.config.ts`），PostCSS 插件 `@tailwindcss/postcss` 处理编译。
- **路径别名**: `tsconfig.json` 中配置了 `@/*` → `./src/*`，导入时使用 `@/components/Header` 等方式。
- **环境变量**: `.env.local`（git-ignored）中的 `ANTHROPIC_API_KEY` 通过 `process.env` 访问。只有以 `NEXT_PUBLIC_` 为前缀的变量可在客户端使用。

## 开发约定
- 使用中文与用户沟通
- 代码注释和变量名优先使用英文
- 提交前确保 `npm run build` 通过
# 项目规则
每次完成功能开发或修改后，必须：
1. 更新 docs/dev-spec.md 的「已知坑 / 非目标」章节
2. 更新 specs/ 下对应功能文档
