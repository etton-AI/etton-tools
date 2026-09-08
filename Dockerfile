# ---- Build Stage ----
FROM node:22-alpine AS builder
WORKDIR /app

# 安装依赖
COPY package.json package-lock.json* ./
RUN npm ci

# 构建
COPY . .
# BL_SERVICE_URL：提单/电放保函后端地址。next.config.ts 的 rewrite 在 build 时求值，
# standalone 模式下运行时不可改，故通过 build-arg 在构建期固化（本地构建不传则默认 localhost:5000）。
ARG BL_SERVICE_URL
ENV BL_SERVICE_URL=$BL_SERVICE_URL
RUN npm run build

# ---- Production Stage ----
FROM node:22-alpine AS runner
WORKDIR /app

ENV NODE_ENV=production
ENV NEXT_TELEMETRY_DISABLED=1

# 只复制 standalone 产出
COPY --from=builder /app/.next/standalone ./
COPY --from=builder /app/.next/static ./.next/static
COPY --from=builder /app/public ./public

EXPOSE 3000
CMD ["node", "server.js"]
