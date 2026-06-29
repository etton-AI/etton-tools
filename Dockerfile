# ---- Build Stage ----
FROM node:22-alpine AS builder
WORKDIR /app

# 安装依赖
COPY package.json package-lock.json* ./
RUN npm ci

# 构建
COPY . .
RUN mkdir -p /app/public/data

# 生成价格数据 (解析 Excel → JSON)
RUN node 比价工具/price_db/parsers/build_db.js
# 复制到 public/ 使其在运行时可通过 /data/prices.json 访问
RUN cp 比价工具/price_db/data/prices.json /app/public/data/prices.json

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
