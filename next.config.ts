import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  // Sealos/容器部署需要 standalone 模式
  output: "standalone",
  experimental: {
    // rewrite 代理转发大文件（提单批量上传多票 PDF，请求体常超 10MB）默认 body clone 上限 10MB，
    // 超限会被 Next.js 截断请求体，转发到 Flask 时 socket hang up / ECONNRESET（后端 ClientDisconnected）。
    // 放宽到 100MB 以支持整周多票上传。
    middlewareClientMaxBodySize: "100mb",
    // rewrite 代理默认 30 秒超时（proxy-request.js: proxyTimeout || 30000），
    // generate 一次生成多票 = 每票 2 次 LibreOffice 转换（提单 + 保函预览），
    // 29 票约 58 次转换，远超 30 秒 → 代理提前断开 → socket hang up / ECONNRESET。
    // 放宽到 10 分钟（单位毫秒），让批量生成有充足时间跑完。
    proxyTimeout: 600000,
  },
  // 提单/电放保函后端（独立 Flask 服务，含 LibreOffice）
  async rewrites() {
    const blService = process.env.BL_SERVICE_URL || "http://localhost:5000";
    return [
      { source: "/api/bl/:path*", destination: `${blService}/api/bl/:path*` },
    ];
  },
};

export default nextConfig;
