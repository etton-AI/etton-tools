import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  // Sealos/容器部署需要 standalone 模式
  output: "standalone",
};

export default nextConfig;
