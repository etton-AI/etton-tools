import type { Metadata } from "next";
import { Header, Footer } from "@/components/Header";
import "./globals.css";

export const metadata: Metadata = {
  title: "ETTON效率提升助手",
  description: "ETTON 电商效率提升工具集",
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="zh-CN">
      <body className="min-h-screen antialiased">
        <div className="grid-bg" />
        <Header />
        <main className="relative z-[1] mx-auto min-h-[calc(100vh-8rem)] max-w-6xl px-4 py-8">
          {children}
        </main>
        <Footer />
      </body>
    </html>
  );
}
