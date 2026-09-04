import Link from "next/link";
import { LOGO_BASE64 } from "./logo";

export function Header() {
  return (
    <header className="sticky top-0 z-50 border-b border-line bg-white/70 backdrop-blur-md">
      <div className="relative">
        <div className="mx-auto flex h-16 max-w-6xl items-center justify-between px-4">
          <Link href="/" className="flex items-center gap-3" aria-label="返回首页">
            {/* eslint-disable-next-line @next/next/no-img-element */}
            <img
              src={`data:image/png;base64,${LOGO_BASE64}`}
              alt="易通科技物流集团"
              className="h-[42px] w-auto"
            />
          </Link>
          <div className="flex items-center gap-2">
            <span className="pulse-dot" />
            <span className="text-[13px] font-medium text-[#16a34a]">系统运行正常</span>
          </div>
        </div>
        <div className="topbar-line" />
      </div>
    </header>
  );
}

export function Footer() {
  return (
    <footer className="border-t border-line bg-white/40">
      <div className="mx-auto max-w-6xl px-4 py-11 text-center">
        <div className="barcode" />
        <p className="font-mono text-[12.5px] tracking-[.5px] text-muted">
          易通科技物流集团 · ETTON 效率提升助手 © {new Date().getFullYear()}
        </p>
      </div>
    </footer>
  );
}
