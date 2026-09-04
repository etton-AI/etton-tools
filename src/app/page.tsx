"use client";

import Link from "next/link";
import type { CSSProperties, MouseEvent, ReactNode } from "react";

interface Tool {
  href: string;
  title: string;
  desc: string;
  badge: string;
  icon: ReactNode;
}

const TOOLS: Tool[] = [
  {
    href: "/insurance-split",
    title: "保单投保区间拆分",
    badge: "高频",
    desc: "按每箱申报价值自动拆分为 5 个投保区间文件",
    icon: (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2} strokeLinecap="round" strokeLinejoin="round">
        <path d="M20 13c0 5-3.5 7.5-7.66 8.95a1 1 0 0 1-.67-.01C7.5 20.5 4 18 4 13V6a1 1 0 0 1 1-1c2 0 4.5-1.2 6.24-2.72a1.17 1.17 0 0 1 1.52 0C14.51 3.81 17 5 19 5a1 1 0 0 1 1 1z" />
        <path d="m9 12 2 2 4-4" />
      </svg>
    ),
  },
  {
    href: "/pacific-convert",
    title: "太平洋货箱清单转换",
    badge: "高效",
    desc: "货箱清单转太平洋投保清单，按单箱货值拆分",
    icon: (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2} strokeLinecap="round" strokeLinejoin="round">
        <path d="M2 21c.6.5 1.2 1 2.5 1 2.5 0 2.5-2 5-2 1.3 0 1.9.5 2.5 1 .6.5 1.2 1 2.5 1 2.5 0 2.5-2 5-2 1.3 0 1.9.5 2.5 1" />
        <path d="M19.38 20A11.6 11.6 0 0 0 21 14l-9-4-9 4c0 2.9.94 5.34 2.81 7.76" />
        <path d="M19 13V7a2 2 0 0 0-2-2H7a2 2 0 0 0-2 2v6" />
        <path d="M12 10v4" />
        <path d="M12 2v3" />
      </svg>
    ),
  },
  {
    href: "/multi-supplier-reconciliation",
    title: "多供应商对账引擎",
    badge: "常用",
    desc: "自动识别 16 家供应商，按 SO 号比对金额差异",
    icon: (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2} strokeLinecap="round" strokeLinejoin="round">
        <path d="m16 16 3-8 3 8c-.87.65-1.92 1-3 1s-2.13-.35-3-1Z" />
        <path d="m2 16 3-8 3 8c-.87.65-1.92 1-3 1s-2.13-.35-3-1Z" />
        <path d="M7 21h10" />
        <path d="M12 3v18" />
        <path d="M3 7h2c2 0 5-1 7-2 2 1 5 2 7 2h2" />
      </svg>
    ),
  },
  {
    href: "/pipixiong-split",
    title: "皮皮熊账单拆分",
    badge: "高频",
    desc: "合并账单拆分为国内 / 国外 / INVOICE 三份",
    icon: (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2} strokeLinecap="round" strokeLinejoin="round">
        <circle cx="6" cy="6" r="3" />
        <path d="M8.12 8.12 12 12" />
        <path d="M20 4 8.12 15.88" />
        <circle cx="6" cy="18" r="3" />
        <path d="M14.8 14.8 20 20" />
      </svg>
    ),
  },
  {
    href: "/yanxun-convert",
    title: "延讯下单优化",
    badge: "高效",
    desc: "延讯发票批量转易通下单模板",
    icon: (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2} strokeLinecap="round" strokeLinejoin="round">
        <path d="m16 3 4 4-4 4" />
        <path d="M20 7H4" />
        <path d="m8 21-4-4 4-4" />
        <path d="M4 17h16" />
      </svg>
    ),
  },
  {
    href: "/warehouse-entry",
    title: "TR入仓数据整理",
    badge: "常用",
    desc: "客户 + 供应商数据自动匹配，生成建议箱规",
    icon: (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2} strokeLinecap="round" strokeLinejoin="round">
        <path d="m7.5 4.27 9 5.15" />
        <path d="M21 8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16Z" />
        <path d="m3.3 7 8.7 5 8.7-5" />
        <path d="M12 22V12" />
      </svg>
    ),
  },
];

export default function HomePage() {
  const handleMouseMove = (e: MouseEvent<HTMLAnchorElement>) => {
    const card = e.currentTarget;
    const r = card.getBoundingClientRect();
    card.style.setProperty("--mx", `${e.clientX - r.left}px`);
    card.style.setProperty("--my", `${e.clientY - r.top}px`);
  };

  return (
    <div>
      <section className="pb-14 pt-[76px] text-center">
        <p className="mb-[18px] text-xs font-bold uppercase tracking-[3px] text-primary">
          ETTON TECHNOLOGY LOGISTICS GROUP
        </p>
        <h1 className="text-[32px] font-extrabold leading-tight tracking-wide text-deep sm:text-[46px]">
          ETTON 效率提升助手
        </h1>
        <div className="scanline" />
        <p className="mt-6 text-[17px] tracking-[.5px] text-muted">
          六款电商效率工具，把繁琐的表格处理变成一次点击。
        </p>
      </section>

      <section className="grid grid-cols-1 gap-6 pb-20 md:grid-cols-2 lg:grid-cols-3">
        {TOOLS.map((t, i) => (
          <Link
            key={t.href}
            href={t.href}
            onMouseMove={handleMouseMove}
            className="brand-card"
            style={{ "--i": i } as unknown as CSSProperties}
          >
            <span className="card-glow" />
            <div className="card-top">
              <span className="icon-box">{t.icon}</span>
              <span className="arrow">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2} strokeLinecap="round" strokeLinejoin="round">
                  <path d="M5 12h14" />
                  <path d="m12 5 7 7-7 7" />
                </svg>
              </span>
            </div>
            <div className="card-body">
              <div className="card-title-row">
                <h3 className="card-title">{t.title}</h3>
                <span className="card-badge">{t.badge}</span>
              </div>
              <p className="card-desc">{t.desc}</p>
            </div>
          </Link>
        ))}
      </section>
    </div>
  );
}
