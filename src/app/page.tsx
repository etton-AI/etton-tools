import Link from "next/link";

export default function HomePage() {
  return (
    <div className="flex min-h-[60vh] flex-col items-center justify-center text-center">
      <h1 className="text-4xl font-bold tracking-tight text-zinc-800">
        ETTON 效率提升助手
      </h1>
      <p className="mt-4 text-lg text-zinc-500">
        电商效率提升工具集
      </p>
      <div className="mt-8 flex flex-wrap justify-center gap-4">
        <Link
          href="/insurance-split"
          className="rounded-lg bg-zinc-800 px-5 py-2.5 text-sm font-medium text-white shadow-sm hover:bg-zinc-700 transition-colors"
        >
          📦 保单投保区间拆分
        </Link>
        <Link
          href="/pacific-convert"
          className="rounded-lg bg-blue-600 px-5 py-2.5 text-sm font-medium text-white shadow-sm hover:bg-blue-500 transition-colors"
        >
          🚢 太平洋货箱清单转换
        </Link>
        <Link
          href="/multi-supplier-reconciliation"
          className="rounded-lg bg-violet-600 px-5 py-2.5 text-sm font-medium text-white shadow-sm hover:bg-violet-500 transition-colors"
        >
          🔄 多供应商对账引擎
        </Link>
        <Link
          href="/pipixiong-split"
          className="rounded-lg bg-fuchsia-600 px-5 py-2.5 text-sm font-medium text-white shadow-sm hover:bg-fuchsia-500 transition-colors"
        >
          🧾 皮皮熊账单拆分
        </Link>
      </div>
    </div>
  );
}
