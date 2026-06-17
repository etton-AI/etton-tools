export function Header() {
  return (
    <header className="sticky top-0 z-50 border-b border-zinc-200 bg-white/80 backdrop-blur-md">
      <div className="mx-auto flex h-16 max-w-6xl items-center px-4">
        <span className="text-xl font-bold text-zinc-900">ETTON 效率提升助手</span>
      </div>
    </header>
  );
}

export function Footer() {
  return (
    <footer className="border-t border-zinc-200 bg-zinc-50">
      <div className="mx-auto max-w-6xl px-4 py-8 text-center text-sm text-zinc-400">
        © {new Date().getFullYear()}
      </div>
    </footer>
  );
}
