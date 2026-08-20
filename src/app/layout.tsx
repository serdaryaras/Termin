import type { Metadata } from "next";
import { ArtiLogo } from "@/components/ArtiLogo";
import "./globals.css";

export const metadata: Metadata = {
  title: "ARTI İş Sıralama ve Gantt",
  description: "Öncelik sırası, paralel yollar ve Gantt çizelgesi",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="tr">
      <body className="antialiased">
        <header className="border-b border-[var(--card-border)] bg-[var(--card)]">
          <div className="mx-auto flex w-full max-w-[1920px] items-center justify-between px-4 py-4">
            <ArtiLogo />
            <span className="text-sm text-[var(--muted)]">Next.js · Supabase · Vercel</span>
          </div>
        </header>
        <main className="mx-auto w-full max-w-[1920px] px-4 py-8">{children}</main>
      </body>
    </html>
  );
}
