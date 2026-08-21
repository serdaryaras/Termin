import { createClient } from "@supabase/supabase-js";

/** personel-izin-app Supabase (resmi_tatil_gunleri) */
function izinEnvUrl(): string {
  return (process.env.NEXT_PUBLIC_IZIN_SUPABASE_URL || "").trim();
}

function izinEnvKey(): string {
  return (process.env.NEXT_PUBLIC_IZIN_SUPABASE_ANON_KEY || "").trim();
}

export function isHolidaySourceConfigured(): boolean {
  const url = izinEnvUrl();
  const key = izinEnvKey();
  if (!url || !key) return false;
  if (url.includes("YOUR_PROJECT") || key.includes("YOUR_")) return false;
  try {
    const parsed = new URL(url);
    return parsed.protocol === "https:" && parsed.hostname.endsWith(".supabase.co");
  } catch {
    return false;
  }
}

function normalizeIsoDate(raw: unknown): string | null {
  if (raw == null) return null;
  if (typeof raw === "string") {
    const m = raw.trim().match(/^(\d{4}-\d{2}-\d{2})/);
    return m ? m[1]! : null;
  }
  return null;
}

/**
 * resmi_tatil satırlarından tam tatil günlerini toplar (arefe yarım gün iş günü sayılır).
 * Dini tatil / 29 Ekim öncesi arefe kuralı personel-izin ile uyumlu.
 */
export function holidayDatesFromRows(
  rows: Array<{ tarih?: unknown; tur?: unknown; tip?: unknown; ad?: unknown; aciklama?: unknown }>
): Set<string> {
  const full = new Set<string>();
  const half = new Set<string>();

  for (const row of rows) {
    const iso = normalizeIsoDate(row.tarih);
    if (!iso) continue;
    const tur = String(row.tur ?? row.tip ?? row.ad ?? row.aciklama ?? "")
      .toLocaleLowerCase("tr")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/ı/g, "i");
    if (/arefe|yarim|half/.test(tur)) {
      half.add(iso);
      continue;
    }
    // Varsayılan: resmi tatil / boş tur → tam tatil
    full.add(iso);
  }

  // Dini tatillerin bir gün öncesi arefe (yarım) — tam tatil listesine eklenmez
  for (const iso of [...full]) {
    const rowHint = rows.find((r) => normalizeIsoDate(r.tarih) === iso);
    const text = String(rowHint?.tur ?? rowHint?.tip ?? rowHint?.ad ?? rowHint?.aciklama ?? "")
      .toLocaleLowerCase("tr");
    if (/dini|ramazan|kurban/.test(text)) {
      half.add(addDaysIso(iso, -1));
    }
    if (iso.slice(5) === "10-29") {
      half.add(addDaysIso(iso, -1));
    }
  }

  // Gantt: yalnızca tam tatilleri atla (arefe iş günü)
  void half;
  return full;
}

function addDaysIso(iso: string, days: number): string {
  const d = new Date(`${iso}T00:00:00`);
  d.setDate(d.getDate() + days);
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

export type HolidayLoadResult = {
  dates: Set<string>;
  count: number;
  error?: string;
  source: "izin-supabase" | "empty";
};

/** personel-izin-app → resmi_tatil_gunleri */
export async function loadOfficialHolidays(): Promise<HolidayLoadResult> {
  if (!isHolidaySourceConfigured()) {
    return { dates: new Set(), count: 0, source: "empty", error: "İzin Supabase ayarı yok" };
  }
  try {
    const sb = createClient(izinEnvUrl(), izinEnvKey());
    const { data, error } = await sb.from("resmi_tatil_gunleri").select("tarih, tur");
    if (error) {
      return { dates: new Set(), count: 0, source: "empty", error: error.message };
    }
    const dates = holidayDatesFromRows((data ?? []) as Array<{ tarih?: unknown; tur?: unknown }>);
    return { dates, count: dates.size, source: "izin-supabase" };
  } catch (e) {
    return {
      dates: new Set(),
      count: 0,
      source: "empty",
      error: e instanceof Error ? e.message : "Tatil listesi okunamadı",
    };
  }
}
