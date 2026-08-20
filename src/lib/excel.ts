import { DEFAULT_PROJECT, normalizeRole, type ParsedCapacityRow, type ParsedRow } from "./types";

function tableHtmlToTsv(html: string): string {
  if (!html || !/<table/i.test(html)) return "";
  const doc = new DOMParser().parseFromString(html, "text/html");
  return [...doc.querySelectorAll("tr")]
    .map((tr) =>
      [...tr.querySelectorAll("th,td")]
        .map((td) => (td.textContent || "").replace(/\s+/g, " ").trim())
        .join("\t")
    )
    .filter(Boolean)
    .join("\n");
}

export function clipboardToText(plain: string, html?: string): string {
  return (tableHtmlToTsv(html || "") || plain).replace(/\r\n/g, "\n").trim();
}

function normalizeHeader(cell: string): string {
  return String(cell || "")
    .toLowerCase()
    .replaceAll("ı", "i")
    .replaceAll("İ", "i")
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

type Kind = "name" | "project" | "role" | "hours" | "people";

function headerKind(cell: string): Kind | null {
  const h = normalizeHeader(cell);
  if (!h) return null;
  if (/(saat|sure|hour|duration|manhour|mh)/.test(h) || h === "sa") return "hours";
  if (/(kisi|kişi|adet|crew size|people|manpower|personel say)/.test(h)) return "people";
  if (/(proje|project|ship|gemi|blok|siparis|sipariş)/.test(h)) return "project";
  if (/(personel|cins|meslek|unvan|role|crew|staff|trade|donatim|donatım|konstruks)/.test(h)) return "role";
  if (/(is|kalem|ad|tanim|job|task|aktivite|faaliyet|name)/.test(h)) return "name";
  return null;
}

function parseHours(raw: string | undefined): number {
  if (raw == null) return NaN;
  let s = String(raw).trim().replace(/\s*(sa|saat|h|hr|hrs)\s*$/i, "");
  s = s.replace(/\s/g, "").replace(",", ".");
  const n = Number(s);
  return n > 0 ? n : NaN;
}

function parsePeople(raw: string | undefined): number {
  if (raw == null || !String(raw).trim()) return 1;
  const n = Number(String(raw).trim().replace(",", "."));
  return n > 0 ? n : 1;
}

function splitClipboardLine(line: string): string[] {
  if (line.includes("\t")) return line.split("\t").map((c) => c.trim());
  if (line.includes(";")) return line.split(";").map((c) => c.trim());
  if (line.includes(",")) {
    const parts = line.split(",").map((c) => c.trim());
    if (parts.length >= 2) return parts;
  }
  return [line.trim()];
}

export function parseExcelText(text: string): { rows: ParsedRow[]; skipped: number } {
  const lines = String(text)
    .replace(/^\uFEFF/, "")
    .split(/\r\n|\n|\r/)
    .map((l) => l.replace(/\u00a0/g, " ").trimEnd())
    .filter((l) => l.trim());

  if (!lines.length) return { rows: [], skipped: 0 };

  let map = { name: 0, project: -1, role: 1, hours: 2, people: -1 };
  let start = 0;
  const first = splitClipboardLine(lines[0]);
  const kinds = first.map(headerKind);
  if (kinds.some(Boolean) && kinds.filter(Boolean).length >= 2) {
    map = { name: -1, project: -1, role: -1, hours: -1, people: -1 };
    kinds.forEach((k, i) => {
      if (k && map[k] === -1) map[k] = i;
    });
    start = 1;
  } else if (first.length >= 4) {
    map = { project: 0, name: 1, role: 2, hours: 3, people: -1 };
  } else if (first.length === 3) {
    // Proje | kalem | saat  veya  kalem | personel | saat
    const c0 = first[0] || "";
    const c1 = first[1] || "";
    if (/^\d+[./]/.test(c0) || /simonsen|proje|project/i.test(c0)) {
      map = { project: 0, name: 1, role: -1, hours: 2, people: -1 };
    } else if (headerKind(c1) === "role" || /donat|konstruk|kaynak|montaj/i.test(c1)) {
      map = { name: 0, project: -1, role: 1, hours: 2, people: -1 };
    } else {
      map = { project: 0, name: 1, role: -1, hours: 2, people: -1 };
    }
  } else if (first.length === 2) {
    map = { name: 0, project: -1, role: -1, hours: 1, people: -1 };
  } else if (first.length === 1) {
    map = { name: 0, project: -1, role: -1, hours: -1, people: -1 };
  }

  const rows: ParsedRow[] = [];
  let skipped = 0;
  for (let i = start; i < lines.length; i++) {
    const cols = splitClipboardLine(lines[i]);
    const name = (map.name >= 0 ? cols[map.name] : cols[0] || "").trim();
    const project = (map.project >= 0 ? cols[map.project] : "").trim();
    const role = (map.role >= 0 ? cols[map.role] : "").trim();
    let hours = map.hours >= 0 ? parseHours(cols[map.hours]) : NaN;
    if (!Number.isFinite(hours)) {
      const used = new Set([map.name, map.project, map.role, map.people].filter((x) => x >= 0));
      const numeric = cols
        .map((c, idx) => ({ idx, n: parseHours(c) }))
        .find((x) => Number.isFinite(x.n) && !used.has(x.idx));
      hours = numeric ? numeric.n : NaN;
    }
    const people = map.people >= 0 ? parsePeople(cols[map.people]) : 1;
    if (!name || !Number.isFinite(hours)) {
      skipped += 1;
      continue;
    }
    rows.push({
      name,
      project: project || DEFAULT_PROJECT,
      role: normalizeRole(role),
      hours,
      people,
    });
  }
  return { rows, skipped };
}

type CapKind = "project" | "year" | "week" | "role" | "people";

function capacityHeaderKind(cell: string): CapKind | null {
  const h = normalizeHeader(cell);
  if (!h) return null;
  if (/(yil|yıl|year)/.test(h)) return "year";
  if (/(hafta|week|hf)/.test(h) && !/(kisi|kişi|people)/.test(h)) return "week";
  if (/(kisi|kişi|adet|sayi|sayı|people|kapasite|capacity|manpower)/.test(h) && !/(tip|cins|tur|tür)/.test(h)) {
    return "people";
  }
  if (/(personel tip|personel cins|meslek|unvan|role|crew|trade|tip|cins|donatim|donatım|konstruks)/.test(h)) {
    return "role";
  }
  if (/(personel|staff)/.test(h) && !/(say|kisi|kişi)/.test(h)) return "role";
  if (/(proje|project)/.test(h)) return "project";
  return null;
}

function parseWeekToken(raw: string): { year?: number; week: number } | null {
  const s = String(raw)
    .trim()
    .replace(/\u00a0/g, " ")
    .replace(",", "."); // TR Excel: 2026,35 → 2026.35
  // 2026.35 | 2026.4 (Excel’in 2026.40 kısaltması) | 2026-35
  const dotted = s.match(/^(\d{4})[./-](\d{1,2})$/);
  if (dotted) {
    const year = Number(dotted[1]);
    let week = Number(dotted[2]);
    const frac = dotted[2];
    if (frac.length === 1 && week >= 1 && week <= 5) week = week * 10;
    if (week >= 1 && week <= 53) return { year, week };
  }
  // Tek hücrede sayısal 2026.35
  const asNum = Number(s);
  if (Number.isFinite(asNum) && asNum >= 2000) {
    const year = Math.floor(asNum);
    let frac = Math.round((asNum - year) * 100);
    // 2026.4 → JS 2026.4 → frac 40; 2026.35 → 35
    if (frac >= 1 && frac <= 53) return { year, week: frac };
    // 2026.4 okunurken 2026.40 kaybı: *10 ile gelmiş olabilir
    const frac1 = Math.round((asNum - year) * 10);
    if (frac1 >= 1 && frac1 <= 5) return { year, week: frac1 * 10 };
  }
  if (Number.isFinite(asNum) && asNum >= 1 && asNum <= 53) return { week: Math.round(asNum) };
  return null;
}

function parseCapacityPeople(raw: string | undefined): number {
  if (raw == null) return NaN;
  const s = String(raw)
    .trim()
    .replace(/\s*kişi\s*$/i, "")
    .replace(/\s*kisi\s*$/i, "")
    .replace(/\s*adet\s*$/i, "")
    .replace(/\s/g, "")
    .replace(",", ".");
  const n = Number(s);
  return Number.isFinite(n) && n >= 0 ? n : NaN;
}

function splitCapacityLine(line: string): string[] {
  const t = line.replace(/\u00a0/g, " ").trim();
  if (t.includes("\t")) return t.split("\t").map((c) => c.trim());
  if (/[·•|]/.test(t)) {
    return t
      .split(/\s*[·•|]\s*/)
      .map((c) => c.trim())
      .filter(Boolean);
  }
  if (t.includes(";")) return t.split(";").map((c) => c.trim());
  // Excel bazen sekmeyi boşluğa çevirir: proje hafta tip kişi
  const spaced = t.split(/\s{2,}|\s+/).map((c) => c.trim()).filter(Boolean);
  if (spaced.length >= 4) {
    // "252.Simonsen 2026.35 Donatım 3"
    return [spaced[0], spaced[1], spaced.slice(2, -1).join(" "), spaced[spaced.length - 1]];
  }
  if (spaced.length === 3) return spaced;
  return splitClipboardLine(t);
}

/**
 * Excel: proje adı | hafta | personel tipi | kişi sayısı
 * Örn: 252.Simonsen | 2026-35 | Donatım | 3
 * (Eski biçimler 2026.35 / 2026,35 de kabul edilir.)
 */
export function parseCapacityText(
  text: string,
  defaultYear: number
): { rows: ParsedCapacityRow[]; skipped: number } {
  const lines = String(text)
    .replace(/^\uFEFF/, "")
    .split(/\r\n|\n|\r/)
    .map((l) => l.replace(/\u00a0/g, " ").trimEnd())
    .filter((l) => l.trim());

  if (!lines.length) return { rows: [], skipped: 0 };

  let map = { project: 0, year: -1, week: 1, role: 2, people: 3 };
  let start = 0;
  const first = splitCapacityLine(lines[0]);
  const kinds = first.map(capacityHeaderKind);
  if (kinds.filter(Boolean).length >= 2) {
    map = { project: -1, year: -1, week: -1, role: -1, people: -1 };
    kinds.forEach((k, i) => {
      if (k && map[k] === -1) map[k] = i;
    });
    start = 1;
  } else if (first.length >= 4) {
    map = { project: 0, year: -1, week: 1, role: 2, people: 3 };
  } else if (first.length === 3) {
    map = { project: 0, year: -1, week: 1, role: -1, people: 2 };
  }

  const rows: ParsedCapacityRow[] = [];
  let skipped = 0;
  for (let i = start; i < lines.length; i++) {
    const cols = splitCapacityLine(lines[i]);
    const project = (map.project >= 0 ? cols[map.project] : "").trim() || DEFAULT_PROJECT;
    const role = normalizeRole(map.role >= 0 ? cols[map.role] : "");
    const people = parseCapacityPeople(map.people >= 0 ? cols[map.people] : "");
    const weekRaw = map.week >= 0 ? cols[map.week] : "";
    const weekTok = parseWeekToken(weekRaw || "");
    let year = map.year >= 0 ? Number(cols[map.year]) : defaultYear;
    if (weekTok?.year) year = weekTok.year;
    const week = weekTok?.week ?? NaN;
    if (
      !project ||
      role === "Belirtilmedi" ||
      !Number.isFinite(year) ||
      !Number.isFinite(week) ||
      !Number.isFinite(people)
    ) {
      skipped += 1;
      continue;
    }
    rows.push({ project, year, week, role, people });
  }
  return { rows, skipped };
}

export function formatCapacityChip(c: {
  project: string;
  year: number;
  week: number;
  role: string;
  people: number;
}): string {
  const ww = String(c.week).padStart(2, "0");
  return `${c.project} · ${c.year}-${ww} · ${c.role} · ${c.people} kişi`;
}
