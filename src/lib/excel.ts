import {
  capacityKey,
  DEFAULT_PROJECT,
  normalizeRole,
  ROLE_OPTIONS,
  type ParsedCapacityRow,
  type ParsedJobImport,
  type ParsedRow,
} from "./types";

const ROLE_OPTIONS_SET = new Set<string>(ROLE_OPTIONS);

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

export function looksLikeActivityCode(raw: string): boolean {
  return /^\d+\.\d+\.\d+/.test(String(raw || "").trim());
}

/** 252.Simonsen gibi kısa proje adı (çok parçalı WBS değil) */
function looksLikeProjectName(raw: string): boolean {
  const t = String(raw || "").trim();
  if (!t || looksLikeActivityCode(t)) return false;
  if (/simonsen|proje|project/i.test(t)) return true;
  if (/^\d+[./]\s*[A-Za-zÀ-ÿğüşıöçĞÜŞİÖÇ]/.test(t)) return true;
  return false;
}

function looksLikeRoleCell(raw: string): boolean {
  const t = String(raw || "").trim();
  if (!t) return false;
  if (headerKind(t) === "role") return true;
  const n = normalizeRole(t);
  if (n !== "Belirtilmedi" && ROLE_OPTIONS_SET.has(n)) return true;
  const key = t
    .toLocaleLowerCase("tr")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/ı/g, "i");
  return /^(donat|konstruk|kaynak|montaj|boru|elektrik|ressam|vinc|tekniker|muhendis|yardimci)/.test(key);
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

/** Hücre: `Donatım:25` → rol + saat */
export function parseRoleHoursCell(raw: string): { role: string; hours: number } | null {
  const t = String(raw || "").trim();
  if (!t) return null;
  const m = t.match(/^(.+?)\s*[:：]\s*([\d.,]+)\s*(?:sa|saat|h|hr|hrs)?\s*$/i);
  if (!m) return null;
  const hours = parseHours(m[2]);
  if (!Number.isFinite(hours)) return null;
  const role = normalizeRole(m[1]!.trim());
  if (!role || role === "Belirtilmedi") return null;
  return { role, hours };
}

/** `252.100.104-Fire Integrity…` → kod + başlık */
export function splitDrawingLabel(raw: string): { code: string; title: string } | null {
  const t = String(raw || "").trim();
  const m = t.match(/^(\d+(?:\.\d+)+)\s*[-–—]\s*(.+)$/);
  if (!m) return null;
  const code = m[1]!.trim();
  const title = m[2]!.trim();
  if (!code || !title) return null;
  return { code, title };
}

function looksLikeCategoryHeader(raw: string): boolean {
  const h = normalizeHeader(raw);
  if (!h) return false;
  if (parseRoleHoursCell(raw)) return false;
  if (looksLikeActivityCode(raw)) return false;
  // Class, 3D Model, ISO, Arrang'nt, Works'p, Manual, …
  if (
    /^(class|sinif|sınıf|3d|model|iso|arrang|arrangement|duzen|düzen|work|workshop|atolye|atölye|manual|kilavuz|kılavuz)/.test(
      h
    )
  ) {
    return true;
  }
  // Kısa başlık; proje/resim sütun etiketi değil
  if (/^(proje|project|resim|cizim|çizim|drawing|job|is|iş)/.test(h)) return false;
  return h.length <= 24 && !/^\d/.test(h);
}

function sanitizeCategoryToken(header: string): string {
  return String(header || "")
    .trim()
    .replace(/\s+/g, " ");
}

function buildActivityName(code: string, category: string, title: string): string {
  return `${code}.${sanitizeCategoryToken(category)}-${title}`;
}

/**
 * Başlık satırı: [Proje Adı] | Class | 3D Model | ISO | …
 * İlk hücre proje adı (veya "Proje Adı" etiketi → proje sonradan doldurulur).
 */
function detectMatrixHeader(cols: string[]): { project: string; categories: string[] } | null {
  if (cols.length < 2) return null;
  const cats = cols.slice(1);
  if (!cats.length) return null;
  const asHeaders = cats.filter((c) => c && looksLikeCategoryHeader(c));
  const asCells = cats.filter((c) => parseRoleHoursCell(c));
  if (asHeaders.length < 1 || asCells.length > 0) return null;

  let project = (cols[0] || "").trim();
  const h = normalizeHeader(project);
  // Sütun başlığı yazılmışsa (Proje Adı / Resim …) — proje değeri değil
  if (
    project &&
    /^(proje|project|resim|cizim|çizim|drawing|job|is adi|is ad|iş adi|iş ad)/.test(h) &&
    !looksLikeProjectName(project)
  ) {
    project = "";
  }

  return {
    project,
    categories: cats.map((c, i) => sanitizeCategoryToken(c) || `Kategori${i + 1}`),
  };
}

/** Veri satırı: resim adı | Donatım:25 | … */
function looksLikeMatrixDataRow(cols: string[]): boolean {
  if (cols.length < 2) return false;
  if (!splitDrawingLabel(cols[0] || "") && !looksLikeActivityCode(cols[0] || "")) return false;
  return cols.slice(1).some((c) => parseRoleHoursCell(c));
}

function parseMatrixExcelText(lines: string[]): ParsedJobImport {
  const first = splitClipboardLine(lines[0]!);
  const header = detectMatrixHeader(first);
  let projectFromHeader = "";
  let categoryHeaders: string[] = [];
  let start = 0;

  if (header) {
    projectFromHeader = header.project;
    categoryHeaders = header.categories;
    start = 1;
  } else {
    const sample = lines.map(splitClipboardLine).find(looksLikeMatrixDataRow) || first;
    const n = Math.max(0, sample.length - 1);
    categoryHeaders = Array.from({ length: n }, (_, i) => `Kategori${i + 1}`);
  }

  const rows: ParsedRow[] = [];
  const dependencies: ParsedJobImport["dependencies"] = [];
  let skipped = 0;

  for (let i = start; i < lines.length; i++) {
    const cols = splitClipboardLine(lines[i]!);
    if (!cols.some((c) => c.trim())) {
      skipped += 1;
      continue;
    }
    const drawingRaw = (cols[0] || "").trim();
    const split = splitDrawingLabel(drawingRaw);
    if (!split) {
      skipped += 1;
      continue;
    }

    const project = projectFromHeader || DEFAULT_PROJECT;
    const chain: ParsedRow[] = [];
    for (let c = 0; c < categoryHeaders.length; c++) {
      const cell = cols[c + 1];
      const parsed = parseRoleHoursCell(cell || "");
      if (!parsed) continue;
      const category = categoryHeaders[c] || `Kategori${c + 1}`;
      chain.push({
        name: buildActivityName(split.code, category, split.title),
        project,
        role: parsed.role,
        hours: parsed.hours,
        people: 1,
      });
    }

    if (!chain.length) {
      skipped += 1;
      continue;
    }

    for (const row of chain) rows.push(row);
    for (let k = 0; k < chain.length - 1; k++) {
      const pred = chain[k]!;
      const succ = chain[k + 1]!;
      dependencies.push({
        predecessor: { project: pred.project, name: pred.name },
        successor: { project: succ.project, name: succ.name },
      });
    }
  }

  return { rows, dependencies, skipped, format: "matrix" };
}

function isMatrixClipboard(lines: string[]): boolean {
  if (!lines.length) return false;
  const first = splitClipboardLine(lines[0]!);
  if (detectMatrixHeader(first)) return true;
  let hits = 0;
  for (let i = 0; i < Math.min(lines.length, 12); i++) {
    if (looksLikeMatrixDataRow(splitClipboardLine(lines[i]!))) hits += 1;
  }
  return hits >= 1 && first.length >= 2;
}

function parseFlatExcelText(lines: string[]): ParsedJobImport {
  let map = { name: 0, project: -1, role: 1, hours: 2, people: -1 };
  let start = 0;
  const first = splitClipboardLine(lines[0]!);
  const kinds = first.map(headerKind);
  if (kinds.some(Boolean) && kinds.filter(Boolean).length >= 2) {
    map = { name: -1, project: -1, role: -1, hours: -1, people: -1 };
    kinds.forEach((k, i) => {
      if (k && map[k] === -1) map[k] = i;
    });
    start = 1;
  } else if (first.length >= 4) {
    const c0 = first[0] || "";
    const c1 = first[1] || "";
    const c2 = first[2] || "";
    if (looksLikeActivityCode(c0) && (looksLikeRoleCell(c1) || !looksLikeProjectName(c0))) {
      map = { name: 0, project: -1, role: 1, hours: 2, people: 3 };
    } else if (looksLikeProjectName(c0) || (!looksLikeActivityCode(c0) && looksLikeRoleCell(c2))) {
      map = { project: 0, name: 1, role: 2, hours: 3, people: first.length >= 5 ? 4 : -1 };
    } else if (looksLikeRoleCell(c1) && Number.isFinite(parseHours(c2))) {
      map = { name: 0, project: -1, role: 1, hours: 2, people: 3 };
    } else {
      map = { project: 0, name: 1, role: 2, hours: 3, people: -1 };
    }
  } else if (first.length === 3) {
    const c0 = first[0] || "";
    const c1 = first[1] || "";
    if (looksLikeActivityCode(c0) || looksLikeRoleCell(c1)) {
      map = { name: 0, project: -1, role: 1, hours: 2, people: -1 };
    } else if (looksLikeProjectName(c0)) {
      map = { project: 0, name: 1, role: -1, hours: 2, people: -1 };
    } else {
      map = { name: 0, project: -1, role: 1, hours: 2, people: -1 };
    }
  } else if (first.length === 2) {
    map = { name: 0, project: -1, role: -1, hours: 1, people: -1 };
  } else if (first.length === 1) {
    map = { name: 0, project: -1, role: -1, hours: -1, people: -1 };
  }

  const rows: ParsedRow[] = [];
  let skipped = 0;
  for (let i = start; i < lines.length; i++) {
    const cols = splitClipboardLine(lines[i]!);
    let name = (map.name >= 0 ? cols[map.name] : cols[0] || "").trim();
    let project = (map.project >= 0 ? cols[map.project] : "").trim();
    let role = (map.role >= 0 ? cols[map.role] : "").trim();
    if (looksLikeActivityCode(project) && !looksLikeActivityCode(name)) {
      const swap = name;
      name = project;
      project = swap;
    }
    if (looksLikeActivityCode(project) && looksLikeRoleCell(name)) {
      role = role || name;
      name = project;
      project = "";
    }
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
  return { rows, dependencies: [], skipped, format: "flat" };
}

/**
 * Excel / panodan iş listesi.
 * Matris: satır1 = ProjeAdı | Class | 3D Model | ISO | …
 *         alt satırlar = resim (kod-ad) | Donatım:25 | …
 */
export function parseExcelText(text: string): ParsedJobImport {
  const lines = String(text)
    .replace(/^\uFEFF/, "")
    .split(/\r\n|\n|\r/)
    .map((l) => l.replace(/\u00a0/g, " ").trimEnd())
    .filter((l) => l.trim());

  if (!lines.length) return { rows: [], dependencies: [], skipped: 0, format: "flat" };

  if (isMatrixClipboard(lines)) return parseMatrixExcelText(lines);
  return parseFlatExcelText(lines);
}

/** .xlsx / .xls / .csv dosyasını okuyup TSV metne çevirir (ilk sayfa) */
export async function readSpreadsheetFileToTsv(file: File): Promise<{ text: string; sheetName: string }> {
  const name = (file.name || "").toLowerCase();
  if (name.endsWith(".csv") || name.endsWith(".tsv") || name.endsWith(".txt")) {
    const text = await file.text();
    return { text: text.replace(/\r\n/g, "\n").trim(), sheetName: file.name };
  }

  const XLSX = await import("xlsx");
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array", cellDates: false, raw: false });
  const sheetName = wb.SheetNames[0] || "";
  if (!sheetName) throw new Error("Excel dosyasında sayfa bulunamadı.");
  const ws = wb.Sheets[sheetName];
  if (!ws) throw new Error(`Sayfa okunamadı: ${sheetName}`);

  // Ham 2D dizi — boş hücreleri koru (kategori boşlukları önemli)
  const rows = XLSX.utils.sheet_to_json<(string | number | boolean | null)[]>(ws, {
    header: 1,
    defval: "",
    blankrows: false,
    raw: false,
  }) as unknown as unknown[][];

  const text = rows
    .map((row) => {
      const cells = Array.isArray(row) ? row : [];
      return cells
        .map((c) => String(c ?? "").replace(/\s+/g, " ").trim())
        .join("\t");
    })
    .filter((line) => line.replace(/\t/g, "").trim())
    .join("\n");

  return { text, sheetName };
}

/** Dosyadan iş listesi matrisini parse et */
export async function parseJobListFile(file: File): Promise<ParsedJobImport & { sheetName: string }> {
  const { text, sheetName } = await readSpreadsheetFileToTsv(file);
  const parsed = parseExcelText(text);
  return { ...parsed, sheetName };
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
    const frac = Math.round((asNum - year) * 100);
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
 * Matris kapasite: satır1 = ProjeAdı | Donatım | Hull | …
 * alt satırlar = 2026-35 | 5 | 2 | …
 */
function isCapacityMatrix(lines: string[]): boolean {
  if (lines.length < 2) return false;
  const header = splitClipboardLine(lines[0]!);
  if (header.length < 2) return false;
  const roles = header.slice(1).filter((c) => c.trim());
  if (!roles.length) return false;
  // Başlıkta rol adları var, Role:saat veya hafta yok
  if (roles.some((c) => parseRoleHoursCell(c) || parseWeekToken(c))) return false;
  // En az bir veri satırı: hafta + sayı
  for (let i = 1; i < Math.min(lines.length, 8); i++) {
    const cols = splitClipboardLine(lines[i]!);
    if (parseWeekToken(cols[0] || "") && cols.slice(1).some((c) => Number.isFinite(parseCapacityPeople(c)))) {
      return true;
    }
  }
  return false;
}

function parseCapacityMatrixText(
  lines: string[],
  defaultYear: number
): { rows: ParsedCapacityRow[]; skipped: number; duplicates: number } {
  const header = splitClipboardLine(lines[0]!);
  let project = (header[0] || "").trim();
  const h = normalizeHeader(project);
  if (
    project &&
    /^(proje|project|hafta|week)/.test(h) &&
    !looksLikeProjectName(project)
  ) {
    project = "";
  }
  project = project || DEFAULT_PROJECT;

  const roles = header.slice(1).map((c) => normalizeRole(c.trim()));
  const rows: ParsedCapacityRow[] = [];
  let skipped = 0;

  for (let i = 1; i < lines.length; i++) {
    const cols = splitClipboardLine(lines[i]!);
    if (!cols.some((c) => c.trim())) {
      skipped += 1;
      continue;
    }
    const weekTok = parseWeekToken(cols[0] || "");
    if (!weekTok) {
      skipped += 1;
      continue;
    }
    const year = weekTok.year ?? defaultYear;
    const week = weekTok.week;
    let any = false;
    for (let c = 0; c < roles.length; c++) {
      const role = roles[c]!;
      if (!role || role === "Belirtilmedi") continue;
      const raw = cols[c + 1];
      if (raw == null || !String(raw).trim()) continue;
      const people = parseCapacityPeople(raw);
      if (!Number.isFinite(people)) {
        skipped += 1;
        continue;
      }
      rows.push({ project, year, week, role, people });
      any = true;
    }
    if (!any) skipped += 1;
  }

  const byKey = new Map<string, ParsedCapacityRow>();
  for (const row of rows) {
    byKey.set(capacityKey(row.project, row.year, row.week, row.role), row);
  }
  const deduped = [...byKey.values()];
  return { rows: deduped, skipped, duplicates: rows.length - deduped.length };
}

/**
 * Excel kapasite.
 * Öncelik matris: ProjeAdı | Donatım | Hull | …
 *                2026-35 | 5 | 2 | …
 * Aksi halde düz: proje | hafta | tip | kişi
 */
export function parseCapacityText(
  text: string,
  defaultYear: number
): { rows: ParsedCapacityRow[]; skipped: number; duplicates: number } {
  const lines = String(text)
    .replace(/^\uFEFF/, "")
    .split(/\r\n|\n|\r/)
    .map((l) => l.replace(/\u00a0/g, " ").trimEnd())
    .filter((l) => l.trim());

  if (!lines.length) return { rows: [], skipped: 0, duplicates: 0 };

  if (isCapacityMatrix(lines)) return parseCapacityMatrixText(lines, defaultYear);

  let map = { project: 0, year: -1, week: 1, role: 2, people: 3 };
  let start = 0;
  const first = splitCapacityLine(lines[0]!);
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
    const cols = splitCapacityLine(lines[i]!);
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

  const byKey = new Map<string, ParsedCapacityRow>();
  for (const row of rows) {
    byKey.set(capacityKey(row.project, row.year, row.week, row.role), row);
  }
  const deduped = [...byKey.values()];
  const duplicates = rows.length - deduped.length;
  return { rows: deduped, skipped, duplicates };
}

/** Dosyadan kapasite matrisi / düz format */
export async function parseCapacityFile(
  file: File,
  defaultYear: number
): Promise<{ rows: ParsedCapacityRow[]; skipped: number; duplicates: number; sheetName: string }> {
  const { text, sheetName } = await readSpreadsheetFileToTsv(file);
  const parsed = parseCapacityText(text, defaultYear);
  return { ...parsed, sheetName };
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
