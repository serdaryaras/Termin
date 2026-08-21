export type Job = {
  id: string;
  name: string;
  /** Proje adı — aynı planda birden fazla proje aktivitesi tutulabilir */
  project: string;
  role: string;
  hours: number;
  /** İşin aynı anda kullandığı kişi sayısı */
  people: number;
};

export type Stage = {
  jobIds: string[];
  /**
   * Bu aşamadan sonra (aşama bitince) sonraki aktivitelerin
   * en erken başlangıcına eklenecek iş günü boşluğu / öteleme.
   */
  gapAfterDays?: number;
};

/** Bitiş-başlangıç (FS): ardıl, öncül bitmeden başlamaz */
export type Dependency = {
  predecessorId: string;
  successorId: string;
};

/** Haftalık aktivite ilerleme takibi (%25 dilimler) */
export type ProgressPercent = 0 | 25 | 50 | 75 | 100;

export type JobWeekProgress = {
  jobId: string;
  year: number;
  week: number;
  percent: ProgressPercent;
  reason: string;
};

/** Eski rol kapasitesi — yedek; asıl kısıt haftalık proje kapasitesidir */
export type ResourceGroup = {
  role: string;
  capacity: number;
};

/** Her proje × ISO hafta × personel tipi için kullanılabilir kişi */
export type ProjectWeekCapacity = {
  project: string;
  year: number;
  week: number;
  /** Personel tipi (Donatım, Konstrüksiyon, …) — işteki role ile eşleşir */
  role: string;
  people: number;
};

export type Plan = {
  name: string;
  startDate: string;
  hoursPerDay: number;
  jobs: Job[];
  stages: Stage[];
  /** Öncül → ardıl (finish-to-start) bağları */
  dependencies: Dependency[];
  resourceGroups: ResourceGroup[];
  weeklyCapacities: ProjectWeekCapacity[];
  /** Aktivite × ISO hafta ilerleme ve takılma nedeni */
  jobProgress: JobWeekProgress[];
};

export type ParsedRow = {
  name: string;
  project: string;
  role: string;
  hours: number;
  people: number;
};

/** Excel aktarım sonucu — matris formatında kategori zinciri bağımlılıkları da gelir */
export type ParsedJobImport = {
  rows: ParsedRow[];
  /** FS: aynı resim içinde kategori hiyerarşisi (ör. Class → 3D Model → ISO) */
  dependencies: Array<{
    predecessor: { project: string; name: string };
    successor: { project: string; name: string };
  }>;
  skipped: number;
  /** matrix = proje|iş|Class|3D Model|… hücreleri Role:saat */
  format: "matrix" | "flat";
};

export type ParsedCapacityRow = {
  project: string;
  year: number;
  week: number;
  role: string;
  people: number;
};

export const DEFAULT_PROJECT = "Genel";

export const ROLE_OPTIONS = [
  "Donatım",
  "Konstrüksiyon",
  "Kaynakçı",
  "Montajcı",
  "Borucu",
  "Elektrikçi",
  "Ressam",
  "Vinç operatörü",
  "Tekniker",
  "Mühendis",
  "Yardımcı personel",
] as const;

/** Excel / kısa yazımları standart personel tipine çevirir */
export function normalizeRole(raw: string | undefined | null): string {
  const s = (raw && String(raw).trim()) || "";
  if (!s) return "Belirtilmedi";
  const key = s
    .toLocaleLowerCase("tr")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/ı/g, "i");
  if (/^donat/.test(key)) return "Donatım";
  if (/^konstr/.test(key)) return "Konstrüksiyon";
  if (/^kaynak/.test(key)) return "Kaynakçı";
  if (/^montaj/.test(key)) return "Montajcı";
  if (/^boru/.test(key)) return "Borucu";
  if (/^elektrik/.test(key)) return "Elektrikçi";
  if (/^ressam|^boya/.test(key)) return "Ressam";
  if (/^vinc/.test(key)) return "Vinç operatörü";
  if (/^tekniker/.test(key)) return "Tekniker";
  if (/^muhendis/.test(key)) return "Mühendis";
  if (/^yardimci/.test(key)) return "Yardımcı personel";
  return s;
}

export const BAR_COLORS = ["#0d4f8b", "#6d28d9", "#0f766e", "#b45309", "#be123c"];

/** İsimdeki kategori — önce `252.100.104.Class-…`, yoksa son `[KATEGORİ]` */
export function activityCategory(name: string): string {
  const s = String(name || "").trim();
  const dotted = s.match(/^\d+(?:\.\d+)+\.([^.]+?)-/);
  if (dotted?.[1]) return dotted[1].trim();
  const matches = [...s.matchAll(/\[([^\]]+)\]/g)];
  if (!matches.length) return "";
  return (matches[matches.length - 1]![1] || "").trim();
}

/** Cyan, sarı, mavi ve ayırt edilebilir pastel tonlar */
export const CATEGORY_PALETTE: ReadonlyArray<{
  label: string;
  bar: string;
  bg: string;
  border: string;
}> = [
  { label: "cyan", bar: "#0891b2", bg: "#ecfeff", border: "#67e8f9" },
  { label: "yellow", bar: "#ca8a04", bg: "#fefce8", border: "#fde047" },
  { label: "blue", bar: "#2563eb", bg: "#eff6ff", border: "#93c5fd" },
  { label: "teal", bar: "#0d9488", bg: "#f0fdfa", border: "#5eead4" },
  { label: "amber", bar: "#d97706", bg: "#fffbeb", border: "#fcd34d" },
  { label: "indigo", bar: "#4f46e5", bg: "#eef2ff", border: "#a5b4fc" },
  { label: "sky", bar: "#0284c7", bg: "#f0f9ff", border: "#7dd3fc" },
  { label: "lime", bar: "#65a30d", bg: "#f7fee7", border: "#bef264" },
  { label: "violet", bar: "#7c3aed", bg: "#f5f3ff", border: "#c4b5fd" },
  { label: "rose", bar: "#e11d48", bg: "#fff1f2", border: "#fda4af" },
  { label: "orange", bar: "#ea580c", bg: "#fff7ed", border: "#fdba74" },
  { label: "fuchsia", bar: "#c026d3", bg: "#fdf4ff", border: "#f0abfc" },
  { label: "emerald", bar: "#059669", bg: "#ecfdf5", border: "#6ee7b7" },
  { label: "red", bar: "#dc2626", bg: "#fef2f2", border: "#fca5a5" },
  { label: "slate", bar: "#475569", bg: "#f1f5f9", border: "#cbd5e1" },
  { label: "pink", bar: "#db2777", bg: "#fdf2f8", border: "#f9a8d4" },
];

export type CategoryTone = {
  label: string;
  bar: string;
  bg: string;
  border: string;
};

const EMPTY_TONE: CategoryTone = {
  label: "none",
  bar: "#64748b",
  bg: "#f8fafc",
  border: "#e2e8f0",
};

/** Palet dışı kategoriler için altın açı ile benzersiz HSL tonu */
export function categoryToneAt(index: number): CategoryTone {
  if (index >= 0 && index < CATEGORY_PALETTE.length) {
    return CATEGORY_PALETTE[index]!;
  }
  const hue = Math.round((index * 137.508) % 360);
  return {
    label: `hue-${hue}`,
    bar: `hsl(${hue} 70% 42%)`,
    bg: `hsl(${hue} 85% 94%)`,
    border: `hsl(${hue} 70% 75%)`,
  };
}

/**
 * Her benzersiz kategoriye farklı renk — alfabetik sırayla paletten,
 * taşınca HSL üretir (çarpışma yok).
 */
export function buildCategoryToneMap(categories: Iterable<string>): Map<string, CategoryTone> {
  const unique = [
    ...new Set(
      [...categories]
        .map((c) => String(c || "").trim())
        .filter(Boolean)
    ),
  ].sort((a, b) => a.localeCompare(b, "tr"));
  const map = new Map<string, CategoryTone>();
  unique.forEach((cat, i) => {
    map.set(cat, categoryToneAt(i));
  });
  return map;
}

export function categoryTone(
  category: string,
  toneMap?: Map<string, CategoryTone>
): CategoryTone {
  const key = (category || "").trim();
  if (!key) return EMPTY_TONE;
  if (toneMap?.has(key)) return toneMap.get(key)!;
  // Harita yoksa tekil çağrı: yine çarpışmasız olmasa da sabit hash yerine
  // tek kategoriyi paletin ilkine bağlama — harita tercih edilir.
  return categoryToneAt(
    [...key].reduce((h, ch) => (h * 33 + ch.charCodeAt(0)) >>> 0, 0) %
      Math.max(CATEGORY_PALETTE.length * 3, 1)
  );
}

export function todayIso(): string {
  const d = new Date();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${d.getFullYear()}-${m}-${day}`;
}

export function emptyPlan(): Plan {
  return {
    name: "İş planı",
    startDate: todayIso(),
    hoursPerDay: 8,
    jobs: [],
    stages: [],
    dependencies: [],
    resourceGroups: [],
    weeklyCapacities: [],
    jobProgress: [],
  };
}

export function newId(): string {
  return crypto.randomUUID();
}

export function normalizeJob(raw: Partial<Job> & { name?: string }): Job | null {
  if (!raw || !raw.name) return null;
  const hours = Number(raw.hours);
  const people = Number(raw.people);
  return {
    id: raw.id || newId(),
    name: String(raw.name).trim(),
    project: (raw.project && String(raw.project).trim()) || DEFAULT_PROJECT,
    role: normalizeRole(raw.role),
    hours: Number.isFinite(hours) && hours > 0 ? hours : 1,
    people: Number.isFinite(people) && people > 0 ? people : 1,
  };
}

function normalizeCapacity(raw: Partial<ProjectWeekCapacity>): ProjectWeekCapacity | null {
  if (!raw || !raw.project) return null;
  const year = Number(raw.year);
  const week = Number(raw.week);
  const people = Number(raw.people);
  if (!Number.isFinite(year) || !Number.isFinite(week) || week < 1 || week > 53) return null;
  if (!Number.isFinite(people) || people < 0) return null;
  return {
    project: String(raw.project).trim(),
    year,
    week,
    role: normalizeRole(raw.role),
    people,
  };
}

const PROGRESS_PERCENTS: ProgressPercent[] = [0, 25, 50, 75, 100];

function normalizeJobProgress(
  raw: Partial<JobWeekProgress>,
  jobIds: Set<string>
): JobWeekProgress | null {
  if (!raw || !raw.jobId || !jobIds.has(String(raw.jobId))) return null;
  const year = Number(raw.year);
  const week = Number(raw.week);
  const percent = Number(raw.percent) as ProgressPercent;
  if (!Number.isFinite(year) || !Number.isFinite(week) || week < 1 || week > 53) return null;
  if (!PROGRESS_PERCENTS.includes(percent)) return null;
  return {
    jobId: String(raw.jobId),
    year,
    week,
    percent,
    reason: raw.reason != null ? String(raw.reason) : "",
  };
}

export function normalizePlan(data: Partial<Plan> | null | undefined): Plan {
  const base = emptyPlan();
  if (!data) return base;
  const jobs = Array.isArray(data.jobs)
    ? data.jobs.map((j) => normalizeJob(j)).filter((j): j is Job => Boolean(j))
    : [];
  const jobIds = new Set(jobs.map((j) => j.id));
  const weeklyCapacities = Array.isArray(data.weeklyCapacities)
    ? data.weeklyCapacities.map((c) => normalizeCapacity(c)).filter((c): c is ProjectWeekCapacity => Boolean(c))
    : [];
  const dependencies = Array.isArray(data.dependencies)
    ? data.dependencies
        .filter(
          (d) =>
            d &&
            jobIds.has(String(d.predecessorId)) &&
            jobIds.has(String(d.successorId)) &&
            d.predecessorId !== d.successorId
        )
        .map((d) => ({
          predecessorId: String(d.predecessorId),
          successorId: String(d.successorId),
        }))
    : [];
  const jobProgress = Array.isArray(data.jobProgress)
    ? data.jobProgress
        .map((p) => normalizeJobProgress(p, jobIds))
        .filter((p): p is JobWeekProgress => Boolean(p))
    : [];
  return {
    name: data.name || base.name,
    startDate: data.startDate || base.startDate,
    hoursPerDay: Number(data.hoursPerDay) || 8,
    jobs,
    stages: Array.isArray(data.stages)
      ? data.stages
          .filter((s) => s && Array.isArray(s.jobIds))
          .map((s) => {
            const gap = Number((s as Stage).gapAfterDays);
            return {
              jobIds: s.jobIds.filter((id) => jobIds.has(String(id))).map(String),
              ...(Number.isFinite(gap) && gap > 0 ? { gapAfterDays: Math.round(gap) } : {}),
            };
          })
          .filter((s) => s.jobIds.length > 0)
      : [],
    dependencies,
    resourceGroups: Array.isArray(data.resourceGroups) ? data.resourceGroups : [],
    weeklyCapacities,
    jobProgress,
  };
}

export function jobKey(project: string, name: string): string {
  return `${project.trim().toLocaleLowerCase("tr")}::${name.trim().toLocaleLowerCase("tr")}`;
}

export function capacityKey(project: string, year: number, week: number, role: string): string {
  return `${project.trim().toLocaleLowerCase("tr")}::${year}::${week}::${role.trim().toLocaleLowerCase("tr")}`;
}

export function listRoles(jobs: Job[], capacities: ProjectWeekCapacity[] = []): string[] {
  const set = new Set<string>();
  for (const r of ROLE_OPTIONS) set.add(r);
  for (const j of jobs) if (j.role?.trim()) set.add(j.role.trim());
  for (const c of capacities) if (c.role?.trim()) set.add(c.role.trim());
  return [...set].sort((a, b) => a.localeCompare(b, "tr"));
}

export function listProjects(jobs: Job[], capacities: ProjectWeekCapacity[] = []): string[] {
  const set = new Set<string>();
  for (const j of jobs) {
    const p = (j.project || DEFAULT_PROJECT).trim() || DEFAULT_PROJECT;
    set.add(p);
  }
  for (const c of capacities) {
    if (c.project?.trim()) set.add(c.project.trim());
  }
  return [...set].sort((a, b) => a.localeCompare(b, "tr"));
}
