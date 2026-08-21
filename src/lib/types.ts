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

/** İsimdeki son `[KATEGORİ]` etiketi — yoksa boş */
export function activityCategory(name: string): string {
  const matches = [...String(name || "").matchAll(/\[([^\]]+)\]/g)];
  if (!matches.length) return "";
  return (matches[matches.length - 1]![1] || "").trim();
}

/** Cyan, sarı, mavi ve uyumlu pastel tonlar */
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
];

const CATEGORY_PREF: Record<string, number> = {
  ARRGMNT: 0,
  ARRANGEMENT: 0,
  ARRANGE: 0,
  STRUCT: 1,
  STRUCTURE: 1,
  STEEL: 1,
  PIPE: 2,
  PIPING: 2,
  HVAC: 3,
  ELEC: 4,
  ELECTRICAL: 4,
  OUTFIT: 5,
  OUTFITTING: 5,
  PAINT: 6,
  INSUL: 7,
  INSULATION: 7,
};

function hashCategory(key: string): number {
  let h = 0;
  const s = key.toLocaleUpperCase("tr");
  for (let i = 0; i < s.length; i++) h = (h * 33 + s.charCodeAt(i)) >>> 0;
  return h;
}

export type CategoryTone = (typeof CATEGORY_PALETTE)[number];

export function categoryTone(category: string): CategoryTone {
  const key = (category || "").trim();
  if (!key) {
    return { label: "none", bar: "#64748b", bg: "#f8fafc", border: "#e2e8f0" };
  }
  const pref = CATEGORY_PREF[key.toLocaleUpperCase("en-US")] ?? CATEGORY_PREF[key.toLocaleUpperCase("tr")];
  const idx =
    pref != null ? pref % CATEGORY_PALETTE.length : hashCategory(key) % CATEGORY_PALETTE.length;
  return CATEGORY_PALETTE[idx]!;
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
    stages: Array.isArray(data.stages) ? data.stages : [],
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
