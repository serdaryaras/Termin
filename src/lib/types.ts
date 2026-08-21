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
