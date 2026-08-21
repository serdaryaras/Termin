import {
  activityCategory,
  BAR_COLORS,
  buildCategoryToneMap,
  capacityKey,
  DEFAULT_PROJECT,
  emptyPlan,
  normalizeRole,
  type Job,
  type Plan,
  type ProjectWeekCapacity,
  type Stage,
} from "./types";

export function scheduledIds(plan: Plan): Set<string> {
  const ids = new Set<string>();
  for (const stage of plan.stages) {
    for (const id of stage.jobIds) ids.add(id);
  }
  return ids;
}

export function jobById(plan: Plan, id: string): Job | undefined {
  return plan.jobs.find((j) => j.id === id);
}

export function findPlacement(plan: Plan, jobId: string) {
  for (let s = 0; s < plan.stages.length; s++) {
    const i = plan.stages[s].jobIds.indexOf(jobId);
    if (i >= 0) return { stageIndex: s, itemIndex: i };
  }
  return null;
}

function prune(stages: Stage[]): Stage[] {
  return stages.filter((s) => s.jobIds.length > 0);
}

function clonePlan(plan: Plan): Plan {
  return {
    ...plan,
    jobs: plan.jobs.map((j) => ({ ...j })),
    stages: plan.stages.map((s) => {
      const gap = Number(s.gapAfterDays);
      return {
        jobIds: [...s.jobIds],
        ...(Number.isFinite(gap) && gap > 0 ? { gapAfterDays: Math.round(gap) } : {}),
      };
    }),
    dependencies: (plan.dependencies || []).map((d) => ({ ...d })),
    resourceGroups: plan.resourceGroups.map((g) => ({ ...g })),
    weeklyCapacities: plan.weeklyCapacities.map((c) => ({ ...c })),
    jobProgress: (plan.jobProgress || []).map((p) => ({ ...p })),
  };
}

export function predecessorsOf(plan: Plan, jobId: string): string[] {
  return (plan.dependencies || [])
    .filter((d) => d.successorId === jobId)
    .map((d) => d.predecessorId);
}

export function successorsOf(plan: Plan, jobId: string): string[] {
  return (plan.dependencies || [])
    .filter((d) => d.predecessorId === jobId)
    .map((d) => d.successorId);
}

export function addDependency(plan: Plan, predecessorId: string, successorId: string): Plan {
  if (!predecessorId || !successorId || predecessorId === successorId) return plan;
  const exists = (plan.dependencies || []).some(
    (d) => d.predecessorId === predecessorId && d.successorId === successorId
  );
  if (exists) return plan;
  // Döngü engeli: successor zaten predecessor'ın atası olmamalı
  const seen = new Set<string>();
  const stack = [successorId];
  while (stack.length) {
    const id = stack.pop()!;
    if (id === predecessorId) return plan;
    if (seen.has(id)) continue;
    seen.add(id);
    for (const s of successorsOf(plan, id)) stack.push(s);
  }
  const next = clonePlan(plan);
  next.dependencies = [...(next.dependencies || []), { predecessorId, successorId }];
  return next;
}

export function removeDependency(plan: Plan, predecessorId: string, successorId: string): Plan {
  const next = clonePlan(plan);
  next.dependencies = (next.dependencies || []).filter(
    (d) => !(d.predecessorId === predecessorId && d.successorId === successorId)
  );
  return next;
}

export function clearDependenciesForJob(plan: Plan, jobId: string): Plan {
  const next = clonePlan(plan);
  next.dependencies = (next.dependencies || []).filter(
    (d) => d.predecessorId !== jobId && d.successorId !== jobId
  );
  return next;
}

export function clearAllDependencies(plan: Plan): Plan {
  return { ...plan, dependencies: [] };
}

/** Seçim zincirindeki ardışık öncül→ardıl bağlarını kaldırır */
export function clearChainDependencies(plan: Plan, chainIds: string[]): Plan {
  if (chainIds.length < 2) return plan;
  const pairs = new Set<string>();
  for (let i = 0; i < chainIds.length - 1; i++) {
    pairs.add(`${chainIds[i]}::${chainIds[i + 1]}`);
  }
  const next = clonePlan(plan);
  next.dependencies = (next.dependencies || []).filter(
    (d) => !pairs.has(`${d.predecessorId}::${d.successorId}`)
  );
  return next;
}

/** Verilen işlerin tüm öncül/ardıl bağlarını kaldırır */
export function clearDependenciesForJobs(plan: Plan, jobIds: string[]): Plan {
  if (!jobIds.length) return plan;
  const set = new Set(jobIds);
  const next = clonePlan(plan);
  next.dependencies = (next.dependencies || []).filter(
    (d) => !set.has(d.predecessorId) && !set.has(d.successorId)
  );
  return next;
}

export function addJobToStage(plan: Plan, jobId: string, stageIndex: number, asNewStage: boolean): Plan {
  const next = clonePlan(plan);
  const place = findPlacement(next, jobId);

  if (place && !asNewStage && place.stageIndex === stageIndex) {
    const ids = next.stages[stageIndex].jobIds;
    ids.splice(place.itemIndex, 1);
    ids.push(jobId);
    return next;
  }

  let target = stageIndex;
  if (place) {
    const emptying = next.stages[place.stageIndex].jobIds.length === 1;
    if (emptying && place.stageIndex < target) target -= 1;
    next.stages[place.stageIndex].jobIds = next.stages[place.stageIndex].jobIds.filter((id) => id !== jobId);
    next.stages = prune(next.stages);
  }

  if (asNewStage) {
    const idx = Math.max(0, Math.min(target, next.stages.length));
    next.stages.splice(idx, 0, { jobIds: [jobId] });
  } else if (!next.stages[target]) {
    next.stages.push({ jobIds: [jobId] });
  } else {
    // Aşama üzerine bırakma = o konumda yeni sıralı aşama (paralel birleştirme yok)
    const idx = Math.max(0, Math.min(target + 1, next.stages.length));
    next.stages.splice(idx, 0, { jobIds: [jobId] });
  }
  return next;
}

export function removeFromSchedule(plan: Plan, jobId: string): Plan {
  let next = clonePlan(plan);
  next.stages = prune(
    next.stages.map((s) => {
      const gap = Number(s.gapAfterDays);
      return {
        jobIds: s.jobIds.filter((id) => id !== jobId),
        ...(Number.isFinite(gap) && gap > 0 ? { gapAfterDays: Math.round(gap) } : {}),
      };
    })
  );
  next = clearDependenciesForJob(next, jobId);
  return next;
}

/** Öncelik sırasını tamamen temizler (iş kalemleri kalır) */
export function clearPriorityOrder(plan: Plan): Plan {
  return { ...plan, stages: [], dependencies: [] };
}

/** Tüm haftalık kapasite kayıtlarını siler */
export function clearAllWeeklyCapacities(plan: Plan): Plan {
  return { ...plan, weeklyCapacities: [] };
}

/** Belirli bir projenin tüm haftalık kapasite / kaynak tanımlarını siler */
export function clearWeeklyCapacitiesForProject(plan: Plan, project: string): Plan {
  const key = (project || "").trim();
  if (!key) return plan;
  const next = clonePlan(plan);
  next.weeklyCapacities = next.weeklyCapacities.filter(
    (c) => (c.project || DEFAULT_PROJECT) !== key
  );
  return next;
}

function moveStage(stages: Stage[], from: number, to: number): Stage[] {
  if (from === to || from < 0 || to < 0 || from >= stages.length) return stages;
  const copy = stages.map((s) => {
    const gap = Number(s.gapAfterDays);
    return {
      jobIds: [...s.jobIds],
      ...(Number.isFinite(gap) && gap > 0 ? { gapAfterDays: Math.round(gap) } : {}),
    };
  });
  const [row] = copy.splice(from, 1);
  const insert = to > from ? to - 1 : to;
  const clamped = Math.max(0, Math.min(insert, copy.length));
  copy.splice(clamped, 0, row);
  return copy;
}

/** Seçili işlerin aşamalarını (sıra korunarak) hedef işin aşamasının önüne taşır */
export function moveJobsGroupBefore(plan: Plan, jobIds: string[], targetJobId: string): Plan {
  const want = [...new Set(jobIds.filter((id) => findPlacement(plan, id)))];
  if (!want.length) return plan;
  if (want.includes(targetJobId)) return plan;

  const next = clonePlan(plan);
  const stageIndexOf = (jobId: string): number => {
    for (let s = 0; s < next.stages.length; s++) {
      if (next.stages[s]!.jobIds.includes(jobId)) return s;
    }
    return -1;
  };

  const stagesToMove = [
    ...new Set(want.map((id) => stageIndexOf(id)).filter((i) => i >= 0)),
  ].sort((a, b) => a - b);
  if (!stagesToMove.length) return plan;

  const extracted = stagesToMove.map((i) => {
    const s = next.stages[i]!;
    const gap = Number(s.gapAfterDays);
    return {
      jobIds: [...s.jobIds],
      ...(Number.isFinite(gap) && gap > 0 ? { gapAfterDays: Math.round(gap) } : {}),
    };
  });
  for (let k = stagesToMove.length - 1; k >= 0; k--) {
    next.stages.splice(stagesToMove[k]!, 1);
  }

  let insertAt = stageIndexOf(targetJobId);
  if (insertAt < 0) insertAt = next.stages.length;
  next.stages.splice(insertAt, 0, ...extracted);
  return next;
}

export function moveSelectedStage(plan: Plan, jobId: string, dir: "top" | "bottom" | "up" | "down"): Plan {
  const place = findPlacement(plan, jobId);
  if (!place) return plan;
  const next = clonePlan(plan);
  const { stageIndex } = place;
  if (dir === "top") next.stages = moveStage(next.stages, stageIndex, 0);
  if (dir === "bottom") next.stages = moveStage(next.stages, stageIndex, next.stages.length);
  if (dir === "up") next.stages = moveStage(next.stages, stageIndex, Math.max(0, stageIndex - 1));
  if (dir === "down") {
    next.stages = moveStage(next.stages, stageIndex, Math.min(next.stages.length, stageIndex + 2));
  }
  return next;
}

/** Seçili aşamayı öncelik sırasında delta kadar kaydır (− yukarı / + aşağı) */
export function moveSelectedStageBy(plan: Plan, jobId: string, delta: number): Plan {
  const place = findPlacement(plan, jobId);
  if (!place || !delta) return plan;
  const from = place.stageIndex;
  const target = Math.max(0, Math.min(plan.stages.length - 1, from + delta));
  if (target === from) return plan;
  const next = clonePlan(plan);
  next.stages = moveStage(next.stages, from, delta > 0 ? target + 1 : target);
  return next;
}

/** Seçili işlerin aşamalarını (tekrarsız) sıra indeksine göre döndür */
function selectedStageReps(plan: Plan, jobIds: string[]): { jobId: string; stageIndex: number }[] {
  const byStage = new Map<number, string>();
  for (const id of jobIds) {
    const place = findPlacement(plan, id);
    if (!place) continue;
    if (!byStage.has(place.stageIndex)) byStage.set(place.stageIndex, id);
  }
  return [...byStage.entries()]
    .sort((a, b) => a[0] - b[0])
    .map(([stageIndex, jobId]) => ({ jobId, stageIndex }));
}

/** Birden fazla seçili aşamayı aynı delta ile kaydır (göreli sıra korunur) */
export function moveJobsByDelta(plan: Plan, jobIds: string[], delta: number): Plan {
  if (!delta || !jobIds.length) return plan;
  const reps = selectedStageReps(plan, jobIds);
  if (!reps.length) return plan;
  let next = plan;
  const ordered = delta > 0 ? [...reps].reverse() : reps;
  for (const { jobId } of ordered) {
    next = moveSelectedStageBy(next, jobId, delta);
  }
  return next;
}

/** Birden fazla seçili aşamayı üste / alta / bir adım kaydır */
export function moveJobsDir(
  plan: Plan,
  jobIds: string[],
  dir: "top" | "bottom" | "up" | "down"
): Plan {
  if (!jobIds.length) return plan;
  if (dir === "up") return moveJobsByDelta(plan, jobIds, -1);
  if (dir === "down") return moveJobsByDelta(plan, jobIds, 1);

  const reps = selectedStageReps(plan, jobIds);
  if (!reps.length) return plan;
  const next = clonePlan(plan);
  const indices = reps.map((r) => r.stageIndex).sort((a, b) => a - b);
  const block = indices.map((i) => {
    const s = next.stages[i]!;
    const gap = Number(s.gapAfterDays);
    return {
      jobIds: [...s.jobIds],
      ...(Number.isFinite(gap) && gap > 0 ? { gapAfterDays: Math.round(gap) } : {}),
    };
  });
  const remaining = next.stages.filter((_, i) => !indices.includes(i));
  if (dir === "top") {
    next.stages = [...block, ...remaining];
  } else {
    next.stages = [...remaining, ...block];
  }
  return next;
}

/** Seçili aşamadan sonra N iş günü boşluk (sonraki tüm aktiviteleri öteleyebilir) */
export function setStageGapAfterDays(plan: Plan, stageIndex: number, days: number): Plan {
  if (stageIndex < 0 || stageIndex >= plan.stages.length) return plan;
  const next = clonePlan(plan);
  const gap = Math.max(0, Math.round(Number(days) || 0));
  if (gap > 0) next.stages[stageIndex]!.gapAfterDays = gap;
  else delete next.stages[stageIndex]!.gapAfterDays;
  return next;
}

export function deleteJob(plan: Plan, jobId: string): Plan {
  const next = removeFromSchedule(plan, jobId);
  next.jobs = next.jobs.filter((j) => j.id !== jobId);
  next.jobProgress = (next.jobProgress || []).filter((p) => p.jobId !== jobId);
  return next;
}

/** Tüm işler, sıra, bağlar ve takip kayıtlarını siler (kapasite kalır) */
export function clearAllJobs(plan: Plan): Plan {
  const next = clonePlan(plan);
  next.jobs = [];
  next.stages = [];
  next.dependencies = [];
  next.jobProgress = [];
  return next;
}

/** Plan içeriğini boşaltır; ad / başlangıç / günlük saat korunur */
export function clearEntirePlan(plan: Plan): Plan {
  return {
    ...emptyPlan(),
    name: plan.name || emptyPlan().name,
    startDate: plan.startDate || emptyPlan().startDate,
    hoursPerDay: plan.hoursPerDay || 8,
  };
}

export function clearJobProgressForWeek(plan: Plan, year: number, week: number): Plan {
  const next = clonePlan(plan);
  next.jobProgress = (next.jobProgress || []).filter(
    (p) => !(p.year === year && p.week === week)
  );
  return next;
}

export function clearAllJobProgress(plan: Plan): Plan {
  const next = clonePlan(plan);
  next.jobProgress = [];
  return next;
}

export function upsertWeeklyCapacity(
  plan: Plan,
  entry: ProjectWeekCapacity
): Plan {
  const next = clonePlan(plan);
  const normalized: ProjectWeekCapacity = {
    ...entry,
    project: entry.project.trim(),
    role: normalizeRole(entry.role),
    people: Number(entry.people) || 0,
  };
  const key = capacityKey(normalized.project, normalized.year, normalized.week, normalized.role);
  const idx = next.weeklyCapacities.findIndex(
    (c) => capacityKey(c.project, c.year, c.week, c.role) === key
  );
  if (idx >= 0) next.weeklyCapacities[idx] = normalized;
  else next.weeklyCapacities.push(normalized);
  return next;
}

/** Seçili projenin (veya tümünün) haftalık kişi kapasitelerini delta kadar kaydırır */
export function adjustWeeklyCapacitiesByDelta(
  plan: Plan,
  project: string,
  delta: number
): Plan {
  const d = Math.round(Number(delta) || 0);
  if (!d) return plan;
  const next = clonePlan(plan);
  const all = !project || project === "all";
  let changed = false;
  for (const c of next.weeklyCapacities) {
    if (!all && (c.project || DEFAULT_PROJECT) !== project) continue;
    const people = Math.max(0, (Number(c.people) || 0) + d);
    if (people !== c.people) {
      c.people = people;
      changed = true;
    }
  }
  return changed ? next : plan;
}

/** Kimlik (proje/hafta/tip) değişince eski kaydı silip yenisini yazar */
export function replaceWeeklyCapacity(
  plan: Plan,
  from: { project: string; year: number; week: number; role: string },
  to: ProjectWeekCapacity
): Plan {
  let next = removeWeeklyCapacity(plan, from.project, from.year, from.week, from.role);
  next = upsertWeeklyCapacity(next, to);
  return next;
}

export function removeWeeklyCapacity(
  plan: Plan,
  project: string,
  year: number,
  week: number,
  role: string
): Plan {
  const next = clonePlan(plan);
  const key = capacityKey(project, year, week, role);
  next.weeklyCapacities = next.weeklyCapacities.filter(
    (c) => capacityKey(c.project, c.year, c.week, c.role) !== key
  );
  return next;
}

/** İş günü / hafta */
export const WORK_DAYS_PER_WEEK = 5;

export function formatWeekLabel(year: number, week: number): string {
  return `${year}-${String(week).padStart(2, "0")}`;
}

export function formatWeekOnly(week: number): string {
  return String(week).padStart(2, "0");
}

export function getIsoWeekParts(date: Date): { year: number; week: number } {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  const week = Math.ceil(((d.getTime() - yearStart.getTime()) / 86400000 + 1) / 7);
  return { year: d.getUTCFullYear(), week };
}

export function weeksInIsoYear(year: number): number {
  return getIsoWeekParts(new Date(year, 11, 28)).week;
}

export function parseIsoDate(iso: string): Date {
  return new Date(`${iso}T00:00:00`);
}

/** ISO haftanın pazartesi tarihi (YYYY-MM-DD) */
export function mondayOfIsoWeek(year: number, week: number): string {
  const jan4 = new Date(Date.UTC(year, 0, 4));
  const day = jan4.getUTCDay() || 7;
  const monday = new Date(jan4);
  monday.setUTCDate(jan4.getUTCDate() - (day - 1) + (week - 1) * 7);
  const y = monday.getUTCFullYear();
  const m = String(monday.getUTCMonth() + 1).padStart(2, "0");
  const d = String(monday.getUTCDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

/** En erken kapasite haftasının pazartesisi — Gantt ekseni buradan başlamalı */
export function earliestCapacityStartDate(plan: Plan): string | null {
  if (!plan.weeklyCapacities.length) return null;
  let year = plan.weeklyCapacities[0].year;
  let week = plan.weeklyCapacities[0].week;
  for (const c of plan.weeklyCapacities) {
    if (c.year < year || (c.year === year && c.week < week)) {
      year = c.year;
      week = c.week;
    }
  }
  return mondayOfIsoWeek(year, week);
}

function nextIsoWeek(year: number, week: number): { year: number; week: number } {
  const maxW = weeksInIsoYear(year);
  if (week >= maxW) return { year: year + 1, week: 1 };
  return { year, week: week + 1 };
}

export function prevIsoWeek(year: number, week: number): { year: number; week: number } {
  if (week <= 1) {
    const py = year - 1;
    return { year: py, week: weeksInIsoYear(py) };
  }
  return { year, week: week - 1 };
}

export { nextIsoWeek };

/** Plan başlangıcından hedef ISO haftaya kadar kaçıncı hafta (0 tabanlı); hedef önceyse null */
export function isoWeekOffsetFromPlanStart(
  planStartIso: string,
  year: number,
  week: number
): number | null {
  const start = getIsoWeekParts(parseIsoDate(planStartIso));
  if (year < start.year || (year === start.year && week < start.week)) return null;
  let y = start.year;
  let w = start.week;
  for (let i = 0; i < 800; i++) {
    if (y === year && w === week) return i;
    const n = nextIsoWeek(y, w);
    y = n.year;
    w = n.week;
  }
  return null;
}

/** Gantt satırı seçilen ISO haftayla kesişiyor mu? */
export function ganttRowOverlapsIsoWeek(
  row: { startDay: number; durationDays: number },
  planStartIso: string,
  year: number,
  week: number
): boolean {
  const offset = isoWeekOffsetFromPlanStart(planStartIso, year, week);
  if (offset == null) return false;
  const weekStart = offset * WORK_DAYS_PER_WEEK;
  const weekEnd = weekStart + WORK_DAYS_PER_WEEK;
  const jobStart = row.startDay;
  const jobEnd = row.startDay + row.durationDays;
  return jobStart < weekEnd && jobEnd > weekStart - 1e-9;
}

export function weekIndexFromPlanStart(_planStartIso: string, workDayOffset: number): number {
  return Math.max(0, workDayOffset) / WORK_DAYS_PER_WEEK;
}

export type WeekTick = { key: string; label: string; year: number; week: number };

export function buildWeekTicks(planStartIso: string, totalWorkDays: number): WeekTick[] {
  const start = parseIsoDate(planStartIso);
  const { year: y0, week: w0 } = getIsoWeekParts(start);
  const weeksNeeded = Math.max(1, Math.ceil(Math.max(0, totalWorkDays) / WORK_DAYS_PER_WEEK - 1e-9) || 1);

  const ticks: WeekTick[] = [];
  let year = y0;
  let week = w0;
  for (let i = 0; i < weeksNeeded; i++) {
    ticks.push({
      key: `${year}-${week}`,
      label: formatWeekLabel(year, week),
      year,
      week,
    });
    const next = nextIsoWeek(year, week);
    year = next.year;
    week = next.week;
  }
  return ticks;
}

export function weekRangeLabel(ticks: WeekTick[]): string {
  if (!ticks.length) return "";
  const first = ticks[0];
  const last = ticks[ticks.length - 1];
  if (first.year === last.year) {
    return `${first.year} · ${formatWeekOnly(first.week)}–${formatWeekOnly(last.week)}`;
  }
  return `${formatWeekLabel(first.year, first.week)} → ${formatWeekLabel(last.year, last.week)}`;
}

export function yearBands(ticks: WeekTick[]): Array<{ year: number; count: number }> {
  const bands: Array<{ year: number; count: number }> = [];
  for (const t of ticks) {
    const last = bands[bands.length - 1];
    if (last && last.year === t.year) last.count += 1;
    else bands.push({ year: t.year, count: 1 });
  }
  return bands;
}

export function isoWeekAtOffset(planStartIso: string, weekOffset: number): { year: number; week: number } {
  const start = parseIsoDate(planStartIso);
  let { year, week } = getIsoWeekParts(start);
  const steps = Math.max(0, Math.floor(weekOffset));
  for (let i = 0; i < steps; i++) {
    const n = nextIsoWeek(year, week);
    year = n.year;
    week = n.week;
  }
  return { year, week };
}

function jobPeople(job: Job): number {
  const n = Number(job.people);
  return Number.isFinite(n) && n > 0 ? n : 1;
}

/** Öncelik sırası: aşama sırası + aşama içi sıra */
export function priorityJobs(plan: Plan): Array<{ job: Job; stage: number }> {
  const out: Array<{ job: Job; stage: number }> = [];
  plan.stages.forEach((stage, si) => {
    for (const id of stage.jobIds) {
      const job = jobById(plan, id);
      if (job) out.push({ job, stage: si + 1 });
    }
  });
  return out;
}

function capacityLookup(plan: Plan): Map<string, number> {
  const map = new Map<string, number>();
  for (const c of plan.weeklyCapacities) {
    const key = capacityKey(c.project, c.year, c.week, normalizeRole(c.role));
    const people = Number(c.people);
    if (!Number.isFinite(people) || people < 0) continue;
    map.set(key, (map.get(key) ?? 0) + people);
  }
  return map;
}

/** Kapasite tablosu olan (proje+personel tipi) çiftleri */
function cappedProjectRoles(plan: Plan): Set<string> {
  const set = new Set<string>();
  for (const c of plan.weeklyCapacities) {
    const role = normalizeRole(c.role);
    set.add(`${c.project.trim().toLocaleLowerCase("tr")}::${role.trim().toLocaleLowerCase("tr")}`);
  }
  return set;
}

function daysOccupied(durationDays: number): number {
  return Math.max(1, Math.ceil(durationDays - 1e-9));
}

/** Kapasite ızgarası her zaman tam iş günü indeksi kullanır (ondalık anahtar kaymasını önler). */
function dayIndex(day: number): number {
  return Math.floor(day + 1e-9);
}

function jobDurationDays(job: Job, hoursPerDay: number): number {
  const hpd = hoursPerDay > 0 ? hoursPerDay : 8;
  return Math.max(1 / WORK_DAYS_PER_WEEK, job.hours / hpd);
}

function pairKey(project: string, role: string): string {
  return `${project.trim().toLocaleLowerCase("tr")}::${normalizeRole(role).trim().toLocaleLowerCase("tr")}`;
}

function weekCapacityAvailable(
  plan: Plan,
  capMap: Map<string, number>,
  capped: Set<string>,
  project: string,
  role: string,
  weekOffset: number
): number {
  const pair = pairKey(project, role);
  const hasCap = capped.has(pair);
  const planHasCapacity = plan.weeklyCapacities.length > 0;
  if (!hasCap) return planHasCapacity ? 0 : Infinity;
  const { year, week } = isoWeekAtOffset(plan.startDate, weekOffset);
  const ck = capacityKey(project, year, week, normalizeRole(role));
  const raw = capMap.get(ck);
  if (raw == null) return 0;
  const n = Number(raw);
  return Number.isFinite(n) ? n : 0;
}

/**
 * İş-günü çözünürlüğünde yerleştir: biten işin hemen ardından kapasite açılır.
 * Haftalık kişi limiti o takvim haftasındaki her iş günü için geçerlidir.
 */
function canPlaceAtDay(
  plan: Plan,
  capMap: Map<string, number>,
  capped: Set<string>,
  used: Map<string, number>,
  project: string,
  role: string,
  startDay: number,
  durationDays: number,
  people: number
): boolean {
  const start = dayIndex(startDay);
  const days = daysOccupied(durationDays);
  const pair = pairKey(project, role);
  const need = Number(people) || 1;
  for (let i = 0; i < days; i++) {
    const day = start + i;
    const weekOffset = Math.floor(day / WORK_DAYS_PER_WEEK);
    const available = weekCapacityAvailable(plan, capMap, capped, project, role, weekOffset);
    const already = used.get(`${pair}::d${day}`) ?? 0;
    if (already + need > available + 1e-9) return false;
  }
  return true;
}

function placeJobAtDay(
  used: Map<string, number>,
  project: string,
  role: string,
  startDay: number,
  durationDays: number,
  people: number
) {
  const start = dayIndex(startDay);
  const days = daysOccupied(durationDays);
  const pair = pairKey(project, role);
  const need = Number(people) || 1;
  for (let i = 0; i < days; i++) {
    const key = `${pair}::d${start + i}`;
    used.set(key, (used.get(key) ?? 0) + need);
  }
}

export type GanttRow = {
  job: Job;
  stage: number;
  /** Paralel hat (H1…): aynı anda çalışan akış */
  lane: number;
  startDay: number;
  durationDays: number;
  color: string;
  /** Satır gölgesi ([KATEGORİ] pastel) */
  tint: string;
  category: string;
  capacityOk: boolean;
};

export type GanttModel = {
  rows: GanttRow[];
  totalHours: number;
  totalDays: number;
  roleColors: Map<string, string>;
  categoryColors: Map<string, string>;
  warnings: string[];
};

/**
 * Öncelik sırası: kim önce kapasiteyi kapar.
 * Öncül-ardıl (FS): ardıl, tüm öncülleri bitmeden başlamaz (kapasite olsa bile).
 * Yerleşim iş günü bazlı — kapasite elverdikçe aynı hafta içinde devam edilebilir.
 */
export function computeGantt(plan: Plan): GanttModel {
  const hpd = Number(plan.hoursPerDay) || 8;
  const ordered = priorityJobs(plan);
  const capMap = capacityLookup(plan);
  const capped = cappedProjectRoles(plan);
  const used = new Map<string, number>();
  const endByJob = new Map<string, number>();
  const rows: GanttRow[] = [];
  const roleColors = new Map<string, string>();
  const categoryColors = new Map<string, string>();
  const warnings: string[] = [];
  let colorI = 0;
  let maxEndDay = 0;
  let totalHours = 0;

  const categoryTones = buildCategoryToneMap(plan.jobs.map((j) => activityCategory(j.name)));

  let currentStage = 0;
  let stageMaxEnd = 0;
  let forcedEarliest = 0;

  for (const { job, stage } of ordered) {
    if (stage !== currentStage) {
      if (currentStage > 0) {
        const prev = plan.stages[currentStage - 1];
        const gap = Math.max(0, Math.round(Number(prev?.gapAfterDays) || 0));
        if (gap > 0) {
          forcedEarliest = Math.max(forcedEarliest, Math.ceil(stageMaxEnd - 1e-9) + gap);
        }
      }
      currentStage = stage;
      stageMaxEnd = 0;
    }

    const project = job.project || DEFAULT_PROJECT;
    const role = normalizeRole(job.role);
    const people = jobPeople(job);
    const durationDays = jobDurationDays(job, hpd);
    totalHours += job.hours;

    const pair = pairKey(project, role);
    const hasCap = capped.has(pair);
    if (plan.weeklyCapacities.length > 0 && !hasCap) {
      warnings.push(
        `${job.name}: kapasite yok — proje “${project}” ve tip “${role}” haftalık personel tablosuyla birebir aynı olmalı`
      );
    }

    let earliest = forcedEarliest;
    for (const predId of predecessorsOf(plan, job.id)) {
      const predEnd = endByJob.get(predId);
      if (predEnd == null) {
        warnings.push(`${job.name}: öncül henüz yerleşmedi — öncelik sırasını kontrol edin`);
        continue;
      }
      // Öncül biter bitmez sonraki tam iş günü (ondalık gün anahtarı kapasite kaçırıyordu)
      earliest = Math.max(earliest, Math.ceil(predEnd - 1e-9));
    }

    let startDay = earliest;
    let placed = false;
    const maxSearch = 520 * WORK_DAYS_PER_WEEK;
    for (let d = earliest; d < earliest + maxSearch; d++) {
      if (canPlaceAtDay(plan, capMap, capped, used, project, role, d, durationDays, people)) {
        startDay = d;
        placed = true;
        break;
      }
    }

    const usedBefore = used.get(`${pair}::d${dayIndex(startDay)}`) ?? 0;
    const lane = Math.floor(usedBefore) + 1;

    if (placed) {
      placeJobAtDay(used, project, role, startDay, durationDays, people);
    } else {
      warnings.push(
        `${job.name} (${role}): uygun kapasite bulunamadı (haftalık kişi yetersiz veya tanımsız)`
      );
      startDay = earliest;
    }

    const endDay = dayIndex(startDay) + durationDays;
    endByJob.set(job.id, endDay);
    stageMaxEnd = Math.max(stageMaxEnd, endDay);

    if (!roleColors.has(role)) {
      roleColors.set(role, BAR_COLORS[colorI % BAR_COLORS.length]);
      colorI += 1;
    }

    const category = activityCategory(job.name);
    const tone = categoryTones.get(category) ?? {
      label: "none",
      bar: roleColors.get(role)!,
      bg: "#f8fafc",
      border: "#e2e8f0",
    };
    if (category && !categoryColors.has(category)) {
      categoryColors.set(category, tone.bar);
    }

    maxEndDay = Math.max(maxEndDay, endDay);

    rows.push({
      job,
      stage,
      lane,
      startDay,
      durationDays,
      color: category ? tone.bar : roleColors.get(role)!,
      tint: tone.bg,
      category,
      capacityOk: placed,
    });
  }

  return {
    rows,
    totalHours,
    totalDays: maxEndDay,
    roleColors,
    categoryColors,
    warnings,
  };
}

export function formatHours(h: number): string {
  const v = Math.round(h * 10) / 10;
  return Number.isInteger(v) ? `${v} sa` : `${v} sa`;
}

export function stageDurationLabel(plan: Plan, stage: Stage): string {
  const model = computeGantt({
    ...plan,
    stages: [stage],
  });
  if (!model.rows.length) return "—";
  const days = Math.max(...model.rows.map((r) => r.startDay + r.durationDays));
  return `${(days / WORK_DAYS_PER_WEEK).toFixed(1)} hf`;
}

/** Gantt yerleşimine göre hafta aralığı: 2026-35 veya 2026-35/36 */
export function formatGanttSpanWeeks(
  planStartIso: string,
  startDay: number,
  durationDays: number
): string | null {
  if (!Number.isFinite(startDay) || !Number.isFinite(durationDays) || durationDays <= 0) return null;
  const startOff = Math.max(0, Math.floor(startDay / WORK_DAYS_PER_WEEK));
  const endOff = Math.max(
    startOff,
    Math.floor((startDay + durationDays - 1e-9) / WORK_DAYS_PER_WEEK)
  );
  const a = isoWeekAtOffset(planStartIso, startOff);
  const b = isoWeekAtOffset(planStartIso, endOff);
  const aw = String(a.week).padStart(2, "0");
  const bw = String(b.week).padStart(2, "0");
  if (a.year === b.year) {
    if (a.week === b.week) return `${a.year}-${aw}`;
    return `${a.year}-${aw}/${bw}`;
  }
  return `${a.year}-${aw}/${b.year}-${bw}`;
}

/** Aşamadaki işlerin Gantt birleşik hafta aralığı; satır yoksa null */
export function stageGanttWeekRangeLabel(
  planStartIso: string,
  stage: Stage,
  ganttRows: GanttRow[]
): string | null {
  const ids = new Set(stage.jobIds);
  const rows = ganttRows.filter((r) => ids.has(r.job.id));
  if (!rows.length) return null;
  let minStart = Infinity;
  let maxEnd = -Infinity;
  for (const r of rows) {
    minStart = Math.min(minStart, r.startDay);
    maxEnd = Math.max(maxEnd, r.startDay + r.durationDays);
  }
  if (!Number.isFinite(minStart) || !(maxEnd > minStart)) return null;
  return formatGanttSpanWeeks(planStartIso, minStart, maxEnd - minStart);
}

export function addDaysLabel(iso: string, days: number): string {
  const d = new Date(`${iso}T00:00:00`);
  d.setDate(d.getDate() + Math.floor(days));
  return d.toLocaleDateString("tr-TR");
}
