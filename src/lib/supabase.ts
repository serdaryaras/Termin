import { createClient, type SupabaseClient } from "@supabase/supabase-js";
import { emptyPlan, normalizePlan, type Plan } from "./types";

const PLAN_SLUG = process.env.NEXT_PUBLIC_PLAN_SLUG || "default";
const LOCAL_KEY = "arti-gantt-plan-v1";

function envUrl(): string {
  return (process.env.NEXT_PUBLIC_SUPABASE_URL || "").trim();
}

function envKey(): string {
  return (process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY || "").trim();
}

export function isSupabaseConfigured(): boolean {
  const url = envUrl();
  const key = envKey();
  if (!url || !key) return false;
  if (url.includes("YOUR_PROJECT") || key === "your-anon-key") return false;
  try {
    const parsed = new URL(url);
    return parsed.protocol === "https:" && parsed.hostname.endsWith(".supabase.co");
  } catch {
    return false;
  }
}

export function getSupabase(): SupabaseClient | null {
  if (!isSupabaseConfigured()) return null;
  return createClient(envUrl(), envKey());
}

export function loadLocalPlan(): Plan | null {
  if (typeof window === "undefined") return null;
  try {
    const raw = localStorage.getItem(LOCAL_KEY);
    if (!raw) return null;
    const data = JSON.parse(raw) as Partial<Plan>;
    return normalizePlan(data);
  } catch {
    return null;
  }
}

export function saveLocalPlan(plan: Plan) {
  if (typeof window === "undefined") return;
  localStorage.setItem(LOCAL_KEY, JSON.stringify(plan));
}

type PlanRow = {
  name: string;
  start_date: string;
  hours_per_day: number;
  jobs: Plan["jobs"];
  stages: Plan["stages"];
  dependencies?: Plan["dependencies"];
  resource_groups?: Plan["resourceGroups"];
  weekly_capacities?: Plan["weeklyCapacities"];
  job_progress?: Plan["jobProgress"];
};

export async function loadPlan(): Promise<{ plan: Plan; source: "supabase" | "local" | "empty"; error?: string }> {
  const local = loadLocalPlan();
  const supabase = getSupabase();
  if (!supabase) {
    if (local) return { plan: local, source: "local" };
    return { plan: emptyPlan(), source: "empty" };
  }

  try {
    const { data, error } = await supabase
      .from("gantt_plans")
      .select(
        "name,start_date,hours_per_day,jobs,stages,dependencies,resource_groups,weekly_capacities,job_progress"
      )
      .eq("slug", PLAN_SLUG)
      .maybeSingle();

    if (error) {
      return {
        plan: local || emptyPlan(),
        source: local ? "local" : "empty",
        error: error.message,
      };
    }
    if (!data) {
      return { plan: local || emptyPlan(), source: local ? "local" : "empty" };
    }
    const row = data as PlanRow;
    return {
      plan: normalizePlan({
        name: row.name || "İş planı",
        startDate: row.start_date,
        hoursPerDay: Number(row.hours_per_day) || 8,
        jobs: Array.isArray(row.jobs) ? row.jobs : [],
        stages: Array.isArray(row.stages) ? row.stages : [],
        dependencies: Array.isArray(row.dependencies) ? row.dependencies : [],
        resourceGroups: Array.isArray(row.resource_groups) ? row.resource_groups : [],
        weeklyCapacities: Array.isArray(row.weekly_capacities) ? row.weekly_capacities : [],
        jobProgress: Array.isArray(row.job_progress) ? row.job_progress : [],
      }),
      source: "supabase",
    };
  } catch (err) {
    const message = err instanceof Error ? err.message : "Supabase’e bağlanılamadı";
    return { plan: local || emptyPlan(), source: local ? "local" : "empty", error: message };
  }
}

export async function savePlan(plan: Plan): Promise<{ ok: boolean; message?: string }> {
  saveLocalPlan(plan);
  const supabase = getSupabase();
  if (!supabase) return { ok: true, message: "Tarayıcıda kaydedildi" };

  const payload = {
    slug: PLAN_SLUG,
    name: plan.name,
    start_date: plan.startDate,
    hours_per_day: plan.hoursPerDay,
    jobs: plan.jobs,
    stages: plan.stages,
    dependencies: plan.dependencies || [],
    resource_groups: plan.resourceGroups,
    weekly_capacities: plan.weeklyCapacities,
    job_progress: plan.jobProgress || [],
    updated_at: new Date().toISOString(),
  };

  try {
    const { error } = await supabase.from("gantt_plans").upsert(payload, { onConflict: "slug" });
    if (error) return { ok: false, message: error.message };
    return { ok: true };
  } catch (err) {
    const message = err instanceof Error ? err.message : "Supabase’e bağlanılamadı";
    return { ok: false, message };
  }
}
