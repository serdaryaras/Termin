"use client";

import { useCallback, useEffect, useMemo, useRef, useState, type DragEvent, type MouseEvent } from "react";
import { flushSync } from "react-dom";
import { clipboardToText, looksLikeActivityCode, parseCapacityText, parseExcelText } from "@/lib/excel";
import { exportElementToPdf } from "@/lib/export-pdf";
import { ARTI_LOGO, artiLogoDisplayWidth } from "@/lib/arti-logo";
import {
  addDependency,
  addJobToStage,
  buildWeekTicks,
  clearAllJobProgress,
  clearAllJobs,
  clearAllWeeklyCapacities,
  clearWeeklyCapacitiesForProject,
  clearDependenciesForJobs,
  clearEntirePlan,
  clearJobProgressForWeek,
  clearPriorityOrder,
  computeGantt,
  deleteJob,
  earliestCapacityStartDate,
  findPlacement,
  formatHours,
  formatWeekLabel,
  formatWeekOnly,
  ganttRowOverlapsIsoWeek,
  getIsoWeekParts,
  getWeeklyCapacityPeople,
  jobById,
  moveJobsByDelta,
  moveJobsDir,
  moveJobsGroupBefore,
  nextIsoWeek,
  parseIsoDate,
  predecessorsOf,
  prevIsoWeek,
  removeDependency,
  removeWeeklyCapacity,
  replaceWeeklyCapacity,
  scheduledIds,
  setStageGapAfterDays,
  stageGanttWeekRangeLabel,
  successorsOf,
  upsertWeeklyCapacity,
  weekIndexFromPlanStart,
  weekRangeLabel,
  WORK_DAYS_PER_WEEK,
} from "@/lib/schedule";
import { isSupabaseConfigured, loadPlan, saveLocalPlan, savePlan } from "@/lib/supabase";
import {
  capacityKey,
  DEFAULT_PROJECT,
  activityCategory,
  buildCategoryToneMap,
  categoryTone,
  emptyPlan,
  jobKey,
  listProjects,
  newId,
  normalizeRole,
  ROLE_OPTIONS,
  type Plan,
  type ProgressPercent,
} from "@/lib/types";

/** Tıklama döngüsü: 100 → 75 → 50 → 25 → 0 → 100… */
const PROGRESS_CYCLE: ProgressPercent[] = [100, 75, 50, 25, 0];

function nextProgressPercent(current: ProgressPercent | undefined): ProgressPercent {
  if (current == null) return 100;
  const i = PROGRESS_CYCLE.indexOf(current);
  return PROGRESS_CYCLE[(i < 0 ? 0 : i + 1) % PROGRESS_CYCLE.length];
}

function parseTrackWeekLabel(raw: string): { year: number; week: number } | null {
  const m = String(raw)
    .trim()
    .match(/^(\d{4})\s*[-./,]\s*(\d{1,2})$/);
  if (!m) return null;
  const year = Number(m[1]);
  const week = Number(m[2]);
  if (!Number.isFinite(year) || !Number.isFinite(week) || week < 1 || week > 53) return null;
  return { year, week };
}

function progressButtonClass(percent: ProgressPercent | undefined): string {
  if (percent == null) {
    return "border border-dashed border-[var(--card-border)] bg-[var(--background)] text-[var(--muted)]";
  }
  if (percent === 100) return "bg-emerald-100 text-emerald-900 hover:bg-emerald-200";
  if (percent === 75) return "bg-sky-100 text-sky-900 hover:bg-sky-200";
  if (percent === 50) return "bg-amber-100 text-amber-900 hover:bg-amber-200";
  if (percent === 25) return "bg-orange-100 text-orange-900 hover:bg-orange-200";
  return "bg-rose-100 text-rose-900 hover:bg-rose-200";
}

export function GanttPlanner() {
  const [plan, setPlan] = useState<Plan>(emptyPlan);
  const [selected, setSelected] = useState<string | null>(null);
  /** CTRL ile seçim sırası — ardışık çiftler öncül→ardıl olur */
  const [selectChain, setSelectChain] = useState<string[]>([]);
  const [pasteText, setPasteText] = useState("");
  const [pasteStatus, setPasteStatus] = useState("");
  const [pasteError, setPasteError] = useState(false);
  const [capPaste, setCapPaste] = useState("");
  const [capStatus, setCapStatus] = useState("");
  const [capError, setCapError] = useState(false);
  const [capProject, setCapProject] = useState("");
  const [capWeek, setCapWeek] = useState("2026-35");
  const [capRole, setCapRole] = useState("Donatım");
  const [capPeople, setCapPeople] = useState(3);
  const [capDeleteProject, setCapDeleteProject] = useState("");
  /** Kapasite satırı seçimleri (capacityKey) — Haftalar ±1 yalnız bunlara */
  const [capSelected, setCapSelected] = useState<string[]>([]);
  /** Kişi değişiklikleri Güncelle’ye kadar taslakta; Gantt / DB henüz değişmez */
  const [capDraft, setCapDraft] = useState<Record<string, number>>({});
  const [projectFilter, setProjectFilter] = useState<string>("all");
  const groupByProject = false;
  const [saveStatus, setSaveStatus] = useState("Yükleniyor…");
  const [overId, setOverId] = useState<string | null>(null);
  const [pdfStatus, setPdfStatus] = useState("");
  const [pdfCompact, setPdfCompact] = useState(false);
  const [activeTab, setActiveTab] = useState<"jobs" | "priority" | "gantt" | "track">("jobs");
  const [prioritySearch, setPrioritySearch] = useState("");
  const [editingJobId, setEditingJobId] = useState<string | null>(null);
  const [ganttSelectIds, setGanttSelectIds] = useState<string[]>([]);
  const [trackWeek, setTrackWeek] = useState(() => {
    const { year, week } = getIsoWeekParts(new Date());
    return formatWeekLabel(year, week);
  });
  const loaded = useRef(false);
  const dragRef = useRef<{ jobId: string } | null>(null);
  const ganttExportRef = useRef<HTMLDivElement | null>(null);
  const supabaseOn = isSupabaseConfigured();

  useEffect(() => {
    let cancelled = false;
    (async () => {
      const data = await loadPlan();
      if (cancelled) return;
      setPlan(data.plan);
      loaded.current = true;
      if (data.error) {
        setSaveStatus(`Supabase bağlanamadı · ${data.error}`);
      } else if (data.source === "supabase") {
        setSaveStatus("Supabase bağlı");
      } else {
        setSaveStatus("Bu cihazda kaydedilir (Supabase henüz yok)");
      }
    })();
    return () => {
      cancelled = true;
    };
  }, []);

  useEffect(() => {
    if (!loaded.current) return;
    const t = setTimeout(async () => {
      setSaveStatus("Kaydediliyor…");
      const result = await savePlan(plan);
      if (result.ok) {
        setSaveStatus(supabaseOn ? "Kaydedildi" : "Bu cihazda kaydedildi");
      } else {
        setSaveStatus(result.message || "Kayıt hatası");
      }
    }, 700);
    return () => clearTimeout(t);
  }, [plan, supabaseOn]);

  useEffect(() => {
    const on = activeTab === "gantt" || activeTab === "priority";
    document.body.dataset.ganttFocus = on ? "1" : "";
    return () => {
      delete document.body.dataset.ganttFocus;
    };
  }, [activeTab]);

  useEffect(() => {
    if (activeTab !== "priority") return;
    const onKey = (e: KeyboardEvent) => {
      const t = e.target as HTMLElement | null;
      if (
        t &&
        (t.tagName === "INPUT" ||
          t.tagName === "TEXTAREA" ||
          t.tagName === "SELECT" ||
          t.isContentEditable)
      ) {
        return;
      }
      const ids =
        selectChain.length > 0 ? selectChain : selected ? [selected] : [];
      if (!ids.length) return;

      const go = (dir: "top" | "bottom" | "up" | "down") => {
        e.preventDefault();
        setPlan((p) => moveJobsDir(p, ids, dir));
        requestAnimationFrame(() => {
          document.getElementById(`priority-job-${ids[ids.length - 1]}`)?.scrollIntoView({
            behavior: "smooth",
            block: "center",
          });
        });
      };

      if (e.key === "ArrowUp") {
        go(e.shiftKey ? "top" : "up");
      } else if (e.key === "ArrowDown") {
        go(e.shiftKey ? "bottom" : "down");
      } else if (e.key === "Home") {
        go("top");
      } else if (e.key === "End") {
        go("bottom");
      }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [activeTab, selectChain, selected]);

  const gantt = useMemo(() => computeGantt(plan), [plan]);
  const categoryToneMap = useMemo(
    () => buildCategoryToneMap(plan.jobs.map((j) => activityCategory(j.name))),
    [plan.jobs]
  );
  const place = selected ? findPlacement(plan, selected) : null;
  const ganttGapPlace = useMemo(() => {
    const id = ganttSelectIds.length ? ganttSelectIds[ganttSelectIds.length - 1]! : null;
    return id ? findPlacement(plan, id) : null;
  }, [ganttSelectIds, plan]);
  const linkTargetIds = useMemo(() => {
    if (selectChain.length) return selectChain;
    return selected ? [selected] : [];
  }, [selectChain, selected]);
  const selectedStageIndices = useMemo(() => {
    const idxs: number[] = [];
    for (const id of linkTargetIds) {
      const p = findPlacement(plan, id);
      if (p) idxs.push(p.stageIndex);
    }
    return idxs;
  }, [linkTargetIds, plan]);
  const canShiftUp =
    selectedStageIndices.length > 0 && Math.min(...selectedStageIndices) > 0;
  const canShiftDown =
    selectedStageIndices.length > 0 &&
    Math.max(...selectedStageIndices) < plan.stages.length - 1;
  const canClearSelectedLinks = useMemo(() => {
    return linkTargetIds.some(
      (id) => predecessorsOf(plan, id).length > 0 || successorsOf(plan, id).length > 0
    );
  }, [linkTargetIds, plan]);
  const projects = useMemo(
    () => listProjects(plan.jobs, plan.weeklyCapacities),
    [plan.jobs, plan.weeklyCapacities]
  );

  const planYear = useMemo(() => getIsoWeekParts(parseIsoDate(plan.startDate)).year, [plan.startDate]);

  const filteredGantt = useMemo(() => {
    if (projectFilter === "all") return gantt;
    return {
      ...gantt,
      rows: gantt.rows.filter((r) => (r.job.project || DEFAULT_PROJECT) === projectFilter),
    };
  }, [gantt, projectFilter]);

  /** Öncelik grid’i: proje + arama görünümü daraltır; si gerçek sıra indeksidir */
  const filteredPriorityStages = useMemo(() => {
    const q = prioritySearch.trim().toLocaleLowerCase("tr");
    const out: { si: number; jobIds: string[] }[] = [];
    plan.stages.forEach((stage, si) => {
      let jobIds = stage.jobIds.filter((id) => {
        const job = jobById(plan, id);
        if (!job) return false;
        if (projectFilter !== "all" && (job.project || DEFAULT_PROJECT) !== projectFilter) {
          return false;
        }
        return true;
      });
      if (!jobIds.length) return;
      if (q) {
        jobIds = jobIds.filter((id) => {
          const job = jobById(plan, id);
          if (!job) return false;
          const week = stageGanttWeekRangeLabel(plan.startDate, stage, gantt.rows) || "";
          const hay = [
            job.name,
            job.role,
            job.project || "",
            `A${si + 1}`,
            week,
            formatHours(job.hours),
          ]
            .join(" ")
            .toLocaleLowerCase("tr");
          return hay.includes(q);
        });
      }
      if (jobIds.length) out.push({ si, jobIds });
    });
    return out;
  }, [plan, prioritySearch, projectFilter, gantt.rows]);

  const sortedCapacities = useMemo(() => {
    const list = [...plan.weeklyCapacities].sort((a, b) => {
      const pc = a.project.localeCompare(b.project, "tr");
      if (pc !== 0) return pc;
      const rc = a.role.localeCompare(b.role, "tr");
      if (rc !== 0) return rc;
      if (a.year !== b.year) return a.year - b.year;
      return a.week - b.week;
    });
    if (!capDeleteProject || capDeleteProject === "all") return list;
    return list.filter((c) => (c.project || DEFAULT_PROJECT) === capDeleteProject);
  }, [plan.weeklyCapacities, capDeleteProject]);

  const capacityProjects = useMemo(() => {
    const set = new Set<string>();
    for (const c of plan.weeklyCapacities) {
      set.add(c.project || DEFAULT_PROJECT);
    }
    return [...set].sort((a, b) => a.localeCompare(b, "tr"));
  }, [plan.weeklyCapacities]);

  useEffect(() => {
    if (!capacityProjects.length) {
      if (capDeleteProject) setCapDeleteProject("");
      return;
    }
    if (capDeleteProject === "all") return;
    if (!capacityProjects.includes(capDeleteProject)) {
      setCapDeleteProject(capacityProjects[0]!);
    }
  }, [capacityProjects, capDeleteProject]);

  useEffect(() => {
    if (capDeleteProject && capDeleteProject !== "all") setCapProject(capDeleteProject);
    setCapSelected([]);
  }, [capDeleteProject]);

  const capRowKey = useCallback(
    (c: { project: string; year: number; week: number; role: string }) =>
      capacityKey(c.project, Number(c.year), Number(c.week), normalizeRole(c.role)),
    []
  );

  const capPeopleOf = useCallback(
    (c: { project: string; year: number; week: number; role: string; people: number }) => {
      const k = capRowKey(c);
      return k in capDraft ? capDraft[k]! : c.people;
    },
    [capDraft, capRowKey]
  );

  const bumpSelectedCapWeeks = useCallback(
    (delta: number) => {
      if (!capSelected.length) {
        setCapStatus("Önce hafta kutucuklarından seçim yapın.");
        setCapError(true);
        return;
      }
      setCapDraft((prev) => {
        const next = { ...prev };
        for (const key of capSelected) {
          const row = plan.weeklyCapacities.find((c) => capRowKey(c) === key);
          if (!row) continue;
          const base = key in next ? next[key]! : row.people;
          next[key] = Math.max(0, base + delta);
        }
        return next;
      });
      setCapStatus(
        `${capSelected.length} hafta ${delta > 0 ? "+" : ""}${delta} kişi (taslak). Güncelle ile Gantt ve veritabanına yazın.`
      );
      setCapError(false);
    },
    [capSelected, plan.weeklyCapacities, capRowKey]
  );

  const trackWeekParts = useMemo(() => parseTrackWeekLabel(trackWeek), [trackWeek]);

  const trackRows = useMemo(() => {
    const rows: { jobId: string; stageIndex: number; seq: number }[] = [];
    if (!trackWeekParts) return rows;
    const inWeek = new Set(
      gantt.rows
        .filter((r) =>
          ganttRowOverlapsIsoWeek(
            r,
            plan.startDate,
            trackWeekParts.year,
            trackWeekParts.week
          )
        )
        .map((r) => r.job.id)
    );
    let seq = 0;
    plan.stages.forEach((stage, stageIndex) => {
      for (const jobId of stage.jobIds) {
        if (!inWeek.has(jobId)) continue;
        const job = jobById(plan, jobId);
        if (!job) continue;
        if (projectFilter !== "all" && (job.project || DEFAULT_PROJECT) !== projectFilter) continue;
        seq += 1;
        rows.push({ jobId, stageIndex, seq });
      }
    });
    return rows;
  }, [plan, projectFilter, gantt.rows, trackWeekParts]);

  const shiftTrackWeek = useCallback((delta: -1 | 1) => {
    const base =
      parseTrackWeekLabel(trackWeek) || getIsoWeekParts(new Date());
    const next =
      delta > 0
        ? nextIsoWeek(base.year, base.week)
        : prevIsoWeek(base.year, base.week);
    setTrackWeek(formatWeekLabel(next.year, next.week));
  }, [trackWeek]);

  const cycleTrackProgress = useCallback(
    (jobId: string) => {
      if (!trackWeekParts) return;
      const { year, week } = trackWeekParts;
      setPlan((prev) => {
        const list = [...(prev.jobProgress || [])];
        const idx = list.findIndex((p) => p.jobId === jobId && p.year === year && p.week === week);
        if (idx < 0) {
          list.push({ jobId, year, week, percent: 100, reason: "" });
        } else {
          list[idx] = {
            ...list[idx],
            percent: nextProgressPercent(list[idx].percent),
          };
        }
        return { ...prev, jobProgress: list };
      });
    },
    [trackWeekParts]
  );

  const setTrackReason = useCallback(
    (jobId: string, reason: string) => {
      if (!trackWeekParts) return;
      const { year, week } = trackWeekParts;
      setPlan((prev) => {
        const list = [...(prev.jobProgress || [])];
        const idx = list.findIndex((p) => p.jobId === jobId && p.year === year && p.week === week);
        if (idx < 0) {
          list.push({ jobId, year, week, percent: 0, reason });
        } else {
          list[idx] = { ...list[idx], reason };
        }
        return { ...prev, jobProgress: list };
      });
    },
    [trackWeekParts]
  );

  const forceSave = useCallback(async () => {
    setSaveStatus("Kaydediliyor…");
    const result = await savePlan(plan);
    if (result.ok) {
      setSaveStatus(supabaseOn ? "Kaydedildi (veritabanı)" : "Bu cihazda kaydedildi");
    } else {
      setSaveStatus(result.message || "Kayıt hatası");
    }
  }, [plan, supabaseOn]);

  const applyCapDraftAndSave = useCallback(async () => {
    const keys = Object.keys(capDraft);
    if (!keys.length) {
      setCapStatus("Uygulanacak kapasite değişikliği yok. Önce hafta seçip ±1 yapın.");
      setCapError(true);
      return;
    }
    let next = plan;
    for (const c of plan.weeklyCapacities) {
      const k = capRowKey(c);
      if (!(k in capDraft)) continue;
      next = upsertWeeklyCapacity(next, {
        project: c.project,
        year: c.year,
        week: c.week,
        role: c.role,
        people: capDraft[k]!,
      });
    }
    setPlan(next);
    setCapDraft({});
    setCapSelected([]);
    setSaveStatus("Kaydediliyor…");
    const result = await savePlan(next);
    if (result.ok) {
      setSaveStatus(supabaseOn ? "Kaydedildi (veritabanı)" : "Bu cihazda kaydedildi");
      setCapStatus(
        `${keys.length} hafta güncellendi · Gantt yenilendi · ${
          supabaseOn ? "veritabanına yazıldı" : "bu cihazda kaydedildi"
        }.`
      );
      setCapError(false);
    } else {
      setSaveStatus(result.message || "Kayıt hatası");
      setCapStatus(result.message || "Veritabanı kaydı başarısız.");
      setCapError(true);
    }
  }, [capDraft, plan, capRowKey, supabaseOn]);

  const wipeEntirePlan = useCallback(async () => {
    if (
      !confirm(
        "Tüm plan verisi (işler, kapasite, öncelik, takip) silinsin mi?\nBu işlem veritabanına da yazılır."
      )
    ) {
      return;
    }
    const cleared = clearEntirePlan(plan);
    setPlan(cleared);
    setSelected(null);
    setSelectChain([]);
    setPasteStatus("");
    setCapStatus("");
    saveLocalPlan(cleared);
    setSaveStatus("Kaydediliyor…");
    const result = await savePlan(cleared);
    setSaveStatus(
      result.ok
        ? supabaseOn
          ? "Tüm veri silindi · veritabanı güncellendi"
          : "Tüm veri silindi · bu cihazda kaydedildi"
        : result.message || "Silme kaydı başarısız"
    );
  }, [plan, supabaseOn]);

  const importPaste = (text: string) => {
    const { rows, skipped } = parseExcelText(text);
    if (!rows.length) {
      setPasteError(true);
      setPasteStatus(
        skipped
          ? `${skipped} satır okunamadı. Sütunlar: proje, iş kalemi, personel tipi, saat.`
          : "Yapıştırılacak satır bulunamadı."
      );
      return;
    }
    const fallbackProject =
      (projectFilter !== "all" ? projectFilter : "") ||
      plan.weeklyCapacities[0]?.project ||
      projects[0] ||
      DEFAULT_PROJECT;
    const jobs = [...plan.jobs];
    const stages = plan.stages.map((s) => ({ jobIds: [...s.jobIds] }));
    const byKey = new Map(jobs.map((j) => [jobKey(j.project || DEFAULT_PROJECT, j.name), j]));
    const scheduled = scheduledIds({ ...plan, stages });
    let added = 0;
    let queued = 0;
    let dup = 0;
    let filledProject = 0;

    rows.forEach((row) => {
      let project = (row.project || "").trim();
      if (!project || project === DEFAULT_PROJECT || looksLikeActivityCode(project)) {
        project = fallbackProject;
        filledProject += 1;
      }
      const key = jobKey(project, row.name);
      let job = byKey.get(key);
      if (!job) {
        job = {
          id: newId(),
          name: row.name,
          project,
          role: row.role,
          hours: row.hours,
          people: row.people,
        };
        jobs.push(job);
        byKey.set(key, job);
        added += 1;
      } else {
        dup += 1;
      }
      if (!scheduled.has(job.id)) {
        stages.push({ jobIds: [job.id] });
        scheduled.add(job.id);
        queued += 1;
      }
    });

    setPlan({ ...plan, jobs, stages });
    setPasteError(queued === 0 && added === 0);
    const bits: string[] = [];
    if (queued) bits.push(`${queued} kalem sıraya eklendi (yapıştırma sırası)`);
    if (added && added !== queued) bits.push(`${added} yeni iş tanımı`);
    if (filledProject) bits.push(`proje: ${fallbackProject}`);
    if (dup && queued < rows.length) bits.push(`${rows.length - queued} zaten sırada / tekrar`);
    if (skipped) bits.push(`${skipped} satır atlandı`);
    setPasteStatus((bits.join(" · ") || "Değişiklik yok") + ".");
  };

  const importCapacity = (text: string) => {
    const { rows, skipped, duplicates } = parseCapacityText(text, planYear);
    if (!rows.length) {
      setCapError(true);
      setCapStatus(
        skipped
          ? `${skipped} satır okunamadı. Beklenen: proje | hafta (2026-35) | Donatım/Konstrüksiyon | kişi`
          : "Kapasite satırı bulunamadı."
      );
      return;
    }
    let next = plan;
    const roles = new Set<string>();
    rows.forEach((row) => {
      roles.add(row.role);
      next = upsertWeeklyCapacity(next, row);
    });
    const aligned = earliestCapacityStartDate(next);
    if (aligned) next = { ...next, startDate: aligned };
    setPlan(next);
    setCapError(false);
    setCapStatus(
      `${rows.length} kayıt uygulandı · tipler: ${[...roles].join(", ")}` +
        (duplicates ? ` · ${duplicates} mükerrer satır yok sayıldı (son değer alındı)` : "") +
        (skipped ? ` · ${skipped} atlandı` : "") +
        (aligned ? ` · eksen ${aligned}` : "") +
        "."
    );
  };

  const addCapacityManual = () => {
    const project = (capProject.trim() || projects[0] || DEFAULT_PROJECT).trim();
    const role = normalizeRole(capRole);
    const raw = capWeek.trim();
    const m = raw.match(/^(\d{4})[./-](\d{1,2})$/);
    const year = m ? Number(m[1]) : planYear;
    let week = m ? Number(m[2]) : Number(raw);
    if (m && m[2].length === 1 && week >= 1 && week <= 5) week = week * 10;
    if (!project || role === "Belirtilmedi" || !(week >= 1 && week <= 53) || !(capPeople >= 0)) return;
    let next = upsertWeeklyCapacity(plan, { project, year, week, role, people: capPeople });
    const aligned = earliestCapacityStartDate(next);
    if (aligned) next = { ...next, startDate: aligned };
    setPlan(next);
    setCapStatus(
      `${project} · ${formatWeekLabel(year, week)} · ${role} · ${capPeople} kişi eklendi${
        aligned ? ` · eksen ${aligned}` : ""
      }.`
    );
    setCapError(false);
  };

  const exportGanttPdf = async () => {
    if (filteredGantt.rows.length === 0) {
      setPdfStatus("Önce Gantt’ta satır olmalı.");
      return;
    }
    setPdfStatus("PDF hazırlanıyor…");
    try {
      flushSync(() => {
        setActiveTab("gantt");
        setPdfCompact(true);
      });
      await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(() => r(undefined))));
      const exportRoot = ganttExportRef.current;
      if (!exportRoot) {
        setPdfStatus("Gantt alanı bulunamadı.");
        return;
      }
      await Promise.all(
        [...exportRoot.querySelectorAll("img")].map((img) =>
          img.complete
            ? Promise.resolve()
            : new Promise<void>((res) => {
                img.onload = () => res();
                img.onerror = () => res();
              })
        )
      );
      const stamp = new Date().toISOString().slice(0, 10);
      const name =
        projectFilter !== "all"
          ? `Gantt_${projectFilter}_${stamp}`
          : `Gantt_${plan.name || "plan"}_${stamp}`;
      const safe = name.replace(/[^\w.\-ğüşıöçĞÜŞİÖÇ]+/gi, "_") || `Gantt_${stamp}`;
      await exportElementToPdf(exportRoot, safe);
      setPdfStatus("PDF indirildi.");
    } catch (err) {
      console.error("PDF export failed", err);
      setPdfStatus(err instanceof Error ? err.message : "PDF oluşturulamadı.");
    } finally {
      setPdfCompact(false);
    }
  };

  const onPriorityClick = (jobId: string, e: MouseEvent) => {
    if (e.shiftKey) {
      e.preventDefault();
      setEditingJobId(jobId);
      setSelectChain([jobId]);
      setSelected(jobId);
      return;
    }
    if (editingJobId && editingJobId !== jobId) {
      setEditingJobId(null);
    }
    if (e.ctrlKey || e.metaKey) {
      e.preventDefault();
      setSelectChain((prev) => {
        if (prev.includes(jobId)) return prev;
        const next = [...prev, jobId];
        if (prev.length > 0) {
          const pred = prev[prev.length - 1];
          setPlan((p) => addDependency(p, pred, jobId));
        }
        return next;
      });
      setSelected(jobId);
      return;
    }
    // Gantt gibi: tıkla = seçime ekle / çıkar
    setSelectChain((prev) => {
      if (prev.includes(jobId)) {
        const next = prev.filter((x) => x !== jobId);
        setSelected(next.length ? next[next.length - 1]! : null);
        return next;
      }
      setSelected(jobId);
      return [...prev, jobId];
    });
  };

  const patchJob = (jobId: string, patch: Partial<{ name: string; role: string; hours: number; people: number; project: string }>) => {
    setPlan((p) => ({
      ...p,
      jobs: p.jobs.map((j) => {
        if (j.id !== jobId) return j;
        return {
          ...j,
          name: patch.name != null ? patch.name : j.name,
          role: patch.role != null ? normalizeRole(patch.role) || j.role : j.role,
          hours:
            patch.hours != null && Number.isFinite(patch.hours)
              ? Math.max(0.5, patch.hours)
              : j.hours,
          people:
            patch.people != null && Number.isFinite(patch.people)
              ? Math.max(1, Math.round(patch.people))
              : j.people,
          project: patch.project != null ? patch.project.trim() || j.project : j.project,
        };
      }),
    }));
  };

  const scrollPriorityJobIntoView = (jobId: string) => {
    requestAnimationFrame(() => {
      requestAnimationFrame(() => {
        document.getElementById(`priority-job-${jobId}`)?.scrollIntoView({
          behavior: "smooth",
          block: "center",
          inline: "nearest",
        });
      });
    });
  };

  const shiftSelectedBy = (delta: number) => {
    const ids = linkTargetIds;
    if (!ids.length) return;
    setPlan(moveJobsByDelta(plan, ids, delta));
    scrollPriorityJobIntoView(ids[ids.length - 1]!);
  };

  const shiftSelectedDir = (dir: "top" | "bottom" | "up" | "down") => {
    const ids = linkTargetIds;
    if (!ids.length) return;
    setPlan(moveJobsDir(plan, ids, dir));
    scrollPriorityJobIntoView(ids[ids.length - 1]!);
  };

  const onDragStart = (jobId: string) => (e: DragEvent) => {
    window.getSelection()?.removeAllRanges();
    dragRef.current = { jobId };
    e.dataTransfer.effectAllowed = "move";
    e.dataTransfer.setData("text/plain", jobId);
  };

  const dropJob = useCallback(
    (stageIndex: number, asNewStage: boolean) => (e: DragEvent) => {
      e.preventDefault();
      e.stopPropagation();
      setOverId(null);
      const jobId = dragRef.current?.jobId || e.dataTransfer.getData("text/plain");
      dragRef.current = null;
      if (!jobId || !jobById(plan, jobId)) return;
      setPlan((p) => addJobToStage(p, jobId, stageIndex, asNewStage));
      setSelected(jobId);
    },
    [plan]
  );

  const allowDrop = (id: string) => (e: DragEvent) => {
    e.preventDefault();
    setOverId(id);
  };

  return (
    <div className={activeTab === "gantt" || activeTab === "priority" ? "space-y-1" : "space-y-6"}>
      {activeTab !== "gantt" && activeTab !== "priority" && (
      <div className="no-print flex flex-wrap items-center justify-between gap-2 rounded-xl border border-[var(--card-border)] bg-[var(--card)] px-3 py-2">
        <div className="min-w-0">
          <p className="text-xs text-[var(--muted)]">
            {supabaseOn
              ? "Değişiklikler otomatik olarak Supabase veritabanına kaydedilir."
              : "Supabase yok · kayıt yalnızca bu tarayıcıda tutulur."}
          </p>
          <p className="text-sm text-[var(--foreground)]">{saveStatus}</p>
        </div>
        <div className="flex flex-wrap gap-2">
          <button
            type="button"
            onClick={() => void forceSave()}
            className="rounded-lg bg-[var(--accent)] px-3 py-1.5 text-sm text-white hover:bg-[var(--accent-hover)]"
          >
            Şimdi kaydet
          </button>
          <button
            type="button"
            onClick={() => void wipeEntirePlan()}
            className="rounded-lg border border-rose-200 px-3 py-1.5 text-sm text-rose-700 hover:bg-rose-50"
          >
            Tüm veriyi sil
          </button>
        </div>
      </div>
      )}

      <div className={`no-print flex flex-wrap gap-1 border-b border-[var(--card-border)] ${activeTab === "gantt" || activeTab === "priority" ? "mb-1" : ""}`}>
        {(
          [
            ["jobs", "İş listesi & kapasite", plan.jobs.length],
            ["priority", "Öncelik sırası", plan.stages.length],
            ["gantt", "Gantt çizelgesi", filteredGantt.rows.length],
            ["track", "Takip", trackRows.length],
          ] as const
        ).map(([id, label, count]) => (
          <button
            key={id}
            type="button"
            onClick={() => setActiveTab(id)}
            className={`rounded-t-lg font-medium transition-colors ${
              activeTab === "gantt" || activeTab === "priority"
                ? "px-3 py-1.5 text-xs"
                : "px-4 py-2.5 text-sm"
            } ${
              activeTab === id
                ? "border border-b-0 border-[var(--card-border)] bg-[var(--card)] text-[var(--accent)]"
                : "text-[var(--muted)] hover:text-[var(--foreground)]"
            }`}
          >
            {label}
            <span className="ml-2 tabular-nums text-xs opacity-70">{count}</span>
          </button>
        ))}
      </div>

      {activeTab === "jobs" && (
      <div className="w-full max-w-[1920px] space-y-6">
        <header>
          <h1 className="text-3xl font-bold tracking-tight">İş sıralama ve Gantt</h1>
          <p className="mt-3 max-w-3xl text-[var(--muted)]">
            İşleri öncelik sırasına dizin; her proje için haftalık personel sayısını girin. Sistem
            kapasiteye göre paralel hatları açarak Gantt’ı üretir.
          </p>
        </header>

        <div className="grid w-full gap-4 lg:grid-cols-2">
        <section className="section-card section-card--tone-0 min-w-0 overflow-hidden rounded-xl border border-[var(--card-border)] bg-[var(--card)]">
          <div className="section-card__header flex items-baseline justify-between gap-3 px-4 py-3 text-sm font-semibold uppercase tracking-wider">
            <span>Kapasite</span>
            <span className="normal-case tracking-normal text-[var(--muted)]">
              {capDeleteProject && capDeleteProject !== "all"
                ? `${sortedCapacities.length}/${plan.weeklyCapacities.length} · ${capDeleteProject}`
                : `${plan.weeklyCapacities.length} kayıt · proje · hafta · tip · kişi`}
            </span>
          </div>

          <div className="no-print space-y-3 p-4">
            <textarea
              rows={3}
              value={capPaste}
              placeholder={
                "proje\thafta\ttip\tkişi\n252.Simonsen\t2026-35\tDonatım\t3\n252.Simonsen\t2026-35\tKonstrüks\t2"
              }
              onChange={(e) => setCapPaste(e.target.value)}
              onPaste={(e) => {
                const text = clipboardToText(
                  e.clipboardData.getData("text/plain"),
                  e.clipboardData.getData("text/html")
                );
                if (!text) return;
                e.preventDefault();
                setCapPaste(text);
                importCapacity(text);
              }}
              className="w-full rounded-lg border border-dashed border-[var(--card-border)] bg-[var(--background)] p-3 text-sm"
            />
            <div className="flex flex-wrap gap-2">
              <button
                type="button"
                onClick={() => importCapacity(capPaste)}
                className="rounded-lg bg-[var(--accent)] px-3 py-2 text-sm text-white hover:bg-[var(--accent-hover)]"
              >
                Kapasiteyi aktar
              </button>
              <button
                type="button"
                onClick={() => {
                  setCapPaste("");
                  setCapStatus("");
                }}
                className="rounded-lg border border-[var(--card-border)] px-3 py-2 text-sm"
              >
                Alanı temizle
              </button>
              <button
                type="button"
                disabled={plan.weeklyCapacities.length === 0}
                className="rounded-lg border border-rose-200 px-3 py-2 text-sm text-rose-700 disabled:opacity-40"
                title="Tüm haftalık kapasite kayıtlarını sil"
                onClick={() => {
                  if (!plan.weeklyCapacities.length) return;
                  if (
                    !confirm(
                      `${plan.weeklyCapacities.length} kapasite kaydı tamamen silinsin mi?`
                    )
                  )
                    return;
                  setPlan(clearAllWeeklyCapacities(plan));
                  setCapStatus("Tüm kapasite silindi.");
                  setCapError(false);
                }}
              >
                Tüm kapasiteyi sil
              </button>
            </div>
            <div className="flex flex-wrap items-center gap-2">
              <select
                value={capDeleteProject}
                onChange={(e) => setCapDeleteProject(e.target.value)}
                disabled={capacityProjects.length === 0}
                className="h-9 min-w-[12rem] rounded-lg border border-[var(--card-border)] bg-[var(--background)] px-2 text-sm disabled:opacity-40"
                title="Listeyi bu projeye göre filtrele; silme de aynı seçimi kullanır"
              >
                {capacityProjects.length === 0 ? (
                  <option value="">Kapasite yok</option>
                ) : (
                  <>
                    <option value="all">Tüm projeler ({plan.weeklyCapacities.length})</option>
                    {capacityProjects.map((p) => (
                      <option key={p} value={p}>
                        {p} (
                        {
                          plan.weeklyCapacities.filter(
                            (c) => (c.project || DEFAULT_PROJECT) === p
                          ).length
                        }
                        )
                      </option>
                    ))}
                  </>
                )}
              </select>
              <button
                type="button"
                disabled={
                  !capDeleteProject ||
                  capDeleteProject === "all" ||
                  capacityProjects.length === 0
                }
                className="rounded-lg border border-rose-200 px-3 py-2 text-sm text-rose-700 disabled:opacity-40"
                title="Seçili projenin tüm kapasite / kaynak tanımlarını sil"
                onClick={() => {
                  const project = capDeleteProject.trim();
                  if (!project || project === "all") return;
                  const count = plan.weeklyCapacities.filter(
                    (c) => (c.project || DEFAULT_PROJECT) === project
                  ).length;
                  if (!count) return;
                  if (
                    !confirm(
                      `“${project}” projesinin ${count} kapasite / kaynak kaydı silinsin mi?`
                    )
                  )
                    return;
                  setPlan(clearWeeklyCapacitiesForProject(plan, project));
                  setCapStatus(`“${project}” kapasitesi silindi (${count} kayıt).`);
                  setCapError(false);
                }}
              >
                Proje kapasitesini sil
              </button>
              <button
                type="button"
                disabled={capSelected.length === 0}
                title="Seçili haftalarda kişi −1 (taslak; Güncelle ile uygulanır)"
                className="h-9 rounded-lg border border-[var(--card-border)] px-3 text-sm font-semibold disabled:opacity-40"
                onClick={() => bumpSelectedCapWeeks(-1)}
              >
                Haftalar −1
              </button>
              <button
                type="button"
                disabled={capSelected.length === 0}
                title="Seçili haftalarda kişi +1 (taslak; Güncelle ile uygulanır)"
                className="h-9 rounded-lg border border-amber-200 bg-amber-50 px-3 text-sm font-semibold text-amber-900 disabled:opacity-40"
                onClick={() => bumpSelectedCapWeeks(1)}
              >
                Haftalar +1
              </button>
              <button
                type="button"
                disabled={Object.keys(capDraft).length === 0}
                title="Taslak kapasiteyi Gantt’a uygula ve veritabanına yaz"
                className="h-9 rounded-lg bg-[var(--accent)] px-3 text-sm font-semibold text-white hover:bg-[var(--accent-hover)] disabled:opacity-40"
                onClick={() => void applyCapDraftAndSave()}
              >
                Güncelle
                {Object.keys(capDraft).length > 0
                  ? ` (${Object.keys(capDraft).length})`
                  : ""}
              </button>
            </div>
            {capStatus && (
              <p className={`text-xs ${capError ? "text-rose-700" : "text-emerald-700"}`}>{capStatus}</p>
            )}

            <div className="flex flex-wrap items-center gap-1.5 pt-1">
              <input
                value={capProject}
                onChange={(e) => setCapProject(e.target.value)}
                list="projects"
                placeholder="Proje adı"
                className="w-[7.5rem] rounded-lg border border-[var(--card-border)] bg-[var(--background)] px-2 py-1.5 text-sm"
              />
              <span className="text-[var(--muted)]">·</span>
              <input
                value={capWeek}
                onChange={(e) => setCapWeek(e.target.value)}
                placeholder="Hafta"
                className="w-[5.5rem] rounded-lg border border-[var(--card-border)] bg-[var(--background)] px-2 py-1.5 text-sm tabular-nums"
              />
              <span className="text-[var(--muted)]">·</span>
              <select
                value={capRole}
                onChange={(e) => setCapRole(e.target.value)}
                className="w-[9.5rem] rounded-lg border border-[var(--card-border)] bg-[var(--background)] px-2 py-1.5 text-sm"
              >
                {ROLE_OPTIONS.map((r) => (
                  <option key={r} value={r}>
                    {r}
                  </option>
                ))}
              </select>
              <span className="text-[var(--muted)]">·</span>
              <input
                type="number"
                min={0}
                step={1}
                value={capPeople}
                onChange={(e) => setCapPeople(Math.max(0, Number(e.target.value) || 0))}
                className="w-14 rounded-lg border border-[var(--card-border)] bg-[var(--background)] px-2 py-1.5 text-sm"
              />
              <span className="text-xs text-[var(--muted)]">kişi</span>
              <button
                type="button"
                onClick={addCapacityManual}
                className="rounded-lg border border-[var(--card-border)] px-3 py-1.5 text-sm hover:border-[var(--accent)]"
              >
                Ekle
              </button>
            </div>

            {sortedCapacities.length > 0 ? (
              <div className="max-h-[min(70vh,36rem)] space-y-2 overflow-auto">
                <div className="flex items-center gap-2 px-1 text-[11px] text-[var(--muted)]">
                  <label className="flex cursor-pointer items-center gap-1.5">
                    <input
                      type="checkbox"
                      className="size-3.5 accent-[var(--accent)]"
                      checked={
                        sortedCapacities.length > 0 &&
                        sortedCapacities.every((c) => capSelected.includes(capRowKey(c)))
                      }
                      onChange={(e) => {
                        if (e.target.checked) {
                          setCapSelected(sortedCapacities.map((c) => capRowKey(c)));
                        } else {
                          setCapSelected([]);
                        }
                      }}
                    />
                    Tümünü seç
                  </label>
                  <span className="tabular-nums">
                    {capSelected.length} seçili
                    {Object.keys(capDraft).length > 0
                      ? ` · ${Object.keys(capDraft).length} taslak`
                      : ""}
                  </span>
                </div>
                {sortedCapacities.map((c) => {
                  const rowKey = capRowKey(c);
                  const selected = capSelected.includes(rowKey);
                  const people = capPeopleOf(c);
                  const drafted = rowKey in capDraft;
                  return (
                  <div
                    key={rowKey}
                    className={`flex flex-wrap items-center gap-1.5 rounded-lg border px-2 py-1.5 text-sm ${
                      selected
                        ? "border-[var(--accent)] bg-[var(--accent)]/5"
                        : drafted
                          ? "border-amber-300 bg-amber-50/50"
                          : "border-[var(--card-border)] bg-[var(--background)]"
                    }`}
                  >
                    <input
                      defaultValue={c.project}
                      list="projects"
                      className="min-w-[5rem] flex-1 rounded border border-transparent bg-transparent px-1 py-0.5 font-semibold hover:border-[var(--card-border)] focus:border-[var(--accent)] focus:outline-none"
                      onBlur={(e) => {
                        const project = e.target.value.trim() || DEFAULT_PROJECT;
                        if (project === c.project) return;
                        setPlan(
                          replaceWeeklyCapacity(
                            plan,
                            { project: c.project, year: c.year, week: c.week, role: c.role },
                            { project, year: c.year, week: c.week, role: c.role, people: c.people }
                          )
                        );
                      }}
                    />
                    <input
                      type="checkbox"
                      className="size-3.5 shrink-0 accent-[var(--accent)]"
                      checked={selected}
                      title="Bu haftayı Haftalar ±1 için seç"
                      onChange={(e) => {
                        setCapSelected((prev) =>
                          e.target.checked
                            ? prev.includes(rowKey)
                              ? prev
                              : [...prev, rowKey]
                            : prev.filter((k) => k !== rowKey)
                        );
                      }}
                    />
                    <span className="text-[var(--muted)]">·</span>
                    <input
                      key={`${rowKey}-week`}
                      defaultValue={formatWeekLabel(c.year, c.week)}
                      className="w-[5.5rem] rounded border border-transparent bg-transparent px-1 py-0.5 tabular-nums hover:border-[var(--card-border)] focus:border-[var(--accent)] focus:outline-none"
                      onBlur={(e) => {
                        const m = e.target.value.trim().match(/^(\d{4})[./-](\d{1,2})$/);
                        if (!m) {
                          e.target.value = formatWeekLabel(c.year, c.week);
                          return;
                        }
                        const year = Number(m[1]);
                        let week = Number(m[2]);
                        if (m[2].length === 1 && week >= 1 && week <= 5) week *= 10;
                        if (year === c.year && week === c.week) return;
                        setPlan(
                          replaceWeeklyCapacity(
                            plan,
                            { project: c.project, year: c.year, week: c.week, role: c.role },
                            { project: c.project, year, week, role: c.role, people: c.people }
                          )
                        );
                      }}
                    />
                    <span className="text-[var(--muted)]">·</span>
                    <select
                      value={c.role}
                      className="w-[9.5rem] rounded border border-[var(--card-border)] bg-[var(--card)] px-1 py-0.5 text-sm"
                      onChange={(e) => {
                        const role = normalizeRole(e.target.value);
                        if (role === c.role) return;
                        setPlan(
                          replaceWeeklyCapacity(
                            plan,
                            { project: c.project, year: c.year, week: c.week, role: c.role },
                            { project: c.project, year: c.year, week: c.week, role, people: c.people }
                          )
                        );
                      }}
                    >
                      {c.role === "Belirtilmedi" && (
                        <option value="Belirtilmedi">Belirtilmedi</option>
                      )}
                      {!ROLE_OPTIONS.includes(c.role as (typeof ROLE_OPTIONS)[number]) &&
                        c.role !== "Belirtilmedi" && (
                          <option value={c.role}>{c.role}</option>
                        )}
                      {ROLE_OPTIONS.map((r) => (
                        <option key={r} value={r}>
                          {r}
                        </option>
                      ))}
                    </select>
                    <span className="text-[var(--muted)]">·</span>
                    <div className="flex items-center gap-0.5">
                      <button
                        type="button"
                        disabled={selected}
                        title={
                          selected
                            ? "Seçili hafta: yalnızca Haftalar ±1"
                            : "Taslak kişi −1"
                        }
                        className="h-8 w-8 rounded border border-[var(--card-border)] text-sm font-semibold disabled:opacity-40"
                        onClick={() =>
                          setCapDraft((prev) => ({
                            ...prev,
                            [rowKey]: Math.max(0, (prev[rowKey] ?? c.people) - 1),
                          }))
                        }
                      >
                        −
                      </button>
                      <input
                        type="number"
                        min={0}
                        step={1}
                        value={people}
                        disabled={selected}
                        title={
                          selected
                            ? "Seçili hafta: yalnızca Haftalar ±1 ile değiştirin"
                            : "Taslak kişi; Güncelle ile uygulanır"
                        }
                        className="h-8 w-14 rounded border border-amber-200 bg-amber-50/70 px-1 text-center text-sm tabular-nums text-amber-950 focus:border-[var(--accent)] focus:outline-none disabled:opacity-50"
                        onChange={(e) => {
                          const v = Math.max(0, Number(e.target.value) || 0);
                          setCapDraft((prev) => ({ ...prev, [rowKey]: v }));
                        }}
                      />
                      <button
                        type="button"
                        disabled={selected}
                        title={
                          selected
                            ? "Seçili hafta: yalnızca Haftalar ±1"
                            : "Taslak kişi +1"
                        }
                        className="h-8 w-8 rounded border border-amber-200 bg-amber-50 text-sm font-semibold text-amber-900 disabled:opacity-40"
                        onClick={() =>
                          setCapDraft((prev) => ({
                            ...prev,
                            [rowKey]: (prev[rowKey] ?? c.people) + 1,
                          }))
                        }
                      >
                        +
                      </button>
                    </div>
                    <span className="text-xs text-[var(--muted)]">kişi</span>
                    <button
                      type="button"
                      className="ml-auto text-xs text-rose-700"
                      onClick={() =>
                        setPlan((p) =>
                          removeWeeklyCapacity(p, c.project, c.year, c.week, c.role)
                        )
                      }
                    >
                      Sil
                    </button>
                  </div>
                  );
                })}
              </div>
            ) : plan.weeklyCapacities.length > 0 &&
              capDeleteProject &&
              capDeleteProject !== "all" ? (
              <p className="text-xs text-[var(--muted)]">
                “{capDeleteProject}” için kapasite kaydı yok. Combo’dan başka proje seçin.
              </p>
            ) : null}
            <p className="text-[11px] text-[var(--muted)]">
              Hafta kutucuğunu işaretleyin → <strong>Haftalar ±1</strong> ile kişi sayısını değiştirin →{" "}
              <strong>Güncelle</strong> ile Gantt’a uygulayın ve veritabanına yazın. Seçili satırlarda satır
              ± düğmeleri kapalıdır.
            </p>
          </div>
        </section>

        <section className="section-card section-card--tone-0 min-w-0 overflow-hidden rounded-xl border border-[var(--card-border)] bg-[var(--card)]">
          <div className="section-card__header flex items-baseline justify-between gap-3 px-4 py-3 text-sm font-semibold uppercase tracking-wider">
            <span>İş listesi</span>
            <span className="normal-case tracking-normal text-[var(--muted)]">Aktarım</span>
          </div>

          <div className="no-print space-y-3 border-b border-[var(--card-border)] p-4">
            <label className="flex flex-wrap justify-between gap-2 text-sm font-semibold">
              Excel’den yapıştır
              <span className="font-normal text-[var(--muted)]">Proje · kalem · personel tipi · saat</span>
            </label>
            <textarea
              rows={4}
              value={pasteText}
              placeholder={
                "iş kalemi\tpersonel tipi\tsaat\n252.100.101-General Arrangement\tDonatım\t25\n\nveya:\nproje\tiş kalemi\tpersonel tipi\tsaat\n252.Simonsen\tBorulama\tDonatım\t40"
              }
              onChange={(e) => setPasteText(e.target.value)}
              onPaste={(e) => {
                const text = clipboardToText(
                  e.clipboardData.getData("text/plain"),
                  e.clipboardData.getData("text/html")
                );
                if (!text) return;
                e.preventDefault();
                setPasteText(text);
                importPaste(text);
              }}
              className="w-full rounded-lg border border-dashed border-[var(--card-border)] bg-[var(--background)] p-3 text-sm"
            />
            <p className="text-[11px] text-[var(--muted)]">
              Proje adı kapasiteyle aynı olmalı (ör. <strong>252.Simonsen</strong>). Eksikse seçili / kapasite projesi kullanılır.
            </p>
            <div className="flex gap-2">
              <button
                type="button"
                onClick={() => importPaste(pasteText)}
                className="rounded-lg bg-[var(--accent)] px-3 py-2 text-sm text-white hover:bg-[var(--accent-hover)]"
              >
                Sıraya aktar
              </button>
              <button
                type="button"
                onClick={() => {
                  setPasteText("");
                  setPasteStatus("");
                }}
                className="rounded-lg border border-[var(--card-border)] px-3 py-2 text-sm"
              >
                Alanı temizle
              </button>
              <button
                type="button"
                disabled={plan.jobs.length === 0}
                className="rounded-lg border border-rose-200 px-3 py-2 text-sm text-rose-700 disabled:opacity-40"
                title="Tüm işleri, öncelik sırasını ve takip kayıtlarını sil (kapasite kalır)"
                onClick={() => {
                  if (!plan.jobs.length) return;
                  if (
                    !confirm(
                      `${plan.jobs.length} iş kalemi, öncelik sırası ve takip kayıtları silinsin mi? Kapasite kalır.`
                    )
                  )
                    return;
                  setPlan(clearAllJobs(plan));
                  setSelected(null);
                  setSelectChain([]);
                  setPasteStatus("Tüm işler silindi.");
                  setPasteError(false);
                }}
              >
                Tüm işleri sil
              </button>
            </div>
            {pasteStatus && (
              <p className={`text-xs ${pasteError ? "text-rose-700" : "text-emerald-700"}`}>{pasteStatus}</p>
            )}
            <p className="text-[11px] text-[var(--muted)]">
              Yapıştırma sırası = öncelik sırası. Yeni satırlar en alta eklenir; sırayı sadece gerektiğinde değiştirin.
            </p>
          </div>

          <details className="no-print px-4 py-2">
            <summary className="cursor-pointer text-xs text-[var(--muted)]">
              İsteğe bağlı: tek kalem ekle
            </summary>
            <form
              className="mt-2 grid grid-cols-2 gap-2 pb-2"
              onSubmit={(e) => {
                e.preventDefault();
                const form = e.currentTarget;
                const name = (form.elements.namedItem("name") as HTMLInputElement).value.trim();
                const project = (form.elements.namedItem("project") as HTMLInputElement).value.trim();
                const role = normalizeRole(
                  (form.elements.namedItem("role") as HTMLInputElement).value
                );
                const hours = Number((form.elements.namedItem("hours") as HTMLInputElement).value);
                const people = Number((form.elements.namedItem("people") as HTMLInputElement).value);
                if (!name || !project || !role || role === "Belirtilmedi" || !(hours > 0) || !(people > 0)) return;
                const id = newId();
                setPlan({
                  ...plan,
                  jobs: [...plan.jobs, { id, name, project, role, hours, people }],
                  stages: [...plan.stages, { jobIds: [id] }],
                });
                form.reset();
              }}
            >
              <input
                name="project"
                list="projects"
                required
                maxLength={80}
                placeholder="Proje adı (ör. 252.Simonsen)"
                defaultValue={projectFilter !== "all" ? projectFilter : ""}
                className="col-span-2 rounded-lg border border-[var(--card-border)] bg-[var(--background)] px-3 py-2 text-sm"
              />
              <datalist id="projects">
                {projects.map((p) => (
                  <option key={p} value={p} />
                ))}
              </datalist>
              <input
                name="name"
                required
                maxLength={80}
                placeholder="İş kalemi / aktivite"
                className="col-span-2 rounded-lg border border-[var(--card-border)] bg-[var(--background)] px-3 py-2 text-sm"
              />
              <input
                name="role"
                required
                list="roles"
                maxLength={40}
                placeholder="Personel tipi"
                className="rounded-lg border border-[var(--card-border)] bg-[var(--background)] px-3 py-2 text-sm"
              />
              <datalist id="roles">
                {ROLE_OPTIONS.map((r) => (
                  <option key={r} value={r} />
                ))}
              </datalist>
              <div className="col-span-2 flex gap-2">
                <input
                  name="hours"
                  type="number"
                  required
                  min={0.5}
                  step={0.5}
                  placeholder="Saat"
                  className="min-w-0 flex-1 rounded-lg border border-[var(--card-border)] bg-[var(--background)] px-3 py-2 text-sm"
                />
                <input
                  name="people"
                  type="number"
                  required
                  min={0.5}
                  step={0.5}
                  placeholder="Kişi/akış"
                  defaultValue={1}
                  className="w-28 rounded-lg border border-[var(--card-border)] bg-[var(--background)] px-3 py-2 text-sm"
                />
                <button
                  type="submit"
                  className="rounded-lg bg-[var(--accent)] px-3 py-2 text-sm text-white hover:bg-[var(--accent-hover)]"
                >
                  Ekle
                </button>
              </div>
            </form>
          </details>
        </section>
        </div>
      </div>
      )}

      {activeTab === "priority" && (
        <section className="flex w-full min-w-0 select-none flex-col overflow-hidden rounded-xl border border-[var(--card-border)] bg-[var(--card)] max-h-[calc(100vh-4.25rem)]">
          <datalist id="priority-roles">
            {ROLE_OPTIONS.map((r) => (
              <option key={r} value={r} />
            ))}
          </datalist>
          <datalist id="priority-projects">
            {projects.map((p) => (
              <option key={p} value={p} />
            ))}
          </datalist>
          <div className="shrink-0 border-b border-[var(--card-border)] bg-[var(--card)]">
            <div className="flex flex-wrap items-center gap-1.5 px-2 py-1">
              <div className="flex shrink-0 items-center gap-1.5">
                <h2 className="text-xs font-semibold uppercase tracking-wider text-[var(--muted)]">Öncelik</h2>
                <select
                  value={projectFilter}
                  onChange={(e) => setProjectFilter(e.target.value)}
                  className="h-7 max-w-[12rem] rounded border border-[var(--card-border)] bg-[var(--background)] px-1.5 text-[11px] text-[var(--foreground)]"
                  title="İşlem yapılan proje"
                >
                  <option value="all">Tüm projeler</option>
                  {projects.map((p) => (
                    <option key={p} value={p}>
                      {p}
                    </option>
                  ))}
                </select>
                <span className="text-[10px] tabular-nums text-[var(--muted)]">
                  {projectFilter !== "all" || prioritySearch.trim()
                    ? `${filteredPriorityStages.length}/${plan.stages.length}`
                    : plan.stages.length}{" "}
                  aşama
                </span>
              </div>
              <input
                type="search"
                value={prioritySearch}
                onChange={(e) => setPrioritySearch(e.target.value)}
                placeholder="Aktivite ara…"
                className="h-9 w-[min(100%,400px)] shrink-0 rounded border border-[var(--card-border)] bg-[var(--background)] px-2 text-[11px] text-[var(--foreground)] placeholder:text-[var(--muted)]"
                title="İsim, rol, proje veya A numarasına göre filtrele; tıkla=çoklu seçim · CTRL=öncül→ardıl · ↑/↓=bir sıra · Shift+↑/↓=üste/alta"
              />
              <div className="flex min-w-0 flex-1 flex-nowrap items-center justify-end gap-1 overflow-x-auto">
                {(
                  [
                    ["Üst", () => shiftSelectedDir("top"), !linkTargetIds.length],
                    ["↑", () => shiftSelectedDir("up"), !canShiftUp],
                    ["↓", () => shiftSelectedDir("down"), !canShiftDown],
                    ["Alt", () => shiftSelectedDir("bottom"), !linkTargetIds.length],
                  ] as [string, () => void, boolean][]
                ).map(([label, fn, disabled]) => (
                  <button
                    key={label}
                    type="button"
                    disabled={disabled}
                    onClick={fn}
                    title={
                      label === "Üst"
                        ? "Shift+↑ · en üste"
                        : label === "Alt"
                          ? "Shift+↓ · en alta"
                          : label === "↑"
                            ? "↑ · bir sıra yukarı"
                            : label === "↓"
                              ? "↓ · bir sıra aşağı"
                              : label
                    }
                    className="h-8 w-[4.5rem] shrink-0 rounded border border-[var(--card-border)] text-xs disabled:opacity-40"
                  >
                    {label}
                  </button>
                ))}
                {([-20, -10, -1, 1, 10, 20] as const).map((delta) => {
                  const disabled =
                    !linkTargetIds.length ||
                    (delta < 0 && !canShiftUp) ||
                    (delta > 0 && !canShiftDown);
                  return (
                    <button
                      key={`shift-bar-${delta}`}
                      type="button"
                      disabled={disabled}
                      title={
                        delta > 0
                          ? `Seçili iş(ler)i ${delta} sıra aşağı`
                          : `Seçili iş(ler)i ${Math.abs(delta)} sıra yukarı`
                      }
                      onClick={() => shiftSelectedBy(delta)}
                      className={`h-8 w-[4.5rem] shrink-0 rounded text-xs font-semibold tabular-nums disabled:opacity-40 ${
                        delta > 0
                          ? "bg-amber-100 text-amber-900"
                          : "bg-sky-100 text-sky-900"
                      }`}
                    >
                      {delta > 0 ? `+${delta}` : delta}
                    </button>
                  );
                })}
              </div>
              <div className="flex shrink-0 flex-wrap items-center gap-1">
                <label
                  className="flex items-center gap-1 rounded border border-[var(--card-border)] bg-[var(--background)] px-1.5 py-0.5 text-[10px] text-[var(--muted)]"
                  title="Seçili aşamadan sonra N iş günü boşluk; sonraki aktiviteler ötelenir"
                >
                  Boşluk
                  <input
                    type="number"
                    min={0}
                    step={1}
                    disabled={!place}
                    value={place ? plan.stages[place.stageIndex]?.gapAfterDays ?? 0 : 0}
                    onChange={(e) => {
                      if (!place) return;
                      setPlan(setStageGapAfterDays(plan, place.stageIndex, Number(e.target.value)));
                    }}
                    className="h-6 w-12 rounded border border-[var(--card-border)] bg-[var(--card)] px-1 text-[11px] tabular-nums text-[var(--foreground)] disabled:opacity-40"
                  />
                  <span className="whitespace-nowrap">iş günü</span>
                </label>
                <button
                  type="button"
                  disabled={!canClearSelectedLinks}
                  className="no-print rounded border border-[var(--card-border)] px-1.5 py-0.5 text-[10px] disabled:opacity-40"
                  title="Seçili aktivitelerin öncül–ardıl bağlarını sil"
                  onClick={() => {
                    setPlan(clearDependenciesForJobs(plan, linkTargetIds));
                    setSelectChain(selected ? [selected] : []);
                  }}
                >
                  Bağları sil
                </button>
                <button
                  type="button"
                  disabled={plan.stages.length === 0}
                  className="no-print rounded border border-rose-200 px-1.5 py-0.5 text-[10px] text-rose-700 disabled:opacity-40"
                  title="Tüm öncelik sırasını sil"
                  onClick={() => {
                    if (!plan.stages.length) return;
                    if (!confirm("Öncelik sırası tamamen silinsin mi? İş kalemleri listede kalır.")) return;
                    setPlan(clearPriorityOrder(plan));
                    setSelected(null);
                    setSelectChain([]);
                  }}
                >
                  Sırayı sil
                </button>
                <button
                  type="button"
                  onClick={() => void forceSave()}
                  className="rounded border border-[var(--card-border)] px-1.5 py-0.5 text-[10px]"
                >
                  Kaydet
                </button>
                {selectChain.length > 0 && (
                  <button
                    type="button"
                    className="rounded border border-[var(--card-border)] px-1.5 py-0.5 text-[10px]"
                    onClick={() => setSelectChain(selected ? [selected] : [])}
                  >
                    Seçimi sıfırla
                  </button>
                )}
              </div>
            </div>
            <div className="mx-2 mb-1 flex gap-1">
              <div
                className={`flex-1 rounded border border-dashed px-2 py-1 text-center text-[10px] font-medium ${
                  overId === "start-top"
                    ? "drop-over border-[var(--accent)]"
                    : "border-sky-300 bg-sky-50 text-sky-800"
                }`}
                onDragOver={allowDrop("start-top")}
                onDragLeave={() => setOverId(null)}
                onDrop={dropJob(0, true)}
              >
                Başa bırak
              </div>
              <div
                className={`flex-1 rounded border border-dashed px-2 py-1 text-center text-[10px] font-medium ${
                  overId === "end-top"
                    ? "drop-over border-[var(--accent)]"
                    : "border-amber-300 bg-amber-50 text-amber-900"
                }`}
                onDragOver={allowDrop("end-top")}
                onDragLeave={() => setOverId(null)}
                onDrop={dropJob(plan.stages.length, true)}
              >
                Sona bırak
              </div>
            </div>
          </div>

          <div className="min-h-0 flex-1 overflow-auto px-1.5 py-1">
            {plan.stages.length === 0 ? (
              <p className="py-4 text-center text-xs text-[var(--muted)]">
                Soldan bir işi buraya sürükleyin.
              </p>
            ) : filteredPriorityStages.length === 0 ? (
              <p className="py-4 text-center text-xs text-[var(--muted)]">
                {projectFilter !== "all"
                  ? `“${projectFilter}” için görünen aşama yok.`
                  : "Aramaya uyan aktivite yok."}
              </p>
            ) : (
              <div className="w-full border-t border-[var(--card-border)] bg-white text-xs text-slate-900">
                {filteredPriorityStages.map(({ si, jobIds }) => {
                  const stage = plan.stages[si];
                  return (
                    <div key={`stage-${si}`}>
                      {(stage.gapAfterDays ?? 0) > 0 && (
                        <div className="border-b border-[var(--card-border)] bg-amber-50 px-2 py-0.5 text-[9px] font-medium text-amber-900">
                          A{si + 1} sonrası +{stage.gapAfterDays} iş günü boşluk
                          {stageGanttWeekRangeLabel(plan.startDate, stage, gantt.rows)
                            ? ` · ${stageGanttWeekRangeLabel(plan.startDate, stage, gantt.rows)}`
                            : ""}
                        </div>
                      )}
                      {jobIds.map((id) => {
                        const job = jobById(plan, id);
                        if (!job) return null;
                        const preds = predecessorsOf(plan, id);
                        const succs = successorsOf(plan, id);
                        const chainIdx = selectChain.indexOf(id);
                        const inChain = chainIdx >= 0;
                        const isSelected = selected === id || inChain;
                        const tone = categoryTone(activityCategory(job.name), categoryToneMap);
                        const isEditing = editingJobId === id;
                        return (
                          <div
                            key={id}
                            id={`priority-job-${id}`}
                            draggable={!isEditing}
                            onClick={(e) => onPriorityClick(id, e)}
                            onDragStart={isEditing ? undefined : onDragStart(id)}
                            onDragOver={allowDrop(`st-${si}`)}
                            onDragLeave={() => setOverId(null)}
                            onDrop={dropJob(si, true)}
                            style={{
                              background: isSelected
                                ? "color-mix(in srgb, var(--accent) 14%, white)"
                                : tone.bg || "#fff",
                            }}
                            className={`flex min-w-0 items-stretch border-b border-[var(--card-border)] ${
                              isEditing
                                ? "cursor-default ring-1 ring-inset ring-[var(--accent)]/40"
                                : isSelected
                                  ? "cursor-grab select-none ring-1 ring-inset ring-[var(--accent)]"
                                  : preds.length || succs.length
                                    ? "cursor-grab select-none"
                                    : "cursor-grab select-none hover:brightness-[0.99]"
                            } ${overId === `st-${si}` ? "drop-over" : ""}`}
                            title={`${job.name} · tıkla: seç · CTRL: öncül→ardıl · ↑/↓: kaydır · Shift+↑/↓: üste/alta · Shift+tık: düzenle`}
                          >
                            {isEditing ? (
                              <div
                                className="w-full space-y-1 px-2 py-1.5"
                                onClick={(ev) => ev.stopPropagation()}
                                onKeyDown={(ev) => {
                                  if (ev.key === "Escape") setEditingJobId(null);
                                }}
                              >
                                <input
                                  autoFocus
                                  value={job.name}
                                  onChange={(e) => patchJob(id, { name: e.target.value })}
                                  className="w-full rounded border border-[var(--card-border)] bg-[var(--background)] px-1.5 py-0.5 text-[11px] font-medium"
                                  placeholder="İş kalemi"
                                />
                                <div className="grid grid-cols-2 gap-1 sm:grid-cols-4">
                                  <label className="flex flex-col gap-0.5 text-[8px] text-[var(--muted)]">
                                    Saat
                                    <input
                                      type="number"
                                      min={0.5}
                                      step={0.5}
                                      value={job.hours}
                                      onChange={(e) =>
                                        patchJob(id, { hours: Number(e.target.value) })
                                      }
                                      className="rounded border border-[var(--card-border)] bg-[var(--background)] px-1.5 py-0.5 text-[10px] tabular-nums"
                                    />
                                  </label>
                                  <label className="flex flex-col gap-0.5 text-[8px] text-[var(--muted)]">
                                    Kişi
                                    <input
                                      type="number"
                                      min={1}
                                      step={1}
                                      value={job.people}
                                      onChange={(e) =>
                                        patchJob(id, { people: Number(e.target.value) })
                                      }
                                      className="rounded border border-[var(--card-border)] bg-[var(--background)] px-1.5 py-0.5 text-[10px] tabular-nums"
                                    />
                                  </label>
                                  <label className="flex flex-col gap-0.5 text-[8px] text-[var(--muted)]">
                                    Personel
                                    <input
                                      list="priority-roles"
                                      value={job.role}
                                      onChange={(e) => patchJob(id, { role: e.target.value })}
                                      className="rounded border border-[var(--card-border)] bg-[var(--background)] px-1.5 py-0.5 text-[10px]"
                                    />
                                  </label>
                                  <label className="flex flex-col gap-0.5 text-[8px] text-[var(--muted)]">
                                    Proje
                                    <input
                                      list="priority-projects"
                                      value={job.project || ""}
                                      onChange={(e) => patchJob(id, { project: e.target.value })}
                                      className="rounded border border-[var(--card-border)] bg-[var(--background)] px-1.5 py-0.5 text-[10px]"
                                    />
                                  </label>
                                </div>
                                <div className="flex justify-end">
                                  <button
                                    type="button"
                                    className="rounded bg-[var(--accent)] px-2 py-0.5 text-[10px] text-white"
                                    onClick={() => setEditingJobId(null)}
                                  >
                                    Tamam
                                  </button>
                                </div>
                              </div>
                            ) : (
                              <>
                                <div className="flex min-w-0 flex-1 items-center gap-1 truncate whitespace-nowrap px-2 py-0.5 text-[11px]">
                                  {inChain && selectChain.length > 1 ? (
                                    <span className="mr-0.5 inline-flex h-3.5 min-w-3.5 shrink-0 items-center justify-center rounded-full bg-[var(--accent)] px-0.5 text-[8px] font-bold text-white">
                                      {chainIdx + 1}
                                    </span>
                                  ) : null}
                                  <span className="min-w-0 truncate">
                                    <span className="mr-1.5 tabular-nums text-[var(--muted)]">
                                      A{si + 1}
                                    </span>
                                    {job.name}
                                    {(preds.length > 0 || succs.length > 0) && (
                                      <span className="ml-1 text-[9px] text-amber-700">
                                        {preds.length ? `←${preds.length}` : ""}
                                        {succs.length ? ` →${succs.length}` : ""}
                                      </span>
                                    )}
                                  </span>
                                </div>
                                <div className="flex shrink-0 items-center gap-1 px-2 py-0.5">
                                  <span className="text-[10px] tabular-nums text-slate-900">
                                    {formatHours(job.hours)}
                                  </span>
                                  <button
                                    type="button"
                                    title="İşi kalıcı sil"
                                    className="no-print px-0.5 text-[10px] leading-none text-[var(--muted)] hover:text-rose-700"
                                    onClick={(ev) => {
                                      ev.stopPropagation();
                                      if (!confirm(`“${job.name}” silinsin mi?`)) return;
                                      setPlan(deleteJob(plan, id));
                                      setSelectChain((c) => c.filter((x) => x !== id));
                                      if (selected === id) setSelected(null);
                                      if (editingJobId === id) setEditingJobId(null);
                                    }}
                                  >
                                    ×
                                  </button>
                                </div>
                              </>
                            )}
                          </div>
                        );
                      })}
                    </div>
                  );
                })}
              </div>
            )}
            {plan.stages.length > 0 && (
              <div
                className={`mt-0.5 h-1 rounded ${overId === "ins-end" ? "drop-over h-3" : ""}`}
                onDragOver={allowDrop("ins-end")}
                onDragLeave={() => setOverId(null)}
                onDrop={dropJob(plan.stages.length, true)}
              />
            )}
          </div>

          <div className="mx-2 mb-1 mt-0.5 shrink-0 flex gap-1 border-t border-[var(--card-border)] pt-1">
            <div
              className={`flex-1 rounded border border-dashed px-2 py-1 text-center text-[10px] font-medium ${
                overId === "start-bottom"
                  ? "drop-over border-[var(--accent)]"
                  : "border-sky-300 bg-sky-50 text-sky-800"
              }`}
              onDragOver={allowDrop("start-bottom")}
              onDragLeave={() => setOverId(null)}
              onDrop={dropJob(0, true)}
            >
              Başa bırak
            </div>
            <div
              className={`flex-1 rounded border border-dashed px-2 py-1 text-center text-[10px] font-medium ${
                overId === "end-bottom"
                  ? "drop-over border-[var(--accent)]"
                  : "border-amber-300 bg-amber-50 text-amber-900"
              }`}
              onDragOver={allowDrop("end-bottom")}
              onDragLeave={() => setOverId(null)}
              onDrop={dropJob(plan.stages.length, true)}
            >
              Sona bırak
            </div>
          </div>
        </section>
      )}

      {activeTab === "gantt" && (
      <section className="rounded-xl border border-[var(--card-border)] bg-[var(--card)] p-1.5">
        <div className="no-print mb-1 flex flex-wrap items-center justify-between gap-2">
          <div className="flex min-w-0 flex-wrap items-center gap-2">
            <h2 className="text-xs font-semibold uppercase tracking-wider text-[var(--muted)]">Gantt</h2>
            <select
              value={projectFilter}
              onChange={(e) => setProjectFilter(e.target.value)}
              className="h-7 max-w-[14rem] rounded border border-[var(--card-border)] bg-[var(--background)] px-1.5 text-[11px] text-[var(--foreground)]"
              title="İşlem yapılan proje"
            >
              <option value="all">Tüm projeler</option>
              {projects.map((p) => (
                <option key={p} value={p}>
                  {p}
                </option>
              ))}
            </select>
            {filteredGantt.rows.length > 0 && (
              <p className="truncate text-[11px] text-[var(--muted)]">
                {filteredGantt.rows.length} satır · {formatHours(gantt.totalHours)} ·{" "}
                {(gantt.totalDays / WORK_DAYS_PER_WEEK).toFixed(1)} hf
              </p>
            )}
            <span className="hidden text-[10px] text-[var(--muted)] lg:inline">
              Tıkla = seç · sürükle = grubu taşı
            </span>
            <label
              className="flex items-center gap-1 rounded border border-[var(--card-border)] bg-[var(--background)] px-1.5 py-0.5 text-[10px] text-[var(--muted)]"
              title="Seçili aktivitenin aşamasından sonra N iş günü boşluk; sonraki aktiviteler ötelenir"
            >
              Boşluk
              <input
                type="number"
                min={0}
                step={1}
                disabled={!ganttGapPlace}
                value={
                  ganttGapPlace
                    ? plan.stages[ganttGapPlace.stageIndex]?.gapAfterDays ?? 0
                    : 0
                }
                onChange={(e) => {
                  if (!ganttGapPlace) return;
                  setPlan(
                    setStageGapAfterDays(plan, ganttGapPlace.stageIndex, Number(e.target.value))
                  );
                }}
                className="h-6 w-12 rounded border border-[var(--card-border)] bg-[var(--card)] px-1 text-[11px] tabular-nums text-[var(--foreground)] disabled:opacity-40"
              />
              <span className="whitespace-nowrap">iş günü</span>
            </label>
            {ganttSelectIds.length > 0 && (
              <button
                type="button"
                className="rounded border border-[var(--card-border)] px-1.5 py-0.5 text-[10px]"
                onClick={() => setGanttSelectIds([])}
              >
                Seçimi temizle ({ganttSelectIds.length})
              </button>
            )}
            {saveStatus && (
              <span className="hidden text-[10px] text-[var(--muted)] sm:inline">{saveStatus}</span>
            )}
          </div>
          <div className="flex flex-wrap items-center gap-2">
            <button
              type="button"
              onClick={() => void forceSave()}
              className="rounded-md border border-[var(--card-border)] px-2 py-1 text-[11px]"
            >
              Kaydet
            </button>
            <button
              type="button"
              onClick={() => void exportGanttPdf()}
              disabled={filteredGantt.rows.length === 0}
              className="rounded-md bg-[var(--accent)] px-2.5 py-1 text-[11px] text-white hover:bg-[var(--accent-hover)] disabled:opacity-40"
            >
              PDF
            </button>
          </div>
        </div>
        {gantt.warnings.length > 0 && (
          <div className="mb-2 max-h-16 overflow-auto rounded-lg border border-rose-200 bg-rose-50 px-2 py-1 text-[10px] text-rose-800">
            {gantt.warnings.map((w) => (
              <div key={w}>{w}</div>
            ))}
          </div>
        )}
        {filteredGantt.rows.length === 0 ? (
          <p className="py-8 text-sm text-[var(--muted)]">
            {projectFilter !== "all"
              ? `“${projectFilter}” için Gantt satırı yok. Başka proje seçin veya sıralamayı kontrol edin.`
              : "Sıralama oluştukça çizelge burada üretilir. Zaman ekseni proje başlangıç haftasından başlar."}
          </p>
        ) : (
          <div className="max-h-[calc(100vh-5.5rem)] overflow-auto">
            <div ref={ganttExportRef} className="inline-block min-w-full bg-white text-slate-900">
              {pdfCompact && (
                <div className="mb-3 flex items-center gap-3 border-b border-slate-200 pb-3">
                  {/* next/image PDF yakalamada sorun çıkarabilir — düz img */}
                  <img
                    src={typeof ARTI_LOGO === "string" ? ARTI_LOGO : ARTI_LOGO.src}
                    alt="ARTI Engineering"
                    width={artiLogoDisplayWidth(40)}
                    height={40}
                    className="h-10 w-auto object-contain"
                  />
                  <span className="text-xl font-semibold tracking-tight text-[#0d4f8b]">Gantt</span>
                </div>
              )}
              {pdfCompact && (
                <>
                  <div className="mb-2 text-sm font-semibold text-slate-800">
                    İş Planı
                    {projectFilter !== "all" ? ` · ${projectFilter}` : ""}
                  </div>
                  <p className="mb-3 text-[11px] text-slate-600">
                    {filteredGantt.rows.length} satır · {formatHours(gantt.totalHours)} ·{" "}
                    {(gantt.totalDays / WORK_DAYS_PER_WEEK).toFixed(1)} hafta ·{" "}
                    {weekRangeLabel(buildWeekTicks(plan.startDate, gantt.totalDays))}
                  </p>
                </>
              )}
              <GanttView
                plan={plan}
                capacityProject={projectFilter !== "all" ? projectFilter : null}
                planStart={plan.startDate}
                model={filteredGantt}
                groupByProject={groupByProject}
                compactLabels={pdfCompact}
                projectHeading={
                  projectFilter !== "all"
                    ? projectFilter
                    : [...new Set(filteredGantt.rows.map((r) => r.job.project || DEFAULT_PROJECT))].sort((a, b) =>
                        a.localeCompare(b, "tr")
                      ).join(" · ") || "İş planı"
                }
                selectedIds={ganttSelectIds}
                onToggleSelect={(jobId) => {
                  setGanttSelectIds((prev) =>
                    prev.includes(jobId) ? prev.filter((x) => x !== jobId) : [...prev, jobId]
                  );
                }}
                onMoveGroupBefore={(jobIds, targetJobId) => {
                  setPlan((p) => moveJobsGroupBefore(p, jobIds, targetJobId));
                }}
              />
            </div>
          </div>
        )}
      </section>
      )}

      {activeTab === "track" && (
        <section className="rounded-xl border border-[var(--card-border)] bg-[var(--card)] p-4">
          <div className="mb-4 flex flex-wrap items-end justify-between gap-3">
            <div>
              <h2 className="text-sm font-semibold uppercase tracking-wider">Haftalık takip</h2>
              <p className="mt-1 text-xs text-[var(--muted)]">
                İlerleme: %100 → %75 → %50 → %25 → %0. Süre ve kişi burada; haftalık kapasite İş listesi & kapasite ekranından.
              </p>
            </div>
            <div className="flex flex-wrap items-end gap-3">
              <div className="flex flex-col gap-1">
                <span className="text-xs text-[var(--muted)]">Takip haftası</span>
                <div className="flex h-9 items-center gap-1">
                  <button
                    type="button"
                    title="Önceki hafta"
                    onClick={() => shiftTrackWeek(-1)}
                    className="flex h-9 w-9 shrink-0 items-center justify-center rounded-lg border border-[var(--card-border)] text-sm hover:bg-[var(--background)]"
                  >
                    ←
                  </button>
                  <input
                    type="text"
                    value={trackWeek}
                    onChange={(e) => setTrackWeek(e.target.value)}
                    placeholder="2026-35"
                    className="h-9 w-28 rounded-lg border border-[var(--card-border)] bg-[var(--background)] px-2 text-center text-sm tabular-nums"
                  />
                  <button
                    type="button"
                    title="Sonraki hafta"
                    onClick={() => shiftTrackWeek(1)}
                    className="flex h-9 w-9 shrink-0 items-center justify-center rounded-lg border border-[var(--card-border)] text-sm hover:bg-[var(--background)]"
                  >
                    →
                  </button>
                </div>
              </div>
              <div className="flex flex-col gap-1">
                <span className="text-xs text-[var(--muted)]">Proje</span>
                <select
                  value={projectFilter}
                  onChange={(e) => setProjectFilter(e.target.value)}
                  className="h-9 min-w-[160px] rounded-lg border border-[var(--card-border)] bg-[var(--background)] px-3 text-sm"
                >
                  <option value="all">Tüm projeler</option>
                  {projects.map((p) => (
                    <option key={p} value={p}>
                      {p}
                    </option>
                  ))}
                </select>
              </div>
              <button
                type="button"
                disabled={!trackWeekParts || !(plan.jobProgress || []).some(
                  (p) => p.year === trackWeekParts.year && p.week === trackWeekParts.week
                )}
                className="h-9 rounded-lg border border-rose-200 px-3 text-sm text-rose-700 disabled:opacity-40"
                onClick={() => {
                  if (!trackWeekParts) return;
                  if (!confirm(`${trackWeek} haftasının takip kayıtları silinsin mi?`)) return;
                  setPlan(clearJobProgressForWeek(plan, trackWeekParts.year, trackWeekParts.week));
                }}
              >
                Bu haftayı sil
              </button>
              <button
                type="button"
                disabled={!(plan.jobProgress || []).length}
                className="h-9 rounded-lg border border-rose-200 px-3 text-sm text-rose-700 disabled:opacity-40"
                onClick={() => {
                  if (!(plan.jobProgress || []).length) return;
                  if (!confirm("Tüm haftaların takip kayıtları silinsin mi?")) return;
                  setPlan(clearAllJobProgress(plan));
                }}
              >
                Tüm takibi sil
              </button>
            </div>
          </div>

          {!trackWeekParts ? (
            <p className="py-6 text-sm text-rose-700">Hafta biçimi geçersiz. Örnek: 2026-35</p>
          ) : trackRows.length === 0 ? (
            <p className="py-8 text-sm text-[var(--muted)]">
              Bu haftada Gantt’a göre planlanmış aktivite yok. Haftayı ← → ile değiştirin veya öncelik / kapasiteyi kontrol edin.
            </p>
          ) : (
            <div className="overflow-auto">
              <table className="w-full min-w-[720px] border-collapse text-sm">
                <thead>
                  <tr className="border-b border-[var(--card-border)] text-left text-[11px] uppercase tracking-wider text-[var(--muted)]">
                    <th className="px-2 py-2 font-medium">Sıra</th>
                    <th className="px-2 py-2 font-medium">Aktivite</th>
                    <th className="px-2 py-2 font-medium">Proje</th>
                    <th className="px-2 py-2 font-medium">Rol</th>
                    <th className="px-2 py-2 font-medium">Kişi</th>
                    <th className="px-2 py-2 font-medium">Saat / süre</th>
                    <th className="px-2 py-2 font-medium">İlerleme</th>
                    <th className="px-2 py-2 font-medium">Takılma / neden</th>
                  </tr>
                </thead>
                <tbody>
                  {trackRows.map(({ jobId, stageIndex, seq }) => {
                    const job = jobById(plan, jobId);
                    if (!job) return null;
                    const entry = (plan.jobProgress || []).find(
                      (p) =>
                        p.jobId === jobId &&
                        p.year === trackWeekParts.year &&
                        p.week === trackWeekParts.week
                    );
                    const hpd = plan.hoursPerDay || 8;
                    return (
                      <tr
                        key={`${jobId}-${trackWeekParts.year}-${trackWeekParts.week}`}
                        className="border-b border-[var(--card-border)]/70 align-top"
                      >
                        <td className="px-2 py-2 tabular-nums text-[var(--muted)]">
                          A{stageIndex + 1}
                          <span className="ml-1 text-[10px]">#{seq}</span>
                        </td>
                        <td className="px-2 py-2 font-medium leading-snug">{job.name}</td>
                        <td className="px-2 py-2 text-[var(--muted)]">{job.project}</td>
                        <td className="px-2 py-2 text-[var(--muted)]">{job.role}</td>
                        <td className="px-2 py-2">
                          <div className="flex items-center gap-0.5">
                            <button
                              type="button"
                              title="Kişi −1"
                              onClick={() => patchJob(jobId, { people: job.people - 1 })}
                              className="rounded border border-[var(--card-border)] px-1.5 py-1 text-[10px] font-medium"
                            >
                              −
                            </button>
                            <input
                              type="number"
                              min={1}
                              step={1}
                              value={job.people}
                              title="İşe atanan kişi (tüm süre boyunca)"
                              onChange={(e) => patchJob(jobId, { people: Number(e.target.value) })}
                              className="w-12 rounded-lg border border-[var(--card-border)] bg-[var(--background)] px-1.5 py-1 text-center text-xs tabular-nums"
                            />
                            <button
                              type="button"
                              title="Kişi +1"
                              onClick={() => patchJob(jobId, { people: job.people + 1 })}
                              className="rounded border border-[var(--card-border)] px-1.5 py-1 text-[10px] font-medium"
                            >
                              +
                            </button>
                          </div>
                        </td>
                        <td className="px-2 py-2">
                          <div className="flex flex-wrap items-center gap-1">
                            <input
                              type="number"
                              min={0.5}
                              step={0.5}
                              value={job.hours}
                              title="Aktivite süresi (saat). Uzatınca Gantt yeniden planlanır."
                              onChange={(e) => patchJob(jobId, { hours: Number(e.target.value) })}
                              className="w-16 rounded-lg border border-[var(--card-border)] bg-[var(--background)] px-2 py-1 text-xs tabular-nums"
                            />
                            <button
                              type="button"
                              title={`−1 iş günü (−${hpd} saat)`}
                              onClick={() => patchJob(jobId, { hours: job.hours - hpd })}
                              className="rounded border border-[var(--card-border)] px-1.5 py-1 text-[10px] font-medium"
                            >
                              −1g
                            </button>
                            <button
                              type="button"
                              title={`−5 iş günü (−${hpd * 5} saat)`}
                              onClick={() => patchJob(jobId, { hours: job.hours - hpd * 5 })}
                              className="rounded border border-sky-200 bg-sky-50 px-1.5 py-1 text-[10px] font-medium text-sky-900"
                            >
                              −5g
                            </button>
                            <button
                              type="button"
                              title={`+1 iş günü (+${hpd} saat)`}
                              onClick={() => patchJob(jobId, { hours: job.hours + hpd })}
                              className="rounded border border-[var(--card-border)] px-1.5 py-1 text-[10px] font-medium"
                            >
                              +1g
                            </button>
                            <button
                              type="button"
                              title={`+5 iş günü (+${hpd * 5} saat)`}
                              onClick={() => patchJob(jobId, { hours: job.hours + hpd * 5 })}
                              className="rounded border border-amber-200 bg-amber-50 px-1.5 py-1 text-[10px] font-medium text-amber-900"
                            >
                              +5g
                            </button>
                          </div>
                        </td>
                        <td className="px-2 py-2">
                          <button
                            type="button"
                            title="Tıklayarak ilerlemeyi değiştir"
                            onClick={() => cycleTrackProgress(jobId)}
                            className={`min-w-[4.5rem] rounded-md px-2.5 py-1.5 text-xs font-semibold tabular-nums ${progressButtonClass(entry?.percent)}`}
                          >
                            {entry?.percent == null ? "—" : `%${entry.percent}`}
                          </button>
                        </td>
                        <td className="px-2 py-2">
                          <input
                            type="text"
                            value={entry?.reason ?? ""}
                            placeholder="Takılma nedeni…"
                            onChange={(e) => setTrackReason(jobId, e.target.value)}
                            className="w-full min-w-[12rem] rounded-lg border border-[var(--card-border)] bg-[var(--background)] px-2.5 py-1.5 text-xs"
                          />
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )}
        </section>
      )}
    </div>
  );
}

function GanttView({
  plan,
  capacityProject = null,
  planStart,
  model,
  groupByProject,
  compactLabels = false,
  projectHeading = "",
  selectedIds = [],
  onToggleSelect,
  onMoveGroupBefore,
}: {
  plan: Plan;
  /** Proje seçiliyse hafta başlığında toplam kapasite (kişi) gösterilir */
  capacityProject?: string | null;
  planStart: string;
  model: ReturnType<typeof computeGantt>;
  groupByProject: boolean;
  /** PDF: Saat kolonunu da gizle */
  compactLabels?: boolean;
  /** Proje adı başlıkta (satırda tekrarlanmaz) */
  projectHeading?: string;
  selectedIds?: string[];
  onToggleSelect?: (jobId: string) => void;
  onMoveGroupBefore?: (jobIds: string[], targetJobId: string) => void;
}) {
  const ticks = useMemo(
    () => buildWeekTicks(planStart, model.totalDays),
    [planStart, model.totalDays]
  );
  const weekPx = Math.max(28, Math.min(48, Math.floor(900 / Math.max(ticks.length, 1))));
  const nameColPx = 480;
  const hourColPx = compactLabels ? 0 : 52;
  const labelsW = nameColPx + hourColPx;
  const labelCols = `${labelsW}px`;
  const selectedSet = useMemo(() => new Set(selectedIds), [selectedIds]);
  const [dropTargetId, setDropTargetId] = useState<string | null>(null);
  const dragIdsRef = useRef<string[]>([]);
  const suppressClickRef = useRef(false);

  const sections = useMemo(() => {
    if (!groupByProject) return [{ project: "", rows: model.rows }];
    const map = new Map<string, typeof model.rows>();
    for (const row of model.rows) {
      const p = row.job.project || DEFAULT_PROJECT;
      if (!map.has(p)) map.set(p, []);
      map.get(p)!.push(row);
    }
    return [...map.entries()]
      .sort((a, b) => a[0].localeCompare(b[0], "tr"))
      .map(([project, rows]) => ({ project, rows }));
  }, [model.rows, groupByProject]);

  const stickyLabelClass =
    "sticky left-0 z-10 flex shrink-0 items-center border-r border-[var(--card-border)] shadow-[3px_0_6px_-3px_rgba(15,23,42,0.18)]";

  const interactive = Boolean(onToggleSelect && onMoveGroupBefore && !compactLabels);

  return (
    <div className="inline-block min-w-[780px] bg-white text-xs text-slate-900">
        <div className="sticky top-0 z-20 border-b border-[var(--card-border)] bg-white shadow-[0_1px_0_0_var(--card-border)]">
        {compactLabels && projectHeading ? (
          <div className="border-b border-[var(--card-border)] bg-white px-2 py-1 text-[11px] font-semibold text-[var(--accent)]">
            {projectHeading}
          </div>
        ) : null}
        <div
          className="grid items-center bg-white font-semibold"
          style={{ gridTemplateColumns: `${labelCols} 1fr` }}
        >
          <div
            className={`${stickyLabelClass} z-30 bg-white`}
            style={{ width: labelsW, minWidth: labelsW }}
          >
            <div className="px-2 py-1" style={{ width: nameColPx, minWidth: nameColPx }}>
              İş kalemi
            </div>
            {!compactLabels && (
              <div className="px-2 py-1" style={{ width: hourColPx, minWidth: hourColPx }}>
                Saat
              </div>
            )}
          </div>
          <div className="flex">
            {ticks.map((t, i) => {
              const yearStart = i === 0 || ticks[i - 1]!.year !== t.year;
              const cap =
                capacityProject && !compactLabels
                  ? getWeeklyCapacityPeople(plan, capacityProject, t.year, t.week)
                  : null;
              return (
                <div
                  key={t.key}
                  className="border-l border-[var(--card-border)] bg-white px-0.5 py-0.5 text-center tabular-nums"
                  style={{ width: weekPx, minWidth: weekPx }}
                  title={
                    cap != null
                      ? `${t.year} hafta ${formatWeekOnly(t.week)} · kapasite ${cap} kişi`
                      : `${t.year} hafta ${formatWeekOnly(t.week)}`
                  }
                >
                  {yearStart ? (
                    <div className="text-[9px] leading-none text-[var(--accent)]">{t.year}</div>
                  ) : (
                    <div className="h-[9px]" aria-hidden />
                  )}
                  <div className="text-[10px] leading-tight">{formatWeekOnly(t.week)}</div>
                  {cap != null && (
                    <div
                      className={`text-[9px] font-semibold leading-none ${
                        cap > 4 ? "text-amber-700" : "text-[var(--muted)]"
                      }`}
                    >
                      {cap}k
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        </div>
        </div>
        {sections.map((section) => (
          <div key={section.project || "all"}>
            {groupByProject && section.project && (
              <div className="sticky left-0 z-[5] border-b border-[var(--card-border)] bg-[var(--background)] px-2 py-1 text-[10px] font-semibold uppercase tracking-wider text-[var(--accent)]">
                {section.project}
                <span className="ml-2 font-normal normal-case tracking-normal text-[var(--muted)]">
                  {section.rows.length} aktivite
                </span>
              </div>
            )}
            {section.rows.map((r) => {
              const left = weekIndexFromPlanStart(planStart, r.startDay) * weekPx;
              const width = Math.max(12, (r.durationDays / WORK_DAYS_PER_WEEK) * weekPx);
              const isSelected = selectedSet.has(r.job.id);
              const isDropTarget = dropTargetId === r.job.id;
              return (
                <div
                  key={r.job.id}
                  className={`grid items-center border-b border-[var(--card-border)] ${
                    isDropTarget ? "bg-sky-50/80" : ""
                  }`}
                  style={{ gridTemplateColumns: `${labelCols} 1fr` }}
                >
                  <div
                    className={`${stickyLabelClass} ${interactive ? "cursor-pointer" : ""} ${
                      isSelected ? "ring-1 ring-inset ring-[var(--accent)]" : ""
                    }`}
                    style={{
                      width: labelsW,
                      minWidth: labelsW,
                      background: isSelected
                        ? "color-mix(in srgb, var(--accent) 14%, white)"
                        : r.tint || "#fff",
                    }}
                    draggable={interactive}
                    title={
                      interactive
                        ? `${r.job.name} · tıkla: seç/kaldır · sürükle: grubu taşı`
                        : `${r.job.name} · ${r.job.role} · ${r.job.people} kişi`
                    }
                    onClick={(e) => {
                      if (!interactive || !onToggleSelect) return;
                      if (suppressClickRef.current) {
                        suppressClickRef.current = false;
                        return;
                      }
                      e.preventDefault();
                      onToggleSelect(r.job.id);
                    }}
                    onDragStart={(e) => {
                      if (!interactive) return;
                      suppressClickRef.current = true;
                      const ids =
                        selectedSet.has(r.job.id) && selectedIds.length > 0
                          ? [...selectedIds]
                          : [r.job.id];
                      dragIdsRef.current = ids;
                      e.dataTransfer.effectAllowed = "move";
                      e.dataTransfer.setData("text/plain", ids.join(","));
                    }}
                    onDragEnd={() => {
                      dragIdsRef.current = [];
                      setDropTargetId(null);
                    }}
                    onDragOver={(e) => {
                      if (!interactive) return;
                      e.preventDefault();
                      e.dataTransfer.dropEffect = "move";
                      setDropTargetId(r.job.id);
                    }}
                    onDragLeave={() => {
                      if (dropTargetId === r.job.id) setDropTargetId(null);
                    }}
                    onDrop={(e) => {
                      if (!interactive || !onMoveGroupBefore) return;
                      e.preventDefault();
                      setDropTargetId(null);
                      const raw = e.dataTransfer.getData("text/plain");
                      const ids = dragIdsRef.current.length
                        ? dragIdsRef.current
                        : raw.split(",").map((s) => s.trim()).filter(Boolean);
                      dragIdsRef.current = [];
                      if (!ids.length || ids.includes(r.job.id)) return;
                      onMoveGroupBefore(ids, r.job.id);
                    }}
                  >
                    <div
                      className="truncate whitespace-nowrap px-2 py-0.5 text-[11px]"
                      style={{ width: nameColPx, minWidth: nameColPx }}
                    >
                      {isSelected && selectedIds.length > 1 ? (
                        <span className="mr-1 inline-flex h-3.5 min-w-3.5 items-center justify-center rounded-full bg-[var(--accent)] px-0.5 text-[8px] font-bold text-white">
                          {selectedIds.indexOf(r.job.id) + 1}
                        </span>
                      ) : null}
                      {r.job.name}
                    </div>
                    {!compactLabels && (
                      <div
                        className="px-2 py-0.5 text-[10px] tabular-nums"
                        style={{ width: hourColPx, minWidth: hourColPx }}
                      >
                        {formatHours(r.job.hours)}
                      </div>
                    )}
                  </div>
                  <div
                    className="relative h-6"
                    style={{
                      width: ticks.length * weekPx,
                      backgroundColor: r.tint || "transparent",
                      backgroundImage: `repeating-linear-gradient(to right, transparent 0, transparent ${weekPx - 1}px, var(--card-border) ${weekPx - 1}px, var(--card-border) ${weekPx}px)`,
                    }}
                  >
                    <div
                      className={`absolute top-1 flex h-4 items-center overflow-hidden px-1 text-[9px] leading-none text-white ${
                        r.capacityOk ? "" : "opacity-50 outline outline-1 outline-rose-500"
                      }`}
                      style={{
                        left,
                        width,
                        background: r.color,
                      }}
                      title={`A${r.stage} · Hat ${r.lane} · ${r.job.role} · ${r.job.people} kişi · ${r.durationDays.toFixed(1)} iş günü · ${(r.durationDays / WORK_DAYS_PER_WEEK).toFixed(1)} hf · ${r.job.project || DEFAULT_PROJECT}`}
                    >
                      H{r.lane}
                    </div>
                  </div>
                </div>
              );
            })}
          </div>
        ))}
      <div className="mt-2 flex flex-wrap gap-3 text-[10px] text-[var(--muted)]">
        {(model.categoryColors.size
          ? [...model.categoryColors.entries()]
          : [...model.roleColors.entries()]
        ).map(([label, color]) => (
          <span key={label} className="inline-flex items-center gap-1.5">
            <i className="inline-block h-2.5 w-2.5 rounded-sm" style={{ background: color }} />
            {model.categoryColors.size ? `[${label}]` : label}
          </span>
        ))}
        <span>
          Haftalık kişi = paralel akış. Sol isim: tıkla seç, sürükle bırak ile grubu taşı.
        </span>
        {groupByProject && <span>Satırlar projelere göre gruplandı</span>}
      </div>
    </div>
  );
}
