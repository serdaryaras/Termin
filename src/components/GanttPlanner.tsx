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
  jobById,
  moveSelectedStage,
  moveSelectedStageBy,
  nextIsoWeek,
  parseIsoDate,
  predecessorsOf,
  prevIsoWeek,
  removeDependency,
  removeWeeklyCapacity,
  replaceWeeklyCapacity,
  scheduledIds,
  stageDurationLabel,
  successorsOf,
  upsertWeeklyCapacity,
  weekIndexFromPlanStart,
  weekRangeLabel,
  WORK_DAYS_PER_WEEK,
  yearBands,
} from "@/lib/schedule";
import { isSupabaseConfigured, loadPlan, saveLocalPlan, savePlan } from "@/lib/supabase";
import {
  DEFAULT_PROJECT,
  emptyPlan,
  jobKey,
  listProjects,
  newId,
  normalizeRole,
  ROLE_OPTIONS,
  type Plan,
  type ProgressPercent,
} from "@/lib/types";

const PRIORITY_COLS = 4;

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

/** S-şekli (boustrophedon): tek satır L→R, çift satır R→L */
function priorityCellPos(si: number, cols = PRIORITY_COLS) {
  const row = Math.floor(si / cols);
  const i = si % cols;
  const col = row % 2 === 0 ? i : cols - 1 - i;
  return { row, col };
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
  const [projectFilter, setProjectFilter] = useState<string>("all");
  const [groupByProject, setGroupByProject] = useState(false);
  const [saveStatus, setSaveStatus] = useState("Yükleniyor…");
  const [overId, setOverId] = useState<string | null>(null);
  const [pdfStatus, setPdfStatus] = useState("");
  const [pdfCompact, setPdfCompact] = useState(false);
  const [activeTab, setActiveTab] = useState<"jobs" | "priority" | "gantt" | "track">("jobs");
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

  const gantt = useMemo(() => computeGantt(plan), [plan]);
  const place = selected ? findPlacement(plan, selected) : null;
  const linkTargetIds = useMemo(() => {
    if (selectChain.length) return selectChain;
    return selected ? [selected] : [];
  }, [selectChain, selected]);
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

  const sortedCapacities = useMemo(() => {
    return [...plan.weeklyCapacities].sort((a, b) => {
      const pc = a.project.localeCompare(b.project, "tr");
      if (pc !== 0) return pc;
      const rc = a.role.localeCompare(b.role, "tr");
      if (rc !== 0) return rc;
      if (a.year !== b.year) return a.year - b.year;
      return a.week - b.week;
    });
  }, [plan.weeklyCapacities]);

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
    if (selected === jobId && selectChain.length <= 1) {
      setSelected(null);
      setSelectChain([]);
      return;
    }
    setSelectChain([jobId]);
    setSelected(jobId);
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
    if (!selected) return;
    const id = selected;
    setPlan(moveSelectedStageBy(plan, id, delta));
    scrollPriorityJobIntoView(id);
  };

  const shiftSelectedDir = (dir: "top" | "bottom" | "up" | "down") => {
    if (!selected) return;
    const id = selected;
    setPlan(moveSelectedStage(plan, id, dir));
    scrollPriorityJobIntoView(id);
  };

  const shortJobName = (id: string) => {
    const j = jobById(plan, id);
    if (!j) return "?";
    const n = j.name;
    return n.length > 22 ? `${n.slice(0, 20)}…` : n;
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
    <div className="space-y-6">
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

      <div className="no-print flex flex-wrap gap-1 border-b border-[var(--card-border)]">
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
            className={`rounded-t-lg px-4 py-2.5 text-sm font-medium transition-colors ${
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

        <div className="no-print flex flex-wrap items-start gap-x-5 gap-y-3 rounded-xl border border-[var(--card-border)] bg-[var(--card)] p-4">
          <div className="flex min-w-[11rem] flex-col gap-1">
            <span className="text-xs leading-4 text-[var(--muted)]">Başlangıç (Gantt hafta ekseni)</span>
            <input
              type="date"
              value={plan.startDate}
              onChange={(e) => setPlan({ ...plan, startDate: e.target.value })}
              className="h-9 rounded-lg border border-[var(--card-border)] bg-[var(--background)] px-3 text-sm text-[var(--foreground)]"
            />
            <button
              type="button"
              className="h-5 text-left text-[11px] leading-5 text-[var(--accent)] hover:underline disabled:opacity-40"
              disabled={!plan.weeklyCapacities.length}
              onClick={() => {
                const aligned = earliestCapacityStartDate(plan);
                if (aligned) setPlan({ ...plan, startDate: aligned });
              }}
            >
              Kapasitenin ilk haftasına hizala
            </button>
          </div>
          <div className="flex flex-col gap-1">
            <span className="text-xs leading-4 text-[var(--muted)]">Günlük saat</span>
            <input
              type="number"
              min={1}
              max={24}
              step={0.5}
              value={plan.hoursPerDay}
              onChange={(e) => setPlan({ ...plan, hoursPerDay: Math.max(1, Number(e.target.value) || 8) })}
              className="h-9 w-28 rounded-lg border border-[var(--card-border)] bg-[var(--background)] px-3 text-sm text-[var(--foreground)]"
            />
            <span className="h-5" aria-hidden />
          </div>
          <div className="flex flex-col gap-1">
            <span className="text-xs leading-4 text-[var(--muted)]">Proje filtresi</span>
            <select
              value={projectFilter}
              onChange={(e) => setProjectFilter(e.target.value)}
              className="h-9 min-w-[160px] rounded-lg border border-[var(--card-border)] bg-[var(--background)] px-3 text-sm text-[var(--foreground)]"
            >
              <option value="all">Tüm projeler</option>
              {projects.map((p) => (
                <option key={p} value={p}>
                  {p}
                </option>
              ))}
            </select>
            <span className="h-5" aria-hidden />
          </div>
          <div className="flex flex-col gap-1">
            <span className="h-4" aria-hidden />
            <label className="flex h-9 items-center gap-2 text-sm text-[var(--foreground)]">
              <input
                type="checkbox"
                checked={groupByProject}
                onChange={(e) => setGroupByProject(e.target.checked)}
                className="size-4 accent-[var(--accent)]"
              />
              Projelere göre grupla
            </label>
            <span className="h-5" aria-hidden />
          </div>
          <div className="ml-auto flex flex-col gap-1">
            <span className="h-4" aria-hidden />
            <div className="flex h-9 items-center gap-3">
              {pdfStatus && <span className="text-sm text-[var(--muted)]">{pdfStatus}</span>}
            </div>
            <span className="h-5" aria-hidden />
          </div>
        </div>

        <div className="grid w-full gap-4 lg:grid-cols-2">
        <section className="section-card section-card--tone-0 min-w-0 overflow-hidden rounded-xl border border-[var(--card-border)] bg-[var(--card)]">
          <div className="section-card__header flex items-baseline justify-between gap-3 px-4 py-3 text-sm font-semibold uppercase tracking-wider">
            <span>Kapasite</span>
            <span className="normal-case tracking-normal text-[var(--muted)]">proje · hafta · tip · kişi</span>
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

            {sortedCapacities.length > 0 && (
              <div className="max-h-[min(70vh,36rem)] space-y-2 overflow-auto">
                {sortedCapacities.map((c) => (
                  <div
                    key={`${c.project}-${c.year}-${c.week}-${c.role}`}
                    className="flex flex-wrap items-center gap-1.5 rounded-lg border border-[var(--card-border)] bg-[var(--background)] px-2 py-1.5 text-sm"
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
                    <span className="text-[var(--muted)]">·</span>
                    <input
                      key={`${c.project}-${c.year}-${c.week}-${c.role}-week`}
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
                    <input
                      type="number"
                      min={0}
                      step={1}
                      defaultValue={c.people}
                      className="w-14 rounded border border-transparent bg-transparent px-1 py-0.5 tabular-nums hover:border-[var(--card-border)] focus:border-[var(--accent)] focus:outline-none"
                      onBlur={(e) => {
                        const people = Math.max(0, Number(e.target.value) || 0);
                        if (people === c.people) return;
                        setPlan(
                          upsertWeeklyCapacity(plan, {
                            project: c.project,
                            year: c.year,
                            week: c.week,
                            role: c.role,
                            people,
                          })
                        );
                      }}
                    />
                    <span className="text-xs text-[var(--muted)]">kişi</span>
                    <button
                      type="button"
                      className="ml-auto text-xs text-rose-700"
                      onClick={() =>
                        setPlan(removeWeeklyCapacity(plan, c.project, c.year, c.week, c.role))
                      }
                    >
                      Sil
                    </button>
                  </div>
                ))}
              </div>
            )}
            <p className="text-[11px] text-[var(--muted)]">
              Excel: <strong>proje adı</strong> · <strong>hafta</strong> · <strong>personel tipi</strong> ·{" "}
              <strong>kişi</strong>. Örn. Donatım=3 → aynı haftada 3 paralel akış (işlerdeki kişi genelde 1).
              Plan başlangıcı kapasite haftalarıyla uyumlu olmalı.
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
        <section className="flex w-full min-w-0 select-none flex-col overflow-hidden rounded-xl border border-[var(--card-border)] bg-[var(--card)] max-h-[calc(100vh-6.5rem)]">
          <div className="shrink-0 border-b border-[var(--card-border)] bg-[var(--card)]">
          <div className="flex items-baseline justify-between gap-2 px-2.5 py-2">
            <h2 className="text-sm font-semibold uppercase tracking-wider">Öncelik sırası</h2>
            <div className="flex flex-wrap items-center gap-2">
              <span className="text-[10px] text-[var(--muted)]">CTRL+tık = öncül→ardıl</span>
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
                title="Tüm öncelik sırasını sil (iş kalemleri kalır)"
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
            </div>
          </div>

          <div className="no-print flex flex-wrap gap-1 border-t border-[var(--card-border)] px-2 py-1.5">
            {(
              [
                ["Üst", () => shiftSelectedDir("top"), !place],
                ["↑", () => shiftSelectedDir("up"), !place || place.stageIndex === 0],
                [
                  "↓",
                  () => shiftSelectedDir("down"),
                  !place || place.stageIndex === plan.stages.length - 1,
                ],
                ["Alt", () => shiftSelectedDir("bottom"), !place],
              ] as [string, () => void, boolean][]
            ).map(([label, fn, disabled]) => (
              <button
                key={label}
                type="button"
                disabled={disabled}
                onClick={fn}
                title={label}
                className="rounded border border-[var(--card-border)] px-1.5 py-0.5 text-[11px] disabled:opacity-40"
              >
                {label}
              </button>
            ))}
            {selectChain.length > 0 && (
              <button
                type="button"
                className="rounded border border-[var(--card-border)] px-1.5 py-0.5 text-[11px]"
                onClick={() => setSelectChain(selected ? [selected] : [])}
              >
                Seçimi sıfırla
              </button>
            )}
          </div>
          <p className="no-print border-t border-[var(--card-border)] px-2 py-1 text-[10px] text-[var(--muted)]">
            CTRL ile sırayla tıklayın: 1. öncül, 2. ardıl… Ardıl, öncül bitmeden (kapasite olsa bile) başlamaz.
          </p>

          <div className="mx-2 mb-1.5 mt-1.5 flex flex-wrap items-stretch gap-2">
            <div
              className={`min-w-[7rem] flex-1 rounded-lg border-2 border-dashed px-3 py-3 text-center text-sm font-medium ${
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
            <div className="flex flex-wrap items-center justify-center gap-1.5 rounded-lg border border-[var(--card-border)] bg-[var(--background)] px-2 py-1.5">
              {([-20, -10, -1, 1, 10, 20] as const).map((delta) => {
                const disabled =
                  !place ||
                  (delta < 0 && place.stageIndex === 0) ||
                  (delta > 0 && place.stageIndex >= plan.stages.length - 1);
                return (
                  <button
                    key={`shift-top-${delta}`}
                    type="button"
                    disabled={disabled}
                    title={
                      delta > 0
                        ? `Seçili işi ${delta} sıra aşağı kaydır`
                        : `Seçili işi ${Math.abs(delta)} sıra yukarı kaydır`
                    }
                    onClick={() => shiftSelectedBy(delta)}
                    className={`min-w-[2.75rem] rounded-md px-2 py-2 text-xs font-semibold tabular-nums disabled:opacity-40 ${
                      delta > 0
                        ? "bg-amber-100 text-amber-900 hover:bg-amber-200"
                        : "bg-sky-100 text-sky-900 hover:bg-sky-200"
                    }`}
                  >
                    {delta > 0 ? `+${delta}` : delta}
                  </button>
                );
              })}
            </div>
            <div
              className={`min-w-[7rem] flex-1 rounded-lg border-2 border-dashed px-3 py-3 text-center text-sm font-medium ${
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

          <div className="min-h-0 flex-1 overflow-auto px-2 py-1.5">
            {plan.stages.length === 0 ? (
              <p className="py-4 text-center text-xs text-[var(--muted)]">
                Soldan bir işi buraya sürükleyin.
              </p>
            ) : (
              <div className="grid w-full grid-cols-4 gap-x-1 gap-y-0">
                {plan.stages.map((stage, si) => {
                  const { row, col } = priorityCellPos(si);
                  const next = si < plan.stages.length - 1 ? priorityCellPos(si + 1) : null;
                  const goesDown = !!next && next.row > row;
                  const goesRight = !!next && next.row === row && next.col > col;
                  const goesLeft = !!next && next.row === row && next.col < col;
                  return (
                  <div
                    key={`stage-${si}`}
                    className="relative flex min-w-0 flex-col"
                    style={{ gridRow: row + 1, gridColumn: col + 1 }}
                  >
                    <div className="flex min-w-0 flex-1 items-start gap-0.5">
                    <div className="flex w-3.5 shrink-0 items-center justify-center self-center">
                      {goesLeft ? (
                        <span
                          className="text-[11px] leading-none text-[var(--muted)]"
                          title="Sıradaki (sola)"
                          aria-hidden
                        >
                          ←
                        </span>
                      ) : null}
                    </div>
                    <div
                      className={`min-w-0 flex-1 rounded-md border px-1.5 py-1 ${
                        overId === `st-${si}` || overId === `ins-${si}`
                          ? "drop-over"
                          : "border-[var(--card-border)] bg-[var(--background)]"
                      }`}
                      onDragOver={allowDrop(`st-${si}`)}
                      onDragLeave={() => setOverId(null)}
                      onDrop={dropJob(si, true)}
                    >
                      <div className="mb-0.5 flex items-center justify-between gap-1 px-0.5 text-[10px] uppercase tracking-wider text-[var(--muted)]">
                        <span>A{si + 1}</span>
                        <span className="tabular-nums">{stageDurationLabel(plan, stage)}</span>
                      </div>
                      <div className="space-y-0.5">
                        {stage.jobIds.map((id) => {
                          const job = jobById(plan, id);
                          if (!job) return null;
                          const preds = predecessorsOf(plan, id);
                          const succs = successorsOf(plan, id);
                          const chainIdx = selectChain.indexOf(id);
                          const inChain = chainIdx >= 0;
                          const isSelected = selected === id || inChain;
                          return (
                            <article
                              key={id}
                              id={`priority-job-${id}`}
                              draggable
                              onClick={(e) => onPriorityClick(id, e)}
                              onDragStart={onDragStart(id)}
                              className={`relative flex min-w-0 cursor-grab select-none flex-col gap-0.5 rounded border px-1.5 py-1 ${
                                isSelected
                                  ? "border-[var(--accent)] bg-[var(--card)] ring-1 ring-[var(--accent)]/30"
                                  : preds.length || succs.length
                                    ? "border-amber-400/70 bg-[var(--card)]"
                                    : "border-[var(--card-border)]/60 bg-[var(--card)] hover:border-[var(--card-border)]"
                              }`}
                            >
                              <div className="flex items-start gap-1">
                                {inChain && (
                                  <span
                                    className="mt-0.5 flex h-4 w-4 shrink-0 items-center justify-center rounded-full bg-[var(--accent)] text-[9px] font-bold text-white"
                                    title={`Seçim sırası ${chainIdx + 1}`}
                                  >
                                    {chainIdx + 1}
                                  </span>
                                )}
                                <h3
                                  className="min-w-0 flex-1 text-[11px] font-medium leading-snug"
                                  title={job.name}
                                >
                                  {job.name}
                                </h3>
                                <button
                                  type="button"
                                  title="İşi kalıcı sil"
                                  className="no-print shrink-0 px-0.5 text-xs leading-none text-[var(--muted)] hover:text-rose-700"
                                  onClick={(ev) => {
                                    ev.stopPropagation();
                                    if (!confirm(`“${job.name}” silinsin mi?`)) return;
                                    setPlan(deleteJob(plan, id));
                                    setSelectChain((c) => c.filter((x) => x !== id));
                                    if (selected === id) setSelected(null);
                                  }}
                                >
                                  ×
                                </button>
                              </div>
                              <div className="flex items-baseline justify-between gap-1 text-[10px]">
                                <span className="truncate text-[var(--muted)]" title={job.role}>
                                  {job.role}
                                </span>
                                <span className="shrink-0 tabular-nums text-[var(--accent)]">
                                  {formatHours(job.hours)}
                                </span>
                              </div>
                              {(preds.length > 0 || succs.length > 0) && (
                                <div className="mt-0.5 space-y-0.5 border-t border-[var(--card-border)]/60 pt-0.5">
                                  {preds.map((pid) => (
                                    <div
                                      key={`p-${pid}`}
                                      className="flex items-center gap-1 text-[9px] leading-tight text-amber-800"
                                    >
                                      <span className="shrink-0 rounded bg-amber-100 px-1 font-semibold uppercase tracking-wide">
                                        Öncül
                                      </span>
                                      <span className="min-w-0 truncate" title={shortJobName(pid)}>
                                        ← {shortJobName(pid)}
                                      </span>
                                      <button
                                        type="button"
                                        className="no-print ml-auto shrink-0 text-[var(--muted)] hover:text-rose-700"
                                        title="Bağı kaldır"
                                        onClick={(ev) => {
                                          ev.stopPropagation();
                                          setPlan(removeDependency(plan, pid, id));
                                        }}
                                      >
                                        ×
                                      </button>
                                    </div>
                                  ))}
                                  {succs.map((sid) => (
                                    <div
                                      key={`s-${sid}`}
                                      className="flex items-center gap-1 text-[9px] leading-tight text-sky-800"
                                    >
                                      <span className="shrink-0 rounded bg-sky-100 px-1 font-semibold uppercase tracking-wide">
                                        Ardıl
                                      </span>
                                      <span className="min-w-0 truncate" title={shortJobName(sid)}>
                                        → {shortJobName(sid)}
                                      </span>
                                      <button
                                        type="button"
                                        className="no-print ml-auto shrink-0 text-[var(--muted)] hover:text-rose-700"
                                        title="Bağı kaldır"
                                        onClick={(ev) => {
                                          ev.stopPropagation();
                                          setPlan(removeDependency(plan, id, sid));
                                        }}
                                      >
                                        ×
                                      </button>
                                    </div>
                                  ))}
                                </div>
                              )}
                            </article>
                          );
                        })}
                      </div>
                    </div>
                    <div className="flex w-3.5 shrink-0 items-center justify-center self-center">
                      {goesRight ? (
                        <span
                          className="text-[11px] leading-none text-[var(--muted)]"
                          title="Sıradaki (sağa)"
                          aria-hidden
                        >
                          →
                        </span>
                      ) : null}
                    </div>
                    </div>
                    <div className="flex h-3.5 shrink-0 items-start justify-center">
                      {goesDown ? (
                        <span
                          className="text-[11px] leading-none text-[var(--muted)]"
                          title="Alt satıra geç"
                          aria-hidden
                        >
                          ↓
                        </span>
                      ) : null}
                    </div>
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

          <div className="mx-2 mb-2 mt-1 shrink-0 flex flex-wrap items-stretch gap-2 border-t border-[var(--card-border)] pt-2">
            <div
              className={`min-w-[7rem] flex-1 rounded-lg border-2 border-dashed px-3 py-3 text-center text-sm font-medium ${
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
            <div className="flex flex-wrap items-center justify-center gap-1.5 rounded-lg border border-[var(--card-border)] bg-[var(--background)] px-2 py-1.5">
              {([-20, -10, -1, 1, 10, 20] as const).map((delta) => {
                const disabled =
                  !place ||
                  (delta < 0 && place.stageIndex === 0) ||
                  (delta > 0 && place.stageIndex >= plan.stages.length - 1);
                return (
                  <button
                    key={`shift-bottom-${delta}`}
                    type="button"
                    disabled={disabled}
                    title={
                      delta > 0
                        ? `Seçili işi ${delta} sıra aşağı kaydır`
                        : `Seçili işi ${Math.abs(delta)} sıra yukarı kaydır`
                    }
                    onClick={() => shiftSelectedBy(delta)}
                    className={`min-w-[2.75rem] rounded-md px-2 py-2 text-xs font-semibold tabular-nums disabled:opacity-40 ${
                      delta > 0
                        ? "bg-amber-100 text-amber-900 hover:bg-amber-200"
                        : "bg-sky-100 text-sky-900 hover:bg-sky-200"
                    }`}
                  >
                    {delta > 0 ? `+${delta}` : delta}
                  </button>
                );
              })}
            </div>
            <div
              className={`min-w-[7rem] flex-1 rounded-lg border-2 border-dashed px-3 py-3 text-center text-sm font-medium ${
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
      <section className="rounded-xl border border-[var(--card-border)] bg-[var(--card)] p-4">
        <div className="mb-4 flex flex-wrap items-baseline justify-between gap-2">
          <h2 className="text-sm font-semibold uppercase tracking-wider">Gantt çizelgesi</h2>
          <div className="flex flex-wrap items-center gap-2">
            {filteredGantt.rows.length > 0 && (
              <p className="text-xs text-[var(--muted)]">
                {projectFilter !== "all" ? `${projectFilter} · ` : ""}
                {filteredGantt.rows.length} satır · plan {formatHours(gantt.totalHours)} ·{" "}
                {(gantt.totalDays / WORK_DAYS_PER_WEEK).toFixed(1)} hafta ·{" "}
                {weekRangeLabel(buildWeekTicks(plan.startDate, gantt.totalDays))}
              </p>
            )}
            <button
              type="button"
              onClick={() => void exportGanttPdf()}
              disabled={filteredGantt.rows.length === 0}
              className="no-print rounded-lg bg-[var(--accent)] px-3 py-1.5 text-xs text-white hover:bg-[var(--accent-hover)] disabled:opacity-40"
            >
              PDF kaydet
            </button>
          </div>
        </div>
        {gantt.warnings.length > 0 && (
          <div className="mb-3 rounded-lg border border-rose-200 bg-rose-50 px-3 py-2 text-xs text-rose-800">
            {gantt.warnings.map((w) => (
              <div key={w}>{w}</div>
            ))}
          </div>
        )}
        {filteredGantt.rows.length === 0 ? (
          <p className="py-8 text-sm text-[var(--muted)]">
            Sıralama oluştukça çizelge burada üretilir. Zaman ekseni proje başlangıç haftasından başlar.
          </p>
        ) : (
          <div className="overflow-auto">
            <div ref={ganttExportRef} className="inline-block min-w-full bg-white p-2 text-slate-900">
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
              <div className="mb-2 text-sm font-semibold text-slate-800">
                İş Planı
                {projectFilter !== "all" ? ` · ${projectFilter}` : ""}
              </div>
              <p className="mb-3 text-[11px] text-slate-600">
                {filteredGantt.rows.length} satır · {formatHours(gantt.totalHours)} ·{" "}
                {(gantt.totalDays / WORK_DAYS_PER_WEEK).toFixed(1)} hafta ·{" "}
                {weekRangeLabel(buildWeekTicks(plan.startDate, gantt.totalDays))}
              </p>
              <GanttView
                planStart={plan.startDate}
                model={filteredGantt}
                groupByProject={groupByProject}
                compactLabels={pdfCompact}
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
                İlerleme düğmesine tıklayın: %100 → %75 → %50 → %25 → %0. Takılan işlere neden yazın.
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
  planStart,
  model,
  groupByProject,
  compactLabels = false,
}: {
  planStart: string;
  model: ReturnType<typeof computeGantt>;
  groupByProject: boolean;
  /** PDF: Proje / Personel / Saat kolonlarını gizle */
  compactLabels?: boolean;
}) {
  const ticks = useMemo(
    () => buildWeekTicks(planStart, model.totalDays),
    [planStart, model.totalDays]
  );
  const bands = useMemo(() => yearBands(ticks), [ticks]);
  const weekPx = Math.max(28, Math.min(48, Math.floor(900 / Math.max(ticks.length, 1))));
  const labelCols = compactLabels ? "220px" : "200px 90px 120px 70px";
  const labelSpan = compactLabels ? 1 : 4;

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

  return (
    <div className="inline-block min-w-[780px] bg-white text-xs text-slate-900">
        <div
          className="grid items-stretch border-b border-[var(--card-border)] bg-[var(--background)] font-semibold"
          style={{ gridTemplateColumns: `${labelCols} 1fr` }}
        >
          <div style={{ gridColumn: `span ${labelSpan}` }} />
          <div className="flex">
            {bands.map((b) => (
              <div
                key={b.year}
                className="border-l border-[var(--card-border)] px-1 py-1.5 text-center text-[var(--accent)]"
                style={{ width: b.count * weekPx, minWidth: b.count * weekPx }}
              >
                {b.year}
              </div>
            ))}
          </div>
        </div>
        <div
          className="grid items-center border-b border-[var(--card-border)] bg-[var(--background)] font-semibold"
          style={{ gridTemplateColumns: `${labelCols} 1fr` }}
        >
          <div className="px-2 py-2">İş kalemi</div>
          {!compactLabels && (
            <>
              <div className="px-2 py-2">Proje</div>
              <div className="px-2 py-2">Personel</div>
              <div className="px-2 py-2">Saat</div>
            </>
          )}
          <div className="flex">
            {ticks.map((t) => (
              <div
                key={t.key}
                className="border-l border-[var(--card-border)] px-0.5 py-1.5 text-center tabular-nums text-[10px]"
                style={{ width: weekPx, minWidth: weekPx }}
                title={`${t.year} hafta ${formatWeekOnly(t.week)}`}
              >
                {formatWeekOnly(t.week)}
              </div>
            ))}
          </div>
        </div>
        {sections.map((section) => (
          <div key={section.project || "all"}>
            {groupByProject && section.project && (
              <div className="border-b border-[var(--card-border)] bg-[var(--background)] px-2 py-2 text-[11px] font-semibold uppercase tracking-wider text-[var(--accent)]">
                {section.project}
                <span className="ml-2 font-normal normal-case tracking-normal text-[var(--muted)]">
                  {section.rows.length} aktivite
                </span>
              </div>
            )}
            {section.rows.map((r) => {
              const left = weekIndexFromPlanStart(planStart, r.startDay) * weekPx;
              const width = Math.max(12, (r.durationDays / WORK_DAYS_PER_WEEK) * weekPx);
              return (
                <div
                  key={r.job.id}
                  className="grid items-stretch border-b border-[var(--card-border)]"
                  style={{ gridTemplateColumns: `${labelCols} 1fr` }}
                >
                  <div className="px-2 py-2">{r.job.name}</div>
                  {!compactLabels && (
                    <>
                      <div className="px-2 py-2 text-[var(--muted)]">{r.job.project || DEFAULT_PROJECT}</div>
                      <div className="px-2 py-2 text-[var(--muted)]">
                        {r.job.role} · {r.job.people} kişi
                      </div>
                      <div className="px-2 py-2">{formatHours(r.job.hours)}</div>
                    </>
                  )}
                  <div
                    className="relative min-h-9"
                    style={{
                      width: ticks.length * weekPx,
                      backgroundImage: `repeating-linear-gradient(to right, transparent 0, transparent ${weekPx - 1}px, var(--card-border) ${weekPx - 1}px, var(--card-border) ${weekPx}px)`,
                    }}
                  >
                    <div
                      className={`absolute top-2 flex h-5 items-center overflow-hidden px-1.5 text-[10px] text-white ${
                        r.capacityOk ? "" : "opacity-50 outline outline-1 outline-rose-500"
                      }`}
                      style={{
                        left,
                        width,
                        background: r.color,
                      }}
                      title={`A${r.stage} · Hat ${r.lane} · ${r.durationDays.toFixed(1)} iş günü · ${(r.durationDays / WORK_DAYS_PER_WEEK).toFixed(1)} hf`}
                    >
                      H{r.lane}
                    </div>
                  </div>
                </div>
              );
            })}
          </div>
        ))}
      <div className="mt-3 flex flex-wrap gap-4 text-xs text-[var(--muted)]">
        {[...model.roleColors.entries()].map(([role, color]) => (
          <span key={role} className="inline-flex items-center gap-2">
            <i className="inline-block h-3 w-3" style={{ background: color }} />
            {role}
          </span>
        ))}
        <span>
          Haftalık kişi = paralel akış. CTRL ile öncül→ardıl: ardıl öncül bitmeden başlamaz.
        </span>
        {groupByProject && <span>Satırlar projelere göre gruplandı</span>}
      </div>
    </div>
  );
}
