-- Run in Supabase SQL editor once.

create table if not exists public.gantt_plans (
  id uuid primary key default gen_random_uuid(),
  slug text unique not null default 'default',
  name text not null default 'İş planı',
  start_date date not null default current_date,
  hours_per_day numeric not null default 8,
  jobs jsonb not null default '[]'::jsonb,
  stages jsonb not null default '[]'::jsonb,
  resource_groups jsonb not null default '[]'::jsonb,
  weekly_capacities jsonb not null default '[]'::jsonb,
  dependencies jsonb not null default '[]'::jsonb,
  updated_at timestamptz not null default now()
);

alter table public.gantt_plans
  add column if not exists resource_groups jsonb not null default '[]'::jsonb;

alter table public.gantt_plans
  add column if not exists weekly_capacities jsonb not null default '[]'::jsonb;

alter table public.gantt_plans
  add column if not exists dependencies jsonb not null default '[]'::jsonb;

insert into public.gantt_plans (slug, name)
values ('default', 'İş planı')
on conflict (slug) do nothing;

alter table public.gantt_plans enable row level security;

drop policy if exists "gantt_plans_select" on public.gantt_plans;
drop policy if exists "gantt_plans_insert" on public.gantt_plans;
drop policy if exists "gantt_plans_update" on public.gantt_plans;

create policy "gantt_plans_select" on public.gantt_plans for select using (true);
create policy "gantt_plans_insert" on public.gantt_plans for insert with check (true);
create policy "gantt_plans_update" on public.gantt_plans for update using (true);
