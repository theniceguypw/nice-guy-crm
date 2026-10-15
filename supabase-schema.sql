-- The Nice Guy Washing Services - CRM database
-- Run this once in Supabase: SQL Editor > New query > paste > Run.

-- ---------- Tables ----------

create table if not exists public.customers (
  id          uuid primary key default gen_random_uuid(),
  user_id     uuid not null default auth.uid() references auth.users(id) on delete cascade,
  first_name  text not null,
  last_name   text not null default '',
  phone       text not null default '',
  email       text not null default '',
  address     text not null default '',
  segment     text not null default 'residential' check (segment in ('residential', 'commercial')),
  stage       text not null default 'new_lead'
              check (stage in ('new_lead', 'quoted', 'follow_up', 'booked', 'won', 'lost')),
  created_at  timestamptz not null default now(),
  updated_at  timestamptz not null default now()
);

create table if not exists public.notes (
  id          uuid primary key default gen_random_uuid(),
  user_id     uuid not null default auth.uid() references auth.users(id) on delete cascade,
  customer_id uuid not null references public.customers(id) on delete cascade,
  body        text not null,
  created_at  timestamptz not null default now()
);

create table if not exists public.follow_ups (
  id               uuid primary key default gen_random_uuid(),
  user_id          uuid not null default auth.uid() references auth.users(id) on delete cascade,
  customer_id      uuid not null references public.customers(id) on delete cascade,
  title            text not null,
  due_at           timestamptz not null,
  outlook_event_id text,
  done             boolean not null default false,
  created_at       timestamptz not null default now()
);

create index if not exists customers_user_idx on public.customers (user_id, segment, stage);
create index if not exists notes_customer_idx on public.notes (customer_id, created_at);
create index if not exists follow_ups_customer_idx on public.follow_ups (customer_id, due_at);

-- Keep updated_at current on customers
create or replace function public.touch_updated_at()
returns trigger language plpgsql as $$
begin
  new.updated_at = now();
  return new;
end;
$$;

drop trigger if exists customers_touch on public.customers;
create trigger customers_touch before update on public.customers
for each row execute function public.touch_updated_at();

-- ---------- Security: each logged-in user sees only their own rows ----------

alter table public.customers  enable row level security;
alter table public.notes      enable row level security;
alter table public.follow_ups enable row level security;

drop policy if exists "own customers" on public.customers;
create policy "own customers" on public.customers
  for all to authenticated
  using (user_id = auth.uid())
  with check (user_id = auth.uid());

drop policy if exists "own notes" on public.notes;
create policy "own notes" on public.notes
  for all to authenticated
  using (user_id = auth.uid())
  with check (user_id = auth.uid());

drop policy if exists "own follow_ups" on public.follow_ups;
create policy "own follow_ups" on public.follow_ups
  for all to authenticated
  using (user_id = auth.uid())
  with check (user_id = auth.uid());

-- Anonymous visitors get nothing.
revoke all on public.customers  from anon;
revoke all on public.notes      from anon;
revoke all on public.follow_ups from anon;
