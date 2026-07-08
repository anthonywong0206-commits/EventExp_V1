create table if not exists public.user_app_data (
  user_id uuid primary key references auth.users(id) on delete cascade,
  activities jsonb not null default '[]'::jsonb,
  settings jsonb not null default '{}'::jsonb,
  updated_at timestamptz not null default now()
);

alter table public.user_app_data enable row level security;

drop policy if exists "Users can read own app data" on public.user_app_data;
drop policy if exists "Users can insert own app data" on public.user_app_data;
drop policy if exists "Users can update own app data" on public.user_app_data;

create policy "Users can read own app data"
on public.user_app_data
for select
to authenticated
using ((select auth.uid()) = user_id);

create policy "Users can insert own app data"
on public.user_app_data
for insert
to authenticated
with check ((select auth.uid()) = user_id);

create policy "Users can update own app data"
on public.user_app_data
for update
to authenticated
using ((select auth.uid()) = user_id)
with check ((select auth.uid()) = user_id);

grant select, insert, update on public.user_app_data to authenticated;

do $$
begin
  if not exists (
    select 1
    from pg_publication_tables
    where pubname = 'supabase_realtime'
      and schemaname = 'public'
      and tablename = 'user_app_data'
  ) then
    alter publication supabase_realtime add table public.user_app_data;
  end if;
end $$;
