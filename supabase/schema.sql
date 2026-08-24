-- Run this once in Supabase: Dashboard → SQL Editor → New query → paste → Run

-- One table stores every data collection as a JSON blob (settings, customers,
-- products, documents, transactions, billingNotes), keyed by name — this
-- mirrors the key/value design the app already uses.
create table if not exists app_data (
  key text primary key,
  value jsonb not null,
  updated_at timestamptz not null default now()
);

alter table app_data enable row level security;

-- No public sign-up will be enabled (see README.md), so every account is one
-- you created yourself for a specific family member. "authenticated" here
-- effectively means "someone I invited".
create policy "authenticated can read app_data"
  on app_data for select
  using (auth.role() = 'authenticated');

create policy "authenticated can insert app_data"
  on app_data for insert
  with check (auth.role() = 'authenticated');

create policy "authenticated can update app_data"
  on app_data for update
  using (auth.role() = 'authenticated');

create policy "authenticated can delete app_data"
  on app_data for delete
  using (auth.role() = 'authenticated');

-- Keep updated_at fresh on every write
create or replace function set_updated_at()
returns trigger as $$
begin
  new.updated_at = now();
  return new;
end;
$$ language plpgsql;

drop trigger if exists app_data_updated_at on app_data;
create trigger app_data_updated_at
before update on app_data
for each row execute function set_updated_at();
