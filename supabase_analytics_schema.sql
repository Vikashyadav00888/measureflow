create table if not exists public.mf_events (
  id bigint generated always as identity primary key,
  user_id text not null,
  session_id text not null,
  event_type text not null,
  route text,
  tab text,
  details jsonb not null default '{}'::jsonb,
  user_agent text,
  referrer text,
  created_at timestamptz not null default now()
);

create index if not exists mf_events_created_at_idx on public.mf_events (created_at desc);
create index if not exists mf_events_user_id_idx on public.mf_events (user_id);
create index if not exists mf_events_event_type_idx on public.mf_events (event_type);
