# Lab School Video Behavior Database React Clone

This is a React + Vite clone of the Lab School app, built to run as a static site on GitHub Pages.

## Live site

- [https://ryanlay.github.io/tcfdlabschool.github.io/](https://ryanlay.github.io/tcfdlabschool.github.io/)

## Deployment

- Repository: `https://github.com/ryanlay/tcfdlabschool.github.io`
- GitHub Actions workflow: `.github/workflows/deploy.yml`
- Deploy trigger: push to `main`

## Run locally

```powershell
cd "c:\Users\rlay\OneDrive - The Center For Discovery\Projects\Active\lab-school-database-react"
npm install
npm run dev
```

## Build for GitHub Pages

```powershell
npm run build
```

The app uses `base: '/tcfdlabschool.github.io/'` for project-site hosting on GitHub Pages.

## Features

- Intake wizard with subjects and behavior selection
- Query-style review tables
- Searchable data view
- Admin lists for subjects and behaviors
- Shared persistent data in Supabase
- SharePoint folder links based on the recording date

## Shared Database Setup (Supabase)

This app stores `subjects`, `behaviors`, and `videos` in a single shared row in Supabase, so data is usable across browsers and devices.

### 1) Create table and policies in Supabase SQL Editor

Run this SQL in your Supabase project:

```sql
create table if not exists public.lab_school_state (
  state_key text primary key,
  subjects jsonb not null default '[]'::jsonb,
  behaviors jsonb not null default '[]'::jsonb,
  videos jsonb not null default '[]'::jsonb,
  updated_at timestamptz not null default now()
);

alter table public.lab_school_state enable row level security;

create policy "public read" on public.lab_school_state
for select using (true);

create policy "public insert" on public.lab_school_state
for insert with check (true);

create policy "public update" on public.lab_school_state
for update using (true) with check (true);
```

### 2) Create local environment file

Create `.env.local` in the project root:

```env
VITE_SUPABASE_URL=https://uqedhpsjugpnlzohearq.supabase.co
VITE_SUPABASE_ANON_KEY=sb_publishable_OvVPFSDDRNxQGnw8lIyXOA_KWB0G6GP
VITE_SUPABASE_TABLE=lab_school_state
```

If your URL is copied as `https://...supabase.co/rest/v1/`, use `https://...supabase.co` in `VITE_SUPABASE_URL`.

## Notes

- The frontend must use the publishable/anon key only.
- Never put the Supabase secret/service key in client code.
