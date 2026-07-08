# Supabase Cloud Sync

This version stores application data in Supabase after the user signs in with
Email/Password. Devices signed in to the same account share activities,
settings, and receipt PDF data through Supabase Realtime.

## Environment variables

Copy `.env.example` to `.env.local` for local development, or add these values
to the deployment provider:

```bash
VITE_SUPABASE_URL=https://jciqwdzuptvmwdmmqdaj.supabase.co
VITE_SUPABASE_ANON_KEY=your_supabase_publishable_or_anon_key
```

Use a Supabase publishable key or legacy anon key. Never expose a secret key or
service role key in this frontend application.

## Database

The schema is stored in:

```text
supabase/migrations/20260709000000_create_user_app_data.sql
```

The migration creates `public.user_app_data`, enables RLS, restricts every row
to its authenticated owner, and adds the table to the `supabase_realtime`
publication.

The migration has already been applied to project
`jciqwdzuptvmwdmmqdaj`.

## Local fallback

The existing localStorage keys remain as a local cache. If Supabase is not
configured or the user is signed out, the application continues to work with
local browser data.
