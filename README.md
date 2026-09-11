# BDA File Cleaner

A tiny Next.js app that reproduces the pandas cleaning script as a web upload
tool: pick a CSV/Excel export, it runs the same cleaning rules server-side,
and a `<name>_cleaned.xlsx` file downloads automatically.

## What's in this project

```
app/
  layout.tsx          Root layout, loads global styles
  page.tsx             The single upload page (client component)
  globals.css          Visual styling
  api/process/route.ts  POST endpoint: file in, cleaned .xlsx out
lib/
  bdaProcessor.ts      The actual cleaning logic (line-by-line port of the python script)
```

Every column-manipulation step in `lib/bdaProcessor.ts` is commented with the
exact line from the original Python script it replaces, so you can diff the
two side by side.

## Importing into Vercel

1. Push this folder to a new GitHub/GitLab/Bitbucket repo (or use `vercel deploy`
   directly from this folder with the Vercel CLI — no git required).
2. In the Vercel dashboard: **Add New… → Project → Import** the repo (or drag
   this folder in if using CLI). Vercel auto-detects Next.js — no config needed.
3. Deploy. No environment variables are required.

To run it locally first:

```bash
npm install
npm run dev
# open http://localhost:3000
```

## Assumptions carried over from the Python script

These match the original script's behaviour exactly, including a couple of
its quirks:

- **Date format is strict.** `createdAt`, `updatedAt`, and
  `date_time_of_closing` must match `DD/MM/YYYY HH:mm` exactly (e.g.
  `05/03/2026 14:30`). A non-matching value raises an error and stops
  processing, same as `pd.to_datetime(..., exact=True)` would.
- **Required columns.** The uploaded file must contain: `createdAt`,
  `updatedAt`, `date_time_of_closing`, `original_phone_number`,
  `phoneNumber`, `paymentStatus`, `paymentAmount`, `refundAmount`,
  `closing_amt`, `rest_collection`, `alternateEmails`,
  `alternatePhoneNumbers`, `paymentId`, `paymentproof`, `remark`.
- **The "still 0" phone number quirk.** If `original_phone_number` is itself
  blank/NaN, a `phoneNumber` of `0` becomes blank rather than `0` — so it
  will *not* get one of the sequential 101, 102, 103… placeholders. This
  mirrors the original script's behaviour exactly (it's arguably a bug in
  the source script, but the port preserves it rather than silently "fixing"
  it).
- **First sheet only.** If an `.xlsx`/`.xls` file has multiple sheets, only
  the first one is read (same as `pd.read_excel()`'s default).

## Adjusting for your data

If your export uses different column names, edit the `requiredColumns` list
and the field references inside `processBda()` in `lib/bdaProcessor.ts` —
every reference to a column name is a plain string, so a find-and-replace is
enough.
