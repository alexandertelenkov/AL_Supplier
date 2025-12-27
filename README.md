# Supplier Risk Ops Dashboard

Single-page React dashboard to **import your XLSX tracker**, normalize suppliers (Supplier Code = golden key), run **risk analytics**, drive **bulk actions**, and keep a full **audit log + action calendar**.

## Capabilities

- **Import XLSX** (your tracker) or **Import JSON snapshot**
- **Supplier Master** (1 supplier code → 1 canonical name)
- **Duplicate detection** (same code, different names) + recommended canonical selection
- **Risk ops tracking** (High / Non-risk / Unknown) + contract status (Sent / Signed / etc.)
- **Bulk actions** (mass risk decisions, mass contract updates, mass canonical name)
- **Policy layer** (Category/Family rules) + per-supplier overrides
- **Audit log** for every action
- **Calendar / notes** (date, owner, what was done, what next)
- **Overview drill-down**: click a chart bar (category/country) → supplier table (ID, name, risk, PO, spend)

## Tech stack

- Vite + React + TypeScript
- TailwindCSS
- Recharts (charts)
- XLSX (file import)
- Framer Motion (micro animations)
- Lucide React (icons)

## Run locally

```bash
npm install
npm run dev
```

Open the dev server URL printed in the terminal.

## Build

```bash
npm run build
npm run preview
```

## Using the dashboard

1. Click **Import XLSX/JSON**
2. Select your tracker file (XLSX)
3. Go to **Overview** to see KPIs + charts
4. Use **Suppliers** for filtering + per-supplier updates
5. Use **Bulk** to paste supplier codes and mass-update risk / contract / canonical name
6. Use **Duplicates** to resolve conflicting names under the same supplier code
7. Use **Log** for auditability (what changed, when, by whom)
8. Use **Calendar** to record actions / follow-ups

## Data model (logic)

The UI is a thin layer over a local “database” stored in `localStorage`:

- `factRows`: imported lines from your XLSX (Raw/MetaDataBase/Risk Evaluation Overview)
- `overrides`: per-supplier overrides (canonical name, risk decision, contract status, notes)
- `categoryRules`: policy defaults by Category or Family
- `tasks`: calendar items
- `log`: immutable audit trail

**Hierarchy of truth** (risk + identity):

1. Supplier overrides
2. Category/Family rules
3. Risk flags from imported data

## GitHub push

```bash
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin <YOUR_GITHUB_REPO_URL>
git push -u origin main
```

## CI

A basic GitHub Actions workflow is included (installs deps + runs `npm run build`).

## How to use

1. Click **Import XLSX/JSON**
2. Import your tracker file (recommended: the original XLSX)
3. Use:
   - **Overview** for risk analytics + drill-down
   - **Suppliers** for day-to-day ops (filters + quick actions)
   - **Duplicates** to unify names
   - **Bulk** for mass updates
   - **Log** for audit trail
   - **Calendar** to track follow-ups

## Data persistence

The dashboard stores your working state in **localStorage**. Use **Export DB** to create a portable JSON snapshot.

## GitHub deploy / CI

This repo includes a simple GitHub Actions workflow that runs `npm install` + `npm run build` on every push.
