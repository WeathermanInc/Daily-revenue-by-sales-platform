# Daily Revenue by Sales Platform

Automated previous-day revenue reporting across **Amazon (Sellerboard)**, **Shopify Direct**, **DICK'S SPORTING GOODS**, **Nordstrom**, and **Walmart**, with:

1. A plain-text-style **Brevo HTML email** (Image 1 layout)
2. A hosted **WEATHERMAN Daily Revenue Dashboard** under `docs/` (Image 2 layout)

The GitHub Actions workflow **`Daily Revenue Report Pipeline`** (file: `.github/workflows/daily_report.yml`) runs daily at **14:35 UTC (10:35 AM ET)** (`35 14 * * *`) and can also be triggered via `workflow_dispatch` (`target_date`, `backfill`, `force_backfill`, `send_email`).

```bash
# List / trigger (use the workflow display name, not the file name):
gh workflow list --repo WeathermanInc/Daily-revenue-by-sales-platform
gh run list --repo WeathermanInc/Daily-revenue-by-sales-platform --workflow "Daily Revenue Report Pipeline" --limit 5
gh workflow run "Daily Revenue Report Pipeline" --repo WeathermanInc/Daily-revenue-by-sales-platform \
  -f send_email=true -f target_date=YYYY-MM-DD -f backfill=false -f force_backfill=false
```

Hosted dashboard: https://mg22mex.github.io/Daily-revenue-by-sales-platform/

---

## Project Architecture & Stats

| Integration | Mechanism | Auth | Refresh / window | Rate limits & notes |
|---|---|---|---|---|
| **Shopify** | Admin GraphQL `orders` (+ line items, tags, channel info) | Dev Dashboard `client_credentials` (`SHOPIFY_CLIENT_ID` / `SHOPIFY_CLIENT_SECRET`) → short-lived `X-Shopify-Access-Token` | Target day `America/New_York` **00:00:00–23:59:59 ET** (half-open `[day, day+1)`) | Split into Shopify Direct / DSG / Nordstrom via tags + channel name. Nordstrom uses the same **createdAt** ET calendar day as the Nordstrom Summary digest (no EDI day-lag). Metric: **Gross Revenue** (`totalPriceSet` / Shopify `total_price`). Optional Mirakl OR11 merge when `MIRAKL_API_KEY` is set. |
| **Walmart** | Marketplace Orders API after OAuth | Client credentials | Target day ET | Status × fulfillment matrix, then dedupe. Category + SKU rollups included. |
| **Sellerboard** | Permanent automation CSV `GET` | Token in URL | Daily + product exports | Daily CSV → Amazon revenue / units / Real ACOS. **Product (Dashboard by Product) CSV** → Amazon SKU drilldown + category matrix. Dates default to **DD/MM** (EU). |
| **Brevo** | `POST /v3/smtp/email` | API key | Each successful run | Subject: `Daily revenue by sales platform — YYYY-MM-DD`. CTA links to hosted dashboard. |

**Failure model:** each source is isolated. Failed channels render as `Unavailable (<reason>)` in the email/dashboard (missing secrets, OAuth errors, HTTP failures) without blocking Brevo for the rest. Startup logs a source-env preflight naming which integrations are missing credentials.

---

## Data Flow Architecture

```mermaid
flowchart LR
  subgraph Sources
    S[Shopify GraphQL]
    W[Walmart Orders API]
    B[Sellerboard CSV]
  end

  subgraph Pipeline["main.py + reporting/"]
    I[Ingest + classify channels]
    M[DailyReport model]
    A[data/daily_archive.json]
    E[Email HTML Image 1]
    D[Dashboard HTML Image 2]
  end

  subgraph Delivery
    Brevo[Brevo email]
    Pages["docs/ GitHub Pages"]
  end

  S --> I
  W --> I
  B --> I
  I --> M --> A
  M --> E --> Brevo
  M --> D --> Pages
```

---

## Report outputs

### Email (Image 1)
- Greeting (`Hi Rick,`) + long-form target date
- Revenue by sales platform (5 channels + Total)
- Category unit totals
- Ad metrics (Amazon Real ACOS, Shopify blended COS, revenue per ad dollar)
- Blue CTA `#1A73E8` → **dated** hosted dashboard URL (`…/{YYYY-MM-DD}.html`), labeled `Daily Revenue Dashboard MM-DD-YYYY`

### Dashboard (Image 2)
Written to:
- `docs/index.html` (latest)
- `docs/YYYY-MM-DD.html` (dated snapshot)
- `docs/email-YYYY-MM-DD.html` (email preview for QA)

Sections: navy WEATHERMAN header, period overview cards, platform performance, ad/reconciliation cards, **category units by platform** (including Amazon), and **SKU drilldown · five channels**.

There is **no** green status banner on the dashboard; operational notes stay in logs / email only.

Enable **GitHub Pages** from the `docs/` folder on `main`, then set secret `DASHBOARD_PUBLIC_URL` to the Pages **site root** (no trailing file name), e.g. `https://mg22mex.github.io/Daily-revenue-by-sales-platform`. Each run writes `docs/{YYYY-MM-DD}.html` and copies it to `docs/index.html`; Brevo CTAs always link the dated file so older emails keep resolving.

---

## Sellerboard (Amazon) details

| Concern | Behavior |
|---|---|
| Daily totals | `SELLERBOARD_DAILY_URL` — exact calendar-day match only (never sums a multi-day export) |
| SKU + categories | `SELLERBOARD_PRODUCT_URL` — prefer **Dashboard by Product** (Date + SKU/ASIN + Units + Sales). Rows for `target_date` feed Amazon column in category matrix and Amazon lines in SKU drilldown |
| Date locale | Numeric dates default to **DD/MM/YYYY** (`11/09/2026` = 11 Sep). Day-first is auto-detected when any day component is `> 12`; override with `SELLERBOARD_DAYFIRST` |
| Day-scoped snapshot | If a CSV has **no** date column, it is treated as a same-day automation snapshot for `target_date` |
| Archive merge | If a live parse returns Amazon `$0` but a prior archive day had revenue, the prior channel total is preserved so MTD/forecast stay continuous |

---

## Roadblocks & Technical Considerations

| Risk | Mitigation |
|---|---|
| Sellerboard temporary URLs return HTML | Use permanent automation links only; HTML responses mark Amazon Unavailable |
| Sellerboard DD/MM vs MM/DD | Default DMY + day-first detection; set `SELLERBOARD_DAYFIRST=us` only for US-formatted exports |
| Amazon KPI without SKUs | Ensure product URL is itemized (Dashboard by Product / Orders), not daily aggregate-only |
| Shopify Admin API 401 / Unavailable channels | Prefer secret `SHOPIFY_ACCESS_TOKEN` (`shpat_…`) — used directly as `X-Shopify-Access-Token` (skips OAuth). Only when unset: Dev Dashboard `SHOPIFY_CLIENT_ID` + `SHOPIFY_CLIENT_SECRET` → `client_credentials`. Also set `SHOPIFY_STORE_URL`. |
| Sellerboard HTTP 401 | Rotate/verify `SELLERBOARD_DAILY_URL` + `SELLERBOARD_PRODUCT_URL` automation links. Pipeline logs a clear notice and falls back to archived Amazon totals for that date when available. |
| Amazon umbrella titles mentioning “Backpack” | Classify **umbrella** before use-case keywords; map `WM-40002-*` Venture Dry Pack to backpack |
| Shopify Direct inflation | Date-only query tokens + ET calendar-day post-filter; exclude draft/POS/void/cancelled |
| Nordstrom vs Summary mismatch | Nordstrom uses **Gross Revenue** + **createdAt** ET calendar day + optional Mirakl OR11 (same as Nordstrom Summary). Not Net Product Sales (`subtotal_price`). |
| Shopify ad spend not in Admin API | `SHOPIFY_AD_SPEND`, or dated JSON/CSV (`data/shopify_ad_spend.json`) |
| Partial API failures | Per-source try/except; email/dashboard still publish |
| MTD / forecast / last month | Rolling `data/daily_archive.json`; optional `--backfill` + validated Aug/Sep baseline seed |

---

## Environment Configuration

| Secret / env | Required | Purpose |
|---|---|---|
| `SHOPIFY_CLIENT_ID` / `SHOPIFY_CLIENT_SECRET` | Yes | Dev Dashboard app credentials → `client_credentials` Admin token |
| `SHOPIFY_STORE_URL` | Yes | `weatherman3.myshopify.com` (or shop slug). Aliases: `SHOPIFY_STORE_DOMAIN`, `SHOPIFY_SHOP_NAME` |
| `SHOPIFY_ACCESS_TOKEN` | Optional | Legacy static Admin token fallback only |
| `WALMART_CLIENT_ID` / `WALMART_CLIENT_SECRET` | Yes | Marketplace OAuth |
| `SELLERBOARD_DAILY_URL` / `SELLERBOARD_PRODUCT_URL` | Yes | Permanent CSV automation URLs (daily + product) |
| `SELLERBOARD_DAYFIRST` | Optional | `true`/`dmy` (default) or `us`/`0` for MM/DD exports |
| `BREVO_API_KEY` | Yes | Transactional email |
| `REPORT_RECIPIENTS` | Yes | Comma/semicolon emails. Defaults include marco/rick/christine/margo/sajjad/mollie + `slease@saxadvisorygroup.com` |
| `BREVO_SENDER_EMAIL` | Recommended | Verified Brevo sender |
| `BREVO_SENDER_NAME` | Optional | Default `Weatherman Revenue` |
| `DASHBOARD_PUBLIC_URL` | Recommended | GitHub Pages **site root** (no filename). CTA always appends `/{YYYY-MM-DD}.html` |
| `SHOPIFY_AD_SPEND` | Optional | Day’s Shopify ad spend (USD) |
| `SHOPIFY_ADS_JSON_PATH` / `SHOPIFY_ADS_CSV_*` | Optional | Dated ad-spend feeds |
| `MIRAKL_API_KEY` / `MIRAKL_SHOP_ID` | Optional | Align Nordstrom with Summary (OR11 Mirakl-only orders) |
| `MIRAKL_BASE_URL` | Optional | Default `https://nordstromus-prod.mirakl.net` |
| `REPORT_GREETING_NAME` | Optional | Default `Rick` |
| `REPORT_DATE` | Optional | Same as `--date YYYY-MM-DD` |
| `SKIP_EMAIL` | Optional | `1`/`true` skips Brevo |

---

## Local Development & Testing

```bash
cd Daily-revenue-by-sales-platform
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env   # fill secrets
```

**UI-only demo** (no APIs / no Brevo) — renders Image 1/2 sample (classic Sep 10 mockup, or synthetic SKUs/categories for other dates):

```bash
python main.py --demo
python main.py --date 2026-09-11 --demo --skip-email
# open docs/index.html and docs/email-*.html
```

Without API secrets, `python main.py --date …` **auto-falls back to demo** so local verification still produces dashboard artifacts.

**Targeted / live run:**

```bash
python main.py --date 2026-09-11 --skip-email
python main.py --backfill --date 2026-09-10 --skip-email
python main.py --backfill --force --no-seed-baseline --skip-email
```

---

## Repository layout

```
├── main.py                         # Ingestion + orchestration
├── reporting/
│   ├── models.py                   # DailyReport data model
│   ├── categories.py               # SKU → category mapping
│   ├── email_report.py             # Image 1 Brevo HTML
│   └── dashboard.py                # Image 2 hosted dashboard HTML
├── docs/                           # Generated dashboard (GitHub Pages)
├── data/
│   ├── daily_archive.json          # MTD / last-month rollup archive
│   └── shopify_ad_spend.json       # Optional dated Shopify ad spend
├── scripts/
│   ├── fetch_walmart_sales.py
│   └── backfill_archive.py
└── .github/workflows/daily_report.yml
```
