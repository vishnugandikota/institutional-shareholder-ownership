# Data setup

The prior‑quarter history comes from the SEC's **Form 13F Data Sets** — one ZIP per filing window, each containing every institutional manager's holdings for that quarter. These files are **not** committed to the repo (they're ~80MB each and freely re‑downloadable).

## 1. Download the bulk data sets

Go to the SEC's data page and download the most recent quarterly ZIPs:

**<https://www.sec.gov/data-research/sec-markets-data/form-13f-data-sets>**

Place the `*_form13f.zip` files in the project root (next to `app.py`). Four quarters gives you the latest quarter plus three priors. As an example, the 2025 calendar quarters are:

| Report quarter (period end) | Data‑set file |
| --- | --- |
| Q1 2025 (2025‑03‑31) | `01mar2025-31may2025_form13f.zip` |
| Q2 2025 (2025‑06‑30) | `01jun2025-31aug2025_form13f.zip` |
| Q3 2025 (2025‑09‑30) | `01sep2025-30nov2025_form13f.zip` |
| Q4 2025 (2025‑12‑31) | `01dec2025-28feb2026_form13f.zip` |

> A 13F for a quarter is due **45 days** after quarter end, and the SEC publishes the bulk data set a few weeks after that. To analyze a quarter **before** the bulk file is posted, use `edgar_13f_scraper.py` to pull it live from EDGAR (see the README).

## 2. First run builds the index

On first launch, the app reads the **newest** ZIP and builds `securities_index_*.json` (a CUSIP ↔ issuer‑name map used for ticker resolution, ~30s) and fetches SEC's `company_tickers.json`. Both are cached locally and reused.

## What's generated locally (and git‑ignored)

| File | Purpose |
| --- | --- |
| `securities_index_*.json` | CUSIP ↔ issuer‑name index, built from the newest ZIP |
| `company_tickers.json` | SEC ticker → CIK/name map (fetched once) |
| `ticker_cusip_cache.json` | Resolved ticker → CUSIP cache |
| `scraped_<cusip>_<period>.csv` | A scraped latest quarter |
| `<TICKER>_13F_Institutional_Ownership.xlsx` | Generated workbook output |

Delete any of these to force a rebuild/refresh.
