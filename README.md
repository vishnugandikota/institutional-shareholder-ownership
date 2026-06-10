# Institutional Shareholder Ownership

> **Top 25 institutional holders for any public company** — sourced directly from the 13Fs of all filers, rigorously validating all holders and aggregating to present "true ownership." Provides top holders, quarter‑over‑quarter top‑10 buyers/sellers, and new entrants as an interactive web app, with the option to export to an Excel document.

![Python](https://img.shields.io/badge/python-3.9%2B-blue)
![License: MIT](https://img.shields.io/badge/License-MIT-green)
![Data: SEC EDGAR](https://img.shields.io/badge/data-SEC%20EDGAR%2013F-orange)

Institutional managers with over $100M in U.S. equities must disclose their holdings each quarter on **SEC Form 13F**. This project consolidates every filer's holdings of a given security, cleans them up (corporate‑family consolidation, as‑filed dollar values, options excluded), and produces:

1. **Top 25 institutional holders** — ranked by the latest quarter, with the three prior quarters, % ownership, and share counts.
2. **Top buyers & sellers** — the largest position changes quarter‑over‑quarter.
3. **New entrants** — institutions that initiated a position this quarter.

---

## Features

- **Any ticker** — type a ticker; it resolves the CUSIP from SEC's ticker file + a securities index built from the 13F data. Handles abbreviations (`FINL SVCS` → `FINANCIAL SERVICES`), reversed names (`SCHWAB CHARLES`), multi‑class names (GOOGL/GOOG), and **company renames** (via SEC former‑name records). Direct CUSIP entry also works.
- **Always‑fresh latest quarter** — pulls the newest quarter live from EDGAR the moment filings are in (no waiting for SEC's bulk dataset), with prior quarters from the bulk data sets.
- **Corporate‑family consolidation** — BlackRock, Vanguard, State Street, etc. are merged across all sub‑entities into one holder.
- **As‑filed dollar values** — values are summed exactly as reported (filers who report in $thousands are flagged, not silently altered).
- **Two front‑ends** — a formatted multi‑sheet **Excel workbook** and a local **web dashboard** with a live progress bar and one‑click Excel export.
- **Private by design** — runs locally; can be shared privately over [Tailscale](#private-hosting).

---

## How it works

```
ticker ──▶ SEC company_tickers.json ──▶ company name ──┐
                                                        ├──▶ CUSIP   (securities index, built once from bulk 13F data)
direct CUSIP ───────────────────────────────────────────┘
                                                         │
                 ┌───────────────────────────────────────┴───────────────────────────────┐
                 ▼                                                                         ▼
   prior 3 quarters: SEC Form 13F bulk data sets (ZIPs in this folder)       latest quarter: live scrape of EDGAR
                 │                                                                         │
                 └──────────────────┬──────────────────────────────────────────────────────┘
                                    ▼
        consolidate corporate families · exclude options · as‑filed $ · rank
                                    ▼
                 Excel workbook (build_13f.py)   ·   Web dashboard (app.py)
```

---

## Repository structure

| File | What it is |
| --- | --- |
| `app.py` | **Web app** (Flask). Type a ticker → top‑25 holders table + buyers/sellers chart, live scrape with progress bar, Excel download. |
| `build_13f.py` | **CLI engine.** Builds the multi‑sheet Excel workbook for any `--ticker` / `--cusips` from the bulk data sets (+ any scraped quarters). |
| `edgar_13f_scraper.py` | Scrapes **every** 13F filer of a CUSIP for one quarter from EDGAR, writes `scraped_<cusip>_<period>.csv`. `--build` then runs `build_13f.py`. |
| `holdings_core.py` | Shared computation (loading, family consolidation, ranking) used by the web app. |
| `aapl_13f_pipeline.py` | Thin Apple‑default wrapper around `build_13f.py` (back‑compat). |
| `requirements.txt`, `LICENSE`, `CONTRIBUTING.md`, `docs/` | Project scaffolding. |

> **Data files are not committed** (they're large and reproducible). See [Data setup](#data-setup) and `.gitignore`.

---

## Requirements

- Python 3.9+
- `pip install -r requirements.txt` (pandas, openpyxl, requests, flask)
- Internet access for the live scrape and ticker/shares‑outstanding lookups (SEC EDGAR)

```bash
git clone https://github.com/<your-username>/institutional-shareholder-ownership.git
cd institutional-shareholder-ownership
python3 -m venv .venv && source .venv/bin/activate     # optional but recommended
pip install -r requirements.txt
```

## Data setup

The prior‑quarter history comes from SEC's **Form 13F Data Sets** (one ZIP per quarter, ~80MB each). Download the four most recent quarters from
<https://www.sec.gov/data-research/sec-markets-data/form-13f-data-sets> and place the `*_form13f.zip` files in the project folder. See [`docs/DATA_SETUP.md`](docs/DATA_SETUP.md) for the exact links and quarter mapping.

The first run builds a one‑time `securities_index_*.json` from the newest ZIP (~30s) and fetches SEC's `company_tickers.json`; both are cached locally.

---

## Usage

### 1. Web app (easiest)

```bash
python3 app.py            # builds the index on first run, then serves
# open http://127.0.0.1:5050
```

Type a ticker (or a 9‑character CUSIP). The latest quarter is scraped live with a progress bar, then cached; use **↻ Refresh** to re‑pull, **⬇ Excel** to download the workbook.

Environment variables: `PORT` (default 5050), `HOST` (default 127.0.0.1; set `0.0.0.0` to reach it over a LAN/Tailnet), `LATEST_PERIOD` (e.g. `2026-03-31`).

### 2. Build an Excel workbook for any ticker

```bash
# resolve a CUSIP from the bulk data
python3 build_13f.py --resolve "NVIDIA"

# build the workbook
python3 build_13f.py --cusips 67066G104 --name "NVIDIA Corporation" --ticker NVDA
# -> NVDA_13F_Institutional_Ownership.xlsx
```

Multi‑class names take a comma list, e.g. `--cusips 02079K305,02079K107` for Alphabet.

### 3. Scrape the latest quarter (before the bulk dataset is published)

```bash
python3 edgar_13f_scraper.py \
  --cusip 67066G104 --period 2026-03-31 \
  --ticker NVDA --name "NVIDIA Corporation" \
  --ua "Your Name your@email" --build
```

This enumerates every 13F filer of the CUSIP via EDGAR full‑text search, downloads each information table, writes `scraped_<cusip>_<period>.csv`, and (`--build`) rebuilds the workbook with the new quarter as the latest column. Honors SEC's ~8 req/s fair‑access limit and logs any failed downloads to `scraped_<cusip>_<period>_FAILURES.csv`.

---

## Output

The Excel workbook contains: **Overview** (methodology), **Reference** (quarter‑end prices), **Top 25 Holders** (latest quarter first + 3 prior, change, % of 13F, sub‑entity count), **Top 25 Changes** (largest buys/sells), **New Entrants**, and **All Holders**. The web app renders the top‑25 holders table and a buyers/sellers bar chart.

---

## Private hosting

This is a local app — keep it that way and expose it only to your own devices with [Tailscale](https://tailscale.com):

```bash
python3 app.py                  # serves on 127.0.0.1:5050
tailscale serve --bg 5050       # -> https://<your-machine>.<tailnet>.ts.net  (tailnet-only, HTTPS)
```

`tailscale serve` keeps it private to your devices. (`tailscale funnel` would make it public — only do that behind authentication.)

---

## Methodology & caveats

- **Consolidation** — multiple records under one corporate parent are merged (45+ families mapped by name).
- **As‑filed values** — dollar amounts are summed exactly as filers reported them. A few managers (e.g. T. Rowe Price) file values in **$thousands**; these are shown as‑filed and flagged, not adjusted.
- **New entrants** — defined as holders present in the latest quarter but in **none** of the prior quarters shown, to avoid filing‑gap false positives.
- **Filing gaps** — a holder absent in a quarter means it did not report the security that quarter (sold out, or its filing isn't on record yet); such cells show as not reported.
- **Ticker → CUSIP** — there is no official SEC ticker↔CUSIP table, so resolution bridges via company name. It's robust for securities held by institutions, but a micro‑cap with no 13F holders cannot be resolved (paste a CUSIP if needed).
- **Latest quarter completeness** — the live scrape captures filers that have filed by run time; re‑run after the 45‑day deadline for full coverage.

---

## Roadmap

- Optional precompute that flattens the bulk ZIPs into a fast per‑CUSIP store (instant prior‑quarter lookups).
- Quarter‑over‑quarter % shading in the holders table; recent‑tickers list.
- Sector/peer comparison view.

---

## Disclaimer

This project uses **public** data from SEC EDGAR and is provided for informational and research purposes only. It is **not investment advice**. Data is presented as filed and may contain filer errors; verify against original filings before relying on it. When scraping EDGAR, respect SEC's [fair‑access policy](https://www.sec.gov/os/webmaster-faq#developers) (descriptive User‑Agent, ≤10 requests/second).

## License

[MIT](LICENSE) © 2026 Vishnu

## Acknowledgements

Data source: **U.S. Securities and Exchange Commission — EDGAR / Form 13F Data Sets**.
