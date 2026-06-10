# Contributing

Thanks for your interest in improving **Institutional Shareholder Ownership**.

## Getting set up

```bash
git clone https://github.com/<your-username>/institutional-shareholder-ownership.git
cd institutional-shareholder-ownership
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
```

Download a few quarterly Form 13F data‑set ZIPs (see [`docs/DATA_SETUP.md`](docs/DATA_SETUP.md)) into the project folder so you have data to test against. **Do not commit data files** — `.gitignore` excludes the ZIPs, scraped CSVs, generated indexes, and output workbooks.

## How the pieces fit

- `holdings_core.py` / `build_13f.py` — data loading, corporate‑family consolidation, and ranking. Keep the family‑mapping and CUSIP‑resolution logic in sync between them.
- `edgar_13f_scraper.py` — the EDGAR scraper. Respect SEC fair access (≤10 req/s, descriptive User‑Agent) in any change here.
- `app.py` — the Flask web app and the ticker→CUSIP resolver.

If you change name normalization or the securities index schema, bump the `IDX_PATH` version (e.g. `securities_index_v5.json`) so existing caches rebuild.

## Guidelines

- Keep changes focused; one logical change per pull request.
- Match the existing code style (compact, dependency‑light).
- Don't introduce paid/credentialed data sources — this project relies only on public SEC data.
- Add a short note to the README if you add a user‑facing feature or flag.

## Reporting issues

Open an issue with the ticker/CUSIP, the command or URL used, and what you expected vs. what happened. For resolution problems, include the company's SEC name and the CUSIP if you know it.
