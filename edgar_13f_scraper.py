#!/usr/bin/env python3
"""
EDGAR 13F scraper — builds a full institutional-ownership snapshot for ONE security
for ONE report quarter, directly from individual SEC filings, the day after the
13F deadline (no need to wait for the SEC bulk data set).

WHAT IT DOES
  1. Uses EDGAR full-text search (efts.sec.gov) to enumerate EVERY 13F-HR filing that
     reports the target CUSIP for the target quarter.
  2. Downloads each filing's information table and extracts that holder's shares + value.
  3. Writes scraped_<period>.csv  (cik, filer_name, period_report, shares, value).
  4. With --build, runs build_13f.py for the scraped CUSIP, which merges this quarter with the
     prior quarters (from the bulk ZIPs already in the folder) and rebuilds the Excel
     workbook (Top 25 holders latest-first, Q-o-Q change, and New Entrants).

WHERE TO RUN
  Run on a machine with normal internet access (your computer). Requires: python3, pandas,
  openpyxl, requests.  Install:  pip install pandas openpyxl requests

SEC FAIR ACCESS
  SEC requires a descriptive User-Agent with contact info and <= 10 requests/sec.
  Set --ua "Your Name your@email".  Default uses the project owner's email.

USAGE
  # Apple (defaults):
  python3 edgar_13f_scraper.py --period 2026-03-31 --ticker AAPL --name "Apple Inc." --build
  # Any other ticker -- pass --cusip, --ticker, --name so --build labels the workbook correctly:
  python3 edgar_13f_scraper.py --cusip 918284100 --period 2026-03-31 \
        --ticker VSEC --name "VSE Corporation" --ua "Vishnu vishgand@gmail.com" --build
  # (--build runs build_13f.py for the scraped CUSIP -> <TICKER>_13F_Institutional_Ownership.xlsx)
"""
import argparse, os, re, sys, time, json, threading, queue
from concurrent.futures import ThreadPoolExecutor
import requests, pandas as pd

EFTS = "https://efts.sec.gov/LATEST/search-index"
ARCH = "https://www.sec.gov/Archives/edgar/data/{cik}/{adsh}/{fn}"
MONTHS = {1:'JAN',2:'FEB',3:'MAR',4:'APR',5:'MAY',6:'JUN',7:'JUL',8:'AUG',9:'SEP',10:'OCT',11:'NOV',12:'DEC'}

class Rate:
    """<= max_per_sec across all threads."""
    def __init__(self, max_per_sec=8):
        self.interval = 1.0/max_per_sec; self.lock = threading.Lock(); self.next = 0.0
    def wait(self):
        with self.lock:
            now = time.time(); self.next = max(now, self.next) ; t = self.next; self.next += self.interval
        d = t - time.time()
        if d > 0: time.sleep(d)

def get(url, ua, rate, params=None, tries=6):
    for i in range(tries):
        rate.wait()
        try:
            r = requests.get(url, headers={'User-Agent': ua, 'Accept-Encoding':'gzip,deflate'}, params=params, timeout=120)
            if r.status_code == 200: return r
            # EDGAR returns transient 5xx (incl. 500) and throttling codes; back off and retry
            if r.status_code in (429,403,500,502,503,504): time.sleep(2*(i+1)); continue
            return r
        except requests.RequestException:
            time.sleep(2*(i+1))
    return r if 'r' in dir() else None

def enumerate_filings(cusip, startdt, enddt, ua, rate):
    """Return list of dicts: cik, adsh, fn, period, form, file_date, name.

    EDGAR FTS returns ~100 hits/page and caps reachable results at 10000 (from<10000).
    Transient 5xx on a single page must NOT abort the whole sweep: skip that page and
    keep going so we don't silently truncate the filer universe."""
    seen, out, frm = set(), [], 0
    total, page = None, 100
    while frm < 10000:
        params = {'q': f'"{cusip}"', 'forms':'13F-HR', 'startdt':startdt, 'enddt':enddt, 'from':frm}
        r = get(EFTS, ua, rate, params=params)
        if r is None or r.status_code != 200:
            # page failed after retries; advance past it rather than losing the tail
            print(f"  WARN page from={frm} failed (status {getattr(r,'status_code','none')}); skipping")
            frm += page
            if total is not None and frm >= total: break
            continue
        j = r.json(); hits = j.get('hits',{}).get('hits',[])
        if total is None:
            total = j.get('hits',{}).get('total',{}).get('value', 0)
        if not hits:
            if total and frm < total:  # gap from a skipped page; keep stepping
                frm += page; continue
            break
        for h in hits:
            s = h['_source']; _id = h['_id']  # "{adsh}:{filename}"
            adsh, fn = _id.split(':',1)
            if adsh in seen: continue
            seen.add(adsh)
            out.append({'cik': s['ciks'][0], 'adsh': adsh, 'fn': fn,
                        'period': s.get('period_ending',''), 'form': s.get('form',''),
                        'file_date': s.get('file_date',''),
                        'name': re.sub(r'\s+\(CIK.*$','', s['display_names'][0]).strip()})
        frm += len(hits)
        print(f"  enumerated {frm}/{total}", end='\r')
        if frm >= total: break
    print()
    return out

# ---- information-table parsing (namespace-agnostic; handles raw XML and rendered HTML) ----
def parse_holding(text, cusip):
    shares = value = 0.0
    if '<' in text and ('infoTable' in text or 'informationTable' in text.lower()):
        for blk in re.findall(r'<(?:\w+:)?infoTable\b.*?</(?:\w+:)?infoTable>', text, re.S|re.I):
            if cusip not in blk: continue
            def g(tag):
                m = re.search(rf'<(?:\w+:)?{tag}\b[^>]*>(.*?)</(?:\w+:)?{tag}>', blk, re.S|re.I)
                return (m.group(1).strip() if m else '')
            if g('cusip').replace(' ','') != cusip: continue
            if re.search(r'<(?:\w+:)?putCall\b[^>]*>\s*(Put|Call)', blk, re.I): continue
            typ = g('sshPrnamtType').upper()
            if typ and typ != 'SH': continue
            sh = re.sub(r'[^0-9]','', g('sshPrnamt')); va = re.sub(r'[^0-9]','', g('value'))
            shares += int(sh) if sh else 0; value += int(va) if va else 0
    else:  # rendered HTML table fallback
        for row in re.findall(r'<tr[^>]*>.*?</tr>', text, re.S|re.I):
            if cusip not in row: continue
            c = [re.sub(r'<[^>]+>','',x).replace('&nbsp;','').strip() for x in re.findall(r'<td[^>]*>(.*?)</td>', row, re.S|re.I)]
            if len(c) < 8 or c[2].replace(' ','') != cusip: continue
            if c[7].strip() in ('Put','Call'): continue
            if c[6].strip().upper() not in ('SH',''): continue
            shares += int(re.sub(r'[^0-9]','',c[5]) or 0); value += int(re.sub(r'[^0-9]','',c[4]) or 0)
    return shares, value

def fetch_one(f, cusip, ua, rate):
    url = ARCH.format(cik=int(f['cik']), adsh=f['adsh'].replace('-',''), fn=f['fn'])
    r = get(url, ua, rate)
    if r is None or r.status_code != 200: return None
    sh, va = parse_holding(r.text, cusip)
    return sh, va

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--cusip', default='037833100')
    ap.add_argument('--period', required=True, help='report quarter end, e.g. 2026-03-31')
    ap.add_argument('--startdt', default=None, help='filing-date window start (default = period month)')
    ap.add_argument('--enddt', default=None, help='filing-date window end (default = period +5 months)')
    ap.add_argument('--ua', default='Vishnu Research vishgand@gmail.com')
    ap.add_argument('--workers', type=int, default=5)
    ap.add_argument('--rate', type=float, default=8.0)
    ap.add_argument('--folder', default=os.path.dirname(os.path.abspath(__file__)))
    ap.add_argument('--ticker', default='', help='ticker label for the workbook (e.g. VSEC); defaults to the CUSIP')
    ap.add_argument('--name', default='', help='issuer display name (e.g. "VSE Corporation")')
    ap.add_argument('--build', action='store_true', help='rebuild <TICKER>_13F_Institutional_Ownership.xlsx via build_13f.py after scraping')
    a = ap.parse_args()
    y, m, d = map(int, a.period.split('-'))
    startdt = a.startdt or f'{y}-{m:02d}-01'
    em = m+5; ey=y
    if em>12: em-=12; ey+=1
    enddt = a.enddt or f'{ey}-{em:02d}-28'
    rate = Rate(a.rate)

    print(f'Enumerating 13F-HR filers of {a.cusip} for period {a.period} (filed {startdt}..{enddt}) ...')
    flist = enumerate_filings(a.cusip, startdt, enddt, a.ua, rate)
    # keep target period; per CIK keep latest filing (by file_date) that reports the cusip
    flist = [f for f in flist if f['period'] == a.period]
    flist.sort(key=lambda f: f['file_date'])
    by_cik = {}
    for f in flist: by_cik[f['cik']] = f   # last wins = latest filing
    targets = list(by_cik.values())
    print(f'{len(targets)} unique filers report this security for {a.period}. Downloading info tables ...')

    rows = []; fails = []; done = [0]; lock = threading.Lock()
    def work(f):
        res = fetch_one(f, a.cusip, a.ua, rate)
        with lock:
            done[0]+=1
            if done[0] % 50 == 0: print(f'  downloaded {done[0]}/{len(targets)}', end='\r')
        if res is None:
            with lock: fails.append(f)
            return None
        if res[0] > 0:
            return {'cik': f['cik'], 'filer_name': f['name'], 'shares': res[0], 'value': res[1]}
        return None
    with ThreadPoolExecutor(max_workers=a.workers) as ex:
        for r in ex.map(work, targets):
            if r: rows.append(r)
    print()
    if fails:
        pd.DataFrame(fails)[['cik','name','adsh','fn']].to_csv(
            os.path.join(a.folder, f'scraped_{a.cusip}_{a.period}_FAILURES.csv'), index=False)
        print(f'WARNING: {len(fails)} filing(s) failed to download after retries — see '
              f'scraped_{a.cusip}_{a.period}_FAILURES.csv. Re-run to retry them.')
    df = pd.DataFrame(rows, columns=['cik','filer_name','shares','value'])
    pr = f'{d:02d}-{MONTHS[m]}-{y}'
    df['period_report'] = pr; df['cusip'] = a.cusip
    outcsv = os.path.join(a.folder, f'scraped_{a.cusip}_{a.period}.csv')
    df.to_csv(outcsv, index=False)
    if len(df):
        print(f'Wrote {outcsv}: {len(df)} holders, {df.shares.sum():,.0f} shares, ${df.value.sum()/1e9:,.1f}B')
    else:
        print(f'Wrote {outcsv}: 0 holders — nothing extracted (check period/window).')

    if a.build and len(df):
        import subprocess
        tkr = a.ticker or a.cusip
        nm = a.name or a.ticker or a.cusip
        print(f'Building workbook for {tkr} ...')
        subprocess.run([sys.executable, os.path.join(a.folder,'build_13f.py'),
                        '--folder', a.folder, '--cusips', a.cusip, '--ticker', tkr, '--name', nm])

if __name__ == '__main__':
    main()
