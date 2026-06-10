#!/usr/bin/env python3
"""
13F Institutional Ownership — local web app.
Type a ticker -> resolves its CUSIP, pulls the latest quarter live from SEC EDGAR
(prior 3 quarters come from the bulk ZIPs in this folder), and renders:
  (1) Top 25 institutional holders (latest quarter + 3 prior), with % ownership & total shares
  (2) A bar chart of the top 10 buyers and top 10 sellers (shares bought/sold) for the quarter
Plus a one-click Excel download.

RUN (on a machine with internet, in this folder):
    pip install flask pandas openpyxl requests
    python3 app.py           # then open http://127.0.0.1:5000
Latest-quarter scrapes are cached as scraped_<cusip>_<period>.csv; use the Refresh button to re-pull.
"""
import os, sys, json, re, datetime, subprocess
import pandas as pd
from flask import Flask, request, render_template_string, send_file, redirect, url_for, jsonify
import threading
from concurrent.futures import ThreadPoolExecutor
import holdings_core as hc
import edgar_13f_scraper as S

FOLDER = os.path.dirname(os.path.abspath(__file__))
UA = "Vishnu Research vishgand@gmail.com"
LATEST_PERIOD = os.environ.get("LATEST_PERIOD", "2026-03-31")   # quarter end to treat as "latest"
CACHE = os.path.join(FOLDER, "ticker_cusip_cache.json")

def load_cache():
    try: return json.load(open(CACHE))
    except Exception: return {}
def save_cache(c): json.dump(c, open(CACHE,"w"), indent=2)

def period_report(period):  # 2026-03-31 -> 31-MAR-2026
    y,m,d = map(int, period.split("-")); return f"{d:02d}-{S.MONTHS[m]}-{y}"

# ---------- securities index (CUSIP <-> issuer name) built once from bulk data ----------
SUFFIX=set('INC CORP CORPORATION CO COMPANY LTD LIMITED PLC LP LLC LLP SA AG NV THE GROUP HOLDINGS HOLDING CLASS CL A B C COM COMMON STOCK SHS SH ORD ADR ADS NEW CAP CAPITAL PAR REIT TR TRUST & N V S DEL COMPANIES COS OF AND FOR'.split())
ABBREV={'FINL':'FINANCIAL','FIN':'FINANCIAL','SVCS':'SERVICES','SVC':'SERVICES','SERV':'SERVICES',
 'GRP':'GROUP','INTL':'INTERNATIONAL','INTERNATL':'INTERNATIONAL','TECH':'TECHNOLOGY','TECHS':'TECHNOLOGIES',
 'MGMT':'MANAGEMENT','HLDG':'HOLDINGS','HLDGS':'HOLDINGS','HLD':'HOLDINGS','COMM':'COMMUNICATIONS','COMMS':'COMMUNICATIONS',
 'NATL':'NATIONAL','INDS':'INDUSTRIES','IND':'INDUSTRIES','PHARM':'PHARMACEUTICALS','PHARMA':'PHARMACEUTICALS',
 'RES':'RESOURCES','SYS':'SYSTEMS','INSTRS':'INSTRUMENTS','INSTR':'INSTRUMENTS','PROD':'PRODUCTS','PRODS':'PRODUCTS','PPTYS':'PROPERTIES','PPTY':'PROPERTIES',
 'MFG':'MANUFACTURING','AMER':'AMERICA','AMERN':'AMERICAN','BK':'BANK','BKS':'BANKS','CTR':'CENTER',
 'ENTPR':'ENTERPRISES','ENTERPRISE':'ENTERPRISES','PWR':'POWER','ENGY':'ENERGY','ELEC':'ELECTRIC',
 'MTRS':'MOTORS','TRANS':'TRANSPORTATION','LABS':'LABORATORIES','CONSOL':'CONSOLIDATED','DEV':'DEVELOPMENT',
 'NORTHN':'NORTHERN','STHN':'SOUTHERN','UTD':'UNITED','CMNTYS':'COMMUNITIES','SCIENTIFIC':'SCIENTIFIC'}
def _norm(s):
    s=re.sub(r'[^A-Z0-9 ]',' ',str(s).upper())
    out=[]
    for t in s.split():
        t=ABBREV.get(t,t)
        if t not in SUFFIX: out.append(t)
    return ' '.join(out)

IDX_PATH=os.path.join(FOLDER,'securities_index_v4.json')   # v4: + connector-word drop, count-ranked merge
def build_index():
    """cusip->name, norm-name->[cusips], sorted-token-name->[cusips], cusip->holder count. Cached.
    Common stock only (options/derivatives excluded)."""
    if os.path.exists(IDX_PATH):
        try:
            d=json.load(open(IDX_PATH)); return d['canon'], d['idx'], d['idxs'], d['count']
        except Exception: pass
    import zipfile, collections
    names=collections.defaultdict(collections.Counter)
    zz=sorted([z for z in os.listdir(FOLDER) if z.endswith('form13f.zip')])
    if zz:
        with zipfile.ZipFile(os.path.join(FOLDER,zz[-1])) as zf:
            nm=[n for n in zf.namelist() if n.upper().endswith('INFOTABLE.TSV')][0]
            with zf.open(nm) as fh:
                for ch in pd.read_csv(fh,sep='\t',dtype=str,chunksize=300000,
                        usecols=['NAMEOFISSUER','CUSIP','SSHPRNAMTTYPE','PUTCALL']):
                    ch=ch[(ch.SSHPRNAMTTYPE=='SH')&(ch.PUTCALL.fillna('').str.strip()=='')]
                    for n,c in zip(ch.NAMEOFISSUER.fillna(''),ch.CUSIP.fillna('')):
                        if c: names[c.upper()][n.upper()]+=1
    canon={c:cnt.most_common(1)[0][0] for c,cnt in names.items()}
    count={c:int(sum(cnt.values())) for c,cnt in names.items()}
    idx=collections.defaultdict(list); idxs=collections.defaultdict(list)
    for c in canon:
        nk=_norm(canon[c]); idx[nk].append(c)
        idxs[' '.join(sorted(nk.split()))].append(c)
    srt=lambda d:{k:sorted(set(v), key=lambda c:-count[c]) for k,v in d.items()}
    idx=srt(idx); idxs=srt(idxs)
    json.dump({'canon':canon,'idx':idx,'idxs':idxs,'count':count}, open(IDX_PATH,'w'))
    return canon, idx, idxs, count

def _match(title, idx, idxs, count):
    k=_norm(title)
    if not k: return []
    # merge exact (ordered) + order-independent matches, then rank by holder count
    pool=set(idx.get(k, [])) | set(idxs.get(' '.join(sorted(k.split())), []))
    if not pool:                                               # token-overlap fallback
        kt=set(k.split())
        for nm2,cs in idx.items():
            n2=set(nm2.split()) if nm2 else set()
            inter=kt & n2
            if inter and (len(inter)>=2 or kt<=n2 or n2<=kt) and len(inter)/len(kt|n2)>=0.5:
                pool.update(cs)
    if not pool: return []
    cands=sorted(pool, key=lambda c:-count.get(c,0))           # real common stock = most holders
    top=count.get(cands[0],0)                                  # drop option/derivative CUSIPs (tiny counts)
    return [c for c in cands if count.get(c,0) >= max(20, 0.05*top)]

def _company_tickers():
    """SEC ticker->name/CIK map, cached locally. Re-fetches if missing or looks like a stub (<1000 entries)."""
    p=os.path.join(FOLDER,'company_tickers.json')
    need=True
    if os.path.exists(p):
        try: need = len(json.load(open(p))) < 1000
        except Exception: need=True
    if need:
        import requests
        r=requests.get('https://www.sec.gov/files/company_tickers.json',headers={'User-Agent':UA},timeout=30)
        open(p,'w').write(r.text)
    return json.load(open(p))

# ---------- ticker (or CUSIP) -> CUSIP(s) + CIK + name ----------
def resolve_ticker(token):
    token=token.upper().strip()
    cache=load_cache()
    if token in cache: return cache[token]
    canon,idx,idxs,count=build_index()
    # direct CUSIP entry (9 chars incl a digit)
    if re.fullmatch(r'[0-9A-Z]{9}', token) and any(ch.isdigit() for ch in token):
        return {'ticker':token,'name':canon.get(token,f'CUSIP {token}').title(),'cik':None,
                'cusips':[token],'candidates':[token]}
    try: ct=_company_tickers()
    except Exception: return None
    name=cik=None
    for row in ct.values():
        if row['ticker'].upper()==token: name=row['title']; cik=str(row['cik_str']).zfill(10); break
    if not name: return None
    cands=_match(name, idx, idxs, count)
    if not cands and cik:                      # renamed company? bulk data may list a FORMER name
        try:
            import requests
            sub=requests.get(f"https://data.sec.gov/submissions/CIK{cik}.json",headers={'User-Agent':UA},timeout=30).json()
            for alt in [fn.get('name','') for fn in sub.get('formerNames',[])]:
                cands=_match(alt, idx, idxs, count)
                if cands: break
        except Exception: pass
    if not cands: return None
    info={'ticker':token,'name':name,'cik':cik,'cusips':[cands[0]],'candidates':cands[:6]}
    cache[token]=info; save_cache(cache)
    return info

def shares_outstanding(cik):
    try:
        import requests
        u=f"https://data.sec.gov/api/xbrl/companyconcept/CIK{cik}/dei/EntityCommonStockSharesOutstanding.json"
        j=requests.get(u, headers={"User-Agent":UA}, timeout=30).json()
        vals=[x for arr in j.get("units",{}).values() for x in arr if x.get("val")]
        vals.sort(key=lambda x:x.get("end",""))
        return int(vals[-1]["val"]) if vals else None
    except Exception:
        return None

# ---------- latest quarter: cache or live scrape (reports progress via prog dict) ----------
def get_latest_df(cusip, force=False, prog=None):
    path = os.path.join(FOLDER, f"scraped_{cusip}_{LATEST_PERIOD}.csv")
    if os.path.exists(path) and not force:
        if prog is not None: prog.update(phase="cached", msg="Using cached latest quarter")
        return pd.read_csv(path, dtype={'cik':str})
    rate = S.Rate(8)
    y,m,d = map(int, LATEST_PERIOD.split("-")); start=f"{y}-{m:02d}-01"
    em=m+5; ey=y
    if em>12: em-=12; ey+=1
    end=f"{ey}-{em:02d}-28"
    if prog is not None: prog.update(phase="enumerating", msg="Finding filers via SEC full-text search…", cur=0, total=0)
    flist=[f for f in S.enumerate_filings(cusip,start,end,UA,rate) if f["period"]==LATEST_PERIOD]
    flist.sort(key=lambda f:f["file_date"]); by={}
    for f in flist: by[f["cik"]]=f
    targets=list(by.values())
    if prog is not None: prog.update(phase="scraping", msg="Downloading 13F filings from EDGAR…", cur=0, total=len(targets))
    rows=[]; done=[0]; lock=threading.Lock()
    def work(f):
        res=S.fetch_one(f,cusip,UA,rate)
        with lock:
            done[0]+=1
            if prog is not None: prog["cur"]=done[0]
        if res and res[0]>0: return {"cik":f["cik"],"filer_name":f["name"],"shares":res[0],"value":res[1]}
        return None
    with ThreadPoolExecutor(max_workers=5) as ex:
        for r in ex.map(work, targets):
            if r: rows.append(r)
    df=pd.DataFrame(rows, columns=["cik","filer_name","shares","value"])
    df["period_report"]=period_report(LATEST_PERIOD); df["cusip"]=cusip
    # Only persist a cache file when the scrape actually returned holders. Caching an empty
    # result (bad/typo'd CUSIP, or a transient SEC failure) poisons the cache: later valid
    # requests would read the empty file via the cache branch and never re-scrape.
    if len(df):
        df.to_csv(path, index=False)
    return df

app = Flask(__name__)

PAGE = """
<!doctype html><html><head><meta charset="utf-8"><title>13F Ownership{% if t %} — {{t}}{% endif %}</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<style>
 :root{--bg:#0f1720;--card:#16212e;--ink:#e6edf3;--mut:#8b9aa9;--line:#27323f;--buy:#2ea36b;--sell:#d2544b;--accent:#4f8cff}
 *{box-sizing:border-box} body{margin:0;background:var(--bg);color:var(--ink);font:14px/1.45 -apple-system,Segoe UI,Roboto,Arial}
 .wrap{max-width:1100px;margin:0 auto;padding:28px}
 h1{font-size:20px;margin:0 0 4px} .sub{color:var(--mut);margin:0 0 22px}
 form.search{display:flex;gap:8px;margin:0 0 24px} input[type=text]{flex:0 0 240px;padding:10px 12px;border-radius:8px;border:1px solid var(--line);background:var(--card);color:var(--ink);font-size:15px;text-transform:uppercase}
 button{padding:10px 16px;border:0;border-radius:8px;background:var(--accent);color:#fff;font-weight:600;cursor:pointer}
 a.btn{padding:8px 12px;border:1px solid var(--line);border-radius:8px;color:var(--ink);text-decoration:none;font-size:13px;background:var(--card)}
 .card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:18px 20px;margin:0 0 22px}
 .hd{display:flex;align-items:baseline;justify-content:space-between;flex-wrap:wrap;gap:8px;margin:0 0 12px}
 .hd h2{font-size:16px;margin:0} .hd .meta{color:var(--mut);font-size:12px}
 table{width:100%;border-collapse:collapse;font-variant-numeric:tabular-nums}
 th,td{padding:7px 8px;text-align:right;border-bottom:1px solid var(--line);white-space:nowrap}
 th:nth-child(2),td:nth-child(2){text-align:left} th{color:var(--mut);font-weight:600;font-size:12px}
 td.h{font-weight:600} .up{color:var(--buy)} .dn{color:var(--sell)} .err{color:var(--sell)}
 .pill{display:inline-block;background:#22303f;color:var(--mut);border-radius:20px;padding:2px 10px;font-size:12px;margin-left:6px}
</style></head><body><div class="wrap">
 <h1>Institutional 13F Ownership</h1>
 <p class="sub">Type a ticker — top holders &amp; quarterly buyers/sellers from SEC 13F filings.</p>
 <form class="search" action="/" method="get">
   <input type="text" name="ticker" placeholder="e.g. AAPL" value="{{t or ''}}" autofocus>
   <button>Search</button>
   {% if t %}<a class="btn" href="/?ticker={{t}}&refresh=1">↻ Refresh latest</a>
   <a class="btn" href="/download?ticker={{t}}">⬇ Excel</a>{% endif %}
 </form>
 {% if error %}<div class="card err">{{error}}</div>{% endif %}
 {% if holders %}
 <div class="card">
   <div class="hd"><h2>{{name}} ({{t}}) — Top 25 Institutional Holders</h2>
     <span class="meta">CUSIP {{cusip}} · Latest: {{order[0]}} · {{nholders}} holders{% if shares_out %} · {{ '{:,}'.format(shares_out) }} shares out{% endif %}
     {% if alts %}<br>Other share classes: {% for c in alts %}<a class="btn" style="padding:1px 8px;margin-left:4px" href="/?ticker={{t}}&cusip={{c}}">{{c}}</a>{% endfor %}{% endif %}</span></div>
   <table><thead><tr><th>#</th><th>Holder</th>
     {% for q in order %}<th>{{q}}</th>{% endfor %}
     <th>% own</th></tr></thead><tbody>
   {% for h in holders %}<tr>
     <td>{{h.rank}}</td><td class="h">{{h.holder}}</td>
     {% for v in h.quarters %}<td>{{ '{:,}'.format(v) if v else '—' }}</td>{% endfor %}
     <td>{% if h.pct_out %}{{ '%.2f'|format(h.pct_out) }}%{% else %}{{ '%.2f'|format(h.pct_13f) }}%<span class="pill">of 13F</span>{% endif %}</td>
   </tr>{% endfor %}</tbody></table>
 </div>
 <div class="card">
   <div class="hd"><h2>Top 10 Buyers &amp; Sellers — {{lat}} vs {{pri}}</h2>
     <span class="meta">bars = shares bought (green) / sold (red)</span></div>
   <canvas id="bs" height="120"></canvas>
 </div>
 <script>
  const labels={{bs_labels|tojson}}, data={{bs_values|tojson}};
  new Chart(document.getElementById('bs'),{type:'bar',data:{labels:labels,datasets:[{data:data,
    backgroundColor:data.map(v=>v>=0?'#2ea36b':'#d2544b')}]},
    options:{indexAxis:'y',plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>c.parsed.x.toLocaleString()+' sh'}}},
      scales:{x:{ticks:{color:'#8b9aa9',callback:v=>v.toLocaleString()},grid:{color:'#27323f'}},
              y:{ticks:{color:'#e6edf3'},grid:{display:false}}}}});
 </script>
 {% endif %}
</div></body></html>
"""

# ---------- background jobs + progress state ----------
PROGRESS={}; RESULTS={}; JOBLOCK=threading.Lock()
RUNNING={"starting","resolving","enumerating","scraping","cached","building","finishing"}

def run_job(cusip, info, refresh):
    prog=PROGRESS[cusip]
    try:
        latest=get_latest_df(cusip, force=refresh, prog=prog)
        prog.update(phase="building", msg="Building consolidated tables (4 quarters)…", cur=0, total=0)
        res=hc.build(FOLDER,[cusip],latest_df=latest,latest_period=period_report(LATEST_PERIOD))
        if not res:
            prog.update(phase="error", error="No 13F data found for this security."); return
        prog.update(phase="finishing", msg="Fetching shares outstanding & ranking…")
        so=shares_outstanding(info["cik"])
        holders,order,tot=hc.top_holders(res,n=25,shares_out=so)
        buyers,sellers,lat,pri=hc.buyers_sellers(res,n=10)
        bs=buyers+sellers[::-1]
        RESULTS[cusip]={"t":info["ticker"],"name":info["name"],"holders":holders,"order":order,
            "nholders":len(res["wide"]),"shares_out":so,"lat":lat,"pri":pri,"cusip":cusip,
            "alts":[c for c in info.get("candidates",[cusip]) if c!=cusip][:5],
            "bs_labels":[x["holder"] for x in bs],"bs_values":[x["chg"] for x in bs]}
        prog.update(phase="done", msg="Done")
    except Exception as e:
        prog.update(phase="error", error=str(e))

@app.route("/")
def home():
    t=request.args.get("ticker","").upper().strip()
    if not t: return render_template_string(PAGE, t=None)
    info=resolve_ticker(t)
    if not info: return render_template_string(PAGE, t=t, error=f"Couldn't resolve '{t}'. Try the exact ticker, or paste the 9-character CUSIP.")
    cusip=request.args.get("cusip","").upper().strip() or info["cusips"][0]
    refresh=request.args.get("refresh")=="1"
    with JOBLOCK:
        running = PROGRESS.get(cusip,{}).get("phase") in RUNNING
        if refresh and not running: RESULTS.pop(cusip,None)
        if cusip in RESULTS and not refresh:
            return redirect(url_for("result", ticker=t, cusip=cusip))
        if not running:
            PROGRESS[cusip]={"phase":"starting","msg":"Starting…","cur":0,"total":0}
            threading.Thread(target=run_job, args=(cusip, info, refresh), daemon=True).start()
    return render_template_string(PROCESSING, t=t, cusip=cusip, name=info["name"])

@app.route("/favicon.ico")
def favicon():
    return ("", 204)  # silence the browser's automatic favicon request (was logging 404)

@app.route("/status")
def status():
    cusip=request.args.get("cusip","").upper().strip()
    return jsonify(PROGRESS.get(cusip, {"phase":"unknown"}))

@app.route("/result")
def result():
    t=request.args.get("ticker","").upper().strip()
    cusip=request.args.get("cusip","").upper().strip()
    ctx=RESULTS.get(cusip)
    if not ctx: return redirect(url_for("home", ticker=t, cusip=cusip))
    return render_template_string(PAGE, **ctx)

@app.route("/download")
def download():
    t=request.args.get("ticker","").upper().strip()
    info=resolve_ticker(t)
    if not info: return f"Couldn't resolve {t}", 404
    cusip=request.args.get("cusip","").upper().strip() or info["cusips"][0]
    get_latest_df(cusip)
    subprocess.run([sys.executable, os.path.join(FOLDER,"build_13f.py"), "--folder", FOLDER,
                    "--cusips", cusip, "--ticker", t, "--name", info["name"]])
    f=os.path.join(FOLDER, f"{t}_13F_Institutional_Ownership.xlsx")
    return send_file(f, as_attachment=True) if os.path.exists(f) else (f"Build failed for {t}", 500)

PROCESSING="""
<!doctype html><html><head><meta charset="utf-8"><title>Loading {{t}}…</title>
<style>
 :root{--bg:#0f1720;--card:#16212e;--ink:#e6edf3;--mut:#8b9aa9;--line:#27323f}
 body{margin:0;background:var(--bg);color:var(--ink);font:14px/1.5 -apple-system,Segoe UI,Roboto,Arial}
 .wrap{max-width:680px;margin:90px auto;padding:28px}
 .card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:26px}
 h2{margin:0 0 6px;font-size:18px} .meta{color:var(--mut)}
 .phase{display:inline-block;background:#22303f;border-radius:20px;padding:2px 10px;font-size:12px;color:var(--mut);margin-bottom:8px;text-transform:capitalize}
 .barwrap{height:12px;background:#0c141d;border:1px solid var(--line);border-radius:8px;overflow:hidden;margin:18px 0 8px;position:relative}
 #bar{height:100%;width:0;background:linear-gradient(90deg,#4f8cff,#2ea36b);transition:width .4s;border-radius:8px}
 #bar.indet{position:absolute;width:35%;animation:slide 1.2s ease-in-out infinite}
 @keyframes slide{0%{left:-35%}100%{left:100%}}
</style></head><body><div class="wrap"><div class="card">
 <div class="phase" id="phase">working</div>
 <h2>Loading {{name}} ({{t}})</h2>
 <div class="meta" id="msg">Starting…</div>
 <div class="barwrap"><div id="bar" class="indet"></div></div>
 <div class="meta" id="cnt"></div>
 <p class="meta" style="margin-top:16px;font-size:12px">Large names (thousands of filers) take a few minutes the first time, then they're cached.</p>
</div></div>
<script>
const cusip="{{cusip}}", t="{{t}}";
async function poll(){
  try{
    const s=await (await fetch('/status?cusip='+encodeURIComponent(cusip))).json();
    document.getElementById('phase').textContent=s.phase||'working';
    document.getElementById('msg').textContent=s.msg||'';
    const bar=document.getElementById('bar'), cnt=document.getElementById('cnt');
    if(s.total>0){ const p=Math.round(100*s.cur/s.total); bar.classList.remove('indet'); bar.style.width=p+'%';
      cnt.textContent=s.cur.toLocaleString()+' / '+s.total.toLocaleString()+' filings ('+p+'%)'; }
    else { bar.classList.add('indet'); bar.style.width=''; cnt.textContent=''; }
    if(s.phase==='done'){ location.href='/result?ticker='+encodeURIComponent(t)+'&cusip='+encodeURIComponent(cusip); return; }
    if(s.phase==='error'){ document.getElementById('msg').innerHTML='<span style="color:#d2544b">'+(s.error||'Error')+'</span>'; bar.classList.remove('indet'); bar.style.width='100%'; bar.style.background='#d2544b'; return; }
  }catch(e){}
  setTimeout(poll, 1000);
}
poll();
</script></body></html>
"""

if __name__ == "__main__":
    HOST=os.environ.get("HOST","127.0.0.1")   # set HOST=0.0.0.0 to reach it over your tailnet by IP
    PORT=int(os.environ.get("PORT","5050"))   # 5050 avoids macOS AirPlay Receiver, which holds :5000
    print("Building securities index (one-time, ~30s)…")
    try: build_index()
    except Exception as e: print("  index build skipped:", e)
    print(f"Ready → http://{HOST}:{PORT}")
    app.run(host=HOST, port=PORT, debug=False, threaded=True)
