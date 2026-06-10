"""Shared 13F holdings logic for the web app. Computes consolidated institutional
ownership for one security across the 4 most recent quarters from the local bulk
ZIPs plus the latest (scraped) quarter."""
import os, re, glob, zipfile
import pandas as pd, numpy as np

FAMILY_PATTERNS=[
 (r'BLACKROCK','BlackRock Inc.'),(r'VANGUARD','Vanguard Group Inc.'),
 (r'STATE STREET','State Street Corp.'),(r'GEODE','Geode Capital Management'),
 (r'BERKSHIRE HATHAWAY','Berkshire Hathaway Inc.'),
 (r'\bFMR\b|FIDELITY MANAGEMENT|FIDELITY INSTITUTIONAL|FIDELITY INVESTMENTS|FIDELITY DIVERSIFYING','FMR LLC (Fidelity)'),
 (r'MORGAN STANLEY','Morgan Stanley'),(r'GOLDMAN SACHS','Goldman Sachs Group Inc.'),
 (r'JP ?MORGAN|J\.?P\.? MORGAN','JPMorgan Chase & Co.'),(r'BANK OF AMERICA|MERRILL LYNCH','Bank of America Corp.'),
 (r'NORGES BANK','Norges Bank (Norway)'),
 (r'WELLINGTON MANAGEMENT|WELLINGTON TRUST|CABOT-WELLINGTON','Wellington Management Group'),
 (r'NORTHERN TRUST','Northern Trust Corp.'),(r'BANK OF NEW YORK MELLON|BNY MELLON|MELLON CAPITAL|DREYFUS','Bank of New York Mellon'),
 (r'CHARLES SCHWAB|\bSCHWAB\b','Charles Schwab'),
 (r'^CAPITAL WORLD INVESTORS|^CAPITAL RESEARCH (&|AND|GLOBAL)|^CAPITAL INTERNATIONAL (INVESTORS|INC|LTD|SARL)|^CAPITAL GROUP (COMPANIES|INTERNATIONAL|INVESTMENT|PRIVATE)|^CAPITAL GUARDIAN','Capital Group (American Funds)'),
 (r'\bUBS\b','UBS Group AG'),(r'DEUTSCHE BANK|DWS ','Deutsche Bank AG'),
 (r'T ROWE|ROWE PRICE','T. Rowe Price Group'),(r'INVESCO','Invesco Ltd.'),
 (r'DIMENSIONAL FUND','Dimensional Fund Advisors'),(r'FRANKLIN RESOURCES|FRANKLIN ADVIS|FRANKLIN MUTUAL','Franklin Resources'),
 (r'AMUNDI','Amundi'),(r'LEGAL & GENERAL|LEGAL AND GENERAL','Legal & General Group'),
 (r'NUVEEN|\bTIAA\b|TEACHERS ADVISORS|TEACHERS INSURANCE','Nuveen / TIAA'),(r'WELLS FARGO','Wells Fargo & Co.'),
 (r'CITIGROUP|CITIBANK','Citigroup Inc.'),(r'AMERIPRISE|COLUMBIA THREADNEEDLE','Ameriprise Financial'),
 (r'PRINCIPAL FINANCIAL|PRINCIPAL GLOBAL|PRINCIPAL TRUST','Principal Financial Group'),(r'AMERICAN CENTURY','American Century'),
 (r'MASSACHUSETTS FINANCIAL|\bMFS\b','MFS (Massachusetts Financial)'),(r'ALLIANZ','Allianz SE'),
 (r'ROYAL BANK OF CANADA|\bRBC\b','Royal Bank of Canada'),(r'TORONTO[- ]DOMINION','Toronto-Dominion Bank'),
 (r'BANK OF MONTREAL|\bBMO\b','Bank of Montreal'),(r'\bHSBC\b','HSBC Holdings'),
 (r'BNP PARIBAS','BNP Paribas'),(r'MITSUBISHI UFJ|\bMUFG\b','Mitsubishi UFJ'),
 (r'SUMITOMO MITSUI','Sumitomo Mitsui'),(r'\bNOMURA\b','Nomura'),
 (r'FISHER ASSET|FISHER INVESTMENTS','Fisher Investments'),(r'JANUS HENDERSON','Janus Henderson'),
 (r'SCHRODER','Schroders'),(r'RAYMOND JAMES','Raymond James'),
]
def clean(n):
    return re.sub(r'\s+',' ',re.sub(r'[^A-Z0-9 &.\-/]',' ',str(n).upper())).strip()
def fam(n):
    u=clean(n)
    for p,c in FAMILY_PATTERNS:
        if re.search(p,u): return c
    return None

def _zip_members(zf):
    g=lambda b:[n for n in zf.namelist() if n.upper().endswith(b)]
    return g('INFOTABLE.TSV')[0], g('SUBMISSION.TSV')[0], g('COVERPAGE.TSV')[0]

# ---- in-memory caches for the long-running web app ----
# Re-reading the four ~85MB bulk ZIPs on every request is the dominant latency (~30s).
# Cache by a ZIP "signature" (name+mtime+size) so caches auto-invalidate when a new
# quarter's ZIP is dropped in. _META_CACHE (submission+coverpage) is CUSIP-independent
# and shared across all tickers; _AF_CACHE keys on the CUSIP set for instant repeat hits.
_META_CACHE={}   # (zip_path, mtime) -> merged submission+coverpage DataFrame
_AF_CACHE={}     # (frozenset(cusips), zip_signature) -> af DataFrame

def _zip_sig(folder):
    return tuple(sorted((os.path.basename(z), int(os.path.getmtime(z)), os.path.getsize(z))
                        for z in glob.glob(os.path.join(folder,'*form13f.zip'))))

def _load_meta(zf, snm, cnm, zpath):
    key=(zpath, int(os.path.getmtime(zpath)))
    m=_META_CACHE.get(key)
    if m is not None: return m
    with zf.open(snm) as fh: sub=pd.read_csv(fh,sep='\t',dtype=str)
    with zf.open(cnm) as fh: cov=pd.read_csv(fh,sep='\t',dtype=str)[['ACCESSION_NUMBER','AMENDMENTNO','FILINGMANAGER_NAME']]
    m=sub.merge(cov,on='ACCESSION_NUMBER',how='left')
    _META_CACHE[key]=m
    return m

def load_af(folder, cusips):
    """Return per-accession AAPL-style holding rows for the given CUSIP set, from all bulk ZIPs.
    Memoized in-memory by (CUSIP set, ZIP signature); returns a copy so callers can't corrupt the cache."""
    cusips=set(cusips)
    ckey=(frozenset(cusips), _zip_sig(folder))
    hit=_AF_CACHE.get(ckey)
    if hit is not None: return hit.copy()
    acc=[]; meta=[]
    for z in sorted(glob.glob(os.path.join(folder,'*form13f.zip'))):
        with zipfile.ZipFile(z) as zf:
            inm,snm,cnm=_zip_members(zf)
            parts=[]
            with zf.open(inm) as fh:
                for ch in pd.read_csv(fh,sep='\t',dtype=str,chunksize=200000,
                        usecols=['ACCESSION_NUMBER','CUSIP','VALUE','SSHPRNAMT','SSHPRNAMTTYPE','PUTCALL']):
                    ch=ch[ch.CUSIP.isin(cusips)]
                    if len(ch): parts.append(ch)
            if parts:
                info=pd.concat(parts); info=info[(info.SSHPRNAMTTYPE=='SH')&(info.PUTCALL.fillna('').str.strip()=='')].copy()
                info['SSHPRNAMT']=pd.to_numeric(info.SSHPRNAMT,errors='coerce').fillna(0)
                info['VALUE']=pd.to_numeric(info.VALUE,errors='coerce').fillna(0)
                acc.append(info.groupby('ACCESSION_NUMBER').agg(shares=('SSHPRNAMT','sum'),value=('VALUE','sum')).reset_index())
            meta.append(_load_meta(zf, snm, cnm, z))
    if not acc:
        af=pd.DataFrame(columns=['ACCESSION_NUMBER','shares','value','FILING_DATE','AMENDMENTNO_n','PERIODOFREPORT','CIK','FILINGMANAGER_NAME','pdate'])
        _AF_CACHE[ckey]=af
        return af.copy()
    a=pd.concat(acc).drop_duplicates('ACCESSION_NUMBER'); m=pd.concat(meta).drop_duplicates('ACCESSION_NUMBER')
    af=a.merge(m,on='ACCESSION_NUMBER',how='left')
    af['FILING_DATE']=pd.to_datetime(af.FILING_DATE,format='%d-%b-%Y',errors='coerce')
    af['AMENDMENTNO_n']=pd.to_numeric(af.AMENDMENTNO,errors='coerce').fillna(0)
    af['pdate']=pd.to_datetime(af.PERIODOFREPORT,format='%d-%b-%Y',errors='coerce')
    af=af[af.shares>0]
    _AF_CACHE[ckey]=af
    return af.copy()

def add_latest(af, latest_df, period_report):
    """Append a scraped latest quarter (cols cik,filer_name,shares,value) as synthetic rows."""
    if latest_df is None or not len(latest_df): return af
    s=latest_df.rename(columns={'cik':'CIK','filer_name':'FILINGMANAGER_NAME'}).copy()
    s['CIK']=s.CIK.astype(str).str.split('.').str[0].str.zfill(10)
    s['ACCESSION_NUMBER']='SCRAPED-'+s.CIK.astype(str); s['FILING_DATE']=pd.Timestamp('2100-01-01')
    s['AMENDMENTNO_n']=0; s['PERIODOFREPORT']=period_report
    s['pdate']=pd.to_datetime(period_report,format='%d-%b-%Y',errors='coerce')
    s['shares']=pd.to_numeric(s.shares,errors='coerce').fillna(0); s['value']=pd.to_numeric(s.value,errors='coerce').fillna(0)
    return pd.concat([af, s], ignore_index=True, sort=False)

def qlabel(d): return f"Q{(d.month-1)//3+1} {d.year}"

def consolidate(af, pdate):
    p=af[af.pdate==pdate].sort_values(['FILING_DATE','AMENDMENTNO_n'])
    last=p.groupby('CIK').tail(1).copy()
    last['family']=last.FILINGMANAGER_NAME.apply(fam)
    last['holder_key']=np.where(last.family.notna(),last.family,'CIK:'+last.CIK)
    last['holder_name']=np.where(last.family.notna(),last.family,last.FILINGMANAGER_NAME.str.title())
    return last.groupby('holder_key').agg(shares=('shares','sum'),value=('value','sum'),
        n_ciks=('CIK','nunique'),holder_name=('holder_name','first')).reset_index()

def build(folder, cusips, latest_df=None, latest_period=None):
    af=load_af(folder, cusips)
    if latest_df is not None and latest_period:
        af=add_latest(af, latest_df, latest_period)
    if not len(af): return None
    cnt=af.groupby('pdate').ACCESSION_NUMBER.nunique()
    mx=int(cnt.max()); thr=max(10,0.2*mx)
    good=sorted([d for d,c in cnt.items() if c>=thr])[-4:]
    labels=[qlabel(d) for d in good]
    wide=None
    for d,lab in zip(good,labels):
        c=consolidate(af,d).rename(columns={'shares':f'sh_{lab}','value':f'val_{lab}','n_ciks':f'nc_{lab}'})
        c=c[['holder_key','holder_name',f'sh_{lab}',f'val_{lab}',f'nc_{lab}']]
        wide=c if wide is None else wide.merge(c,on='holder_key',how='outer').assign(
            holder_name=lambda x:x.holder_name_y.combine_first(x.holder_name_x)).drop(columns=['holder_name_x','holder_name_y'])
    for lab in labels:
        for pre in ['sh_','val_','nc_']: wide[pre+lab]=wide[pre+lab].fillna(0)
    price={}
    for lab in labels:
        imp=(wide[f'val_{lab}']/wide[f'sh_{lab}'])[(wide[f'sh_{lab}']>0)&(wide[f'val_{lab}']>0)]
        imp=imp[(imp>1)&(imp<100000)]; price[lab]=float(imp.median()) if len(imp) else 0.0
    order=labels[::-1]  # latest first
    return {'wide':wide,'order':order,'price':price}

def top_holders(res, n=25, shares_out=None):
    wide=res['wide']; order=res['order']; latest=order[0]
    tot=wide[f'sh_{latest}'].sum()
    t=wide.sort_values(f'sh_{latest}',ascending=False).head(n)
    out=[]
    for i,(_,r) in enumerate(t.iterrows(),1):
        row={'rank':i,'holder':r['holder_name'],
             'shares_latest':int(r[f'sh_{latest}']),
             'pct_out':(100*r[f'sh_{latest}']/shares_out) if shares_out else None,
             'pct_13f':100*r[f'sh_{latest}']/tot if tot else 0,
             'quarters':[int(r[f'sh_{q}']) for q in order]}
        out.append(row)
    return out, order, tot

def buyers_sellers(res, n=10):
    wide=res['wide']; order=res['order']; latest, prior = order[0], order[1]
    both=wide[(wide[f'sh_{latest}']>0)&(wide[f'sh_{prior}']>0)].copy()
    both['chg']=both[f'sh_{latest}']-both[f'sh_{prior}']
    buyers=both.sort_values('chg',ascending=False).head(n)
    sellers=both.sort_values('chg',ascending=True).head(n)
    mk=lambda df:[{'holder':r['holder_name'],'chg':int(r['chg'])} for _,r in df.iterrows()]
    return mk(buyers), mk(sellers), latest, prior
