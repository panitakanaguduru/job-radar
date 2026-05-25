"""
Job Radar - Cloud Version
Reads: real_companies_1000.xlsx
Writes: jobs_all.json (accumulates 7 days), dashboard.html (GitHub Pages)
Run by GitHub Actions every 6 hours
"""
import sys, io, os, json, re, time
from datetime import datetime, timezone, timedelta
from pathlib import Path
import os
import requests
from bs4 import BeautifulSoup

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace', line_buffering=True)

# ── Config ────────────────────────────────────────────────────────────
HOURS_BACK  = 24        # fetch jobs from last 24hrs each run
KEEP_DAYS   = 7         # keep jobs for 7 days total
DELAY       = 0.3
TIMEOUT     = 12
SCRIPT_DIR  = Path(__file__).parent
EXCEL_FILE  = SCRIPT_DIR / "real_companies_1000.xlsx"
JSON_FILE   = SCRIPT_DIR / "docs" / "jobs_all.json"   # docs/ = GitHub Pages root
HTML_FILE   = SCRIPT_DIR / "docs" / "index.html"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                  "AppleWebKit/537.36 Chrome/124.0 Safari/537.36",
    "Accept-Language": "en-US,en;q=0.9",
}
S = requests.Session()
S.headers.update(HEADERS)

# ── Keywords ──────────────────────────────────────────────────────────
KEYWORDS = [
    "data analyst","business intelligence","bi analyst","bi developer",
    "product analyst","analytics engineer","data engineer","data scientist",
    "risk analyst","fraud analyst","growth analyst","revenue analyst",
    "pricing analyst","reporting analyst","marketing analyst","financial analyst",
    "quantitative analyst","research analyst","operations analyst","systems analyst",
    "business analyst","data science analyst","intelligence analyst",
    "credit analyst","investment analyst","portfolio analyst","real estate analyst",
    "property analyst","asset analyst","market analyst","valuation analyst",
    "underwriting analyst","compliance analyst","security analyst","soc analyst",
    "policy analyst","strategy analyst","supply chain analyst","logistics analyst",
    "hr analyst","people analyst","workforce analyst","claims analyst",
    "actuarial analyst","treasury analyst","tax analyst","accounting analyst",
    "budget analyst","cost analyst","sales analyst","customer analyst",
    "software engineer","software developer","backend engineer","frontend engineer",
    "full stack engineer","full-stack engineer","fullstack engineer",
    "ios engineer","android engineer","mobile engineer","web developer",
    "platform engineer","devops engineer","site reliability engineer",
    "ml engineer","machine learning engineer","ai engineer","data engineer",
    "cloud engineer","systems engineer","security engineer","sde","swe",
    "sales engineer","solutions engineer","presales engineer","pre-sales engineer",
    "technical sales","solutions architect","customer engineer","field engineer",
]

EXCLUDE_KEYWORDS = [
    "senior","sr.","sr ","lead ","principal","staff ",
    "distinguished","director","vp ","vice president",
    "head of","manager","chief","architect","fellow",
    "executive","president","cto","ceo","cfo","cdo","ciso",
    "associate director","managing director","group manager",
    "team lead","tech lead","engineering manager",
    "4+ years","5+ years","6+ years","7+ years","8+ years",
    "9+ years","10+ years","4 years","5 years","6 years",
    "experienced","seasoned","expert level","veteran",
]

NON_US = [
    "london","united kingdom"," uk ","england","scotland",
    "canada","toronto","vancouver","india","bangalore","bengaluru",
    "hyderabad","mumbai","pune","delhi","germany","berlin","munich",
    "france","paris","netherlands","amsterdam","ireland","dublin",
    "australia","sydney","melbourne","singapore","hong kong","tokyo",
    "japan","china","beijing","shanghai","brazil","mexico",
    "poland","warsaw","sweden","stockholm","norway","denmark",
    "spain","madrid","italy","milan","israel","tel aviv",
    "uae","dubai","emea","apac","latam","remote - uk","remote - canada",
    "remote - india","remote - eu",
]

def kw_match(text):
    t = text.lower()
    if any(ex in t for ex in EXCLUDE_KEYWORDS): return False
    return any(k in t for k in KEYWORDS)

def is_us(loc):
    if not loc or loc.strip() == "": return True
    l = loc.lower()
    return not any(n in l for n in NON_US)

def is_early_career(title, desc=""):
    text = (title + " " + desc).lower()
    bad = [
        "4+ years","4 or more years","four years","five years",
        "5+ years","5 years of","6+ years","six years",
        "7+ years","8+ years","9+ years","10+ years",
        "minimum 4","minimum 5","at least 4 years","at least 5 years",
        "requires 4","requires 5","requires 6",
        "4 years of experience","5 years of experience",
        "3+ years of professional experience",
        "3+ years of industry experience",
        "3 or more years of experience",
    ]
    return not any(b in text for b in bad)

def strip_html(text):
    if text and "<" in text:
        return BeautifulSoup(text, "html.parser").get_text(" ")
    return text or ""

def make_job(company, sector, title, loc, url, source, posted_at="", note="",
             opt="YES", f500="No"):
    if not posted_at:
        posted_at = datetime.now(timezone.utc).isoformat()
    return {
        "id": f"{company}::{title}::{url}",
        "company": company, "sector": sector, "title": title,
        "location": loc, "url": url, "source": source,
        "posted_at": posted_at, "fetched_at": datetime.now(timezone.utc).isoformat(),
        "remote": bool(re.search(r"\bremote\b", title+" "+loc, re.I)),
        "note": note, "opt_friendly": opt, "fortune500": f500,
    }

# ── Fetchers ──────────────────────────────────────────────────────────
def fetch_greenhouse(url, company, sector, cutoff, opt, f500):
    slug = url.rstrip("/").split("/")[-1]
    api  = f"https://boards-api.greenhouse.io/v1/boards/{slug}/jobs?content=true"
    try:
        r = S.get(api, timeout=TIMEOUT)
        if r.status_code != 200: return []
        jobs_raw = r.json().get("jobs", [])
    except: return []
    results = []
    for j in jobs_raw:
        title = j.get("title","")
        if not kw_match(title): continue
        offices = j.get("offices",[]) or []
        loc_obj = j.get("location",{}) or {}
        loc = offices[0].get("name","") if offices else loc_obj.get("name","")
        if not is_us(loc): continue
        updated = j.get("updated_at","")
        try:
            dt = datetime.fromisoformat(updated.replace("Z","+00:00"))
            if dt < cutoff: continue
        except: pass
        desc = strip_html(j.get("content",""))
        if not is_early_career(title, desc): continue
        results.append(make_job(company,sector,title,loc,
            j.get("absolute_url",url),"Greenhouse",updated,"",opt,f500))
    return results

def fetch_ashby(url, company, sector, cutoff, opt, f500):
    slug = url.rstrip("/").split("/")[-1]
    api  = f"https://api.ashbyhq.com/posting-api/job-board/{slug}"
    try:
        r = S.get(api, timeout=TIMEOUT)
        if r.status_code != 200: return []
        data = r.json()
    except: return []
    results = []
    for j in data.get("jobs",[]):
        title = j.get("title","")
        if not kw_match(title): continue
        published = j.get("publishedDate","") or j.get("updatedAt","")
        try:
            dt = datetime.fromisoformat(published.replace("Z","+00:00"))
            if dt < cutoff: continue
        except: pass
        loc = j.get("location","") or j.get("locationName","") or ""
        if isinstance(loc, dict): loc = loc.get("name","")
        if not is_us(str(loc)): continue
        desc = strip_html(j.get("descriptionHtml","") or j.get("description",""))
        if not is_early_career(title, desc): continue
        job_url = j.get("jobUrl","") or f"https://jobs.ashbyhq.com/{slug}/{j.get('id','')}"
        results.append(make_job(company,sector,title,str(loc),job_url,
            "Ashby",published,"",opt,f500))
    return results

def fetch_lever(url, company, sector, cutoff, opt, f500):
    slug = url.rstrip("/").split("/")[-1]
    try:
        r = S.get(f"https://api.lever.co/v0/postings/{slug}?mode=json", timeout=TIMEOUT)
        if r.status_code != 200: return []
        jobs_raw = r.json()
        if not isinstance(jobs_raw, list): return []
    except: return []
    results = []
    for j in jobs_raw:
        title = j.get("text","")
        if not kw_match(title): continue
        ms = j.get("createdAt",0)
        try:
            dt = datetime.fromtimestamp(ms/1000, tz=timezone.utc)
            if dt < cutoff: continue
            pa = dt.isoformat()
        except: pa = ""
        cats = j.get("categories",{})
        loc  = cats.get("location","") or j.get("workplaceType","")
        if not is_us(loc): continue
        desc = strip_html(j.get("description","") or j.get("descriptionPlain",""))
        if not is_early_career(title, desc): continue
        results.append(make_job(company,sector,title,loc,
            j.get("hostedUrl",url),"Lever",pa,"",opt,f500))
    return results

def fetch_workday(url, company, sector, cutoff, opt, f500):
    from urllib.parse import urlparse
    parsed = urlparse(url)
    host   = parsed.netloc
    parts  = [p for p in parsed.path.strip("/").split("/") if p]
    if "myworkdayjobs.com" not in host or not parts: return []
    tenant    = host.split(".")[0]
    job_board = parts[0]
    api       = f"https://{host}/wday/cxs/{tenant}/{job_board}/jobs"
    try:
        r = S.post(api, json={"appliedFacets":{},"limit":50,"offset":0,"searchText":""},
                   headers={**HEADERS,"Content-Type":"application/json"}, timeout=TIMEOUT)
        if r.status_code != 200: return []
        jobs_raw = r.json().get("jobPostings",[])
    except: return []
    results = []
    for j in jobs_raw:
        title  = j.get("title","")
        if not kw_match(title): continue
        posted = j.get("postedOn","").lower()
        if posted:
            ok = ("today" in posted or "just now" in posted or
                  "hour" in posted or posted in ("posted 1 day ago","1 day ago","0 days ago"))
            if not ok: continue
        loc = j.get("locationsText","")
        if not is_us(loc): continue
        ext     = j.get("externalPath","")
        job_url = f"https://{host}{ext}" if ext else url
        results.append(make_job(company,sector,title,loc,job_url,"Workday",
            datetime.now(timezone.utc).isoformat(),
            "Posted today" if "today" in posted else posted,opt,f500))
    return results

def fetch_smartrecruiters(url, company, sector, cutoff, opt, f500):
    slug = url.rstrip("/").split("/")[-1]
    api  = f"https://api.smartrecruiters.com/v1/companies/{slug}/postings?limit=50"
    try:
        r = S.get(api, timeout=TIMEOUT, headers={**HEADERS,"Accept":"application/json"})
        if r.status_code != 200: return []
        data = r.json()
    except: return []
    results = []
    for j in data.get("content",[]):
        title = j.get("name","")
        if not kw_match(title): continue
        created = j.get("releasedDate","") or j.get("createDate","")
        try:
            dt = datetime.fromisoformat(created.replace("Z","+00:00"))
            if dt < cutoff: continue
        except: pass
        loc_obj = j.get("location",{})
        country = loc_obj.get("country","") or loc_obj.get("countryCode","")
        if country and country.upper() not in ("US","USA","UNITED STATES"): continue
        city    = loc_obj.get("city","")
        loc     = f"{city}, {country}".strip(", ")
        ref     = j.get("id","") or j.get("refNumber","")
        job_url = f"https://jobs.smartrecruiters.com/{slug}/{ref}" if ref else url
        results.append(make_job(company,sector,title,loc,job_url,"SmartRecruiters",
            created,"",opt,f500))
    return results


FETCHERS = {
    "Greenhouse":     fetch_greenhouse,
    "Ashby":          fetch_ashby,
    "Lever":          fetch_lever,
    "Workday":        fetch_workday,
    "SmartRecruiters":fetch_smartrecruiters,
}

# ── Main ──────────────────────────────────────────────────────────────
def main():
    import pandas as pd
    cutoff = datetime.now(timezone.utc) - timedelta(hours=HOURS_BACK)
    keep_cutoff = datetime.now(timezone.utc) - timedelta(days=KEEP_DAYS)

    print(f"Job Radar Cloud - {datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M')} UTC")
    print(f"Cutoff: last {HOURS_BACK}h | Keep: {KEEP_DAYS} days")

    if not EXCEL_FILE.exists():
        print(f"ERROR: {EXCEL_FILE} not found"); return

    df = pd.read_excel(EXCEL_FILE)
    print(f"Loaded {len(df)} companies")

    # Load existing jobs (accumulate)
    JSON_FILE.parent.mkdir(parents=True, exist_ok=True)
    existing = []
    if JSON_FILE.exists():
        try:
            with open(JSON_FILE, encoding="utf-8") as f:
                data = json.load(f)
            existing = data.get("jobs", [])
            print(f"Loaded {len(existing)} existing jobs")
        except: pass

    # Remove jobs older than KEEP_DAYS
    existing = [j for j in existing
                if j.get("fetched_at","") >= keep_cutoff.isoformat()]
    print(f"After 7-day cleanup: {len(existing)} jobs")

    for i, row in df.iterrows():
        company = str(row.get("Company","")).strip()
        sector  = str(row.get("Sector","")).strip()
        ats     = str(row.get("ATS","")).strip()
        url     = str(row.get("Careers_URL","")).strip()
        opt     = str(row.get("OPT_Friendly","YES")).strip()
        f500    = str(row.get("Fortune500","No")).strip()

        if not company or not url or url == "nan": continue
        fetcher = FETCHERS.get(ats)
        if not fetcher: continue

        print(f"[{i+1}/{len(df)}] {ats:15} {company[:35]}", end="", flush=True)
        try:
            jobs = fetcher(url, company, sector, cutoff, opt, f500)
            # Only add truly new jobs (dedup by id)
            fresh = [j for j in jobs if j["id"] not in existing_ids]
            new_jobs.extend(fresh)
            for j in fresh: existing_ids.add(j["id"])
            stats[ats] = stats.get(ats,0) + len(fresh)
            print(f" -> {len(fresh)} new")
        except Exception as e:
            print(f" -> ERR: {e}")
        time.sleep(DELAY)

    # Merge and sort
    all_jobs = existing + new_jobs
    all_jobs.sort(key=lambda j: j.get("posted_at",""), reverse=True)

    print(f"\nTotal: {len(all_jobs)} jobs ({len(new_jobs)} new this run)")
    for k,v in stats.items():
        if v: print(f"  {k}: {v} new")

    # Save JSON
    output = {
        "jobs": all_jobs,
        "stats": stats,
        "total": len(all_jobs),
        "new_this_run": len(new_jobs),
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "keep_days": KEEP_DAYS,
    }
    with open(JSON_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False)
    print(f"Saved {len(all_jobs)} jobs -> {JSON_FILE}")

    # Build dashboard
    build_dashboard(all_jobs, stats, len(new_jobs))
    print(f"Dashboard -> {HTML_FILE}")

def build_dashboard(jobs, stats, new_count):
    now_str = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")
    companies = sorted(set(j["company"] for j in jobs))
    sources   = sorted(set(j["source"]  for j in jobs))
    sectors   = sorted(set(j.get("sector","") for j in jobs if j.get("sector")))
    jj = json.dumps(jobs, ensure_ascii=False)
    co_opts  = "".join(f'<option value="{c}">{c}</option>' for c in companies)
    src_opts = "".join(f'<option value="{s}">{s}</option>' for s in sources)
    sec_opts = "".join(f'<option value="{s}">{s}</option>' for s in sectors)

    html = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Job Radar - """ + str(len(jobs)) + """ Jobs</title>
<link href="https://fonts.googleapis.com/css2?family=Space+Mono:wght@400;700&family=Cabinet+Grotesk:wght@400;600;700;800;900&display=swap" rel="stylesheet">
<style>
:root{--bg:#07070f;--s1:#0e0e1c;--s2:#15152a;--bd:#252540;--bd2:#333355;--ac:#6c63ff;--a3:#43e8a0;--a4:#ffd166;--tx:#e2e2f0;--mu:#5a5a7a;}
*{box-sizing:border-box;margin:0;padding:0}
body{background:var(--bg);color:var(--tx);font-family:'Cabinet Grotesk',sans-serif;min-height:100vh}
.hdr{padding:1.2rem 2rem;display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:1rem;border-bottom:1px solid var(--bd)}
.logo{font-size:1.6rem;font-weight:900;letter-spacing:-.04em}
.logo span{background:linear-gradient(135deg,#6c63ff,#ff6584);-webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text}
.hdr-right{display:flex;align-items:center;gap:.6rem;flex-wrap:wrap}
.pills{display:flex;gap:.5rem;flex-wrap:wrap}
.pill{background:var(--s2);border:1px solid var(--bd2);border-radius:8px;padding:.45rem .8rem;text-align:center;min-width:65px}
.pill-n{font-size:1.2rem;font-weight:900;line-height:1;color:var(--ac)}
.pill-l{font-size:.55rem;font-family:'Space Mono',monospace;color:var(--mu);text-transform:uppercase;letter-spacing:.1em;margin-top:2px}
.btn{padding:.45rem .9rem;border:none;border-radius:7px;font-family:'Space Mono',monospace;font-size:.7rem;font-weight:700;cursor:pointer}
.btn-dl{background:var(--a3);color:#07070f}
.sub{background:var(--s1);border-bottom:1px solid var(--bd);padding:.45rem 2rem;font-family:'Space Mono',monospace;font-size:.65rem;color:var(--mu)}
.sub b{color:var(--a3)}
.ctrl{background:var(--s1);border-bottom:1px solid var(--bd);padding:.8rem 2rem;display:flex;gap:.5rem;flex-wrap:wrap;align-items:center;position:sticky;top:0;z-index:50}
.sb{position:relative;flex:1;min-width:160px;max-width:280px}
.sb svg{position:absolute;left:8px;top:50%;transform:translateY(-50%);color:var(--mu);pointer-events:none}
input,select{background:var(--s2);border:1px solid var(--bd2);color:var(--tx);font-family:'Space Mono',monospace;font-size:.72rem;border-radius:6px;padding:.45rem .8rem;outline:none}
.sb input{padding-left:1.8rem;width:100%}
input:focus,select:focus{border-color:var(--ac)}
select{cursor:pointer}select option{background:var(--s2)}
.chips{display:flex;gap:.3rem;flex-wrap:wrap}
.chip{padding:.3rem .65rem;border-radius:16px;border:1px solid var(--bd2);background:transparent;color:var(--mu);font-size:.65rem;font-family:'Space Mono',monospace;cursor:pointer;white-space:nowrap;transition:all .15s}
.chip:hover{border-color:var(--ac);color:var(--ac)}.chip.on{background:var(--ac);border-color:var(--ac);color:#fff}
.rc{margin-left:auto;font-family:'Space Mono',monospace;font-size:.68rem;color:var(--mu);white-space:nowrap}.rc b{color:var(--a3)}
.wrap{padding:1.2rem 2rem 3rem}
.grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(320px,1fr));gap:.75rem}
.card{background:var(--s1);border:1px solid var(--bd);border-radius:10px;padding:.9rem 1rem;display:flex;flex-direction:column;gap:.5rem;transition:border-color .2s;position:relative;overflow:hidden}
.card:hover{border-color:var(--ac)}
.bar{position:absolute;top:0;left:0;right:0;height:2px}
.bar-Greenhouse{background:linear-gradient(90deg,#43e8a0,#22d3ee)}
.bar-Ashby{background:linear-gradient(90deg,#6c63ff,#a78bfa)}
.bar-Lever{background:linear-gradient(90deg,#fb923c,#f59e0b)}
.bar-Workday{background:linear-gradient(90deg,#38bdf8,#6c63ff)}
.bar-SmartRecruiters{background:linear-gradient(90deg,#f43f5e,#fb923c)}
.ct{display:flex;justify-content:space-between;align-items:flex-start;gap:.3rem}
.co{font-size:.62rem;font-family:'Space Mono',monospace;color:var(--mu);text-transform:uppercase;letter-spacing:.07em}
.src{font-size:.55rem;font-family:'Space Mono',monospace;padding:.1rem .35rem;border-radius:3px;flex-shrink:0;background:var(--s2);color:var(--mu)}
.ttl{font-size:.88rem;font-weight:700;line-height:1.3}
.meta{display:flex;gap:.5rem;flex-wrap:wrap;align-items:center}
.m{font-size:.64rem;color:var(--mu);font-family:'Space Mono',monospace}
.badge{padding:.08rem .35rem;border-radius:3px;font-size:.58rem;font-family:'Space Mono',monospace}
.b-new{background:rgba(67,232,160,.2);color:#43e8a0;font-weight:700}
.b-rem{background:rgba(67,232,160,.1);color:var(--a3)}
.b-opt{background:rgba(67,232,160,.15);color:#43e8a0}
.b-opt2{background:rgba(108,99,255,.15);color:#a78bfa}
.b-f500{background:rgba(255,209,102,.15);color:var(--a4)}
.apply{display:inline-flex;align-items:center;gap:.25rem;padding:.4rem .8rem;background:var(--ac);color:#fff;border-radius:6px;font-size:.68rem;font-weight:700;text-decoration:none;margin-top:auto;align-self:flex-start;font-family:'Space Mono',monospace}
.apply:hover{background:#8b85ff}
.empty{grid-column:1/-1;text-align:center;padding:4rem 2rem;color:var(--mu)}
.empty h3{font-size:1rem;color:var(--tx);margin-bottom:.5rem}
::-webkit-scrollbar{width:4px}::-webkit-scrollbar-track{background:var(--bg)}::-webkit-scrollbar-thumb{background:var(--bd2);border-radius:2px}
</style>
</head>
<body>
<div class="hdr">
  <div class="logo">&#9889; <span>Job Radar</span></div>
  <div class="hdr-right">
    <div class="pills">
      <div class="pill"><div class="pill-n">TOTAL_JOBS</div><div class="pill-l">Total</div></div>
      <div class="pill"><div class="pill-n" style="color:#43e8a0">NEW_COUNT</div><div class="pill-l">New Today</div></div>
      <div class="pill"><div class="pill-n" style="color:#6c63ff">GH_COUNT</div><div class="pill-l">Greenhouse</div></div>
      <div class="pill"><div class="pill-n" style="color:#a78bfa">ASH_COUNT</div><div class="pill-l">Ashby</div></div>
    </div>
    <button class="btn btn-dl" onclick="downloadCSV()">&#11015; CSV</button>
  </div>
</div>
<div class="sub">Last updated: TIMESTAMP &nbsp;|&nbsp; Jobs kept for 7 days &nbsp;|&nbsp; Auto-updated every 6 hours via GitHub Actions</div>
<div class="ctrl">
  <div class="sb">
    <svg width="11" height="11" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="11" cy="11" r="8"/><path d="m21 21-4.35-4.35"/></svg>
    <input type="text" id="srch" placeholder="Search title or company...">
  </div>
  <select id="selCo"><option value="">All companies</option>CO_OPTS</select>
  <select id="selSrc"><option value="">All sources</option>SRC_OPTS</select>
  <select id="selSec"><option value="">All sectors</option>SEC_OPTS</select>
  <select id="selSort">
    <option value="newest">Newest first</option>
    <option value="company">Company A-Z</option>
    <option value="title">Title A-Z</option>
  </select>
  <div class="chips">
    <button class="chip on" data-f="all">All</button>
    <button class="chip" data-f="new">&#127381; New Today</button>
    <button class="chip" data-f="remote">Remote</button>
    <button class="chip" data-f="analyst">Analyst</button>
    <button class="chip" data-f="swe">Software Eng</button>
    <button class="chip" data-f="sales">Sales Eng</button>
    <button class="chip" data-f="opt">OPT Friendly</button>
    <button class="chip" data-f="f500">Fortune 500</button>

  </div>
  <div class="rc"><b id="rcN">0</b> results</div>
</div>
<div class="wrap"><div class="grid" id="grid"></div></div>
<script>
var ALL = JOBS_JSON;
var TODAY = new Date(); TODAY.setHours(0,0,0,0);
var chip='all',srch='',co='',src='',sec='',srt='newest';
var F={
  all:    function(j){return true;},
  new:    function(j){return new Date(j.fetched_at)>=TODAY;},
  remote: function(j){return !!j.remote;},
  analyst:function(j){return /analyst|data science/i.test(j.title);},
  swe:    function(j){return /software eng|software dev|backend|frontend|full.?stack|mobile eng|devops|site reliability|ml eng|machine learning|ai eng|data eng|cloud eng|systems eng|security eng|\\bsde\\b|\\bswe\\b/i.test(j.title);},
  sales:  function(j){return /sales eng|solutions eng|presales|pre-sales|technical sales|customer eng|field eng/i.test(j.title);},
  opt:    function(j){return j.opt_friendly==='HIGH'||j.opt_friendly==='YES';},
  f500:   function(j){return j.fortune500==='Yes';},

};
function ago(iso){
  if(!iso)return'';
  try{var h=Math.floor((Date.now()-new Date(iso))/3600000);
    return h<1?'<1h':h<24?h+'h ago':Math.floor(h/24)+'d ago';}catch(e){return'';}
}
function esc(s){return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');}
function isNew(j){return new Date(j.fetched_at)>=TODAY;}
function card(j){
  var a=ago(j.posted_at);
  var newBadge=isNew(j)?'<span class="badge b-new">NEW</span>':'';
  var optB=j.opt_friendly==='HIGH'?'<span class="badge b-opt">OPT+</span>':j.opt_friendly==='YES'?'<span class="badge b-opt2">OPT</span>':'';
  var f5=j.fortune500==='Yes'?'<span class="badge b-f500">F500</span>':'';
  var rem=j.remote?'<span class="badge b-rem">Remote</span>':'';
  return '<div class="card"><div class="bar bar-'+esc(j.source)+'"></div>'
    +'<div class="ct"><span class="co">'+esc(j.company)+'</span><span class="src">'+esc(j.source)+'</span></div>'
    +'<div class="ttl">'+esc(j.title)+'</div>'
    +'<div class="meta">'+(j.location?'<span class="m">&#128205;'+esc(j.location)+'</span>':'')
    +(a?'<span class="m">&#128336;'+a+'</span>':'')+newBadge+rem+optB+f5+'</div>'
    +'<a class="apply" href="'+esc(j.url)+'" target="_blank" rel="noopener">Apply &#8599;</a></div>';
}
function render(){
  var ff=F[chip]||F.all,s=srch.toLowerCase();
  var list=ALL.filter(function(j){
    if(!ff(j))return false;
    if(co&&j.company!==co)return false;
    if(src&&j.source!==src)return false;
    if(sec&&(j.sector||'')!==sec)return false;
    if(s&&j.title.toLowerCase().indexOf(s)===-1&&j.company.toLowerCase().indexOf(s)===-1)return false;
    return true;
  });
  if(srt==='newest')list.sort(function(a,b){return(b.posted_at||'').localeCompare(a.posted_at||'');});
  else if(srt==='company')list.sort(function(a,b){return a.company.localeCompare(b.company);});
  else list.sort(function(a,b){return a.title.localeCompare(b.title);});
  document.getElementById('rcN').textContent=list.length;
  var g=document.getElementById('grid');
  g.innerHTML=list.length?list.map(card).join(''):'<div class="empty"><h3>No results</h3><p>Try adjusting filters.</p></div>';
}
function downloadCSV(){
  var ff=F[chip]||F.all,s=srch.toLowerCase();
  var list=ALL.filter(function(j){
    if(!ff(j))return false;
    if(co&&j.company!==co)return false;
    if(src&&j.source!==src)return false;
    if(sec&&(j.sector||'')!==sec)return false;
    if(s&&j.title.toLowerCase().indexOf(s)===-1&&j.company.toLowerCase().indexOf(s)===-1)return false;
    return true;
  });
  if(!list.length){alert('No jobs to download');return;}
  var hdr=['Company','Title','Location','URL','Source','Posted','Remote','OPT','F500','Sector'];
  var rows=list.map(function(j){
    return[j.company,j.title,j.location||'',j.url,j.source,(j.posted_at||'').slice(0,10),
      j.remote?'Yes':'No',j.opt_friendly||'',j.fortune500||'',j.sector||'']
    .map(function(v){return'"'+String(v).replace(/"/g,'""')+'"';}).join(',');
  });
  var a=document.createElement('a');
  a.href=URL.createObjectURL(new Blob([[hdr.join(',')].concat(rows).join('\\n')],{type:'text/csv'}));
  a.download='jobs_'+new Date().toISOString().slice(0,10)+'.csv';
  a.click();
}
document.getElementById('srch').addEventListener('input',function(e){srch=e.target.value;render();});
document.getElementById('selCo').addEventListener('change',function(e){co=e.target.value;render();});
document.getElementById('selSrc').addEventListener('change',function(e){src=e.target.value;render();});
document.getElementById('selSec').addEventListener('change',function(e){sec=e.target.value;render();});
document.getElementById('selSort').addEventListener('change',function(e){srt=e.target.value;render();});
document.querySelectorAll('.chip').forEach(function(b){
  b.addEventListener('click',function(){
    document.querySelectorAll('.chip').forEach(function(x){x.classList.remove('on');});
    b.classList.add('on');chip=b.dataset.f;render();
  });
});
render();
</script>
</body>
</html>"""

    html = html.replace("TOTAL_JOBS", str(len(jobs)))
    html = html.replace("NEW_COUNT",  str(new_count))
    html = html.replace("GH_COUNT",   str(stats.get("Greenhouse",0)))
    html = html.replace("ASH_COUNT",  str(stats.get("Ashby",0)))
    html = html.replace("WD_COUNT",   str(stats.get("Workday",0)))
    html = html.replace("TIMESTAMP",  now_str)
    html = html.replace("CO_OPTS",    co_opts)
    html = html.replace("SRC_OPTS",   src_opts)
    html = html.replace("SEC_OPTS",   sec_opts)
    html = html.replace("JOBS_JSON",  jj)

    with open(HTML_FILE, "w", encoding="utf-8") as f:
        f.write(html)

if __name__ == "__main__":
    main()
