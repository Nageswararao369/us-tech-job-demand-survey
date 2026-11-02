#!/usr/bin/env python3
"""
Role Demand Survey
------------------
Reads fortune500_careers.xlsx, finds which job system (ATS) each company uses,
pulls every public US job opening, groups titles into role families, and
writes a ranked Excel report.

Setup (once):
    pip3 install pandas openpyxl requests playwright
    python3 -m playwright install chromium

Run:
    python3 role_demand_survey.py fortune500_careers.xlsx

Steps it runs (each step is cached, so you can stop and re-run safely):
    1. detect  - opens each career page in a headless browser and looks for
                 Workday / Greenhouse / Lever / Ashby / SmartRecruiters links
    2. fetch   - downloads job postings from those systems' public feeds
    3. report  - classifies titles and writes role_demand_report.xlsx

Optional: add a column named "ATS Override" to the sheet and paste a job-board
URL (e.g. https://walmart.wd5.myworkdayjobs.com/WalmartExternal) for any
company that detection misses. Overrides always win.
"""

import argparse
import asyncio
import json
import re
import sys
import time
from collections import Counter, defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

import pandas as pd
import requests

CACHE_DIR = Path("survey_cache")
DETECT_CACHE = CACHE_DIR / "detected_ats.json"
JOBS_DIR = CACHE_DIR / "jobs"
UA = ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/124.0 Safari/537.36")

# --------------------------------------------------------------------------
# 1. ATS detection
# --------------------------------------------------------------------------

ATS_PATTERNS = [
    ("workday", re.compile(
        r"https?://([\w-]+)\.(wd\d+)\.myworkdayjobs\.com/(?:wday/cxs/[\w-]+/)?"
        r"(?:[a-z]{2}-[A-Z]{2}/)?([\w-]+)", re.I)),
    ("greenhouse", re.compile(
        r"(?:boards|job-boards)(?:\.eu)?\.greenhouse\.io/(?:embed/job_board\?for=)?([\w-]+)", re.I)),
    ("greenhouse", re.compile(r"boards-api\.greenhouse\.io/v1/boards/([\w-]+)", re.I)),
    ("greenhouse", re.compile(r"greenhouse\.io/embed/job_board(?:/js)?\?for=([\w-]+)", re.I)),
    ("lever", re.compile(r"jobs\.lever\.co/([\w.-]+)", re.I)),
    ("lever", re.compile(r"api\.lever\.co/v0/postings/([\w.-]+)", re.I)),
    ("ashby", re.compile(r"jobs\.ashbyhq\.com/([\w.%-]+)", re.I)),
    ("smartrecruiters", re.compile(
        r"(?:jobs|careers)\.smartrecruiters\.com/([\w-]+)", re.I)),
    ("smartrecruiters", re.compile(
        r"api\.smartrecruiters\.com/v1/companies/([\w-]+)", re.I)),
]
BAD_TOKENS = {"embed", "js", "v1", "wday", "cxs", "jobs", "job", "api", "login",
              "en-us", "careers", "search", "static", "assets", "widget"}
FOLLOW_WORDS = re.compile(
    r"search (all )?jobs|view (all )?(jobs|openings|positions)|open (positions|roles)|"
    r"job (openings|search)|find (a )?job|explore (jobs|roles|careers|opportunities)|"
    r"current openings|see (all )?(jobs|openings)|browse jobs|apply now", re.I)


def find_ats(text):
    """Return (ats, info) for the first ATS reference found in text."""
    for ats, pat in ATS_PATTERNS:
        for m in pat.finditer(text):
            if ats == "workday":
                tenant, wd, site = m.groups()
                if site.lower() in BAD_TOKENS:
                    continue
                return ats, {"tenant": tenant, "wd": wd, "site": site}
            token = m.group(1).strip("/.")
            if token.lower() in BAD_TOKENS or len(token) < 2:
                continue
            return ats, {"token": token}
    return None, None


async def detect_one(browser, company, url, sem):
    async with sem:
        seen = []
        ctx = await browser.new_context(user_agent=UA)
        page = await ctx.new_page()
        page.on("request", lambda r: seen.append(r.url))
        page.on("framenavigated", lambda f: seen.append(f.url))

        async def scan(target):
            try:
                await page.goto(target, wait_until="domcontentloaded", timeout=30000)
                await page.wait_for_timeout(3500)
            except Exception:
                pass
            try:
                html = await page.content()
            except Exception:
                html = ""
            hrefs = []
            try:
                hrefs = await page.eval_on_selector_all(
                    "a[href], iframe[src]",
                    "els => els.map(e => [e.href || e.src, (e.innerText||'').trim()])")
            except Exception:
                pass
            blob = "\n".join([page.url, html] + seen + [h for h, _ in hrefs])
            return find_ats(blob), hrefs

        result = {"company": company, "url": url, "ats": None, "info": None}
        try:
            (ats, info), hrefs = await scan(url)
            if not ats:
                # follow up to 3 "search jobs"-style links one level deep
                cands = [h for h, t in hrefs if h and h.startswith("http")
                         and (FOLLOW_WORDS.search(t) or re.search(r"/(jobs|search|openings)", h))]
                for link in list(dict.fromkeys(cands))[:3]:
                    (ats, info), _ = await scan(link)
                    if ats:
                        break
            result.update(ats=ats, info=info, final_url=page.url)
        except Exception as e:
            result["error"] = str(e)[:200]
        finally:
            await ctx.close()
        print(f"  [{result['ats'] or '-':15}] {company}", flush=True)
        return result


async def detect_all(rows, concurrency):
    from playwright.async_api import async_playwright
    cache = json.loads(DETECT_CACHE.read_text()) if DETECT_CACHE.exists() else {}
    todo = [r for r in rows if r["company"] not in cache and not r.get("override")]
    print(f"Detecting job systems for {len(todo)} companies "
          f"({len(cache)} already cached)...")
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        sem = asyncio.Semaphore(concurrency)
        tasks = [detect_one(browser, r["company"], r["url"], sem) for r in todo]
        for coro in asyncio.as_completed(tasks):
            res = await coro
            cache[res["company"]] = res
            DETECT_CACHE.write_text(json.dumps(cache, indent=1))
        await browser.close()
    return cache


# --------------------------------------------------------------------------
# 2. Fetching jobs from public ATS feeds
# --------------------------------------------------------------------------

S = requests.Session()
S.headers.update({"User-Agent": UA, "Accept": "application/json"})


def get_json(url, **kw):
    for attempt in range(3):
        try:
            r = S.get(url, timeout=30, **kw)
            if r.status_code == 429:
                time.sleep(5 * (attempt + 1))
                continue
            r.raise_for_status()
            return r.json()
        except requests.RequestException:
            if attempt == 2:
                raise
            time.sleep(2)


def fetch_greenhouse(info, cap):
    d = get_json(f"https://boards-api.greenhouse.io/v1/boards/{info['token']}/jobs")
    return [{"title": j.get("title", ""), "location": (j.get("location") or {}).get("name", ""),
             "country": None} for j in d.get("jobs", [])][:cap]


def fetch_lever(info, cap):
    d = get_json(f"https://api.lever.co/v0/postings/{info['token']}?mode=json")
    out = []
    for j in d:
        cat = j.get("categories") or {}
        locs = cat.get("allLocations") or [cat.get("location", "")]
        out.append({"title": j.get("text", ""), "location": "; ".join(filter(None, locs)),
                    "country": j.get("country")})
    return out[:cap]


def fetch_ashby(info, cap):
    d = get_json(f"https://api.ashbyhq.com/posting-api/job-board/{info['token']}")
    out = []
    for j in d.get("jobs", []):
        if j.get("isListed") is False:
            continue
        addr = ((j.get("address") or {}).get("postalAddress") or {})
        out.append({"title": j.get("title", ""), "location": j.get("location", ""),
                    "country": addr.get("addressCountry")})
    return out[:cap]


def fetch_smartrecruiters(info, cap):
    out, offset = [], 0
    while len(out) < cap:
        d = get_json(f"https://api.smartrecruiters.com/v1/companies/{info['token']}/postings",
                     params={"limit": 100, "offset": offset, "country": "us"})
        items = d.get("content", [])
        for j in items:
            loc = j.get("location") or {}
            out.append({"title": j.get("name", ""),
                        "location": ", ".join(filter(None, [loc.get("city"), loc.get("region")])),
                        "country": loc.get("country")})
        offset += 100
        if not items or offset >= d.get("totalFound", 0):
            break
    return out[:cap]


def _workday_us_facet(facets):
    """Find the Workday facet value for 'United States' so we can filter server-side."""
    def walk(fs):
        for f in fs or []:
            param = f.get("facetParameter", "")
            for v in f.get("values", []) or []:
                if "values" in v:  # nested facet group
                    yield from walk([v])
                desc = (v.get("descriptor") or "").lower()
                if desc in ("united states of america", "united states", "usa", "us"):
                    yield param, v.get("id")
    for param, vid in walk(facets):
        if "country" in param.lower() or "location" in param.lower():
            return param, vid
    return None, None


def fetch_workday(info, cap):
    base = f"https://{info['tenant']}.{info['wd']}.myworkdayjobs.com"
    api = f"{base}/wday/cxs/{info['tenant']}/{info['site']}/jobs"
    hdr = {"Content-Type": "application/json", "Accept": "application/json"}

    def page(offset, facets):
        body = {"appliedFacets": facets, "limit": 20, "offset": offset, "searchText": ""}
        for attempt in range(3):
            try:
                r = S.post(api, json=body, headers=hdr, timeout=30)
                if r.status_code == 429:
                    time.sleep(5 * (attempt + 1)); continue
                r.raise_for_status()
                return r.json()
            except requests.RequestException:
                if attempt == 2:
                    raise
                time.sleep(2)

    first = page(0, {})
    param, vid = _workday_us_facet(first.get("facets"))
    facets = {param: [vid]} if param else {}
    us_filtered = bool(param)
    d = page(0, facets) if param else first
    total = d.get("total", 0)
    out, offset = [], 0
    while True:
        for j in d.get("jobPostings", []):
            out.append({"title": j.get("title", ""), "location": j.get("locationsText", ""),
                        "country": "US" if us_filtered else None})
        offset += 20
        if offset >= min(total, cap) or not d.get("jobPostings"):
            break
        time.sleep(0.3)
        d = page(offset, facets)
    return out[:cap]


FETCHERS = {"greenhouse": fetch_greenhouse, "lever": fetch_lever, "ashby": fetch_ashby,
            "smartrecruiters": fetch_smartrecruiters, "workday": fetch_workday}


def fetch_company(company, ats, info, cap):
    safe = re.sub(r"[^\w-]+", "_", company)
    path = JOBS_DIR / f"{safe}.json"
    if path.exists():
        return company, json.loads(path.read_text()), None
    try:
        jobs = FETCHERS[ats](info, cap)
        path.write_text(json.dumps(jobs))
        return company, jobs, None
    except Exception as e:
        return company, None, str(e)[:200]


# --------------------------------------------------------------------------
# 3. US filter + role classification
# --------------------------------------------------------------------------

STATE_NAMES = [s.replace("_", " ") for s in (
    "Alabama Alaska Arizona Arkansas California Colorado Connecticut Delaware Florida Georgia "
    "Hawaii Idaho Illinois Indiana Iowa Kansas Kentucky Louisiana Maine Maryland Massachusetts "
    "Michigan Minnesota Mississippi Missouri Montana Nebraska Nevada New_Hampshire New_Jersey "
    "New_Mexico New_York North_Carolina North_Dakota Ohio Oklahoma Oregon Pennsylvania "
    "Rhode_Island South_Carolina South_Dakota Tennessee Texas Utah Vermont Virginia Washington "
    "West_Virginia Wisconsin Wyoming District_of_Columbia").split()]
ABBR = ("AL AK AZ AR CA CO CT DE FL GA HI ID IL IN IA KS KY LA ME MD MA MI MN MS MO MT NE NV NH "
        "NJ NM NY NC ND OH OK OR PA RI SC SD TN TX UT VT VA WA WV WI WY DC").split()
US_RE = re.compile(
    r"\b(united states|u\.s\.a?\.?|usa)\b|\bUS\b|\bUS[-_ ]|"
    r"(?:,|-)\s*(?:" + "|".join(ABBR) + r")\b|\b(?:" + "|".join(STATE_NAMES) + r")\b")
NON_US_RE = re.compile(
    r"\b(canada|ontario|quebec|british columbia|alberta|toronto|vancouver|montreal|mexico|"
    r"india|bangalore|bengaluru|hyderabad|pune|chennai|united kingdom|\buk\b|london|ireland|"
    r"germany|france|poland|spain|italy|netherlands|brazil|argentina|china|japan|singapore|"
    r"philippines|australia|costa rica|colombia|israel|romania|czech|hungary|portugal|"
    r"sweden|switzerland|belgium|malaysia|vietnam|korea|taiwan|uae|dubai|south africa)\b", re.I)
US_COUNTRY = {"us", "usa", "united states", "united states of america", "u.s."}


def is_us(job):
    c = (job.get("country") or "").strip().lower()
    if c:
        return c in US_COUNTRY
    loc = job.get("location") or ""
    if not loc:
        return False
    us = bool(US_RE.search(loc))
    if NON_US_RE.search(loc) and not re.search(r"united states|\bUSA\b", loc, re.I):
        # multi-location strings like "Austin, TX; Toronto, ON" still count as US
        return us and ";" in loc
    return us


# Order matters: more specific families first. Edit freely.
ROLE_RULES = [
    ("Machine Learning / AI Engineer", r"machine learning|\bml\b|\bai\b|artificial intelligence|deep learning|\bllm\b|genai|generative ai|computer vision|\bnlp\b"),
    ("Data Scientist", r"data scien|applied scientist|research scientist"),
    ("Data Engineer", r"data engineer|etl|data pipeline|big data|data platform engineer|analytics engineer"),
    ("Data Analyst / BI", r"data analyst|business intelligence|\bbi (developer|analyst|engineer)|reporting analyst|analytics (analyst|specialist|manager|consultant)|tableau|power ?bi|insights analyst|data analytics"),
    ("Cybersecurity", r"secur|cyber|soc analyst|penetration|infosec|threat|identity and access|\biam\b|vulnerability"),
    ("DevOps / SRE / Cloud", r"devops|site reliability|\bsre\b|cloud (engineer|architect|ops)|platform engineer|infrastructure engineer|kubernetes|\baws\b|azure|\bgcp\b"),
    ("Full Stack Developer", r"full[- ]?stack"),
    ("Frontend Developer", r"front[- ]?end|\bui (developer|engineer)|react developer|angular developer|web developer"),
    ("Backend Developer", r"back[- ]?end|\bjava developer|\.net developer|python developer|api developer"),
    ("Mobile Developer", r"\bios\b|android|mobile (developer|engineer)"),
    ("QA / Test Engineer", r"\bqa\b|quality assurance|\btest(ing)? (engineer|analyst|automation)|\bsdet\b|software tester"),
    ("Embedded / Firmware Engineer", r"embedded|firmware"),
    ("Software Engineer / Developer", r"software|developer|programmer|\bswe\b|application engineer|systems engineer.*software|engineer,? (software|platform)|salesforce|\bsap\b|servicenow|mainframe|cobol"),
    ("Database Administrator", r"database|\bdba\b|sql server"),
    ("Network Engineer", r"network (engineer|administrator|architect|analyst)|\bnoc\b"),
    ("Systems Administrator", r"system(s)? administrator|sysadmin|systems engineer|linux admin|windows admin"),
    ("IT Support / Help Desk", r"help ?desk|desktop support|it support|technical support|service desk|it technician|it specialist|it analyst"),
    ("Solutions / Enterprise Architect", r"architect"),
    ("Product Manager / Owner", r"product manag|product owner|product lead"),
    ("Project / Program Manager", r"project manag|program manag|scrum master|\bpmo\b|delivery manager"),
    ("Business Analyst", r"business analyst|business systems analyst|systems analyst|functional analyst"),
    ("UX / UI / Product Designer", r"\bux\b|\bui\b|user experience|product designer|interaction designer|visual designer|graphic designer|designer"),
    ("Controls / Automation Engineer", r"controls? engineer|automation engineer|\bplc\b|scada|instrumentation|\bbms\b|epms"),
    ("Electrical Engineer", r"electrical engineer|power engineer|electrical design"),
    ("Mechanical Engineer", r"mechanical engineer|mechanical design"),
    ("Manufacturing / Process / Quality Engineer", r"manufacturing engineer|process engineer|quality engineer|industrial engineer|reliability engineer|production engineer"),
    ("Other Engineer", r"engineer"),
    ("Nursing", r"\brn\b|nurse|nursing|\blpn\b|cna\b|nurse practitioner"),
    ("Pharmacy", r"pharmac"),
    ("Clinical / Healthcare", r"physician|clinical|therapist|medical|patient|technologist|radiolog|surgical|dental|phlebotom|caregiver|behavioral health|social worker|counselor"),
    ("Accounting / Finance", r"account|financ|audit|tax|controller|treasury|payroll|underwrit|actuar|credit|risk analyst|fp&a|bookkeep"),
    ("Sales / Account Management", r"sales|account (executive|manager)|business development|relationship manager|territory|inside sales"),
    ("Marketing / Communications", r"marketing|brand|communications|content|social media|seo|public relations|copywriter"),
    ("Customer Service / Call Center", r"customer (service|care|support|success)|call center|contact center|client service"),
    ("Human Resources / Recruiting", r"human resources|\bhr\b|recruit|talent|people partner|benefits|compensation"),
    ("Legal / Compliance", r"legal|attorney|counsel|paralegal|compliance|regulatory"),
    ("Warehouse / Logistics / Supply Chain", r"warehouse|logistics|supply chain|distribution|forklift|material handler|inventory|fulfillment|shipping|receiving|procurement|buyer|purchasing"),
    ("Driver / Transportation", r"driver|cdl|truck|delivery|courier|transport"),
    ("Retail / Store Associate", r"store|retail|cashier|sales associate|merchandis|team member|crew member|associate"),
    ("Technician / Maintenance / Skilled Trades", r"technician|mechanic|maintenance|electrician|hvac|plumb|welder|machinist|installer|lineman|operator"),
    ("Manufacturing / Production Worker", r"production|assembler|manufacturing|packag"),
    ("Food Service / Hospitality", r"cook|chef|food|restaurant|barista|server|kitchen|hospitality|housekeep"),
    ("Administrative / Office", r"administrative|assistant|receptionist|office|clerk|coordinator|data entry|secretary"),
    ("Operations / General Management", r"operations|general manager|store manager|director|manager|supervisor|lead"),
    ("Consultant", r"consult|advisor"),
    ("Analyst (General)", r"analyst"),
    ("Intern / Early Career", r"intern|co-op|apprentice|graduate|new grad"),
]
ROLE_RES = [(name, re.compile(p, re.I)) for name, p in ROLE_RULES]


def classify(title):
    t = title or ""
    for name, rx in ROLE_RES:
        if rx.search(t):
            return name
    return "Unclassified"


# --------------------------------------------------------------------------
# Report
# --------------------------------------------------------------------------

def write_report(rows, detected, results, errors, out_path):
    job_rows, coverage = [], []
    for r in rows:
        c = r["company"]
        d = detected.get(c) or {}
        ats = d.get("ats")
        jobs = results.get(c)
        us = [j for j in (jobs or []) if is_us(j)]
        for j in us:
            job_rows.append({"Company": c, "Industry": r["industry"], "Title": j["title"],
                             "Location": j["location"], "Role Family": classify(j["title"])})
        if jobs is not None:
            status = "Fetched"
        elif c in errors:
            status = f"Fetch failed: {errors[c]}"
        elif ats:
            status = "Detected, not fetched"
        else:
            status = "Job system not detected / not supported"
        coverage.append({"Company": c, "Industry": r["industry"], "Career URL": r["url"],
                         "Job System": ats or "", "Status": status,
                         "All Jobs Pulled": len(jobs or []), "US Jobs": len(us)})

    jobs_df = pd.DataFrame(job_rows, columns=["Company", "Industry", "Title", "Location", "Role Family"])
    cov_df = pd.DataFrame(coverage)
    covered = int((cov_df["Status"] == "Fetched").sum())

    rank = (jobs_df.groupby("Role Family")
            .agg(Companies_Hiring=("Company", "nunique"), US_Openings=("Title", "size"))
            .reset_index()
            .sort_values(["Companies_Hiring", "US_Openings"], ascending=False))
    rank.insert(0, "Rank", range(1, len(rank) + 1))
    rank["% of Covered Companies"] = (rank["Companies_Hiring"] / max(covered, 1)).round(3)
    rank = rank.rename(columns={"Companies_Hiring": "Companies Hiring",
                                "US_Openings": "Total US Openings"})

    by_industry = (jobs_df.pivot_table(index="Role Family", columns="Industry",
                                       values="Company", aggfunc="nunique", fill_value=0)
                   if not jobs_df.empty else pd.DataFrame())
    unclassified = (jobs_df[jobs_df["Role Family"] == "Unclassified"]["Title"]
                    .value_counts().head(300).rename_axis("Title").reset_index(name="Count"))

    summary = pd.DataFrame({
        "Metric": ["Companies in sheet", "Companies with jobs fetched",
                   "Total US openings counted", "Role families found",
                   "Ranking method"],
        "Value": [len(rows), covered, len(jobs_df), len(rank),
                  "Sorted by number of companies hiring, then total openings"]})

    with pd.ExcelWriter(out_path, engine="openpyxl") as xw:
        summary.to_excel(xw, sheet_name="Summary", index=False)
        rank.to_excel(xw, sheet_name="Role Ranking", index=False)
        if not by_industry.empty:
            by_industry.to_excel(xw, sheet_name="Roles by Industry")
        cov_df.to_excel(xw, sheet_name="Company Coverage", index=False)
        unclassified.to_excel(xw, sheet_name="Unclassified Titles", index=False)
        jobs_df.to_excel(xw, sheet_name="All US Jobs", index=False)

    from openpyxl import load_workbook
    from openpyxl.styles import Font, PatternFill
    wb = load_workbook(out_path)
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                cell.font = Font(name="Arial", bold=cell.row == 1)
                if cell.row == 1:
                    cell.fill = PatternFill("solid", fgColor="DDE7F3")
        for col in ws.columns:
            width = max(len(str(c.value or "")) for c in list(col)[:200])
            ws.column_dimensions[col[0].column_letter].width = min(max(10, width + 2), 60)
        ws.freeze_panes = "A2"
    ws = wb["Role Ranking"]
    for cell in ws["E"][1:]:
        cell.number_format = "0.0%"
    wb.save(out_path)
    return rank, covered, len(jobs_df)


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx")
    ap.add_argument("--out", default="role_demand_report.xlsx")
    ap.add_argument("--concurrency", type=int, default=5, help="browser tabs at once")
    ap.add_argument("--max-jobs", type=int, default=5000, help="cap per company")
    ap.add_argument("--skip-detect", action="store_true")
    ap.add_argument("--limit", type=int, default=0, help="only first N companies (testing)")
    a = ap.parse_args()

    CACHE_DIR.mkdir(exist_ok=True)
    JOBS_DIR.mkdir(exist_ok=True)

    df = pd.read_excel(a.xlsx)
    rows = []
    for _, r in df.iterrows():
        override = str(r.get("ATS Override", "") or "").strip()
        rows.append({"company": str(r["Company"]).strip(),
                     "industry": str(r.get("Industry", "")).strip(),
                     "url": str(r["Career Page URL"]).strip(),
                     "override": override if override.startswith("http") else ""})
    if a.limit:
        rows = rows[:a.limit]

    # Step 1: detect
    detected = json.loads(DETECT_CACHE.read_text()) if DETECT_CACHE.exists() else {}
    if not a.skip_detect:
        detected = asyncio.run(detect_all(rows, a.concurrency))
    for r in rows:  # direct URLs and overrides
        for src in (r["override"], r["url"]):
            if src:
                ats, info = find_ats(src)
                if ats:
                    detected[r["company"]] = {"company": r["company"], "ats": ats, "info": info}
                    break
    DETECT_CACHE.write_text(json.dumps(detected, indent=1))

    # Step 2: fetch
    todo = [(r["company"], detected[r["company"]]["ats"], detected[r["company"]]["info"])
            for r in rows if (detected.get(r["company"]) or {}).get("ats")]
    print(f"\nFetching jobs for {len(todo)} companies with a supported job system...")
    results, errors = {}, {}
    with ThreadPoolExecutor(max_workers=8) as ex:
        futs = [ex.submit(fetch_company, c, ats, info, a.max_jobs) for c, ats, info in todo]
        for f in as_completed(futs):
            c, jobs, err = f.result()
            if jobs is not None:
                results[c] = jobs
                print(f"  {len(jobs):6d} jobs  {c}", flush=True)
            else:
                errors[c] = err
                print(f"  FAILED       {c}: {err}", flush=True)

    # Step 3: report
    rank, covered, n = write_report(rows, detected, results, errors, a.out)
    print(f"\nDone. {covered}/{len(rows)} companies covered, {n} US openings counted.")
    print(f"Report: {Path(a.out).resolve()}\n")
    print(rank.head(20).to_string(index=False))


if __name__ == "__main__":
    sys.exit(main())
