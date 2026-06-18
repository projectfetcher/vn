import os
import re
import sys
import csv
import time
import math
import logging
from datetime import datetime

import requests
from bs4 import BeautifulSoup
import pandas as pd

# ----------------------------------------------------------------------------
#  CONFIG
# ----------------------------------------------------------------------------
BASE        = "https://careerone.vn"
JOBS_URL    = f"{BASE}/jobs"
MAX_PAGES   = int(os.environ.get("CO_MAX_PAGES", "6"))   # 0 => crawl until empty
DELAY       = float(os.environ.get("CO_DELAY", "1.0"))
OUTPUT_CSV  = os.environ.get("CO_OUT_CSV", "careerone_jobs.csv")
OUTPUT_XLSX = os.environ.get("CO_OUT_XLSX", "careerone_jobs.xlsx")
TRACKER     = os.environ.get("CO_TRACKER", "processed_urls.csv")

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.8",
}

OUTPUT_COLUMNS = [
    "Job Title", "Job Type", "Job Qualifications", "Job Experience",
    "Job Location", "Job Field", "Date Posted", "Deadline",
    "Job Description", "Application", "Company URL", "Company Name",
    "Company Logo", "Company Industry", "Company Founded", "Company Type",
    "Company Website", "Company Address", "Company Details",
    "Job URL", "Estimated Deadline", "Salary Range",
]

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("careerone")


# ----------------------------------------------------------------------------
#  SMALL HELPERS
# ----------------------------------------------------------------------------
def s(val) -> str:
    if val is None:
        return ""
    if isinstance(val, float) and math.isnan(val):
        return ""
    return str(val).strip()


_SESSION = requests.Session()
_SESSION.headers.update(HEADERS)


def get_soup(url: str, retries: int = 3, timeout: int = 25):
    """Fetch a URL and return a BeautifulSoup, or None on failure."""
    for attempt in range(1, retries + 1):
        try:
            resp = _SESSION.get(url, timeout=timeout, allow_redirects=True)
            if resp.status_code == 200:
                return BeautifulSoup(resp.text, "html.parser")
            log.warning(f"  HTTP {resp.status_code} for {url}")
        except requests.RequestException as e:
            log.warning(f"  request error ({attempt}/{retries}) {url}: {e}")
        time.sleep(DELAY * attempt)
    return None


# ----------------------------------------------------------------------------
#  CLOUDFLARE EMAIL DECODER
# ----------------------------------------------------------------------------
def cf_decode(cfhex: str) -> str:
    """Decode a Cloudflare-obfuscated email from its hex string."""
    try:
        key = int(cfhex[:2], 16)
        return "".join(
            chr(int(cfhex[i:i + 2], 16) ^ key) for i in range(2, len(cfhex), 2)
        )
    except (ValueError, IndexError):
        return ""


def extract_cf_emails(soup: BeautifulSoup) -> list:
    """Return all decoded emails found anywhere on the page, in document order."""
    emails = []

    # 1) <span|a class="__cf_email__" data-cfemail="...">
    for el in soup.select("[data-cfemail]"):
        dec = cf_decode(el.get("data-cfemail", ""))
        if "@" in dec:
            emails.append(dec)

    # 2) <a href=".../cdn-cgi/l/email-protection#HEX">
    for a in soup.find_all("a", href=True):
        m = re.search(r"/cdn-cgi/l/email-protection#([0-9a-fA-F]+)", a["href"])
        if m:
            dec = cf_decode(m.group(1))
            if "@" in dec:
                emails.append(dec)

    # 3) plain mailto:
    for a in soup.find_all("a", href=True):
        if a["href"].lower().startswith("mailto:"):
            addr = a["href"][7:].split("?")[0].strip()
            if "@" in addr:
                emails.append(addr)

    # de-dup, preserve order
    seen, out = set(), []
    for e in emails:
        el = e.lower()
        if el not in seen:
            seen.add(el)
            out.append(e)
    return out


# ----------------------------------------------------------------------------
#  STANDARDISERS  (qualification / experience / field / type / company / salary)
# ----------------------------------------------------------------------------
QUALIFICATION_TIERS = [
    ("PhD / Doctorate", ["phd", "ph.d", "doctorate", "doctoral"]),
    ("Master's Degree", ["master", "msc", "m.sc", "mba", "mphil", "postgraduate", "master of"]),
    ("Bachelor's Degree", ["bachelor", "bsc", "b.sc", "b.a", "beng", "llb", "degree in",
                            "university degree", "undergraduate", "honours"]),
    ("Diploma", ["diploma", "associate degree", "college degree"]),
    ("Professional Certification", ["acca", "cpa", "cfa", "cima", "pmp", "prince2", "chartered"]),
    ("A-Levels / High School", ["high school", "secondary school", "a-level", "a level"]),
]


def extract_qualification(text: str) -> str:
    low = (text or "").lower()
    for label, kws in QUALIFICATION_TIERS:
        if any(k in low for k in kws):
            return label
    return ""


NO_EXP_KW = ["no experience", "fresh graduate", "freshers", "entry level",
             "entry-level", "0 years", "training provided"]
LESS1_KW  = ["less than 1 year", "under 1 year", "6 months", "less than a year"]


def years_to_band(n: int) -> str:
    if n <= 0:  return "No Experience Required"
    if n <= 2:  return "1 - 2 Years"
    if n <= 5:  return "3 - 5 Years"
    if n <= 10: return "6 - 10 Years"
    return "10+ Years"


def extract_experience(text: str) -> str:
    if not text:
        return ""
    low = text.lower()
    if any(k in low for k in NO_EXP_KW):
        return "No Experience Required"
    if any(k in low for k in LESS1_KW):
        return "Less than 1 Year"
    patterns = [
        r"(\d+)\s*[-\u2013to]+\s*(\d+)\s*\+?\s*years?",
        r"(?:minimum|at\s+least|over|more\s+than)\s+(\d+)\s*\+?\s*years?",
        r"(\d+)\s*\+?\s*years?\s*(?:of\s+)?(?:relevant\s+)?(?:work\s+)?experience",
        r"experience\s*(?:of\s+)?(\d+)\s*years?",
    ]
    for p in patterns:
        m = re.search(p, text, re.IGNORECASE)
        if m:
            raw = int(m.group(1))
            if 0 < raw <= 25:
                return years_to_band(raw)
    return ""


FIELD_KEYWORD_MAP = [
    ("Programme / Project Management",
     ["programme officer", "program officer", "project officer", "project manager",
      "programme manager", "project coordinator", "programme coordinator",
      "project administrator", "project lead", "project assistant"]),
    ("Monitoring & Evaluation",
     ["m&e", "monitoring and evaluation", "mel officer", "evaluation"]),
    ("Finance & Accounting",
     ["finance officer", "finance manager", "accountant", "auditor", "grants officer",
      "financial", "bookkeeper", "treasury"]),
    ("Human Resources",
     ["human resources", "hr officer", "hr manager", "talent", "recruitment"]),
    ("Communications & Advocacy",
     ["communications", "advocacy", "media", "outreach", "public relations"]),
    ("Consultancy",
     ["consultant", "consultancy", "request for proposals", "terms of reference",
      "rfp", "rfa", "service provider"]),
    ("Health & Nutrition",
     ["health", "nutrition", "wash", "medical", "nurse", "clinical"]),
    ("Information Technology",
     ["it officer", "software", "developer", "gis", "qgis", "database", "data analyst"]),
    ("Logistics & Procurement",
     ["logistics", "procurement", "supply chain", "warehouse", "fleet"]),
    ("Research",
     ["researcher", "research", "study", "survey"]),
    ("Administration",
     ["administrator", "admin", "secretary", "receptionist", "operations"]),
]


def infer_job_field(title: str, category: str, desc: str) -> str:
    # Prefer the site's own category if present
    if category:
        return category
    combined = ((title or "") + " " + (desc or "")).lower()
    for label, kws in FIELD_KEYWORD_MAP:
        if any(k in combined for k in kws):
            return label
    return "Development / NGO"


def detect_job_type(site_type: str, title: str, desc: str) -> str:
    combined = (site_type + " " + title + " " + desc).lower()
    if re.search(r"\bvolunteer\b", combined):
        return "Volunteer"
    if re.search(r"\bintern(ship)?\b", combined):
        return "Internship"
    if re.search(r"\bconsultan(t|cy)\b|request for proposals?|terms of reference|\btor\b|\brfp\b|\brfa\b|service provider",
                 combined):
        return "Consultancy / Contract"
    if re.search(r"\bpart[-\s]?time\b", combined):
        return "Part-time"
    if re.search(r"\bcontract\b|\bfixed[-\s]term\b|\btemporary\b", combined):
        return "Consultancy / Contract"
    return "Full-time"


def detect_company_type(text: str) -> str:
    tl = (text or "").lower()
    if re.search(r"\bundp\b|\bunicef\b|\bwfp\b|\bunhcr\b|\bwho\b|\bilo\b|united nations|\biom\b|\bunesco\b|\bunfpa\b", tl):
        return "UN Agency"
    if re.search(r"\bingo\b|international ngo|international non-governmental", tl):
        return "INGO"
    if re.search(r"\bngo\b|non.?governmental|non.?profit|nonprofit|foundation|charity|humanitarian", tl):
        return "NGO / Non-Profit"
    if re.search(r"\bembassy\b|\bgovernment\b|ministry of|\bgovernmental\b", tl):
        return "Government / Embassy"
    if re.search(r"\bsocial enterprise\b", tl):
        return "Social Enterprise"
    if re.search(r"\buniversity\b|research institute|\binstitute\b", tl):
        return "Academic / Research"
    return "NGO / Development"


CURRENCY_PATTERNS = [
    r"USD\s*[\d,]+(?:\s*[-\u2013]\s*[\d,]+)?(?:\s*/\s*\w+)?",
    r"\$\s*[\d,]+(?:\s*[-\u2013]\s*\$?\s*[\d,]+)?(?:\s*/\s*\w+)?",
    r"VND\s*[\d,\.]+(?:\s*[-\u2013]\s*[\d,\.]+)?",
    r"[\d,\.]+\s*(?:VND|\u20ab)\b",
    r"[\d,]+(?:\s*[-\u2013]\s*[\d,]+)?\s*/\s*(?:month|year|day|hour)",
]


def extract_salary(text: str) -> str:
    src = (text or "")[:4000]
    for pat in CURRENCY_PATTERNS:
        m = re.search(pat, src, re.IGNORECASE)
        if m:
            val = m.group(0).strip().rstrip(".,")
            if re.search(r"\d", val):
                return val
    return ""


# ----------------------------------------------------------------------------
#  LISTING PARSER
# ----------------------------------------------------------------------------
def parse_listing(soup: BeautifulSoup) -> list:
    """Return list of dicts from the jobs table: title, url, org, deadline, location."""
    rows = []
    # Find the table whose header contains Deadline + Location
    target = None
    for table in soup.find_all("table"):
        head = " ".join(th.get_text(" ", strip=True).lower()
                        for th in table.find_all(["th", "td"])[:5])
        if "deadline" in head and "location" in head:
            target = table
            break
    if target is None:
        # fallback: first table containing a /jobs/ link
        for table in soup.find_all("table"):
            if table.find("a", href=re.compile(r"/jobs/")):
                target = table
                break
    if target is None:
        return rows

    for tr in target.find_all("tr"):
        cells = tr.find_all(["td", "th"])
        if len(cells) < 4:
            continue
        link = cells[0].find("a", href=re.compile(r"/jobs/[^/]"))
        if not link:
            continue  # header row
        url = link["href"].split("?")[0].rstrip("/") + "/"
        rows.append({
            "title":    link.get_text(" ", strip=True),
            "url":      url,
            "org":      cells[1].get_text(" ", strip=True),
            "deadline": cells[2].get_text(" ", strip=True),
            "location": cells[3].get_text(" ", strip=True),
        })
    return rows


# ----------------------------------------------------------------------------
#  DETAIL PARSER
# ----------------------------------------------------------------------------
DETAIL_LABELS = {
    "organisation": "org",
    "organization": "org",
    "job location": "location",
    "application deadline": "deadline",
    "send application to": "apply_label",
    "job categories": "category",
    "job types": "site_type",
}


def _labelled_fields(text_lines: list) -> dict:
    """Read 'Label:' followed by a value line from the Job Details block."""
    out = {}
    n = len(text_lines)
    for i, line in enumerate(text_lines):
        label = line.strip().rstrip(":").strip().lower()
        if line.strip().endswith(":") and label in DETAIL_LABELS:
            # value = next non-empty line
            for j in range(i + 1, min(i + 4, n)):
                val = text_lines[j].strip().strip("'\u2018\u2019\"")
                if val:
                    out[DETAIL_LABELS[label]] = val
                    break
    return out


def _main_content(soup: BeautifulSoup):
    """Return the soup node most likely to hold the post body."""
    for sel in ["article", "main", ".entry-content", ".post-content",
                "#content .post", "#main", "#content"]:
        node = soup.select_one(sel)
        if node and len(node.get_text(strip=True)) > 200:
            return node
    return soup.body or soup


def _published_date(soup: BeautifulSoup) -> str:
    meta = soup.find("meta", attrs={"property": "article:published_time"})
    if meta and meta.get("content"):
        return meta["content"][:10]
    t = soup.find("time")
    if t and t.get("datetime"):
        return t["datetime"][:10]
    return ""


def _external_website(node, desc: str) -> str:
    blocked = ("careerone.vn", "google.com", "drive.google", "facebook.com",
               "linkedin.com", "twitter.com", "youtube.com", "cdn-cgi",
               "wp-content", "gravatar", "vxtgroup.com")
    for a in node.find_all("a", href=True):
        href = a["href"].strip()
        if href.startswith("http") and not any(b in href for b in blocked):
            return href.rstrip("/")
    return ""


def parse_detail(url: str, listing_row: dict) -> dict:
    soup = get_soup(url)
    if soup is None:
        return None

    # Drop obvious chrome before reading text
    for tag in soup.find_all(["script", "style", "nav", "header", "footer", "aside"]):
        tag.decompose()

    title = ""
    h1 = soup.find(["h1"])
    if h1:
        title = h1.get_text(" ", strip=True)
    title = title or listing_row.get("title", "")

    content = _main_content(soup)
    text = content.get_text("\n", strip=True)
    lines = [l for l in text.split("\n")]

    # Labelled Job Details block
    fields = _labelled_fields(lines)

    # Description = text from after H1 up to "Job Details"
    desc_text = text
    if title and title in desc_text:
        desc_text = desc_text.split(title, 1)[-1]
    desc_text = re.split(r"\n\s*Job Details\s*\n", desc_text, maxsplit=1)[0]
    desc_text = re.sub(r"\n{2,}", "\n\n", desc_text).strip()

    # Apply contact: prefer decoded emails near apply wording, else any cf email
    emails = extract_cf_emails(soup)
    application = ""
    if emails:
        application = emails[0]
    # an application URL (e.g. workday/greenhouse) overrides if clearly an apply link
    for a in content.find_all("a", href=True):
        href = a["href"]
        if re.search(r"workday|greenhouse|lever\.co|bamboohr|smartrecruiters|/apply|careers?\.",
                     href, re.IGNORECASE):
            application = href.rstrip("/")
            break
    if not application:
        application = listing_row.get("apply", "")

    category  = fields.get("category", "")
    site_type = fields.get("site_type", "")
    org       = listing_row.get("org") or fields.get("org", "")
    location  = listing_row.get("location") or fields.get("location", "")
    deadline  = listing_row.get("deadline") or fields.get("deadline", "")

    website = _external_website(content, desc_text)

    record = {
        "Job Title":          title,
        "Job Type":           detect_job_type(site_type, title, desc_text),
        "Job Qualifications": extract_qualification(desc_text),
        "Job Experience":     extract_experience(desc_text),
        "Job Location":       location,
        "Job Field":          infer_job_field(title, category, desc_text),
        "Date Posted":        _published_date(soup),
        "Deadline":           deadline,
        "Job Description":    desc_text[:6000],
        "Application":        application,
        "Company URL":        website,
        "Company Name":       org,
        "Company Logo":       "",   # site does not provide per-employer logos
        "Company Industry":   category or "NGO / Development",
        "Company Founded":    "",
        "Company Type":       detect_company_type(org + " " + desc_text),
        "Company Website":    website,
        "Company Address":    location,
        "Company Details":    "",
        "Job URL":            url,
        "Estimated Deadline": deadline,
        "Salary Range":       extract_salary(desc_text),
    }
    return record


# ----------------------------------------------------------------------------
#  TRACKER
# ----------------------------------------------------------------------------
def load_processed() -> set:
    if not os.path.exists(TRACKER):
        return set()
    try:
        df = pd.read_csv(TRACKER)
        return set(df["Job URL"].fillna("").astype(str))
    except Exception:
        return set()


def save_processed(urls: set):
    pd.DataFrame({"Job URL": sorted(urls)}).to_csv(TRACKER, index=False)


# ----------------------------------------------------------------------------
#  MAIN
# ----------------------------------------------------------------------------
def main():
    print("=" * 60)
    print("  CareerONE.vn Job Scraper")
    print(f"  Started   : {datetime.now():%Y-%m-%d %H:%M:%S}")
    print(f"  Max pages : {MAX_PAGES or 'until empty'}  ·  delay {DELAY}s")
    print("=" * 60)

    processed = load_processed()
    log.info(f"Already processed: {len(processed)} jobs")

    # 1) Crawl listing pages
    listing_rows = []
    page = 1
    while True:
        url = JOBS_URL if page == 1 else f"{JOBS_URL}/?paged={page}"
        log.info(f"[listing] page {page}: {url}")
        soup = get_soup(url)
        if soup is None:
            break
        rows = parse_listing(soup)
        if not rows:
            log.info("  no rows — stopping pagination")
            break
        listing_rows.extend(rows)
        page += 1
        if MAX_PAGES and page > MAX_PAGES:
            break
        time.sleep(DELAY)

    # de-dup listing by URL
    seen, unique_rows = set(), []
    for r in listing_rows:
        if r["url"] not in seen:
            seen.add(r["url"])
            unique_rows.append(r)
    log.info(f"Collected {len(unique_rows)} unique listings")

    # 2) Visit detail pages (skip already processed)
    records = []
    for i, row in enumerate(unique_rows, 1):
        if row["url"] in processed:
            log.info(f"[{i}/{len(unique_rows)}] skip (done): {row['title']}")
            continue
        log.info(f"[{i}/{len(unique_rows)}] {row['title']}")
        rec = parse_detail(row["url"], row)
        if rec:
            records.append(rec)
            processed.add(row["url"])
        time.sleep(DELAY)

    if not records:
        log.info("No new jobs scraped. Nothing to write.")
        return

    # 3) Write outputs (append to existing CSV if present)
    new_df = pd.DataFrame(records, columns=OUTPUT_COLUMNS)
    if os.path.exists(OUTPUT_CSV):
        try:
            old = pd.read_csv(OUTPUT_CSV)
            combined = pd.concat([old, new_df], ignore_index=True)
            combined.drop_duplicates(subset=["Job URL"], keep="last", inplace=True)
        except Exception:
            combined = new_df
    else:
        combined = new_df

    combined.to_csv(OUTPUT_CSV, index=False, encoding="utf-8-sig",
                    quoting=csv.QUOTE_ALL)
    log.info(f"Wrote {OUTPUT_CSV} ({len(combined)} rows total)")

    try:
        with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as writer:
            combined.to_excel(writer, index=False, sheet_name="CareerONE Jobs")
            ws = writer.sheets["CareerONE Jobs"]
            for col in ws.columns:
                width = max((len(str(c.value or "")) for c in col), default=10)
                ws.column_dimensions[col[0].column_letter].width = min(width + 3, 60)
            ws.freeze_panes = "A2"
        log.info(f"Wrote {OUTPUT_XLSX}")
    except Exception as e:
        log.warning(f"XLSX write skipped: {e}")

    save_processed(processed)

    print("=" * 60)
    print(f"  Done. New jobs: {len(records)}  ·  Total rows: {len(combined)}")
    print(f"  Output: {OUTPUT_CSV}, {OUTPUT_XLSX}")
    print(f"  Finished: {datetime.now():%Y-%m-%d %H:%M:%S}")
    print("=" * 60)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        sys.exit(130)
