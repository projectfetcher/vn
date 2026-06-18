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
MAX_PAGES   = int(os.environ.get("CO_MAX_PAGES", "6"))
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
#  TITLE CLEANER  (FIX #1 — strip [Extension], \xa0, long RFQ/tender titles)
# ----------------------------------------------------------------------------
# Prefixes like "[Extension]" are status tags, not part of the actual title
_TITLE_PREFIX_RE = re.compile(
    r"^\s*\[(?:extension|extended?|re-?advertised?|readvertised?|revised?|updated?|open)\]\s*",
    re.IGNORECASE,
)
# Non-breaking spaces and stray whitespace
_WHITESPACE_RE = re.compile(r"[\xa0\u200b\u200c\u200d\ufeff]+")

# Procurement / tender noise: titles that are full sentences describing an RFQ/tender
# are truncated to a clean short label
_PROCUREMENT_STARTS = re.compile(
    r"^(?:request\s+for\s+(?:proposals?|quotations?|expressions?\s+of\s+interest|tender)|"
    r"invitation\s+to\s+tender|call\s+for\s+(?:proposals?|quotations?)|"
    r"consultancy\s+(?:for|service|to))\b",
    re.IGNORECASE,
)

MAX_TITLE_LEN = 120  # hard cap — titles longer than this are almost certainly noise


def clean_title(raw: str) -> str:
    """Normalise a scraped job title."""
    t = _WHITESPACE_RE.sub(" ", raw or "").strip()
    # Strip [Extension] / [Re-advertised] etc.
    t = _TITLE_PREFIX_RE.sub("", t).strip()
    # Collapse inner whitespace
    t = re.sub(r"\s{2,}", " ", t)
    # Hard-truncate absurdly long titles (RFQ sentence dumps)
    if len(t) > MAX_TITLE_LEN:
        # Try to cut at a sensible boundary (semicolon, em-dash, colon after 40 chars)
        for sep in [";", " – ", " - ", ":"]:
            idx = t.find(sep, 40)
            if 40 < idx < MAX_TITLE_LEN:
                t = t[:idx].strip()
                break
        else:
            t = t[:MAX_TITLE_LEN].rstrip(" ,;:").strip()
    return t


# ----------------------------------------------------------------------------
#  CLOUDFLARE EMAIL DECODER
# ----------------------------------------------------------------------------
def cf_decode(cfhex: str) -> str:
    try:
        key = int(cfhex[:2], 16)
        return "".join(
            chr(int(cfhex[i:i + 2], 16) ^ key) for i in range(2, len(cfhex), 2)
        )
    except (ValueError, IndexError):
        return ""


def extract_cf_emails(soup: BeautifulSoup) -> list:
    emails = []
    for el in soup.select("[data-cfemail]"):
        dec = cf_decode(el.get("data-cfemail", ""))
        if "@" in dec:
            emails.append(dec)
    for a in soup.find_all("a", href=True):
        m = re.search(r"/cdn-cgi/l/email-protection#([0-9a-fA-F]+)", a["href"])
        if m:
            dec = cf_decode(m.group(1))
            if "@" in dec:
                emails.append(dec)
    for a in soup.find_all("a", href=True):
        if a["href"].lower().startswith("mailto:"):
            addr = a["href"][7:].split("?")[0].strip()
            if "@" in addr:
                emails.append(addr)
    seen, out = set(), []
    for e in emails:
        el = e.lower()
        if el not in seen:
            seen.add(el)
            out.append(e)
    return out


# ----------------------------------------------------------------------------
#  STANDARDISERS
# ----------------------------------------------------------------------------
QUALIFICATION_TIERS = [
    ("PhD / Doctorate",          ["phd", "ph.d", "doctorate", "doctoral"]),
    ("Master's Degree",          ["master", "msc", "m.sc", "mba", "mphil",
                                   "postgraduate", "master of"]),
    ("Bachelor's Degree",        ["bachelor", "bsc", "b.sc", "b.a", "beng", "llb",
                                   "degree in", "university degree", "undergraduate",
                                   "honours"]),
    ("Diploma",                  ["diploma", "associate degree", "college degree"]),
    ("Professional Certification", ["acca", "cpa", "cfa", "cima", "pmp",
                                     "prince2", "chartered"]),
    ("A-Levels / High School",   ["high school", "secondary school", "a-level", "a level"]),
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


# FIX #2 — expanded & better-ordered Job Field map
# Rules: match against job TITLE first (more reliable), then fall back to description.
# Keep most-specific first. IT keywords are title-only to avoid false matches in body text.

FIELD_KEYWORD_MAP = [
    # --- match on TITLE only (tuples of (label, title_keywords, desc_keywords)) ---
    # We use a unified list approach: keywords matched against title+desc[:500] combined
    ("Monitoring & Evaluation",
     ["monitoring and evaluation officer", "m&e officer", "mel officer",
      "evaluation officer", "evaluation specialist"]),
    ("Health & Nutrition",
     ["health officer", "health specialist", "health advisor", "nutrition officer",
      "wash officer", "nurse", "epidemiology officer", "surveillance officer",
      "hiv officer", "aids officer", "laboratory advisor", "lab specialist",
      "laboratory specialist", "laboratory officer"]),
    ("Finance & Accounting",
     ["finance officer", "finance manager", "accountant", "auditor",
      "grants officer", "grant officer", "financial analyst", "bookkeeper", "treasurer"]),
    ("Human Resources",
     ["human resource officer", "human resources officer", "hr officer", "hr manager",
      "hr & admin", "hr and admin", "talent acquisition", "recruitment officer"]),
    ("Communications & Advocacy",
     ["communications officer", "communications manager", "communications assistant",
      "advocacy officer", "media officer", "pr officer", "pr and communications",
      "outreach officer", "public relations officer"]),
    ("Research",
     ["research officer", "research assistant", "research specialist",
      "baseline survey", "endline survey", "evaluation consultant",
      "survey consultant"]),
    # IT: only match very explicit IT job titles, not body-text mentions of "database"
    ("Information Technology",
     ["it officer", "software developer", "software engineer", "gis officer",
      "qgis specialist", "data analyst", "digital transformation officer",
      "mis officer", "ict officer", "systems developer"]),
    ("Logistics & Procurement",
     ["logistics officer", "procurement officer", "supply chain officer",
      "request for quotation", "rfq", "invitation to tender",
      "call for quotation", "request for tender",
      "flight ticket", "beverages", "food items procurement",
      "equipment procurement"]),
    ("Administration",
     ["administrative assistant", "admin officer", "admin assistant",
      "secretary", "receptionist"]),
    ("Environment & Natural Resources",
     ["environment officer", "environmental officer", "climate officer",
      "water governance", "fisheries officer", "wildlife officer",
      "agribusiness specialist", "forestry officer", "aquaculture",
      "resilience specialist", "environment management officer",
      "environment management"]),
    ("Education & Training",
     ["teacher", "trainer", "tvet specialist", "training specialist",
      "education officer", "education program assistant",
      "teacher guideline", "patisserie trainer", "hospitality trainer"]),
    ("Social Work & Community",
     ["social worker", "community outreach officer", "community development officer"]),
    ("Programme / Project Management",
     ["programme officer", "program officer", "project officer", "project manager",
      "programme manager", "project coordinator", "programme coordinator",
      "project administrator", "project lead", "project assistant",
      "program coordinator", "program manager", "team leader", "deputy team leader",
      "program assistant", "programme assistant", "program manager",
      "senior program coordinator", "program development"]),
    ("Consultancy",
     ["consultant", "consultancy", "request for proposals",
      "terms of reference", "rfp", "rfa", "service provider",
      "expressions of interest", "technical advisor", "technical consultant"]),
]


def infer_job_field(title: str, category: str, desc: str) -> str:
    """Match title first (high confidence), then title+desc[:600] for broader context."""
    if category and category not in ("NGO / Development", "Development / NGO", ""):
        return category

    title_low = (title or "").lower()
    # Phase 1: title-only match (very precise)
    for label, kws in FIELD_KEYWORD_MAP:
        if any(k in title_low for k in kws):
            return label

    # Phase 2: title + first 600 chars of description (for M&E, Health etc.)
    combined = (title_low + " " + (desc or "")[:600].lower())
    for label, kws in FIELD_KEYWORD_MAP:
        if any(k in combined for k in kws):
            return label

    return "Development / NGO"


def infer_job_field(title: str, category: str, desc: str) -> str:
    if category and category not in ("NGO / Development", "Development / NGO", ""):
        return category
    combined = ((title or "") + " " + (desc or ""))[:3000].lower()
    for label, kws in FIELD_KEYWORD_MAP:
        if any(k in combined for k in kws):
            return label
    return "Development / NGO"


# FIX #3 — tighter Job Type detection (avoid false "Internship" from "intern" in company names)
def detect_job_type(site_type: str, title: str, desc: str) -> str:
    combined = (site_type + " " + title + " " + desc).lower()
    if re.search(r"\bvolunteer\b", combined):
        return "Volunteer"
    # "intern" only as a standalone word / "internship" — not inside other words
    if re.search(r"\binternship\b|\binter[ns]\b(?!\w)", combined):
        return "Internship"
    if re.search(r"\bpart[-\s]?time\b", combined):
        return "Part-time"
    # Procurement / tender documents → Consultancy / Contract
    if re.search(
        r"\bconsultan(?:t|cy)\b|request\s+for\s+(?:proposals?|quotations?|tender)|"
        r"invitation\s+to\s+tender|call\s+for\s+quotations?|terms\s+of\s+reference|"
        r"\btor\b|\brfp\b|\brfa\b|\brfq\b|service\s+provider|expressions?\s+of\s+interest",
        combined,
    ):
        return "Consultancy / Contract"
    if re.search(r"\bcontract\b|\bfixed[-\s]?term\b|\btemporary\b", combined):
        return "Consultancy / Contract"
    return "Full-time"


def detect_company_type(text: str) -> str:
    tl = (text or "").lower()
    if re.search(
        r"\bundp\b|\bunicef\b|\bwfp\b|\bunhcr\b|\bwho\b|\bilo\b|"
        r"united nations|\biom\b|\bunesco\b|\bunfpa\b|\bfao\b|\bunops\b",
        tl,
    ):
        return "UN Agency"
    if re.search(r"\bingo\b|international ngo|international non.?governmental", tl):
        return "INGO"
    if re.search(
        r"\bngo\b|non.?governmental|non.?profit|nonprofit|foundation|"
        r"charity|humanitarian|civil society",
        tl,
    ):
        return "NGO / Non-Profit"
    if re.search(r"\bembassy\b|\bgovernment\b|ministry of|\bgovernmental\b", tl):
        return "Government / Embassy"
    if re.search(r"\bsocial enterprise\b", tl):
        return "Social Enterprise"
    if re.search(r"\buniversity\b|research institute|\binstitute\b", tl):
        return "Academic / Research"
    return "NGO / Development"


# FIX #4 — salary: avoid false matches like "USD5" (budget line) or VND totals
CURRENCY_PATTERNS = [
    # USD range: USD 1,000 - 2,000  or  $1,000 – $2,000 /month
    r"USD\s*[\d,]{4,}(?:\s*[-\u2013]\s*(?:USD\s*)?[\d,]{3,})?(?:\s*/\s*\w+)?",
    r"\$\s*[\d,]{4,}(?:\s*[-\u2013]\s*\$?\s*[\d,]{3,})?(?:\s*/\s*\w+)?",
    # VND amounts (must be large numbers: at least 7 digits = millions of VND)
    r"VND\s*[\d,\.]{7,}(?:\s*[-\u2013]\s*[\d,\.]{6,})?",
    r"[\d,\.]{7,}\s*(?:VND|\u20ab)\b",
    # Per-time-period salary (must have at least 3-digit number)
    r"[\d,]{3,}(?:\s*[-\u2013]\s*[\d,]{3,})?\s*/\s*(?:month|year|day|hour)",
]


def extract_salary(text: str) -> str:
    src = (text or "")[:4000]
    for pat in CURRENCY_PATTERNS:
        m = re.search(pat, src, re.IGNORECASE)
        if m:
            val = m.group(0).strip().rstrip(".,")
            if re.search(r"\d{3,}", val):   # sanity: at least a 3-digit number
                return val
    return ""


# FIX #5 — Date Posted: also look for WordPress-style date meta and visible date text
_DATE_META_ATTRS = [
    {"property": "article:published_time"},
    {"name": "date"},
    {"itemprop": "datePublished"},
]
_DATE_TEXT_RE = re.compile(
    r"(?:posted|published|date)\s*[:\-]?\s*"
    r"(\d{1,2}[\s/\-]\w+[\s/\-]\d{2,4}|\w+\s+\d{1,2},?\s+\d{4}|\d{4}-\d{2}-\d{2})",
    re.IGNORECASE,
)


def _published_date(soup: BeautifulSoup) -> str:
    # 1) meta tags
    for attrs in _DATE_META_ATTRS:
        meta = soup.find("meta", attrs=attrs)
        if meta and meta.get("content"):
            raw = meta["content"][:10]
            if re.match(r"\d{4}-\d{2}-\d{2}", raw):
                return raw
    # 2) <time> element
    t = soup.find("time")
    if t and t.get("datetime"):
        return t["datetime"][:10]
    # 3) visible "Posted: DD Month YYYY" text
    m = _DATE_TEXT_RE.search(soup.get_text(" ", strip=True)[:2000])
    if m:
        try:
            for fmt in ("%d %B %Y", "%B %d, %Y", "%d/%m/%Y", "%Y-%m-%d"):
                try:
                    return datetime.strptime(m.group(1).strip(), fmt).strftime("%Y-%m-%d")
                except ValueError:
                    pass
        except Exception:
            pass
    return ""


# ----------------------------------------------------------------------------
#  LISTING + DETAIL PARSERS  (unchanged except title now goes through clean_title)
# ----------------------------------------------------------------------------
def parse_listing(soup: BeautifulSoup) -> list:
    rows = []
    target = None
    for table in soup.find_all("table"):
        head = " ".join(th.get_text(" ", strip=True).lower()
                        for th in table.find_all(["th", "td"])[:5])
        if "deadline" in head and "location" in head:
            target = table
            break
    if target is None:
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
            continue
        url = link["href"].split("?")[0].rstrip("/") + "/"
        rows.append({
            "title":    clean_title(link.get_text(" ", strip=True)),   # ← cleaned here
            "url":      url,
            "org":      cells[1].get_text(" ", strip=True),
            "deadline": cells[2].get_text(" ", strip=True),
            "location": cells[3].get_text(" ", strip=True),
        })
    return rows


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
    out = {}
    n = len(text_lines)
    for i, line in enumerate(text_lines):
        label = line.strip().rstrip(":").strip().lower()
        if line.strip().endswith(":") and label in DETAIL_LABELS:
            for j in range(i + 1, min(i + 4, n)):
                val = text_lines[j].strip().strip("'\u2018\u2019\"")
                if val:
                    out[DETAIL_LABELS[label]] = val
                    break
    return out


def _main_content(soup: BeautifulSoup):
    for sel in ["article", "main", ".entry-content", ".post-content",
                "#content .post", "#main", "#content"]:
        node = soup.select_one(sel)
        if node and len(node.get_text(strip=True)) > 200:
            return node
    return soup.body or soup


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

    for tag in soup.find_all(["script", "style", "nav", "header", "footer", "aside"]):
        tag.decompose()

    title = ""
    h1 = soup.find(["h1"])
    if h1:
        title = clean_title(h1.get_text(" ", strip=True))   # ← cleaned here too
    title = title or listing_row.get("title", "")

    content = _main_content(soup)
    text = content.get_text("\n", strip=True)
    lines = [l for l in text.split("\n")]

    fields = _labelled_fields(lines)

    desc_text = text
    if title and title in desc_text:
        desc_text = desc_text.split(title, 1)[-1]
    desc_text = re.split(r"\n\s*Job Details\s*\n", desc_text, maxsplit=1)[0]
    desc_text = re.sub(r"\n{2,}", "\n\n", desc_text).strip()

    emails = extract_cf_emails(soup)
    application = ""
    if emails:
        application = emails[0]
    for a in content.find_all("a", href=True):
        href = a["href"]
        if re.search(
            r"workday|greenhouse|lever\.co|bamboohr|smartrecruiters|/apply|careers?\.",
            href, re.IGNORECASE,
        ):
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
        "Company Logo":       "",
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
#  CSV RE-CLEANER  — apply all fixes to an existing CSV without re-scraping
# ----------------------------------------------------------------------------
def reclean_existing_csv(path: str) -> pd.DataFrame:
    """Read an existing careerone CSV and apply all post-processing fixes in-place."""
    df = pd.read_csv(path)
    log.info(f"Re-cleaning {len(df)} rows from {path}")

    for idx, row in df.iterrows():
        title = clean_title(s(row.get("Job Title", "")))
        desc  = s(row.get("Job Description", ""))
        cat   = s(row.get("Company Industry", ""))
        org   = s(row.get("Company Name", ""))

        df.at[idx, "Job Title"]          = title
        df.at[idx, "Job Type"]           = detect_job_type("", title, desc)
        df.at[idx, "Job Qualifications"] = extract_qualification(desc)
        df.at[idx, "Job Experience"]     = extract_experience(desc)
        df.at[idx, "Job Field"]          = infer_job_field(title, cat, desc)
        df.at[idx, "Company Type"]       = detect_company_type(org + " " + desc)
        df.at[idx, "Salary Range"]       = extract_salary(desc)

    return df


# ----------------------------------------------------------------------------
#  MAIN
# ----------------------------------------------------------------------------
def main():
    # If a CSV already exists and we're just re-cleaning, do that first
    if "--reclean" in sys.argv and os.path.exists(OUTPUT_CSV):
        df = reclean_existing_csv(OUTPUT_CSV)
        df.to_csv(OUTPUT_CSV, index=False, encoding="utf-8-sig", quoting=csv.QUOTE_ALL)
        log.info(f"Re-cleaned CSV saved → {OUTPUT_CSV}")
        return

    print("=" * 60)
    print("  CareerONE.vn Job Scraper")
    print(f"  Started   : {datetime.now():%Y-%m-%d %H:%M:%S}")
    print(f"  Max pages : {MAX_PAGES or 'until empty'}  ·  delay {DELAY}s")
    print("=" * 60)

    processed = load_processed()
    log.info(f"Already processed: {len(processed)} jobs")

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

    seen, unique_rows = set(), []
    for r in listing_rows:
        if r["url"] not in seen:
            seen.add(r["url"])
            unique_rows.append(r)
    log.info(f"Collected {len(unique_rows)} unique listings")

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

    combined.to_csv(OUTPUT_CSV, index=False, encoding="utf-8-sig", quoting=csv.QUOTE_ALL)
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
