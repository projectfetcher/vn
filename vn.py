#!/usr/bin/env python3
"""
careerone_scraper.py — CareerONE.vn job scraper with integrated data cleaning,
Mistral paraphrasing, column mapping, and WordPress posting.

Pipeline:
  1. Crawl listing pages → collect job URLs
  2. Parse each detail page → structured record
  3. Map columns → standard schema
  4. Paraphrase title / description / company via Mistral
  5. Post to WordPress (WP Job Manager) with logo upload & taxonomy terms
  6. Track processed jobs in a local CSV tracker
  7. Save raw scraped data to CSV + XLSX
"""

import os
import re
import sys
import csv
import time
import math
import base64
import hashlib
import logging
from datetime import date, datetime

import requests
from bs4 import BeautifulSoup
import pandas as pd

# ── Optional: sentence-transformers for similarity scoring ────────────────
try:
    from sentence_transformers import SentenceTransformer, util as st_util
    _ST_AVAILABLE = True
except ImportError:
    _ST_AVAILABLE = False

# ── Optional: language-tool for grammar correction ────────────────────────
try:
    import language_tool_python
    _LT_AVAILABLE = True
except ImportError:
    _LT_AVAILABLE = False

# ════════════════════════════════════════════════════════════════════════════
# CONFIG  — all tuneable via environment variables
# ════════════════════════════════════════════════════════════════════════════
BASE        = "https://careerone.vn"
JOBS_URL    = f"{BASE}/jobs"
MAX_PAGES   = int(os.environ.get("CO_MAX_PAGES", "6"))    # 0 = until empty
DELAY       = float(os.environ.get("CO_DELAY", "1.0"))
OUTPUT_CSV  = os.environ.get("CO_OUT_CSV",  "careerone_jobs.csv")
OUTPUT_XLSX = os.environ.get("CO_OUT_XLSX", "careerone_jobs.xlsx")
TRACKER     = os.environ.get("CO_TRACKER",  "processed_urls.csv")
MAX_TITLE   = int(os.environ.get("CO_MAX_TITLE", "80"))

# ── Mistral ───────────────────────────────────────────────────────────────
MISTRAL_API_KEY = os.environ.get("MISTRAL_API_KEY", "")
MISTRAL_MODEL   = os.environ.get("MISTRAL_MODEL", "mistral-small-latest")
MISTRAL_URL     = "https://api.mistral.ai/v1/chat/completions"

# ── WordPress ─────────────────────────────────────────────────────────────
_WP_BASE        = os.environ.get("WP_BASE_URL", "").rstrip("/")
WP_BASE         = _WP_BASE
WP_URL          = f"{WP_BASE}/job-listings"
WP_COMPANY_URL  = f"{WP_BASE}/companies"
WP_MEDIA_URL    = f"{WP_BASE}/media"
WP_USERNAME     = os.environ.get("WP_USERNAME", "")
WP_APP_PASSWORD = os.environ.get("WP_APP_PASSWORD", "")

# ── Feature flags ─────────────────────────────────────────────────────────
ENABLE_PARAPHRASE = os.environ.get("CO_PARAPHRASE", "true").lower() == "true"
ENABLE_WP_POST    = os.environ.get("CO_WP_POST",    "true").lower() == "true"

# ── Tracker file ──────────────────────────────────────────────────────────
PROCESSED_IDS_FILE = os.environ.get("CO_IDS_FILE", "careerone_processed_ids.csv")

JOB_TYPE_MAPPING = {
    "full-time": "full-time", "full time": "full-time", "fulltime": "full-time",
    "part-time": "part-time", "part time": "part-time", "parttime": "part-time",
    "contract": "contract",   "contractor": "contract", "contracting": "contract",
    "temporary": "temporary", "temp": "temporary",
    "freelance": "freelance",
    "internship": "internship", "intern": "internship",
    "volunteer": "volunteer",
    "consultancy / contract": "contract",
}

APPSCRIPT_COLUMNS = [
    "Job Title", "Job Type", "Job Qualifications", "Job Experience",
    "Job Location", "Job Field", "Date Posted", "Deadline",
    "Job Description", "Application", "Company URL", "Company Name",
    "Company Logo", "Company Industry", "Company Founded", "Company Type",
    "Company Website", "Company Address", "Company Details",
    "Job URL", "Estimated Deadline", "Salary Range",
]

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.8",
}

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("careerone")

# ════════════════════════════════════════════════════════════════════════════
# LAZY-LOADED MODELS
# ════════════════════════════════════════════════════════════════════════════
_similarity_model = None
_grammar_tool     = None


def _get_similarity_model():
    global _similarity_model
    if _similarity_model is None and _ST_AVAILABLE:
        _similarity_model = SentenceTransformer("all-MiniLM-L6-v2", device="cpu")
    return _similarity_model


def _get_grammar_tool():
    global _grammar_tool
    if _grammar_tool is None and _LT_AVAILABLE:
        try:
            _grammar_tool = language_tool_python.LanguageTool(
                "en-US", remote_server="https://api.languagetool.org"
            )
        except Exception:
            pass
    return _grammar_tool


# ════════════════════════════════════════════════════════════════════════════
# SMALL HELPERS
# ════════════════════════════════════════════════════════════════════════════
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


def is_ats_or_bad(url: str) -> bool:
    bad = ("sr-company-attachments", "cdn-cgi", "wp-content", "gravatar",
           "facebook.com", "linkedin.com", "twitter.com", "youtube.com")
    return not url or any(b in url for b in bad)


# ════════════════════════════════════════════════════════════════════════════
# TEXT NORMALISATION
# ════════════════════════════════════════════════════════════════════════════
_WS_MAP = {
    "\xa0": " ", "\u2007": " ", "\u202f": " ", "\u2009": " ", "\u2002": " ",
    "\u2003": " ", "\t": " ", "\u200b": "", "\u200c": "", "\u200d": "",
    "\ufeff": "", "\u2028": " ", "\u2029": " ",
}
_QUOTES = {"\u201c": '"', "\u201d": '"', "\u2018": "'", "\u2019": "'"}

_MOJIBAKE = [
    ("Â", ""), ("â€™", "'"), ("â€œ", '"'), ("â€\x9d", '"'), ("â€", '"'),
    ("â€¢", "•"), ("â„¢", "™"), ("\u00a0", " "), ("\u200b", ""), ("\ufeff", ""),
]


def _fix_mojibake(text: str) -> str:
    for pattern, replacement in _MOJIBAKE:
        text = text.replace(pattern, replacement)
    text = re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]", "", text)
    return text


def normalize_text(val) -> str:
    """Flatten invisible chars, normalise quotes, collapse whitespace."""
    if val is None:
        return ""
    t = str(val)
    for bad, good in _WS_MAP.items():
        t = t.replace(bad, good)
    for bad, good in _QUOTES.items():
        t = t.replace(bad, good)
    t = _fix_mojibake(t)
    t = re.sub(r"[ \t]+", " ", t)
    t = re.sub(r" *\n *", "\n", t)
    t = re.sub(r"\n{3,}", "\n\n", t)
    return t.strip()


def sanitize_text(text, is_url=False, is_email=False) -> str:
    if not isinstance(text, str):
        text = str(text) if pd.notna(text) else ""
    text = text.strip()
    if text in ("nan", "None", "NaN", "", "N/A", "n/a", "NA", "na"):
        return ""
    text = _fix_mojibake(text)
    if is_url or is_email:
        return re.sub(r"[ \t\r\n\f\v]+", " ", text).strip()
    text = re.sub(r"#+\s*", "", text)
    text = re.sub(r"\*\*", "", text)
    text = re.sub(r"[^\x20-\x7E\n\u00C0-\u017F\u2013\u2014\u2018-\u201D\u2022]", "", text)
    text = re.sub(r"[ \t]+", " ", text)
    return re.sub(r"\n{3,}", "\n\n", text).strip()


# ════════════════════════════════════════════════════════════════════════════
# TITLE CLEANING
# ════════════════════════════════════════════════════════════════════════════
_STOP_TAIL = re.compile(
    r"\s+(?:and|for|to|of|the|on|in|with|a|an|&|that|which|from|by|at|or)$",
    re.IGNORECASE,
)
_BOILER = re.compile(
    r"^(?:request\s+for\s+(?:proposals?|quotations?|expressions?\s+of\s+interest|"
    r"applications?|information)|invitation\s+to\s+tender|call\s+for\s+(?:proposals?|"
    r"applications?|expressions?\s+of\s+interest|nominations?)|terms\s+of\s+reference|"
    r"consultancy\s+service(?:\s+provider)?|vacancy\s+announcement|job\s+announcement|"
    r"notice\s+of\s+vacancy|rfq|rfp|eoi|tor)\b",
    re.IGNORECASE,
)
_EXTENSION_TAG = re.compile(r"^\s*\[Extension\]\s*", re.IGNORECASE)
_QUOTED_SCOPE  = re.compile(r"[\"\u201c\u201d]([^\"\u201c\u201d]{10,})[\"\u201c\u201d]")
_BAD_APPLY     = re.compile(
    r"sr-company-attachments|cdn-cgi/l/email-protection(?!.*@)", re.IGNORECASE
)


def _shrink(text: str, max_len: int) -> str:
    text = text.strip()
    if len(text) <= max_len:
        return text
    for delim in (" \u2013 ", " \u2014 ", " - "):
        idx = text.find(delim)
        if 25 <= idx <= max_len:
            return text[:idx].rstrip(" :\u2013\u2014-")
    cut = text[:max_len]
    if " " in cut:
        cut = cut[:cut.rfind(" ")]
    cut = re.sub(r"[\s,;:\u2013\u2014\-]+$", "", cut)
    cut = _STOP_TAIL.sub("", cut)
    return cut + "\u2026"


def clean_title(raw: str, max_len: int = 80) -> str:
    t = normalize_text(raw)
    if len(t) > 1 and t[0] in "\"'" and t[-1] == t[0]:
        t = t[1:-1].strip()

    segs = [x.strip() for x in t.split(";") if x.strip()]
    if len(segs) >= 2 and all(2 < len(x) <= 60 for x in segs):
        extra = len(segs) - 1
        return f"{segs[0]} (+{extra} more role{'s' if extra > 1 else ''})"

    extension_flag = bool(_EXTENSION_TAG.match(t))
    if extension_flag:
        t = _EXTENSION_TAG.sub("", t).strip()

    chosen = t
    if _BOILER.match(t):
        q = _QUOTED_SCOPE.search(t)
        if q and len(q.group(1)) >= 12:
            chosen = q.group(1).strip()
        else:
            for delim in (": ", " \u2013 ", " \u2014 ", " - "):
                idx = t.find(delim)
                if idx != -1 and len(t) - idx > 14:
                    chosen = t[idx + len(delim):].strip()
                    break
        if len(chosen) < 8:
            chosen = t

    final = _shrink(chosen, max_len)
    if extension_flag:
        final = f"{final} (Deadline Extended)"
    return final


# ════════════════════════════════════════════════════════════════════════════
# CLOUDFLARE EMAIL DECODER
# ════════════════════════════════════════════════════════════════════════════
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


# ════════════════════════════════════════════════════════════════════════════
# FIELD STANDARDISERS
# ════════════════════════════════════════════════════════════════════════════
QUALIFICATION_TIERS = [
    ("PhD / Doctorate",            ["phd", "ph.d", "doctorate", "doctoral"]),
    ("Master's Degree",            ["master", "msc", "m.sc", "mba", "mphil",
                                    "postgraduate", "master of"]),
    ("Bachelor's Degree",          ["bachelor", "bsc", "b.sc", "b.a", "beng",
                                    "llb", "degree in", "university degree",
                                    "undergraduate", "honours"]),
    ("Diploma",                    ["diploma", "associate degree", "college degree"]),
    ("Professional Certification", ["acca", "cpa", "cfa", "cima", "pmp",
                                    "prince2", "chartered"]),
    ("A-Levels / High School",     ["high school", "secondary school",
                                    "a-level", "a level"]),
]

NO_EXP_KW = ["no experience", "fresh graduate", "freshers", "entry level",
              "entry-level", "0 years", "training provided"]
LESS1_KW  = ["less than 1 year", "under 1 year", "6 months", "less than a year"]


def extract_qualification(text: str) -> str:
    low = (text or "").lower()
    for label, kws in QUALIFICATION_TIERS:
        if any(k in low for k in kws):
            return label
    return ""


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
      "programme manager", "project coordinator", "project administrator",
      "project lead", "project assistant"]),
    ("Monitoring & Evaluation",
     ["m&e", "monitoring and evaluation", "mel officer", "evaluation"]),
    ("Finance & Accounting",
     ["finance officer", "finance manager", "accountant", "auditor",
      "grants officer", "financial", "bookkeeper", "treasury"]),
    ("Human Resources",
     ["human resources", "hr officer", "hr manager", "talent", "recruitment"]),
    ("Communications & Advocacy",
     ["communications", "advocacy", "media", "outreach", "public relations"]),
    ("Consultancy",
     ["consultant", "consultancy", "request for proposals",
      "terms of reference", "rfp", "rfa", "service provider"]),
    ("Health & Nutrition",
     ["health", "nutrition", "wash", "medical", "nurse", "clinical"]),
    ("Information Technology",
     ["it officer", "software", "developer", "gis", "qgis",
      "database", "data analyst"]),
    ("Logistics & Procurement",
     ["logistics", "procurement", "supply chain", "warehouse", "fleet"]),
    ("Research",
     ["researcher", "research", "study", "survey"]),
    ("Administration",
     ["administrator", "admin", "secretary", "receptionist", "operations"]),
]


def infer_job_field(title: str, category: str, desc: str) -> str:
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
    if re.search(
        r"\bconsultan(t|cy)\b|request for proposals?|terms of reference"
        r"|\btor\b|\brfp\b|\brfa\b|service provider", combined
    ):
        return "Consultancy / Contract"
    if re.search(r"\bpart[-\s]?time\b", combined):
        return "Part-time"
    if re.search(r"\bcontract\b|\bfixed[-\s]term\b|\btemporary\b", combined):
        return "Consultancy / Contract"
    return "Full-time"


def detect_company_type(text: str) -> str:
    tl = (text or "").lower()
    if re.search(
        r"\bundp\b|\bunicef\b|\bwfp\b|\bunhcr\b|\bwho\b|\bilo\b"
        r"|united nations|\biom\b|\bunesco\b|\bunfpa\b", tl
    ):
        return "UN Agency"
    if re.search(r"\bingo\b|international ngo|international non-governmental", tl):
        return "INGO"
    if re.search(
        r"\bngo\b|non.?governmental|non.?profit|nonprofit"
        r"|foundation|charity|humanitarian", tl
    ):
        return "NGO / Non-Profit"
    if re.search(r"\bembassy\b|\bgovernment\b|ministry of|\bgovernmental\b", tl):
        return "Government / Embassy"
    if re.search(r"\bsocial enterprise\b", tl):
        return "Social Enterprise"
    if re.search(r"\buniversity\b|research institute|\binstitute\b", tl):
        return "Academic / Research"
    return "NGO / Development"


_SALARY_PATTERNS = [
    (r"VND[\s\xa0]*([\d,\.]+)[\s\xa0]*(?:to|-|–)[\s\xa0]*(?:VND[\s\xa0]*)?([\d,\.]+)",
     "vnd_range"),
    (r"VND[\s\xa0]*([\d,\.]+)", "vnd_single"),
    (r"USD[\s\xa0]*([\d]{1,3}(?:,\d{3})+(?:[\s\xa0]*[-–][\s\xa0]*[\d,\.]+)?)", "usd"),
    (r"\$([\d]{1,3}(?:,\d{3})+(?:[\s\xa0]*[-–][\s\xa0]*\$?[\d,\.]+)?"
     r"(?:[\s\xa0]*/[\s\xa0]*\w+)?)", "usd"),
    (r"([\d,]+(?:[\s\xa0]*[-–][\s\xa0]*[\d,]+)?[\s\xa0]*/[\s\xa0]*"
     r"(?:month|year|day|hour))", "rate"),
]


def extract_salary(text: str) -> str:
    src = (text or "")[:4000]
    for pat, kind in _SALARY_PATTERNS:
        m = re.search(pat, src, re.IGNORECASE)
        if not m:
            continue
        if kind == "vnd_range" and m.lastindex >= 2 and m.group(2):
            return f"VND {m.group(1).strip()} - {m.group(2).strip()}"
        if kind == "vnd_single":
            return f"VND {m.group(1).strip()}"
        val = m.group(0).strip().rstrip(".,")
        if re.search(r"\d{4,}|\d{1,3},\d{3}", val):
            return val
    return ""


# ════════════════════════════════════════════════════════════════════════════
# ORG LOOKUP TABLE
# ════════════════════════════════════════════════════════════════════════════
ORG_DB = {
    "fhi 360": {
        "website": "https://www.fhi360.org",
        "details": "FHI 360 is a nonprofit human development organization dedicated to improving lives in lasting ways by advancing integrated, locally driven solutions. Working with a diverse mix of partners, FHI 360 serves more than 70 countries and all U.S. states and territories.",
        "founded": "1971", "logo": "https://www.fhi360.org/themes/custom/fhi360/logo.svg",
        "industry": "International Development / Public Health", "type": "INGO",
    },
    "wwf": {
        "website": "https://wwf.org",
        "details": "WWF (World Wide Fund for Nature) is one of the world's largest and most respected independent conservation organizations.",
        "founded": "1961", "logo": "https://wwf.org/wp-content/themes/wwf-rw/assets/images/header/wwf-logo.svg",
        "industry": "Environmental Conservation", "type": "INGO",
    },
    "oxfam": {
        "website": "https://www.oxfam.org",
        "details": "Oxfam is a global movement of people who are fighting inequality to end poverty and injustice.",
        "founded": "1942", "logo": "https://www.oxfam.org/themes/contrib/oxfam/logo.svg",
        "industry": "Humanitarian / Development", "type": "INGO",
    },
    "care": {
        "website": "https://www.care.org",
        "details": "CARE is a leading humanitarian organization fighting global poverty, placing special focus on working alongside poor women and girls.",
        "founded": "1945", "logo": "https://www.care.org/wp-content/themes/care2019/assets/images/logo.svg",
        "industry": "Humanitarian / Development", "type": "INGO",
    },
    "wvi": {
        "website": "https://www.wvi.org",
        "details": "World Vision International is a Christian humanitarian organization dedicated to working with children, families, and their communities.",
        "founded": "1950", "logo": "https://www.wvi.org/sites/default/files/2019-10/WV_Logo_RGB.png",
        "industry": "Humanitarian / Development", "type": "INGO",
    },
    "snv": {
        "website": "https://www.snv.org",
        "details": "SNV Netherlands Development Organisation is a mission-driven global development partner working in agriculture, energy, and water.",
        "founded": "1965", "logo": "https://www.snv.org/themes/custom/snv/logo.svg",
        "industry": "International Development", "type": "INGO",
    },
    "crs": {
        "website": "https://www.crs.org",
        "details": "Catholic Relief Services is the official international humanitarian agency of the Catholic community in the United States.",
        "founded": "1943", "logo": "https://www.crs.org/sites/default/files/crs-logo.png",
        "industry": "Humanitarian / Development", "type": "INGO",
    },
    "plan international": {
        "website": "https://plan-international.org",
        "details": "Plan International is an independent development and humanitarian organization that advances children's rights and equality for girls.",
        "founded": "1937", "logo": "https://plan-international.org/uploads/2022/01/Plan_International_logo.svg",
        "industry": "Children's Rights / Development", "type": "INGO",
    },
    "giz": {
        "website": "https://www.giz.de/en",
        "details": "GIZ is a federal enterprise supporting the German Government in international cooperation for sustainable development.",
        "founded": "1975", "logo": "https://www.giz.de/static/en/images/giz-logo.png",
        "industry": "International Development / Technical Cooperation", "type": "Government / Embassy",
    },
    "unfpa": {
        "website": "https://www.unfpa.org",
        "details": "UNFPA, the United Nations Population Fund, is the UN sexual and reproductive health agency.",
        "founded": "1969", "logo": "https://www.unfpa.org/sites/default/files/pub-pdf/UNFPA_logo_blue.png",
        "industry": "Sexual & Reproductive Health", "type": "UN Agency",
    },
    "icraf": {
        "website": "https://www.cifor-icraf.org",
        "details": "CIFOR-ICRAF is a research center dedicated to transforming lives and landscapes through forest, tree and agroforestry science.",
        "founded": "1978", "logo": "https://www.cifor-icraf.org/wp-content/uploads/2021/06/CIFOR-ICRAF_logo.svg",
        "industry": "Forestry / Agroforestry Research", "type": "Academic / Research",
    },
    "wcs": {
        "website": "https://www.wcs.org",
        "details": "The Wildlife Conservation Society saves wildlife and wild places worldwide through science, conservation action, and education.",
        "founded": "1895", "logo": "https://www.wcs.org/images/wcs-logo.png",
        "industry": "Wildlife Conservation", "type": "NGO / Non-Profit",
    },
    "helvetas": {
        "website": "https://www.helvetas.org",
        "details": "HELVETAS is a Swiss organization for international cooperation committed to a just world in which all people determine the course of their lives.",
        "founded": "1955", "logo": "https://www.helvetas.org/typo3conf/ext/sitepackage/Resources/Public/Images/logo.svg",
        "industry": "International Development", "type": "INGO",
    },
    "reach": {
        "website": "https://www.reach-initiative.org",
        "details": "REACH Initiative supports humanitarian communities through information management, assessments, and research.",
        "founded": "2010", "logo": "https://www.reach-initiative.org/wp-content/uploads/2019/05/reach-logo.png",
        "industry": "Humanitarian Information Management", "type": "NGO / Non-Profit",
    },
    "rikolto": {
        "website": "https://www.rikolto.org",
        "details": "Rikolto is an international NGO with 40+ years of experience partnering with farmer organizations worldwide.",
        "founded": "1976", "logo": "https://www.rikolto.org/sites/default/files/rikolto_logo.png",
        "industry": "Agriculture / Food Systems", "type": "INGO",
    },
    "chai": {
        "website": "https://www.clintonhealthaccess.org",
        "details": "The Clinton Health Access Initiative is a global health organization committed to saving lives in low- and middle-income countries.",
        "founded": "2002", "logo": "https://www.clintonhealthaccess.org/content/uploads/2021/01/CHAI-logo.png",
        "industry": "Global Health", "type": "NGO / Non-Profit",
    },
    "blue dragon": {
        "website": "https://www.bluedragon.org",
        "details": "Blue Dragon Children's Foundation is an Australian NGO working in Vietnam to rescue children from slavery and exploitation.",
        "founded": "2002", "logo": "https://www.bluedragon.org/wp-content/uploads/2020/10/blue-dragon-logo.png",
        "industry": "Child Protection", "type": "INGO",
    },
    "sci": {
        "website": "https://www.savethechildren.org",
        "details": "Save the Children International is the world's leading independent organization for children, working in around 100 countries.",
        "founded": "1919", "logo": "https://www.savethechildren.org/content/dam/usa/images/logos/save-the-children-logo.png",
        "industry": "Children's Rights / Humanitarian", "type": "INGO",
    },
    "samaritan's purse": {
        "website": "https://www.samaritanspurse.org",
        "details": "Samaritan's Purse is an international Christian relief and development organization providing aid to victims of war, famine, disease, and poverty.",
        "founded": "1970", "logo": "https://www.samaritanspurse.org/images/sp-logo.png",
        "industry": "Humanitarian Relief / Development", "type": "INGO",
    },
    "samaritans purse": {
        "website": "https://www.samaritanspurse.org",
        "details": "Samaritan's Purse is an international Christian relief and development organization providing aid to victims of war, famine, disease, and poverty.",
        "founded": "1970", "logo": "https://www.samaritanspurse.org/images/sp-logo.png",
        "industry": "Humanitarian Relief / Development", "type": "INGO",
    },
    "vvob": {
        "website": "https://www.vvob.org",
        "details": "VVOB is a Belgian NGO that partners with governments and education systems in the Global South to improve education quality.",
        "founded": "1982", "logo": "https://www.vvob.org/sites/default/files/vvob-logo.png",
        "industry": "Education / Development", "type": "INGO",
    },
    "ide": {
        "website": "https://www.ideglobal.org",
        "details": "iDE is an international NGO that creates income and livelihood opportunities for poor rural households.",
        "founded": "1982", "logo": "https://www.ideglobal.org/wp-content/uploads/2020/01/iDE-logo.png",
        "industry": "Agriculture / WASH", "type": "INGO",
    },
    "mrc": {
        "website": "https://www.mrcmekong.org",
        "details": "The Mekong River Commission is an inter-governmental organization managing shared water resources of the Mekong River.",
        "founded": "1995", "logo": "https://www.mrcmekong.org/assets/images/logo.png",
        "industry": "Water Resources / Environmental Governance", "type": "Government / Embassy",
    },
    "streets": {
        "website": "https://www.streetsinternational.org",
        "details": "Streets International provides disadvantaged Vietnamese youth with professional culinary and hospitality training.",
        "founded": "2007", "logo": "https://www.streetsinternational.org/wp-content/uploads/2020/01/Streets-Logo.png",
        "industry": "Vocational Training / Hospitality", "type": "NGO / Non-Profit",
    },
    "koto": {
        "website": "https://koto.com.au",
        "details": "KOTO (Know One Teach One) is a social enterprise providing disadvantaged youth with hospitality and life skills training.",
        "founded": "1999", "logo": "https://koto.com.au/wp-content/uploads/2021/06/KOTO-Logo.png",
        "industry": "Vocational Training / Social Enterprise", "type": "Social Enterprise",
    },
}


def _norm_key(name: str) -> str:
    n = name.lower().strip()
    n = re.sub(r"['\u2018\u2019\u201c\u201d]", "", n)
    n = re.sub(r"[^\w\s]", " ", n)
    n = re.sub(r"\s+", " ", n).strip()
    return n


def _enrich_company(name_raw: str, existing_web: str) -> dict:
    key = _norm_key(name_raw)
    info = ORG_DB.get(key, {})
    best_web = info.get("website", "")
    if not best_web:
        best_web = "" if is_ats_or_bad(existing_web) else existing_web
    return {
        "website":  best_web,
        "logo":     info.get("logo",     ""),
        "founded":  info.get("founded",  ""),
        "details":  info.get("details",  ""),
        "industry": info.get("industry", ""),
        "type":     info.get("type",     ""),
    }


# ════════════════════════════════════════════════════════════════════════════
# COLUMN MAPPING
# ════════════════════════════════════════════════════════════════════════════
def map_columns(df: pd.DataFrame) -> pd.DataFrame:
    col_lookup = {c.lower().strip(): c for c in df.columns}
    ALIASES = {
        "Job Title":          ["title", "position", "role", "vacancy", "job name"],
        "Job Type":           ["type", "employment type", "contract type", "work type"],
        "Job Qualifications": ["qualifications", "qualification", "education", "degree"],
        "Job Experience":     ["experience", "exp", "years experience", "work experience"],
        "Job Location":       ["location", "city", "town", "region", "county", "place"],
        "Job Field":          ["field", "sector", "category", "job category"],
        "Date Posted":        ["posted", "post date", "published", "created"],
        "Deadline":           ["closing date", "expiry", "apply by", "close date", "end date"],
        "Job Description":    ["description", "details", "duties", "responsibilities",
                               "summary", "content", "job details"],
        "Application":        ["apply", "apply url", "apply link", "application url",
                               "application link", "apply email", "application email",
                               "email", "contact email", "how to apply"],
        "Company URL":        ["company url", "company link", "employer url"],
        "Company Name":       ["company", "employer", "organisation", "organization", "firm"],
        "Company Logo":       ["logo", "logo url", "company image", "company logo url"],
        "Company Industry":   ["industry", "company sector", "business type"],
        "Company Founded":    ["founded", "year founded", "established"],
        "Company Type":       ["company type", "org type"],
        "Company Website":    ["website", "company web", "web", "site", "company site"],
        "Company Address":    ["address", "location address"],
        "Company Details":    ["company description", "about company", "company bio",
                               "company profile", "about", "company info"],
        "Job URL":            ["url", "source url", "source", "link", "job link",
                               "original url", "reference url"],
        "Estimated Deadline": ["estimated expiry", "calculated deadline", "auto deadline"],
        "Salary Range":       ["salary", "salary range", "pay", "remuneration",
                               "compensation", "pay range", "wage", "wages"],
    }
    for internal, aliases in ALIASES.items():
        if internal in df.columns:
            continue
        for alias in aliases:
            if alias in col_lookup:
                df = df.rename(columns={col_lookup[alias]: internal})
                col_lookup = {c.lower().strip(): c for c in df.columns}
                break
    for col in APPSCRIPT_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    return df


def print_column_mapping(df: pd.DataFrame):
    print("┌─ COLUMN MAPPING " + "─" * 42 + "┐")
    for col in APPSCRIPT_COLUMNS:
        has_data = (
            col in df.columns
            and df[col].notna().any()
            and (df[col].astype(str).str.strip() != "").any()
        )
        if has_data:
            sample = str(df[col].replace("", pd.NA).dropna().iloc[0])[:50]
            print(f"│ ✅ {col:<25} → {sample!r}")
        else:
            print(f"│ ⚠️  {col:<25} → NOT FOUND / empty")
    print("└" + "─" * 59 + "┘\n")


# ════════════════════════════════════════════════════════════════════════════
# DUPLICATE / PROCESSED-ID TRACKER
# ════════════════════════════════════════════════════════════════════════════
def _init_tracker():
    if not os.path.exists(PROCESSED_IDS_FILE):
        pd.DataFrame(columns=[
            "Job ID", "Job URL", "Job Title", "Company Name",
            "Status", "Timestamp", "Sheet Row",
        ]).to_csv(PROCESSED_IDS_FILE, index=False)


def load_processed_ids() -> tuple:
    _init_tracker()
    df = pd.read_csv(PROCESSED_IDS_FILE)
    return (
        set(df["Job ID"].fillna("").astype(str)),
        set(df.get("Job URL", pd.Series()).fillna("").astype(str)),
    )


def _upsert_row(job_id: str, updates: dict):
    _init_tracker()
    df = pd.read_csv(PROCESSED_IDS_FILE)
    mask = df["Job ID"].astype(str) == str(job_id)
    if mask.any():
        for col, val in updates.items():
            if col in df.columns:
                df.loc[mask, col] = val
        df.loc[mask, "Timestamp"] = datetime.now().isoformat()
    else:
        row = {"Job ID": job_id, "Timestamp": datetime.now().isoformat()}
        row.update(updates)
        df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    df.to_csv(PROCESSED_IDS_FILE, index=False)


def mark_read(job_id, job_url, title, company, sheet_row):
    _upsert_row(job_id, {"Job URL": job_url, "Job Title": title,
                          "Company Name": company, "Status": "read",
                          "Sheet Row": sheet_row})


def mark_posted(job_id, wp_id, wp_url):
    _upsert_row(job_id, {"Status": f"posted|wp_id={wp_id}|{wp_url}"})


def mark_failed(job_id, reason):
    _upsert_row(job_id, {"Status": f"failed|{reason}"})


def print_tracker_summary():
    if not os.path.exists(PROCESSED_IDS_FILE):
        return
    df = pd.read_csv(PROCESSED_IDS_FILE)
    print(f"\n{'═'*55}")
    print(f" TRACKER SUMMARY ({len(df)} total records)")
    print(f"{'═'*55}")
    counts = df["Status"].str.split("|").str[0].value_counts()
    icons = {"read": "🔵", "paraphrased": "🟡", "posted": "✅", "failed": "❌"}
    for status, count in counts.items():
        print(f" {icons.get(status, '⚪')} {status:<15} {count}")
    print(f"{'═'*55}\n")


def make_job_id(row: pd.Series, idx: int) -> str:
    src = sanitize_text(str(row.get("Job URL", "")), is_url=True)
    if src:
        return hashlib.md5(src.encode()).hexdigest()[:16]
    seed = f"{row.get('Job Title', '')}{row.get('Company Name', '')}{idx}"
    return hashlib.md5(seed.encode()).hexdigest()[:16]


# ════════════════════════════════════════════════════════════════════════════
# MISTRAL API
# ════════════════════════════════════════════════════════════════════════════
def mistral_generate(prompt: str, max_tokens: int = 400, temperature: float = 0.7) -> str:
    if not MISTRAL_API_KEY:
        log.warning("MISTRAL_API_KEY not set — skipping paraphrase")
        return ""
    try:
        response = requests.post(
            MISTRAL_URL,
            headers={
                "Authorization": f"Bearer {MISTRAL_API_KEY}",
                "Content-Type": "application/json",
            },
            json={
                "model": MISTRAL_MODEL,
                "messages": [{"role": "user", "content": prompt}],
                "max_tokens": max_tokens,
                "temperature": temperature,
            },
            timeout=30,
        )
        response.raise_for_status()
        return response.json()["choices"][0]["message"]["content"].strip()
    except Exception as e:
        log.error(f"Mistral API error: {e}")
        return ""


def grammar_correct(text: str) -> str:
    gt = _get_grammar_tool()
    if gt is None:
        return text
    try:
        return language_tool_python.utils.correct(text, gt.check(text))
    except Exception:
        return text


def similarity_score(a: str, b: str) -> float:
    model = _get_similarity_model()
    if model is None:
        return 0.8   # assume OK when model unavailable
    try:
        emb = model.encode([a, b], convert_to_tensor=True)
        return float(st_util.pytorch_cos_sim(emb[0], emb[1]))
    except Exception:
        return 0.0


def clean_output(text: str) -> str:
    text = _fix_mojibake(text)
    for pat in [r"\[/?INST\]", r"</?s>",
                r"(?i)(rewritten?|rephrased?|output|paraphrase[d]?)[:\s]+",
                r"\*\*", r"###", r"---"]:
        text = re.sub(pat, "", text)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return grammar_correct(text.strip())


def paraphrase_title(title: str) -> str:
    clean = sanitize_text(title)
    if not clean:
        return title

    print(f"\n ┌─ TITLE PARAPHRASE {'─'*45}")
    print(f" │ Original : \"{clean}\"")
    print(f" │ {'─'*60}")

    best_result, best_sim = None, 0.0

    for attempt in range(4):
        temp = round(0.68 + attempt * 0.06, 2)
        print(f" │ Attempt {attempt+1} (temp={temp}):")
        prompt = (
            f"Rewrite this job title professionally using different words. "
            f"Output ONLY the rewritten title, nothing else. "
            f"Keep it between 4 and 12 words.\n\nJob title: {clean}"
        )
        raw = mistral_generate(prompt, max_tokens=50, temperature=temp)
        result = clean_output(raw).split("\n")[0].strip().strip('"').strip("'")
        wc = len(result.split()) if result else 0
        sim = similarity_score(clean, result) if result else 0.0
        is_dup = result.lower().strip() == clean.lower().strip()

        print(f" │    Output  : \"{result}\"")
        print(f" │    Words   : {wc} | Similarity: {sim:.3f} | Duplicate: {'Yes ⚠️' if is_dup else 'No'}")

        valid = bool(result) and 4 <= wc <= 14 and sim >= 0.55 and not is_dup
        if valid:
            if sim > best_sim:
                best_sim, best_result = sim, result
                print(f" │    → ✅ ACCEPTED (sim={sim:.3f})")
            else:
                print(f" │    → ✅ VALID but not better than best (best={best_sim:.3f})")
        else:
            reasons = []
            if not result:  reasons.append("empty output")
            if wc < 4:      reasons.append(f"too short ({wc}w, min=4)")
            if wc > 14:     reasons.append(f"too long ({wc}w, max=14)")
            if sim < 0.55:  reasons.append(f"sim={sim:.3f} < 0.55")
            if is_dup:      reasons.append("identical to original")
            print(f" │    → ❌ REJECTED — {', '.join(reasons)}")
        print(f" │ {'─'*60}")
        time.sleep(1)

    if best_result:
        print(f" │ 🏆 FINAL: \"{best_result}\" (sim={best_sim:.3f})")
        print(f" └{'─'*65}")
        return best_result
    print(f" │ ⚠️  No valid paraphrase — keeping original")
    print(f" └{'─'*65}")
    return clean


def paraphrase_description(text: str) -> str:
    clean = sanitize_text(text)
    if not clean:
        return text

    paragraphs = [p.strip() for p in clean.split("\n") if p.strip()]
    rewritten, success_count = [], 0

    print(f"\n ┌─ DESCRIPTION PARAPHRASE ({len(paragraphs)} paragraphs) {'─'*25}")

    for i, para in enumerate(paragraphs):
        print(f"\n │ ┌─ Paragraph {i+1}/{len(paragraphs)} {'─'*50}")
        print(f" │ │ ORIGINAL ({len(para.split())} words): {para[:120]}...")
        print(f" │ │ {'─'*60}")

        prompt = (
            f"Rewrite this job description paragraph professionally. "
            f"Keep ALL facts, requirements, and responsibilities. "
            f"Use different sentence structure and vocabulary. "
            f"Output ONLY the rewritten paragraph — no labels, no explanation.\n\n"
            f"Original:\n{para}"
        )

        best_result, best_sim, accepted_text = None, 0.0, None

        for attempt in range(3):
            temp = round(0.65 + attempt * 0.08, 2)
            raw = mistral_generate(prompt, max_tokens=500, temperature=temp)
            result = clean_output(raw).strip()
            rw = len(result.split()) if result else 0
            sim = similarity_score(para, result) if result and rw >= 5 else 0.0

            print(f" │ │ Attempt {attempt+1} (temp={temp}): {rw}w sim={sim:.3f}")
            valid = bool(result) and rw >= 8 and sim >= 0.48
            if valid:
                print(f" │ │    → ✅ ACCEPTED")
                rewritten.append(result)
                success_count += 1
                accepted_text = result
                break
            else:
                reasons = []
                if not result:  reasons.append("empty")
                if rw < 8:      reasons.append(f"too short ({rw}w)")
                if sim < 0.48:  reasons.append(f"sim={sim:.3f} < 0.48")
                print(f" │ │    → ❌ REJECTED — {', '.join(reasons)}")
                if result and sim > best_sim:
                    best_sim, best_result = sim, result
            print(f" │ │ {'─'*60}")
            time.sleep(1)

        if accepted_text is None:
            if best_result and best_sim >= 0.40:
                print(f" │ │ 🔁 FALLBACK — best attempt (sim={best_sim:.3f})")
                rewritten.append(best_result)
                success_count += 1
            else:
                print(f" │ │ ⚠️  KEPT ORIGINAL (best sim={best_sim:.3f} < 0.40)")
                rewritten.append(para)

        print(f" │ └{'─'*62}")

    print(f"\n │ SUMMARY: {success_count}/{len(paragraphs)} paragraphs paraphrased")
    print(f" └{'─'*80}\n")
    return "\n\n".join(rewritten)


def paraphrase_company(text: str) -> str:
    clean = sanitize_text(text)
    if not clean:
        return text

    print(f"\n ┌─ COMPANY PARAPHRASE {'─'*43}")
    print(f" │ Original ({len(clean.split())} words): {clean[:120]}...")
    prompt = (
        f"Rewrite this company description professionally. "
        f"Preserve all facts. Use different wording. "
        f"Output ONLY the rewritten description.\n\nOriginal:\n{clean}"
    )
    raw = mistral_generate(prompt, max_tokens=600, temperature=0.68)
    result = clean_output(raw)
    rw = len(result.split()) if result else 0
    sim = similarity_score(clean, result) if result and rw >= 10 else 0.0

    if result and rw >= 10:
        print(f" │ → ✅ ACCEPTED ({rw}w, sim={sim:.3f})")
        print(f" └{'─'*65}")
        time.sleep(1)
        return result
    reasons = []
    if not result: reasons.append("empty output")
    if rw < 10:    reasons.append(f"too short ({rw}w)")
    print(f" │ → ❌ REJECTED — {', '.join(reasons)} — keeping original")
    print(f" └{'─'*65}")
    time.sleep(1)
    return clean


def paraphrase_tagline(text: str) -> str:
    clean = sanitize_text(text[:300])
    if not clean:
        return text
    print(f"\n ┌─ TAGLINE PARAPHRASE {'─'*43}")
    print(f" │ Original : \"{clean}\"")
    prompt = (
        f"Rewrite this company tagline as a crisp, professional phrase. "
        f"Output ONLY the rewritten tagline (5–12 words). No explanation.\n\n"
        f"Original: {clean}"
    )
    raw = mistral_generate(prompt, max_tokens=35, temperature=0.75)
    result = clean_output(raw).split("\n")[0].strip().strip('"').strip("'")
    wc = len(result.split()) if result else 0

    if result and 3 <= wc <= 15:
        print(f" │ → ✅ \"{result}\"")
        print(f" └{'─'*65}")
        time.sleep(1)
        return result
    print(f" │ → ❌ keeping original")
    print(f" └{'─'*65}")
    time.sleep(1)
    return clean


# ════════════════════════════════════════════════════════════════════════════
# WORDPRESS HELPERS
# ════════════════════════════════════════════════════════════════════════════
def wp_headers():
    token = base64.b64encode(f"{WP_USERNAME}:{WP_APP_PASSWORD}".encode()).decode()
    return {"Authorization": f"Basic {token}", "Content-Type": "application/json"}


def upload_logo(logo_url: str):
    logo_url = sanitize_text(logo_url, is_url=True)
    if not logo_url or not logo_url.startswith("http"):
        return None
    ext = logo_url.lower().rsplit(".", 1)[-1]
    if ext not in ("png", "jpg", "jpeg", "webp", "svg"):
        return None
    try:
        img = requests.get(logo_url, timeout=10)
        img.raise_for_status()
        h = wp_headers()
        h["Content-Disposition"] = f"attachment; filename={logo_url.split('/')[-1]}"
        h["Content-Type"] = img.headers.get("content-type", "image/jpeg")
        r = requests.post(WP_MEDIA_URL, headers=h, data=img.content,
                          auth=(WP_USERNAME, WP_APP_PASSWORD), timeout=15, verify=False)
        r.raise_for_status()
        return r.json().get("id")
    except Exception as e:
        log.error(f"Logo upload error: {e}")
        return None


def get_or_create_term(taxonomy_url: str, name: str):
    if not name or not name.strip():
        return None
    slug = re.sub(r"[^a-z0-9-]", "-", name.lower().strip())
    try:
        r = requests.get(f"{taxonomy_url}?slug={slug}",
                         headers=wp_headers(), timeout=10, verify=False)
        terms = r.json()
        if isinstance(terms, list) and terms:
            return terms[0]["id"]
    except Exception:
        pass
    try:
        r = requests.post(taxonomy_url,
                          json={"name": name, "slug": slug},
                          headers=wp_headers(),
                          auth=(WP_USERNAME, WP_APP_PASSWORD),
                          timeout=10, verify=False)
        return r.json().get("id")
    except Exception as e:
        log.error(f"Term create error '{name}': {e}")
        return None


def save_company(company_data: dict):
    name = sanitize_text(company_data.get("company_name", ""))
    if not name or name in ("Unknown Company", "nan"):
        return None, None
    slug = re.sub(r"[^a-z0-9-]", "-", name.lower())
    try:
        r = requests.get(f"{WP_COMPANY_URL}?slug={slug}",
                         headers=wp_headers(), timeout=10, verify=False)
        posts = r.json()
        if isinstance(posts, list) and posts:
            log.info(f"⏭ Company exists: {name}")
            return posts[0]["id"], posts[0].get("link")
    except Exception:
        pass
    attachment_id = upload_logo(company_data.get("company_logo", ""))
    raw     = company_data.get("company_details", "")
    details = paraphrase_company(raw) if raw and ENABLE_PARAPHRASE else sanitize_text(raw)
    tagline = paraphrase_tagline(raw[:300]) if raw and ENABLE_PARAPHRASE else ""
    payload = {
        "title":  name,
        "content": details,
        "status": "publish",
        "featured_media": attachment_id or 0,
        "meta": {
            "_company_name":     name,
            "_company_logo":     str(attachment_id) if attachment_id else "",
            "_company_industry": sanitize_text(company_data.get("company_industry", "")),
            "_company_website":  sanitize_text(company_data.get("company_website", ""), is_url=True),
            "_company_tagline":  tagline,
        },
    }
    try:
        r = requests.post(WP_COMPANY_URL, json=payload, headers=wp_headers(),
                          auth=(WP_USERNAME, WP_APP_PASSWORD), timeout=15, verify=False)
        r.raise_for_status()
        post = r.json()
        log.info(f"✅ Company posted: {name} → ID {post.get('id')}")
        return post.get("id"), post.get("link")
    except Exception as e:
        log.error(f"Company post error '{name}': {e}")
        return None, None


def normalise_job_type(raw: str) -> str:
    return JOB_TYPE_MAPPING.get(raw.lower().strip(), "full-time")


def save_job(row: pd.Series, title: str, description: str) -> tuple:
    if not WP_BASE or not WP_USERNAME or not WP_APP_PASSWORD:
        log.warning("WordPress credentials not configured — skipping WP post")
        return None, None

    h = wp_headers()
    for jt_label in ["Full Time", "Part Time", "Contract",
                     "Temporary", "Freelance", "Internship", "Volunteer"]:
        get_or_create_term(f"{WP_BASE}/job_listing_type", jt_label)

    location      = sanitize_text(str(row.get("Job Location", "Vietnam")))
    raw_type      = sanitize_text(str(row.get("Job Type", "Full-time")))
    job_type_s    = normalise_job_type(raw_type)
    company       = sanitize_text(str(row.get("Company Name", "")))
    application   = sanitize_text(str(row.get("Application", "")), is_url=True)
    deadline      = sanitize_text(str(row.get("Deadline", "")))
    logo_url      = sanitize_text(str(row.get("Company Logo", "")), is_url=True)
    co_website    = sanitize_text(str(row.get("Company Website", "")), is_url=True)
    qualif        = sanitize_text(str(row.get("Job Qualifications", "")))
    experience    = sanitize_text(str(row.get("Job Experience", "")))
    industry      = sanitize_text(str(row.get("Company Industry", "")))
    co_address    = sanitize_text(str(row.get("Company Address", "")))
    job_field     = sanitize_text(str(row.get("Job Field", "")))
    job_url       = sanitize_text(str(row.get("Job URL", "")), is_url=True)
    co_founded    = sanitize_text(str(row.get("Company Founded", "")))
    co_type       = sanitize_text(str(row.get("Company Type", "")))
    salary        = sanitize_text(str(row.get("Salary Range", "")))
    if not deadline:
        deadline = sanitize_text(str(row.get("Estimated Deadline", "")))

    is_email = bool(re.match(
        r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$", application))
    is_url_v = bool(re.match(r"^https?://[^\s]+$", application))
    if not (is_email or is_url_v):
        application = ""

    slug = re.sub(r"[^a-z0-9-]", "-", title.lower())[:80]
    try:
        r = requests.get(f"{WP_URL}?slug={slug}", headers=h, timeout=10, verify=False)
        posts = r.json()
        if isinstance(posts, list) and posts:
            log.info(f"⏭ Job already on WP: {title}")
            return posts[0]["id"], posts[0].get("link")
    except Exception:
        pass

    attachment_id    = upload_logo(logo_url)
    region_term_id   = get_or_create_term(f"{WP_BASE}/job_listing_region", location)
    job_type_term_id = get_or_create_term(f"{WP_BASE}/job_listing_type",
                                           job_type_s.replace("-", " ").title())
    payload = {
        "title":   title,
        "content": description,
        "status":  "publish",
        "featured_media": attachment_id or 0,
        "meta": {
            "_job_title":          title,
            "_job_location":       location,
            "_job_type":           job_type_s,
            "_job_description":    description,
            "_application":        application,
            "_job_expires":        deadline,
            "_company_name":       company,
            "_company_website":    co_website,
            "_company_logo":       str(attachment_id) if attachment_id else "",
            "_company_industry":   industry,
            "_company_address":    co_address,
            "_company_founded":    co_founded,
            "_company_type":       co_type,
            "_job_qualifications": qualif,
            "_job_experiences":    experience,
            "_job_field":          job_field,
            "_job_source_url":     job_url,
            "_job_salary":         salary,
        },
    }
    if region_term_id:
        payload["job_listing_region"] = [region_term_id]
    if job_type_term_id:
        payload["job_listing_type"] = [job_type_term_id]

    for attempt in range(3):
        try:
            r = requests.post(WP_URL, json=payload, headers=h,
                              auth=(WP_USERNAME, WP_APP_PASSWORD),
                              timeout=20, verify=False)
            r.raise_for_status()
            post = r.json()
            log.info(f"✅ Job posted: '{title}' → WP ID {post.get('id')}")
            return post.get("id"), post.get("link")
        except Exception as e:
            log.error(f"Job post attempt {attempt+1} failed: {e}")
            if attempt < 2:
                time.sleep(2 ** attempt)
    return None, None


# ════════════════════════════════════════════════════════════════════════════
# LISTING PARSER  (CareerONE.vn specific)
# ════════════════════════════════════════════════════════════════════════════
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
            "title":    link.get_text(" ", strip=True),
            "url":      url,
            "org":      cells[1].get_text(" ", strip=True),
            "deadline": cells[2].get_text(" ", strip=True),
            "location": cells[3].get_text(" ", strip=True),
        })
    return rows


# ════════════════════════════════════════════════════════════════════════════
# DESCRIPTION CLEANER  ← NEW: strips all boilerplate from the body text
# ════════════════════════════════════════════════════════════════════════════

# Lines that signal the start of the application / submission block.
# Once matched, everything from here to the end is discarded.
_APPLY_LINE = re.compile(
    r"^\s*(?:"
    r"to\s+(?:apply|submit)|"
    r"(?:please\s+)?send\s+(?:your\s+)?(?:cv|resume|application)|"
    r"interested\s+candidates\s+(?:should|may|are)|"
    r"how\s+to\s+apply|"
    r"application\s+(?:process|instructions?|deadline|submission)|"
    r"submit\s+(?:your\s+)?(?:application|cv|resume)|"
    r"(?:please\s+)?(?:email|send)\s+(?:your\s+)?(?:cv|resume|application)|"
    r"closing\s+date|"
    r"deadline\s+for\s+(?:applications?|submission)|"
    r"applications?\s+(?:close|due|deadline)|"
    r"we\s+kindly\s+request|"
    r"professional\s+candidates\s+are\s+encouraged\s+to\s+apply|"
    r"only\s+applications?\s+received|"
    r"questions\s+due\s+date|"
    r"responses?\s+to\s+(?:any\s+)?inquir|"
    r"prepare\s+and\s+submit\s+a\s+(?:competitive\s+)?quotation|"
    r"please\s+ensure\s+the\s+subject\s+line|"
    r"we\s+invite\s+qualified\s+candidates|"
    r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}"   # bare email
    r")",
    re.IGNORECASE,
)

# Standalone date lines like "June 22, 2026" or "5 June 26, 2026 at 5:00 PM"
_DATE_LINE = re.compile(
    r"^\s*\d{0,2}\s*(?:january|february|march|april|may|june|july|august|"
    r"september|october|november|december)\s+\d{1,2}(?:st|nd|rd|th)?"
    r"(?:,?\s*\d{4})?(?:\s+at\s+[\d:]+\s*(?:AM|PM))?\s*$",
    re.IGNORECASE,
)

# Whole-sentence RFQ / procurement boilerplate patterns — stripped anywhere in body.
_BOILER_SENTENCES = re.compile(
    r"(?:"
    # RFQ intro / availability notices
    r"the\s+request\s+for\s+quotation\s+\(rfq\)[^.]*\."
    r"|rfq\s+document\s+is\s+(?:now\s+)?available[^.]*\."
    r"|interested\s+parties\s+must\s+submit[^.]*\."
    r"|ensure\s+all\s+required\s+documentation[^.]*\."
    r"|the\s+rfq\s+outlines[^.]*\."
    r"|reviewing\s+the\s+document\s+thoroughly[^.]*\."
    # Submission / quotation instructions
    r"|prepare\s+and\s+submit\s+a\s+competitive\s+quotation[^.]*\."
    r"|please\s+ensure\s+the\s+subject\s+line[^.]*\."
    r"|quote-it\s+equipment[^.]*\."
    # Responses / Q&A schedule lines
    r"|responses\s+to\s+any\s+inquiries\s+will\s+be\s+published[^.]*\."
    r"|questions\s+due\s+date[^.]*\."
    # "Only applications received by…" closers
    r"|only\s+applications?\s+received\s+by[^.]*\."
    # Generic "we invite qualified candidates" block
    r"|we\s+invite\s+qualified\s+candidates\s+to\s+submit[^.]*\."
    r"|to\s+be\s+eligible[^.]*\."
    r")",
    re.IGNORECASE | re.DOTALL,
)

# Unfilled template placeholders like [insert date], [X years], [specific field]
_TEMPLATE_PLACEHOLDER = re.compile(
    r"\[\s*(?:insert\s+\w[\w\s]*|x\s+years?|specific\s+[\w\s]+|degree[^\]]*"
    r"|certification[^\]]*|industry[^\]]*|field[^\]]*|list\s+key[^\]]*"
    r"|relevant\s+[\w\s]*|required\s+[\w\s]*|\d+\s+[\w\s]*)\s*\]",
    re.IGNORECASE,
)

# Lines consisting ONLY of a placeholder (after stripping)
_PLACEHOLDER_ONLY_LINE = re.compile(
    r"^\s*\[[^\]]{1,80}\]\s*$"
)

# Meaningless fragment lines — link-text debris, lone punctuation, "see attachment" etc.
_FRAGMENT_LINE = re.compile(
    r"^\s*(?:"
    r"[:\-–•·]\s*(?:please\s+see\s+details?\s+in\s+the)?"
    r"|please\s+see\s+details?\s+in\s+the"
    r"|see\s+(?:details?|attachment|document|below|above)\s*\.?"
    r"|[.\-–•·:,;]+\s*"
    r")\s*$",
    re.IGNORECASE,
)

# Sentences whose subject is a filled-in placeholder — e.g. starts with "[…]"
# or contains only generic boilerplate around a placeholder
_PLACEHOLDER_SENTENCE = re.compile(
    r"[^.]*\[\s*(?:insert|x\s+years?|specific|degree|certification|industry|"
    r"field|list\s+key|relevant|required|\d+\s+\w+)[^\]]*\][^.]*\.",
    re.IGNORECASE,
)

# Sentences that are pure procurement/RFQ boilerplate prose (no bracket needed)
_PROCUREMENT_PROSE = re.compile(
    r"(?:"
    r"fhi\s+360\s+is\s+in\s+search\s+of\s+a\s+skilled\s+vendor[^.]*\."
    r"|we\s+seek\s+candidates\s+with\s+a\s+minimum\s+of\s+five\s+years[^.]*\."
    r"|proficiency\s+in\s+\[specific\s+software[^\]]*\][^.]*\."
    r"|the\s+ideal\s+applicant\s+should\s+hold\s+a\s+\[[^\]]*\][^.]*\."
    r"|excellent\s+communication\s+and\s+teamwork\s+abilities\s+are\s+required[^.]*\."
    r"|additionally,\s+experience\s+with\s+\[[^\]]*\][^.]*\."
    r")",
    re.IGNORECASE | re.DOTALL,
)


def clean_description(raw: str) -> str:
    """
    Multi-pass pipeline that removes:
      1. Application / submission block (cutoff on first trigger line)
      2. Inline email addresses
      3. Standalone date-only lines
      4. RFQ / procurement boilerplate sentences (regex)
      5. Unfilled template placeholders [insert …] etc.
      6. Sentences that are entirely placeholder or procurement prose
      7. Leftover blank lines / excess whitespace
    """
    if not raw:
        return raw

    # ── Pass 1: cutoff at first application/submission trigger line ──────
    kept_lines = []
    for line in raw.split("\n"):
        if _APPLY_LINE.match(line) or _DATE_LINE.match(line):
            break
        # Also stop if the line contains a bare email address anywhere
        if re.search(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}", line):
            break
        kept_lines.append(line)
    text = "\n".join(kept_lines)

    # ── Pass 2: scrub any stray emails that slipped through ───────────────
    text = re.sub(
        r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}", "", text
    )

    # ── Pass 3: remove standalone date-only lines (mid-text) ─────────────
    text = re.sub(
        r"\n[ \t]*\d{0,2}\s*(?:january|february|march|april|may|june|july|august|"
        r"september|october|november|december)\s+\d{1,2}(?:st|nd|rd|th)?"
        r"(?:,?\s*\d{4})?(?:\s+at\s+[\d:]+\s*(?:AM|PM))?[ \t]*\n",
        "\n", text, flags=re.IGNORECASE,
    )

    # ── Pass 4: strip RFQ / procurement boilerplate sentences ────────────
    text = _BOILER_SENTENCES.sub("", text)
    text = _PROCUREMENT_PROSE.sub("", text)

    # ── Pass 5: strip lines that are ONLY a placeholder or a fragment ────
    lines = []
    for line in text.split("\n"):
        if _PLACEHOLDER_ONLY_LINE.match(line):
            continue
        if _FRAGMENT_LINE.match(line):
            continue
        lines.append(line)
    text = "\n".join(lines)

    # ── Pass 6: strip sentences containing unfilled placeholders ─────────
    text = _PLACEHOLDER_SENTENCE.sub("", text)
    # Also zap any residual bracket content that didn't form a full sentence
    text = _TEMPLATE_PLACEHOLDER.sub("", text)

    # ── Pass 7: normalise whitespace ──────────────────────────────────────
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r" *\n *", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


# ════════════════════════════════════════════════════════════════════════════
# DETAIL PARSER
# ════════════════════════════════════════════════════════════════════════════
DETAIL_LABELS = {
    "organisation":        "org",
    "organization":        "org",
    "job location":        "location",
    "application deadline":"deadline",
    "send application to": "apply_label",
    "job categories":      "category",
    "job types":           "site_type",
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


def _published_date(soup: BeautifulSoup) -> str:
    meta = soup.find("meta", attrs={"property": "article:published_time"})
    if meta and meta.get("content"):
        return meta["content"][:10]
    t = soup.find("time")
    if t and t.get("datetime"):
        return t["datetime"][:10]
    return str(date.today())


def _external_website(node, desc: str) -> str:
    blocked = ("careerone.vn", "google.com", "drive.google", "facebook.com",
               "linkedin.com", "twitter.com", "youtube.com", "cdn-cgi",
               "wp-content", "gravatar", "vxtgroup.com", "sr-company-attachments")
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

    h1 = soup.find("h1")
    title = h1.get_text(" ", strip=True) if h1 else ""
    title = title or listing_row.get("title", "")

    content = _main_content(soup)
    text = content.get_text("\n", strip=True)
    lines = [l for l in text.split("\n")]
    fields = _labelled_fields(lines)

    desc_text = text
    if title and title in desc_text:
        desc_text = desc_text.split(title, 1)[-1]
    desc_text = re.split(r"\n\s*Job Details\s*\n", desc_text, maxsplit=1)[0]

    # ── Run the full description cleaning pipeline ────────────────────────
    desc_text = clean_description(desc_text)

    emails = extract_cf_emails(soup)
    application = emails[0] if emails else ""
    for a in content.find_all("a", href=True):
        href = a["href"]
        if re.search(r"workday|greenhouse|lever\.co|bamboohr|smartrecruiters|/apply|careers?\.",
                     href, re.IGNORECASE):
            if not _BAD_APPLY.search(href):
                application = href.rstrip("/")
                break
    if not application:
        application = listing_row.get("apply", "")

    category  = fields.get("category", "")
    site_type = fields.get("site_type", "")
    org       = listing_row.get("org") or fields.get("org", "")
    location  = listing_row.get("location") or fields.get("location", "")
    deadline  = listing_row.get("deadline") or fields.get("deadline", "")
    website   = _external_website(content, desc_text)

    clean = clean_title(title, MAX_TITLE)
    co    = _enrich_company(org, website)

    return {
        "Job Title":          clean,
        "Job Type":           detect_job_type(site_type, title, desc_text),
        "Job Qualifications": extract_qualification(desc_text),
        "Job Experience":     extract_experience(desc_text),
        "Job Location":       location,
        "Job Field":          infer_job_field(title, category, desc_text) if not co["industry"] else co["industry"],
        "Date Posted":        _published_date(soup),
        "Deadline":           deadline,
        "Job Description":    normalize_text(desc_text[:6000]),
        "Application":        application,
        "Company URL":        co["website"] or website,
        "Company Name":       org,
        "Company Logo":       co["logo"],
        "Company Industry":   co["industry"] or category or "NGO / Development",
        "Company Founded":    co["founded"],
        "Company Type":       co["type"] or detect_company_type(org + " " + desc_text),
        "Company Website":    co["website"] or website,
        "Company Address":    location,
        "Company Details":    co["details"],
        "Job URL":            url,
        "Estimated Deadline": deadline,
        "Salary Range":       extract_salary(desc_text),
    }


# ════════════════════════════════════════════════════════════════════════════
# RAW-URL TRACKER  (original simple tracker for scrape deduplication)
# ════════════════════════════════════════════════════════════════════════════
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


# ════════════════════════════════════════════════════════════════════════════
# POST PIPELINE  (paraphrase + WP post for a single scraped record)
# ════════════════════════════════════════════════════════════════════════════
def post_record(row: pd.Series, idx: int,
                processed_ids: set, processed_urls: set,
                processed_companies: set) -> str:
    """
    Paraphrase and post one record to WordPress.
    Returns: "posted" | "skipped" | "failed"
    """
    title       = sanitize_text(str(row.get("Job Title", "")))
    desc        = sanitize_text(str(row.get("Job Description", "")))
    company     = sanitize_text(str(row.get("Company Name", "")))
    job_url     = sanitize_text(str(row.get("Job URL", "")), is_url=True)
    application = sanitize_text(str(row.get("Application", "")))
    job_id      = make_job_id(row, idx)

    print(f"\n{'═'*60}")
    print(f" Job ID: {job_id}")
    print(f" Title  : {title or '(empty)'}")
    print(f" Company: {company or '(empty)'}")
    print(f"{'═'*60}")

    if not title:
        print(" ⏭ SKIP — empty Job Title")
        return "skipped"
    if not desc:
        print(" ⏭ SKIP — empty Job Description")
        return "skipped"
    if not application:
        print(" ⏭ SKIP — empty Application")
        return "skipped"
    if job_id in processed_ids:
        print(" ⏭ SKIP — already processed")
        return "skipped"
    if job_url and job_url in processed_urls:
        print(" ⏭ SKIP — URL already processed")
        return "skipped"

    mark_read(job_id, job_url, title, company, idx)
    processed_ids.add(job_id)
    if job_url:
        processed_urls.add(job_url)

    if company and company not in processed_companies:
        company_data = {
            "company_name":     company,
            "company_logo":     sanitize_text(str(row.get("Company Logo", "")), is_url=True),
            "company_website":  sanitize_text(str(row.get("Company Website", "")), is_url=True),
            "company_industry": sanitize_text(str(row.get("Company Industry", ""))),
            "company_details":  sanitize_text(str(row.get("Company Details", ""))),
        }
        print(f"\n 🏢 Processing company: {company}")
        save_company(company_data)
        processed_companies.add(company)

    if ENABLE_PARAPHRASE:
        print(f"\n ✍️  Paraphrasing with Mistral…")
        new_title = paraphrase_title(title)
        new_desc  = paraphrase_description(desc)
        _upsert_row(job_id, {"Status": "paraphrased"})
    else:
        new_title = title
        new_desc  = desc

    if ENABLE_WP_POST:
        print(f"\n 📤 Posting to WordPress…")
        post_id, post_url = save_job(row, new_title, new_desc)
        if post_id:
            mark_posted(job_id, post_id, post_url or "")
            print(f" ✅ SUCCESS — WP ID={post_id} 🔗 {post_url}")
            return "posted"
        else:
            mark_failed(job_id, "wp_post_failed")
            print(f" ❌ WordPress post failed")
            return "failed"
    else:
        log.info(f"WP posting disabled — paraphrased only: {new_title}")
        return "posted"


# ════════════════════════════════════════════════════════════════════════════
# MAIN
# ════════════════════════════════════════════════════════════════════════════
def main():
    print("=" * 60)
    print("  CareerONE.vn Scraper + Paraphrase + WordPress Poster")
    print(f"  Started     : {datetime.now():%Y-%m-%d %H:%M:%S}")
    print(f"  Max pages   : {MAX_PAGES or 'until empty'}  ·  delay {DELAY}s")
    print(f"  Paraphrase  : {'ON' if ENABLE_PARAPHRASE else 'OFF'}")
    print(f"  WP posting  : {'ON' if ENABLE_WP_POST else 'OFF'}")
    print("=" * 60)

    # ── 1) Scrape listing pages ──────────────────────────────────────────
    processed_raw = load_processed()
    log.info(f"Already processed (raw tracker): {len(processed_raw)} jobs")

    listing_rows, page = [], 1
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

    # ── 2) Parse detail pages ────────────────────────────────────────────
    raw_records = []
    for i, row in enumerate(unique_rows, 1):
        if row["url"] in processed_raw:
            log.info(f"[{i}/{len(unique_rows)}] skip (done): {row['title']}")
            continue
        log.info(f"[{i}/{len(unique_rows)}] scraping: {row['title']}")
        rec = parse_detail(row["url"], row)
        if rec:
            raw_records.append(rec)
            processed_raw.add(row["url"])
        time.sleep(DELAY)

    if not raw_records:
        log.info("No new jobs scraped.")
    else:
        # ── 3) Write raw output (CSV + XLSX) ────────────────────────────
        new_df = pd.DataFrame(raw_records, columns=APPSCRIPT_COLUMNS)
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

        save_processed(processed_raw)

    # ── 4) Paraphrase + post to WordPress ───────────────────────────────
    if not ENABLE_WP_POST and not ENABLE_PARAPHRASE:
        log.info("Both WP posting and paraphrasing are OFF — done.")
        _summarise(raw_records)
        return

    # Load the full scraped CSV as the posting source
    if not os.path.exists(OUTPUT_CSV):
        log.warning(f"{OUTPUT_CSV} not found — nothing to post")
        _summarise(raw_records)
        return

    df_raw = pd.read_csv(OUTPUT_CSV)
    df = map_columns(df_raw)
    print_column_mapping(df)

    processed_ids, processed_urls = load_processed_ids()
    print(f"📋 {len(processed_ids)} jobs already in ID tracker.")
    print_tracker_summary()

    processed_companies: set = set()
    posted_count = skipped_count = failed_count = 0
    total = len(df)

    for idx, row in df.iterrows():
        result = post_record(
            row, idx,
            processed_ids, processed_urls, processed_companies
        )
        if result == "posted":
            posted_count += 1
        elif result == "skipped":
            skipped_count += 1
        else:
            failed_count += 1

        if (idx + 1) % 10 == 0:
            print(f"\n ⏸ Pausing 20s after every 10 jobs…")
            time.sleep(20)

    print(f"\n{'#'*60}")
    print(f" CYCLE COMPLETE  ({datetime.now():%Y-%m-%d %H:%M})")
    print(f" ✅ Posted  : {posted_count}")
    print(f" ⏭ Skipped : {skipped_count}")
    print(f" ❌ Failed  : {failed_count}")
    print(f" 📄 Total   : {total}")
    print(f"{'#'*60}")
    print_tracker_summary()
    _summarise(raw_records)


def _summarise(raw_records):
    print("=" * 60)
    print(f"  Scrape complete. New jobs this run: {len(raw_records)}")
    print(f"  Output: {OUTPUT_CSV}, {OUTPUT_XLSX}")
    print(f"  Finished: {datetime.now():%Y-%m-%d %H:%M:%S}")
    print("=" * 60)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        sys.exit(130)
