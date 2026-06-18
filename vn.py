#!/usr/bin/env python3
"""
careerone_scraper.py — CareerONE.vn job scraper with integrated data cleaning.

Cleaning applied at scrape time:
  1. Invisible / non-breaking characters (\xa0, zero-width, BOM, tabs) stripped
     from Job Title and Job Description.
  2. [Extension] tag stripped from titles; replaced with "(Deadline Extended)".
  3. Multi-role titles ("A; B; C") -> lead role + "(+N more roles)".
  4. Long RFQ/consultancy boilerplate titles shortened to meaningful scope.
  5. Salary Range: false positives like "USD5 million" rejected; VND ranges
     captured as "VND X - Y" instead of just lower bound.
  6. SmartRecruiters attachment URLs (sr-company-attachments) replaced with
     Job URL fallback — they are PDFs, not apply pages.
  7. Date Posted: backfilled with today's date where site omits publish time.
"""

import os
import re
import sys
import csv
import time
import math
import logging
from datetime import date, datetime

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
MAX_TITLE   = int(os.environ.get("CO_MAX_TITLE", "80"))

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
#  TITLE CLEANING
# ----------------------------------------------------------------------------
_WS_MAP = {
    "\xa0": " ", "\u2007": " ", "\u202f": " ", "\u2009": " ", "\u2002": " ",
    "\u2003": " ", "\t": " ", "\u200b": "", "\u200c": "", "\u200d": "",
    "\ufeff": "", "\u2028": " ", "\u2029": " ",
}
_QUOTES = {"\u201c": '"', "\u201d": '"', "\u2018": "'", "\u2019": "'"}

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
_BAD_APPLY     = re.compile(r"sr-company-attachments|cdn-cgi/l/email-protection(?!.*@)",
                             re.IGNORECASE)


def normalize_text(val) -> str:
    """Flatten invisible chars, normalise quotes, collapse whitespace."""
    if val is None:
        return ""
    t = str(val)
    for bad, good in _WS_MAP.items():
        t = t.replace(bad, good)
    for bad, good in _QUOTES.items():
        t = t.replace(bad, good)
    t = re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]", "", t)
    t = re.sub(r"[ \t]+", " ", t)
    t = re.sub(r" *\n *", "\n", t)
    t = re.sub(r"\n{3,}", "\n\n", t)
    return t.strip()


def _shrink(text: str, max_len: int) -> str:
    """Word-boundary truncate with ellipsis, preferring clause breaks."""
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
    """Return a tidy, sensibly-short title."""
    t = normalize_text(raw)
    if len(t) > 1 and t[0] in "\"'" and t[-1] == t[0]:
        t = t[1:-1].strip()

    # multi-role: "Role A; Role B; Role C" -> lead + count
    segs = [x.strip() for x in t.split(";") if x.strip()]
    if len(segs) >= 2 and all(2 < len(x) <= 60 for x in segs):
        extra = len(segs) - 1
        return f"{segs[0]} (+{extra} more role{'s' if extra > 1 else ''})"

    # strip [Extension] tag, note it at the end instead
    extension_flag = bool(_EXTENSION_TAG.match(t))
    if extension_flag:
        t = _EXTENSION_TAG.sub("", t).strip()

    # procurement/notice boilerplate: keep the meaningful scope
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
    ("Master's Degree",          ["master", "msc", "m.sc", "mba", "mphil", "postgraduate", "master of"]),
    ("Bachelor's Degree",        ["bachelor", "bsc", "b.sc", "b.a", "beng", "llb", "degree in",
                                   "university degree", "undergraduate", "honours"]),
    ("Diploma",                  ["diploma", "associate degree", "college degree"]),
    ("Professional Certification", ["acca", "cpa", "cfa", "cima", "pmp", "prince2", "chartered"]),
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


# Salary: ordered most-specific first.
# VND range must come before single VND to capture both bounds.
# USD requires comma-separated thousands or 4+ raw digits to avoid "USD5 million" false positives.
_SALARY_PATTERNS = [
    # VND range: "VND 579,293,760 to VND 888,245,530" (with possible \xa0)
    (r"VND[\s\xa0]*([\d,\.]+)[\s\xa0]*(?:to|-|–)[\s\xa0]*(?:VND[\s\xa0]*)?([\d,\.]+)", "vnd_range"),
    # standalone VND
    (r"VND[\s\xa0]*([\d,\.]+)", "vnd_single"),
    # USD with commas e.g. "USD 5,000" or "USD 5,000 - 8,000"
    (r"USD[\s\xa0]*([\d]{1,3}(?:,\d{3})+(?:[\s\xa0]*[-–][\s\xa0]*[\d,\.]+)?)", "usd"),
    # $ with commas
    (r"\$([\d]{1,3}(?:,\d{3})+(?:[\s\xa0]*[-–][\s\xa0]*\$?[\d,\.]+)?(?:[\s\xa0]*/[\s\xa0]*\w+)?)", "usd"),
    # "X,XXX / month" style
    (r"([\d,]+(?:[\s\xa0]*[-–][\s\xa0]*[\d,]+)?[\s\xa0]*/[\s\xa0]*(?:month|year|day|hour))", "rate"),
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
            val = f"VND {m.group(1).strip()}"
            return val
        val = m.group(0).strip().rstrip(".,")
        if re.search(r"\d{4,}|\d{1,3},\d{3}", val):
            return val
    return ""


# ----------------------------------------------------------------------------
#  ORG LOOKUP TABLE
# ----------------------------------------------------------------------------
ORG_DB = {
    "fhi 360": {
        "website":  "https://www.fhi360.org",
        "details":  "FHI 360 is a nonprofit human development organization dedicated to improving lives in lasting ways by advancing integrated, locally driven solutions. Working with a diverse mix of partners, FHI 360 serves more than 70 countries and all U.S. states and territories.",
        "founded":  "1971",
        "logo":     "https://www.fhi360.org/themes/custom/fhi360/logo.svg",
        "industry": "International Development / Public Health",
        "type":     "INGO",
    },
    "wwf": {
        "website":  "https://wwf.org",
        "details":  "WWF (World Wide Fund for Nature) is one of the world's largest and most respected independent conservation organizations. WWF's mission is to stop the degradation of the planet's natural environment and to build a future in which humans live in harmony with nature.",
        "founded":  "1961",
        "logo":     "https://wwf.org/wp-content/themes/wwf-rw/assets/images/header/wwf-logo.svg",
        "industry": "Environmental Conservation",
        "type":     "INGO",
    },
    "oxfam": {
        "website":  "https://www.oxfam.org",
        "details":  "Oxfam is a global movement of people who are fighting inequality to end poverty and injustice. The Oxfam confederation comprises 21 independent organizations working in 79 countries, collectively fighting poverty, supporting communities in disasters, and campaigning for systemic change.",
        "founded":  "1942",
        "logo":     "https://www.oxfam.org/themes/contrib/oxfam/logo.svg",
        "industry": "Humanitarian / Development",
        "type":     "INGO",
    },
    "care": {
        "website":  "https://www.care.org",
        "details":  "CARE is a leading humanitarian organization fighting global poverty. Founded in 1945, CARE places special focus on working alongside poor women and girls because, equipped with the proper resources, women have the power to lift whole families and entire communities out of poverty. CARE has been working in Vietnam since 1989.",
        "founded":  "1945",
        "logo":     "https://www.care.org/wp-content/themes/care2019/assets/images/logo.svg",
        "industry": "Humanitarian / Development",
        "type":     "INGO",
    },
    "wvi": {
        "website":  "https://www.wvi.org",
        "details":  "World Vision International is a Christian humanitarian organization dedicated to working with children, families, and their communities to reach their full potential by tackling the causes of poverty and injustice. World Vision serves all people, regardless of religion, race, ethnicity, or gender.",
        "founded":  "1950",
        "logo":     "https://www.wvi.org/sites/default/files/2019-10/WV_Logo_RGB.png",
        "industry": "Humanitarian / Development",
        "type":     "INGO",
    },
    "snv": {
        "website":  "https://www.snv.org",
        "details":  "SNV Netherlands Development Organisation is a mission-driven global development partner working in agriculture, energy, and water across more than 20 countries. SNV supports people to pursue bright futures by strengthening capacity and catalyzing partnerships that transform food, energy, and water systems.",
        "founded":  "1965",
        "logo":     "https://www.snv.org/themes/custom/snv/logo.svg",
        "industry": "International Development",
        "type":     "INGO",
    },
    "crs": {
        "website":  "https://www.crs.org",
        "details":  "Catholic Relief Services (CRS) is the official international humanitarian agency of the Catholic community in the United States. CRS works to save, protect, and transform lives in need in more than 100 countries, without regard to race, religion or nationality.",
        "founded":  "1943",
        "logo":     "https://www.crs.org/sites/default/files/crs-logo.png",
        "industry": "Humanitarian / Development",
        "type":     "INGO",
    },
    "plan international": {
        "website":  "https://plan-international.org",
        "details":  "Plan International is an independent development and humanitarian organization that advances children's rights and equality for girls. Founded in 1937, Plan International works in more than 70 countries to enable children and young people to be heard, to thrive, and to achieve their potential.",
        "founded":  "1937",
        "logo":     "https://plan-international.org/uploads/2022/01/Plan_International_logo.svg",
        "industry": "Children's Rights / Development",
        "type":     "INGO",
    },
    "giz": {
        "website":  "https://www.giz.de/en",
        "details":  "The Deutsche Gesellschaft für Internationale Zusammenarbeit (GIZ) GmbH is a federal enterprise supporting the German Government in achieving its objectives in the field of international cooperation for sustainable development. GIZ operates in more than 120 countries worldwide.",
        "founded":  "1975",
        "logo":     "https://www.giz.de/static/en/images/giz-logo.png",
        "industry": "International Development / Technical Cooperation",
        "type":     "Government / Embassy",
    },
    "unfpa": {
        "website":  "https://www.unfpa.org",
        "details":  "UNFPA, the United Nations Population Fund, is the United Nations sexual and reproductive health agency. UNFPA's mission is to deliver a world where every pregnancy is wanted, every childbirth is safe and every young person's potential is fulfilled.",
        "founded":  "1969",
        "logo":     "https://www.unfpa.org/sites/default/files/pub-pdf/UNFPA_logo_blue.png",
        "industry": "Sexual & Reproductive Health",
        "type":     "UN Agency",
    },
    "icraf": {
        "website":  "https://www.cifor-icraf.org",
        "details":  "CIFOR-ICRAF (Center for International Forestry Research and World Agroforestry) is a research center dedicated to transforming lives and landscapes through forest, tree and agroforestry science. It works in more than 50 countries to advance sustainable land use and improve livelihoods.",
        "founded":  "1978",
        "logo":     "https://www.cifor-icraf.org/wp-content/uploads/2021/06/CIFOR-ICRAF_logo.svg",
        "industry": "Forestry / Agroforestry Research",
        "type":     "Academic / Research",
    },
    "wcs": {
        "website":  "https://www.wcs.org",
        "details":  "The Wildlife Conservation Society (WCS) saves wildlife and wild places worldwide through science, conservation action, education, and inspiring people to value nature. WCS manages more than 500 conservation projects in 60+ countries and has maintained a presence in Vietnam since 1989.",
        "founded":  "1895",
        "logo":     "https://www.wcs.org/images/wcs-logo.png",
        "industry": "Wildlife Conservation",
        "type":     "NGO / Non-Profit",
    },
    "helvetas": {
        "website":  "https://www.helvetas.org",
        "details":  "HELVETAS is a Swiss organization for international cooperation committed to a just world in which all people determine the course of their lives and contribute to their societies. HELVETAS works in more than 30 countries across Africa, Asia, Latin America, and Eastern Europe.",
        "founded":  "1955",
        "logo":     "https://www.helvetas.org/typo3conf/ext/sitepackage/Resources/Public/Images/logo.svg",
        "industry": "International Development",
        "type":     "INGO",
    },
    "reach": {
        "website":  "https://www.reach-initiative.org",
        "details":  "REACH Initiative is a humanitarian assessment and information management organization. Co-founded by ACTED and IMPACT Initiatives, REACH supports humanitarian communities through information management, assessments, and research to enable more effective programming and advocacy.",
        "founded":  "2010",
        "logo":     "https://www.reach-initiative.org/wp-content/uploads/2019/05/reach-logo.png",
        "industry": "Humanitarian Information Management",
        "type":     "NGO / Non-Profit",
    },
    "rikolto": {
        "website":  "https://www.rikolto.org",
        "details":  "Rikolto is an international NGO with 40+ years of experience partnering with farmer organizations worldwide. Rikolto works across Africa, Asia, Latin America, and Europe to co-create fair and sustainable food systems through farmer-centered market development.",
        "founded":  "1976",
        "logo":     "https://www.rikolto.org/sites/default/files/rikolto_logo.png",
        "industry": "Agriculture / Food Systems",
        "type":     "INGO",
    },
    "chai": {
        "website":  "https://www.clintonhealthaccess.org",
        "details":  "The Clinton Health Access Initiative (CHAI) is a global health organization committed to saving lives and reducing the burden of disease in low- and middle-income countries. CHAI partners with governments to strengthen health systems by expanding access to medicines, diagnostics, and health services.",
        "founded":  "2002",
        "logo":     "https://www.clintonhealthaccess.org/content/uploads/2021/01/CHAI-logo.png",
        "industry": "Global Health",
        "type":     "NGO / Non-Profit",
    },
    "blue dragon": {
        "website":  "https://www.bluedragon.org",
        "details":  "Blue Dragon Children's Foundation is an Australian NGO working in Vietnam to rescue children from slavery and exploitation, provide education and support, and walk alongside street kids and vulnerable families to rebuild their lives.",
        "founded":  "2002",
        "logo":     "https://www.bluedragon.org/wp-content/uploads/2020/10/blue-dragon-logo.png",
        "industry": "Child Protection",
        "type":     "INGO",
    },
    "mwf": {
        "website":  "https://www.miral.ae",
        "details":  "Miral World (formerly Miral Asset Management) is the UAE's leading creator of immersive destinations and experiences, responsible for developing Yas Island Abu Dhabi's world-class leisure and entertainment portfolio.",
        "founded":  "2003",
        "logo":     "https://www.miral.ae/images/miral-logo.svg",
        "industry": "Leisure / Entertainment",
        "type":     "NGO / Development",
    },
    "aop": {
        "website":  "https://www.actiononpoverty.org",
        "details":  "Action on Poverty (AOP) is an Australian aid organization working in Vietnam and Southeast Asia to deliver practical solutions to poverty. AOP focuses on sustainable agriculture, clean water, education, and community enterprise development.",
        "founded":  "1968",
        "logo":     "https://www.actiononpoverty.org/wp-content/themes/aop/img/logo.png",
        "industry": "International Development",
        "type":     "INGO",
    },
    "ic vvaf": {  # normalized form of IC-VVAF
        "website":  "https://www.vvaf.org",
        "details":  "The Vietnam Veterans of America Foundation (VVAF) / International Center (IC) works on victims assistance and land mine/UXO clearance programs in Vietnam and other affected countries, improving the lives of victims of war and conflict.",
        "founded":  "1980",
        "logo":     "",
        "industry": "Humanitarian / UXO Clearance",
        "type":     "NGO / Non-Profit",
    },
    "ic-vvaf": {
        "website":  "https://www.vvaf.org",
        "details":  "The Vietnam Veterans of America Foundation (VVAF) / International Center (IC) works on victims assistance and land mine/UXO clearance programs in Vietnam and other affected countries, improving the lives of victims of war and conflict.",
        "founded":  "1980",
        "logo":     "",
        "industry": "Humanitarian / UXO Clearance",
        "type":     "NGO / Non-Profit",
    },
    "sci": {  # Save the Children International
        "website":  "https://www.savethechildren.org",
        "details":  "Save the Children International (SCI) is the world's leading independent organization for children. Founded in 1919, Save the Children works in around 100 countries to give children a healthy start in life, the opportunity to learn, and protection from harm.",
        "founded":  "1919",
        "logo":     "https://www.savethechildren.org/content/dam/usa/images/logos/save-the-children-logo.png",
        "industry": "Children's Rights / Humanitarian",
        "type":     "INGO",
    },
    "samaritan's purse": {
        "website":  "https://www.samaritanspurse.org",
        "details":  "Samaritan's Purse is an international Christian relief and development organization providing aid to victims of war, famine, disease, poverty, disasters, and persecution. Samaritan's Purse serves in more than 100 countries and is best known for Operation Christmas Child.",
        "founded":  "1970",
        "logo":     "https://www.samaritanspurse.org/images/sp-logo.png",
        "industry": "Humanitarian Relief / Development",
        "type":     "INGO",
    },
    "samaritans purse": {  # alternate key (apostrophe stripped)
        "website":  "https://www.samaritanspurse.org",
        "details":  "Samaritan's Purse is an international Christian relief and development organization providing aid to victims of war, famine, disease, poverty, disasters, and persecution. Samaritan's Purse serves in more than 100 countries and is best known for Operation Christmas Child.",
        "founded":  "1970",
        "logo":     "https://www.samaritanspurse.org/images/sp-logo.png",
        "industry": "Humanitarian Relief / Development",
        "type":     "INGO",
    },
    "epic project": {
        "website":  "https://www.epicproject.vn",
        "details":  "The EPIC (Engaging People in Conservation) Project is an environmental conservation initiative operating in Vietnam, working with local communities to protect biodiversity and natural resources through education, outreach, and community engagement.",
        "founded":  "",
        "logo":     "",
        "industry": "Environmental Conservation",
        "type":     "NGO / Non-Profit",
    },
    "msd": {
        "website":  "https://www.msdvietnam.org",
        "details":  "MSD Vietnam (Market Systems for Development) is an NGO that develops sustainable market systems and private sector engagement to create inclusive economic opportunities for vulnerable populations in Vietnam.",
        "founded":  "2011",
        "logo":     "",
        "industry": "Market Systems Development",
        "type":     "NGO / Non-Profit",
    },
    "koto": {
        "website":  "https://koto.com.au",
        "details":  "KOTO (Know One Teach One) is a social enterprise and training restaurant in Hanoi, Vietnam, providing disadvantaged youth with hospitality, life skills, and English training to help them build sustainable futures.",
        "founded":  "1999",
        "logo":     "https://koto.com.au/wp-content/uploads/2021/06/KOTO-Logo.png",
        "industry": "Vocational Training / Social Enterprise",
        "type":     "Social Enterprise",
    },
    "vvob": {
        "website":  "https://www.vvob.org",
        "details":  "VVOB (Education for Development) is a Belgian NGO that partners with governments, teachers, and education systems in the Global South to improve the quality of education, focusing on teacher professional development and school leadership.",
        "founded":  "1982",
        "logo":     "https://www.vvob.org/sites/default/files/vvob-logo.png",
        "industry": "Education / Development",
        "type":     "INGO",
    },
    "ide": {
        "website":  "https://www.ideglobal.org",
        "details":  "iDE (International Development Enterprises) is an international NGO that creates income and livelihood opportunities for poor rural households. iDE specializes in market-based approaches to agriculture, water, and sanitation in Asia and Africa.",
        "founded":  "1982",
        "logo":     "https://www.ideglobal.org/wp-content/uploads/2020/01/iDE-logo.png",
        "industry": "Agriculture / WASH",
        "type":     "INGO",
    },
    "mrc": {
        "website":  "https://www.mrcmekong.org",
        "details":  "The Mekong River Commission (MRC) is an inter-governmental organization that works directly with the governments of Cambodia, Laos, Thailand, and Vietnam to jointly manage the shared water resources and the sustainable development of the Mekong River.",
        "founded":  "1995",
        "logo":     "https://www.mrcmekong.org/assets/images/logo.png",
        "industry": "Water Resources / Environmental Governance",
        "type":     "Government / Embassy",
    },
    "aaf": {
        "website":  "https://www.asianaustralianfoundation.org",
        "details":  "The Asian Australian Foundation supports initiatives that strengthen ties between Australia and Asia, funding community development, education, and cultural exchange programs.",
        "founded":  "",
        "logo":     "",
        "industry": "Foundation / Philanthropy",
        "type":     "NGO / Non-Profit",
    },
    "aea": {
        "website":  "https://www.aea-international.org",
        "details":  "AEA International Holdings (now SOS International) is a global provider of medical assistance, security services, and emergency response to organizations and individuals in challenging environments worldwide.",
        "founded":  "1974",
        "logo":     "",
        "industry": "Medical Assistance / Security Services",
        "type":     "NGO / Development",
    },
    "amperes": {
        "website":  "https://www.amperes.com.vn",
        "details":  "AMPERES is a renewable energy and clean technology organization operating in Vietnam, working with local and international partners to accelerate the transition to sustainable energy solutions.",
        "founded":  "",
        "logo":     "",
        "industry": "Renewable Energy",
        "type":     "Social Enterprise",
    },
    "cov": {
        "website":  "https://www.caritas.org",
        "details":  "Caritas (COV) is an international confederation of Catholic charitable organizations working to fight poverty and promote human development in more than 200 countries and territories worldwide.",
        "founded":  "1897",
        "logo":     "https://www.caritas.org/wp-content/uploads/2020/06/caritas-logo.png",
        "industry": "Humanitarian / Development",
        "type":     "NGO / Non-Profit",
    },
    "env": {
        "website":  "https://www.env.org.vn",
        "details":  "Education for Nature – Vietnam (ENV) is a Vietnamese non-governmental organization focused on protecting Vietnam's wildlife and natural environment. ENV works to combat illegal wildlife trade, promote environmental education, and strengthen environmental governance.",
        "founded":  "2000",
        "logo":     "",
        "industry": "Wildlife Conservation / Environmental Education",
        "type":     "NGO / Non-Profit",
    },
    "sprint": {
        "website":  "https://www.sprintvietnam.org",
        "details":  "SPRINT (Social Protection Rights and Inclusive Networks in Training) is a development project operating in Vietnam, supporting social protection systems and inclusive programming for vulnerable groups.",
        "founded":  "",
        "logo":     "",
        "industry": "Social Protection / Development",
        "type":     "NGO / Development",
    },
    "streets": {
        "website":  "https://www.streetsinternational.org",
        "details":  "Streets International is a Hoi An-based NGO that provides disadvantaged Vietnamese youth with professional culinary and hospitality training, mentorship, and job placement to help them build sustainable careers and break the cycle of poverty.",
        "founded":  "2007",
        "logo":     "https://www.streetsinternational.org/wp-content/uploads/2020/01/Streets-Logo.png",
        "industry": "Vocational Training / Hospitality",
        "type":     "NGO / Non-Profit",
    },
}


def _norm_key(name: str) -> str:
    """Normalize org name for lookup: lowercase, collapse whitespace, strip punctuation."""
    n = name.lower().strip()
    n = re.sub(r"['\u2018\u2019\u201c\u201d]", "", n)   # strip smart quotes/apostrophes
    n = re.sub(r"[^\w\s]", " ", n)
    n = re.sub(r"\s+", " ", n).strip()
    return n

# ─────────────────────────────────────────────────────────────────────────────
#  WEB FETCH FALLBACK  (for orgs not in the lookup table)
# ─────────────────────────────────────────────────────────────────────────────
_SESS = requests.Session()
_SESS.headers.update({
    "User-Agent": "Mozilla/5.0 (compatible; enricher/1.0)",
    "Accept": "text/html,application/xhtml+xml;q=0.9,*/*;q=0.8",
})

_ABOUT_SLUGS = ["/about", "/about-us", "/who-we-are", "/about-us/", "/about/"]


def _enrich_company(name_raw: str, existing_web: str) -> dict:
    """Return enriched company fields from ORG_DB, or empty dict if unknown."""
    key = _norm_key(name_raw)
    info = ORG_DB.get(key, {})
    best_web = info.get("website", "")
    if not best_web:
        best_web = "" if is_ats_or_bad(existing_web) else existing_web
    return {
        "website": best_web,
        "logo":    info.get("logo", ""),
        "founded": info.get("founded", ""),
        "details": info.get("details", ""),
        "industry":info.get("industry", ""),
        "type":    info.get("type", ""),
    }


# ----------------------------------------------------------------------------
#  LISTING PARSER
# ----------------------------------------------------------------------------
def parse_listing(soup: BeautifulSoup) -> list:
    """Return list of dicts from the jobs table: title, url, org, deadline, location."""
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


# ----------------------------------------------------------------------------
#  DETAIL PARSER
# ----------------------------------------------------------------------------
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
    # site omits publish date — fall back to today as scrape-time marker
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

    title = ""
    h1 = soup.find(["h1"])
    if h1:
        title = h1.get_text(" ", strip=True)
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

    # Application: prefer decoded CF emails, then ATS apply links (not attachments)
    emails = extract_cf_emails(soup)
    application = emails[0] if emails else ""
    for a in content.find_all("a", href=True):
        href = a["href"]
        if re.search(r"workday|greenhouse|lever\.co|bamboohr|smartrecruiters|/apply|careers?\.",
                     href, re.IGNORECASE):
            if not _BAD_APPLY.search(href):          # skip attachment CDN links
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

    # Clean title at scrape time
    clean = clean_title(title, MAX_TITLE)

    # Enrich company data from built-in lookup table
    co = _enrich_company(org, website)

    record = {
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
    print(f"  Max title : {MAX_TITLE} chars")
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
