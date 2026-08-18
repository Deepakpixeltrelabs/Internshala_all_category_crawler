""" Scrapes internship listings from Internshala category listing pages
(e.g. https://internshala.com/internships/net-development-internship/),
then visits each individual internship detail page to pull the FULL
job description, "Who can apply" / "Other requirements" / "Perks" /
"Number of openings" / company info, and saves everything to an Excel
file.

Built with the same two-phase, fully-resumable design as the Coursesity
crawler you supplied:

    Phase 0 (category discovery): visits Internshala's category
        directory page (https://internshala.com/internships-by-category/)
        and collects every category listing URL
        (".../internships/<slug>-internship/"). You can also just hard-
        code the categories you care about in CATEGORY_URLS below and
        skip discovery entirely (see DISCOVER_ALL_CATEGORIES).

    Phase 1 (listing crawl): visits each category's paginated listing
        pages (Internshala uses plain "page-2/", "page-3/", ... URLs -
        no clicking required) and collects the basic card info (title,
        company, location, stipend, duration, posted date, short
        description, skills, job type, internship URL, internship ID).

    Phase 2 (detail crawl): visits every individual internship_url
        collected in Phase 1 and scrapes the FULL job description plus
        "Who can apply", "Other requirements", "Perks", "Number of
        openings", company website/about text, applicant count, and
        activity info. This is what gives you the full posting instead
        of just the listing-card blurb.

    Both phases are independently resumable (see RESUMING below).

BLOCKERS / VERIFICATION WALLS HANDLED:
    - Internshala renders these pages server-side (confirmed against the
      saved HTML you provided), so no special JS-rendering tricks are
      needed for normal pages - Selenium is used mainly for a realistic
      browser fingerprint and because it's the same proven approach as
      the Coursesity crawler.
    - The only client-side widget on these pages by default is an
      invisible reCAPTCHA v3 badge used for login/signup forms; it does
      not gate viewing listing or detail pages. classify_page() still
      watches for a *rendered* CAPTCHA/verification widget (g-recaptcha,
      h-captcha, cf-turnstile) actually replacing the real content, in
      case that ever changes or a specific IP/session gets challenged.
    - Every fetched page is classified as 'ok', 'soft_block' (a timing/
      auto-redirect JS challenge, Cloudflare "Attention Required",
      "unusual traffic", rate-limit message, a login-wall replacing the
      real content, etc.) or 'hard_block' (an actual interactive
      CAPTCHA widget rendered on the page):
        * soft_block  -> wait_out_soft_block() sits tight for a few
          rounds since these often clear themselves automatically,
          then re-checks before giving up.
        * hard_block  -> if MANUAL_SOLVE_ON_HARD_BLOCK = True and you
          run with HEADLESS = False, try_manual_solve() pauses the
          crawl and lets you solve it by hand in the visible browser
          window, then resumes automatically. Otherwise it's logged
          and skipped/retried like any other failure - this script
          never attempts to defeat a real CAPTCHA programmatically.
    - Session cookies persist to COOKIE_FILE across runs so repeat runs
      look like a returning browser instead of a brand-new anonymous
      session, and stealth JS patches (navigator.webdriver, plugins,
      languages, etc.) plus small randomized scrolling avoid the
      easy automated-browser fingerprint checks.
    - Pagination uses Internshala's real "isLastPage" hidden input
      (confirmed in the saved HTML) instead of guessing from a page
      count, so it won't stop early or loop forever.
    - All of this is best-effort and purely defensive (backing off,
      waiting, and optionally asking a human) - it does not attempt to
      circumvent deliberate anti-automation security measures.

INSTALL:
    pip install selenium webdriver-manager beautifulsoup4 pandas openpyxl

REQUIREMENTS:
    - Google Chrome (or Chromium) installed on your machine.
      webdriver-manager will download a matching chromedriver automatically.

RUN:
    python internshala_crawler.py

RESUMING:
    Everything is checkpointed to disk (see the *_FILE / DATA_DIR
    constants below). If the script is interrupted (Ctrl+C, network
    error, rate limiting, etc.) just re-run it - it picks up where it
    left off in both Phase 1 and Phase 2.
"""

import os
import re
import json
import time
import random
import logging
from dataclasses import dataclass, asdict, fields
from typing import Optional, List, Dict
from urllib.parse import urljoin

import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException
from webdriver_manager.chrome import ChromeDriverManager

logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(message)s")
log = logging.getLogger("internshala")

BASE_URL = "https://internshala.com"
CATEGORY_DIRECTORY_URL = "https://internshala.com/internships-by-category/"

# ----------------------------------------------------------------------
# WHICH CATEGORIES TO CRAWL
# ----------------------------------------------------------------------
# Internshala's category directory lists 250+ categories. Crawling all of
# them is possible (set DISCOVER_ALL_CATEGORIES = True) but will take a
# long time and produce a very large file. By default this script crawls
# only the categories listed in CATEGORY_URLS below (seeded with the
# .NET Development example you gave). Add more URLs from
# https://internshala.com/internships-by-category/ to this list, or flip
# DISCOVER_ALL_CATEGORIES to True to auto-crawl every category found on
# that directory page.

DISCOVER_ALL_CATEGORIES = True   # True = auto-discover & crawl EVERY category
MAX_CATEGORIES = None              # cap total categories when discovering (None = no cap)

CATEGORY_URLS = [
    "https://internshala.com/internships/net-development-internship/",
    # "https://internshala.com/internships/computer-science-internship/",
    # "https://internshala.com/internships/marketing-internship/",
    # ... add more category URLs here, or set DISCOVER_ALL_CATEGORIES = True
]

MAX_PAGES_PER_CATEGORY = None      # cap pages per category (None = all pages)
SCRAPE_FULL_DETAILS = True         # set False to skip Phase 2 entirely
DETAIL_MAX_INTERNSHIPS = None      # cap total internships visited in Phase 2 (testing); None = all
DETAIL_CATEGORY_FILTER = None      # e.g. {"Net Development"} to only fetch details for
                                    # those categories' internships; None = all

PAGE_LOAD_DELAY = 1.8               # base seconds, politeness delay between page loads
PAGE_LOAD_JITTER = 1.2              # + up to this many extra random seconds
LISTING_WAIT_TIMEOUT = 20           # seconds to wait for a listing page to render
DETAIL_WAIT_TIMEOUT = 15            # seconds to wait for a detail page to render
MAX_PAGE_RETRIES = 4                # retries for a stuck listing page
MAX_DETAIL_RETRIES = 3              # retries for a stuck internship detail page
RETRY_BACKOFF = 6                   # seconds, multiplied by attempt number

OUTPUT_FILE = "internshala_internships.xlsx"
DATA_DIR = "internshala_data"                              # per-category listing CSVs
PROGRESS_FILE = "internshala_progress.json"                 # listing progress (per category)
DETAILS_FILE = "internshala_details.csv"                    # detail data, keyed by internship_url
DETAIL_PROGRESS_FILE = "internshala_detail_progress.json"   # (unused directly - DETAILS_FILE doubles as progress)
HEADLESS = True                     # set False to watch the browser while debugging

# ----------------------------------------------------------------------
# ANTI-BLOCKING / VERIFICATION-WALL HANDLING
# ----------------------------------------------------------------------
COOKIE_FILE = "internshala_cookies.json"   # persisted session cookies, reused across runs
                                            # so Internshala sees a "returning browser"
                                            # instead of a fresh one every run
USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
]
CHALLENGE_WAIT_SECONDS = 12         # how long to just sit and let an automatic
                                     # JS/Cloudflare challenge resolve itself
                                     # before re-checking the page
MAX_CHALLENGE_WAIT_ROUNDS = 3       # how many CHALLENGE_WAIT_SECONDS rounds to try
MANUAL_SOLVE_ON_HARD_BLOCK = False  # if True and HEADLESS is False, pause and let a
                                     # human solve a real CAPTCHA in the visible window
                                     # instead of giving up immediately
MANUAL_SOLVE_TIMEOUT = 180          # seconds to wait for a human to clear a hard block


@dataclass
class Internship:
    # --- fields in the final normalized schema ---
    internship_id: str = ""            # Internshala's own internshipid attribute (stable, no need to invent one)
    title: str = ""
    company_name: str = ""
    internship_url: str = ""
    category: str = ""                 # which CATEGORY_URLS slug this was found under
    location: str = ""
    work_from_home: bool = False
    stipend: str = ""
    duration: str = ""
    start_date: str = ""
    apply_by: str = ""
    posted: str = ""                   # e.g. "3 days ago"
    job_type: str = ""                 # e.g. "Part time" (blank = full time / unspecified)
    applicants: str = ""
    skills_required: str = ""          # comma-joined
    description: str = ""              # short listing blurb, upgraded to full description in Phase 2
    who_can_apply: str = ""
    other_requirements: str = ""
    perks: str = ""                    # comma-joined
    number_of_openings: str = ""
    company_website: str = ""
    about_company: str = ""
    hiring_since: str = ""
    opportunities_posted: str = ""
    # --- internal-only fields, kept for resuming/debugging, dropped from
    #     the final exported file (see EXPORT_FIELD_ORDER / save()) ---
    detail_scraped: bool = False
    page_found_on: int = 1


COLUMN_RENAME = {
    "internship_id": "internshipId",
    "title": "title",
    "company_name": "companyName",
    "internship_url": "internshipUrl",
    "category": "category",
    "location": "location",
    "work_from_home": "workFromHome",
    "stipend": "stipend",
    "duration": "duration",
    "start_date": "startDate",
    "apply_by": "applyBy",
    "posted": "posted",
    "job_type": "jobType",
    "applicants": "applicants",
    "skills_required": "skillsRequired",
    "description": "description",
    "who_can_apply": "whoCanApply",
    "other_requirements": "otherRequirements",
    "perks": "perks",
    "number_of_openings": "numberOfOpenings",
    "company_website": "companyWebsite",
    "about_company": "aboutCompany",
    "hiring_since": "hiringSince",
    "opportunities_posted": "opportunitiesPosted",
}
EXPORT_FIELD_ORDER = list(COLUMN_RENAME.keys())


# ----------------------------------------------------------------------
# DRIVER SETUP
# ----------------------------------------------------------------------

def build_driver(headless: bool = True):
    opts = Options()
    if headless:
        opts.add_argument("--headless=new")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--window-size=1400,2200")
    opts.add_argument(f"user-agent={random.choice(USER_AGENTS)}")
    opts.add_argument("--lang=en-US,en")
    # Reduce obvious automation fingerprints
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option("useAutomationExtension", False)

    chrome_bin = os.environ.get("CHROME_BIN")
    if chrome_bin:
        opts.binary_location = chrome_bin

    driver_path = os.environ.get("CHROMEDRIVER_PATH")
    if driver_path and os.path.exists(driver_path):
        service = Service(driver_path)
    else:
        service = Service(ChromeDriverManager().install())

    driver = webdriver.Chrome(service=service, options=opts)

    # Stealth: patch the handful of navigator/window properties that
    # naive bot-detection scripts check for (webdriver flag, plugin
    # list length, languages, chrome.runtime presence, permissions API
    # quirk). None of this touches page content or bypasses a real
    # human-verification challenge - it just avoids getting flagged as
    # an automated browser for the easy, automatic checks.
    stealth_js = """
        Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
        Object.defineProperty(navigator, 'languages', {get: () => ['en-US', 'en']});
        Object.defineProperty(navigator, 'plugins', {get: () => [1, 2, 3, 4, 5]});
        window.chrome = window.chrome || { runtime: {} };
        const origQuery = window.navigator.permissions && window.navigator.permissions.query;
        if (origQuery) {
            window.navigator.permissions.query = (parameters) => (
                parameters.name === 'notifications'
                    ? Promise.resolve({ state: Notification.permission })
                    : origQuery(parameters)
            );
        }
    """
    try:
        driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {"source": stealth_js})
    except Exception:
        pass

    load_cookies(driver)
    return driver


def polite_delay():
    time.sleep(PAGE_LOAD_DELAY + random.uniform(0, PAGE_LOAD_JITTER))


def human_like_scroll(driver):
    """Small randomized scroll so the page fires lazy-load/scroll
    listeners the way a real visitor would, and so every page load
    doesn't look identical to a bot-detection script."""
    try:
        height = driver.execute_script("return document.body.scrollHeight") or 0
        if height <= 0:
            return
        for _ in range(random.randint(1, 3)):
            y = random.randint(0, height)
            driver.execute_script(f"window.scrollTo(0, {y});")
            time.sleep(random.uniform(0.2, 0.6))
    except Exception:
        pass


def save_cookies(driver):
    try:
        cookies = driver.get_cookies()
        save_json(COOKIE_FILE, cookies)
    except Exception as exc:
        log.debug("Could not save cookies: %s", exc)


def load_cookies(driver):
    """Loads a previously saved session so repeat runs look like a
    returning browser rather than a brand-new anonymous session (which
    is what usually trips rate-based bot walls). Cookies can only be
    added once we're already on the target domain, so this does a
    lightweight visit first."""
    if not os.path.exists(COOKIE_FILE):
        return
    cookies = load_json(COOKIE_FILE, [])
    if not cookies:
        return
    try:
        driver.get(BASE_URL)
        for cookie in cookies:
            cookie.pop("sameSite", None)  # selenium is picky about this field's values
            try:
                driver.add_cookie(cookie)
            except Exception:
                continue
        log.info("Restored %d cookies from previous session", len(cookies))
    except Exception as exc:
        log.debug("Could not restore cookies: %s", exc)


# ----------------------------------------------------------------------
# BLOCKER DETECTION
# ----------------------------------------------------------------------

BLOCKER_SIGNS = [
    "attention required",
    "checking your browser",
    "unusual traffic",
    "access denied",
    "just a moment",
    "are you a human",
    "please verify you are a human",
    "verify you are human",
    "ddos protection by",
    "sorry, you have been blocked",
    "request unsuccessful",
    "rate limit exceeded",
    "temporarily blocked",
    "enable javascript and cookies",
]

# Signs of a real, solvable interactive challenge (as opposed to a
# silent block) - if we see these we know sitting still and waiting
# won't help; it needs either a human or is simply not passable
# automatically.
HARD_CHALLENGE_SIGNS = [
    "g-recaptcha",
    "h-captcha",
    "cf-turnstile",
    "hcaptcha.com",
    "recaptcha/api2/anchor",
]


def classify_page(html: str, content_ok: Optional[bool] = None) -> str:
    """Returns 'ok', 'soft_block' (likely an automatic JS/timing
    challenge that resolves itself if we wait), or 'hard_block' (an
    interactive CAPTCHA/verification widget actually blocking the real
    content).

    `content_ok` lets the caller confirm the page-type-specific content
    it actually expects is present (e.g. internship cards on a listing
    page, the details container on a detail page, category links on
    the directory page). Internshala embeds a site-wide invisible
    reCAPTCHA v3 badge on every single page (including normal, working
    ones), so its mere presence in the HTML is NOT evidence of a block
    - only trust the HARD_CHALLENGE_SIGNS when the expected real
    content is also missing.
    """
    if content_ok:
        return "ok"

    lowered = html.lower()
    if any(sign in lowered for sign in BLOCKER_SIGNS):
        return "soft_block"
    if any(sign in lowered for sign in HARD_CHALLENGE_SIGNS):
        return "hard_block"
    if content_ok is False:
        # Caller already confirmed the expected content isn't there and
        # no explicit signs matched either - treat as a transient/soft
        # issue (empty results, slow render, etc.) rather than assuming
        # a hard block with no evidence for it.
        return "soft_block"
    if "internshala.com" not in lowered and "internship" not in lowered:
        return "soft_block"
    return "ok"


def detect_blocker(html: str) -> bool:
    """Backwards-compatible boolean wrapper around classify_page()."""
    return classify_page(html) != "ok"


def wait_out_soft_block(driver, get_html) -> Optional[str]:
    """When classify_page() says 'soft_block', many bot walls (timing
    checks, auto-redirecting JS challenges) clear themselves within a
    few seconds without any interaction. Sit tight and re-check a few
    times before treating it as a real block."""
    for round_num in range(1, MAX_CHALLENGE_WAIT_ROUNDS + 1):
        log.info("  soft block detected, waiting %ds for it to clear (round %d/%d)...",
                  CHALLENGE_WAIT_SECONDS, round_num, MAX_CHALLENGE_WAIT_ROUNDS)
        time.sleep(CHALLENGE_WAIT_SECONDS)
        html = get_html()
        if classify_page(html) == "ok":
            log.info("  block cleared on its own.")
            return html
    return None


def try_manual_solve(driver, get_html) -> Optional[str]:
    """Only runs when MANUAL_SOLVE_ON_HARD_BLOCK is True and the
    browser is visible (HEADLESS = False). Pauses the crawl and lets a
    human clear an actual CAPTCHA/verification widget in the open
    browser window, then continues automatically once the page looks
    clear. This does not attempt to solve anything programmatically."""
    if not (MANUAL_SOLVE_ON_HARD_BLOCK and not HEADLESS):
        return None
    log.warning("  hard block (interactive verification widget) detected. "
                "Please solve it in the browser window - waiting up to %ds...",
                MANUAL_SOLVE_TIMEOUT)
    deadline = time.time() + MANUAL_SOLVE_TIMEOUT
    while time.time() < deadline:
        time.sleep(3)
        html = get_html()
        if classify_page(html) == "ok":
            log.info("  verification cleared, resuming crawl.")
            save_cookies(driver)
            return html
    log.warning("  timed out waiting for manual verification.")
    return None


# ----------------------------------------------------------------------
# PHASE 0: CATEGORY DISCOVERY
# ----------------------------------------------------------------------

CATEGORY_HREF_RE = re.compile(r"/internships/[a-z0-9\-]+-internship/?$")
# Same pattern but usable directly against raw page HTML (matches up to
# the closing quote of the href attribute, since raw HTML obviously
# doesn't end right after the URL the way an isolated href string does).
CATEGORY_HREF_PROBE_RE = re.compile(r'/internships/[a-z0-9\-]+-internship/?["\']')


def discover_category_urls(driver) -> List[str]:
    log.info("Discovering categories from %s", CATEGORY_DIRECTORY_URL)
    driver.get(CATEGORY_DIRECTORY_URL)
    polite_delay()
    human_like_scroll(driver)
    html = driver.page_source
    status = classify_page(html, content_ok=bool(CATEGORY_HREF_PROBE_RE.search(html)))

    if status == "soft_block":
        html = wait_out_soft_block(driver, lambda: driver.page_source) or html
        status = classify_page(html, content_ok=bool(CATEGORY_HREF_PROBE_RE.search(html)))
    if status == "hard_block":
        html = try_manual_solve(driver, lambda: driver.page_source) or html
        status = classify_page(html, content_ok=bool(CATEGORY_HREF_PROBE_RE.search(html)))

    if status != "ok":
        log.error("Category directory page looks blocked (%s); falling back to CATEGORY_URLS as given.", status)
        return []
    save_cookies(driver)

    soup = BeautifulSoup(html, "html.parser")
    urls = set()
    for a in soup.find_all("a", href=True):
        href = a["href"]
        if CATEGORY_HREF_RE.search(href):
            urls.add(urljoin(BASE_URL, href).rstrip("/") + "/")

    urls = sorted(urls)
    if MAX_CATEGORIES:
        urls = urls[:MAX_CATEGORIES]
    log.info("Discovered %d category URLs", len(urls))
    return urls


def get_category_name(url: str) -> str:
    slug = url.rstrip("/").split("/")[-1]
    return slug.replace("-internship", "").replace("-", " ").title()


def category_csv_path(url: str) -> str:
    slug = url.rstrip("/").split("/")[-1]
    return os.path.join(DATA_DIR, f"{slug}.csv")


# ----------------------------------------------------------------------
# PHASE 1: LISTING PARSING
# ----------------------------------------------------------------------

def get_total_count(soup: BeautifulSoup) -> Optional[int]:
    """Reads the confirmed <h1>N Category Internships</h1> heading."""
    h1 = soup.find("h1")
    if not h1:
        return None
    m = re.match(r"\s*([\d,]+)", h1.get_text(strip=True))
    return int(m.group(1).replace(",", "")) if m else None


def is_last_page(soup: BeautifulSoup) -> bool:
    """Reads the confirmed <input id="isLastPage" value="0|1"> flag."""
    inp = soup.find("input", id="isLastPage")
    if inp and inp.get("value") is not None:
        return inp.get("value").strip() == "1"
    # No pagination control at all (e.g. only one page of results) also
    # means there's nothing further to fetch.
    return soup.select_one(".pagination_desktop") is None


def _row1_field(card, icon_classes: List[str]) -> str:
    for cls in icon_classes:
        icon = card.select_one(f".row-1-item i.{cls}")
        if icon:
            item = icon.find_parent("div", class_="row-1-item")
            if item:
                span = item.find("span")
                return span.get_text(" ", strip=True) if span else ""
    return ""


def parse_listing_page(html: str, category_name: str, page_num: int) -> List[Internship]:
    soup = BeautifulSoup(html, "html.parser")
    internships: List[Internship] = []

    container = soup.find("div", id="internship_list_container") or soup.find("div", id="internships_list_container")
    if container is None:
        return internships

    cards = container.find_all("div", class_="individual_internship", recursive=True)
    seen_ids = set()
    for card in cards:
        internship_id = card.get("internshipid", "")
        if internship_id in seen_ids:
            continue  # dedupe (mobile/desktop duplicate markup can double-count)
        seen_ids.add(internship_id)

        href = card.get("data-href", "")
        url = urljoin(BASE_URL, href) if href else ""

        title_el = card.select_one("h2.job-internship-name a") or card.select_one(".job-title-href")
        title = title_el.get_text(strip=True) if title_el else ""

        company_el = card.select_one(".company-name") or card.select_one(".company_name")
        company_name = company_el.get_text(strip=True) if company_el else ""

        location = _row1_field(card, ["ic-16-home", "ic-16-map-pin"])
        work_from_home = card.select_one(".row-1-item i.ic-16-home") is not None
        stipend = _row1_field(card, ["ic-16-money"])
        duration = _row1_field(card, ["ic-16-calendar"])

        desc_el = card.select_one(".about_job .text")
        description = desc_el.get_text(" ", strip=True) if desc_el else ""
        description = re.sub(r"^About the internship\s*", "", description).strip()

        skills = [s.get_text(strip=True) for s in card.select(".job_skills .job_skill")]

        posted_el = card.select_one(".detail-row-2 .color-labels span") or card.select_one(".status-inactive span, .status-info span")
        posted = posted_el.get_text(strip=True) if posted_el else ""

        job_type_el = card.select_one(".detail-row-2 .gray-labels .status-li span")
        job_type = job_type_el.get_text(strip=True) if job_type_el else ""

        internships.append(Internship(
            internship_id=internship_id,
            title=title,
            company_name=company_name,
            internship_url=url,
            category=category_name,
            location=location,
            work_from_home=work_from_home,
            stipend=stipend,
            duration=duration,
            posted=posted,
            job_type=job_type,
            skills_required=", ".join(skills),
            description=description,
            page_found_on=page_num,
        ))
    return internships


def build_page_url(category_url: str, page_num: int) -> str:
    base = category_url.rstrip("/") + "/"
    if page_num <= 1:
        return base
    return f"{base}page-{page_num}/"


def fetch_listing_page_with_retry(driver, url: str):
    """Returns (html, soup) or (None, None) after exhausting retries.
    Classifies every load as ok / soft_block / hard_block and handles
    each accordingly (see classify_page / wait_out_soft_block /
    try_manual_solve above) before falling back to a normal retry."""
    for attempt in range(1, MAX_PAGE_RETRIES + 1):
        try:
            driver.get(url)
            WebDriverWait(driver, LISTING_WAIT_TIMEOUT).until(
                EC.presence_of_element_located((By.TAG_NAME, "body"))
            )
            human_like_scroll(driver)
            html = driver.page_source

            def _listing_ok(h):
                return ("internship_list_container" in h or "internships_list_container" in h
                        or "isLastPage" in h or "pagination_desktop" in h)

            status = classify_page(html, content_ok=_listing_ok(html))

            if status == "soft_block":
                html = wait_out_soft_block(driver, lambda: driver.page_source) or html
                status = classify_page(html, content_ok=_listing_ok(html))

            if status == "hard_block":
                html = try_manual_solve(driver, lambda: driver.page_source) or html
                status = classify_page(html, content_ok=_listing_ok(html))

            if status == "ok":
                save_cookies(driver)
                return html, BeautifulSoup(html, "html.parser")

            log.warning("  attempt %d/%d: %s detected on %s",
                        attempt, MAX_PAGE_RETRIES, status, url)
        except (TimeoutException, WebDriverException) as exc:
            log.warning("  attempt %d/%d: error loading %s (%s)",
                        attempt, MAX_PAGE_RETRIES, url, exc)

        if attempt < MAX_PAGE_RETRIES:
            backoff = RETRY_BACKOFF * attempt
            log.info("  retrying in %.0fs...", backoff)
            time.sleep(backoff)
    return None, None


def crawl_category(driver, category_url: str, max_pages: Optional[int],
                    progress: Dict[str, dict]) -> List[Internship]:
    category_name = get_category_name(category_url)
    csv_path = category_csv_path(category_url)
    state = progress.get(category_url, {"last_completed_page": 0, "done": False})

    if state.get("done"):
        log.info("=== %s: already complete (skipping) ===", category_name)
        return load_internships_csv(csv_path)

    log.info("=== %s ===", category_name)
    internships = load_internships_csv(csv_path) if state["last_completed_page"] else []
    page = state["last_completed_page"] + 1
    total_count = None

    while True:
        url = build_page_url(category_url, page)
        html, soup = fetch_listing_page_with_retry(driver, url)
        if soup is None:
            log.warning("  giving up on %s at page %d after %d retries. "
                        "Progress is saved - just re-run the script and it "
                        "will resume from page %d.",
                        category_name, page, MAX_PAGE_RETRIES, page)
            break

        if total_count is None:
            total_count = get_total_count(soup)
            if total_count is not None:
                log.info("  %d total internships reported", total_count)

        page_internships = parse_listing_page(html, category_name, page)
        log.info("  page %d -> %d internships", page, len(page_internships))

        if not page_internships and page > 1:
            # Nothing left even though isLastPage wasn't set - stop cleanly.
            state["done"] = True
            progress[category_url] = state
            save_json(PROGRESS_FILE, progress)
            break

        internships.extend(page_internships)
        state.update({"last_completed_page": page})
        progress[category_url] = state
        save_json(PROGRESS_FILE, progress)
        save_internships_csv(internships, csv_path)

        reached_cap = bool(max_pages) and page >= max_pages
        if reached_cap or is_last_page(soup):
            state["done"] = True
            progress[category_url] = state
            save_json(PROGRESS_FILE, progress)
            break

        polite_delay()
        page += 1

    return internships


# ----------------------------------------------------------------------
# PHASE 2: DETAIL PARSING
# ----------------------------------------------------------------------

def _other_detail(soup, label: str) -> str:
    for item in soup.select(".other_detail_item"):
        heading = item.select_one(".item_heading span")
        if heading and label.lower() in heading.get_text(strip=True).lower():
            body = item.select_one(".item_body")
            if not body:
                return ""
            # The "Start Date" body duplicates its text in a hidden
            # mobile-only <span> and a desktop-only <span> (CSS shows
            # only one at a time) - drop the mobile copy so we don't
            # concatenate "Starts immediately Immediately".
            mobile_dup = body.select_one("[class*='mobile']")
            if mobile_dup:
                mobile_dup.extract()
            return body.get_text(" ", strip=True)
    return ""


def _section_after_heading(soup, text_fragment: str):
    """Finds a section_heading tag by (partial) text and returns the
    next sibling element - used only for headings without a unique
    class (e.g. 'Number of openings')."""
    for h in soup.select("h2.section_heading, h3.section_heading, p.section_heading"):
        if text_fragment.lower() in h.get_text(strip=True).lower():
            nxt = h.find_next_sibling()
            while nxt is not None and getattr(nxt, "name", None) is None:
                nxt = nxt.find_next_sibling()
            return nxt
    return None


def parse_internship_detail(html: str) -> dict:
    soup = BeautifulSoup(html, "html.parser")
    result = {
        "description": "",
        "who_can_apply": "",
        "other_requirements": "",
        "perks": "",
        "number_of_openings": "",
        "company_website": "",
        "about_company": "",
        "hiring_since": "",
        "opportunities_posted": "",
        "location": "",
        "stipend": "",
        "duration": "",
        "start_date": "",
        "apply_by": "",
        "posted": "",
        "applicants": "",
        "company_name": "",
        "skills_required": "",
    }

    container = soup.find("div", id="details_container")
    if container is None:
        return result

    # --- header/meta fields ---
    company_el = container.select_one(".company_name a") or container.select_one(".company_name")
    if company_el:
        result["company_name"] = company_el.get_text(strip=True)

    loc_el = container.select_one("#location_names")
    if loc_el:
        result["location"] = loc_el.get_text(" ", strip=True)

    result["start_date"] = _other_detail(container, "Start Date")
    result["duration"] = _other_detail(container, "Duration")
    result["stipend"] = _other_detail(container, "Stipend")
    result["apply_by"] = _other_detail(container, "Apply By")

    posted_el = container.select_one(".status-inactive, .status-info")
    if posted_el:
        result["posted"] = posted_el.get_text(" ", strip=True)

    applicants_el = container.select_one(".applications_message")
    if applicants_el:
        result["applicants"] = applicants_el.get_text(strip=True)

    # --- About the internship ---
    about_h = container.select_one("h2.about_heading")
    if about_h:
        desc_div = about_h.find_next_sibling("div", class_="text-container")
        if desc_div:
            result["description"] = desc_div.get_text("\n", strip=True)

    # --- Skills required ---
    skills_h = container.select_one("h3.skills_heading")
    if skills_h:
        skills_box = skills_h.find_next_sibling("div", class_="round_tabs_container")
        if skills_box:
            result["skills_required"] = ", ".join(
                s.get_text(strip=True) for s in skills_box.select(".round_tabs")
            )

    # --- Who can apply ---
    who_div = container.select_one(".text-container.who_can_apply")
    if who_div:
        result["who_can_apply"] = who_div.get_text("\n", strip=True)

    # --- Other requirements ---
    other_div = container.select_one(".text-container.additional_detail")
    if other_div:
        result["other_requirements"] = other_div.get_text("\n", strip=True)

    # --- Perks ---
    perks_h = container.select_one("h3.perks_heading")
    if perks_h:
        perks_box = perks_h.find_next_sibling("div", class_="round_tabs_container")
        if perks_box:
            result["perks"] = ", ".join(
                s.get_text(strip=True) for s in perks_box.select(".round_tabs")
            )

    # --- Number of openings (no unique class - match by heading text) ---
    openings_div = _section_after_heading(container, "Number of openings")
    if openings_div:
        result["number_of_openings"] = openings_div.get_text(strip=True)

    # --- Company info ---
    website_el = container.select_one(".company_info .website_link a")
    if website_el:
        result["company_website"] = website_el.get("href", "")

    about_company_div = container.select_one(".about_company_text_container")
    if about_company_div:
        result["about_company"] = about_company_div.get_text("\n", strip=True)

    activity_texts = [a.get_text(" ", strip=True) for a in container.select(".activity_container .activity .text")]
    for t in activity_texts:
        if "hiring since" in t.lower():
            result["hiring_since"] = t
        elif "opportunit" in t.lower():
            result["opportunities_posted"] = t

    return result


DETAIL_READY_SELECTOR = "#details_container"


def fetch_internship_detail_with_retry(driver, url: str) -> Optional[dict]:
    for attempt in range(1, MAX_DETAIL_RETRIES + 1):
        try:
            driver.get(url)
            WebDriverWait(driver, DETAIL_WAIT_TIMEOUT).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, DETAIL_READY_SELECTOR))
            )
            human_like_scroll(driver)
            html = driver.page_source
            status = classify_page(html, content_ok="details_container" in html)

            if status == "soft_block":
                html = wait_out_soft_block(driver, lambda: driver.page_source) or html
                status = classify_page(html, content_ok="details_container" in html)

            if status == "hard_block":
                html = try_manual_solve(driver, lambda: driver.page_source) or html
                status = classify_page(html, content_ok="details_container" in html)

            if status == "ok":
                save_cookies(driver)
                return parse_internship_detail(html)

            log.warning("  attempt %d/%d: %s detected on %s",
                        attempt, MAX_DETAIL_RETRIES, status, url)
        except (TimeoutException, WebDriverException) as exc:
            log.warning("  attempt %d/%d: error loading %s (%s)",
                        attempt, MAX_DETAIL_RETRIES, url, exc)

        if attempt < MAX_DETAIL_RETRIES:
            time.sleep(RETRY_BACKOFF * attempt)
    return None


def run_detail_phase(driver, all_internships: List[Internship]) -> None:
    if not SCRAPE_FULL_DETAILS:
        log.info("SCRAPE_FULL_DETAILS is False, skipping Phase 2.")
        return

    targets = all_internships
    if DETAIL_CATEGORY_FILTER:
        targets = [i for i in targets if i.category in DETAIL_CATEGORY_FILTER]

    seen_urls = set()
    unique_targets = []
    for i in targets:
        if i.internship_url and i.internship_url not in seen_urls:
            seen_urls.add(i.internship_url)
            unique_targets.append(i)
    targets = unique_targets

    if DETAIL_MAX_INTERNSHIPS:
        targets = targets[:DETAIL_MAX_INTERNSHIPS]

    details_map = load_details_map()
    done_urls = set(details_map.keys())
    remaining = [i for i in targets if i.internship_url not in done_urls]
    log.info("=== Phase 2: internship details === %d already done, %d remaining",
             len(done_urls & {i.internship_url for i in targets}), len(remaining))

    for idx, internship in enumerate(remaining, start=1):
        detail = fetch_internship_detail_with_retry(driver, internship.internship_url)
        if detail is None:
            log.warning("  [%d/%d] giving up on %s after %d retries. "
                        "Progress is saved - re-run the script to retry it.",
                        idx, len(remaining), internship.internship_url, MAX_DETAIL_RETRIES)
            continue

        append_detail_row(internship.internship_url, detail)
        details_map[internship.internship_url] = detail
        if idx % 25 == 0 or idx == len(remaining):
            log.info("  [%d/%d] internship details scraped", idx, len(remaining))
        polite_delay()

    for internship in all_internships:
        detail = details_map.get(internship.internship_url)
        if not detail:
            continue
        internship.detail_scraped = True
        if detail.get("description"):
            internship.description = detail["description"]
        for f in ("who_can_apply", "other_requirements", "perks", "number_of_openings",
                  "company_website", "about_company", "hiring_since", "opportunities_posted",
                  "applicants"):
            val = detail.get(f, "")
            if val:
                setattr(internship, f, val)
        # Prefer confirmed detail-page values, fall back to listing values
        for f in ("location", "stipend", "duration", "start_date", "apply_by", "posted",
                  "company_name", "skills_required"):
            val = detail.get(f, "")
            if val:
                setattr(internship, f, val)


# ----------------------------------------------------------------------
# PERSISTENCE HELPERS
# ----------------------------------------------------------------------

def load_json(path: str, default):
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    return default


def save_json(path: str, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)


def save_internships_csv(internships: List[Internship], path: str):
    df = pd.DataFrame([asdict(i) for i in internships])
    df.to_csv(path, index=False)


def load_internships_csv(path: str) -> List[Internship]:
    if not os.path.exists(path):
        return []
    df = pd.read_csv(path, keep_default_na=False)
    valid_fields = {f.name for f in fields(Internship)}
    internships = []
    for _, row in df.iterrows():
        kwargs = {k: row[k] for k in df.columns if k in valid_fields}
        if "work_from_home" in kwargs:
            kwargs["work_from_home"] = str(kwargs["work_from_home"]).strip().lower() in ("true", "1")
        if "detail_scraped" in kwargs:
            kwargs["detail_scraped"] = str(kwargs["detail_scraped"]).strip().lower() in ("true", "1")
        if "page_found_on" in kwargs:
            try:
                kwargs["page_found_on"] = int(kwargs["page_found_on"])
            except (ValueError, TypeError):
                kwargs["page_found_on"] = 1
        internships.append(Internship(**kwargs))
    return internships


def load_details_map() -> Dict[str, dict]:
    if not os.path.exists(DETAILS_FILE):
        return {}
    df = pd.read_csv(DETAILS_FILE, keep_default_na=False)
    out = {}
    for _, row in df.iterrows():
        out[row["internship_url"]] = {k: row[k] for k in df.columns if k != "internship_url"}
    return out


def append_detail_row(url: str, detail: dict):
    row = {"internship_url": url, **detail}
    df = pd.DataFrame([row])
    write_header = not os.path.exists(DETAILS_FILE)
    df.to_csv(DETAILS_FILE, mode="a", header=write_header, index=False)


def save(all_internships: List[Internship], path: str):
    df = pd.DataFrame([asdict(i) for i in all_internships])
    if df.empty:
        log.warning("Nothing scraped yet, skipping save of %s", path)
        return
    df.index = range(1, len(df) + 1)
    df.index.name = "#"
    df = df[EXPORT_FIELD_ORDER]
    df = df.rename(columns=COLUMN_RENAME)
    if path.endswith(".xlsx"):
        df.to_excel(path, sheet_name="Internships")
    else:
        df.to_csv(path)
    log.info("Saved %d rows -> %s", len(df), path)


# ----------------------------------------------------------------------
# MAIN
# ----------------------------------------------------------------------

def main():
    os.makedirs(DATA_DIR, exist_ok=True)
    progress = load_json(PROGRESS_FILE, {})

    driver = build_driver(headless=HEADLESS)
    all_rows: List[Internship] = []
    try:
        category_urls = CATEGORY_URLS
        if DISCOVER_ALL_CATEGORIES:
            discovered = discover_category_urls(driver)
            if discovered:
                category_urls = discovered
            else:
                log.warning("Discovery failed/blocked; falling back to CATEGORY_URLS list.")

        for url in category_urls:
            try:
                rows = crawl_category(driver, url, MAX_PAGES_PER_CATEGORY, progress)
                all_rows.extend(rows)
            except Exception as exc:
                log.error("Unexpected failure on %s: %s", url, exc)
                all_rows.extend(load_internships_csv(category_csv_path(url)))

        run_detail_phase(driver, all_rows)
    finally:
        save_cookies(driver)
        driver.quit()

    save(all_rows, OUTPUT_FILE)

    unfinished = [get_category_name(u) for u, s in progress.items() if not s.get("done")]
    if unfinished:
        log.info("Listing crawl NOT fully complete for: %s. Re-run to resume.",
                  ", ".join(unfinished))
    total_with_details = sum(1 for i in all_rows if i.detail_scraped)
    log.info("DONE. %d internships total, %d with full details scraped.",
              len(all_rows), total_with_details)


if __name__ == "__main__":
    main()
