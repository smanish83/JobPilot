#!/usr/bin/env python3
"""
JobPilot - Tailor your base resume to a job, write a matching cover letter,
and create a job requirements document, all as Word (.docx) files named after the job.

See README.md, SETUP.md and HOW_TO_USE.md for the full guides.

ONE-TIME SETUP
    Windows: double-click setup_windows.bat     Mac/Linux: ./setup_mac_linux.sh
    (installs everything, then runs the guided setup:  jobpilot init)
    Check it any time with:  jobpilot check

EXPERIENCE RULES (optional, one time)
    jobpilot rules --edit
    Opens ~/.resume_agent/experience_rules.json. Each rule switches on when the job posting
    mentions certain keywords, then replaces terms and adds your bullets/skills in a specific
    role. Example: if the job mentions Playwright, present your Robot Framework work at one employer
    as Playwright and add your Playwright bullets; if it mentions AWS/Docker/cloud, add your cloud
    bullets to the employers where you did that work. Run `jobpilot rules` to check which are ready.

HOW IT WORKS
    Three focused model calls: (1) analyze the posting against your resume, (2) tailor the
    resume using your base resume as the baseline, (3) write the cover letter from that tailored
    resume. Every reply is validated; a broken or incomplete reply is retried, never saved empty.
    Each resume then gets an ATS compliance check (keywords, title, contact, headings, dates,
    symbols, length, layout), shown in the terminal and in the requirements document.

APPLICATION TRACKER
    Job_Application_Tracker.xlsx is kept next to the applications folder. Every tailor/render run
    adds or refreshes that job's row; your own columns (Status, Date Applied, Follow-up Date,
    Contact, Next Step, Notes) are never overwritten. Catch up or rebuild all rows:  jobpilot track
    Needs openpyxl:  pip install openpyxl

FOR EACH JOB
    jobpilot tailor "https://www.linkedin.com/jobs/view/4012345678/"
    jobpilot tailor "https://www.indeed.com/viewjob?jk=abc123"
    jobpilot tailor job_description.txt
    jobpilot tailor --paste        (paste the description into the terminal)

SEVERAL JOBS AT ONCE
    jobpilot tailor "<link 1>" "<link 2>" "<link 3>"              (one after another)
    jobpilot tailor "<link 1>" "<link 2>" "<link 3>" --parallel 3 (three at the same time)
    jobpilot tailor --file jobs.txt --parallel 3                  (one link or file per line)
    A link that cannot be read is skipped (or you can paste it); the others still run.

    If a LinkedIn/Indeed page blocks the script, it asks you to paste the description instead.

OUTPUT  ->  applications/<YYYY-MM-DD>_<Company>_<Job_Title>/     (date = application date)
    Resume_<Company>_<Job_Title>.docx
    CoverLetter_<Company>_<Job_Title>.docx
    JobRequirements_<Company>_<Job_Title>.docx
                   - job details, responsibilities, required/preferred qualifications
                     with how you match each one, gaps, likely interview topics,
                     and the full original posting (kept in case it is taken down)
    content.json   - the generated text; edit it, then rebuild the Word files with:
                     jobpilot render "applications/<folder>/content.json"
"""
import argparse
import json
import threading
import os
import re
import shutil
import sys
from datetime import date
from pathlib import Path

__version__ = "2.3.0"
HOME = Path.home() / ".resume_agent"
SETTINGS_FILE = HOME / "settings.json"
DEFAULT_SETTINGS = {
    "work_authorization": "",            # e.g. "U.S. Work Authorized | No Visa Sponsorship Needed" ("" = hidden)
    "model": "claude-sonnet-5-5",        # Anthropic model used for all three steps
    "max_resume_pages": 2,               # longer resumes are trimmed automatically
    "shorten_role_first": "",            # company whose bullets go first when too long ("" = your oldest role)
    "drop_bachelors_if_higher_degree": True,  # when still too long and a Master's/PhD/MBA is listed
    "min_fit_score": 70,                 # skip jobs that match your real experience less than this (%); 0 = off
}


def load_settings() -> dict:
    settings = dict(DEFAULT_SETTINGS)
    if SETTINGS_FILE.exists():
        try:
            settings.update(json.loads(SETTINGS_FILE.read_text(encoding="utf-8")))
        except json.JSONDecodeError as err:
            sys.exit(f"Your settings file has a JSON error (line {err.lineno}): {SETTINGS_FILE}")
    # environment variables still work and take priority
    for prefix in ("JOBAPPLY_", "JOBPILOT_"):  # JOBAPPLY_* is the old name, still honored
        if prefix + "WORK_AUTH" in os.environ:
            settings["work_authorization"] = os.environ[prefix + "WORK_AUTH"]
        if os.environ.get(prefix + "MODEL"):
            settings["model"] = os.environ[prefix + "MODEL"]
    return settings


SETTINGS = load_settings()
MODEL = SETTINGS["model"]
_PRINT_LOCK = threading.Lock()
_TASK = threading.local()  # per-thread label such as "[2/3 Acme]" for parallel runs


def log(message: str = ""):
    """print() that is safe from several threads and prefixes the current job's label."""
    label = getattr(_TASK, "label", "")
    with _PRINT_LOCK:
        print(f"{label} {message}" if label and message else message, flush=True)
# Highlighted under your contact details on the resume and cover letter (settings: work_authorization).
WORK_AUTH = str(SETTINGS.get("work_authorization") or "").strip()
BASE_TEXT = HOME / "base_resume.txt"
OUTPUT_DIR = Path("applications")

# ---------------------------------------------------------------------------
# Instructions for the model (three focused calls: analyze job -> resume -> cover letter)
# ---------------------------------------------------------------------------
TRUTH_RULES = """Truthfulness rules (most important):
- NEVER invent employers, titles, dates, degrees, certifications, tools, projects or metrics
  that are not in the base resume or in <experience_rules>. Do not inflate seniority or numbers.
- <experience_rules>, when present, are additional verified facts written by the candidate.
  They override the base resume where they conflict: apply every term replacement and include
  every listed achievement in the named role (light rephrasing to mirror the posting is fine;
  keep the facts and numbers).
"""

ANALYZE_PROMPT = """You are an expert technical recruiter. Analyze the job posting against the
candidate's base resume (plus any <experience_rules>, which are verified facts).
Ignore website boilerplate (sign-in prompts, "show more", applicant counts, EEO statements).

Return ONLY one JSON object, no other text, exactly in this shape:
{
  "company": "hiring company",
  "job_title": "job title exactly as posted",
  "job_location": "location as posted",
  "job_requirements": {
    "overview": "2-3 sentence summary of the role",
    "work_arrangement": "Remote / Hybrid / On-site, or Not listed",
    "employment_type": "Full-time, Contract, W2... or Not listed",
    "experience_level": "e.g. 6+ years, or Not listed",
    "compensation": "as posted, or Not listed",
    "salary_min": "lowest base pay number in the range that applies (remote/candidate location, else the general or target range), or null",
    "salary_max": "highest base pay number of that same range, or null if no maximum is stated",
    "salary_period": "year or hour, or null",
    "sponsorship": "what the posting says about visa sponsorship / work authorization, or Not listed",
    "responsibilities": ["..."],
    "required": [{"requirement": "...", "match": "Meets|Partial|Gap", "evidence": "resume evidence or what is missing"}],
    "preferred": [{"requirement": "...", "match": "Meets|Partial|Gap", "evidence": "..."}],
    "tech_stack": ["ATS keywords: concise 1-4 word tools, technologies, methods and skills, spelled as posted"],
    "required_keywords": ["at most 15 of the tech_stack keywords that come from REQUIRED qualifications"],
    "ats_phrases": ["up to 12 short NON-technical phrases recruiters screen for, copied word for word from the posting"],
    "interview_focus": ["4-8 topics interviewers will likely probe"]
  },
  "matched_keywords": ["posting keywords the candidate genuinely has"],
  "gaps": ["requirements the candidate does not clearly meet"],
  "fit": {
    "score": 0-100,
    "role_family": "what kind of job this really is, e.g. 'Software QA automation (SDET)', 'ATM/banking hardware automation', 'Industrial/PLC automation', 'Robotic Process Automation (RPA) development', 'Executive leadership (VP)'",
    "seniority": "Match | Below candidate | Above candidate",
    "blockers": ["hard mismatches that make an interview very unlikely; empty list if none"],
    "reason": "one or two plain sentences explaining the score"
  }
}
Fit rules (fit decides whether a resume is written at all, so judge like a strict hiring manager):
- First decide what the job REALLY is from its duties, not from shared words. "Automation" or "testing"
  can mean industrial/PLC, robotics, ATM or kiosk hardware, network, marketing, RPA or manufacturing
  test work. If the job's core work is a different profession from the candidate's, that is a blocker.
  (Testing the SOFTWARE of such a product with the candidate's kind of tools is NOT a blocker.)
- Seniority: a VP, Head of, Director or other executive role is "Above candidate" and a blocker unless the
  resume shows comparable people-management scope; one level up (e.g. Lead, Staff, Manager of a small
  team) is not a blocker when the resume shows leadership evidence.
- Also blockers: a required license, security clearance, degree or certification the candidate lacks;
  a core primary language/platform the candidate has never used when the posting says it is mandatory;
  required years of experience far above the candidate's.
- score = the share of the job's core requirements the candidate's real experience covers. Weigh the
  main duties and required qualifications most; preferred ones count little. Meets = full, Partial = half.
  With any blocker, score must be 40 or lower. 85+ strong fit, 70-84 reasonable, below 70 unlikely.
- Be honest: do not inflate the score to be encouraging, and do not count self-directed learning
  (e.g. a Professional Development section) as work experience.
Keyword rules (these drive the ATS score, so keep them meaningful):
- tech_stack: at most 30 concise keywords an ATS would literally search for (e.g. "Playwright",
  "REST API", "CI/CD", "contract testing"). No sentences, soft skills, or business-domain nouns
  (e.g. not "Claims", "Billing", "Retries").
- required_keywords: at most 15, only from required qualifications. When the posting offers
  alternatives ("Jenkins, GitHub Actions, or GitLab CI"; "Java or Python"), include only ONE:
  the one the candidate has, or the first one if the candidate has none.
- ats_phrases: soft skills, collaboration and work-model, domain and responsibility phrases exactly as
  written in the posting, 2-6 words each, e.g. "team player", "onshore/offshore teams",
  "healthcare domain experience", "communicate testing status", "stakeholder management".
List every required and preferred qualification separately. "Meets" = clear evidence,
"Partial" = related but weaker evidence, "Gap" = no evidence.
"""

RESUME_PROMPT = TRUTH_RULES + """
You are an expert resume writer and ATS-optimization specialist. Rewrite the candidate's base
resume for the job, using the base resume as the baseline: keep its structure, every employer,
every title, every date, its education, and its strongest metrics.

Content rules:
- Reorder, rephrase, emphasize or trim content to fit the job; most relevant first.
- Keep every employer and date range. Recent roles 4-7 bullets, older roles 1-3 bullets.
- Bullets start with strong action verbs, stay under ~30 words, keep the original metrics.
- headline: the posting's job title, then 2-4 of the posting's top keywords,
  e.g. "QA Automation Engineer | Playwright, Cypress, REST API & CI/CD".
- summary: 3-4 sentences, includes the posting's job title and top keywords.
- Keep the whole resume to at most two pages (roughly 800-1000 words). If space is tight, shorten the
  oldest role to its title line first.
- Courses, training and self-directed learning from the base resume (e.g. a "Professional Development"
  section) go in professional_development, keeping their dates and "self-directed"/"learning" wording.
  NEVER present them as work done at an employer or move them into an experience entry, and do not
  list a skill that appears ONLY in that learning section under Core Skills either.

ATS rules (the resume MUST pass applicant tracking system screening):
- Use the posting's exact keyword spelling for every skill the candidate genuinely has
  (e.g. "CI/CD", "REST API", "Agile/Scrum"). Each such keyword must appear in core skills AND,
  where the candidate has used it, in at least one experience bullet.
- Spell out acronyms once next to the short form, e.g. "Behavior-Driven Development (BDD)".
- WORDING ALIGNMENT: for every phrase in <posting_phrases> and <posting_keywords>, check whether the base
  resume shows the same experience in other words. If it does, use the posting's exact words for it in the
  summary, core skills or the bullet that holds the evidence. Examples: base "worked with offshore QA teams"
  + posting "collaborates well with onshore/offshore teams" -> write "collaborated with onshore/offshore teams";
  base "Healthcare Systems" + posting "healthcare domain experience" -> write "healthcare domain experience";
  base "reported progress to management" + posting "communicate testing status to stakeholders" -> use
  "communicated testing status to stakeholders". Keep the facts, change only the wording.
  If the base resume has NO evidence for a phrase, do not use it anywhere.
- Plain text only: no tables, icons, emojis, columns, graphics, or special symbols.
- Dates as "Month YYYY - Month YYYY" (or "- Present"), same format everywhere.
- Do not put work authorization or visa text in the contact line or summary.

Return ONLY one JSON object, no other text, exactly in this shape:
{
  "name": "...",
  "headline": "...",
  "contact": "City, ST | phone | email | LinkedIn URL",
  "summary": "...",
  "skills": [{"category": "Test Automation Tools", "items": ["Playwright", "Selenium (WebDriver)"]}],
  "experience": [{"company": "...", "location": "City, ST (Remote/Hybrid)", "title": "...",
                  "dates": "Month YYYY - Month YYYY", "description": "optional one-line company/role context, or empty",
                  "bullets": ["..."]}],
  "education": ["Degree | School, City, ST | dates"],
  "certifications": [],
  "professional_development": ["Title (Self-Directed, YYYY): what was learned"]
}
"""

COVER_LETTER_PROMPT = """You are an expert career writer. Write a cover letter for the job, based ONLY on
the candidate's tailored resume (do not add any experience that is not in it).

Rules:
- 250-350 words, 3-4 body paragraphs, confident and specific, plain professional tone.
- Name the company and role, connect 2-3 real achievements (with their numbers) to the job's
  key requirements, and say why this company/role specifically (using only facts in the posting).
- Avoid cliches such as "I am writing to express my interest" and "I believe I would be a great fit".
- Address the hiring manager by name only if the posting names them; otherwise "Dear Hiring Manager,".
- No placeholders like [Company] or [Your Name].

Return ONLY one JSON object, no other text, exactly in this shape:
{"greeting": "Dear Hiring Manager,", "paragraphs": ["...", "...", "..."], "closing": "Sincerely,"}
"""
# ---------------------------------------------------------------------------
# Reading files
# ---------------------------------------------------------------------------
def extract_text(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix == ".docx":
        from docx import Document
        doc = Document(path)
        lines = [p.text for p in doc.paragraphs]
        for table in doc.tables:
            for row in table.rows:
                lines.append(" | ".join(cell.text.strip() for cell in row.cells))
        return "\n".join(line for line in lines if line.strip())
    if suffix == ".pdf":
        from pypdf import PdfReader
        return "\n".join(page.extract_text() or "" for page in PdfReader(path).pages)
    return path.read_text(encoding="utf-8")


# ---------------------------------------------------------------------------
# Getting the job description
# ---------------------------------------------------------------------------
HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/126.0 Safari/537.36"),
    "Accept-Language": "en-US,en;q=0.9",
}
BLOCKED_MARKERS = ("captcha", "verify you are human", "just a moment", "security check",
                   "access denied", "unusual activity")


def html_to_text(html: str) -> str:
    from bs4 import BeautifulSoup
    soup = BeautifulSoup(html, "html.parser")
    for tag in soup(["script", "style", "nav", "footer", "header", "noscript"]):
        tag.decompose()
    return re.sub(r"\n\s*\n+", "\n\n", soup.get_text("\n")).strip()


def fetch_linkedin(url: str, requests) -> str | None:
    """LinkedIn's public 'guest' job endpoint works without logging in for most postings."""
    match = re.search(r"currentJobId=(\d{6,})", url) or re.search(r"/jobs/view/(?:[^/?]*-)?(\d{6,})", url)
    if not match:
        return None
    api = f"https://www.linkedin.com/jobs-guest/jobs/api/jobPosting/{match.group(1)}"
    response = requests.get(api, headers=HEADERS, timeout=30)
    if not response.ok:
        return None
    from bs4 import BeautifulSoup
    soup = BeautifulSoup(response.text, "html.parser")
    description = soup.select_one(".show-more-less-html__markup, .description__text")
    if description is None:
        return html_to_text(response.text)
    header = [el.get_text(" ", strip=True) for el in
              soup.select(".top-card-layout__title, .topcard__org-name-link, .topcard__flavor--bullet")]
    criteria = [el.get_text(": ", strip=True) for el in soup.select(".description__job-criteria-item")]
    return "\n".join(header + [""] + [html_to_text(str(description))] + [""] + criteria)


class AlreadyDone(Exception):
    """This job (same company and job title) was already prepared or applied for."""
    def __init__(self, previous: dict):
        super().__init__(describe_previous(previous))
        self.previous = previous


class NotAFit(Exception):
    """The posting matches the candidate's real experience too little to be worth applying."""
    def __init__(self, analysis: dict, score: int):
        fit = analysis.get("fit") or {}
        super().__init__(f"fit {score}%: {fit.get('reason', '')}")
        self.analysis, self.score = analysis, score


def fit_score(analysis: dict) -> int | None:
    """The model's 0-100 fit score, capped at 40 when it listed blockers. None if missing."""
    fit = analysis.get("fit")
    if not isinstance(fit, dict):
        return None
    m = re.search(r"\d+(?:\.\d+)?", str(fit.get("score", "")))
    if not m:
        return None
    score = max(0, min(100, round(float(m.group()))))
    blockers = [b for b in fit.get("blockers") or [] if str(b).strip()]
    return min(score, 40) if blockers else score


def describe_fit(analysis: dict, score: int | None) -> str:
    fit = analysis.get("fit") or {}
    lines = [f"Role fit: {score}%" + (f" ({fit['role_family']})" if fit.get("role_family") else "")]
    if fit.get("seniority") and fit["seniority"] != "Match":
        lines.append(f"Seniority: {fit['seniority']}")
    for b in fit.get("blockers") or []:
        if str(b).strip():
            lines.append(f"Blocker: {b}")
    if fit.get("reason"):
        lines.append(f"Why: {fit['reason']}")
    return "\n  ".join(lines)


SKIP_LOG_NAME = "skipped_not_a_fit.csv"
SKIP_LOG_FIELDS = ["Date", "Company", "Job Title", "Location", "Fit %", "Role Type", "Reason", "Job Link"]


def skip_log_path() -> Path:
    return OUTPUT_DIR / SKIP_LOG_NAME


def log_not_a_fit(err: "NotAFit", source: str):
    """Remember skipped jobs, so the same link is not analyzed (and paid for) again."""
    import csv
    path = skip_log_path()
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        new = not path.exists()
        with path.open("a", newline="", encoding="utf-8-sig" if new else "utf-8") as f:
            w = csv.writer(f)
            if new:
                w.writerow(SKIP_LOG_FIELDS)
            a, fit = err.analysis, err.analysis.get("fit") or {}
            reason = "; ".join([str(b) for b in fit.get("blockers") or [] if str(b).strip()] + [fit.get("reason", "")])
            w.writerow([date.today().isoformat(), a.get("company", ""), a.get("job_title", ""),
                        a.get("job_location", ""), err.score, fit.get("role_family", ""), reason.strip("; "), source])
    except OSError as e:  # e.g. open in Excel
        print(f"  Note: could not record it in {path} ({e}).")


def previous_skips() -> dict:
    """Links skipped earlier as not a fit, keyed by job_key."""
    import csv
    path = skip_log_path()
    if not path.exists():
        return {}
    skips = {}
    try:
        with path.open(newline="", encoding="utf-8-sig") as f:
            for row in csv.DictReader(f):
                key = job_key(row.get("Job Link", ""))
                if key:
                    skips[key] = row
    except (OSError, csv.Error):
        pass
    return skips


class JobInputError(Exception):
    """A job posting could not be read (blocked page, skipped paste, missing file)."""


def read_pasted(reason: str = "", allow_skip: bool = False) -> str:
    if reason:
        print(reason)
    eof_key = "Ctrl+Z then Enter" if sys.platform.startswith("win") else "Ctrl+D"
    skip = " (or press it right away to skip this job)" if allow_skip else ""
    print(f"Paste the job description below, then press {eof_key} on a new line{skip}:\n")
    text = sys.stdin.read().strip()
    if allow_skip and not text:
        raise JobInputError("no description pasted")
    if len(text) < 200:
        raise JobInputError("the pasted job description was too short")
    return text


def fetch_job(url: str, allow_skip: bool = False) -> str:
    import requests
    text = None
    try:
        if "linkedin.com" in url:
            text = fetch_linkedin(url, requests)
        if not text:
            response = requests.get(url, headers=HEADERS, timeout=30)
            if response.ok:
                text = html_to_text(response.text)
    except requests.RequestException as err:
        print(f"Could not load the page ({err}).")

    if not text or len(text) < 400 or any(m in text.lower()[:3000] for m in BLOCKED_MARKERS):
        return read_pasted(f"The job site blocked automatic reading of {url}\n"
                           "(common for Indeed and some LinkedIn pages).", allow_skip=allow_skip)
    return text[:20000]


# ---------------------------------------------------------------------------
# The agent call
# ---------------------------------------------------------------------------
def _parse_json_from_text(text: str) -> dict | None:
    """Fallback: pull a JSON object out of a plain-text reply."""
    text = re.sub(r"^```(?:json)?\s*|\s*```$", "", text.strip())
    start, end = text.find("{"), text.rfind("}")
    if start == -1 or end <= start:
        return None
    try:
        return json.loads(text[start:end + 1])
    except json.JSONDecodeError:
        return None


# ---------------------------------------------------------------------------
# Normalizing the model's output (models sometimes return a slightly different shape,
# e.g. a list as JSON text, skills as plain strings, or bullets as one block of text)
# ---------------------------------------------------------------------------
def _decode(value):
    """Turn JSON text like '[...]' or '{...}' into real lists/dicts."""
    if isinstance(value, str) and value.strip()[:1] in "[{":
        try:
            return json.loads(value)
        except json.JSONDecodeError:
            pass
    return value


def _str_list(value) -> list[str]:
    value = _decode(value)
    if value is None:
        return []
    if isinstance(value, str):
        parts = re.split(r"\n+", value) if "\n" in value else [value]
        return [re.sub(r"^\s*[-*•]\s*", "", p).strip() for p in parts if p.strip()]
    if isinstance(value, dict):
        return [f"{k}: {v}" for k, v in value.items()]
    out = []
    for item in value:
        if isinstance(item, dict):
            out.append(" - ".join(str(v) for v in item.values() if v))
        elif item is not None and str(item).strip():
            out.append(str(item).strip())
    return out


def _skills(value) -> list[dict]:
    value = _decode(value)
    if not value:
        return []
    if isinstance(value, str):
        value = [value]
    if isinstance(value, dict):  # {"Languages": ["Python", ...], ...}
        return [{"category": str(k), "items": _split_items(v)} for k, v in value.items()]
    groups, loose = [], []
    for item in value:
        item = _decode(item)
        if isinstance(item, dict):
            category = item.get("category") or item.get("name") or item.get("title") or "Skills"
            items = item.get("items", item.get("skills", item.get("values", [])))
            groups.append({"category": str(category), "items": _split_items(items)})
        elif isinstance(item, str) and ":" in item:  # "Languages: Python, Java"
            category, _, rest = item.partition(":")
            groups.append({"category": category.strip(), "items": _split_items(rest)})
        elif item:
            loose.append(str(item).strip())
    if loose:
        groups.append({"category": "Skills", "items": loose})
    return groups


def _split_items(value) -> list[str]:
    value = _decode(value)
    if isinstance(value, str):
        return [s.strip() for s in value.split(",") if s.strip()]
    return _str_list(value)


def _requirement_items(value) -> list[dict]:
    value = _decode(value) or []
    if isinstance(value, (str, dict)):
        value = [value]
    out = []
    for item in value:
        item = _decode(item)
        if isinstance(item, dict):
            match = str(item.get("match", "")).strip().capitalize()
            out.append({"requirement": str(item.get("requirement") or item.get("name") or ""),
                        "match": match if match in ("Meets", "Partial", "Gap") else match or "-",
                        "evidence": str(item.get("evidence") or "")})
        elif item:
            out.append({"requirement": str(item), "match": "-", "evidence": ""})
    return out


def normalize_data(data: dict) -> dict:
    data = {k: _decode(v) for k, v in data.items()}
    for key in ("company", "job_title", "name", "headline", "contact", "summary", "job_location"):
        if isinstance(data.get(key), list):
            data[key] = " | ".join(map(str, data[key]))
        data[key] = str(data.get(key) or "").strip()
    # Work authorization is shown on its own line; keep it out of the contact line.
    parts = [p.strip() for p in data["contact"].split("|")]
    data["contact"] = " | ".join(p for p in parts if p and not re.search(r"sponsor|visa|work auth", p, re.I))
    data["company"] = data["company"] or "Company"
    data["job_title"] = data["job_title"] or "Job"

    data["skills"] = _skills(data.get("skills"))

    experience = data.get("experience") or []
    if isinstance(experience, dict):
        experience = [experience]
    jobs = []
    for job in experience:
        job = _decode(job)
        if not isinstance(job, dict):
            continue
        jobs.append({"title": str(job.get("title") or job.get("role") or ""),
                     "company": str(job.get("company") or ""),
                     "location": str(job.get("location") or ""),
                     "dates": str(job.get("dates") or job.get("date") or ""),
                     "description": str(job.get("description") or ""),
                     "bullets": _str_list(job.get("bullets") or job.get("highlights"))})
    data["experience"] = jobs

    for key in ("education", "certifications", "professional_development", "matched_keywords", "gaps"):
        data[key] = _str_list(data.get(key))

    letter = _decode(data.get("cover_letter")) or {}
    if isinstance(letter, str):
        letter = {"paragraphs": [p for p in re.split(r"\n\s*\n", letter) if p.strip()]}
    data["cover_letter"] = {"greeting": str(letter.get("greeting") or "Dear Hiring Manager,"),
                            "paragraphs": _str_list(letter.get("paragraphs") or letter.get("body")),
                            "closing": str(letter.get("closing") or "Sincerely,")}

    req = _decode(data.get("job_requirements")) or {}
    if not isinstance(req, dict):
        req = {}
    req = {k: _decode(v) for k, v in req.items()}
    for key in ("overview", "work_arrangement", "employment_type", "experience_level", "compensation",
                "sponsorship"):
        req[key] = str(req.get(key) or "").strip()
    for key in ("responsibilities", "tech_stack", "required_keywords", "ats_phrases", "interview_focus"):
        req[key] = _str_list(req.get(key))
    for key in ("salary_min", "salary_max"):
        req[key] = _to_number(req.get(key))
    req["salary_period"] = req.get("salary_period") if req.get("salary_period") in ("year", "hour") else None
    req["required"] = _requirement_items(req.get("required"))
    req["preferred"] = _requirement_items(req.get("preferred"))
    data["job_requirements"] = req
    return data


JUNK_MARKERS = ("<parameter", "</parameter", "<invoke", "antml:", "placeholder\"")
LAST_REPLY = HOME / "last_model_reply.txt"


def _has_junk(value) -> bool:
    text = json.dumps(value) if not isinstance(value, str) else value
    return any(marker in text for marker in JUNK_MARKERS)


def _ask_json(client, system: str, prompt: str, check, label: str, max_tokens: int = 12000) -> dict:
    """One model call that must return a JSON object passing `check` (returns an error text or None).
    Retries once with feedback; never returns empty or broken content."""
    messages = [{"role": "user", "content": prompt}]
    problem = "no reply"
    for attempt in range(2):
        with client.messages.stream(model=MODEL, max_tokens=max_tokens, system=system,
                                    messages=messages) as stream:
            response = stream.get_final_message()
        text = "".join(b.text for b in response.content if b.type == "text")
        HOME.mkdir(parents=True, exist_ok=True)
        LAST_REPLY.write_text(f"--- {label} (attempt {attempt + 1}) ---\n{text}", encoding="utf-8")
        if response.stop_reason == "max_tokens":
            problem = "the reply was cut off because it was too long"
        else:
            data = _parse_json_from_text(text)
            if data is None:
                problem = "the reply was not a valid JSON object"
            elif _has_junk(data):
                problem = "the reply contained broken markup instead of resume text"
            else:
                problem = check(data)
                if not problem:
                    return data
        log(f"  {label}: {problem}. Retrying...")
        messages = [{"role": "user", "content": prompt},
                    {"role": "assistant", "content": text or "(empty)"},
                    {"role": "user", "content": f"That reply had a problem: {problem}. Return ONLY the complete, "
                                                "corrected JSON object in the required shape, with no other text."}]
    raise RuntimeError(f"Could not get a valid {label} from the model ({problem}). Nothing was saved. "
                       f"The last reply is in {LAST_REPLY}. Please run the command again.")


def _check_analysis(d: dict):
    if not d.get("company") or not d.get("job_title"):
        return "company or job_title is missing"
    req = d.get("job_requirements")
    if not isinstance(req, dict) or not req.get("tech_stack") or not (req.get("required") or req.get("responsibilities")):
        return "job_requirements is missing tech_stack, required or responsibilities"
    return None


def _check_resume(d: dict):
    exp = d.get("experience")
    if not d.get("name") or not d.get("summary"):
        return "name or summary is missing"
    if not isinstance(exp, list) or not exp or not all(isinstance(j, dict) and j.get("bullets") for j in exp):
        return "experience is empty or a role has no bullets"
    if not d.get("skills"):
        return "skills are missing"
    if not d.get("education"):
        return "education is missing"
    return None


def _check_letter(d: dict):
    paragraphs = d.get("paragraphs")
    if not isinstance(paragraphs, list) or len([p for p in paragraphs if str(p).strip()]) < 3:
        return "the letter needs at least 3 body paragraphs"
    return None


def generate(base_resume: str, job_text: str, rules_text: str = "", after_analysis=None) -> dict:
    import anthropic
    client = anthropic.Anthropic()  # uses ANTHROPIC_API_KEY
    shared = (f"<base_resume>\n{base_resume}\n</base_resume>\n\n{rules_text}"
              f"<job_posting>\n{job_text}\n</job_posting>\n\n")

    log("  1/3 Analyzing the job posting...")
    analysis = _ask_json(client, ANALYZE_PROMPT, shared + "Analyze this job against the resume.",
                         _check_analysis, "job analysis")
    if after_analysis:
        after_analysis(analysis)  # may raise AlreadyDone: stops before the resume and cover letter

    log("  2/3 Tailoring the resume...")
    keywords = ", ".join(analysis["job_requirements"].get("tech_stack", []))
    phrases = " | ".join(analysis["job_requirements"].get("ats_phrases", []))
    resume = _ask_json(client, RESUME_PROMPT,
                       shared + f"<posting_keywords>{keywords}</posting_keywords>\n"
                       f"<posting_phrases>{phrases}</posting_phrases>\n\n"
                       "Rewrite the base resume for this job. Use every posting keyword the candidate "
                       "genuinely has, spelled exactly as in <posting_keywords>, and align the wording to "
                       "<posting_phrases> wherever the base resume shows that experience.",
                       _check_resume, "tailored resume", max_tokens=16000)

    log("  3/3 Writing the cover letter...")
    auth = ""
    if WORK_AUTH:
        auth = (f"<work_authorization>{WORK_AUTH}</work_authorization>\nIn the final paragraph, state this "
                "in one short, confident sentence (e.g. 'I do not require visa sponsorship.').\n\n")
    letter = _ask_json(client, COVER_LETTER_PROMPT,
                       f"<tailored_resume>\n{json.dumps(resume, indent=1)}\n</tailored_resume>\n\n{auth}"
                       f"<job_posting>\n{job_text}\n</job_posting>\n\nWrite the cover letter.",
                       _check_letter, "cover letter", max_tokens=4000)

    return {**analysis, **resume, "cover_letter": letter}


# ---------------------------------------------------------------------------
# Word document builders
# ---------------------------------------------------------------------------
def _docx_imports():
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Inches, Pt, RGBColor
    return Document, WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, OxmlElement, qn, Inches, Pt, RGBColor


Document, WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, OxmlElement, qn, Inches, Pt, RGBColor = _docx_imports()
ACCENT = RGBColor(0x1F, 0x3A, 0x5F)
GRAY = RGBColor(0x44, 0x44, 0x44)
TEXT_WIDTH = Inches(7.3)  # 8.5" page minus 0.6" margins


def new_document(font_size: float):
    doc = Document()
    section = doc.sections[0]
    section.page_width, section.page_height = Inches(8.5), Inches(11)
    section.left_margin = section.right_margin = Inches(0.6)
    section.top_margin = section.bottom_margin = Inches(0.5)
    normal = doc.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(font_size)
    normal.element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(0)
    normal.paragraph_format.line_spacing = 1.05
    heading = doc.styles["Heading 1"]
    heading.font.name, heading.font.size, heading.font.bold = "Calibri", Pt(12), True
    heading.font.color.rgb = ACCENT
    fonts = heading.element.rPr.rFonts
    for attr in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
        fonts.attrib.pop(qn(attr), None)  # use Calibri, not the theme's heading font
    return doc


def set_properties(doc, **values):
    """Word limits each document property (title, author, keywords) to 255 characters."""
    for name, value in values.items():
        value = clean_xml_text(value).strip()
        if len(value) > 255:
            cut = value[:255]
            value = cut[:cut.rfind(",")] if "," in cut else cut  # end on a whole keyword
        setattr(doc.core_properties, name, value)


def add_name_header(doc, name: str, contact: str, headline: str = ""):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(name)
    run.bold, run.font.size, run.font.color.rgb = True, Pt(20), ACCENT
    if headline:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(1)
        run = p.add_run(headline)
        run.bold, run.font.size = True, Pt(11.5)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(0 if WORK_AUTH else 4)
    run = p.add_run(contact)
    run.font.size, run.font.color.rgb = Pt(9.5), GRAY
    if WORK_AUTH:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(4)
        run = p.add_run(WORK_AUTH)
        run.bold, run.font.size, run.font.color.rgb = True, Pt(10), ACCENT


def add_section_heading(doc, title: str):
    """Real Word 'Heading 1' paragraphs: ATS parsers use them to find sections."""
    p = doc.add_paragraph(style="Heading 1")
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.keep_with_next = True
    run = p.add_run(title)
    run.bold, run.font.size, run.font.color.rgb = True, Pt(12), ACCENT
    border = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    for key, value in (("w:val", "single"), ("w:sz", "6"), ("w:space", "1"), ("w:color", "1F3A5F")):
        bottom.set(qn(key), value)
    border.append(bottom)
    p._p.get_or_add_pPr().append(border)


RESUME_SECTIONS = {"summary": "Professional Summary", "skills": "Core Skills",
                   "experience": "Professional Experience",
                   "professional_development": "Professional Development",
                   "education": "Education", "certifications": "Certifications"}


def build_resume(data: dict, path: Path):
    """Layout follows the base resume: name, headline, contact, then standard sections.
    Single column, no tables/text boxes/images, contact info in the body (not the Word header)."""
    doc = new_document(10.5)
    add_name_header(doc, data["name"], data["contact"], data.get("headline", ""))

    add_section_heading(doc, RESUME_SECTIONS["summary"])
    doc.add_paragraph(data["summary"])

    add_section_heading(doc, RESUME_SECTIONS["skills"])
    for group in data.get("skills", []):
        p = doc.add_paragraph(style="List Bullet")
        p.paragraph_format.space_after = Pt(1)
        p.add_run(f"{group['category']}: ").bold = True
        p.add_run(", ".join(group["items"]))

    add_section_heading(doc, RESUME_SECTIONS["experience"])
    for job in data.get("experience", []):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(job["company"] + (f" - {job['location']}" if job.get("location") else ""))
        run.bold, run.font.size = True, Pt(11)
        p = doc.add_paragraph()
        p.paragraph_format.keep_with_next = True
        p.paragraph_format.space_after = Pt(2)
        p.add_run(f"{job['title']} | {job['dates']}").bold = True
        if job.get("description"):
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(2)
            p.add_run(job["description"]).italic = True
        for bullet in job["bullets"]:
            bp = doc.add_paragraph(bullet, style="List Bullet")
            bp.paragraph_format.space_after = Pt(1)

    if data.get("professional_development"):
        add_section_heading(doc, RESUME_SECTIONS["professional_development"])
        for item in data["professional_development"]:
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(2)
            title, sep, rest = item.partition(":")
            if sep and len(title) <= 80:  # "AI Agent Development (Self-Directed, 2026): ..."
                p.add_run(title.strip() + ": ").bold = True
                p.add_run(rest.strip())
            else:
                p.add_run(item)
    if data.get("education"):
        add_section_heading(doc, RESUME_SECTIONS["education"])
        for item in data["education"]:
            doc.add_paragraph(item)
    if data.get("certifications"):
        add_section_heading(doc, RESUME_SECTIONS["certifications"])
        for item in data["certifications"]:
            doc.add_paragraph(item)

    set_properties(doc, author=data["name"], title=f"{data['name']} - Resume - {data['job_title']}",
                   keywords=", ".join(data.get("ats", {}).get("found", [])))
    doc.save(path)


def build_cover_letter(data: dict, path: Path):
    letter = data["cover_letter"]
    doc = new_document(11)
    add_name_header(doc, data["name"], data["contact"])

    def para(text="", space_after=10, bold=False):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(space_after)
        p.add_run(text).bold = bold
        return p

    applied = applied_date(data)
    para(f"{applied:%B} {applied.day}, {applied.year}", space_after=12)
    para("Hiring Team", space_after=0)
    para(data["company"], space_after=0)
    if data.get("job_location"):
        para(data["job_location"], space_after=0)
    para("", space_after=6)
    para(f"Re: {data['job_title']}", bold=True)
    para(letter["greeting"])
    paragraphs = list(letter["paragraphs"])
    if WORK_AUTH and paragraphs and not any("sponsor" in b.lower() for b in paragraphs):
        # The model forgot: add the statement to the closing paragraph.
        paragraphs[-1] = paragraphs[-1].rstrip() + " I do not require visa sponsorship."
    for body in paragraphs:
        para(body)
    para(letter.get("closing") or "Sincerely,", space_after=24)
    para(data["name"], space_after=0)

    set_properties(doc, author=data["name"], title=f"{data['name']} - Cover Letter - {data['job_title']}")
    doc.save(path)


MATCH_COLORS = {"Meets": ("1E7B3A", "E3F4E8"), "Partial": ("8A5A00", "FFF4D6"), "Gap": ("B42318", "FDE7E5")}


def clean_xml_text(text: str) -> str:
    """Strip control characters that Word documents cannot store (common in scraped pages)."""
    return re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", str(text or ""))


def shade_cell(cell, hex_fill: str):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_fill)
    tc_pr.append(shd)


def set_cell(cell, text: str, bold=False, color: RGBColor | None = None, size=10):
    cell.text = ""
    run = cell.paragraphs[0].add_run(clean_xml_text(text))
    run.bold, run.font.size = bold, Pt(size)
    if color is not None:
        run.font.color.rgb = color


def set_col_widths(table, widths: list[float]):
    """Set widths on the grid, the columns and every cell so Word and LibreOffice both honor them."""
    table.autofit = False
    for i, width in enumerate(widths):
        table.columns[i].width = Inches(width)
        for cell in table.columns[i].cells:
            cell.width = Inches(width)


def add_table(doc, headers: list[str], rows: list[list[str]], widths: list[float]):
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    for i, header in enumerate(headers):
        cell = table.rows[0].cells[i]
        set_cell(cell, header, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))
        shade_cell(cell, "1F3A5F")
    for row in rows:
        cells = table.add_row().cells
        for i, value in enumerate(row):
            set_cell(cells[i], value)
    set_col_widths(table, widths)
    return table


def add_match_table(doc, items: list[dict]):
    if not items:
        doc.add_paragraph("None listed in the posting.")
        return
    table = add_table(doc, ["Requirement", "My match", "Evidence from my resume"],
                      [[i["requirement"], i.get("match", ""), i.get("evidence", "")] for i in items],
                      [3.0, 0.9, 3.4])
    for row, item in zip(table.rows[1:], items):
        text_hex, fill_hex = MATCH_COLORS.get(item.get("match"), ("444444", "FFFFFF"))
        set_cell(row.cells[1], item.get("match", ""), bold=True, color=RGBColor.from_string(text_hex))
        shade_cell(row.cells[1], fill_hex)


def build_requirements(data: dict, path: Path):
    req = data.get("job_requirements", {})
    doc = new_document(10.5)

    p = doc.add_paragraph()
    run = p.add_run(data["job_title"])
    run.bold, run.font.size, run.font.color.rgb = True, Pt(18), ACCENT
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(6)
    run = p.add_run(f"{data['company']}  |  Job Requirements")
    run.font.size, run.font.color.rgb = Pt(11), GRAY

    details = [
        ["Application date", f"{applied_date(data):%B} {applied_date(data).day}, {applied_date(data).year}"],
        ["Company", data["company"]],
        ["Job title", data["job_title"]],
        ["Location", data.get("job_location") or "Not listed"],
        ["Work arrangement", req.get("work_arrangement") or "Not listed"],
        ["Employment type", req.get("employment_type") or "Not listed"],
        ["Experience level", req.get("experience_level") or "Not listed"],
        ["Compensation", req.get("compensation") or "Not listed"],
        ["Suggested salary ask", salary_text(data.get("salary"))],
        ["Visa sponsorship (posting)", req.get("sponsorship") or "Not listed"],
        ["Source", data.get("source") or "Not recorded"],
    ]
    if data.get("fit_score") is not None:
        fit = data.get("fit") or {}
        details.append(["Role fit", f"{data['fit_score']}%" + (f" - {fit['reason']}" if fit.get("reason") else "")])
    ats = data.get("ats") or {}
    if ats.get("total"):
        details.append(["ATS keyword coverage", f"{ats.get('required_score', ats['score'])}% of required keywords; "
                                                f"{ats['score']}% of all posting keywords ({len(ats['found'])} of {ats['total']})"])
    if data.get("rules_applied"):
        details.append(["Experience rules applied", "; ".join(data["rules_applied"])])
    table = doc.add_table(rows=0, cols=2)
    table.style = "Table Grid"
    for label, value in details:
        cells = table.add_row().cells
        set_cell(cells[0], label, bold=True)
        shade_cell(cells[0], "EEF2F7")
        set_cell(cells[1], value)
    set_col_widths(table, [1.8, 5.5])

    all_items = req.get("required", []) + req.get("preferred", [])
    if all_items:
        counts = {k: sum(1 for i in all_items if i.get("match") == k) for k in MATCH_COLORS}
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(8)
        p.add_run("Match summary: ").bold = True
        p.add_run(f"Meets {counts['Meets']}  |  Partial {counts['Partial']}  |  Gap {counts['Gap']}  "
                  f"(out of {len(all_items)} listed qualifications)")

    add_section_heading(doc, "Role Overview")
    doc.add_paragraph(clean_xml_text(req.get("overview", "")))

    add_section_heading(doc, "Key Responsibilities")
    for item in req.get("responsibilities", []):
        doc.add_paragraph(clean_xml_text(item), style="List Bullet")

    add_section_heading(doc, "Required Qualifications")
    add_match_table(doc, req.get("required", []))

    add_section_heading(doc, "Preferred Qualifications")
    add_match_table(doc, req.get("preferred", []))

    if req.get("tech_stack"):
        add_section_heading(doc, "Tech Stack and Keywords")
        doc.add_paragraph(clean_xml_text(", ".join(req["tech_stack"])))
        if ats.get("total"):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(4)
            p.add_run("In your resume: ").bold = True
            p.add_run(clean_xml_text(", ".join(ats["found"]) or "none"))
            if ats.get("phrases_found") or ats.get("phrases_missing"):
                p = doc.add_paragraph()
                p.add_run("Posting wording mirrored: ").bold = True
                p.add_run(clean_xml_text("; ".join(ats.get("phrases_found", [])) or "none"))
                if ats.get("phrases_missing"):
                    p = doc.add_paragraph()
                    p.add_run("Posting wording not mirrored: ").bold = True
                    run = p.add_run(clean_xml_text("; ".join(ats["phrases_missing"])))
                    run.font.color.rgb = RGBColor(0x8A, 0x5A, 0x00)
            p = doc.add_paragraph()
            p.add_run("Missing from your resume: ").bold = True
            run = p.add_run(clean_xml_text(", ".join(ats["missing"]) or "none"))
            if ats["missing"]:
                run.font.color.rgb = RGBColor(0xB4, 0x23, 0x18)

    if ats.get("checks"):
        add_section_heading(doc, "ATS Compliance Check: " + ("PASS" if ats["passed"] else "NEEDS ATTENTION"))
        table = add_table(doc, ["Check", "Result", "Detail"],
                          [[c["check"], "Pass" if c["passed"] else "Fix", c["detail"]] for c in ats["checks"]],
                          [3.3, 0.7, 3.3])
        for row, c in zip(table.rows[1:], ats["checks"]):
            text_hex, fill_hex = MATCH_COLORS["Meets" if c["passed"] else "Gap"]
            set_cell(row.cells[1], "Pass" if c["passed"] else "Fix", bold=True, color=RGBColor.from_string(text_hex))
            shade_cell(row.cells[1], fill_hex)

    if data.get("gaps"):
        add_section_heading(doc, "Gaps to Prepare For")
        for gap in data["gaps"]:
            doc.add_paragraph(clean_xml_text(gap), style="List Bullet")

    if req.get("interview_focus"):
        add_section_heading(doc, "Likely Interview Focus")
        for topic in req["interview_focus"]:
            doc.add_paragraph(clean_xml_text(topic), style="List Bullet")

    if data.get("job_description_raw"):
        doc.add_page_break()
        add_section_heading(doc, "Original Job Posting")
        for block in re.split(r"\n\s*\n", clean_xml_text(data["job_description_raw"])):
            if block.strip():
                para = doc.add_paragraph()
                para.paragraph_format.space_after = Pt(6)
                run = para.add_run(re.sub(r"[ \t]*\n[ \t]*", "\n", block.strip()))
                run.font.size = Pt(9.5)

    set_properties(doc, author=data.get("name", ""),
                   title=f"Job Requirements - {data['job_title']} - {data['company']}")
    doc.save(path)


# ---------------------------------------------------------------------------
# Saving everything
# ---------------------------------------------------------------------------
def slug(text: str) -> str:
    """File/folder-safe name: drops tags in brackets like [Remote-US] or (Hybrid), then turns every
    comma, dash, slash, ampersand or other special character into a single underscore."""
    text = re.sub(r"\[[^\]]*\]|\([^)]*\)|\{[^}]*\}", " ", text or "")
    text = re.sub(r"[^A-Za-z0-9]+", "_", text).strip("_")
    return text[:50].rstrip("_") or "Job"


def applied_date(data: dict) -> date:
    try:
        return date.fromisoformat(data.get("applied_date", ""))
    except ValueError:
        return date.today()


def new_folder_for(data: dict) -> Path:
    """applications/2026-10-14_Company_Job_Title (adds _2, _3 ... if run again the same day)."""
    base = f"{applied_date(data).isoformat()}_{slug(data['company'])}_{slug(data['job_title'])}"
    folder, n = OUTPUT_DIR / base, 2
    while folder.exists():
        folder, n = OUTPUT_DIR / f"{base}_{n}", n + 1
    return folder


# ---------------------------------------------------------------------------
# Two-page limit for the resume
# ---------------------------------------------------------------------------
MAX_RESUME_PAGES = int(SETTINGS.get("max_resume_pages") or 2)
HIGHER_DEGREE = re.compile(r"\b(master|m\.?s\.?|mba|ph\.?d|doctor)", re.I)


def _pages_with_word(path: Path) -> int | None:
    """Exact page count from Microsoft Word (Windows, needs: pip install pywin32)."""
    if not sys.platform.startswith("win"):
        return None
    try:
        import win32com.client  # type: ignore
    except ImportError:
        return None
    word = doc = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Open(str(path.resolve()), ReadOnly=True, AddToRecentFiles=False)
        return int(doc.ComputeStatistics(2))  # 2 = wdStatisticPages
    except Exception:
        return None
    finally:
        try:
            if doc is not None:
                doc.Close(False)
            if word is not None:
                word.Quit()
        except Exception:
            pass


def _pages_with_libreoffice(path: Path) -> int | None:
    import subprocess
    import tempfile
    exe = shutil.which("soffice") or shutil.which("libreoffice")
    if not exe:
        return None
    try:
        from pypdf import PdfReader
        with tempfile.TemporaryDirectory() as tmp:
            subprocess.run([exe, "--headless", "--convert-to", "pdf", "--outdir", tmp, str(path)],
                           capture_output=True, timeout=120, check=True)
            return len(PdfReader(Path(tmp) / (path.stem + ".pdf")).pages)
    except Exception:
        return None


def estimate_resume_pages(data: dict) -> float:
    """Layout estimate matching build_resume (Letter, 0.6" margins, Calibri 10.5 pt).
    Used only when neither Word nor LibreOffice is available to count pages exactly."""
    import math
    body_chars, bullet_chars = 128, 120          # characters per line
    line = 10.5 * 1.05 * 1.22                    # points per line of body text
    def lines(text, width):
        return max(1, math.ceil(len(text) / width))
    height = 20 * 1.22 + 2                       # name
    if data.get("headline"):
        height += 11.5 * 1.22 + 1
    height += 9.5 * 1.22 + 4 + (10 * 1.22 + 6 if WORK_AUTH else 0)
    heading = 12 * 1.22 + 11 + 2                 # section heading with spacing and rule
    sections = 3 + sum(1 for k in ("professional_development", "education", "certifications") if data.get(k))
    height += sections * heading
    height += lines(data.get("summary", ""), body_chars) * line
    for g in data.get("skills", []):
        height += lines(f"{g['category']}: {', '.join(g['items'])}", bullet_chars) * line + 1
    for job in data.get("experience", []):
        height += 6 + 11 * 1.22 + line + 2
        if job.get("description"):
            height += lines(job["description"], body_chars) * line + 2
        height += sum(lines(b, bullet_chars) * line + 1 for b in job.get("bullets", []))
    for item in data.get("professional_development", []):
        height += lines(item, body_chars) * line + 2
    for item in data.get("education", []) + data.get("certifications", []):
        height += lines(item, body_chars) * line
    return height / (10.0 * 72)                  # 10" of usable height per page


def count_resume_pages(path: Path, data: dict) -> tuple[float, str]:
    for counter, how in ((_pages_with_word, "Word"), (_pages_with_libreoffice, "LibreOffice")):
        pages = counter(path)
        if pages:
            return float(pages), how
    return estimate_resume_pages(data), "estimate"


def _over_limit(pages: float, how: str) -> bool:
    # The estimate runs slightly low near the limit (checked against 124 real resumes), so it uses a margin.
    return pages > MAX_RESUME_PAGES if how != "estimate" else pages > MAX_RESUME_PAGES - 0.07


def fit_resume_to_two_pages(data: dict, path: Path) -> list[str]:
    """If the resume runs past the page limit: first keep only the company and title of the role named in
    settings (default: your oldest role), then drop a Bachelor's degree when a higher degree is listed."""
    trims = []
    pages, how = count_resume_pages(path, data)
    steps = []
    target = str(SETTINGS.get("shorten_role_first") or "").lower()
    old_jobs = [j for j in data.get("experience", []) if target and target in j.get("company", "").lower()]
    old_jobs = old_jobs or data.get("experience", [])[-1:]
    for job in old_jobs:
        if job.get("bullets") or job.get("description"):
            steps.append((f"kept only the title for {job['company']}",
                          lambda j=job: j.update(bullets=[], description="")))
    education = data.get("education", [])
    has_higher = any(HIGHER_DEGREE.search(e) for e in education if not re.match(r"\s*bachelor", e, re.I))
    if SETTINGS.get("drop_bachelors_if_higher_degree", True) and has_higher and \
            any(re.match(r"\s*bachelor", e, re.I) for e in education):
        steps.append(("removed the Bachelor's degree",
                      lambda: data.update(education=[e for e in data["education"]
                                                     if not re.match(r"\s*bachelor", e, re.I)])))
    for label, apply in steps:
        if not _over_limit(pages, how):
            break
        apply()
        trims.append(label)
        build_resume(data, path)
        pages, how = count_resume_pages(path, data)
    shown = f"{pages:.0f}" if how != "estimate" else f"about {pages:.1f}"
    if trims:
        print(f"Resume was over {MAX_RESUME_PAGES} pages: {'; '.join(trims)}. Now {shown} pages ({how}).")
    if _over_limit(pages, how):
        print(f"Note: the resume is still {shown} pages ({how}). Consider shortening older bullets "
              "in content.json, then run render.")
    return trims


def write_application(data: dict, folder: Path | None = None) -> Path:
    data = normalize_data(data)
    data["ats"] = ats_check(data)
    data["salary"] = salary_suggestion(data)
    data.setdefault("applied_date", date.today().isoformat())
    job_name = f"{slug(data['company'])}_{slug(data['job_title'])}"
    folder = folder or new_folder_for(data)
    folder.mkdir(parents=True, exist_ok=True)

    # Save the content first, so nothing generated is lost if a Word file fails to build.
    content_path = folder / "content.json"
    content_path.write_text(json.dumps(data, indent=2), encoding="utf-8")
    try:
        build_resume(data, folder / f"Resume_{job_name}.docx")
        trims = fit_resume_to_two_pages(data, folder / f"Resume_{job_name}.docx")
        if trims:
            data["resume_trims"] = sorted(set(data.get("resume_trims", []) + trims))
            data["ats"] = ats_check(data)  # re-score the trimmed resume
            content_path.write_text(json.dumps(data, indent=2), encoding="utf-8")
        build_cover_letter(data, folder / f"CoverLetter_{job_name}.docx")
        build_requirements(data, folder / f"JobRequirements_{job_name}.docx")
    except Exception:
        print(f"\nBuilding the Word files failed. The generated content is saved in:\n  {content_path}\n"
              f'After fixing the problem, rebuild without a new API call:\n  jobpilot render "{content_path}"\n')
        raise
    try:
        update_tracker(data, folder)
    except Exception as err:  # the tracker must never break a finished application
        print(f"Tracker not updated ({err}). Run  jobpilot track  to retry.")
    return folder


# ---------------------------------------------------------------------------
# Salary guidance: a suggested ask when the posting lists a maximum
# ---------------------------------------------------------------------------
_NUM = re.compile(r"(?<![\w.])\$?\s*(\d{1,3}(?:,\d{3})+(?:\.\d+)?|\d+(?:\.\d+)?)\s*([kK])?(?![\w%(])")
_HOURLY = re.compile(r"/\s*h(ou)?r\b|per\s+hour|hourly|an\s+hour|\bhr\b", re.I)


def _to_number(value):
    try:
        number = float(str(value).replace(",", "").replace("$", "").strip())
        return number if number > 0 else None
    except (TypeError, ValueError):
        return None


def parse_salary(text: str) -> dict | None:
    """Read min/max/period from free text such as '$115,000 - $130,000', 'Up to $120K', '$44-$54/hr'."""
    text = str(text or "")
    hourly = bool(_HOURLY.search(text))
    numbers = []
    for match in _NUM.finditer(text):
        value = float(match.group(1).replace(",", "")) * (1000 if match.group(2) else 1)
        if (hourly and 10 <= value <= 500) or (not hourly and 20_000 <= value <= 2_000_000):
            numbers.append((value, match.start()))
    if not numbers:
        return None
    first = numbers[0][0]
    up_to = re.search(r"\bup\s+to\b", text[:numbers[0][1]], re.I)
    if not up_to and len(numbers) >= 2 and numbers[1][0] > first:
        low, high = first, numbers[1][0]
    else:
        low, high = (None, first) if up_to else (first, first)
    return {"min": low, "max": high, "period": "hour" if hourly else "year",
            "multiple_ranges": len(numbers) > 2}


def _round_pay(value: float, period: str) -> float:
    step = 1 if period == "hour" else 1000
    return int(value / step + 0.5) * step


def money(value, period="year") -> str:
    if value is None:
        return ""
    return f"${value:,.2f}/hr" if period == "hour" else f"${value:,.0f}"


def salary_suggestion(data: dict) -> dict | None:
    """Suggest what to ask for, only when the posting lists a maximum.
    Narrow range: about 85% of the way up (senior experience). Wide range (max > 1.35 x min): about 70%,
    since the very top of a wide band is usually reserved. Only a maximum ('up to'): 97% of it.
    A single figure: that figure."""
    req = data.get("job_requirements", {})
    low, high = _to_number(req.get("salary_min")), _to_number(req.get("salary_max"))
    period = req.get("salary_period") if req.get("salary_period") in ("year", "hour") else None
    multiple = False
    if high is None:
        parsed = parse_salary(req.get("compensation", ""))
        if not parsed:
            return None
        low, high, period, multiple = parsed["min"], parsed["max"], parsed["period"], parsed["multiple_ranges"]
    period = period or ("hour" if high < 500 else "year")
    if low and low > high:
        low, high = high, low

    if low is None:                              # "up to $X"
        target, basis = high * 0.97, "near the posted maximum (only an upper limit is listed)"
    elif low == high:                            # a single figure
        target, basis = high, "the single figure posted"
    else:
        position = 0.85 if high / low <= 1.35 else 0.70
        target = low + position * (high - low)
        basis = f"about {round(position * 100)}% of the way up the posted range"
    target = min(_round_pay(target, period), high)
    top = min(high, _round_pay(target + 0.15 * (high - (low or high * 0.9)), period)) if high > target else high
    note = basis + ("; the posting lists several ranges (location tiers or target/full), so the first one "
                    "was used: check which applies to you" if multiple else "")
    return {"posted_min": low, "posted_max": high, "period": period,
            "suggested": target, "quote_low": target, "quote_high": max(top, target), "basis": note}


def salary_text(s: dict | None) -> str:
    if not s:
        return "No maximum listed in the posting"
    quote = (f" (quote {money(s['quote_low'], s['period'])} - {money(s['quote_high'], s['period'])})"
             if s["quote_high"] > s["quote_low"] else "")
    return f"{money(s['suggested'], s['period'])}{quote}: {s['basis']}"


# ---------------------------------------------------------------------------
# Application tracker: Job_Application_Tracker.xlsx next to the applications folder
# ---------------------------------------------------------------------------
TRACKER_NAME = "Job_Application_Tracker.xlsx"
STATUSES = ["Prepared", "Applied", "Recruiter Screen", "Technical Interview", "Final Interview",
            "Offer", "Rejected", "Withdrawn", "No Response"]

# (header, width, filled by) - "auto" columns are rewritten on every run,
# "you" columns are yours: set once with a default, then never overwritten.
TRACKER_COLUMNS = [
    ("Applied Date", 13, "auto"), ("Company", 24, "auto"), ("Job Title", 34, "auto"),
    ("Location", 20, "auto"), ("Work Arrangement", 16, "auto"), ("Employment Type", 16, "auto"),
    ("Compensation", 20, "auto"), ("Posted Max", 13, "auto"), ("Suggested Ask", 13, "auto"),
    ("Ask Range to Quote", 22, "auto"), ("Sponsorship (Posting)", 28, "auto"),
    ("ATS Result", 15, "auto"), ("Required Keywords %", 11, "auto"), ("All Keywords %", 10, "auto"),
    ("Qualifications Met", 14, "auto"), ("Experience Rules Applied", 26, "auto"),
    ("Job Link", 12, "auto"), ("Folder", 14, "auto"), ("Resume", 12, "auto"),
    ("Status", 18, "you"), ("Date Applied", 13, "you"), ("Salary Requested", 14, "you"),
    ("Follow-up Date", 13, "you"),
    ("Days Since Prepared", 11, "auto"), ("Contact", 22, "you"), ("Next Step", 24, "you"), ("Notes", 36, "you"),
]
COL = {name: i + 1 for i, (name, _, _) in enumerate(TRACKER_COLUMNS)}


def _letter(name: str) -> str:
    from openpyxl.utils import get_column_letter
    return get_column_letter(COL[name])


def _new_tracker():
    from openpyxl import Workbook
    from openpyxl.formatting.rule import CellIsRule, FormulaRule
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.worksheet.datavalidation import DataValidation

    wb = Workbook()
    ws = wb.active
    ws.title = "Applications"
    you_fill = PatternFill("solid", fgColor="FFF4CC")
    for name, width, kind in TRACKER_COLUMNS:
        cell = ws.cell(row=1, column=COL[name], value=name)
        cell.font = Font(name="Arial", bold=True, color="FFFFFF" if kind == "auto" else "1F3A5F", size=10)
        cell.fill = PatternFill("solid", fgColor="1F3A5F") if kind == "auto" else you_fill
        cell.alignment = Alignment(wrap_text=True, vertical="center")
        ws.column_dimensions[_letter(name)].width = width
    ws.row_dimensions[1].height = 32
    ws.freeze_panes = "D2"
    last = _letter(TRACKER_COLUMNS[-1][0])
    ws.auto_filter.ref = f"A1:{last}1000"

    status = _letter("Status")
    dv = DataValidation(type="list", formula1='"' + ",".join(STATUSES) + '"', allow_blank=True)
    dv.error, dv.errorTitle = "Pick a status from the list.", "Status"
    ws.add_data_validation(dv)
    dv.add(f"{status}2:{status}1000")

    green, amber, red, blue = ("E3F4E8", "1E7B3A"), ("FFF4D6", "8A5A00"), ("FDE7E5", "B42318"), ("E6EEF8", "1F3A5F")
    def fill_rule(op, value, colors, col):
        ws.conditional_formatting.add(f"{col}2:{col}1000", CellIsRule(
            operator=op, formula=[value], fill=PatternFill("solid", fgColor=colors[0]),
            font=Font(name="Arial", color=colors[1], bold=True)))
    ats = _letter("ATS Result")
    fill_rule("equal", '"PASS"', green, ats)
    fill_rule("equal", '"NEEDS ATTENTION"', amber, ats)
    for value, colors in (("Offer", green), ("Technical Interview", blue), ("Final Interview", blue),
                          ("Recruiter Screen", blue), ("Rejected", red), ("Applied", amber)):
        fill_rule("equal", f'"{value}"', colors, status)
    follow = _letter("Follow-up Date")
    ws.conditional_formatting.add(f"{follow}2:{follow}1000", FormulaRule(
        formula=[f'AND({follow}2<>"",{follow}2<=TODAY(),OR({status}2="Applied",{status}2="No Response"))'],
        fill=PatternFill("solid", fgColor=red[0]), font=Font(name="Arial", color=red[1], bold=True)))

    _build_summary(wb)
    _build_help(wb)
    return wb


def _build_summary(wb):
    from openpyxl.styles import Font
    ws = wb.create_sheet("Summary")
    ws.column_dimensions["A"].width, ws.column_dimensions["B"].width = 34, 14
    ws["A1"] = "Job Search Summary"
    ws["A1"].font = Font(name="Arial", size=14, bold=True, color="1F3A5F")
    a, s = "Applications!", _letter("Status")
    d, req = _letter("Applied Date"), _letter("Required Keywords %")
    rows = [("Applications prepared", f"=COUNTA({a}$B$2:$B$1000)"),
            ("Applications sent (not Prepared)", f'=B3-COUNTIF({a}${s}$2:${s}$1000,"Prepared")'),
            ("Prepared in the last 7 days", f'=COUNTIF({a}${d}$2:${d}$1000,">="&(TODAY()-7))'),
            ("Interview rate (of sent)", f'=IF(B4=0,0,(COUNTIF({a}${s}$2:${s}$1000,"*Interview*")'
                                         f'+COUNTIF({a}${s}$2:${s}$1000,"Recruiter Screen")'
                                         f'+COUNTIF({a}${s}$2:${s}$1000,"Offer"))/B4)'),
            ("Average required-keyword coverage", f"=IFERROR(AVERAGE({a}${req}$2:${req}$1000),0)"),
            ("", None), ("By status", None)]
    rows += [(st, f'=COUNTIF({a}${s}$2:${s}$1000,"{st}")') for st in STATUSES]
    for i, (label, formula) in enumerate(rows, 3):
        ws.cell(row=i, column=1, value=label).font = Font(name="Arial", bold=label == "By status", size=10)
        if formula:
            cell = ws.cell(row=i, column=2, value=formula)
            cell.font = Font(name="Arial", size=10)
    ws["B6"].number_format = "0%"
    ws["B7"].number_format = "0%"


def _build_help(wb):
    from openpyxl.styles import Font, Alignment
    ws = wb.create_sheet("How to use")
    ws.column_dimensions["A"].width = 110
    lines = ["How this tracker works",
             "Dark blue columns are filled by JobPilot on every tailor or render run. Edits there are replaced.",
             "Yellow columns are yours (Status, Date Applied, Follow-up Date, Contact, Next Step, Notes). JobPilot never overwrites them.",
             "New rows start as Status = Prepared. Change it from the drop-down when you apply, hear back, or interview.",
             "Follow-up Date turns red when it is due and the status is still Applied or No Response.",
             "Posted Max and Suggested Ask come from the posting's salary range: about 85% of the way up a narrow range, "
             "about 70% up a wide one (max more than 1.35 x min), 97% of an 'up to' maximum. Record what you actually entered in Salary Requested.",
             "Job Link, Folder and Resume are clickable links (keep this file next to the applications folder).",
             "Close this file in Excel before running JobPilot; if it is open, run 'jobpilot track' afterwards to catch up.",
             "Rebuild or catch up all rows from the applications folder at any time:  jobpilot track"]
    for i, line in enumerate(lines, 1):
        cell = ws.cell(row=i, column=1, value=line)
        cell.font = Font(name="Arial", size=12 if i == 1 else 10, bold=i == 1, color="1F3A5F" if i == 1 else "000000")
        cell.alignment = Alignment(wrap_text=True)


def _upgrade_tracker(wb):
    """Move rows from an older tracker layout into the current one, keeping every value you typed."""
    old = wb["Applications"]
    headers = [c.value for c in old[1]]
    if headers[:len(TRACKER_COLUMNS)] == [name for name, _, _ in TRACKER_COLUMNS]:
        return wb
    new_wb = _new_tracker()
    ws = new_wb["Applications"]
    out = 2
    for r in range(2, old.max_row + 1):
        if not any(old.cell(row=r, column=c).value for c in range(1, len(headers) + 1)):
            continue
        for c, name in enumerate(headers, 1):
            if name not in COL:
                continue
            src, dst = old.cell(row=r, column=c), ws.cell(row=out, column=COL[name])
            if isinstance(src.value, str) and src.value.startswith("="):
                continue  # formulas are rebuilt for the new column positions
            dst.value, dst.number_format = src.value, src.number_format
            if src.hyperlink:
                dst.hyperlink = src.hyperlink.target
            dst.font = src.font.copy()
            dst.alignment = src.alignment.copy()
        prepared = _letter("Applied Date")
        days = ws.cell(row=out, column=COL["Days Since Prepared"],
                       value=f'=IF({prepared}{out}="","",TODAY()-{prepared}{out})')
        days.number_format = "0"
        date_col, follow = _letter("Date Applied"), ws.cell(row=out, column=COL["Follow-up Date"])
        if follow.value is None:
            follow.value = f'=IF({date_col}{out}="","",{date_col}{out}+7)'
            follow.number_format = "yyyy-mm-dd"
        out += 1
    print(f"Tracker upgraded to the new layout ({out - 2} rows kept, your entries preserved).")
    return new_wb


def _tracker_values(data: dict, folder: Path) -> dict:
    req = data.get("job_requirements", {})
    items = req.get("required", []) + req.get("preferred", [])
    met = sum(1 for i in items if i.get("match") == "Meets")
    ats = data.get("ats") or {}
    sal = data.get("salary") if "salary" in data else salary_suggestion(data)
    job_name = f"{slug(data['company'])}_{slug(data['job_title'])}"
    resume = next(iter(sorted(folder.glob("Resume_*.docx"))), folder / f"Resume_{job_name}.docx")
    return {
        "Applied Date": applied_date(data),
        "Company": data.get("company", ""), "Job Title": data.get("job_title", ""),
        "Location": data.get("job_location", ""), "Work Arrangement": req.get("work_arrangement", ""),
        "Employment Type": req.get("employment_type", ""), "Compensation": req.get("compensation", ""),
        "Posted Max": sal["posted_max"] if sal else None,
        "Suggested Ask": sal["suggested"] if sal else None,
        "Ask Range to Quote": (f"{money(sal['quote_low'], sal['period'])} - {money(sal['quote_high'], sal['period'])}"
                               if sal and sal["quote_high"] > sal["quote_low"] else
                               (money(sal["suggested"], sal["period"]) if sal else "No max listed")),
        "Sponsorship (Posting)": req.get("sponsorship", ""),
        "ATS Result": ("PASS" if ats.get("passed") else "NEEDS ATTENTION") if ats.get("total") else "Not checked",
        "Required Keywords %": (ats.get("required_score", ats.get("score", 0)) / 100) if ats.get("total") else None,
        "All Keywords %": (ats.get("score", 0) / 100) if ats.get("total") else None,
        "Qualifications Met": f"{met} of {len(items)}" if items else "",
        "Experience Rules Applied": "; ".join(r.split(" (")[0] for r in data.get("rules_applied", [])),
        "Job Link": data.get("source", "") if str(data.get("source", "")).startswith("http") else "",
        "Folder": folder.name, "Resume": resume.name,
    }


def update_tracker(data: dict, folder: Path, quiet: bool = False) -> Path | None:
    """Add or refresh this application's row. Your own columns are never overwritten."""
    try:
        from openpyxl import load_workbook
        from openpyxl.styles import Alignment, Font
    except ImportError:
        print("Tracker skipped: install openpyxl with  pip install openpyxl")
        return None
    path = folder.resolve().parent.parent / TRACKER_NAME
    wb = _upgrade_tracker(load_workbook(path)) if path.exists() else _new_tracker()
    ws = wb["Applications"]

    values = _tracker_values(data, folder)
    folder_col = COL["Folder"]
    row = next((r for r in range(2, ws.max_row + 1) if ws.cell(row=r, column=folder_col).value == folder.name), None)
    is_new = row is None
    if is_new:
        row = next((r for r in range(2, ws.max_row + 2) if not ws.cell(row=r, column=COL["Company"]).value), ws.max_row + 1)

    font, link_font = Font(name="Arial", size=10), Font(name="Arial", size=10, color="1155CC", underline="single")
    for name, value in values.items():
        cell = ws.cell(row=row, column=COL[name], value=value)
        cell.font, cell.alignment = font, Alignment(vertical="top", wrap_text=name in ("Job Title", "Sponsorship (Posting)", "Experience Rules Applied"))
        cell.hyperlink = None
    ws.cell(row=row, column=COL["Applied Date"]).number_format = "yyyy-mm-dd"
    for name in ("Required Keywords %", "All Keywords %"):
        ws.cell(row=row, column=COL[name]).number_format = "0%"
    sal = data.get("salary") if "salary" in data else salary_suggestion(data)
    pay_format = '$#,##0.00"/hr"' if sal and sal["period"] == "hour" else "$#,##0"
    for name in ("Posted Max", "Suggested Ask"):
        ws.cell(row=row, column=COL[name]).number_format = pay_format
    rel = f"{folder.parent.name}/{folder.name}"
    for name, target, label in (("Job Link", values["Job Link"], "Open posting"),
                                ("Folder", rel, folder.name),  # the folder name also identifies the row
                                ("Resume", f"{rel}/{values['Resume']}", "Open resume")):
        if target:
            cell = ws.cell(row=row, column=COL[name], value=label)
            cell.hyperlink, cell.font = target, link_font
    days = ws.cell(row=row, column=COL["Days Since Prepared"],
                   value=f'=IF({_letter("Applied Date")}{row}="","",TODAY()-{_letter("Applied Date")}{row})')
    days.font, days.number_format = font, "0"

    if is_new or ws.cell(row=row, column=COL["Status"]).value is None:  # defaults for your columns, set once
        defaults = {"Status": "Prepared", "Date Applied": None,
                    "Follow-up Date": f'=IF({_letter("Date Applied")}{row}="","",{_letter("Date Applied")}{row}+7)'}
        for name, value in defaults.items():
            cell = ws.cell(row=row, column=COL[name], value=value)
            cell.font = font
        for name in ("Date Applied", "Follow-up Date"):
            ws.cell(row=row, column=COL[name]).number_format = "yyyy-mm-dd"
        ws.cell(row=row, column=COL["Salary Requested"]).number_format = pay_format

    wb.calculation.fullCalcOnLoad = True  # Excel recalculates the formulas when you open it
    try:
        wb.save(path)
    except PermissionError:
        print(f"Tracker not updated: {path.name} is open in Excel. Close it, then run:  jobpilot track")
        return None
    if not quiet:
        print(f"Tracker {'added' if is_new else 'updated'}: {path}")
    return path


def cmd_track(args):
    """Add or refresh a row for every application folder (also catches up older folders)."""
    folders = [p.parent for p in OUTPUT_DIR.glob("*/content.json")]
    if not folders:
        sys.exit(f"No applications found in {OUTPUT_DIR.resolve()}. Run this from your jobApplication folder.")
    apps = []
    for folder in folders:
        data = normalize_data(json.loads((folder / "content.json").read_text(encoding="utf-8")))
        if "applied_date" not in data:  # older folders: use the folder's date
            match = re.match(r"(\d{4}-\d{2}-\d{2})_", folder.name)
            data["applied_date"] = match.group(1) if match else date.fromtimestamp(
                (folder / "content.json").stat().st_mtime).isoformat()
        if not (data.get("ats") or {}).get("checks"):
            data["ats"] = ats_check(data)
        apps.append((data["applied_date"], folder.name, data, folder))
    path = None
    for _, _, data, folder in sorted(apps, key=lambda a: (a[0], a[1])):  # oldest first
        path = update_tracker(data, folder, quiet=True) or path
        if path is None:
            return
    print(f"Tracker synced: {len(folders)} applications -> {path}")


# ---------------------------------------------------------------------------
# Experience rules: your own extra experience, used only when a job asks for it
# ---------------------------------------------------------------------------
RULES_FILE = HOME / "experience_rules.json"

RULES_TEMPLATE = {
    "how_to_use": [
        "Each rule switches on when the job posting mentions ANY of its 'when_job_mentions' keywords.",
        "'company' must match part of the employer name in your resume (e.g. 'Acme' matches 'Acme Health Inc.').",
        "'replace_terms' swaps wording inside that company's bullets, e.g. Robot Framework -> Playwright.",
        "'add_bullets' are achievements added to that company's section. Write ONLY work you actually did,",
        "  with real tools, scale and results: interviewers will ask you about every line.",
        "'add_skills' adds keywords to the Technical Skills section under the given category.",
        "Bullets that still contain TODO are skipped until you replace them with your own text.",
        "Check the file with:  jobpilot rules",
    ],
    "rules": [
        {
            "name": "Playwright",
            "when_job_mentions": ["Playwright"],
            "changes": [
                {
                    "company": "Your Most Recent Employer",
                    "replace_terms": {"Robot Framework": "Playwright"},
                    "add_bullets": [
                        "TODO: Your Playwright framework work there (what you built, language, test count, result)",
                        "TODO: A second Playwright achievement (e.g. API testing, parallel runs, CI integration, flaky-test reduction)",
                    ],
                }
            ],
            "add_skills": {"Test Automation": ["Playwright"]},
        },
        {
            "name": "Cloud (AWS / Docker)",
            "when_job_mentions": ["AWS", "Amazon Web Services", "Docker", "Kubernetes", "cloud",
                                  "container", "containers", "EC2", "S3", "Lambda", "ECS", "EKS"],
            "changes": [
                {
                    "company": "Your Previous Employer",
                    "replace_terms": {},
                    "add_bullets": [
                        "TODO: Your AWS / Docker work there (services used, what you set up, result)",
                    ],
                },
                {
                    "company": "Another Employer",
                    "replace_terms": {},
                    "add_bullets": [
                        "TODO: Your AWS / Docker work there (services used, what you set up, result)",
                    ],
                },
            ],
            "add_skills": {"Cloud & DevOps": ["AWS", "Docker"]},
        },
    ],
}


def is_placeholder(text: str) -> bool:
    return "TODO" in text or bool(re.search(r"<[^>]+>", text))


def keyword_in(keyword: str, text: str) -> bool:
    """Whole-word, case-insensitive match that also works for terms like CI/CD or C#."""
    pattern = rf"(?<![A-Za-z0-9]){re.escape(keyword.strip())}(?![A-Za-z0-9])"
    return bool(keyword.strip()) and re.search(pattern, text, re.IGNORECASE) is not None


def load_rules() -> list[dict]:
    if not RULES_FILE.exists():
        return []
    try:
        rules = json.loads(RULES_FILE.read_text(encoding="utf-8")).get("rules", [])
    except json.JSONDecodeError as err:
        sys.exit(f"Your experience rules file has a JSON error (line {err.lineno}, column {err.colno}):\n"
                 f"  {RULES_FILE}\nFix it (often a missing comma or quote), then run the command again.")
    return [r for r in rules if isinstance(r, dict)]


def _as_list(value) -> list:
    if value is None:
        return []
    return value if isinstance(value, list) else [value]


def active_rules(rules: list[dict], job_text: str) -> list[dict]:
    """Rules whose keywords appear in the job posting, with the usable (non-TODO) bullets only."""
    active = []
    for rule in rules:
        hits = [k for k in _as_list(rule.get("when_job_mentions")) if keyword_in(str(k), job_text)]
        if not hits:
            continue
        changes = []
        for change in _as_list(rule.get("changes")):
            bullets = [b for b in _as_list(change.get("add_bullets")) if b and not is_placeholder(b)]
            skipped = len(_as_list(change.get("add_bullets"))) - len(bullets)
            if skipped:
                log(f"  Note: {skipped} TODO bullet(s) for {change.get('company')} skipped. "
                      f"Add your real text in {RULES_FILE}")
            changes.append({"company": str(change.get("company", "")),
                            "replace_terms": change.get("replace_terms") or {},
                            "add_bullets": bullets})
        active.append({"name": rule.get("name", "Rule"), "matched": hits, "changes": changes,
                       "add_skills": rule.get("add_skills") or {}})
    return active


def rules_prompt(active: list[dict]) -> str:
    if not active:
        return ""
    lines = ["<experience_rules>",
             "Verified experience supplied by the candidate for this job. Apply all of it:"]
    for rule in active:
        lines.append(f"- {rule['name']} (the posting mentions: {', '.join(rule['matched'])})")
        for change in rule["changes"]:
            for old, new in change["replace_terms"].items():
                lines.append(f"  - In the {change['company']} role, the work described with {old} "
                             f"must be presented as {new} work. Do not mention {old} in that role.")
            for bullet in change["add_bullets"]:
                lines.append(f"  - Include in the {change['company']} role: {bullet}")
        for category, items in rule["add_skills"].items():
            lines.append(f"  - Add to Technical Skills under '{category}': {', '.join(_as_list(items))}")
    lines.append("</experience_rules>\n\n")
    return "\n".join(lines)


def rules_facts(active: list[dict]) -> str:
    """Only what you wrote as facts (bullets, replacement terms, skills), never the trigger keywords."""
    parts = []
    for rule in active:
        for change in rule["changes"]:
            parts += list(change["replace_terms"].values()) + change["add_bullets"]
        for items in rule["add_skills"].values():
            parts += [str(i) for i in _as_list(items)]
    return " ".join(parts)


def _similar(a: str, b: str) -> bool:
    import difflib
    words_a = set(re.findall(r"[a-z0-9]{4,}", a.lower()))
    words_b = set(re.findall(r"[a-z0-9]{4,}", b.lower()))
    overlap = len(words_a & words_b) / max(1, min(len(words_a), len(words_b)))
    return overlap >= 0.6 or difflib.SequenceMatcher(None, a.lower(), b.lower()).ratio() >= 0.6


def apply_rules(data: dict, active: list[dict]) -> dict:
    """Guarantee the rules were applied, even if the model missed something."""
    applied = []
    for rule in active:
        for change in rule["changes"]:
            target = change["company"].lower()
            jobs = [j for j in data["experience"] if target and target in j.get("company", "").lower()]
            if not jobs:
                log(f"  Warning: no '{change['company']}' role found in the resume for rule '{rule['name']}'.")
                continue
            for job in jobs:
                for old, new in change["replace_terms"].items():
                    job["title"] = re.sub(re.escape(old), new, job["title"], flags=re.IGNORECASE)
                    job["bullets"] = [re.sub(re.escape(old), new, b, flags=re.IGNORECASE) for b in job["bullets"]]
                for bullet in reversed(change["add_bullets"]):
                    if not any(_similar(bullet, existing) for existing in job["bullets"]):
                        job["bullets"].insert(0, bullet)
                # remove duplicates that a replacement may have created
                seen, unique = set(), []
                for b in job["bullets"]:
                    if b.lower() not in seen:
                        seen.add(b.lower())
                        unique.append(b)
                job["bullets"] = unique
        for category, items in rule["add_skills"].items():
            skills_text = " ".join(f"{g['category']} {' '.join(g['items'])}" for g in data["skills"])
            new_items = [i for i in _as_list(items) if not keyword_in(str(i), skills_text)]
            if not new_items:
                continue  # already listed, e.g. "AWS" inside "AWS (Lambda, S3, SNS, EC2)"
            group = next((g for g in data["skills"] if g["category"].lower() == category.lower()), None)
            if group is None:
                group = {"category": category, "items": []}
                data["skills"].append(group)
            group["items"] = new_items + group["items"]
        applied.append(f"{rule['name']} (posting mentions {', '.join(rule['matched'])})")
    data["rules_applied"] = applied
    return data


def _resume_text(data: dict) -> str:
    return " ".join([data.get("headline", ""), data.get("summary", "")]
                    + [f"{g['category']} {' '.join(g['items'])}" for g in data.get("skills", [])]
                    + [" ".join([j["title"], j.get("description", ""), *j["bullets"]]) for j in data.get("experience", [])]
                    + data.get("education", []) + data.get("certifications", [])
                    + data.get("professional_development", []))


def _dashes(text: str) -> str:
    return re.sub(r"[‐-―]", "-", text or "")


def without_learning_section(resume_text: str) -> str:
    """Drop a 'Professional Development' (learning) section, so skills you only studied are never
    added to Core Skills as if you had used them at work."""
    return re.sub(r"(?ims)^\s*professional development\s*$.*?(?=^\s*(education|certifications|"
                  r"professional experience|core skills|technical skills)\s*$|\Z)", "", resume_text)


def ats_boost(data: dict, base_resume: str) -> list[str]:
    """Add posting keywords the candidate truly has (they appear in the base resume or rules)
    but the tailored resume left out, so ATS keyword matching finds them."""
    keywords = data.get("job_requirements", {}).get("tech_stack", [])
    text = _resume_text(data)
    added = [k for k in keywords
             if len(k) <= 40 and not keyword_in(k, text) and keyword_in(k, base_resume)]
    if added:
        group = next((g for g in data["skills"] if g["category"].lower() == "additional skills"), None)
        if group is None:
            group = {"category": "Additional Skills", "items": []}
            data["skills"].append(group)
        group["items"] += added
    return added


_STOPWORDS = {"and", "the", "with", "for", "you", "our", "your", "are", "who", "will", "can", "able", "work",
              "working", "experience", "strong", "excellent", "good", "great", "ability", "skills", "well"}


def phrase_present(phrase: str, text: str) -> bool:
    """True when the posting phrase appears as written, or all its key words appear in the resume."""
    if keyword_in(phrase, text):
        return True
    words = [w for w in re.findall(r"[a-z0-9/+#-]{3,}", phrase.lower()) if w not in _STOPWORDS]
    return bool(words) and all(re.search(rf"(?<![a-z0-9]){re.escape(w)}", text, re.I) for w in words)


DATE_FORMAT = re.compile(r"^[A-Z][a-z]+ \d{4} - ([A-Z][a-z]+ \d{4}|Present)$")
ALLOWED_CHARS = re.compile(r"[^\x20-\x7E–—‘’“”…é·]")


def ats_check(data: dict) -> dict:
    """Keyword coverage plus an ATS compliance checklist for the tailored resume."""
    keywords = list(dict.fromkeys(k for k in data.get("job_requirements", {}).get("tech_stack", []) if k))
    text = _resume_text(data)
    found = [k for k in keywords if keyword_in(k, text)]
    missing = [k for k in keywords if k not in found]
    score = round(100 * len(found) / len(keywords)) if keywords else 0

    words = len(text.split())
    bad_dates = [j["dates"] for j in data.get("experience", []) if not DATE_FORMAT.match(_dashes(j["dates"]))]
    odd_chars = sorted(set(ALLOWED_CHARS.findall(text)))
    title = _dashes(data.get("job_title", ""))
    title_clean = re.sub(r"\s*[\[(][^\])]*[\])]", "", title).strip()   # drop "[Remote-US]", "(Hybrid)"
    title_core = re.split(r"\s*[-,|]\s*", title_clean)[0].strip()      # "QA Engineer - Automation" -> "QA Engineer"
    top_text = _dashes(data.get("headline", "") + " " + data.get("summary", ""))
    contact = data.get("contact", "")

    required = [k for k in data.get("job_requirements", {}).get("required_keywords", []) if k] or keywords
    req_found = [k for k in required if keyword_in(k, text)]
    req_score = round(100 * len(req_found) / len(required)) if required else 100
    checks = [
        ("Required-keyword coverage of 80% or more", req_score >= 80,
         f"{req_score}% of required keywords ({len(req_found)} of {len(required)}); "
         f"{score}% of all posting keywords ({len(found)} of {len(keywords)})"),
        ("Job title from the posting in headline or summary",
         bool(title) and any(keyword_in(t, top_text) for t in (title, title_clean, title_core) if t), title or "-"),
        ("Email and phone in the document body",
         bool(re.search(r"\S+@\S+\.\w+", contact)) and bool(re.search(r"\d{3}\D{0,2}\d{3}\D?\d{4}", contact)),
         "Contact line is plain body text, not a Word header"),
        ("Standard section headings", True, ", ".join(RESUME_SECTIONS.values())),
        ("Single column, no tables, text boxes, images or graphics", True, "Built as plain paragraphs"),
        ("Standard font and real bullet style", True, "Calibri 10.5 pt, Word List Bullet style"),
        ("Consistent date format (Month YYYY - Month YYYY)", not bad_dates,
         "All dates consistent" if not bad_dates else "Check: " + "; ".join(bad_dates)),
        ("No unusual symbols or emoji", not odd_chars,
         "Plain text" if not odd_chars else "Found: " + " ".join(odd_chars)),
        ("Length about 1-2 pages (450-1200 words)", 450 <= words <= 1200, f"{words} words"),
        ("Saved as .docx (most widely accepted by ATS)", True, "Word document"),
    ]
    checks = [{"check": c, "passed": bool(ok), "detail": d} for c, ok, d in checks]
    phrases = [p for p in data.get("job_requirements", {}).get("ats_phrases", []) if p]
    phrases_found = [p for p in phrases if phrase_present(p, text)]
    return {"score": score, "total": len(keywords), "found": found, "missing": missing,
            "phrases_found": phrases_found, "phrases_missing": [p for p in phrases if p not in phrases_found],
            "required_score": req_score,
            "required_missing": [k for k in required if k not in req_found],
            "checks": checks, "passed": all(c["passed"] for c in checks)}


# ---------------------------------------------------------------------------
# Commands
# ---------------------------------------------------------------------------
def cmd_rules(args):
    if not RULES_FILE.exists():
        HOME.mkdir(parents=True, exist_ok=True)
        RULES_FILE.write_text(json.dumps(RULES_TEMPLATE, indent=2), encoding="utf-8")
        print(f"Created your experience rules file:\n  {RULES_FILE}\n")
    else:
        print(f"Experience rules file:\n  {RULES_FILE}\n")

    for rule in load_rules():
        print(f"[{rule.get('name', 'Rule')}]  when the job mentions: {', '.join(map(str, _as_list(rule.get('when_job_mentions'))))}")
        for change in _as_list(rule.get("changes")):
            for old, new in (change.get("replace_terms") or {}).items():
                print(f"   {change.get('company')}: replace '{old}' with '{new}'")
            for bullet in _as_list(change.get("add_bullets")):
                status = "TODO - not used yet" if is_placeholder(bullet) else "ready"
                print(f"   {change.get('company')}: add bullet ({status}): {bullet[:70]}")
        for category, items in (rule.get("add_skills") or {}).items():
            print(f"   skills > {category}: {', '.join(_as_list(items))}")
        print()

    if args.edit:
        _open_in_editor(RULES_FILE)
    else:
        print("Open it for editing with:  jobpilot rules --edit")


SAMPLE_RESUME = Path(__file__).resolve().parent / "examples" / "sample_resume.docx"


def cmd_setup(args):
    if getattr(args, "sample", False) or str(args.resume or "").strip().lower() == "sample":
        args.resume = str(SAMPLE_RESUME)
        print("Using the sample resume (Alex Morgan, a fictional QA engineer) so you can try JobPilot.")
        print("Replace it with your own before applying:  jobpilot setup --resume \"My_Resume.docx\"\n")
    if not args.resume:
        sys.exit('Give your resume file:  jobpilot setup --resume "My_Resume.docx"   (or --sample to try the sample)')
    source = Path(args.resume)
    if not source.exists():
        sys.exit(f"File not found: {source}")
    text = extract_text(source)
    if len(text.strip()) < 200:
        sys.exit("Couldn't read much text from that file. Try a .docx or .txt version of your resume.")
    HOME.mkdir(parents=True, exist_ok=True)
    shutil.copy(source, HOME / f"base_resume_original{source.suffix.lower()}")
    BASE_TEXT.write_text(text, encoding="utf-8")
    print(f"Base resume saved ({len(text.split())} words) to: {BASE_TEXT}")
    print("Tip: open that file and add any projects, tools or metrics your current resume leaves out.")
    print("The agent only uses what is in your base resume, so a fuller one gives better results.")
    write_default_settings()
    if not RULES_FILE.exists():
        print("\nNext: add experience that should appear only for certain jobs (e.g. Playwright, AWS):")
        print("  jobpilot rules --edit")
    print("Optional: add a work authorization line or change the page limit:  jobpilot settings --edit")


def _open_in_editor(path: Path):
    import subprocess
    if sys.platform.startswith("win"):
        subprocess.Popen(["notepad", str(path)])
    elif sys.platform == "darwin":
        subprocess.Popen(["open", "-t", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)])


def write_default_settings() -> bool:
    if SETTINGS_FILE.exists():
        return False
    HOME.mkdir(parents=True, exist_ok=True)
    SETTINGS_FILE.write_text(json.dumps(DEFAULT_SETTINGS, indent=2), encoding="utf-8")
    print(f"Created your settings file: {SETTINGS_FILE}")
    return True


def cmd_settings(args):
    write_default_settings()
    print(f"Settings file: {SETTINGS_FILE}\n")
    labels = {"work_authorization": "Line under your contact details (\"\" hides it)",
              "model": "Anthropic model",
              "max_resume_pages": "Page limit for resumes",
              "shorten_role_first": "Role to shorten first when too long (\"\" = oldest role)",
              "drop_bachelors_if_higher_degree": "Drop Bachelor's when a higher degree is listed and still too long",
              "min_fit_score": "Skip jobs matching your experience less than this % (0 = never skip)"}
    for key, label in labels.items():
        print(f"  {key:34} = {json.dumps(SETTINGS.get(key))}\n  {'':34}   {label}")
    if any(k in os.environ for k in ("JOBPILOT_WORK_AUTH", "JOBPILOT_MODEL", "JOBAPPLY_WORK_AUTH", "JOBAPPLY_MODEL")):
        print("\n  Note: JOBPILOT_WORK_AUTH / JOBPILOT_MODEL environment variables are set and override this file.")
    if args.edit:
        _open_in_editor(SETTINGS_FILE)
    else:
        print("\nOpen it for editing with:  jobpilot settings --edit")


def cmd_check(args):
    """Checks that everything needed is in place."""
    def line(ok, text, fix=""):
        print(f"  [{'OK' if ok else '!!'}] {text}" + (f"\n        -> {fix}" if not ok and fix else ""))
    print(f"JobPilot {__version__} (Python {sys.version.split()[0]})\n")
    ok_all = True
    py_ok = sys.version_info >= (3, 10)
    line(py_ok, "Python 3.10 or newer", "Install Python 3.10+ from python.org"); ok_all &= py_ok
    for module, package in (("anthropic", "anthropic"), ("docx", "python-docx"), ("requests", "requests"),
                            ("bs4", "beautifulsoup4"), ("pypdf", "pypdf"), ("openpyxl", "openpyxl")):
        try:
            __import__(module); found = True
        except ImportError:
            found = False
        line(found, f"Library {package}", f"pip install {package}"); ok_all &= found
    if sys.platform.startswith("win"):
        try:
            import win32com.client  # noqa: F401
            line(True, "pywin32 (exact page counts from Word)")
        except ImportError:
            line(True, "pywin32 not installed: page counts are estimated (optional: pip install pywin32)")
    key = os.environ.get("ANTHROPIC_API_KEY", "").strip()
    line(bool(key), "ANTHROPIC_API_KEY is set" + (" (looks valid)" if key.startswith("sk-ant-") else ""),
         'setx ANTHROPIC_API_KEY "sk-ant-..." then open a new terminal' if sys.platform.startswith("win")
         else 'export ANTHROPIC_API_KEY="sk-ant-..." in ~/.zshrc or ~/.bashrc'); ok_all &= bool(key)
    line(BASE_TEXT.exists(), f"Base resume saved ({BASE_TEXT})", 'jobpilot setup --resume "My_Resume.docx"')
    ok_all &= BASE_TEXT.exists()
    line(True, f"Settings: {SETTINGS_FILE if SETTINGS_FILE.exists() else 'defaults (no settings file yet)'}")
    rules = load_rules()
    ready = sum(1 for r in rules for c in _as_list(r.get("changes")) for b in _as_list(c.get("add_bullets"))
                if b and not is_placeholder(b))
    line(True, f"Experience rules: {len(rules)} rule(s), {ready} bullet(s) ready" if RULES_FILE.exists()
         else "Experience rules: none yet (optional: jobpilot rules --edit)")
    print("\n" + ("Ready. Try:  jobpilot tailor \"<job link>\"" if ok_all else "Fix the items marked !! and run  jobpilot check  again."))


def save_settings(**changes):
    current = {}
    if SETTINGS_FILE.exists():
        current = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
    merged = {**DEFAULT_SETTINGS, **current, **changes}
    HOME.mkdir(parents=True, exist_ok=True)
    SETTINGS_FILE.write_text(json.dumps(merged, indent=2), encoding="utf-8")
    SETTINGS.update(changes)


def _ask(question: str, default: str = "") -> str:
    try:
        answer = input(question).strip().strip('"').strip("'")
    except EOFError:
        return default
    return answer or default


def _save_api_key(key: str):
    import subprocess
    os.environ["ANTHROPIC_API_KEY"] = key
    if sys.platform.startswith("win"):
        result = subprocess.run(["setx", "ANTHROPIC_API_KEY", key], capture_output=True, text=True)
        print("  Saved for future terminals." if result.returncode == 0
              else f"  Could not save it automatically. Run:  setx ANTHROPIC_API_KEY \"{key[:12]}...\"")
        return
    rc = Path.home() / (".zshrc" if "zsh" in os.environ.get("SHELL", "") else ".bashrc")
    if _ask(f"  Add it to {rc} so new terminals have it? [Y/n] ", "y").lower().startswith("y"):
        with rc.open("a", encoding="utf-8") as f:
            f.write(f'\nexport ANTHROPIC_API_KEY="{key}"  # added by JobPilot\n')
        print(f"  Saved in {rc}. Open a new terminal (or run: source {rc}).")


def cmd_init(args):
    """Guided first-time setup."""
    import getpass
    print(f"=== JobPilot {__version__} setup ===\n")

    print("1) Anthropic API key  (create one at https://console.anthropic.com > Settings > API Keys)")
    key = os.environ.get("ANTHROPIC_API_KEY", "").strip()
    if key.startswith("sk-ant-"):
        print("  Already set. Keeping it.")
    else:
        try:
            key = getpass.getpass("  Paste your API key (hidden as you type; Enter to skip): ").strip()
        except EOFError:
            key = ""
        if key:
            if not key.startswith("sk-ant-"):
                print("  Note: Anthropic API keys normally start with sk-ant-. Saving it anyway.")
            _save_api_key(key)
        else:
            print("  Skipped. You can run  jobpilot init  again later.")

    print("\n2) Your current resume (.docx, .pdf, .md or .txt)")
    if BASE_TEXT.exists() and not _ask("  A base resume is already saved. Replace it? [y/N] ", "n").lower().startswith("y"):
        print("  Keeping the saved base resume.")
    else:
        for _ in range(3):
            path = _ask("  Path to your resume file (tip: drag the file into this window),\n"
                        "  or type  sample  to try JobPilot with a sample resume first: ")
            if not path:
                print("  Skipped. Later run:  jobpilot setup --resume \"My_Resume.docx\"")
                break
            try:
                cmd_setup(argparse.Namespace(resume=path, sample=False))
                break
            except SystemExit as err:
                print(f"  {err}")

    print("\n3) Work authorization line (optional, shown under your contact details)")
    current = SETTINGS.get("work_authorization") or ""
    print(f"  Example: U.S. Work Authorized | No Visa Sponsorship Needed" + (f"\n  Current: {current}" if current else ""))
    line = _ask("  Type the line, '-' to remove it, or press Enter to keep as is: ")
    write_default_settings()
    if line == "-":
        save_settings(work_authorization="")
    elif line:
        save_settings(work_authorization=line)

    print("\n4) Experience rules (optional): extra experience that appears only for matching jobs")
    if not RULES_FILE.exists() and _ask("  Create the rules file now? [y/N] ", "n").lower().startswith("y"):
        cmd_rules(argparse.Namespace(edit=True))

    print("\n5) Checking everything")
    cmd_check(argparse.Namespace())


def check_api_key():
    key = os.environ.get("ANTHROPIC_API_KEY", "").strip()
    if not key:
        windows = sys.platform.startswith("win")
        sys.exit("Your Anthropic API key is not set in this terminal.\n"
                 + ('  Set it permanently:  setx ANTHROPIC_API_KEY "sk-ant-api03-..."\n'
                    "  Then close this terminal and open a new one (setx only applies to new windows).\n"
                    "  Note: 'set' (without x) only lasts until you close the window."
                    if windows else
                    '  Add to ~/.zshrc or ~/.bashrc:  export ANTHROPIC_API_KEY="sk-ant-api03-..."\n'
                    "  Then open a new terminal.")
                 + "\n  Get a key at https://console.anthropic.com (Settings > API Keys).")
    if not key.startswith("sk-ant-"):
        print("Warning: ANTHROPIC_API_KEY does not look like an Anthropic API key (should start with sk-ant-).")


# ---------------------------------------------------------------------------
# Duplicate check: have I already prepared or applied for this job?
# ---------------------------------------------------------------------------
_TRACKING_PARAMS = re.compile(r"^(utm_.*|trk.*|ref.*|src|source|gh_src|lever-.*|trackingid|ebp|alternatechannel|"
                              r"position|pagenum|origin|from|sid|campaign.*|mc_.*|fbclid|gclid|li_fat_id|"
                              r"recommendedflavor|currentjobid)$", re.I)


def job_key(url: str) -> str | None:
    """A stable ID for a job posting, so the same job matches even when its link has tracking codes."""
    from urllib.parse import urlparse, parse_qs, urlencode
    url = (url or "").strip()
    if not re.match(r"https?://", url, re.I):
        return None
    parts = urlparse(url)
    host = parts.netloc.lower().removeprefix("www.").removeprefix("m.")
    path = parts.path.rstrip("/")
    query = {k.lower(): v for k, v in parse_qs(parts.query).items()}
    if "linkedin.com" in host:
        found = re.search(r"/jobs/view/(?:[^/]*?-)?(\d{6,})", path) or re.search(r"(\d{6,})", query.get("currentjobid", [""])[0])
        # a LinkedIn or Indeed link without a job number can't identify a posting: never match it
        return f"linkedin:{found.group(1)}" if found else None
    if "indeed." in host:
        jk = (query.get("jk") or query.get("vjk") or [""])[0]
        return f"indeed:{jk.lower()}" if jk else None
    if "dice.com" in host:
        found = re.search(r"/job-detail/([0-9a-f-]{8,})", path, re.I)
        if found:
            return f"dice:{found.group(1).lower()}"
    if "greenhouse.io" in host or "gh_jid" in query:
        jid = (query.get("gh_jid") or [""])[0] or (re.search(r"/jobs/(\d+)", path) or [None, ""])[1]
        if jid:
            return f"greenhouse:{jid}"
    if "lever.co" in host or "ashbyhq.com" in host:
        found = re.search(r"([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})", path, re.I)
        if found:
            return f"{host.split('.')[-2]}:{found.group(1).lower()}"
    kept = sorted((k, v[0]) for k, v in query.items() if not _TRACKING_PARAMS.match(k))
    return f"{host}{path.lower()}" + (f"?{urlencode(kept)}" if kept else "")


def _same_job_name(company: str, title: str) -> str:
    clean = lambda t: re.sub(r"[^a-z0-9]+", " ", re.sub(r"\[[^\]]*\]|\([^)]*\)", " ", (t or "").lower())).strip()
    return f"{clean(company)}|{clean(title)}"


def previous_applications(include_unlinked: bool = False):
    """Everything already created, keyed by job_key: from the tracker (with your Status) and from every
    application folder (which also covers runs made while the tracker was open in Excel)."""
    seen, unlinked = {}, []
    for content in OUTPUT_DIR.glob("*/content.json"):
        try:
            d = json.loads(content.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            continue
        key = job_key(d.get("source", ""))
        entry = {"company": d.get("company", ""), "title": d.get("job_title", ""),
                 "location": d.get("job_location", ""),
                 "date": d.get("applied_date", ""), "status": "", "folder": content.parent.name}
        if key:
            seen[key] = entry
        else:
            unlinked.append(entry)  # made from a pasted description or a file
    tracker = OUTPUT_DIR.resolve().parent / TRACKER_NAME
    if tracker.exists():
        try:
            from openpyxl import load_workbook
            ws = load_workbook(tracker)["Applications"]
            col = {c.value: i + 1 for i, c in enumerate(ws[1])}
            for r in range(2, ws.max_row + 1):
                link = ws.cell(row=r, column=col["Job Link"])
                key = job_key(link.hyperlink.target if link.hyperlink else str(link.value or ""))
                if not key:
                    continue
                prepared = ws.cell(row=r, column=col["Applied Date"]).value
                entry = seen.setdefault(key, {"folder": ws.cell(row=r, column=col["Folder"]).value or ""})
                entry.update(company=ws.cell(row=r, column=col["Company"]).value or entry.get("company", ""),
                             title=ws.cell(row=r, column=col["Job Title"]).value or entry.get("title", ""),
                             date=str(prepared)[:10] if prepared else entry.get("date", ""),
                             location=ws.cell(row=r, column=col["Location"]).value or entry.get("location", ""),
                             status=ws.cell(row=r, column=col["Status"]).value or "")
        except Exception as err:  # never block a run because the tracker could not be read
            print(f"Note: could not read {tracker.name} for the duplicate check ({err}).")
    return (seen, unlinked) if include_unlinked else seen


def _same_location(a: str, b: str) -> bool:
    """Unknown locations count as the same; otherwise one must contain the other (e.g. Remote / Remote (US))."""
    clean = lambda t: re.sub(r"[^a-z0-9]+", " ", (t or "").lower()).strip()
    a, b = clean(a), clean(b)
    if not a or not b or "not listed" in (a, b):
        return True
    return a in b or b in a


def describe_previous(p: dict) -> str:
    status = f", status: {p['status']}" if p.get("status") else ""
    return f"{p.get('company') or '?'} - {p.get('title') or '?'} (prepared {p.get('date') or '?'}{status})"


def _read_job_list(args) -> list[str]:
    jobs = []
    for item in args.job or []:
        jobs.extend(re.findall(r"https?://.+?(?=https?://|\s|$)", item) or [item])
    if args.file:
        path = Path(args.file)
        if not path.exists():
            sys.exit(f"Job list file not found: {path}")
        for line in path.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            links = re.findall(r"https?://.+?(?=https?://|\s|$)", line)
            jobs.extend(links or [line])  # several links stuck on one line are split apart
    seen, unique = set(), []
    for job in jobs:  # drop duplicates, keep order
        if job not in seen:
            seen.add(job)
            unique.append(job)
    return unique


def _read_job(job: str, batch: bool) -> tuple[str, str]:
    """Returns (job_text, source) for a URL or a job description file."""
    if re.match(r"https?://", job):
        return fetch_job(job, allow_skip=batch), job
    if Path(job).exists():
        return extract_text(Path(job)), job
    raise JobInputError("not a URL and no such file")


def _tailor_one(job_text: str, source: str, base_resume: str, rules: list[dict], check=None) -> dict:
    """The model work for one job (safe to run in parallel threads)."""
    active = active_rules(rules, job_text)
    for rule in active:
        log(f"Experience rule on: {rule['name']} (posting mentions {', '.join(rule['matched'])})")
    data = normalize_data(generate(base_resume, job_text, rules_prompt(active), after_analysis=check))
    data = apply_rules(data, active)
    # only keywords you have used at work (not ones you have only studied)
    boosted = ats_boost(data, without_learning_section(base_resume) + " " + rules_facts(active))
    if boosted:
        log("ATS: added keywords you have but the draft left out: " + ", ".join(boosted))
    data["source"] = source
    data["fit_score"] = fit_score(data)
    data["applied_date"] = date.today().isoformat()
    data["job_description_raw"] = job_text
    return data


def _report(folder: Path, data: dict, detailed: bool):
    saved = json.loads((folder / "content.json").read_text(encoding="utf-8"))
    print(f"\nDone: {data['job_title']} at {data['company']}  ->  {folder}")
    if not detailed:
        ats = saved.get("ats", {})
        line = f"  ATS: {'PASS' if ats.get('passed') else 'NEEDS ATTENTION'}"
        if saved.get("salary"):
            line += f"  |  Suggested ask: {money(saved['salary']['suggested'], saved['salary']['period'])}"
        print(line)
        return
    for file in sorted(folder.glob("*.docx")):
        print(f"  {file}")
    print_ats(saved.get("ats", {}))
    if saved.get("salary"):
        print(f"\nSalary: posted {saved['job_requirements'].get('compensation')}")
        print(f"  Suggested ask: {salary_text(saved['salary'])}")
    if data.get("gaps"):
        print("\nGaps to prepare for:\n  - " + "\n  - ".join(data["gaps"]))


def cmd_tailor(args):
    check_api_key()
    if not BASE_TEXT.exists():
        sys.exit('No base resume yet. Run:  jobpilot setup --resume "My_Resume.docx"')
    base_resume = BASE_TEXT.read_text(encoding="utf-8")
    rules = load_rules()
    jobs = _read_job_list(args)

    # Step 0: skip jobs you already prepared or applied for (compared by job link).
    duplicates, not_fit = [], []
    if jobs and not args.force:
        previous, batch_keys, fresh = previous_applications(), set(), []
        for job in jobs:
            key = job_key(job)
            if key and key in previous:
                duplicates.append((job, previous[key]))
            elif key and key in batch_keys:
                continue  # the same posting twice in this list
            else:
                fresh.append(job)
            if key:
                batch_keys.add(key)
        for job, prev in duplicates:
            print(f"Already done: {describe_previous(prev)}\n  {job}\n  -> applications/{prev.get('folder', '')}  (skipped)")
        if duplicates:
            print("To create it again anyway, add --force.\n")
        if not fresh:
            print("Nothing new to create.")
            return
        jobs = fresh

    # Also skip links an earlier run found to be a poor fit for your experience.
    min_fit = 0 if args.ignore_fit else int(SETTINGS.get("min_fit_score") or 0)
    if jobs and min_fit:
        skips, keep = previous_skips(), []
        for job in jobs:
            row = skips.get(job_key(job) or "")
            if row:
                print(f"Skipped before as not a fit ({row.get('Fit %')}%, {row.get('Date')}): "
                      f"{row.get('Company')} - {row.get('Job Title')}\n  {job}\n  Why: {row.get('Reason')}")
                not_fit.append((job, f"not a fit ({row.get('Fit %')}%), skipped on {row.get('Date')}"))
            else:
                keep.append(job)
        if not_fit:
            print("To create it anyway, add --ignore-fit.\n")
        if not keep:
            print("Nothing new to create.")
            return
        jobs = keep

    # Step 1: read every posting, one at a time (a blocked page may need you to paste it).
    prepared, failed = [], []
    if args.paste or not jobs:
        try:
            prepared.append((read_pasted(), "pasted description"))
        except JobInputError as err:
            sys.exit(f"Stopped: {err}. Please try again with the full posting.")
    else:
        batch = len(jobs) > 1
        for n, job in enumerate(jobs, 1):
            print(f"Reading job posting {n} of {len(jobs)}: {job}" if batch else "Reading the job posting...")
            try:
                prepared.append(_read_job(job, batch))
            except JobInputError as err:
                if not batch:
                    sys.exit(f"Stopped: {err}.")
                print(f"  Skipped: {err}")
                failed.append((job, str(err)))

    # Step 2: the model calls, in parallel when asked (each job takes about a minute).
    workers = max(1, min(args.parallel, 5, len(prepared) or 1))
    batch = len(prepared) > 1
    if batch:
        mode = f"{workers} at a time" if workers > 1 else "one after another"
        print(f"\nTailoring {len(prepared)} applications, {mode}...")
    else:
        print("Tailoring your resume, writing the cover letter and breaking down the requirements...")

    # Same job under another link, a pasted description or a file: caught right after step 1/3
    # (company and job title), before the resume and cover letter are written.
    by_name = {}
    if not args.force:
        linked, unlinked = previous_applications(include_unlinked=True)
        for prev in list(linked.values()) + unlinked:
            by_name.setdefault(_same_job_name(prev.get("company"), prev.get("title")), prev)

    def check(analysis):
        prev = by_name.get(_same_job_name(analysis.get("company"), analysis.get("job_title")))
        if not prev:
            check_fit(analysis)
            return
        if not _same_location(prev.get("location"), analysis.get("job_location")):
            log(f"Note: same company and title as {describe_previous(prev)}, but a different location "
                f"({analysis.get('job_location')} vs {prev.get('location')}): treating it as a different job.")
            check_fit(analysis)
            return
        raise AlreadyDone(prev)

    def check_fit(analysis):
        score = fit_score(analysis)
        if score is None:
            log("Note: the job analysis gave no fit score; continuing.")
            return
        log(f"  Role fit: {score}%" + (f" (minimum {min_fit}%)" if min_fit else ""))
        if min_fit and score < min_fit:
            raise NotAFit(analysis, score)

    def work(index, job_text, source):
        _TASK.label = f"[{index}/{len(prepared)}]" if batch else ""
        try:
            return _tailor_one(job_text, source, base_resume, rules, check)
        finally:
            _TASK.label = ""

    results = [None] * len(prepared)
    if workers > 1:
        from concurrent.futures import ThreadPoolExecutor
        with ThreadPoolExecutor(max_workers=workers) as pool:
            futures = {pool.submit(work, i + 1, text, src): i for i, (text, src) in enumerate(prepared)}
            for future, i in futures.items():
                try:
                    results[i] = future.result()
                except Exception as err:
                    results[i] = err
    else:
        for i, (text, src) in enumerate(prepared):
            try:
                results[i] = work(i + 1, text, src)
            except (AlreadyDone, NotAFit) as err:
                results[i] = err
            except Exception as err:
                if not batch:
                    raise
                results[i] = err

    # Step 3: save the Word files and update the tracker, one job at a time.
    done = []
    for (text, source), result in zip(prepared, results):
        if isinstance(result, AlreadyDone):
            prev = result.previous
            print(f"\nAlready done: {describe_previous(prev)}\n  {source}\n"
                  f"  -> applications/{prev.get('folder', '')}  (skipped after the job analysis; "
                  f"no resume or cover letter written. Use --force to create it anyway.)")
            duplicates.append((source, prev))
            continue
        if isinstance(result, NotAFit):
            a = result.analysis
            print(f"\nNot a fit, skipped: {a.get('job_title')} at {a.get('company')}\n  {source}\n  "
                  + describe_fit(a, result.score)
                  + f"\n  No resume or cover letter written (minimum fit is {min_fit}%). "
                    "Use --ignore-fit to create it anyway.")
            log_not_a_fit(result, source)
            not_fit.append((source, f"not a fit ({result.score}%)"))
            continue
        if isinstance(result, Exception):
            print(f"\nFailed: {source}\n  {result}")
            failed.append((source, str(result)))
            continue
        try:
            folder = write_application(result)
        except Exception as err:
            if not batch:
                raise
            print(f"\nFailed to save: {source}\n  {err}")
            failed.append((source, str(err)))
            continue
        _report(folder, result, detailed=not batch)
        done.append(folder)

    if len(jobs) > 1 or failed or duplicates or not_fit:
        extra = f", {len(duplicates)} already done" if duplicates else ""
        extra += f", {len(not_fit)} not a fit (see {skip_log_path()})" if not_fit else ""
        print(f"\nFinished: {len(done)} created, {len(failed)} failed or skipped{extra}.")
        for source, reason in failed:
            print(f"  - {source}: {reason}")
        if failed:
            print("Re-run just those links (or use --paste for a blocked page).")


def print_ats(ats: dict):
    if not ats.get("checks"):
        return
    print(f"\nATS compliance: {'PASS' if ats['passed'] else 'NEEDS ATTENTION'}")
    for c in ats["checks"]:
        print(f"  [{'x' if c['passed'] else ' '}] {c['check']}: {c['detail']}")
    if ats.get("missing"):
        print("  Posting keywords not in your resume (real gaps, not added): " + ", ".join(ats["missing"]))
    if ats.get("phrases_found") or ats.get("phrases_missing"):
        n = len(ats.get("phrases_found", [])); total = n + len(ats.get("phrases_missing", []))
        print(f"  Posting wording mirrored: {n} of {total} phrases")
        if ats.get("phrases_missing"):
            print("  Not mirrored (no matching experience found, or worded differently): "
                  + "; ".join(ats["phrases_missing"]))


def cmd_render(args):
    path = Path(args.json)
    data = json.loads(path.read_text(encoding="utf-8"))
    folder = write_application(data, path.parent)
    print(f"Rebuilt Word files in: {folder}")
    print_ats(json.loads((folder / "content.json").read_text(encoding="utf-8")).get("ats", {}))


def main():
    parser = argparse.ArgumentParser(
        prog="jobpilot", description="Tailored resume, cover letter and job requirements generator (Word format).")
    parser.add_argument("--version", action="version", version=f"JobPilot {__version__}")
    sub = parser.add_subparsers(dest="command", required=True)

    p_setup = sub.add_parser("setup", help="Save your current resume as the base resume")
    p_setup.add_argument("--resume", help="Your resume (.docx, .pdf, .md, .txt)")
    p_setup.add_argument("--sample", action="store_true",
                         help="Use the sample resume in examples/ (a fictional QA engineer) to try JobPilot")
    p_setup.set_defaults(func=cmd_setup)

    p_tailor = sub.add_parser("tailor", help="Create a resume, cover letter and requirements doc for a job")
    p_tailor.add_argument("job", nargs="*", help="One or more job URLs (LinkedIn/Indeed/company) or job description files")
    p_tailor.add_argument("--file", help="Text file with one job URL or file path per line")
    p_tailor.add_argument("--parallel", type=int, default=1, metavar="N",
                          help="Tailor up to N jobs at the same time (1-5, default 1 = one after another)")
    p_tailor.add_argument("--paste", action="store_true", help="Paste one job description into the terminal")
    p_tailor.add_argument("--force", action="store_true",
                          help="Create the application even if this job link was already done")
    p_tailor.add_argument("--ignore-fit", action="store_true",
                          help="Create the application even if the job matches your experience less than min_fit_score")
    p_tailor.set_defaults(func=cmd_tailor)

    p_track = sub.add_parser("track", help="Add/refresh every application in Job_Application_Tracker.xlsx")
    p_track.set_defaults(func=cmd_track)

    p_rules = sub.add_parser("rules", help="Create/check your experience rules (e.g. Playwright, AWS)")
    p_rules.add_argument("--edit", action="store_true", help="Open the rules file in a text editor")
    p_rules.set_defaults(func=cmd_rules)

    p_settings = sub.add_parser("settings", help="Show/edit your settings (work authorization line, page limit...)")
    p_settings.add_argument("--edit", action="store_true", help="Open the settings file in a text editor")
    p_settings.set_defaults(func=cmd_settings)

    p_init = sub.add_parser("init", help="Guided first-time setup (API key, resume, settings)")
    p_init.set_defaults(func=cmd_init)

    p_check = sub.add_parser("check", help="Check that everything is installed and set up")
    p_check.set_defaults(func=cmd_check)

    p_render = sub.add_parser("render", help="Rebuild the Word files from an edited content.json")
    p_render.add_argument("json")
    p_render.set_defaults(func=cmd_render)

    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
