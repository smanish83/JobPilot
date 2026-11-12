"""Unit tests for the small pure functions: names, links, keywords, salary, fit, normalizing replies."""
import pytest


# --- file and folder names ---------------------------------------------------------------
@pytest.mark.parametrize("text, expected", [
    ("Senior QA Engineer [Remote-US]", "Senior_QA_Engineer"),
    ("QA Engineer, Automation (Hybrid)", "QA_Engineer_Automation"),
    ("R&D / Test Lead", "R_D_Test_Lead"),
    ("", "Job"),
])
def test_slug_drops_tags_and_special_characters(jp, text, expected):
    assert jp.slug(text) == expected


# --- job links ---------------------------------------------------------------------------
@pytest.mark.parametrize("a, b", [
    ("https://www.linkedin.com/jobs/view/4123456789/?trackingId=abc&refId=x",
     "https://linkedin.com/jobs/view/4123456789"),
    ("https://www.linkedin.com/jobs/view/senior-qa-engineer-at-acme-4123456789",
     "https://www.linkedin.com/jobs/view/4123456789/"),
    ("https://www.indeed.com/viewjob?jk=ABC123&from=serp", "https://indeed.com/viewjob?jk=abc123"),
    ("https://boards.greenhouse.io/acme/jobs/1234567?gh_src=xyz", "https://boards.greenhouse.io/acme/jobs/1234567"),
    ("https://careers.acme.com/job/42?utm_source=linkedin", "https://careers.acme.com/job/42"),
])
def test_job_key_ignores_tracking_codes(jp, a, b):
    assert jp.job_key(a) == jp.job_key(b) is not None


@pytest.mark.parametrize("url", [
    "https://www.linkedin.com/jobs/search/?keywords=qa",   # no job number
    "https://www.indeed.com/jobs?q=qa",                    # no jk
    "job_description.txt",                                 # not a link
    "",
])
def test_job_key_none_when_no_posting_id(jp, url):
    assert jp.job_key(url) is None


def test_job_key_different_jobs_differ(jp):
    assert jp.job_key("https://www.linkedin.com/jobs/view/4123456789/") != \
        jp.job_key("https://www.linkedin.com/jobs/view/4123456780/")


# --- keywords -----------------------------------------------------------------------------
@pytest.mark.parametrize("keyword, text, expected", [
    ("CI/CD", "Built CI/CD pipelines", True),
    ("C#", "Wrote tests in C# and Java", True),
    ("Java", "Wrote tests in JavaScript", False),
    ("SQL", "MySQL and PostgreSQL", False),
    ("REST API", "rest api testing", True),
    ("", "anything", False),
])
def test_keyword_in_matches_whole_terms(jp, keyword, text, expected):
    assert jp.keyword_in(keyword, text) is expected


def test_phrase_present_tolerates_small_wording_changes(jp):
    text = "Collaborated with onshore/offshore teams and communicated testing status to stakeholders."
    assert jp.phrase_present("onshore/offshore teams", text)
    assert not jp.phrase_present("people management of 50 engineers", text)


# --- salary -------------------------------------------------------------------------------
@pytest.mark.parametrize("text, low, high, period", [
    ("$115,000 - $130,000", 115000, 130000, "year"),
    ("Up to $120K a year", None, 120000, "year"),
    ("$44-$54/hr", 44, 54, "hour"),
    ("$150,000", 150000, 150000, "year"),
])
def test_parse_salary(jp, text, low, high, period):
    parsed = jp.parse_salary(text)
    assert (parsed["min"], parsed["max"], parsed["period"]) == (low, high, period)


def test_parse_salary_none_when_no_pay(jp):
    assert jp.parse_salary("Competitive salary and great benefits") is None


def _suggest(jp, **req):
    return jp.salary_suggestion({"job_requirements": req})


def test_salary_narrow_range_asks_85_percent_up(jp):
    s = _suggest(jp, salary_min=120000, salary_max=140000, salary_period="year")
    assert s["suggested"] == 137000


def test_salary_wide_range_asks_70_percent_up(jp):
    s = _suggest(jp, salary_min=100000, salary_max=160000, salary_period="year")
    assert s["suggested"] == 142000


def test_salary_up_to_asks_97_percent(jp):
    s = _suggest(jp, compensation="Up to $150,000")
    assert s["suggested"] == 145500 or s["suggested"] == 146000


def test_salary_never_above_maximum(jp):
    s = _suggest(jp, salary_min=139000, salary_max=140000, salary_period="year")
    assert s["suggested"] <= 140000


def test_no_salary_suggestion_without_pay(jp):
    assert _suggest(jp, compensation="Not listed") is None


# --- locations and duplicate names --------------------------------------------------------
@pytest.mark.parametrize("a, b, same", [
    ("Remote", "Remote (US)", True),
    ("Dallas, TX", "Not listed", True),
    ("Dallas, TX", "Austin, TX", False),
    ("", "Austin, TX", True),
])
def test_same_location(jp, a, b, same):
    assert jp._same_location(a, b) is same


def test_same_job_name_ignores_tags_and_punctuation(jp):
    assert jp._same_job_name("Acme, Inc.", "QA Engineer [Remote]") == jp._same_job_name("acme inc", "QA Engineer")


# --- role fit ------------------------------------------------------------------------------
@pytest.mark.parametrize("fit, expected", [
    ({"score": 88, "blockers": []}, 88),
    ({"score": "75%", "blockers": []}, 75),
    ({"score": 90, "blockers": ["VP role"]}, 40),
    ({"score": 150}, 100),
    ({"score": "n/a"}, None),
    (None, None),
])
def test_fit_score(jp, fit, expected):
    assert jp.fit_score({"fit": fit}) == expected


def test_describe_fit_lists_blockers_and_reason(jp):
    text = jp.describe_fit({"fit": {"role_family": "ATM hardware automation", "seniority": "Match",
                                     "blockers": ["ATM firmware"], "reason": "Different profession."}}, 35)
    assert "35%" in text and "Blocker: ATM firmware" in text and "Different profession." in text


# --- normalizing model replies ----------------------------------------------------------------
def test_normalize_data_fixes_shapes_and_removes_visa_text(jp):
    data = jp.normalize_data({
        "company": "Acme", "job_title": "QA", "name": "A", "headline": "h", "summary": "s",
        "contact": "Austin, TX | (555) 010-0199 | a@example.com | No visa sponsorship needed",
        "skills": {"Testing": "Regression, Smoke"},
        "experience": {"company": "Acme", "title": "QA", "dates": "May 2020 - Present", "highlights": "Did things"},
    })
    assert "visa" not in data["contact"].lower()
    assert data["skills"] == [{"category": "Testing", "items": ["Regression", "Smoke"]}]
    assert data["experience"][0]["bullets"] == ["Did things"]


def test_normalize_data_defaults_missing_company_and_title(jp):
    data = jp.normalize_data({"contact": ""})
    assert data["company"] == "Company" and data["job_title"] == "Job"
