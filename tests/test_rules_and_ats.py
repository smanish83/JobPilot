"""Experience rules and the ATS checklist."""
import copy

from conftest import load_fixture

RULES = [{
    "name": "Cypress",
    "when_job_mentions": ["Cypress"],
    "changes": [{"company": "Contoso", "replace_terms": {"Selenium WebDriver": "Cypress"},
                 "add_bullets": ["Built Cypress component tests for the checkout UI.", "TODO: write your bullet"]}],
    "add_skills": {"Test Automation": ["Cypress Component Testing"], "Cloud": ["AWS"]},
}]


def resume_data():
    data = copy.deepcopy(load_fixture("resume"))
    data.update(company="Example Corp", job_title="Senior QA Automation Engineer",
                job_requirements=copy.deepcopy(load_fixture("analysis")["job_requirements"]))
    return data


def test_rule_switches_on_only_when_posting_mentions_it(jp):
    assert jp.active_rules(RULES, "We use Playwright.") == []
    active = jp.active_rules(RULES, "Experience with Cypress is a plus.")
    assert [r["name"] for r in active] == ["Cypress"]


def test_rule_skips_todo_bullets(jp):
    active = jp.active_rules(RULES, "Cypress")
    assert active[0]["changes"][0]["add_bullets"] == ["Built Cypress component tests for the checkout UI."]


def test_apply_rules_replaces_terms_adds_bullets_and_skills(jp):
    data = jp.normalize_data(resume_data())
    data = jp.apply_rules(data, jp.active_rules(RULES, "Cypress"))
    contoso = next(j for j in data["experience"] if j["company"] == "Contoso Retail")
    assert contoso["bullets"][0] == "Built Cypress component tests for the checkout UI."
    assert any("Cypress" in b and "Selenium WebDriver" not in b for b in contoso["bullets"][1:2])
    skills = {g["category"]: g["items"] for g in data["skills"]}
    assert "Cypress Component Testing" in skills["Test Automation"]
    assert "Cloud" not in skills  # AWS is already listed, so no duplicate group is added
    assert data["rules_applied"] == ["Cypress (posting mentions Cypress)"]


def test_rules_never_touch_other_employers(jp):
    data = jp.apply_rules(jp.normalize_data(resume_data()), jp.active_rules(RULES, "Cypress"))
    northwind = next(j for j in data["experience"] if j["company"] == "Northwind Payments")
    assert not any("Cypress component" in b for b in northwind["bullets"])


def test_ats_check_passes_for_well_tailored_resume(jp):
    ats = jp.ats_check(jp.normalize_data(resume_data()))
    failed = [c["check"] for c in ats["checks"] if not c["passed"]]
    assert ats["passed"], failed
    assert ats["required_score"] == 100


def test_ats_check_flags_missing_keywords_and_bad_dates(jp):
    data = jp.normalize_data(resume_data())
    data["job_requirements"]["required_keywords"] = ["Kubernetes", "Terraform", "Go", "Playwright"]
    data["experience"][0]["dates"] = "03/2022 - now"
    ats = jp.ats_check(data)
    checks = {c["check"]: c for c in ats["checks"]}
    assert not checks["Required-keyword coverage of 80% or more"]["passed"]
    assert not checks["Consistent date format (Month YYYY - Month YYYY)"]["passed"]
    assert set(ats["required_missing"]) == {"Kubernetes", "Terraform", "Go"}
    assert not ats["passed"]


def test_ats_boost_adds_only_keywords_from_the_base_resume(jp):
    data = jp.normalize_data(resume_data())
    data["job_requirements"]["tech_stack"] += ["JMeter", "Kubernetes"]
    for group in data["skills"]:  # the draft left JMeter out, although the base resume has it
        group["items"] = [i for i in group["items"] if i != "JMeter"]
    base = jp.BASE_TEXT.read_text(encoding="utf-8")
    added = jp.ats_boost(data, jp.without_learning_section(base))
    assert "JMeter" in added          # in the sample resume
    assert "Kubernetes" not in added  # not in the sample resume: a real gap, never invented
