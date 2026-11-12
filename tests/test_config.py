"""The configuration file: sections, upgrading old files, documents, formats, conditions and the API key."""
import json
import os
from pathlib import Path

import pytest


def configure(jp, **sections):
    """Write a configuration file and load it, like a user editing it with `jobpilot settings --edit`."""
    jp.SETTINGS_FILE.write_text(json.dumps(sections), encoding="utf-8")
    jp.SETTINGS.clear()
    jp.SETTINGS.update(jp.load_settings())


def app_folders():
    return sorted(p for p in Path("applications").glob("*") if p.is_dir())


# --- loading -------------------------------------------------------------------------------
def test_defaults_when_there_is_no_file(jp):
    configure(jp)
    assert jp.cfg("documents.cover_letter") is True
    assert jp.doc_format("resume") == "docx"
    assert jp.max_pages() == 2
    assert jp.cfg("conditions.min_fit_score") == 70


def test_partial_sections_keep_the_other_defaults(jp):
    configure(jp, documents={"cover_letter": False})
    assert jp.cfg("documents.cover_letter") is False
    assert jp.cfg("documents.requirements_document") is True


def test_old_flat_settings_are_understood(jp):
    configure(jp, max_resume_pages=3, min_fit_score=60, shorten_role_first="Fabrikam")
    assert jp.max_pages() == 3
    assert jp.cfg("conditions.min_fit_score") == 60
    assert jp.cfg("resume.shorten_role_first") == "Fabrikam"


def test_old_file_is_upgraded_in_place_with_a_backup(jp, run):
    jp.SETTINGS_FILE.write_text(json.dumps({"work_authorization": "U.S. Work Authorized",
                                            "max_resume_pages": 2, "min_fit_score": 65}), encoding="utf-8")
    run("settings")
    upgraded = json.loads(jp.SETTINGS_FILE.read_text(encoding="utf-8"))
    assert upgraded["work_authorization"] == "U.S. Work Authorized"
    assert upgraded["conditions"]["min_fit_score"] == 65
    assert "min_fit_score" not in upgraded and "max_resume_pages" not in upgraded
    assert "documents" in upgraded
    assert jp.SETTINGS_FILE.with_suffix(".json.bak").exists()


def test_bad_values_are_reported(jp):
    configure(jp, documents={"resume_format": "rtf"}, conditions={"skip_job_titles_with": "VP"})
    problems = jp.settings_problems()
    assert any("resume_format" in p for p in problems)
    assert any("skip_job_titles_with" in p for p in problems)
    assert jp.doc_format("resume") == "docx"  # falls back safely


def test_json_error_is_explained(jp):
    jp.SETTINGS_FILE.write_text('{"model": "x",}', encoding="utf-8")
    with pytest.raises(SystemExit, match="JSON error"):
        jp.load_settings()


# --- API key -------------------------------------------------------------------------------
def test_api_key_from_configuration_file(jp, fake_model, job_file, run, monkeypatch):
    monkeypatch.delenv("ANTHROPIC_API_KEY")
    configure(jp, anthropic_api_key="sk-ant-from-config")
    run("tailor", job_file)
    assert os.environ["ANTHROPIC_API_KEY"] == "sk-ant-from-config"
    assert len(app_folders()) == 1


def test_environment_variable_wins_over_the_file(jp, monkeypatch):
    monkeypatch.setenv("ANTHROPIC_API_KEY", "sk-ant-env")
    configure(jp, anthropic_api_key="sk-ant-file")
    assert jp.api_key() == "sk-ant-env"


def test_settings_never_prints_the_whole_key(jp, run, capsys):
    configure(jp, anthropic_api_key="sk-ant-api03-SECRETSECRETSECRET-abcd")
    run("settings")
    out = capsys.readouterr().out
    assert "SECRETSECRETSECRET" not in out and "abcd" in out


# --- documents and formats -------------------------------------------------------------------
def test_no_cover_letter_saves_a_model_call(jp, fake_model, job_file, run):
    configure(jp, documents={"cover_letter": False})
    run("tailor", job_file)
    assert fake_model.calls == ["ANALYZE_PROMPT", "RESUME_PROMPT"]
    [folder] = app_folders()
    assert not list(folder.glob("CoverLetter_*"))
    assert list(folder.glob("Resume_*.docx")) and list(folder.glob("JobRequirements_*.docx"))


def test_no_requirements_document(jp, fake_model, job_file, run):
    configure(jp, documents={"requirements_document": False})
    run("tailor", job_file)
    [folder] = app_folders()
    assert not list(folder.glob("JobRequirements_*"))


def fake_pdf(monkeypatch, jp):
    def convert(path):
        pdf = path.with_suffix(".pdf")
        pdf.write_bytes(b"%PDF-1.4 fake")
        return pdf
    monkeypatch.setattr(jp, "convert_to_pdf", convert)


def test_pdf_only_resume(jp, fake_model, job_file, run, monkeypatch):
    fake_pdf(monkeypatch, jp)
    configure(jp, documents={"resume_format": "pdf"})
    run("tailor", job_file)
    [folder] = app_folders()
    assert list(folder.glob("Resume_*.pdf")) and not list(folder.glob("Resume_*.docx"))
    assert list(folder.glob("CoverLetter_*.docx"))  # still the default


def test_both_formats(jp, fake_model, job_file, run, monkeypatch):
    fake_pdf(monkeypatch, jp)
    configure(jp, documents={"resume_format": "both", "cover_letter_format": "both"})
    run("tailor", job_file)
    [folder] = app_folders()
    for prefix in ("Resume_", "CoverLetter_"):
        assert list(folder.glob(prefix + "*.pdf")) and list(folder.glob(prefix + "*.docx"))


def test_word_file_is_kept_when_pdf_cannot_be_made(jp, fake_model, job_file, run, monkeypatch, capsys):
    monkeypatch.setattr(jp, "convert_to_pdf", lambda path: None)
    configure(jp, documents={"resume_format": "pdf"})
    run("tailor", job_file)
    [folder] = app_folders()
    assert list(folder.glob("Resume_*.docx"))
    assert "could not make a PDF" in capsys.readouterr().out


def test_page_limit_from_configuration(jp, fake_model, job_file, run, capsys):
    configure(jp, resume={"max_pages": 1})
    run("tailor", job_file)
    assert "Resume was over 1 pages" in capsys.readouterr().out


# --- conditions --------------------------------------------------------------------------------
def test_skip_if_posting_mentions_makes_no_model_call(jp, fake_model, job_file, run, capsys):
    configure(jp, conditions={"skip_if_posting_mentions": ["security clearance", "Postman"]})
    run("tailor", job_file)
    assert fake_model.calls == []
    assert "The posting mentions: Postman" in capsys.readouterr().out
    assert app_folders() == []


def test_skip_job_titles_with(jp, fake_model, job_file, run, capsys):
    fake_model.analysis["job_title"] = "VP of Quality Engineering"
    configure(jp, conditions={"skip_job_titles_with": ["VP", "Director"]})
    run("tailor", job_file)
    out = capsys.readouterr().out
    assert fake_model.calls == ["ANALYZE_PROMPT"]
    assert 'job title contains "VP"' in out
    assert "Your conditions" in Path("applications/skipped_not_a_fit.csv").read_text(encoding="utf-8-sig")


def test_title_words_match_whole_words_only(jp, fake_model, job_file, run):
    fake_model.analysis["job_title"] = "Senior QA Automation Engineer (Vprotect team)"
    configure(jp, conditions={"skip_job_titles_with": ["VP"]})
    run("tailor", job_file)
    assert len(app_folders()) == 1


def test_allowed_work_arrangements(jp, fake_model, job_file, run, capsys):
    fake_model.analysis["job_requirements"]["work_arrangement"] = "On-site"
    configure(jp, conditions={"allowed_work_arrangements": ["Remote", "Hybrid"]})
    run("tailor", job_file)
    assert "work arrangement is On-site" in capsys.readouterr().out
    assert app_folders() == []


def test_unlisted_work_arrangement_is_not_skipped(jp, fake_model, job_file, run):
    fake_model.analysis["job_requirements"]["work_arrangement"] = "Not listed"
    configure(jp, conditions={"allowed_work_arrangements": ["Remote"]})
    run("tailor", job_file)
    assert len(app_folders()) == 1


def test_min_salary(jp, fake_model, job_file, run, capsys):
    configure(jp, conditions={"min_salary": 150000})  # the posting tops out at $140,000
    run("tailor", job_file)
    assert "below your minimum of $150,000" in capsys.readouterr().out
    assert app_folders() == []


def test_min_salary_with_hourly_pay(jp):
    configure(jp, conditions={"min_salary": 100000})
    low = {"job_title": "QA", "job_requirements": {"compensation": "$40-$45/hr"}}
    high = {"job_title": "QA", "job_requirements": {"compensation": "$60-$70/hr"}}
    assert jp.condition_failures(low) and not jp.condition_failures(high)


def test_ignore_fit_also_ignores_conditions(jp, fake_model, job_file, run):
    configure(jp, conditions={"min_salary": 500000, "skip_if_posting_mentions": ["Postman"]})
    run("tailor", job_file, "--ignore-fit")
    assert len(app_folders()) == 1


def test_fit_check_can_be_turned_off(jp, fake_model, job_file, run):
    fake_model.analysis["fit"] = {"score": 10, "blockers": [], "reason": "x"}
    configure(jp, conditions={"min_fit_score": 0})
    run("tailor", job_file)
    assert len(app_folders()) == 1
