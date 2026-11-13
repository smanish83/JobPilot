"""End-to-end runs of `jobpilot tailor` with a fake model: files, tracker, duplicates and the fit check."""
import json
from pathlib import Path

from openpyxl import load_workbook


def app_folders():
    return sorted(p for p in Path("applications").glob("*") if p.is_dir())


def test_tailor_creates_the_three_documents_and_tracker_row(jp, fake_model, job_file, run):
    run("tailor", job_file)

    assert fake_model.calls == ["ANALYZE_PROMPT", "RESUME_PROMPT", "COVER_LETTER_PROMPT"]
    [folder] = app_folders()
    assert folder.name.endswith("_Example_Corp_Senior_QA_Automation_Engineer")
    names = sorted(p.name for p in folder.iterdir())
    assert names == ["CoverLetter_Example_Corp_Senior_QA_Automation_Engineer.docx",
                     "JobRequirements_Example_Corp_Senior_QA_Automation_Engineer.docx",
                     "Resume_Example_Corp_Senior_QA_Automation_Engineer.docx",
                     "content.json"]
    content = json.loads((folder / "content.json").read_text(encoding="utf-8"))
    assert content["ats"]["passed"]
    assert content["fit_score"] == 88
    assert content["salary"]["suggested"] == 137000

    ws = load_workbook("Job_Application_Tracker.xlsx")["Applications"]
    header = [c.value for c in ws[1]]
    row = {h: ws.cell(row=2, column=i + 1).value for i, h in enumerate(header)}
    assert row["Company"] == "Example Corp"
    assert row["Job Title"] == "Senior QA Automation Engineer"
    assert ws.max_row == 2


def test_resume_text_comes_from_the_model_reply(jp, fake_model, job_file, run):
    from docx import Document
    run("tailor", job_file)
    [folder] = app_folders()
    resume = Document(next(folder.glob("Resume_*.docx")))
    text = "\n".join(p.text for p in resume.paragraphs)
    assert "Alex Morgan" in text and "Northwind Payments" in text
    assert "Professional Experience" in text


def test_same_job_again_is_skipped_after_the_analysis(jp, fake_model, job_file, run, capsys):
    run("tailor", job_file)
    fake_model.calls.clear()
    run("tailor", job_file)
    assert fake_model.calls == ["ANALYZE_PROMPT"]  # stopped before the resume and cover letter
    assert "Already done" in capsys.readouterr().out
    assert len(app_folders()) == 1


def test_force_creates_it_again(jp, fake_model, job_file, run):
    run("tailor", job_file)
    run("tailor", job_file, "--force")
    assert len(app_folders()) == 2  # the second folder gets a _2 suffix


def test_known_link_is_skipped_without_any_model_call(jp, fake_model, run, capsys, monkeypatch):
    link = "https://www.linkedin.com/jobs/view/4123456789/"
    monkeypatch.setattr(jp, "_read_job", lambda job, batch: (Path("job_example_corp.txt").read_text(), job))
    Path("job_example_corp.txt").write_text("Example Corp posting", encoding="utf-8")
    run("tailor", link)
    fake_model.calls.clear()
    run("tailor", link + "?trackingId=abc123")
    assert fake_model.calls == []
    assert "Already done" in capsys.readouterr().out


def test_poor_fit_is_skipped_and_recorded(jp, fake_model, job_file, run, capsys):
    fake_model.analysis["fit"] = {"score": 92, "role_family": "Executive leadership (VP)",
                                  "seniority": "Above candidate", "blockers": ["VP of Quality, 60 reports"],
                                  "reason": "Executive scope."}
    run("tailor", job_file)
    out = capsys.readouterr().out
    assert fake_model.calls == ["ANALYZE_PROMPT"]
    assert "Not a fit, skipped" in out and "40%" in out
    assert app_folders() == []
    log = Path("applications/skipped_not_a_fit.csv").read_text(encoding="utf-8-sig")
    assert "VP of Quality, 60 reports" in log


def test_ignore_fit_creates_the_application_anyway(jp, fake_model, job_file, run):
    fake_model.analysis["fit"] = {"score": 30, "blockers": [], "reason": "Mostly hardware."}
    run("tailor", job_file, "--ignore-fit")
    assert len(app_folders()) == 1


def test_skipped_link_is_not_analyzed_twice(jp, fake_model, run, monkeypatch, capsys):
    link = "https://www.linkedin.com/jobs/view/4999999999/"
    monkeypatch.setattr(jp, "_read_job", lambda job, batch: ("Example Corp posting", job))
    fake_model.analysis["fit"] = {"score": 20, "blockers": ["ATM hardware"], "reason": "Different job."}
    run("tailor", link)
    fake_model.calls.clear()
    run("tailor", link)
    assert fake_model.calls == []
    assert "Skipped before as not a fit" in capsys.readouterr().out


def test_missing_fit_score_does_not_block(jp, fake_model, job_file, run):
    del fake_model.analysis["fit"]
    run("tailor", job_file)
    assert len(app_folders()) == 1


def test_several_jobs_in_parallel(jp, fake_model, run):
    for name in ("Alpha", "Beta", "Gamma"):
        Path(f"job_{name}.txt").write_text(f"{name} Corp posting", encoding="utf-8")
    companies = iter(["Alpha Corp", "Beta Corp", "Gamma Corp"])
    base_reply = fake_model.reply_for

    def reply_for(system):  # each job is a different company, so none is a duplicate
        reply = base_reply(system)
        return {**reply, "company": next(companies)} if system == jp.ANALYZE_PROMPT else reply

    fake_model.reply_for = reply_for
    run("tailor", "job_Alpha.txt", "job_Beta.txt", "job_Gamma.txt", "--parallel", "3")
    assert len(app_folders()) == 3
    assert fake_model.calls.count("ANALYZE_PROMPT") == 3


def test_render_rebuilds_without_model_calls(jp, fake_model, job_file, run):
    run("tailor", job_file)
    [folder] = app_folders()
    content = folder / "content.json"
    data = json.loads(content.read_text(encoding="utf-8"))
    data["summary"] = "Edited summary for Senior QA Automation Engineer roles with Playwright and CI/CD."
    content.write_text(json.dumps(data), encoding="utf-8")
    for docx in folder.glob("*.docx"):
        docx.unlink()
    fake_model.calls.clear()
    run("render", content)
    assert fake_model.calls == []
    assert len(list(folder.glob("*.docx"))) == 3


def test_a_list_file_given_without_file_is_read_as_a_list(jp, fake_model, run, capsys):
    Path("job_a.txt").write_text("Posting A", encoding="utf-8")
    Path("jobs.txt").write_text("# my jobs\njob_a.txt\nhttps://www.linkedin.com/jobs/view/4111111111/"
                                "https://www.linkedin.com/jobs/view/4222222222/\n", encoding="utf-8")
    assert jp._list_file_entries(Path("jobs.txt")) == [
        "job_a.txt", "https://www.linkedin.com/jobs/view/4111111111/", "https://www.linkedin.com/jobs/view/4222222222/"]


def test_a_job_description_file_is_not_mistaken_for_a_list(jp, job_file):
    assert jp._list_file_entries(job_file) is None
