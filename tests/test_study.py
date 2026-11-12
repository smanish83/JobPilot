"""Study material: only when asked (configuration, --study/--questions, or `jobpilot study`)."""
import json
from pathlib import Path

from docx import Document

from test_config import configure


def app_folder():
    [folder] = [p for p in Path("applications").glob("*") if p.is_dir()]
    return folder


def docx_text(path):
    return "\n".join(p.text for p in Document(path).paragraphs)


def test_no_study_material_by_default(jp, fake_model, job_file, run):
    run("tailor", job_file)
    assert "STUDY_PROMPT" not in fake_model.calls and "QA_PROMPT" not in fake_model.calls
    assert not list(app_folder().glob("StudyGuide_*")) and not list(app_folder().glob("InterviewQA_*"))


def test_study_flag_makes_a_pdf_study_guide(jp, fake_model, job_file, run):
    run("tailor", job_file, "--study")
    assert fake_model.calls[-1] == "STUDY_PROMPT"
    folder = app_folder()
    assert list(folder.glob("StudyGuide_*.pdf"))          # default format: pdf
    assert not list(folder.glob("InterviewQA_*"))        # only when asked


def test_study_guide_links(jp, fake_model, job_file, run):
    configure(jp, study_material={"format": "docx"})
    run("tailor", job_file, "--study")
    content = json.loads((app_folder() / "content.json").read_text(encoding="utf-8"))
    topics = {t["topic"]: t for t in content["study"]["topics"]}
    k8s = [link["label"] for link in topics["Kubernetes basics for test environments"]["links"]]
    assert k8s == ["Official documentation", "YouTube", "Pluralsight", "LinkedIn Learning", "Coursera", "Udemy"]
    playwright = topics["Playwright fixtures and sharding"]["links"]
    assert all("made-up.example" not in link["url"] for link in playwright)  # a docs link that doesn't work is dropped
    youtube = next(link["url"] for link in playwright if link["label"] == "YouTube")
    assert youtube == "https://www.youtube.com/results?search_query=playwright+fixtures+sharding+tutorial"

    guide = next(app_folder().glob("StudyGuide_*.docx"))
    text = docx_text(guide)
    assert "Kubernetes basics for test environments" in text and "Practice:" in text
    assert text.index("Kubernetes basics") < text.index("Behavioral stories")  # High priority first
    targets = [r.target_ref for r in Document(guide).part.rels.values() if r.is_external]
    assert "https://kubernetes.io/docs/home/" in targets  # clickable links in the document


def test_only_the_sites_you_choose(jp, fake_model, job_file, run):
    configure(jp, study_material={"link_sources": ["YouTube"], "format": "docx"})
    run("tailor", job_file, "--study")
    content = json.loads((app_folder() / "content.json").read_text(encoding="utf-8"))
    labels = {link["label"] for t in content["study"]["topics"] for link in t["links"]}
    assert labels <= {"YouTube", "Official documentation"}


def test_links_can_be_turned_off(jp, fake_model, job_file, run):
    configure(jp, study_material={"learning_links": False})
    run("tailor", job_file, "--study")
    content = json.loads((app_folder() / "content.json").read_text(encoding="utf-8"))
    assert all("links" not in t for t in content["study"]["topics"])


def test_questions_flag_makes_the_interview_qa(jp, fake_model, job_file, run):
    configure(jp, study_material={"format": "both"})
    run("tailor", job_file, "--questions")
    assert fake_model.calls[-1] == "QA_PROMPT" and "STUDY_PROMPT" not in fake_model.calls
    folder = app_folder()
    assert list(folder.glob("InterviewQA_*.pdf")) and list(folder.glob("InterviewQA_*.docx"))
    text = docx_text(next(folder.glob("InterviewQA_*.docx")))
    assert "Q1. How do you handle flaky tests in Playwright?" in text
    assert "Questions About Your Gaps" in text
    assert text.index("Technical Questions") < text.index("Behavioral Questions")


def test_configuration_turns_both_on(jp, fake_model, job_file, run):
    configure(jp, study_material={"enabled": True, "interview_questions": True})
    run("tailor", job_file)
    assert fake_model.calls[-2:] == ["STUDY_PROMPT", "QA_PROMPT"]
    assert list(app_folder().glob("StudyGuide_*.pdf")) and list(app_folder().glob("InterviewQA_*.pdf"))


def test_no_study_overrides_the_configuration(jp, fake_model, job_file, run):
    configure(jp, study_material={"enabled": True, "interview_questions": True})
    run("tailor", job_file, "--no-study")
    assert fake_model.calls == ["ANALYZE_PROMPT", "RESUME_PROMPT", "COVER_LETTER_PROMPT"]


def test_skipped_jobs_get_no_study_material(jp, fake_model, job_file, run):
    fake_model.analysis["fit"] = {"score": 20, "blockers": [], "reason": "x"}
    run("tailor", job_file, "--study", "--questions")
    assert fake_model.calls == ["ANALYZE_PROMPT"]


def test_study_command_for_an_existing_application(jp, fake_model, job_file, run):
    run("tailor", job_file)
    fake_model.calls.clear()
    run("study", app_folder(), "--questions")
    assert fake_model.calls == ["STUDY_PROMPT", "QA_PROMPT"]
    folder = app_folder()
    assert list(folder.glob("StudyGuide_*")) and list(folder.glob("InterviewQA_*"))
    content = json.loads((folder / "content.json").read_text(encoding="utf-8"))
    assert content["study"]["topics"] and content["interview_qa"]["questions"]


def test_study_command_finds_the_folder_by_part_of_its_name(jp, fake_model, job_file, run):
    run("tailor", job_file)
    fake_model.calls.clear()
    run("study", "Example_Corp", "--questions-only")
    assert fake_model.calls == ["QA_PROMPT"]


def test_render_rebuilds_study_documents_without_model_calls(jp, fake_model, job_file, run):
    configure(jp, study_material={"format": "docx"})
    run("tailor", job_file, "--study", "--questions")
    folder = app_folder()
    for f in list(folder.glob("StudyGuide_*")) + list(folder.glob("InterviewQA_*")):
        f.unlink()
    fake_model.calls.clear()
    run("render", folder / "content.json")
    assert fake_model.calls == []
    assert list(folder.glob("StudyGuide_*.docx")) and list(folder.glob("InterviewQA_*.docx"))


def test_failed_study_material_never_loses_the_application(jp, fake_model, job_file, run, capsys):
    fake_model.extra["STUDY_PROMPT"] = {"topics": []}  # invalid twice in a row
    run("tailor", job_file, "--study")
    assert list(app_folder().glob("Resume_*.docx"))
    assert "study material could not be made" in capsys.readouterr().out


def test_bad_study_settings_are_reported(jp):
    configure(jp, study_material={"format": "epub", "link_sources": ["MySpace"]})
    problems = " ".join(jp.settings_problems())
    assert "study_material.format" in problems and "MySpace" in problems
