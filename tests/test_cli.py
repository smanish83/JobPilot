"""Command-line basics: version, setup, settings, check."""
import pytest


def test_version(jp, run, capsys):
    with pytest.raises(SystemExit) as exit_info:
        run("--version")
    assert exit_info.value.code == 0
    assert capsys.readouterr().out.strip() == f"JobPilot {jp.__version__}"


def test_setup_with_the_sample_resume(jp, run):
    jp.BASE_TEXT.unlink()
    run("setup", "--sample")
    text = jp.BASE_TEXT.read_text(encoding="utf-8")
    assert text.startswith("Alex Morgan")
    assert "Professional Experience" in text


def test_setup_with_own_docx(jp, run, tmp_path):
    from docx import Document
    doc = Document()
    for line in ["Jamie Lee", "QA Engineer", "Professional Summary", "Tester " * 120]:
        doc.add_paragraph(line)
    path = tmp_path / "Jamie_Lee_Resume.docx"
    doc.save(path)
    run("setup", "--resume", path)
    assert jp.BASE_TEXT.read_text(encoding="utf-8").startswith("Jamie Lee")


def test_setup_rejects_missing_file(jp, run):
    with pytest.raises(SystemExit, match="File not found"):
        run("setup", "--resume", "no_such_file.docx")


def test_tailor_without_api_key_explains_how_to_set_it(jp, run, monkeypatch, job_file):
    monkeypatch.delenv("ANTHROPIC_API_KEY", raising=False)
    with pytest.raises(SystemExit, match="API key is not set"):
        run("tailor", job_file)


def test_settings_lists_every_setting(jp, run, capsys):
    run("settings")
    out = capsys.readouterr().out
    for key in ("work_authorization", "model", "max_resume_pages", "min_fit_score"):
        assert key in out
    assert jp.SETTINGS_FILE.exists()
