"""Shared test setup.

Every test runs JobPilot in a temporary folder with its own fake home directory, a copy of the sample
resume as the base resume, and a fake Anthropic client, so no test touches your real files or the API.
"""
import copy
import json
import os
import sys
import tempfile
from pathlib import Path
from types import SimpleNamespace

import pytest

ROOT = Path(__file__).resolve().parent.parent
FIXTURES = Path(__file__).resolve().parent / "fixtures"

# JobPilot reads its settings when it is imported, so point HOME at an empty folder first.
_IMPORT_HOME = tempfile.mkdtemp(prefix="jobpilot-tests-")
os.environ["HOME"] = os.environ["USERPROFILE"] = _IMPORT_HOME
for name in [k for k in os.environ if k.startswith(("JOBPILOT_", "JOBAPPLY_"))]:
    del os.environ[name]
sys.path.insert(0, str(ROOT))

import jobpilot  # noqa: E402


def load_fixture(name: str) -> dict:
    return json.loads((FIXTURES / f"{name}.json").read_text(encoding="utf-8"))


class FakeModel:
    """Stands in for anthropic.Anthropic(). Replies are chosen by the system prompt of each call."""

    def __init__(self):
        self.calls = []
        self.analysis = load_fixture("analysis")
        self.resume = load_fixture("resume")
        self.letter = load_fixture("letter")
        self.extra = {"STUDY_PROMPT": load_fixture("study"), "QA_PROMPT": load_fixture("interview_qa")}

    def reply_for(self, system: str) -> dict:
        if system == jobpilot.ANALYZE_PROMPT:
            return self.analysis
        if system == jobpilot.RESUME_PROMPT:
            return self.resume
        if system == jobpilot.COVER_LETTER_PROMPT:
            return self.letter
        for name, reply in self.extra.items():
            if system == getattr(jobpilot, name, None):
                return reply
        raise AssertionError(f"unexpected system prompt: {system[:60]!r}")

    def kind(self, system: str) -> str:
        for name in ("ANALYZE_PROMPT", "RESUME_PROMPT", "COVER_LETTER_PROMPT", *self.extra):
            if system == getattr(jobpilot, name, None):
                return name
        return "?"

    # anthropic.Anthropic() -> client.messages.stream(**kw) -> context manager
    def __call__(self, *args, **kwargs):
        return SimpleNamespace(messages=SimpleNamespace(stream=self._stream))

    def _stream(self, **kw):
        model = self

        class Stream:
            def __enter__(self):
                return self

            def __exit__(self, *exc):
                return False

            def get_final_message(self):
                model.calls.append(model.kind(kw["system"]))
                text = json.dumps(model.reply_for(kw["system"]))
                return SimpleNamespace(content=[SimpleNamespace(type="text", text=text)], stop_reason="end_turn")

        return Stream()


@pytest.fixture
def fake_model(monkeypatch):
    model = FakeModel()
    monkeypatch.setitem(sys.modules, "anthropic", SimpleNamespace(Anthropic=model))
    monkeypatch.setenv("ANTHROPIC_API_KEY", "sk-ant-test-key")
    return model


@pytest.fixture
def jp(tmp_path, monkeypatch):
    """The jobpilot module, isolated in tmp_path: fake home, sample base resume, default settings."""
    home = tmp_path / "home" / ".resume_agent"
    home.mkdir(parents=True)
    work = tmp_path / "work"
    work.mkdir()
    monkeypatch.chdir(work)
    monkeypatch.setenv("COLUMNS", "220")  # wide tables, so each test can find a row on one line
    monkeypatch.setattr(jobpilot, "HOME", home)
    monkeypatch.setattr(jobpilot, "SETTINGS_FILE", home / "settings.json")
    monkeypatch.setattr(jobpilot, "BASE_TEXT", home / "base_resume.txt")
    monkeypatch.setattr(jobpilot, "RULES_FILE", home / "experience_rules.json")
    monkeypatch.setattr(jobpilot, "LAST_REPLY", home / "last_model_reply.txt")
    monkeypatch.setattr(jobpilot, "SETTINGS", copy.deepcopy(jobpilot.DEFAULT_SETTINGS))
    # Page counting through Word or LibreOffice is slow and machine-specific: use the estimate.
    monkeypatch.setattr(jobpilot, "_pages_with_word", lambda path: None)
    monkeypatch.setattr(jobpilot, "_pages_with_libreoffice", lambda path: None)
    # Never reach the internet: official-docs links count as working only for these test URLs.
    monkeypatch.setattr(jobpilot, "_url_works", lambda url: url.startswith("https://kubernetes.io/"))
    # PDFs: a tiny stand-in file instead of Word or LibreOffice.
    def fake_pdf(path):
        pdf = path.with_suffix(".pdf")
        pdf.write_bytes(b"%PDF-1.4 test")
        return pdf
    monkeypatch.setattr(jobpilot, "convert_to_pdf", fake_pdf)
    (home / "base_resume.txt").write_text(
        (ROOT / "examples" / "sample_resume.txt").read_text(encoding="utf-8"), encoding="utf-8")
    return jobpilot


@pytest.fixture
def job_file(jp):
    path = Path("job_example_corp.txt")
    path.write_text((FIXTURES / "job_example_corp.txt").read_text(encoding="utf-8"), encoding="utf-8")
    return path


@pytest.fixture
def run(jp, monkeypatch):
    """Run a JobPilot command line, e.g. run("tailor", "job.txt")."""
    def _run(*argv):
        monkeypatch.setattr(sys, "argv", ["jobpilot", *map(str, argv)])
        jp.main()
    return _run
