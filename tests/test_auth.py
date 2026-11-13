"""How JobPilot reaches Claude: a local key, the Claude Code CLI sign-in, then ANTHROPIC_API_KEY."""
import json
import sys
from pathlib import Path

import pytest

from test_config import configure

FAKE = Path(__file__).resolve().parent / "fake_claude.py"
SIGNED_IN = {"loggedIn": True, "authMethod": "claude.ai", "apiProvider": "firstParty"}


@pytest.fixture
def fake_cli(jp, fake_model, monkeypatch, tmp_path):
    """A fake `claude` program on this machine. Returns a helper to read the calls it received."""
    folder = tmp_path / "fakebin"
    folder.mkdir()
    if sys.platform.startswith("win"):
        exe = folder / "claude.cmd"
        exe.write_text(f'@"{sys.executable}" "{FAKE}" %*\n', encoding="utf-8")
    else:
        exe = folder / "claude"
        exe.write_text(f'#!/bin/sh\nexec "{sys.executable}" "{FAKE}" "$@"\n', encoding="utf-8")
        exe.chmod(0o755)
    log = tmp_path / "fake_claude_log.jsonl"
    replies = tmp_path / "replies.json"
    replies.write_text(json.dumps({jp.ANALYZE_PROMPT[:60]: fake_model.analysis,
                                   jp.RESUME_PROMPT[:60]: fake_model.resume,
                                   jp.COVER_LETTER_PROMPT[:60]: fake_model.letter}), encoding="utf-8")
    monkeypatch.setenv("FAKE_CLAUDE_LOG", str(log))
    monkeypatch.setenv("FAKE_CLAUDE_REPLIES", str(replies))
    monkeypatch.setenv("FAKE_CLAUDE_STATUS", json.dumps(SIGNED_IN))
    monkeypatch.setattr(jp, "find_claude_cli", lambda: str(exe))

    def calls(kind="-p"):
        if not log.exists():
            return []
        rows = [json.loads(line) for line in log.read_text(encoding="utf-8").splitlines()]
        return [r for r in rows if (kind in r["args"]) or (kind == "auth" and r["args"][:1] == ["auth"])]
    return calls


def app_folders():
    return sorted(p for p in Path("applications").glob("*") if p.is_dir())


# --- the order ---------------------------------------------------------------------------------
def test_claude_cli_is_used_when_signed_in(jp, fake_cli, job_file, run, capsys, monkeypatch):
    monkeypatch.delenv("ANTHROPIC_API_KEY")
    run("tailor", job_file)
    assert "Using: Claude Code CLI (your Claude account sign-in)" in capsys.readouterr().out
    requests = fake_cli()
    assert len(requests) == 3  # analysis, resume, cover letter
    args = requests[0]["args"]
    assert args[args.index("--model") + 1] == "claude-sonnet-5-5"
    assert "--output-format" in args and "--system-prompt-file" in args and "--tools" in args
    assert requests[0]["files"] == ["system.txt"]  # a fresh empty folder: no project files are loaded
    assert len(app_folders()) == 1


def test_cli_comes_before_the_environment_variable(jp, fake_cli, monkeypatch):
    monkeypatch.setenv("ANTHROPIC_API_KEY", "sk-ant-env")
    auth = jp.resolve_auth()
    assert auth["method"] == "claude_cli"
    assert auth["authMethod"] == "claude.ai"  # the account sign-in, not the API key from the environment


def test_api_key_in_the_environment_is_hidden_from_claude_code(jp, fake_cli, job_file, run, capsys, monkeypatch):
    monkeypatch.setenv("ANTHROPIC_API_KEY", "sk-ant-env")  # e.g. set with setx on Windows
    run("tailor", job_file)
    assert "Using: Claude Code CLI (your Claude account sign-in)" in capsys.readouterr().out
    assert fake_cli() and not any(call["env_key"] for call in fake_cli() + fake_cli("auth"))


def test_cli_without_a_sign_in_falls_back_to_the_env_key_directly(jp, fake_cli, monkeypatch):
    monkeypatch.setenv("FAKE_CLAUDE_STATUS", json.dumps({"loggedIn": False, "authMethod": "none"}))
    monkeypatch.setenv("ANTHROPIC_API_KEY", "sk-ant-env")
    auth = jp.resolve_auth()
    assert auth["method"] == "api_key" and auth["source"] == "env"


def test_local_key_comes_before_the_cli(jp, fake_cli, job_file, run, capsys):
    configure(jp, anthropic_api_key="sk-ant-local")
    run("tailor", job_file)
    assert "Using: local API key (configuration file)" in capsys.readouterr().out
    assert fake_cli() == []          # no CLI request was made
    assert fake_cli("auth") == []    # not even a sign-in check


def test_falls_back_to_the_api_key_when_cli_is_not_signed_in(jp, fake_cli, monkeypatch):
    monkeypatch.setenv("FAKE_CLAUDE_STATUS", json.dumps({"loggedIn": False, "authMethod": "none"}))
    monkeypatch.setenv("ANTHROPIC_API_KEY", "sk-ant-env")
    auth = jp.resolve_auth()
    assert auth["method"] == "api_key" and auth["source"] == "env"


def test_falls_back_to_the_api_key_when_cli_is_not_installed(jp, monkeypatch):
    monkeypatch.setenv("ANTHROPIC_API_KEY", "sk-ant-env")
    assert jp.resolve_auth()["source"] == "env"


def test_nothing_available_explains_both_options(jp, run, job_file, monkeypatch):
    monkeypatch.delenv("ANTHROPIC_API_KEY", raising=False)
    with pytest.raises(SystemExit) as stop:
        run("tailor", job_file)
    message = str(stop.value)
    assert "claude auth login" in message and "anthropic_api_key" in message


# --- the auth setting ----------------------------------------------------------------------------
def test_api_key_mode_never_uses_the_cli(jp, fake_cli, monkeypatch):
    monkeypatch.setenv("ANTHROPIC_API_KEY", "sk-ant-env")
    configure(jp, auth="api_key")
    assert jp.resolve_auth()["method"] == "api_key"
    assert fake_cli("auth") == []


def test_claude_cli_mode_never_uses_a_key(jp, fake_cli, monkeypatch, run, job_file):
    monkeypatch.setenv("FAKE_CLAUDE_STATUS", json.dumps({"loggedIn": False, "authMethod": "none"}))
    configure(jp, auth="claude_cli", anthropic_api_key="sk-ant-local")
    with pytest.raises(SystemExit, match="claude auth login"):
        run("tailor", job_file)


def test_bad_auth_value_is_reported(jp):
    configure(jp, auth="password")
    assert any("auth is" in p for p in jp.settings_problems())


def test_claude_cli_path_setting(jp, tmp_path):
    from conftest import REAL_FIND_CLAUDE_CLI
    configure(jp, claude_cli_path=str(tmp_path / "missing" / "claude"))
    assert REAL_FIND_CLAUDE_CLI() is None            # a path that doesn't exist is not used
    exe = tmp_path / "claude"
    exe.write_text("", encoding="utf-8")
    configure(jp, claude_cli_path=str(exe))
    assert REAL_FIND_CLAUDE_CLI() == str(exe)


# --- the CLI client -------------------------------------------------------------------------------
def test_cli_failure_explains_how_to_check_the_sign_in(jp, fake_cli, job_file, run, monkeypatch):
    monkeypatch.setenv("FAKE_CLAUDE_FAIL", "Authentication error")
    with pytest.raises(RuntimeError, match="claude auth status"):
        run("tailor", job_file)


def test_older_cli_without_the_extra_flags_still_works(jp, fake_cli, job_file, run, monkeypatch):
    monkeypatch.setenv("FAKE_CLAUDE_OLD", "1")
    run("tailor", job_file)
    assert len(app_folders()) == 1
    assert any("--tools" not in c["args"] for c in fake_cli())


def test_a_retry_sends_the_whole_conversation(jp):
    prompt = jp.ClaudeCLIClient._prompt([
        {"role": "user", "content": "Analyze this job."},
        {"role": "assistant", "content": "not json"},
        {"role": "user", "content": "That reply had a problem."}])
    assert prompt.index("Analyze this job.") < prompt.index("<your_previous_reply>") < prompt.index("<follow_up>")


def test_study_material_also_goes_through_the_cli(jp, fake_cli, fake_model, job_file, run, monkeypatch, tmp_path):
    replies = json.loads(Path(tmp_path / "replies.json").read_text(encoding="utf-8"))
    replies[jp.STUDY_PROMPT[:60]] = fake_model.extra["STUDY_PROMPT"]
    Path(tmp_path / "replies.json").write_text(json.dumps(replies), encoding="utf-8")
    run("tailor", job_file, "--study")
    assert len(fake_cli()) == 4
    assert list(app_folders()[0].glob("StudyGuide_*"))


def test_check_shows_the_connection(jp, fake_cli, run, capsys):
    run("check")
    assert "Connection to Claude: Claude Code CLI" in capsys.readouterr().out
