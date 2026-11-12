"""Terminal output as tables."""
from pathlib import Path


def test_table_fits_the_width_and_wraps_long_cells(jp):
    text = jp.format_table(["Name", "Details"], [["Folder", "x" * 200], ["Fit", "88%"]], width=60)
    lines = text.splitlines()
    assert all(len(line) <= 60 for line in lines)
    assert "".join(line for line in lines if "x" in line).count("x") == 200  # nothing is cut off
    assert lines[0].startswith("+") and lines[1].startswith("| Name")


def test_table_without_wrapped_rows_is_compact(jp):
    text = jp.format_table(["A", "B"], [["1", "2"], ["3", "4"]], width=80)
    rules = [line for line in text.splitlines() if line.startswith("+-")]
    assert len(rules) == 2  # top and bottom only, no lines between the rows


def test_table_handles_multi_line_cells_and_none(jp):
    text = jp.format_table(["Files"], [["a.docx\nb.pdf"], [None]], width=40)
    assert "| a.docx" in text and "| b.pdf" in text


def test_ats_problems_names_the_failed_checks(jp):
    ats = {"required_score": 75, "checks": [
        {"check": "Required-keyword coverage of 80% or more", "passed": False, "detail": ""},
        {"check": "Consistent date format (Month YYYY - Month YYYY)", "passed": False, "detail": ""},
        {"check": "No unusual symbols or emoji", "passed": True, "detail": ""}]}
    assert jp.ats_problems(ats) == "fix: keywords 75%, dates"
    assert jp.ats_label(ats) == "FIX"


def test_single_job_shows_detail_tables(jp, fake_model, job_file, run, capsys):
    run("tailor", job_file)
    out = capsys.readouterr().out
    assert "| Company" in out and "Example Corp" in out
    assert "| Suggested ask" in out and "$137,000" in out
    assert "ATS compliance: PASS" in out and "| Required-keyword coverage of 80% or more" in out
    assert "Summary:" not in out  # one job: the detail tables are enough


def test_several_jobs_end_with_one_summary_table(jp, fake_model, job_file, run, capsys):
    run("tailor", job_file)                     # done once already
    Path("job_other.txt").write_text("Another posting", encoding="utf-8")
    base_reply = fake_model.reply_for
    replies = iter(["Example Corp", "Other Corp"])
    fake_model.reply_for = lambda system: (
        {**base_reply(system), "company": next(replies)} if system == jp.ANALYZE_PROMPT else base_reply(system))
    capsys.readouterr()
    run("tailor", job_file, "job_other.txt", "missing_file.txt")
    out = capsys.readouterr().out
    summary = out[out.index("Summary:"):]
    assert "1 created" in summary and "1 already done" in summary and "1 couldn't read" in summary
    created = next(line for line in summary.splitlines() if "| Created" in line)
    assert "Other Corp" in created and "88%" in created and "PASS" in created and "$137,000" in created
    assert "Use --force" in summary


def test_skips_are_summarized_with_their_reason(jp, fake_model, job_file, run, capsys):
    fake_model.analysis["fit"] = {"score": 30, "blockers": [], "reason": "Mostly ATM hardware."}
    run("tailor", job_file)
    out = capsys.readouterr().out
    summary = out[out.index("Summary:"):]
    assert "| Not a fit" in summary and "30%" in summary and "Mostly ATM hardware." in summary
