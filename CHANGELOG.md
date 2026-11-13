# Changelog

## 2.6.0 (2026-11-13)

- **Use your Claude Code sign-in.** JobPilot now looks for a connection in this order: a local API key in the configuration file, then Claude Code (`claude`) installed and signed in, then the `ANTHROPIC_API_KEY` environment variable, and only then asks for a key. Through Claude Code no API key is needed: requests run with `claude -p` from an empty temporary folder, with Claude Code's tools turned off, and count toward that account's usage limits. Older Claude Code versions that lack some flags still work.
- New settings: `auth` (`auto`, `claude_cli` or `api_key`) and `claude_cli_path`. A local `anthropic_api_key` now comes before the environment variable.
- Every run prints which connection it uses; `jobpilot check` and `jobpilot init` show it too. When nothing is found, the message explains both options.
- 15 new tests, using a stand-in `claude` program.

## 2.5.0 (2026-11-12)

- **Results as tables.** A single job ends with a details table (company, job, location, fit, ATS result, keyword coverage, pay, suggested ask, sponsorship, folder and files), the ATS checklist with OK / FIX for each check, and a table of gaps. Several jobs end with one summary table: result, company, job title, fit, ATS, keyword coverage, suggested ask, and the folder or the reason it was skipped (for FIX, the failed checks are named). Tables wrap to the terminal width and use plain ASCII borders, so they show correctly in every terminal.

## 2.4.0 (2026-11-12)

- **Study material, only when asked.** `--study` adds a study guide for the job: prioritized topics (gaps first), what to learn, a hands-on exercise, and links to the official docs (each checked) and to course and video searches on YouTube, Pluralsight, LinkedIn Learning, Coursera and Udemy. `--questions` adds an Interview Q&A document with likely questions and model answers based on your resume. Both default to PDF. Turn them on for every job in the `study_material` section, make them later with `jobpilot study <folder>`, or skip them for one run with `--no-study`.
- **Configuration file.** `settings.json` now has sections you can edit with `jobpilot settings --edit`:
  - `documents`: cover letter on or off (off also saves a model call), the JobRequirements document on or off, and `docx`, `pdf` or `both` for each document. PDFs are made with Microsoft Word or LibreOffice.
  - `resume`: page limit and how to shorten.
  - `conditions`: minimum fit score, words that rule a posting out before any model call, job-title words to skip (e.g. VP, Director), allowed work arrangements and a minimum salary.
  - `anthropic_api_key`: optional, used when the environment variable is not set.
  - Older flat settings files are upgraded automatically, keeping your values (backup: `settings.json.bak`). `jobpilot check` reports configuration mistakes.
- **Automated tests.** A pytest suite (`tests/`) runs JobPilot end to end with a fake model and the sample resume: no API key or internet needed. GitHub Actions runs it on Windows and Linux with Python 3.10 and 3.13.
- **Sample resume.** `examples/sample_resume.docx` and `.txt`: a complete resume for a fictional QA engineer that shows the expected structure. `jobpilot setup --sample` (or typing `sample` in `jobpilot init`) loads it so you can try JobPilot end to end.

## 2.3.0 (2026-11-12)

First public release as **JobPilot** (previously `jobapply`).

- **Role-fit check.** The job analysis scores how well each posting matches your real work experience (0-100%). Jobs below `min_fit_score` (default 70%) are skipped before the resume and cover letter are written: VP or Director roles, "automation" that means ATM, kiosk, PLC or RPA work, a required clearance or license, or a mandatory stack you have never used. Skipped jobs are recorded in `applications/skipped_not_a_fit.csv`, so the same link is not analyzed (or paid for) twice. Use `--ignore-fit` to create one anyway.
- The fit score appears in the terminal and in the JobRequirements document.
- Renamed to JobPilot: the command is `jobpilot`, the main file is `jobpilot.py`, and the environment variables are `JOBPILOT_MODEL` and `JOBPILOT_WORK_AUTH` (the old `JOBAPPLY_*` names still work).
- Released under the MIT License.

## Earlier versions (as jobapply)

- **2.2:** duplicate detection (by job link, and by company and title after the analysis step), and wording alignment with the posting's own phrases.
- **2.1 and before:** tailored resume, cover letter and JobRequirements document in Word format; ATS checks and keyword boost; experience rules; two-page limit; salary guidance; the Excel application tracker; several jobs at once (`--file`, `--parallel`); guided setup (`init`, `check`, `settings`).
