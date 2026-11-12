# Changelog

## 2.4.0 (unreleased)

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
