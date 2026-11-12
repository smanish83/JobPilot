# JobPilot: How to Use

This guide assumes setup is done (see SETUP.md). Run every command from a terminal opened **in the JobPilot folder**. In PowerShell type `.\jobpilot`, on Mac/Linux `./jobpilot`.

---

## 1. Your everyday workflow

1. **Find a job** and copy its link.
2. **Run** `jobpilot tailor "<job link>"`.
3. **Read the three documents** in the new folder under `applications\`. Fix anything that isn't right.
4. **Apply.** Upload the resume and cover letter. Use the suggested salary ask if the form asks for one.
5. **Update the tracker** (`Job_Application_Tracker.xlsx`): set *Status* to `Applied` and fill in *Date Applied* and *Salary Requested*.

## 2. Commands

| Command | What it does |
|---|---|
| `jobpilot init` | Guided setup: API key, resume, work-authorization line |
| `jobpilot check` | Checks everything is installed and set up |
| `jobpilot setup --resume "My_Resume.docx"` | Saves (or replaces) your base resume |
| `jobpilot setup --sample` | Uses the sample resume, so you can try JobPilot before setting up your own |
| `jobpilot tailor "<link>"` | Creates one application |
| `jobpilot tailor "<link1>" "<link2>" "<link3>"` | Several applications, one after another |
| `jobpilot tailor "<link1>" "<link2>" --parallel 3` | Several applications, up to 3 at the same time |
| `jobpilot tailor --file jobs.txt --parallel 3` | Every link or file listed in `jobs.txt` |
| `jobpilot tailor job_description.txt` | Uses a saved job description (.txt, .docx, .pdf) |
| `jobpilot tailor --paste` | Paste a job description into the terminal |
| `jobpilot tailor "<link>" --force` | Creates the application even if that job was already done |
| `jobpilot tailor "<link>" --ignore-fit` | Creates the application even if the job fails the fit check or your conditions |
| `jobpilot rules` / `jobpilot rules --edit` | Shows / edits your experience rules |
| `jobpilot settings` / `jobpilot settings --edit` | Shows / edits your configuration file (documents, formats, page limit, conditions, API key) |
| `jobpilot tailor "<link>" --study --questions` | Also makes a study guide and an interview Q&A for the job (see section 10) |
| `jobpilot study "applications\<folder>"` | Study guide (and `--questions`) for an application you already have |
| `jobpilot render "applications\<folder>\content.json"` | Rebuilds the Word files after you edit `content.json` (no API cost) |
| `jobpilot track` | Rebuilds or catches up the tracker from all application folders |
| `jobpilot --version` | Shows the version |

## 3. Your base resume

`jobpilot setup` saves the text of your resume as your **base resume**:
`C:\Users\<you>\.resume_agent\base_resume.txt` (Windows) or `~/.resume_agent/base_resume.txt` (Mac/Linux).

- **Everything JobPilot writes comes from this file** (plus your experience rules). It is never allowed to invent experience.
- **Make it complete.** Open it in Notepad or any text editor and add every project, tool, metric and achievement you can honestly claim, even ones your current resume leaves out. JobPilot picks what fits each job and keeps the result to two pages.
- **Sample resume.** `examples\sample_resume.docx` (and `.txt`) is a complete resume for a fictional QA engineer, Alex Morgan. It shows the structure JobPilot reads best. You can copy it as a template, or run `jobpilot setup --sample` to try JobPilot end to end before you use your own resume. **Never apply with it.**
- **Keep the structure:** the section names (`Professional Summary`, `Core Skills`, `Professional Experience`, `Education`), then each job as a "Company - Location" line, a "Title | Month YYYY - Month YYYY" line and one achievement per line.
- **Careful:** running `jobpilot setup` again replaces this file with the text of the Word resume you give it. Edits you made only in `base_resume.txt` are lost unless you also put them in that Word file.
- **Close the file in your editor when you're done.** An open editor window can later save an old copy over your changes.

## 4. Tailoring

### One job
```
jobpilot tailor "https://www.linkedin.com/jobs/view/1234567890/"
```
LinkedIn links usually work directly. Indeed and some company sites block automatic reading. JobPilot then asks you to **paste the description**: copy the full text from your browser, paste it, then press **Ctrl+Z then Enter** (Windows) or **Ctrl+D** (Mac/Linux).

When it finishes, the results are shown as tables: the application (company, job, location, role fit, ATS result, keyword coverage, pay, suggested ask, sponsorship, folder and files), the ATS checklist with **OK** or **FIX** for each check (so you can see exactly why a resume says *NEEDS ATTENTION*), and the gaps to prepare for.

### Already applied? Duplicates are skipped
Before creating anything, JobPilot checks each job link against your tracker and your `applications` folder. If you've already prepared or applied for that job, it shows you when, the status, and the folder, then skips it, with no API cost. Links are matched by the posting's ID, so the same job still matches when its link carries different tracking codes (for example LinkedIn's `?trackingId=...`). This works for LinkedIn, Indeed, Dice, Greenhouse, Lever, Ashby and most company career sites.

- To create it again anyway, add `--force`.
- The same job under a *different* link (for example on Dice instead of LinkedIn), as a pasted description, or from a file is caught right after the first step identifies the company and job title. JobPilot then stops before writing the resume and cover letter, so only one of the three model calls is used.
- Same company and title but a *different location* counts as a separate opening and is created, with a note.

### Not a fit? Skipped before writing anything
Some postings share words with your experience but are a different job: a VP or Director role, "automation" that means ATM, kiosk, industrial/PLC or RPA work, a required clearance or license you don't have, or a mandatory core stack you've never used. Applying to them wastes your time.

So the first step also scores how well the job matches your **real work experience** (0-100%). Self-directed learning doesn't count. Any hard blocker caps the score at 40%. If the score is below **70%** (setting `min_fit_score`), JobPilot stops right there: no resume or cover letter is written, and only one of the three model calls is used. It shows you why:

```
Not a fit, skipped: ATM Automation Engineer at Delta Bank
  Role fit: 35% (ATM/banking hardware automation)
  Blocker: Core work is ATM hardware and firmware validation, not software test automation
  No resume or cover letter written (minimum fit is 70%). Use --ignore-fit to create it anyway.
```

- Skipped jobs are listed in `applications\skipped_not_a_fit.csv` (date, company, title, fit %, reason, link). Open it in Excel to review them.
- If you run the same link again later, it is skipped at once, with no API cost. To check it again (for example, after adding experience to your base resume), delete its line from that file or use `--ignore-fit`.
- The fit score of every job you *do* create is shown in the terminal and in the requirements document.
- Too strict or too loose? Change `min_fit_score` in the `conditions` section of the configuration file (for example 60 or 80). Set it to 0 to turn the check off.
- You can add your own rules too, such as skipping VP or Director titles, on-site jobs, or pay below your minimum. See section 9.

### Several jobs
```
jobpilot tailor "<link 1>" "<link 2>" "<link 3>" --parallel 3
```
- On Windows Command Prompt, keep **all links on one line**. For long lists, use a file instead (below).
- Without `--parallel`, the jobs run one after another. `--parallel N` runs up to N at a time (maximum 5). 2 or 3 is a good choice, because higher values can hit your API rate limit.
- Postings are read first, one by one (you may be asked to paste a blocked one, or press Ctrl+Z / Ctrl+D right away to skip it). Then the applications are created. A failure never stops the other jobs.
- At the end, one **summary table** shows every job on its own row:

```
Summary: 2 created, 1 already done, 1 not a fit
+---+--------------+-------------+---------------+-----+------+----------+----------+----------------------+
| # | Result       | Company     | Job title     | Fit | ATS  | Keywords | Ask      | Folder / reason      |
+===+==============+=============+===============+=====+======+==========+==========+======================+
| 1 | Created      | Quo         | QA Automation | 86% | FIX  | 75%      | $140,000 | Quo_QA_Automation_En |
|   |              |             | Engineer      |     |      |          |          | gineer_Web           |
|   |              |             | (SDET), Web   |     |      |          |          | fix: keywords 75%    |
+---+--------------+-------------+---------------+-----+------+----------+----------+----------------------+
| 2 | Created      | Acme Health | Senior SDET   | 91% | PASS | 94%      | $152,000 | Acme_Health_Senior_S |
|   |              |             |               |     |      |          |          | DET                  |
+---+--------------+-------------+---------------+-----+------+----------+----------+----------------------+
| 3 | Already done | Contoso     | QA Engineer   | -   | -    | -        | -        | prepared 2026-11-05, |
|   |              |             |               |     |      |          |          | status: Applied      |
+---+--------------+-------------+---------------+-----+------+----------+----------+----------------------+
| 4 | Not a fit    | Delta Bank  | ATM           | 35% | -    | -        | -        | Core work is ATM     |
|   |              |             | Automation    |     |      |          |          | hardware and         |
|   |              |             | Engineer      |     |      |          |          | firmware validation  |
+---+--------------+-------------+---------------+-----+------+----------+----------+----------------------+
```

  **ATS** is PASS or FIX; for FIX, the last column names the failed checks (for example `keywords 75%`). Run that job on its own, or open its JobRequirements document, for the full checklist. Tables wrap to fit your terminal window; a wider window shows more on each line.

### A list of jobs in a file
Create `jobs.txt` with one link (or description file) per line; lines starting with `#` are ignored (see `examples\jobs.txt`). Then run:
```
jobpilot tailor --file jobs.txt --parallel 3
```

Each job takes about a minute and makes three model calls (one, if it turns out to be a duplicate or not a fit): one to analyze the job, one to tailor the resume, one to write the cover letter.

## 5. What you get

A folder for each job, named with the date and the job:
`applications\2026-10-29_Acme_Corp_Senior_SDET\`

| File | Contents |
|---|---|
| **Resume_...docx** | Tailored resume: headline with the posting's job title, summary, Core Skills, experience, Professional Development, Education |
| **CoverLetter_...docx** | 250-350 words, written from the tailored resume, addressed to the hiring manager |
| **JobRequirements_...docx** | Job details (location, remote/hybrid, pay, sponsorship), each requirement marked **Meets / Partial / Gap** with evidence from your resume, keywords found and missing, the ATS checklist, gaps to prepare for, likely interview topics, and the **full original posting** (useful if it is taken down before your interview) |
| **StudyGuide_...pdf**, **InterviewQA_...pdf** | Only when you ask for them: topics to learn with course and video links, and likely interview questions with answers (section 10) |
| **content.json** | Everything that was generated. Edit it and run `jobpilot render` to rebuild the Word files without another API call. |

Folder and file names drop bracketed tags such as `[Remote-US]` and special characters. Running the same job twice on the same day creates a `_2` folder instead of overwriting the first.

## 6. ATS checks

Most employers screen resumes with an Applicant Tracking System (ATS) before a person reads them. Every resume is checked for:

- **Keyword coverage:** at least 80% of the posting's *required* keywords appear in your resume. Keywords you genuinely have but the draft left out are added automatically (only if they appear in your base resume or rules). Keywords you don't have are listed as gaps, never added.
- **Job title:** the posting's title appears in your headline or summary.
- **Layout:** contact details in the page body, standard section headings, a single column with no tables or graphics, a standard font, real bullets, consistent dates, no unusual symbols, a sensible length, and the .docx format.

**Wording alignment.** Job postings describe the same experience in different words: "team player", "collaborates with onshore/offshore teams", "healthcare domain experience", "communicate testing status to stakeholders". JobPilot picks out these phrases from each posting. Wherever your base resume shows that experience in other words, the tailored resume uses the posting's exact phrase. It never uses a phrase you have no evidence for. The requirements document lists which phrases were matched and which weren't, so you can add real experience to your base resume if something is missing.

The result (**PASS** or **NEEDS ATTENTION**) is shown in the terminal, in the requirements document and in the tracker.

## 7. Two-page limit

After building a resume, JobPilot counts its pages (exactly with Word on Windows, otherwise estimated). If it is too long:
1. The oldest role (or the one named in your settings) is cut to its title line.
2. If it is still too long and you list a higher degree (Master's, MBA, PhD), the Bachelor's degree is removed.

If it is *still* too long, you are told, and you can shorten older bullets in `content.json` and run `render`. The page limit and both steps can be changed in the configuration file (`resume` section).

## 8. Experience rules

Rules add experience **only when a posting asks for it**. For example: *"If the job mentions Playwright, describe my Robot Framework work at Acme as Playwright work and add my Playwright bullets."*

Run `jobpilot rules --edit`. Each rule has:

| Field | Meaning |
|---|---|
| `when_job_mentions` | The rule switches on if the posting mentions **any** of these words |
| `changes[].company` | Part of the employer name as written in your resume (`Acme` matches `Acme Health Inc.`) |
| `changes[].replace_terms` | Words to swap in that job's bullets, e.g. `{"Robot Framework": "Playwright"}` |
| `changes[].add_bullets` | Achievements to add to that job. Bullets that still contain `TODO` are ignored. |
| `add_skills` | Skills to add under a Core Skills category |

See `examples\experience_rules.example.json` for a filled-in example. Run `jobpilot rules` to see which bullets are ready.

**Only write experience you really have.** Interviewers will ask you about every line.

## 9. Configuration file

Everything you can change lives in one file: `.resume_agent\settings.json` in your home folder. Open it with `jobpilot settings --edit`, or see the current values with `jobpilot settings`. It is plain JSON: keep the quotes, commas and brackets. `jobpilot check` tells you if something is wrong. A filled-in example is `examples\settings.example.json`.

```json
{
  "anthropic_api_key": "",
  "model": "claude-sonnet-5-5",
  "work_authorization": "U.S. Work Authorized | No Visa Sponsorship Needed",
  "documents": {
    "cover_letter": true,
    "requirements_document": true,
    "resume_format": "both",
    "cover_letter_format": "pdf",
    "requirements_format": "docx"
  },
  "resume": {
    "max_pages": 2,
    "shorten_role_first": "",
    "drop_bachelors_if_higher_degree": true
  },
  "conditions": {
    "min_fit_score": 70,
    "skip_if_posting_mentions": ["security clearance", "TS/SCI"],
    "skip_job_titles_with": ["VP", "Vice President", "Director", "Intern"],
    "allowed_work_arrangements": ["Remote", "Hybrid"],
    "min_salary": 110000
  },
  "study_material": {
    "enabled": false,
    "learning_links": true,
    "link_sources": ["YouTube", "Pluralsight", "LinkedIn Learning", "Coursera", "Udemy"],
    "interview_questions": false,
    "question_count": 25,
    "format": "pdf"
  }
}
```

**General**

| Setting | Default | Meaning |
|---|---|---|
| `anthropic_api_key` | `""` | Your API key. Used only when the `ANTHROPIC_API_KEY` environment variable is not set. The file is plain text, so keep it private and never share or commit it. |
| `model` | `claude-sonnet-5-5` | The Anthropic model used |
| `work_authorization` | `""` | A line under your contact details (example below). Empty means hidden. When set, the cover letter also says you need no sponsorship. |

Example work-authorization line: `U.S. Work Authorized | No Visa Sponsorship Needed`

**documents**: what is created, and in which format

| Setting | Default | Meaning |
|---|---|---|
| `cover_letter` | `true` | `false` = no cover letter. This also saves one of the three model calls. |
| `requirements_document` | `true` | `false` = no JobRequirements document |
| `resume_format` | `docx` | `docx`, `pdf` or `both` |
| `cover_letter_format` | `docx` | `docx`, `pdf` or `both` |
| `requirements_format` | `docx` | `docx`, `pdf` or `both` |

PDFs are made with Microsoft Word (Windows, through pywin32, which the setup installs) or LibreOffice (free, any system). If neither is available, the Word file is kept and you are told. For online applications, **.docx is the safest format for ATS**; `both` gives you a PDF to email as well.

**resume**: length

| Setting | Default | Meaning |
|---|---|---|
| `max_pages` | `2` | Page limit (see section 7) |
| `shorten_role_first` | `""` | Company whose bullets are cut first when too long (empty = your oldest role) |
| `drop_bachelors_if_higher_degree` | `true` | Allow removing the Bachelor's degree when still too long |

**conditions**: which jobs are worth a resume

| Setting | Default | Meaning |
|---|---|---|
| `min_fit_score` | `70` | Skip jobs whose fit with your real experience is below this % (0 = never skip). See section 4. |
| `skip_if_posting_mentions` | `[]` | Skip a posting that mentions any of these words, e.g. `["security clearance", "TS/SCI"]`. Checked **before** any model call, so it costs nothing. |
| `skip_job_titles_with` | `[]` | Skip job titles containing any of these whole words, e.g. `["VP", "Director", "Intern"]` |
| `allowed_work_arrangements` | `[]` | Only these, e.g. `["Remote", "Hybrid"]`. Empty = any. Postings that don't say are kept. |
| `min_salary` | `0` | Skip jobs whose posted maximum is below this yearly amount (hourly pay is converted at 2,080 hours). Postings without pay are kept. 0 = off. |

Jobs skipped by a condition are listed in `applications\skipped_not_a_fit.csv` with the reason, like jobs that aren't a fit. `--ignore-fit` overrides the fit check and all conditions for one run.

**study_material**: study guide and interview Q&A, only when you want them. See section 10.

**Older settings files** (version 2.3 and earlier) are upgraded automatically the next time you run `jobpilot settings`: your values are kept, the new options are added, and the old file is saved as `settings.json.bak`.

The environment variables `JOBPILOT_WORK_AUTH` and `JOBPILOT_MODEL`, if set, take priority over this file.

## 10. Study material and interview prep

JobPilot can also help you prepare, **only when you ask for it**. Nothing extra is made (or paid for) otherwise.

| You want | How |
|---|---|
| A study guide for one run | `jobpilot tailor "<link>" --study` |
| The interview Q&A for one run | `jobpilot tailor "<link>" --questions` (add `--study` for both) |
| Study material for every job | Set `"enabled": true` (and `"interview_questions": true` for the Q&A) in the `study_material` section of the configuration file |
| Study material for a job you already have | `jobpilot study "applications\2026-11-12_Acme_QA_Engineer"` (add `--questions`, or `--questions-only`). Part of the folder name works too: `jobpilot study Acme` |
| No study material this run, although the configuration turns it on | `--no-study` |

**StudyGuide_...pdf** lists 6-12 topics for *this* job, most important first: your Gap and Partial requirements, the job's core tools, then what interviewers will probe. Each topic says why it matters, what exactly to learn, a small hands-on exercise, and **links**:

- the official documentation, when the tool has it (each link is checked when the guide is made, so it never points to a page that doesn't exist);
- searches for the topic on **YouTube, Pluralsight, LinkedIn Learning, Coursera and Udemy** (pick the sites in `link_sources`). Search links always work and show the newest courses; choose a recent, well-rated one. Pluralsight, LinkedIn Learning, Coursera and Udemy may need a subscription or a free trial.

It starts with a short day-by-day study plan.

**InterviewQA_...pdf** has the questions you are most likely to be asked for this job, with model answers written from **your own resume**: technical, scenario, behavioral (STAR format) and questions about your gaps (answered honestly: what related experience you have and how you'd close the gap). Each has a tip on what the interviewer is really checking. Practice them out loud and make them your own.

Each document is one extra model call. The settings are in the `study_material` section:

| Setting | Default | Meaning |
|---|---|---|
| `enabled` | `false` | Make a study guide for every job |
| `learning_links` | `true` | Add course and video links to each topic |
| `link_sources` | all five | Any of `YouTube`, `Pluralsight`, `LinkedIn Learning`, `Coursera`, `Udemy`, `freeCodeCamp` |
| `interview_questions` | `false` | Also make the Interview Q&A for every job (with `enabled`) and for `jobpilot study` |
| `question_count` | `25` | How many questions (5-60) |
| `format` | `pdf` | `pdf`, `docx` or `both` (PDFs need Microsoft Word or LibreOffice; otherwise you get the Word file) |

## 11. Salary guidance

When a posting lists a maximum salary, JobPilot suggests what to ask for:

| The posting lists | Suggested ask |
|---|---|
| A narrow range (max no more than 1.35 times the min) | About 85% of the way up the range |
| A wide range | About 70% of the way up (the very top of a wide band is usually reserved) |
| Only a maximum ("up to $X") | 97% of the maximum |
| A single figure | That figure |
| Hourly pay | The same rules, per hour |

It also gives a **range to quote**, starting at the suggested ask. If a posting lists several ranges (by location, or a target and a full range), the first one is used and you are told to check which applies to you. These numbers come only from the posting. Compare them with market data for your city and role, and weigh bonus, equity and benefits, before you negotiate.

## 12. The application tracker

`Job_Application_Tracker.xlsx` is created next to your `applications` folder and updated after every run.

- **Dark blue columns** are filled in by JobPilot: dates, company, job, location, pay, *Posted Max*, *Suggested Ask*, *Ask Range to Quote*, sponsorship, ATS result, keyword coverage, qualifications met, rules applied, and links to the posting, folder and resume.
- **Yellow columns are yours:** *Status* (a drop-down), *Date Applied*, *Salary Requested*, *Follow-up Date*, *Contact*, *Next Step*, *Notes*. JobPilot never overwrites them.
- **Follow-up Date** fills in as Date Applied + 7 days and turns **red** when it is due and the status is still *Applied* or *No Response*.
- The **Summary** tab counts applications by status and shows your interview rate and average keyword coverage.
- **Close the tracker before running JobPilot.** If it was open, the documents are still created; run `jobpilot track` afterwards to catch up.

## 13. Editing and rebuilding

Want to change a bullet, the summary or a cover-letter paragraph? Edit `content.json` in the job's folder (any text editor; keep the quotes and commas), then run:
```
jobpilot render "applications\2026-10-29_Acme_Corp_Senior_SDET\content.json"
```
This rebuilds the documents (including any study guide and interview Q&A) and refreshes the tracker, with no API call. You can also simply edit the Word files directly.

## 14. Honesty and privacy

- JobPilot is instructed never to invent employers, titles, dates, degrees, certifications, tools or metrics. **Still, read everything before sending.**
- Your resume and the job posting are sent to the Anthropic API to generate the documents. Nothing else is sent anywhere.
- Your personal files stay on your computer, in `.resume_agent` and your `applications` folder. Don't put your API key in shared files.

## 15. FAQ

**Can I use it for any profession?** Yes. Nothing is tied to one field; it works from your resume and the posting.

**Can I change the look of the resume?** The layout is deliberately plain for ATS. You can restyle the Word file after it is created.

**It says "Gap" for something I've done.** Add it to your base resume (or an experience rule) and run `tailor` again.

**How much does it cost?** Each application makes three model calls (two without a cover letter), plus one each for a study guide and an interview Q&A if you ask for them. The exact cost depends on the model and the length of your resume and the posting; check your usage in the Anthropic Console.
