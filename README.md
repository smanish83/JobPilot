# JobPilot

*Runs the whole application workflow for you, from the posting to the tracker.*

**Tailor your resume to any job posting, get a matching cover letter, and keep track of every application. Everything comes out as Word documents.**

Give JobPilot a job link (LinkedIn, Indeed or a company careers page) or a saved job description. About a minute later you get a folder with:

| File | What it is |
|---|---|
| `Resume_<Company>_<Job>.docx` | Your resume rewritten for this job, using only what is in your own resume, and checked against ATS rules |
| `CoverLetter_<Company>_<Job>.docx` | A cover letter for this job, written from that tailored resume |
| `JobRequirements_<Company>_<Job>.docx` | The job broken down: requirements you meet / partly meet / miss, keywords, salary, a suggested salary ask, likely interview topics and the full original posting |
| `StudyGuide_...pdf`, `InterviewQA_...pdf` | Only if you ask: what to learn for this job, with links, and likely interview questions with model answers |

Every application is also added to **`Job_Application_Tracker.xlsx`**, where you track status, follow-ups, contacts and the salary you asked for.

## What it does for you

- **Tailors honestly.** It reorders, rephrases and emphasizes your real experience to match the posting. It is told never to invent employers, titles, dates, tools or numbers. Anything the job wants that you do not have is listed as a gap to prepare for, not added to your resume.
- **Speaks the posting's language.** Where your experience is described in different words, it uses the posting's exact phrases for the same experience.
- **Gets past ATS screening.** It uses the posting's exact keywords, standard headings, plain single-column formatting and consistent dates, then runs an ATS checklist and shows the result.
- **Keeps the resume to two pages.** If it runs long, it shortens your oldest role first (you can change this).
- **Experience rules.** Extra experience that should appear only for certain jobs, for example "if the job mentions Playwright, add my Playwright bullets".
- **Salary guidance.** When a posting lists a maximum, it suggests what to ask for and records it in the tracker.
- **No duplicates.** Jobs you've already prepared or applied for are recognized by their link and skipped.
- **Skips jobs that aren't a fit.** Before writing anything, it scores how well the job matches your real experience. A VP role, "automation" that means ATM or factory equipment, or a must-have stack you've never used scores below 70%, so it is skipped with the reason, and you save the time and API cost.
- **Study material, when you ask.** For any job: a study guide with the topics to learn (gaps first), hands-on exercises and links to the official docs and to courses and videos on YouTube, Pluralsight, LinkedIn Learning, Coursera and Udemy, plus an optional PDF of likely interview questions with answers based on your resume.
- **Configurable.** One configuration file decides whether you get a cover letter, Word or PDF (or both) for each document, the page limit, which jobs to skip (titles such as VP, on-site jobs, pay below your minimum, words like "security clearance"), whether you want study material and, optionally, your API key.
- **Several jobs at once.** Pass several links, or a file of links, and run them one after another or in parallel.

## Quick start

1. **Install Python 3.10 or newer** from [python.org](https://www.python.org/downloads/). On Windows, tick *"Add python.exe to PATH"* during installation.
2. **Connect to Claude**, either way:
   - **Already use Claude Code?** If the `claude` command is installed and signed in (`claude auth login`), JobPilot uses that sign-in. No API key needed.
   - **Otherwise, get an Anthropic API key** at [console.anthropic.com](https://console.anthropic.com) (Settings > API Keys). API usage is billed separately from a Claude.ai subscription.
3. **Get JobPilot.** On [the GitHub page](https://github.com/smanish83/JobPilot), click **Code > Download ZIP** and unzip it somewhere easy to find, for example `Documents\JobPilot`. Or, with Git: `git clone https://github.com/smanish83/JobPilot.git`.
4. **Run the setup** in that folder:
   - Windows: double-click **`setup_windows.bat`**
   - Mac / Linux: open a terminal in this folder and run **`./setup_mac_linux.sh`**

   The guided setup finds your Claude Code sign-in or asks for an API key, then your current resume and an optional work-authorization line, then checks that everything is ready. No resume handy? Type `sample` to try JobPilot with the included sample resume first.
5. **Tailor your first application.** Open a terminal in this folder and run:
   ```
   jobpilot tailor "https://www.linkedin.com/jobs/view/1234567890/"
   ```
   On Mac/Linux and in Windows PowerShell, type `./jobpilot` (or `.\jobpilot`) instead of `jobpilot`.

Full instructions are in **[SETUP.md](SETUP.md)**. The complete user guide is **[HOW_TO_USE.md](HOW_TO_USE.md)**. Both are also in **JobPilot_User_Guide.pdf**, ready to print or share.

## Folder contents

```
JobPilot/
├── README.md                 this overview
├── SETUP.md                  installation and troubleshooting
├── HOW_TO_USE.md             the user guide
├── JobPilot_User_Guide.pdf   setup + user guide as one PDF
├── setup_windows.bat         one-click setup for Windows
├── setup_mac_linux.sh        setup for macOS / Linux
├── jobpilot.bat / jobpilot   launchers (Windows / Mac-Linux)
├── jobpilot.py               the tool itself
├── requirements.txt          Python libraries it needs
├── requirements-dev.txt      extra library for running the tests (pytest)
├── tests/                    automated tests (no API key or internet needed)
├── pyproject.toml            for installing with pip instead (optional)
├── examples/                 a sample resume, a sample job, a jobs list, sample experience rules and a sample configuration
├── CHANGELOG.md              what changed in each version
└── LICENSE                   MIT License
```

Your personal data never goes in this folder's code. Your base resume, configuration file and experience rules live in a `.resume_agent` folder in your home directory, and your applications are created in the folder you run JobPilot from. That means you can share or update this folder without sharing your resume. The included `.gitignore` also keeps `applications/`, trackers and Word files out of Git, so a fork or clone never publishes them by accident.

## Important

- **Always read every document before you send it.** The tool is careful, but you are responsible for what goes to an employer.
- Job sites sometimes block automatic reading (Indeed often does). When that happens, JobPilot asks you to paste the job description instead.
- Each application makes three model calls (one when a job is skipped as a duplicate or not a fit). With an API key, each call has a small API cost; through a Claude Code sign-in, the calls count toward that account's usage limits instead.
- JobPilot is not affiliated with LinkedIn, Indeed, Anthropic or any job site. Respect each site's terms of use.

## Contributing

Bug reports and ideas are welcome: open an issue on GitHub. For a code change, open a pull request with a short description, and never include real resumes, job applications or API keys in issues, commits or test files.

### Running the tests

The automated tests run JobPilot end to end with a fake model, in a temporary folder, using the sample resume. They need no API key, no internet and never touch your own resume, settings or applications.

```
pip install -r requirements.txt -r requirements-dev.txt
python -m pytest
```

They cover job-link matching, keyword and ATS checks, salary suggestions, experience rules, the fit check, duplicate detection, the tracker, parallel runs and the command line. GitHub runs them automatically on Windows and Linux for every push and pull request.

## License

JobPilot is released under the [MIT License](LICENSE). Copyright (c) 2026 Manish Sandbhor.
