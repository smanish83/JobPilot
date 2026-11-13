# JobPilot: Setup Guide

Setup takes about 10 minutes. You only do it once.

## What you need

| Item | Details |
|---|---|
| A computer | Windows 10/11, macOS, or Linux |
| For PDF output (optional) | Microsoft Word on Windows, or [LibreOffice](https://www.libreoffice.org) (free) on any system |
| Python 3.10 or newer | Free from [python.org/downloads](https://www.python.org/downloads/) |
| A way to reach Claude | **Either** Claude Code installed and signed in (no API key needed), **or** an Anthropic API key from [console.anthropic.com](https://console.anthropic.com) (pay-as-you-go, billed separately from a Claude.ai subscription) |
| Your current resume | A `.docx` (best), `.pdf`, `.md` or `.txt` file |
| Internet access | To install the libraries and to call the API |
| Microsoft Word (optional) | To open the documents. On Windows it is also used to count resume pages exactly. |

---

## Step 1: Install Python

**Windows**

1. Download the latest Python 3 installer from [python.org/downloads](https://www.python.org/downloads/).
2. On the first installer screen, **tick "Add python.exe to PATH"**, then click *Install Now*.
3. To check it worked, open *Command Prompt* and type `python --version`. You should see `Python 3.10` or higher.

**Mac**

- Install from [python.org/downloads](https://www.python.org/downloads/), or with Homebrew: `brew install python`.
- Check it worked: `python3 --version`.

**Linux**

- Use your package manager, for example `sudo apt install python3 python3-venv`.

## Step 2: Connect to Claude

JobPilot looks for a connection in this order and uses the first one it finds:

1. **A local API key** saved in your configuration file (`anthropic_api_key`).
2. **Claude Code** (the `claude` command), installed and signed in. JobPilot sends its requests through it, so your Claude Code sign-in (for example your Claude subscription) is used and **no API key is needed**. To set it up, [install Claude Code](https://code.claude.com/docs/en/setup), then run `claude auth login` once. `claude auth status` shows whether you are signed in.
3. **The `ANTHROPIC_API_KEY` environment variable.**
4. If none of these is found, JobPilot asks you for an API key (and can save it in the configuration file).

To always use one method, set `"auth"` in the configuration file to `"claude_cli"` or `"api_key"` (the default is `"auto"`, the order above).

**To get an API key** (only if you don't use Claude Code):

1. Sign in at [console.anthropic.com](https://console.anthropic.com) and add a payment method or credits.
2. Go to **Settings > API Keys > Create Key**.
3. Copy the key. It starts with `sk-ant-`. Keep it private: anyone with it can use your account.

## Step 3: Run the setup

Get JobPilot from [github.com/smanish83/JobPilot](https://github.com/smanish83/JobPilot):

- **Without Git:** click **Code > Download ZIP**, then unzip it somewhere easy to find, for example `Documents\JobPilot`.
- **With Git:** `git clone https://github.com/smanish83/JobPilot.git`

This folder is also where your applications and tracker will be created.

**Windows:** double-click **`setup_windows.bat`**.

**Mac / Linux:** open Terminal in the folder and run:
```
chmod +x setup_mac_linux.sh
./setup_mac_linux.sh
```

The setup creates a private Python environment inside the folder (`.venv`), installs the libraries, then starts the **guided setup**, which asks you:

1. **Your connection to Claude.** If a local key or a signed-in Claude Code is found, it is used and nothing is asked. Otherwise paste an API key; it is hidden as you type. On Windows it is saved with `setx`. On Mac/Linux you are asked before it is added to `~/.zshrc` or `~/.bashrc`.
2. **Your resume file.** Type the path, or drag the file into the window. To try JobPilot first, type `sample` to use the included sample resume (`examples/sample_resume.docx`, a fictional person), and replace it with yours later.
3. **A work-authorization line** (optional). It appears under your contact details, for example `U.S. Work Authorized | No Visa Sponsorship Needed`. Press Enter to skip.
4. **Experience rules** (optional). See HOW_TO_USE.md. You can create them later.

It finishes with a readiness check. All lines should say `[OK]`.

## Step 4: Check it any time

Open a terminal **in the JobPilot folder** and run:

| Windows Command Prompt | Windows PowerShell | Mac / Linux |
|---|---|---|
| `jobpilot check` | `.\jobpilot check` | `./jobpilot check` |

The rest of the guides write `jobpilot`; use the form that matches your terminal.

> **Tip (Windows):** to open a terminal in a folder, open the folder in File Explorer, click the address bar, type `cmd` and press Enter.

---

## Optional extras

- **Exact page counts on Windows.** The setup installs `pywin32`, which lets JobPilot ask Word how many pages the resume has. Without Word, JobPilot estimates the page count instead.
- **Install as a normal Python package.** Instead of the setup scripts, you can run `pip install .` (or `pip install .[word]` on Windows) in the folder. You then get a `jobpilot` command, set the API key yourself, and run `jobpilot init`.

## Updating to a new version

If you cloned with Git, run `git pull` in the folder. Otherwise download the new ZIP and copy its files over the old ones, keeping the `.venv` folder (and your `applications` folder and tracker). See CHANGELOG.md for what changed. If `requirements.txt` changed, run the setup again; it keeps your existing resume and settings. Your personal files live in your home folder (`.resume_agent`), so updates never touch them.

## Uninstalling

Delete the JobPilot folder. To remove your personal data too, delete the `.resume_agent` folder in your home directory (`C:\Users\<you>\.resume_agent` on Windows, `~/.resume_agent` on Mac/Linux). On Windows, the API key can be removed under *Edit environment variables for your account*.

---

## Troubleshooting

| Problem | Fix |
|---|---|
| `Python was not found` or the Microsoft Store opens | Install Python from python.org and tick **Add python.exe to PATH**, then run the setup again. |
| `JobPilot could not find a way to reach Claude` | Sign in to Claude Code (`claude auth login`), or set an API key. A key saved with `setx` only reaches new terminal windows, so close this one and open a new one. Then run `jobpilot check`. |
| `Claude Code CLI request failed` | Run `claude auth status`. If you are signed out, run `claude auth login`. If your plan's usage limit is reached, wait, or set `"auth": "api_key"` to use an API key instead. If `claude` is installed somewhere unusual, set `"claude_cli_path"`. |
| `401` / `authentication_error` | The key is wrong or was deleted. Create a new key in the Console and run `jobpilot init`. |
| `credit balance is too low` | Add credits in the Anthropic Console (Plans & Billing). |
| `'jobpilot' is not recognized` | You are not in the JobPilot folder, or you are in PowerShell. Use `.\jobpilot`, or `cd` into the folder first. |
| A job site is blocked | Normal for Indeed and some LinkedIn pages. Paste the description when asked, or save it to a `.txt` file and run `jobpilot tailor job.txt`. |
| `Tracker not updated ... is open in Excel` | Close the tracker in Excel, then run `jobpilot track`. |
| `Building the Word files failed` | Your content was saved first. Fix the issue, then run the `jobpilot render ...` command it shows. No new API cost. |
| Rate-limit errors with `--parallel` | Use `--parallel 2` or run without it, then re-run the failed links. |
| Anything else | Run `jobpilot check` and fix every line marked `!!`. |
