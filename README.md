# GitHub Activity Report

A production-grade Python tool that collects commits, pull requests, and GitHub
Actions workflow runs across one repository or an entire organisation, joins
the data on commit SHA, and writes a colour-coded multi-sheet Excel report that
is committed back to the repository on every run.

---

## Table of Contents

1. [What You Need to Provide](#what-you-need-to-provide)
2. [What the Report Contains](#what-the-report-contains)
3. [Prerequisites](#prerequisites)
4. [Quick Start — Local](#quick-start--local)
5. [Configuration Reference](#configuration-reference)
6. [Token Permissions](#token-permissions)
7. [GitHub Actions Setup](#github-actions-setup)
8. [Organisation-Wide Scanning](#organisation-wide-scanning)
9. [Excel Report Layout](#excel-report-layout)
10. [Row Colour Guide](#row-colour-guide)
11. [What to Change in the Code](#what-to-change-in-the-code)
12. [Troubleshooting](#troubleshooting)
13. [Project Structure](#project-structure)

---

## What You Need to Provide

This is the complete list of things **you must supply** before the tool works
in your organisation. Everything else has a safe default.

### 1. GitHub Personal Access Token — `GH_PAT`

A Classic PAT with the following scopes:

| Scope | Why it is needed |
|---|---|
| `repo` | Read commits, PRs, branches (including private repos) |
| `workflow` | Read workflow run details and logs |
| `read:org` | List all repositories in an organisation |

**How to create it:**
`GitHub → Settings → Developer settings → Personal access tokens → Tokens (classic) → Generate new token`

Add it as a **repository secret** named `GH_PAT`:
`Your repo → Settings → Secrets and variables → Actions → New repository secret`

---

### 2. Target — at least one of these

| Variable | What it does | Example value |
|---|---|---|
| `GH_REPO` | Scan a **single** repository | `my-org/backend-api` |
| `GH_ORG` | Scan **all repos** in an organisation | `my-org` |

> If both are set, `GH_REPO` takes priority.
> If neither is set, the tool scans all repos the token can reach.

**Where to set for GitHub Actions:** edit
[.github/workflows/github_activity_report.yml](.github/workflows/github_activity_report.yml)
and add your values to the `env:` block of the *Run GitHub Activity Report* step.

---

### 3. Report title and contact details (optional but recommended)

Search for `✏ CHANGE` in
[github_activity_report.py](github_activity_report.py) — there are five
clearly marked places:

| Mark in code | What to set |
|---|---|
| `GH_ORG` default | Your GitHub organisation login |
| `REPORT_TITLE` default | Your organisation or project name |
| `"Maintained by"` in cover sheet | Your team name |
| `"Contact"` in cover sheet | Your team email / Slack channel |
| `_BRAND_DARK` hex colour | Your brand primary colour (optional) |

---

## What the Report Contains

| Sheet | Contents |
|---|---|
| **Cover** | Title, generation timestamp, date range, colour key, contact details |
| **Activity** | Every commit × PR × workflow run joined on commit SHA — latest first |
| **Access Control** | Collaborators and teams with permission levels |
| **Failure Summary** | All failed / timed-out runs with root-cause diagnostics |
| **Failure Alerts** | Red banner + summary table for every failure found |

---

## Prerequisites

- Python **3.10** or later
- A GitHub Personal Access Token (see above)

```bash
python3 --version   # must be 3.10+
```

---

## Quick Start — Local

```bash
# 1. Clone the repository
git clone https://github.com/<your-org>/git-report.git
cd git-report

# 2. Create and activate a virtual environment
python3 -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate

# 3. Install dependencies
pip install -r requirements.txt

# 4. Set required environment variables
export GH_PAT=ghp_xxxxxxxxxxxxxxxxxxxx   # your Personal Access Token
export GH_REPO=my-org/my-repo            # OR use GH_ORG for org-wide scan

# 5. Run
python github_activity_report.py

# Output: github_activity_report.xlsx
```

---

## Configuration Reference

All configuration is passed through environment variables — no files to edit
for basic use.

### Required

| Variable | Description | Example |
|---|---|---|
| `GH_PAT` | Personal Access Token | `ghp_abc123...` |
| `GH_REPO` | Single repo to scan (`owner/repo`) | `my-org/backend-api` |
| `GH_ORG` | Organisation login (scans all repos) | `my-org` |

### Optional

| Variable | Default | Description |
|---|---|---|
| `GH_OUTPUT` | `github_activity_report.xlsx` | Output file path |
| `GH_MAX_RUNS` | `200` | Max workflow runs fetched per repo |
| `GH_LOOKBACK_DAYS` | `30` | Report activity from the last N days |
| `GH_SINCE` | *(none)* | ISO-8601 start date — overrides `GH_LOOKBACK_DAYS` |
| `GH_UNTIL` | *(none)* | ISO-8601 end date |
| `GH_LOG_LEVEL` | `INFO` | Logging verbosity: `DEBUG`, `INFO`, `WARNING` |
| `GH_REPORT_TITLE` | `GitHub Activity & Workflow Audit Report` | Banner title in the Excel file |

### Date filtering examples

```bash
# Last 7 days
export GH_LOOKBACK_DAYS=7

# Specific quarter
export GH_SINCE=2024-01-01
export GH_UNTIL=2024-03-31

# Debug mode
export GH_LOG_LEVEL=DEBUG
```

---

## Token Permissions

### Classic PAT — recommended for org-wide use

| Scope | Reason |
|---|---|
| `repo` | Read commits, PRs, branches (private repos) |
| `workflow` | Read workflow runs and logs |
| `read:org` | List organisation repositories |

Create at:
`GitHub → Settings → Developer settings → Personal access tokens → Tokens (classic)`

### Fine-grained PAT — for stricter environments

| Permission | Level |
|---|---|
| Contents | Read |
| Actions | Read |
| Pull requests | Read |
| Members *(org-level)* | Read |

### Built-in `GITHUB_TOKEN` — single repo only

Works automatically inside GitHub Actions for the current repository only.
Cannot read workflow logs or access other repositories.
Used as a fallback when `GH_PAT` is not set.

---

## GitHub Actions Setup

The workflow runs automatically on every push to `main`, on pull requests,
on manual trigger, and after any other workflow completes.  It generates the
report, uploads it as an artifact, and commits the file back to the repository.

### Step 1 — Add the PAT secret

```
Your repo → Settings → Secrets and variables → Actions → New repository secret

Name:  GH_PAT
Value: <your Personal Access Token>
```

### Step 2 — Set your organisation or repository

Edit [.github/workflows/github_activity_report.yml](.github/workflows/github_activity_report.yml),
find the `Run GitHub Activity Report` step, and add your values:

```yaml
- name: Run GitHub Activity Report
  env:
    GH_PAT: ${{ secrets.GH_PAT || github.token }}
    GH_REPO: ${{ github.repository }}   # ← scans this repo only
    # GH_ORG: my-org-name               # ← uncomment to scan all org repos
    GH_LOOKBACK_DAYS: "30"
    GH_REPORT_TITLE: "My Org — GitHub Activity Report"
```

### Step 3 — Push and verify

Push any change to `main`. The workflow will:
1. Run the Python script
2. Upload `github_activity_report.xlsx` as a downloadable **Actions artifact**
3. Commit it back to the repository (`chore: update activity report [...]`)

### Downloading the artifact

`GitHub → Actions → GitHub Activity Report → latest run → Artifacts → github_activity_report`

---

## Organisation-Wide Scanning

To scan every repository in your GitHub organisation:

```bash
export GH_PAT=ghp_xxxxxxxxxxxxxxxxxxxx
export GH_ORG=my-org-name
# Do NOT set GH_REPO
python github_activity_report.py
```

In GitHub Actions, remove `GH_REPO` from the env block and add `GH_ORG`:

```yaml
env:
  GH_PAT: ${{ secrets.GH_PAT }}
  GH_ORG: my-org-name
  GH_LOOKBACK_DAYS: "30"
```

> **Rate limits:** org-wide scans make thousands of API calls. Use a PAT
> (5 000 req/hr) rather than the built-in token (1 000 req/hr). The script
> waits automatically when rate-limited and resumes.

---

## Excel Report Layout

```
┌──────────────────────────────────────────────────────────┐
│  COVER          — title, config, date range, colour key  │  ← opens first
├──────────────────────────────────────────────────────────┤
│  ACTIVITY       — commits × PRs × runs (latest at top)   │
├──────────────────────────────────────────────────────────┤
│  ACCESS CONTROL — collaborators and teams                 │
├──────────────────────────────────────────────────────────┤
│  FAILURE SUMMARY — failed runs with diagnostics           │
├──────────────────────────────────────────────────────────┤
│  FAILURE ALERTS  — banner + alert table                   │  ← opens first if failures
└──────────────────────────────────────────────────────────┘
```

All data sheets have:
- **Latest rows at the top** — sorted by commit date / run start time
- **Frozen header row** — stays visible when scrolling
- **Auto-filter dropdowns** on every column
- **Auto-sized columns** (capped at 62 characters)

### Activity sheet columns

| Column | Description |
|---|---|
| Repository | Repository name |
| Organization | Owner / organisation |
| Visibility | `public` or `private` |
| Default Branch | e.g. `main` |
| Commit ID | Full 40-character SHA |
| Commit Message | First line of the commit message |
| Author | GitHub login (falls back to git author name) |
| Date | ISO 8601 commit timestamp |
| Branch | Branch the commit was first seen on |
| PR ID | Pull request number |
| PR Title | Pull request title |
| PR Author | GitHub login of the PR author |
| PR Status | `open`, `closed`, or `merged` |
| PR Merged | `Yes` / `No` |
| PR Merge Date | ISO 8601 merge timestamp |
| Workflow Name | GitHub Actions workflow name |
| Workflow Run ID | Numeric run ID |
| Trigger Event | `push`, `pull_request`, `schedule`, `workflow_dispatch`, etc. |
| Run Started At | ISO 8601 run start timestamp |
| Workflow Status | `completed`, `in_progress`, `queued` |
| Workflow Conclusion | `success`, `failure`, `cancelled`, `timed_out` |
| Failure Reason | Human-readable failure description |
| Failed Job | Name of the job that failed |
| Failed Step | Name of the step that failed |
| Error Line | Line number from the log (if found) |
| Suggested Fix | Automated fix recommendation |

---

## Row Colour Guide

| Colour | Applies when | Meaning |
|---|---|---|
| 🟢 **Green** (all columns) | `Workflow Conclusion = success` | Run completed successfully |
| 🔴 **Red** (all columns) | `Workflow Conclusion = failure / timed_out` | Run failed or timed out |
| 🟡 **Yellow** (all columns) | `Workflow Conclusion = cancelled / skipped` | Run was cancelled or skipped |
| ⬜ White / grey stripe | No workflow conclusion | Commit or PR with no associated run |

Every cell in a row gets the same colour — not just the status column — so
failures are immediately visible when scanning the sheet.

---

## What to Change in the Code

Open [github_activity_report.py](github_activity_report.py) and search for
`✏ CHANGE` — every customisation point is marked.

### Section 1 — Organisation / project configuration (top of file)

```python
GH_ORG   = os.getenv("GH_ORG", "")          # ✏ CHANGE: your org login
MAX_RUNS = int(os.getenv("GH_MAX_RUNS", "200"))  # ✏ CHANGE: increase for busier orgs
LOOKBACK_DAYS = int(os.getenv("GH_LOOKBACK_DAYS", "30"))  # ✏ CHANGE: history window
REPORT_TITLE = os.getenv("GH_REPORT_TITLE", "GitHub Activity & Workflow Audit Report")
```

### Section 2 — Excel colour theme

```python
_BRAND_DARK  = "1F3864"   # ✏ CHANGE: header background — use your brand colour
_BRAND_LIGHT = "EEF2F7"   # ✏ CHANGE: alternate row stripe
_SUCCESS_BG  = "C6EFCE"   # ✏ CHANGE: success row background
_FAILURE_BG  = "FFC7CE"   # ✏ CHANGE: failure row background
```

### Section 3 — Failure fix-hint rules

Add entries to `_FIX_RULES` that match your stack:

```python
_FIX_RULES = [
    # ✏ CHANGE: add keywords specific to your build tools
    (["your-tool", "your-error-keyword"], "Your suggested fix message."),
    ...
]
```

### Section 4 — Column definitions

Add or remove columns from any of the four column lists:

```python
ACTIVITY_COLUMNS = [
    "Repository", "Organization", ...   # ✏ CHANGE: add/remove columns here
]
```

### Cover sheet — contact details

```python
("Maintained by", "DevOps / Platform Engineering team"),   # ✏ CHANGE
("Contact",        "devops@yourcompany.com"),               # ✏ CHANGE
```

### fetch_repos() — repository type filter

```python
params={"type": "all"},   # ✏ CHANGE: use "public" to skip private repos
```

### GitHub Enterprise Server

Replace `BASE_URL` at the top of the file:

```python
BASE_URL = "https://api.github.com"
# ✏ CHANGE for GHES:
# BASE_URL = "https://github.mycompany.com/api/v3"
```

---

## Troubleshooting

### "No repositories found"
- Check `GH_ORG` is the correct organisation login (case-sensitive)
- Verify the PAT has `read:org` scope
- If your org uses SAML SSO, the token must be authorised for SSO

### "Repository not found or token lacks access"
- Check `GH_REPO` is in `owner/repo` format — e.g. `my-org/backend-api`
- Verify the PAT has `repo` scope for private repos

### "Rate-limited — sleeping N s"
- Normal behaviour — the script waits automatically and resumes
- Use a PAT instead of the built-in token for higher limits (5 000 vs 1 000 req/hr)

### Report committed back but shows no change
- The workflow skips the commit when the Excel file is identical to the previous
  run — this means no new activity occurred in the lookback window

### Access Control sheet is empty
- Requires `repo` scope on a Classic PAT
- Teams endpoint returns 404 for personal repositories — expected behaviour

### Excel file opens on "Failure Alerts" tab
- Intentional — the file opens on the most important sheet when failures exist

### Enable debug logging
```bash
export GH_LOG_LEVEL=DEBUG
python github_activity_report.py
```

---

## Project Structure

```
.
├── github_activity_report.py           # Main script
├── requirements.txt                    # Python dependencies (requests, pandas, openpyxl)
├── README.md                           # This file
└── .github/
    └── workflows/
        └── github_activity_report.yml  # GitHub Actions workflow
```

### Customisation map inside the script

```
github_activity_report.py
│
├── SECTION 1  ✏  Organisation / project configuration
│               GH_ORG, REPORT_TITLE, MAX_RUNS, LOOKBACK_DAYS
│
├── SECTION 2  ✏  Excel colour theme
│               Hex codes for headers, success, failure, warning rows
│
├── SECTION 3  ✏  Failure fix-hint rules
│               Keyword → suggestion mappings for your tech stack
│
├── SECTION 4  ✏  Column definitions
│               Add / remove / rename columns in any sheet
│
└── _write_cover_sheet()  ✏  Team name and contact email on the Cover sheet
```
