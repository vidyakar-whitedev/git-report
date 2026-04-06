# GitHub Activity Report

A Python script that collects commits, pull requests, and workflow runs from any GitHub repository using the GitHub REST API and exports a formatted, colour-coded Excel report.

Runs locally in one command. Also ships a GitHub Actions workflow that generates and uploads the report automatically on every push or pull request.

---

## Table of Contents

- [What it produces](#what-it-produces)
- [Project structure](#project-structure)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Configuration — environment variables](#configuration--environment-variables)
- [Create a GitHub Personal Access Token](#create-a-github-personal-access-token)
- [Run locally](#run-locally)
- [Expected output](#expected-output)
- [Excel report format](#excel-report-format)
- [How the data is assembled](#how-the-data-is-assembled)
- [GitHub Actions — automated report](#github-actions--automated-report)
- [Workflow triggers](#workflow-triggers)
- [Download the artifact](#download-the-artifact)
- [Troubleshooting](#troubleshooting)

---

## What it produces

A single Excel file — `github_activity_report.xlsx` — with one sheet named **Activity**.

Every row represents a unit of repository activity. Commits, pull requests, and workflow runs are joined together on their shared commit SHA so the full picture of each change is visible in one place.

| Column | Description |
|---|---|
| Repository | Repository name |
| Organization | Owner / organization name |
| Commit ID | Full 40-character SHA |
| Commit Message | First line of the commit message |
| Commit Author | GitHub login (falls back to git author name) |
| Commit Date | ISO 8601 timestamp |
| Branch | Branch the commit was first seen on |
| PR ID | Pull request number |
| PR Title | Pull request title |
| PR Author | GitHub login of the PR author |
| PR Status | `open`, `closed`, or `merged` |
| PR Merge Date | ISO 8601 merge timestamp (empty if not merged) |
| Workflow Name | Name of the GitHub Actions workflow |
| Workflow Run ID | Numeric run ID |
| Workflow Status | `completed`, `in_progress`, `queued`, etc. |
| Workflow Conclusion/Error Reason | `success`, `failure`, `cancelled`, or detail of which job/step failed |
| Trigger | Event that triggered the run: `push`, `pull_request`, `schedule`, etc. |

---

## Project structure

```
git-report/
├── github_activity_report.py       # Main script
├── requirements.txt                # Python dependencies
├── demo.txt                        # Test file (placeholder)
└── .github/
    └── workflows/
        └── github_activity_report.yml   # GitHub Actions workflow
```

---

## Prerequisites

| Requirement | Minimum version | Check command |
|---|---|---|
| Python | 3.10 | `python3 --version` |
| pip | Any recent | `pip --version` |
| GitHub Personal Access Token | — | See [Create a PAT](#create-a-github-personal-access-token) |

---

## Installation

### 1. Clone or download the project

```bash
git clone https://github.com/<your-org>/<your-repo>.git
cd git-report/git-report
```

### 2. Install Python dependencies

```bash
pip install -r requirements.txt
```

This installs:

| Package | Version | Purpose |
|---|---|---|
| `requests` | >= 2.31.0 | GitHub REST API calls with pagination and rate-limit handling |
| `pandas` | >= 2.0.0 | DataFrame assembly and Excel writing |
| `openpyxl` | >= 3.1.0 | Excel formatting — colours, column widths, freeze panes, auto-filter |

---

## Configuration — environment variables

All configuration is passed through environment variables. No config files need editing.

| Variable | Required | Default | Description |
|---|---|---|---|
| `GH_PAT` | **Yes** | — | GitHub Personal Access Token for API authentication |
| `GH_REPO` | **Yes** (local) | Auto-set in Actions | Repository in `owner/repo` format — e.g. `octocat/Hello-World` |
| `GH_OUTPUT` | No | `github_activity_report.xlsx` | Output file path |
| `GH_MAX_RUNS` | No | `200` | Maximum number of workflow runs to fetch |

> In GitHub Actions, `GH_REPO` is set automatically from `${{ github.repository }}` — you do not need to set it manually.

---

## Create a GitHub Personal Access Token

The script needs read access to the target repository's contents, pull requests, and Actions.

### Fine-grained token (recommended)

1. Go to **GitHub → Settings → Developer settings → Personal access tokens → Fine-grained tokens**
2. Click **Generate new token**
3. Set **Resource owner** to the org or user that owns the repository
4. Under **Repository access** select the specific repository
5. Under **Permissions → Repository permissions** enable:

   | Permission | Access level |
   |---|---|
   | Contents | Read-only |
   | Pull requests | Read-only |
   | Actions | Read-only |
   | Metadata | Read-only (auto-selected) |

6. Click **Generate token** and copy the value immediately — it is only shown once

### Classic token (alternative)

1. Go to **GitHub → Settings → Developer settings → Personal access tokens → Tokens (classic)**
2. Click **Generate new token (classic)**
3. Select scopes: `repo` (full) and `workflow`
4. Click **Generate token** and copy it

---

## Run locally

### Step 1 — Set environment variables

```bash
export GH_PAT=ghp_xxxxxxxxxxxxxxxxxxxxxxxxxxxx
export GH_REPO=owner/repo-name
```

The format for `GH_REPO` must contain a `/` separating owner and repository name:

```bash
# Correct
export GH_REPO=octocat/Hello-World
export GH_REPO=my-org/my-repo

# Wrong — will exit with an error
export GH_REPO=my-repo-name
```

### Step 2 — Run the script

```bash
python3 github_activity_report.py
```

### Step 3 — Optional overrides

```bash
# Save to a different file path
export GH_OUTPUT=/tmp/my_report.xlsx

# Fetch up to 500 workflow runs instead of the default 200
export GH_MAX_RUNS=500

python3 github_activity_report.py
```

---

## Expected output

The script prints timestamped progress logs as it runs:

```
09:14:01  INFO     === GitHub Activity Report ===
09:14:01  INFO     Repository : octocat/Hello-World
09:14:01  INFO     Output     : github_activity_report.xlsx
09:14:01  INFO     Fetching branches …
09:14:02  INFO       3 branch(es) found
09:14:02  INFO     Fetching commits for 3 branch(es) …
09:14:02  INFO       branch: main
09:14:03  INFO       branch: develop
09:14:04  INFO       branch: feature/new-ui
09:14:05  INFO       47 unique commit(s) found
09:14:05  INFO     Fetching pull requests …
09:14:06  INFO       12 pull request(s) found
09:14:06  INFO     Fetching workflow runs (max 200) …
09:14:07  INFO       38 workflow run(s) found
09:14:07  INFO       fetching job detail for failed run 9871234 …
09:14:08  INFO       fetching job detail for failed run 9865432 …
09:14:08  INFO     Assembling dataset …
09:14:08  INFO     Total rows : 61
09:14:08  INFO     Report saved → github_activity_report.xlsx  (61 data rows)

Done. Report written to: github_activity_report.xlsx
```

The output file is created in the current working directory unless `GH_OUTPUT` overrides the path.

---

## Excel report format

The generated `github_activity_report.xlsx` file has the following formatting applied automatically:

| Element | Format |
|---|---|
| Sheet name | `Activity` |
| Header row | Dark navy background (`#1F3864`), white bold text, centred |
| Alternate rows | Light blue background (`#DCE6F1`) for readability |
| Failed workflow rows | Light red background (`#FFCCCC`) — rows where Workflow Status is `failure` or `timed_out` |
| Column widths | Auto-sized to content, capped at 60 characters |
| Freeze panes | Row 1 (header) is frozen so it stays visible when scrolling |
| Auto-filter | Enabled on all columns — click any header to sort or filter |

### Failure detail

For workflow runs with conclusion `failure` or `timed_out`, the script fetches the job and step level detail and populates **Workflow Conclusion/Error Reason** with the specific location:

```
Failed at: build / Run tests
Timed out at: deploy / Wait for deployment
```

For other conclusions:

| Conclusion | Error Reason value |
|---|---|
| `success` | _(empty)_ |
| `failure` | `Failed at: <job> / <step>` |
| `timed_out` | `Timed out at: <job> / <step>` |
| `startup_failure` | `Workflow startup failure` |
| `cancelled` | `Cancelled` |

---

## How the data is assembled

The script fetches four data sources and joins them by commit SHA:

```
fetch_branches()
    └── fetch_commits()   ← one call per branch, de-duplicated by SHA
                                │
                                ├── joined to PRs via merge_sha or head_sha
                                └── joined to workflow runs via head_sha

fetch_pull_requests()     ← all states: open, closed, merged
fetch_workflow_runs()     ← up to GH_MAX_RUNS, newest first
```

**Row generation rules:**

1. Each commit is an anchor row
2. If a commit matches multiple workflow runs → one row per run
3. If a commit matches a PR → PR columns are filled on the same row
4. PRs with no matching commit → their own row (commit columns empty)
5. Workflow runs with no matching commit → their own row (commit columns contain the head SHA)

**Pagination:** all API calls follow GitHub's `Link: next` header automatically. Repositories with thousands of commits or hundreds of workflow runs are handled correctly.

**Rate limiting:** if the API returns a 429 or 403 rate-limit response, the script reads the `X-RateLimit-Reset` header and sleeps until the window resets before retrying — up to 5 retries per page.

---

## GitHub Actions — automated report

The workflow file at [.github/workflows/github_activity_report.yml](.github/workflows/github_activity_report.yml) runs the script automatically and uploads the Excel file as a downloadable artifact.

### Setup — add the PAT as a repository secret

1. Go to your repository on GitHub
2. Click **Settings → Secrets and variables → Actions → New repository secret**
3. Name: `GH_PAT`
4. Value: your Personal Access Token (from [Create a PAT](#create-a-github-personal-access-token))
5. Click **Add secret**

The workflow reads `GH_REPO` from `${{ github.repository }}` automatically — no additional secrets needed.

### Workflow file overview

```yaml
name: GitHub Activity Report

on:
  push:
    branches: [main]
  pull_request:
    branches: [main]
  workflow_dispatch:        # manual trigger from the Actions UI

jobs:
  report:
    runs-on: ubuntu-latest
    permissions:
      contents: read

    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-python@v5
        with:
          python-version: '3.11'
      - uses: actions/cache@v4          # caches pip downloads
        with:
          path: ~/.cache/pip
          key: ${{ runner.os }}-pip-${{ hashFiles('**/requirements.txt') }}
      - run: pip install -r requirements.txt
      - run: python github_activity_report.py
        env:
          GH_PAT:  ${{ secrets.GH_PAT }}
          GH_REPO: ${{ github.repository }}
      - uses: actions/upload-artifact@v4
        with:
          name: github_activity_report
          path: github_activity_report.xlsx
          if-no-files-found: error
```

---

## Workflow triggers

| Trigger | When it runs |
|---|---|
| `push` to `main` | Every time a commit is pushed to the `main` branch |
| `pull_request` to `main` | Every time a PR is opened, updated, or synchronized against `main` |
| `workflow_dispatch` | Manually from **Actions → GitHub Activity Report → Run workflow** |

---

## Download the artifact

1. Go to your repository on GitHub
2. Click **Actions** in the top navigation
3. Click the **GitHub Activity Report** workflow
4. Click the most recent successful run
5. Scroll to the bottom of the run summary page
6. Under **Artifacts**, click **github_activity_report** to download a `.zip`
7. Unzip and open `github_activity_report.xlsx`

---

## Troubleshooting

### `GH_REPO must be in 'owner/repo' format`

```
ERROR: GH_REPO must be in 'owner/repo' format.
  Current value: 'my-repo-name'
  Example:  export GH_REPO=vidyakar/whitedev
```

The value must include the owner separated by `/`. Find the correct format in the repository URL:  
`https://github.com/<owner>/<repo>` → `export GH_REPO=<owner>/<repo>`

---

### `401 Unauthorized`

The PAT is invalid, expired, or was not set:

```bash
# Check the variable is actually set
echo $GH_PAT
```

If it is empty, re-export it. If it is set but returns 401, regenerate the token on GitHub.

---

### `403 Forbidden` on Actions endpoints

The PAT is missing the `Actions: Read` permission. Regenerate the token and add **Actions → Read-only** to the permissions.

---

### `404 Not Found`

```
ERROR  404 Not Found: https://api.github.com/repos/owner/repo/branches
```

Either the repository name is wrong, or the repository is private and the PAT does not have access to it. Verify the exact `owner/repo` from the repository URL on GitHub.

---

### Rate limit warning during run

```
WARNING  Rate-limited. Sleeping 42s …
```

This is normal behaviour for large repositories or unauthenticated requests. The script waits automatically and resumes. To avoid this, ensure `GH_PAT` is set — authenticated requests have a limit of 5,000 per hour vs 60 per hour unauthenticated.

---

### `ModuleNotFoundError: No module named 'openpyxl'`

Dependencies are not installed:

```bash
pip install -r requirements.txt
```

---

### Workflow runs in Actions but report is empty (0 data rows)

The target repository exists but has no commits, PRs, or workflow runs yet. Run the script against a repository with actual activity.

---

### `if-no-files-found: error` failure in Actions

The script exited before writing the Excel file — check the step logs immediately before the upload step for the actual Python error.

---

### Pylance shows `Import could not be resolved` for openpyxl

This is a VS Code IDE warning, not a runtime error. It means openpyxl is not installed in the Python interpreter VS Code is using. Fix by selecting the correct interpreter:

```
Cmd/Ctrl + Shift + P → Python: Select Interpreter → pick the env where you ran pip install
```

Or install into the active interpreter:

```bash
pip install openpyxl
```
