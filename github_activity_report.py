"""
GitHub Activity & Access Audit Report Generator
================================================
Collects commits, pull requests, workflow runs, and repository access control
across ALL repositories accessible via the provided PAT, then writes a
multi-sheet, colour-coded Excel audit report.

Usage (local):
    export GH_PAT=<your_personal_access_token>
    python github_activity_report.py

    # Optional overrides:
    export GH_REPO=owner/repo       # limit to one specific repo
    export GH_OUTPUT=my_report.xlsx # custom output path
    export GH_MAX_RUNS=500          # max workflow runs per repo (default 200)

In GitHub Actions, GH_REPO is set automatically from ${{ github.repository }}.

Output file: github_activity_audit_report.xlsx
Sheets:
  1. Activity             — commits + PRs + workflow runs per repo
  2. Access Control       — collaborators and teams per repo
  3. Failure Summary      — failed/timed-out runs with reasons and fix suggestions
"""

import io
import os
import sys
import time
import logging
import zipfile

import requests
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

GH_PAT    = os.getenv("GH_PAT", "")
GH_REPO   = os.getenv("GH_REPO") or os.getenv("GITHUB_REPOSITORY", "")
OUTPUT    = os.getenv("GH_OUTPUT", "github_activity_audit_report.xlsx")
MAX_RUNS  = int(os.getenv("GH_MAX_RUNS", "200"))
BASE_URL  = "https://api.github.com"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Excel colour palette
# ---------------------------------------------------------------------------

_HEADER_FILL   = PatternFill("solid", fgColor="1F3864")   # dark navy
_HEADER_FONT   = Font(color="FFFFFF", bold=True, size=10)
_ALT_FILL      = PatternFill("solid", fgColor="EEF2F7")   # very light grey-blue
_SUCCESS_FILL  = PatternFill("solid", fgColor="C6EFCE")   # light green
_SUCCESS_FONT  = Font(color="276221", bold=True)
_FAILURE_FILL  = PatternFill("solid", fgColor="FFC7CE")   # light red
_FAILURE_FONT  = Font(color="9C0006", bold=True)
_WARN_FILL     = PatternFill("solid", fgColor="FFEB9C")   # light yellow
_WARN_FONT     = Font(color="9C5700")
_THIN          = Border(
    left=Side(style="thin", color="D0D7DE"),
    right=Side(style="thin", color="D0D7DE"),
    top=Side(style="thin", color="D0D7DE"),
    bottom=Side(style="thin", color="D0D7DE"),
)

# ---------------------------------------------------------------------------
# HTTP helpers
# ---------------------------------------------------------------------------

def _build_session() -> requests.Session:
    s = requests.Session()
    if not GH_PAT:
        log.warning("GH_PAT is not set — unauthenticated requests are rate-limited to 60/hour")
    s.headers.update({
        "Authorization": f"Bearer {GH_PAT}" if GH_PAT else "",
        "Accept":        "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    })
    return s


SESSION = _build_session()


def _wait_rate_limit(headers: dict) -> None:
    reset = int(headers.get("X-RateLimit-Reset", time.time() + 60))
    wait  = max(reset - int(time.time()), 1)
    log.warning("Rate-limited — sleeping %ds …", wait)
    time.sleep(wait)


def _get(url: str, params: dict | None = None) -> dict | list | None:
    """Single GET with rate-limit retry. Returns None on 404."""
    for _ in range(5):
        resp = SESSION.get(url, params=params or {}, timeout=30)
        if resp.status_code == 403 and "rate limit" in resp.text.lower():
            _wait_rate_limit(resp.headers)
            continue
        if resp.status_code == 404:
            return None
        resp.raise_for_status()
        return resp.json()
    raise RuntimeError(f"Failed after retries: {url}")


def _paginate(url: str, params: dict | None = None, max_items: int = 0) -> list:
    """Follow Link: next pagination, unwrap common wrapper keys, return all items."""
    params = dict(params or {})
    params.setdefault("per_page", 100)
    items: list = []
    page_url: str | None = url

    while page_url:
        for _ in range(5):
            resp = SESSION.get(page_url, params=params, timeout=30)
            if resp.status_code == 403 and "rate limit" in resp.text.lower():
                _wait_rate_limit(resp.headers)
                continue
            if resp.status_code in (404, 403):
                return items          # no access — return what we have
            resp.raise_for_status()
            break

        payload = resp.json()
        if isinstance(payload, dict):
            for key in ("workflow_runs", "jobs", "repositories"):
                if key in payload:
                    payload = payload[key]
                    break
            else:
                payload = []

        items.extend(payload)
        params = {}

        if max_items and len(items) >= max_items:
            return items[:max_items]

        page_url = resp.links.get("next", {}).get("url")

    return items

# ---------------------------------------------------------------------------
# Repository discovery
# ---------------------------------------------------------------------------

def fetch_repos() -> list[dict]:
    """
    Return list of repo dicts.
    If GH_REPO is set, returns only that one repo.
    Otherwise, fetches all repos accessible to the PAT.
    """
    if GH_REPO:
        if "/" not in GH_REPO:
            log.error(
                "GH_REPO must be 'owner/repo' format. Got: %r\n"
                "  Example: export GH_REPO=octocat/Hello-World",
                GH_REPO,
            )
            sys.exit(1)
        owner, name = GH_REPO.split("/", 1)
        data = _get(f"{BASE_URL}/repos/{owner}/{name}")
        if not data:
            log.error("Repository not found or no access: %s", GH_REPO)
            sys.exit(1)
        return [data]

    log.info("Fetching all accessible repositories …")
    repos = _paginate(f"{BASE_URL}/user/repos", params={"affiliation": "owner,collaborator,organization_member"})
    log.info("  %d repository/repositories found", len(repos))
    return repos


def _repo_meta(repo: dict) -> dict:
    return {
        "org":        (repo.get("owner") or {}).get("login", ""),
        "name":       repo.get("name", ""),
        "full_name":  repo.get("full_name", ""),
        "visibility": repo.get("visibility", ""),
        "default_branch": repo.get("default_branch", "main"),
    }

# ---------------------------------------------------------------------------
# Commits
# ---------------------------------------------------------------------------

def fetch_branches(org: str, repo_name: str) -> dict[str, str]:
    branches = _paginate(f"{BASE_URL}/repos/{org}/{repo_name}/branches")
    return {b["name"]: b["commit"]["sha"] for b in branches}


def fetch_commits(org: str, repo_name: str, branches: dict[str, str]) -> list[dict]:
    seen: dict[str, dict] = {}
    for branch_name in branches:
        raw = _paginate(
            f"{BASE_URL}/repos/{org}/{repo_name}/commits",
            params={"sha": branch_name},
        )
        for c in raw:
            sha = c["sha"]
            if sha in seen:
                continue
            git_c   = c.get("commit") or {}
            git_a   = git_c.get("author") or {}
            gh_a    = c.get("author") or {}
            seen[sha] = {
                "sha":     sha,
                "author":  gh_a.get("login") or git_a.get("name", ""),
                "message": git_c.get("message", "").split("\n")[0],
                "date":    git_a.get("date", ""),
                "branch":  branch_name,
            }
    return list(seen.values())

# ---------------------------------------------------------------------------
# Pull Requests
# ---------------------------------------------------------------------------

def fetch_pull_requests(org: str, repo_name: str) -> list[dict]:
    raw = _paginate(
        f"{BASE_URL}/repos/{org}/{repo_name}/pulls",
        params={"state": "all", "sort": "updated", "direction": "desc"},
    )
    result = []
    for pr in raw:
        merged_at = pr.get("merged_at") or ""
        status    = "merged" if merged_at else pr.get("state", "")
        result.append({
            "pr_id":        pr.get("number"),
            "pr_title":     pr.get("title", ""),
            "pr_author":    (pr.get("user") or {}).get("login", ""),
            "pr_status":    status,
            "pr_merged_at": merged_at,
            "merge_sha":    pr.get("merge_commit_sha") or "",
            "head_sha":     (pr.get("head") or {}).get("sha", ""),
        })
    return result

# ---------------------------------------------------------------------------
# Failure diagnostics helpers
# ---------------------------------------------------------------------------

_FIX_RULES: list[tuple[list[str], str]] = [
    (
        ["install", "pip", "npm", "yarn", "dependency", "package", "module not found",
         "cannot find module", "importerror", "modulenotfounderror", "no module"],
        "Dependency error — run `pip install -r requirements.txt` or `npm install` and verify package versions.",
    ),
    (
        ["permission", "forbidden", "401", "403", "unauthorized", "access denied",
         "secret", "token", "credential"],
        "Permission/auth error — check that GH_PAT has the required scopes and that repository secrets are correctly configured.",
    ),
    (
        ["timeout", "timed out", "timed_out", "deadline exceeded"],
        "Timeout — increase the job `timeout-minutes`, optimise long-running steps, or split the job.",
    ),
    (
        ["syntax", "syntaxerror", "parse error", "unexpected token", "invalid syntax",
         "yaml", "yml"],
        "Syntax error — review the workflow YAML and any scripts for syntax mistakes.",
    ),
    (
        ["docker", "container", "image", "pull"],
        "Docker/container error — verify the image name/tag is correct and accessible.",
    ),
    (
        ["test", "assert", "spec", "jest", "pytest", "unittest", "rspec"],
        "Test failure — review failing test output in the workflow logs and fix the failing assertions.",
    ),
    (
        ["build", "compile", "tsc", "webpack", "gradle", "maven"],
        "Build error — check compiler output in the logs for the specific error and fix the source.",
    ),
    (
        ["deploy", "release", "publish"],
        "Deployment error — verify deployment credentials, target environment, and release configuration.",
    ),
]


def _suggest_fix(conclusion: str, job_name: str, step_name: str, log_snippet: str) -> str:
    if conclusion == "timed_out":
        return "Timeout — increase job `timeout-minutes` or optimise the long-running step."
    if conclusion == "startup_failure":
        return "Workflow startup failure — check workflow YAML syntax and runner availability."
    if conclusion == "cancelled":
        return "Run was cancelled manually or by a newer push — no fix needed unless unintended."

    haystack = " ".join([job_name, step_name, log_snippet]).lower()
    for keywords, suggestion in _FIX_RULES:
        if any(kw in haystack for kw in keywords):
            return suggestion

    return "Review the full workflow log in GitHub Actions for the specific error message."


def _fetch_failure_detail(org: str, repo_name: str, run_id: int) -> dict:
    """
    Drill into jobs and steps for a failed run.
    Returns: failed_job, failed_step, log_snippet, error_line, suggested_fix
    """
    jobs = _paginate(f"{BASE_URL}/repos/{org}/{repo_name}/actions/runs/{run_id}/jobs")

    failed_job  = ""
    failed_step = ""
    log_snippet = ""
    error_line  = "See GitHub Actions logs"

    for job in jobs:
        if job.get("conclusion") in ("failure", "timed_out", "cancelled"):
            failed_job = job.get("name", "unknown job")
            for step in job.get("steps", []):
                if step.get("conclusion") in ("failure", "timed_out"):
                    failed_step = step.get("name", "unknown step")
                    break
            break

    # Attempt to pull a log snippet (GitHub returns a zip of text log files)
    log_url = f"{BASE_URL}/repos/{org}/{repo_name}/actions/runs/{run_id}/logs"
    try:
        resp = SESSION.get(log_url, timeout=30, allow_redirects=True)
        if resp.status_code == 200 and resp.content:
            with zipfile.ZipFile(io.BytesIO(resp.content)) as zf:
                for fname in zf.namelist():
                    # Look for the log file that matches the failed job
                    if failed_job.lower().replace(" ", "_") in fname.lower() or not failed_job:
                        raw_log = zf.read(fname).decode("utf-8", errors="replace")
                        # Find lines that look like errors
                        error_lines = [
                            ln.strip() for ln in raw_log.splitlines()
                            if any(kw in ln.lower() for kw in (
                                "error", "fatal", "failed", "exception",
                                "traceback", "stderr", "syntaxerror",
                            ))
                        ]
                        if error_lines:
                            log_snippet = error_lines[0][:200]
                            # Try to find a line number pattern like "line 42" or ":42:"
                            for ln in error_lines[:10]:
                                import re
                                m = re.search(r"(?:line |:)(\d+)", ln, re.IGNORECASE)
                                if m:
                                    error_line = f"Line {m.group(1)}"
                                    break
                        break
    except Exception:
        pass   # log download is best-effort; not all tokens have log access

    suggested_fix = _suggest_fix("failure", failed_job, failed_step, log_snippet)
    return {
        "failed_job":     failed_job,
        "failed_step":    failed_step,
        "log_snippet":    log_snippet,
        "error_line":     error_line,
        "suggested_fix":  suggested_fix,
    }

# ---------------------------------------------------------------------------
# Workflow runs
# ---------------------------------------------------------------------------

def fetch_workflow_runs(org: str, repo_name: str) -> list[dict]:
    log.info("    Fetching workflow runs (max %d) …", MAX_RUNS)
    raw_runs = _paginate(
        f"{BASE_URL}/repos/{org}/{repo_name}/actions/runs",
        max_items=MAX_RUNS,
    )
    log.info("      %d run(s) found", len(raw_runs))

    result = []
    for run in raw_runs:
        conclusion = run.get("conclusion") or ""
        run_id     = run["id"]
        failure_detail: dict = {}

        if conclusion in ("failure", "timed_out", "startup_failure"):
            log.info("      fetching failure detail for run %d …", run_id)
            failure_detail = _fetch_failure_detail(org, repo_name, run_id)
        else:
            failure_detail = {
                "failed_job":    "",
                "failed_step":   "",
                "log_snippet":   "",
                "error_line":    "",
                "suggested_fix": "",
            }

        # Build a human-readable failure reason
        if conclusion == "failure":
            parts = []
            if failure_detail["failed_job"]:
                parts.append(f"Job '{failure_detail['failed_job']}'")
            if failure_detail["failed_step"]:
                parts.append(f"step '{failure_detail['failed_step']}' failed")
            if failure_detail["log_snippet"]:
                parts.append(f"— {failure_detail['log_snippet'][:120]}")
            failure_reason = " ".join(parts) if parts else "Failure (see logs)"
        elif conclusion == "timed_out":
            failure_reason = f"Timed out in job '{failure_detail['failed_job']}'" if failure_detail["failed_job"] else "Timed out"
            failure_detail["suggested_fix"] = _suggest_fix("timed_out", failure_detail["failed_job"], "", "")
        elif conclusion == "startup_failure":
            failure_reason = "Workflow startup failure"
            failure_detail["suggested_fix"] = _suggest_fix("startup_failure", "", "", "")
        elif conclusion == "cancelled":
            failure_reason = "Cancelled"
            failure_detail["suggested_fix"] = _suggest_fix("cancelled", "", "", "")
        else:
            failure_reason = ""

        result.append({
            "run_id":         run_id,
            "workflow":       run.get("name", ""),
            "status":         run.get("status", ""),
            "conclusion":     conclusion,
            "event":          run.get("event", ""),
            "head_sha":       run.get("head_sha", ""),
            "run_author":     (run.get("triggering_actor") or run.get("actor") or {}).get("login", ""),
            "failure_reason": failure_reason,
            **failure_detail,
        })

        # Alert for failures
        if conclusion in ("failure", "timed_out", "startup_failure"):
            log.warning(
                "WORKFLOW FAILURE DETECTED in %s/%s  (Run ID: %d)  — %s",
                org, repo_name, run_id, failure_reason or conclusion,
            )

    return result

# ---------------------------------------------------------------------------
# Access control
# ---------------------------------------------------------------------------

def fetch_access_control(org: str, repo_name: str) -> list[dict]:
    """Fetch collaborators (users) and teams with their permission levels."""
    records: list[dict] = []

    # Collaborators
    collaborators = _paginate(
        f"{BASE_URL}/repos/{org}/{repo_name}/collaborators",
        params={"affiliation": "all"},
    )
    for user in collaborators:
        perms = user.get("permissions") or {}
        if perms.get("admin"):
            level = "Admin"
        elif perms.get("maintain"):
            level = "Maintain"
        elif perms.get("push"):
            level = "Write"
        elif perms.get("triage"):
            level = "Triage"
        else:
            level = "Read"

        records.append({
            "org":         org,
            "repo":        repo_name,
            "entity_name": user.get("login", ""),
            "entity_type": "User",
            "permission":  level,
            "has_admin":   "Yes" if perms.get("admin") else "No",
            "can_delete":  "Yes" if perms.get("admin") else "No",
        })

    # Teams (only works for org repos; 404 for personal repos)
    teams = _paginate(f"{BASE_URL}/repos/{org}/{repo_name}/teams")
    for team in teams:
        perm = team.get("permission", "pull")
        level_map = {
            "admin": "Admin",
            "maintain": "Maintain",
            "push": "Write",
            "triage": "Triage",
            "pull": "Read",
        }
        level = level_map.get(perm, perm.capitalize())
        records.append({
            "org":         org,
            "repo":        repo_name,
            "entity_name": team.get("name", ""),
            "entity_type": "Team",
            "permission":  level,
            "has_admin":   "Yes" if perm == "admin" else "No",
            "can_delete":  "Yes" if perm == "admin" else "No",
        })

    return records

# ---------------------------------------------------------------------------
# Dataset assembly
# ---------------------------------------------------------------------------

ACTIVITY_COLUMNS = [
    "Repository Name", "Organization", "Visibility", "Default Branch",
    "Commit ID", "Commit Message", "Commit Author", "Commit Date", "Branch",
    "PR ID", "PR Title", "PR Author", "PR Status", "PR Merge Date",
    "Workflow Name", "Workflow Run ID", "Workflow Trigger",
    "Workflow Status", "Workflow Conclusion",
    "Failure Reason", "Failed Job", "Failed Step", "Error Line", "Suggested Fix",
]

ACCESS_COLUMNS = [
    "Repository Name", "Organization", "User/Team Name", "Type",
    "Permission Level", "Has Admin", "Has Delete Access",
]

FAILURE_COLUMNS = [
    "Repository Name", "Organization", "Workflow Name", "Workflow Run ID",
    "Workflow Trigger", "Workflow Status", "Workflow Conclusion",
    "Failure Reason", "Failed Job", "Failed Step", "Error Line", "Suggested Fix",
]

ALERTS_COLUMNS = [
    "#", "Repository", "Organization", "Workflow Name", "Run ID",
    "Trigger", "Conclusion", "Failure Reason", "Failed Job", "Failed Step",
    "Error Line", "Suggested Fix",
]


def build_activity_rows(
    repo_meta: dict,
    commits:   list[dict],
    prs:       list[dict],
    runs:      list[dict],
) -> list[dict]:
    org       = repo_meta["org"]
    repo_name = repo_meta["name"]
    vis       = repo_meta["visibility"]
    branch    = repo_meta["default_branch"]

    sha_to_prs: dict[str, list[dict]] = {}
    for pr in prs:
        for key in ("merge_sha", "head_sha"):
            sha = pr.get(key, "")
            if sha:
                sha_to_prs.setdefault(sha, []).append(pr)

    sha_to_runs: dict[str, list[dict]] = {}
    for run in runs:
        sha = run.get("head_sha", "")
        if sha:
            sha_to_runs.setdefault(sha, []).append(run)

    covered_prs:  set = set()
    covered_runs: set = set()
    rows: list[dict] = []

    def _pr_cols(pr: dict | None) -> dict:
        if not pr:
            return {"PR ID": "", "PR Title": "", "PR Author": "",
                    "PR Status": "", "PR Merge Date": ""}
        return {
            "PR ID":        pr["pr_id"],
            "PR Title":     pr["pr_title"],
            "PR Author":    pr["pr_author"],
            "PR Status":    pr["pr_status"],
            "PR Merge Date": pr["pr_merged_at"],
        }

    def _run_cols(run: dict | None) -> dict:
        if not run:
            return {
                "Workflow Name": "", "Workflow Run ID": "",
                "Workflow Trigger": "", "Workflow Status": "", "Workflow Conclusion": "",
                "Failure Reason": "", "Failed Job": "", "Failed Step": "",
                "Error Line": "", "Suggested Fix": "",
            }
        return {
            "Workflow Name":       run["workflow"],
            "Workflow Run ID":     run["run_id"],
            "Workflow Trigger":    run["event"],
            "Workflow Status":     run["status"],
            "Workflow Conclusion": run["conclusion"],
            "Failure Reason":      run.get("failure_reason", ""),
            "Failed Job":          run.get("failed_job", ""),
            "Failed Step":         run.get("failed_step", ""),
            "Error Line":          run.get("error_line", ""),
            "Suggested Fix":       run.get("suggested_fix", ""),
        }

    # Commits as the primary anchor
    for commit in commits:
        sha = commit["sha"]
        matched_prs  = sha_to_prs.get(sha)  or [None]
        matched_runs = sha_to_runs.get(sha) or [None]

        base = {
            "Repository Name": repo_name,
            "Organization":    org,
            "Visibility":      vis,
            "Default Branch":  branch,
            "Commit ID":       sha,
            "Commit Message":  commit["message"],
            "Commit Author":   commit["author"],
            "Commit Date":     commit["date"],
            "Branch":          commit["branch"],
        }

        for pr in matched_prs:
            if pr:
                covered_prs.add(pr["pr_id"])
            for run in matched_runs:
                if run:
                    covered_runs.add(run["run_id"])
                rows.append({**base, **_pr_cols(pr), **_run_cols(run)})

    # Orphan PRs
    for pr in prs:
        if pr["pr_id"] not in covered_prs:
            rows.append({
                "Repository Name": repo_name, "Organization": org,
                "Visibility": vis, "Default Branch": branch,
                "Commit ID": "", "Commit Message": "", "Commit Author": "",
                "Commit Date": "", "Branch": "",
                **_pr_cols(pr), **_run_cols(None),
            })

    # Orphan workflow runs
    for run in runs:
        if run["run_id"] not in covered_runs:
            rows.append({
                "Repository Name": repo_name, "Organization": org,
                "Visibility": vis, "Default Branch": branch,
                "Commit ID": run["head_sha"], "Commit Message": "",
                "Commit Author": run["run_author"], "Commit Date": "", "Branch": "",
                **_pr_cols(None), **_run_cols(run),
            })

    return rows


def build_access_rows(org: str, repo_name: str, access: list[dict]) -> list[dict]:
    rows = []
    for a in access:
        rows.append({
            "Repository Name":  repo_name,
            "Organization":     org,
            "User/Team Name":   a["entity_name"],
            "Type":             a["entity_type"],
            "Permission Level": a["permission"],
            "Has Admin":        a["has_admin"],
            "Has Delete Access": a["can_delete"],
        })
    return rows


def build_failure_rows(org: str, repo_name: str, runs: list[dict]) -> list[dict]:
    rows = []
    for run in runs:
        if run["conclusion"] in ("failure", "timed_out", "startup_failure"):
            rows.append({
                "Repository Name":   repo_name,
                "Organization":      org,
                "Workflow Name":     run["workflow"],
                "Workflow Run ID":   run["run_id"],
                "Workflow Trigger":  run["event"],
                "Workflow Status":   run["status"],
                "Workflow Conclusion": run["conclusion"],
                "Failure Reason":    run.get("failure_reason", ""),
                "Failed Job":        run.get("failed_job", ""),
                "Failed Step":       run.get("failed_step", ""),
                "Error Line":        run.get("error_line", ""),
                "Suggested Fix":     run.get("suggested_fix", ""),
            })
    return rows

# ---------------------------------------------------------------------------
# Excel writer
# ---------------------------------------------------------------------------

def _write_sheet(
    ws,
    columns:        list[str],
    rows:           list[dict],
    conclusion_col: str = "",
) -> None:
    """Write headers + data rows onto an openpyxl worksheet with full formatting."""

    # --- Header row ---
    for col_idx, col_name in enumerate(columns, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font      = _HEADER_FONT
        cell.fill      = _HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border    = _THIN

    ws.row_dimensions[1].height = 28

    # Track max widths for auto-sizing
    col_widths = [len(c) + 2 for c in columns]

    # --- Data rows ---
    for row_idx, row_data in enumerate(rows, start=2):
        conclusion = str(row_data.get(conclusion_col, "")).lower() if conclusion_col else ""
        is_alt     = (row_idx % 2 == 0)

        for col_idx, col_name in enumerate(columns, start=1):
            val  = row_data.get(col_name, "")
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border    = _THIN

            # Row-level fill: conclusion drives colour for the whole row
            if conclusion in ("failure", "timed_out", "startup_failure"):
                cell.fill = _FAILURE_FILL
            elif conclusion == "success":
                cell.fill = _SUCCESS_FILL
            elif conclusion in ("cancelled", "skipped"):
                cell.fill = _WARN_FILL
            elif is_alt:
                cell.fill = _ALT_FILL

            # Conclusion cell gets bold coloured font too
            if col_name == conclusion_col:
                if conclusion in ("failure", "timed_out", "startup_failure"):
                    cell.font = _FAILURE_FONT
                elif conclusion == "success":
                    cell.font = _SUCCESS_FONT
                elif conclusion in ("cancelled", "skipped"):
                    cell.font = _WARN_FONT

            col_widths[col_idx - 1] = max(col_widths[col_idx - 1], min(len(str(val)), 60))

    # --- Column widths ---
    for col_idx, width in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = min(width + 2, 62)

    # --- Freeze header + auto-filter ---
    ws.freeze_panes = "A2"
    if rows:
        ws.auto_filter.ref = ws.dimensions


def _write_alerts_sheet(ws, alert_rows: list[dict]) -> None:
    """
    Write the Failure Alerts sheet.
    Row 1  — merged banner:  "⚠  N WORKFLOW FAILURE(S) DETECTED"
    Row 2  — blank spacer
    Row 3  — column headers
    Row 4+ — one alert per row, all highlighted red
    """
    n_cols = len(ALERTS_COLUMNS)
    total  = len(alert_rows)

    # ── Banner row ────────────────────────────────────────────────────────────
    banner_text = f"⚠   {total} WORKFLOW FAILURE(S) DETECTED" if total else "✔   NO WORKFLOW FAILURES DETECTED"
    banner_fill = PatternFill("solid", fgColor="9C0006") if total else PatternFill("solid", fgColor="375623")
    banner_font = Font(color="FFFFFF", bold=True, size=14)

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
    banner_cell = ws.cell(row=1, column=1, value=banner_text)
    banner_cell.font      = banner_font
    banner_cell.fill      = banner_fill
    banner_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 36

    # ── Spacer row ────────────────────────────────────────────────────────────
    ws.row_dimensions[2].height = 8

    # ── Sub-header: generated timestamp ──────────────────────────────────────
    from datetime import datetime, timezone
    now_utc = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=n_cols)
    ts_cell = ws.cell(row=2, column=1, value=f"Generated: {now_utc}")
    ts_cell.font      = Font(color="9C0006" if total else "375623", italic=True, size=9)
    ts_cell.alignment = Alignment(horizontal="right")

    # ── Column header row (row 3) ─────────────────────────────────────────────
    HEADER_FILL_A = PatternFill("solid", fgColor="7B0000")   # deep red for alert sheet headers
    HEADER_FONT_A = Font(color="FFFFFF", bold=True, size=10)
    thin = Border(
        left=Side(style="thin", color="D0D7DE"),
        right=Side(style="thin", color="D0D7DE"),
        top=Side(style="thin", color="D0D7DE"),
        bottom=Side(style="thin", color="D0D7DE"),
    )

    for col_idx, col_name in enumerate(ALERTS_COLUMNS, start=1):
        cell = ws.cell(row=3, column=col_idx, value=col_name)
        cell.font      = HEADER_FONT_A
        cell.fill      = HEADER_FILL_A
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border    = thin
    ws.row_dimensions[3].height = 24

    # ── Data rows (row 4+) ────────────────────────────────────────────────────
    ROW_FILL  = PatternFill("solid", fgColor="FFC7CE")   # light red
    ROW_FONT  = Font(color="9C0006")
    ALT_FILL2 = PatternFill("solid", fgColor="FFD7DC")   # slightly darker red alternate

    col_widths = [len(c) + 2 for c in ALERTS_COLUMNS]

    for row_idx, row_data in enumerate(alert_rows, start=4):
        fill = ALT_FILL2 if row_idx % 2 == 0 else ROW_FILL
        for col_idx, col_name in enumerate(ALERTS_COLUMNS, start=1):
            val  = row_data.get(col_name, "")
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.fill      = fill
            cell.font      = ROW_FONT
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border    = thin
            col_widths[col_idx - 1] = max(col_widths[col_idx - 1], min(len(str(val)), 60))

    # ── Column widths ─────────────────────────────────────────────────────────
    for col_idx, width in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = min(width + 2, 62)

    # ── Freeze at row 4 (keep banner + headers visible) ──────────────────────
    ws.freeze_panes = "A4"
    if alert_rows:
        # auto-filter on header row
        ws.auto_filter.ref = (
            f"A3:{get_column_letter(n_cols)}{3 + len(alert_rows)}"
        )


def save_excel(
    activity_rows: list[dict],
    access_rows:   list[dict],
    failure_rows:  list[dict],
    alert_rows:    list[dict],
    path:          str,
) -> None:
    log.info("Writing Excel report → %s", path)

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        # Create all sheets via placeholder DataFrames
        pd.DataFrame(columns=ACTIVITY_COLUMNS).to_excel(writer, index=False, sheet_name="Activity")
        pd.DataFrame(columns=ACCESS_COLUMNS).to_excel(writer, index=False, sheet_name="Access Control")
        pd.DataFrame(columns=FAILURE_COLUMNS).to_excel(writer, index=False, sheet_name="Failure Summary")
        pd.DataFrame(columns=ALERTS_COLUMNS).to_excel(writer, index=False, sheet_name="Failure Alerts")

        wb = writer.book

        _write_sheet(
            wb["Activity"],
            ACTIVITY_COLUMNS,
            activity_rows,
            conclusion_col="Workflow Conclusion",
        )
        _write_sheet(
            wb["Access Control"],
            ACCESS_COLUMNS,
            access_rows,
        )
        _write_sheet(
            wb["Failure Summary"],
            FAILURE_COLUMNS,
            failure_rows,
            conclusion_col="Workflow Conclusion",
        )
        _write_alerts_sheet(wb["Failure Alerts"], alert_rows)

        # Tab colours
        wb["Activity"].sheet_properties.tabColor        = "1F3864"   # navy
        wb["Access Control"].sheet_properties.tabColor  = "375623"   # green
        wb["Failure Summary"].sheet_properties.tabColor = "9C0006"   # red
        wb["Failure Alerts"].sheet_properties.tabColor  = "FF0000"   # bright red

        # Open on Failure Alerts if there are failures, else Activity
        wb.active = wb["Failure Alerts"] if alert_rows else wb["Activity"]

    log.info(
        "Report saved → %s  |  Activity: %d  |  Access: %d  |  Failures: %d  |  Alerts: %d",
        path, len(activity_rows), len(access_rows), len(failure_rows), len(alert_rows),
    )

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    log.info("=== GitHub Activity & Access Audit Report ===")
    log.info("Output : %s", OUTPUT)

    repos = fetch_repos()

    all_activity: list[dict] = []
    all_access:   list[dict] = []
    all_failures: list[dict] = []
    all_alerts:   list[dict] = []   # structured rows for the Failure Alerts sheet

    for repo in repos:
        meta = _repo_meta(repo)
        org       = meta["org"]
        repo_name = meta["name"]

        log.info("--- %s/%s  [%s] ---", org, repo_name, meta["visibility"])

        # Branches + commits
        log.info("  Fetching branches …")
        branches = fetch_branches(org, repo_name)
        log.info("    %d branch(es)", len(branches))

        log.info("  Fetching commits …")
        commits = fetch_commits(org, repo_name, branches)
        log.info("    %d unique commit(s)", len(commits))

        # PRs
        log.info("  Fetching pull requests …")
        prs = fetch_pull_requests(org, repo_name)
        log.info("    %d PR(s)", len(prs))

        # Workflow runs
        runs = fetch_workflow_runs(org, repo_name)

        # Access control
        log.info("  Fetching access control …")
        access = fetch_access_control(org, repo_name)
        log.info("    %d user(s)/team(s)", len(access))

        # Assemble rows
        activity_rows = build_activity_rows(meta, commits, prs, runs)
        access_rows   = build_access_rows(org, repo_name, access)
        failure_rows  = build_failure_rows(org, repo_name, runs)

        all_activity.extend(activity_rows)
        all_access.extend(access_rows)
        all_failures.extend(failure_rows)

        # Build structured alert rows for the Failure Alerts sheet
        for idx, run in enumerate(
            (r for r in runs if r["conclusion"] in ("failure", "timed_out", "startup_failure")),
            start=len(all_alerts) + 1,
        ):
            all_alerts.append({
                "#":                  idx,
                "Repository":         repo_name,
                "Organization":       org,
                "Workflow Name":      run["workflow"],
                "Run ID":             run["run_id"],
                "Trigger":            run["event"],
                "Conclusion":         run["conclusion"],
                "Failure Reason":     run.get("failure_reason", run["conclusion"]),
                "Failed Job":         run.get("failed_job", ""),
                "Failed Step":        run.get("failed_step", ""),
                "Error Line":         run.get("error_line", ""),
                "Suggested Fix":      run.get("suggested_fix", ""),
            })

    log.info("=== Summary ===")
    log.info("  Repositories : %d", len(repos))
    log.info("  Activity rows: %d", len(all_activity))
    log.info("  Access rows  : %d", len(all_access))
    log.info("  Failure rows : %d", len(all_failures))
    log.info("  Alert rows   : %d", len(all_alerts))

    if all_alerts:
        log.warning("=== FAILURE ALERTS ===")
        for a in all_alerts:
            log.warning(
                "  [%d] %s/%s | Run ID: %s | %s | %s",
                a["#"], a["Organization"], a["Repository"],
                a["Run ID"], a["Workflow Name"], a["Failure Reason"],
            )

    save_excel(all_activity, all_access, all_failures, all_alerts, OUTPUT)
    print(f"\nDone. Report written to: {OUTPUT}")

    if all_alerts:
        print(f"\n{'='*62}")
        print(f"  ⚠  {len(all_alerts)} WORKFLOW FAILURE(S) DETECTED")
        print(f"{'='*62}")
        for a in all_alerts:
            print(
                f"  [{a['#']:>2}] {a['Organization']}/{a['Repository']}"
                f"  |  Run {a['Run ID']}"
                f"  |  {a['Workflow Name']}"
                f"  |  {a['Failure Reason']}"
            )
        print(f"\n  See the 'Failure Alerts' sheet in {OUTPUT} for full details.")


if __name__ == "__main__":
    main()
