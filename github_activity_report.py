"""
GitHub Activity Report Generator
=================================
Collects commits, pull requests, and workflow runs for a GitHub repository
and writes a consolidated Excel report.

Usage (local):
    export GH_PAT=<your_personal_access_token>
    export GH_REPO=<owner/repo>            # e.g. "octocat/Hello-World"
    python github_activity_report.py

In GitHub Actions the repo is auto-detected via GITHUB_REPOSITORY.

Optional env vars:
    GH_OUTPUT   Path for the output Excel file (default: github_activity_report.xlsx)
    GH_MAX_RUNS Max workflow runs to fetch   (default: 200)
"""

import os
import sys
import time
import logging

import requests
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

GH_PAT   = os.getenv("GH_PAT", "")
# Accept GH_REPO explicitly, or fall back to the variable GitHub Actions sets.
REPO     = os.getenv("GH_REPO") or os.getenv("GITHUB_REPOSITORY", "")
OUTPUT   = os.getenv("GH_OUTPUT", "github_activity_report.xlsx")
MAX_RUNS = int(os.getenv("GH_MAX_RUNS", "200"))

if "/" not in (REPO or ""):
    print(
        "ERROR: GH_REPO must be in 'owner/repo' format.\n"
        f"  Current value: {REPO!r}\n"
        "  Example:  export GH_REPO=vidyakar/whitedev"
    )
    sys.exit(1)

ORG, REPO_NAME = REPO.split("/", 1)
BASE_URL = "https://api.github.com"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# HTTP helpers
# ---------------------------------------------------------------------------

def _build_session() -> requests.Session:
    s = requests.Session()
    if not GH_PAT:
        log.warning("GH_PAT is not set — unauthenticated requests are heavily rate-limited")
    s.headers.update({
        "Authorization": f"Bearer {GH_PAT}" if GH_PAT else "",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    })
    return s


SESSION = _build_session()


def _wait_for_rate_limit(headers: dict) -> None:
    reset = int(headers.get("X-RateLimit-Reset", time.time() + 60))
    wait  = max(reset - int(time.time()), 1)
    log.warning("Rate-limited. Sleeping %ds …", wait)
    time.sleep(wait)


def gh_paginate(url: str, params: dict | None = None, max_items: int = 0) -> list:
    """Follow GitHub Link: next pagination and return all collected items."""
    params = dict(params or {})
    params.setdefault("per_page", 100)
    items: list = []
    page_url: str | None = url

    while page_url:
        for _ in range(5):
            resp = SESSION.get(page_url, params=params, timeout=30)
            if resp.status_code == 403 and "rate limit" in resp.text.lower():
                _wait_for_rate_limit(resp.headers)
                continue
            resp.raise_for_status()
            break

        payload = resp.json()
        # Some endpoints wrap the list in a top-level key
        if isinstance(payload, dict):
            for key in ("workflow_runs", "jobs"):
                if key in payload:
                    payload = payload[key]
                    break
            else:
                # Unexpected dict shape — treat as empty
                payload = []

        items.extend(payload)
        params = {}  # Link: next URL already contains all parameters

        if max_items and len(items) >= max_items:
            items = items[:max_items]
            break

        page_url = resp.links.get("next", {}).get("url")

    return items

# ---------------------------------------------------------------------------
# Data fetchers
# ---------------------------------------------------------------------------

def fetch_branches() -> dict[str, str]:
    """Return {branch_name: head_sha} for every branch in the repo."""
    log.info("Fetching branches …")
    branches = gh_paginate(f"{BASE_URL}/repos/{ORG}/{REPO_NAME}/branches")
    result = {b["name"]: b["commit"]["sha"] for b in branches}
    log.info("  %d branch(es) found", len(result))
    return result


def fetch_commits(branches: dict[str, str]) -> list[dict]:
    """
    Fetch commits for every branch, de-duplicate by SHA, and tag each
    commit with the first branch it was seen on.
    """
    log.info("Fetching commits for %d branch(es) …", len(branches))
    seen: dict[str, dict] = {}   # sha -> commit row

    for branch_name in branches:
        log.info("  branch: %s", branch_name)
        raw_commits = gh_paginate(
            f"{BASE_URL}/repos/{ORG}/{REPO_NAME}/commits",
            params={"sha": branch_name},
        )
        for c in raw_commits:
            sha = c["sha"]
            if sha in seen:
                continue  # already captured from an earlier branch

            git_commit = c.get("commit") or {}
            git_author = git_commit.get("author") or {}
            gh_author  = c.get("author") or {}

            seen[sha] = {
                "sha":     sha,
                # Prefer GitHub login; fall back to git author name
                "author":  gh_author.get("login") or git_author.get("name", ""),
                "message": git_commit.get("message", "").split("\n")[0],  # subject line only
                "date":    git_author.get("date", ""),
                "branch":  branch_name,
            }

    commits_list = list(seen.values())
    log.info("  %d unique commit(s) found", len(commits_list))
    return commits_list


def fetch_pull_requests() -> list[dict]:
    """Fetch all PRs (open + closed/merged) with full detail."""
    log.info("Fetching pull requests …")
    raw_prs = gh_paginate(
        f"{BASE_URL}/repos/{ORG}/{REPO_NAME}/pulls",
        params={"state": "all", "sort": "updated", "direction": "desc"},
    )
    log.info("  %d pull request(s) found", len(raw_prs))

    result = []
    for pr in raw_prs:
        merged_at = pr.get("merged_at") or ""
        status    = "merged" if merged_at else pr.get("state", "")

        result.append({
            "pr_id":       pr.get("number"),
            "pr_title":    pr.get("title", ""),
            "pr_author":   (pr.get("user") or {}).get("login", ""),
            "pr_status":   status,
            "pr_merged_at": merged_at,
            "merge_sha":   pr.get("merge_commit_sha") or "",
            "head_sha":    (pr.get("head") or {}).get("sha", ""),
        })
    return result


def _failure_detail(run_id: int) -> str:
    """
    For a failed workflow run, return '<job> / <step>' for the first
    failed job/step, or an empty string if detail is unavailable.
    """
    jobs = gh_paginate(f"{BASE_URL}/repos/{ORG}/{REPO_NAME}/actions/runs/{run_id}/jobs")
    for job in jobs:
        if job.get("conclusion") in ("failure", "timed_out", "cancelled"):
            job_name = job.get("name", "unknown job")
            for step in job.get("steps", []):
                if step.get("conclusion") in ("failure", "timed_out"):
                    return f"{job_name} / {step.get('name', 'unknown step')}"
            return job_name
    return ""


def fetch_workflow_runs() -> list[dict]:
    """Fetch workflow runs (up to MAX_RUNS) with enriched failure detail."""
    log.info("Fetching workflow runs (max %d) …", MAX_RUNS)
    raw_runs = gh_paginate(
        f"{BASE_URL}/repos/{ORG}/{REPO_NAME}/actions/runs",
        max_items=MAX_RUNS,
    )
    log.info("  %d workflow run(s) found", len(raw_runs))

    result = []
    for run in raw_runs:
        conclusion = run.get("conclusion") or ""
        run_id     = run["id"]

        if conclusion in ("failure", "timed_out", "startup_failure"):
            log.info("    fetching job detail for run %d …", run_id)
            detail = _failure_detail(run_id)
        else:
            detail = ""

        if conclusion == "failure":
            error_reason = f"Failed at: {detail}" if detail else "Failure (step detail unavailable)"
        elif conclusion == "timed_out":
            error_reason = f"Timed out at: {detail}" if detail else "Timed out"
        elif conclusion == "startup_failure":
            error_reason = "Workflow startup failure"
        elif conclusion == "cancelled":
            error_reason = "Cancelled"
        else:
            error_reason = ""

        result.append({
            "run_id":       run_id,
            "workflow":     run.get("name", ""),
            "status":       run.get("status", ""),
            "conclusion":   conclusion,
            "error_reason": error_reason,
            "event":        run.get("event", ""),
            "head_sha":     run.get("head_sha", ""),
            "run_author":   (run.get("triggering_actor") or run.get("actor") or {}).get("login", ""),
        })
    return result

# ---------------------------------------------------------------------------
# Dataset assembly
# ---------------------------------------------------------------------------

def build_dataset(
    commits: list[dict],
    prs:     list[dict],
    runs:    list[dict],
) -> list[dict]:
    """
    Merge commits, PRs, and workflow runs into one flat record set.

    - One anchor row per commit; each matched workflow run produces a separate row.
    - PRs are linked to commits via merge_sha (preferred) then head_sha.
    - PRs with no matching commit get their own row.
    - Workflow runs whose head_sha is not in the fetched commit list get their own row.
    """

    # Build lookup indexes
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

    covered_prs:  set[int] = set()
    covered_runs: set[int] = set()
    rows: list[dict] = []

    def _pr_fields(pr: dict | None) -> dict:
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

    def _run_fields(run: dict | None) -> dict:
        if not run:
            return {"Workflow Name": "", "Workflow Run ID": "",
                    "Workflow Status": "", "Trigger": "",
                    "Workflow Conclusion/Error Reason": ""}
        return {
            "Workflow Name":                    run["workflow"],
            "Workflow Run ID":                  run["run_id"],
            "Workflow Status":                  run["status"],
            "Trigger":                          run["event"],
            "Workflow Conclusion/Error Reason": run["error_reason"] or run["conclusion"],
        }

    # Commits as the primary anchor
    for commit in commits:
        sha = commit["sha"]
        matched_prs  = sha_to_prs.get(sha)  or [None]
        matched_runs = sha_to_runs.get(sha) or [None]

        base = {
            "Repository":    REPO_NAME,
            "Organization":  ORG,
            "Commit ID":     sha,
            "Commit Message": commit["message"],
            "Commit Author": commit["author"],
            "Commit Date":   commit["date"],
            "Branch":        commit["branch"],
        }

        for pr in matched_prs:
            if pr:
                covered_prs.add(pr["pr_id"])
            for run in matched_runs:
                if run:
                    covered_runs.add(run["run_id"])
                rows.append({**base, **_pr_fields(pr), **_run_fields(run)})

    # PRs not linked to any fetched commit
    for pr in prs:
        if pr["pr_id"] not in covered_prs:
            rows.append({
                "Repository":    REPO_NAME,
                "Organization":  ORG,
                "Commit ID":     "",
                "Commit Message": "",
                "Commit Author": "",
                "Commit Date":   "",
                "Branch":        "",
                **_pr_fields(pr),
                **_run_fields(None),
            })

    # Workflow runs whose SHA is not in any fetched commit
    for run in runs:
        if run["run_id"] not in covered_runs:
            rows.append({
                "Repository":    REPO_NAME,
                "Organization":  ORG,
                "Commit ID":     run["head_sha"],
                "Commit Message": "",
                "Commit Author": run["run_author"],
                "Commit Date":   "",
                "Branch":        "",
                **_pr_fields(None),
                **_run_fields(run),
            })

    return rows

# ---------------------------------------------------------------------------
# Excel output
# ---------------------------------------------------------------------------

COLUMN_ORDER = [
    "Repository", "Organization",
    "Commit ID", "Commit Message", "Commit Author", "Commit Date", "Branch",
    "PR ID", "PR Title", "PR Author", "PR Status", "PR Merge Date",
    "Workflow Name", "Workflow Run ID", "Workflow Status",
    "Workflow Conclusion/Error Reason", "Trigger",
]

_HEADER_FILL = PatternFill("solid", fgColor="1F3864")   # dark navy
_HEADER_FONT = Font(color="FFFFFF", bold=True)
_ALT_FILL    = PatternFill("solid", fgColor="DCE6F1")   # light blue
_ERROR_FILL  = PatternFill("solid", fgColor="FFCCCC")   # light red


def save_excel(rows: list[dict], path: str) -> None:
    df = pd.DataFrame(rows, columns=COLUMN_ORDER)

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Activity")
        ws = writer.sheets["Activity"]

        # Header formatting
        for col_idx, col_name in enumerate(COLUMN_ORDER, start=1):
            cell = ws.cell(row=1, column=col_idx)
            cell.font      = _HEADER_FONT
            cell.fill      = _HEADER_FILL
            cell.alignment = Alignment(horizontal="center", wrap_text=True)

        # Track maximum content width per column for auto-sizing
        col_widths = [len(c) for c in COLUMN_ORDER]

        for row_idx, row_data in enumerate(rows, start=2):
            is_alt   = (row_idx % 2 == 0)
            is_error = str(row_data.get("Workflow Status", "")).lower() in (
                "failure", "timed_out"
            )

            for col_idx, col_name in enumerate(COLUMN_ORDER, start=1):
                val  = row_data.get(col_name, "")
                cell = ws.cell(row=row_idx, column=col_idx, value=val)
                cell.alignment = Alignment(wrap_text=True)

                if is_error:
                    cell.fill = _ERROR_FILL
                elif is_alt:
                    cell.fill = _ALT_FILL

                col_widths[col_idx - 1] = max(col_widths[col_idx - 1], len(str(val)))

        # Apply column widths (capped at 60 chars)
        for col_idx, width in enumerate(col_widths, start=1):
            ws.column_dimensions[get_column_letter(col_idx)].width = min(width + 2, 60)

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

    log.info("Report saved → %s  (%d data rows)", path, len(rows))

# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main() -> None:
    if REPO == "ORG/REPO":
        log.error(
            "Repository not configured. "
            "Set GH_REPO=owner/repo (or run inside GitHub Actions where "
            "GITHUB_REPOSITORY is set automatically)."
        )
        sys.exit(1)

    log.info("=== GitHub Activity Report ===")
    log.info("Repository : %s/%s", ORG, REPO_NAME)
    log.info("Output     : %s", OUTPUT)

    branches = fetch_branches()
    commits  = fetch_commits(branches)
    prs      = fetch_pull_requests()
    runs     = fetch_workflow_runs()

    log.info("Assembling dataset …")
    rows = build_dataset(commits, prs, runs)
    log.info("Total rows : %d", len(rows))

    save_excel(rows, OUTPUT)
    print(f"\nDone. Report written to: {OUTPUT}")


if __name__ == "__main__":
    main()
