import os
import sys
import requests
import pandas as pd
import zipfile
import io
import re
import time
from datetime import datetime, timezone

# ---------------- CONFIGURATION ----------------
token = os.getenv("GH_PAT")
organization = os.getenv("GH_ORG")
per_page = 100
# -----------------------------------------------

# Validate required environment variables
if not token:
    print("ERROR: GH_PAT environment variable is not set.")
    sys.exit(1)

if not organization:
    print("ERROR: GH_ORG environment variable is not set.")
    sys.exit(1)

headers = {
    "Authorization": f"token {token}",
    "Accept": "application/vnd.github+json",
    "X-GitHub-Api-Version": "2022-11-28"
}


def make_request(url):
    """Make a GitHub API request with rate-limit and error handling."""
    for attempt in range(3):
        try:
            response = requests.get(url, headers=headers, timeout=30)

            if response.status_code == 401:
                print("ERROR: Invalid GH_PAT token. Check your secret.")
                sys.exit(1)

            if response.status_code == 403:
                remaining = response.headers.get("X-RateLimit-Remaining", "1")
                reset_ts = response.headers.get("X-RateLimit-Reset")
                if remaining == "0" and reset_ts:
                    wait = max(int(reset_ts) - int(time.time()) + 5, 5)
                    print(f"  Rate limit hit. Waiting {wait}s...")
                    time.sleep(wait)
                    continue
                print(f"  WARNING: 403 Forbidden for {url}")
                return None

            if response.status_code == 404:
                print(f"  WARNING: 404 Not Found for {url}")
                return None

            response.raise_for_status()
            return response

        except requests.exceptions.Timeout:
            print(f"  WARNING: Timeout on attempt {attempt + 1} for {url}")
            time.sleep(2)
        except requests.exceptions.RequestException as e:
            print(f"  WARNING: Request error: {e}")
            return None

    return None


def get_error_from_logs(logs_url):
    """Download zip logs and extract first error line."""
    response = make_request(logs_url)
    if not response or not response.content:
        return ""
    try:
        with zipfile.ZipFile(io.BytesIO(response.content)) as z:
            for filename in sorted(z.namelist()):
                if filename.endswith(".txt"):
                    with z.open(filename) as f:
                        for line in f:
                            decoded = line.decode("utf-8", errors="replace")
                            if re.search(r"\berror\b", decoded, re.IGNORECASE):
                                return decoded.strip()[:500]
    except zipfile.BadZipFile:
        pass
    return ""


# 1. Get all repos in the organization
print(f"Fetching repositories for: {organization}")
repos = []
page = 1
while True:
    url = f"https://api.github.com/orgs/{organization}/repos?per_page={per_page}&page={page}&sort=updated"
    response = make_request(url)
    if not response:
        break
    data = response.json()

    if isinstance(data, dict) and "message" in data:
        print(f"ERROR: GitHub API responded with: {data['message']}")
        sys.exit(1)

    if not isinstance(data, list) or not data:
        break

    for repo in data:
        repos.append(repo["full_name"])

    if len(data) < per_page:
        break
    page += 1
    time.sleep(0.5)

print(f"Found {len(repos)} repositories")

if not repos:
    print("ERROR: No repositories found. Check GH_ORG value and token permissions.")
    sys.exit(1)

# 2. Collect workflow runs for each repo
all_runs = []

for repo in repos:
    print(f"  Scanning: {repo}")
    page = 1
    while True:
        url = f"https://api.github.com/repos/{repo}/actions/runs?per_page={per_page}&page={page}"
        response = make_request(url)
        if not response:
            break

        data = response.json()
        runs = data.get("workflow_runs", [])
        if not runs:
            break

        for run in runs:
            # Calculate run duration
            duration_seconds = None
            try:
                created = datetime.fromisoformat(run["created_at"].replace("Z", "+00:00"))
                updated = datetime.fromisoformat(run["updated_at"].replace("Z", "+00:00"))
                duration_seconds = int((updated - created).total_seconds())
            except Exception:
                pass

            record = {
                "Repository":      repo,
                "Workflow Name":   run.get("name", ""),
                "Branch":          run.get("head_branch", ""),
                "Event":           run.get("event", ""),
                "Status":          run.get("status", ""),
                "Conclusion":      run.get("conclusion") or "in_progress",
                "Run Number":      run.get("run_number", ""),
                "Created At":      run.get("created_at", ""),
                "Updated At":      run.get("updated_at", ""),
                "Duration (s)":    duration_seconds,
                "Actor":           (run.get("actor") or {}).get("login", ""),
                "Run URL":         run.get("html_url", ""),
                "Failure Message": ""
            }

            if run.get("conclusion") == "failure":
                logs_url = run.get("logs_url", "")
                if logs_url:
                    record["Failure Message"] = get_error_from_logs(logs_url)

            all_runs.append(record)

        if len(runs) < per_page:
            break
        page += 1
        time.sleep(0.5)

print(f"Total workflow runs collected: {len(all_runs)}")

# 3. Build DataFrames
columns = ["Repository", "Workflow Name", "Branch", "Event", "Status",
           "Conclusion", "Run Number", "Created At", "Updated At",
           "Duration (s)", "Actor", "Run URL", "Failure Message"]

df = pd.DataFrame(all_runs, columns=columns) if all_runs else pd.DataFrame(columns=columns)

# Summary by conclusion
if not df.empty:
    summary = df["Conclusion"].fillna("unknown").value_counts().reset_index()
    summary.columns = ["Conclusion", "Count"]
else:
    summary = pd.DataFrame(columns=["Conclusion", "Count"])

# Summary by repository
if not df.empty:
    repo_summary = (
        df.groupby("Repository")["Conclusion"]
        .value_counts()
        .unstack(fill_value=0)
        .reset_index()
    )
else:
    repo_summary = pd.DataFrame(columns=["Repository"])

# Failed runs only
failed_df = df[df["Conclusion"] == "failure"].copy() if not df.empty else pd.DataFrame(columns=columns)

# 4. Write Excel report with multiple sheets
excel_file = "github_workflow_report.xlsx"
with pd.ExcelWriter(excel_file, engine="openpyxl") as writer:
    summary.to_excel(writer, sheet_name="Summary", index=False)
    repo_summary.to_excel(writer, sheet_name="By Repository", index=False)
    df.to_excel(writer, sheet_name="All Runs", index=False)
    if not failed_df.empty:
        failed_df.to_excel(writer, sheet_name="Failures", index=False)

print(f"Excel report saved: {excel_file}")
