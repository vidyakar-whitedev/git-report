import os
import requests
import pandas as pd
import zipfile
import io
import re
import time

# ---------------- CONFIG ----------------
token = os.getenv("GH_PAT")  # GitHub PAT from Secrets
organization_or_user = "vidyakar-whitedev"  # Change to your username/org
per_page = 100
# ---------------------------------------

headers = {"Authorization": f"token {token}"}

# Helper function to get all repositories
def get_all_repos():
    repos = []
    page = 1
    while True:
        url = f"https://api.github.com/users/{organization_or_user}/repos?per_page={per_page}&page={page}"
        response = requests.get(url, headers=headers)
        data = response.json()
        if not data:
            break
        for repo in data:
            repos.append(repo["full_name"])
        page += 1
        time.sleep(0.5)
    return repos

# Helper function to fetch events
def get_repo_events(repo):
    events_list = []
    page = 1
    while True:
        url = f"https://api.github.com/repos/{repo}/events?per_page={per_page}&page={page}"
        resp = requests.get(url, headers=headers)
        data = resp.json()
        if not data or isinstance(data, dict) and data.get("message"):
            break
        for event in data:
            events_list.append({
                "repository": repo,
                "event_type": event["type"],
                "created_at": event["created_at"],
                "actor": event["actor"]["login"],
                "details": str(event.get("payload"))
            })
        page += 1
        time.sleep(0.2)
    return events_list

# Helper function to fetch workflow runs
def get_workflow_runs(repo):
    runs_list = []
    page = 1
    while True:
        url = f"https://api.github.com/repos/{repo}/actions/runs?per_page={per_page}&page={page}"
        resp = requests.get(url, headers=headers)
        data = resp.json()
        runs = data.get("workflow_runs", [])
        if not runs:
            break
        for run in runs:
            run_record = {
                "repository": repo,
                "workflow_name": run["name"],
                "branch": run["head_branch"],
                "status": run["status"],
                "conclusion": run["conclusion"],
                "created_at": run["created_at"],
                "updated_at": run["updated_at"],
                "run_url": run["html_url"],
                "failure_message": ""
            }
            if run["conclusion"] == "failure":
                logs_url = run["logs_url"]
                log_resp = requests.get(logs_url, headers=headers)
                with zipfile.ZipFile(io.BytesIO(log_resp.content)) as z:
                    error_found = False
                    for filename in z.namelist():
                        if filename.endswith(".txt"):
                            with z.open(filename) as f:
                                for line in f:
                                    line = line.decode('utf-8')
                                    if re.search(r'error', line, re.IGNORECASE):
                                        run_record["failure_message"] = line.strip()
                                        error_found = True
                                        break
                        if error_found:
                            break
            runs_list.append(run_record)
        page += 1
        time.sleep(0.5)
    return runs_list

# Get all repositories
repos = get_all_repos()
print(f"Found {len(repos)} repositories")

all_events = []
all_runs = []

# Fetch events and workflow runs
for repo in repos:
    print(f"Processing repo: {repo}")
    all_events.extend(get_repo_events(repo))
    all_runs.extend(get_workflow_runs(repo))

# Create DataFrames
df_events = pd.DataFrame(all_events)
df_runs = pd.DataFrame(all_runs)

# Summary sheet
summary_data = {
    "Metric": ["Push Events", "Pull Requests", "Issues", "Workflow Success", "Workflow Failure", "Workflow Cancelled"],
    "Count": [
        len(df_events[df_events['event_type']=="PushEvent"]),
        len(df_events[df_events['event_type']=="PullRequestEvent"]),
        len(df_events[df_events['event_type']=="IssuesEvent"]),
        len(df_runs[df_runs['conclusion']=="success"]),
        len(df_runs[df_runs['conclusion']=="failure"]),
        len(df_runs[df_runs['conclusion']=="cancelled"])
    ]
}
df_summary = pd.DataFrame(summary_data)

# Write to Excel
excel_file = "github_full_activity_report.xlsx"
with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
    df_summary.to_excel(writer, sheet_name="Summary", index=False)
    df_events.to_excel(writer, sheet_name="GitHub Events", index=False)
    df_runs.to_excel(writer, sheet_name="Workflow Runs", index=False)

print(f"Excel report generated: {excel_file}")