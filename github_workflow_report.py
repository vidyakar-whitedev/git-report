import os
import requests
import pandas as pd
import zipfile
import io
import re
import time

# ---------------- CONFIGURATION ----------------
token = os.getenv("GH_PAT")
organization = "your_org_name"  # Replace with your GitHub org name
per_page = 100
# -----------------------------------------------

headers = {"Authorization": f"token {token}"}

# 1. Get all repos in organization
repos = []
page = 1
while True:
    url = f"https://api.github.com/orgs/{organization}/repos?per_page={per_page}&page={page}"
    response = requests.get(url, headers=headers)
    data = response.json()
    if not data:
        break
    for repo in data:
        repos.append(repo["full_name"])
    page += 1
    time.sleep(1)

print(f"Found {len(repos)} repositories in {organization}")

all_runs = []

for repo in repos:
    print(f"Fetching workflow runs for {repo}...")
    page = 1
    while True:
        url = f"https://api.github.com/repos/{repo}/actions/runs?per_page={per_page}&page={page}"
        response = requests.get(url, headers=headers)
        data = response.json()
        runs = data.get("workflow_runs", [])
        if not runs:
            break

        for run in runs:
            record = {
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
                # Download logs
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
                                        record["failure_message"] = line.strip()
                                        error_found = True
                                        break
                        if error_found:
                            break

            all_runs.append(record)

        if len(runs) < per_page:
            break
        page += 1
        time.sleep(1)

df = pd.DataFrame(all_runs)

# Summary counts
summary = df['conclusion'].value_counts().reset_index()
summary.columns = ['conclusion', 'count']

# Save Excel
excel_file = "github_workflow_report.xlsx"
with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
    summary.to_excel(writer, sheet_name="Summary", index=False)
    df.to_excel(writer, sheet_name="Details", index=False)

print(f"Excel report generated: {excel_file}")
