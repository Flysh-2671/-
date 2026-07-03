#!/usr/bin/env python3
import os
import textwrap
from datetime import datetime
import math

import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import requests

GITHUB_REPOSITORY = os.environ.get('GITHUB_REPOSITORY')  # owner/repo
GITHUB_TOKEN = os.environ.get('GITHUB_TOKEN')
TOTAL_TARGET = int(os.environ.get('TOTAL_TARGET', '4313'))

DATA_PATHS = ["data/tasks.csv", "data/tasks_sample.csv"]

SAMPLE_CSV = "id,completed_date,assignee,type,priority\n1,2026-01-05,alice,normal,medium\n2,2026-01-12,bob,normal,low\n3,2026-02-03,carol,normal,high\n"

REPORT_DIR = "reports"
os.makedirs(REPORT_DIR, exist_ok=True)


def create_issue_missing_data(repo, token):
    if not repo or not token:
        print("Missing GITHUB_REPOSITORY or GITHUB_TOKEN; cannot create issue.")
        return
    owner_repo = repo
    url = f"https://api.github.com/repos/{owner_repo}/issues"
    title = "[action required] Add data/tasks.csv for analysis"
    body = textwrap.dedent(f"""
    The scheduled analysis workflow couldn't find a `data/tasks.csv` file in the repository.

    Please add a CSV file at `data/tasks.csv` with at least the following columns: `id`, `completed_date` (YYYY-MM-DD). Optional columns: `assignee`, `type`, `priority`, `count` (if you want to specify multiple counts per row).

    Example (copy & paste):

    ```csv
    {SAMPLE_CSV}
    ```

    After you add the file, re-run the workflow (Actions -> analysis -> Run workflow).
    """)

    payload = {"title": title, "body": body}
    headers = {"Authorization": f"token {token}", "Accept": "application/vnd.github.v3+json"}
    r = requests.post(url, json=payload, headers=headers)
    if r.status_code in (200, 201):
        print("Created issue to request data/tasks.csv")
    else:
        print("Failed to create issue", r.status_code, r.text)


def find_data_file():
    for p in DATA_PATHS:
        if os.path.exists(p):
            return p
    return None


def load_data(path):
    df = pd.read_csv(path)
    df.columns = [c.lower() for c in df.columns]

    date_col = None
    for candidate in ['completed_date', 'date', '完成日期', 'completed_at']:
        if candidate in df.columns:
            date_col = candidate
            break
    if date_col is None:
        raise ValueError('No date column found in CSV. Expected one of completed_date,date,completed_at')

    df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
    df = df.dropna(subset=[date_col]).copy()
    df['year_month'] = df[date_col].dt.to_period('M').astype(str)
    return df, date_col


def summarize(df, date_col):
    total_completed = int(df.shape[0])
    if 'count' in df.columns:
        try:
            total_completed = int(df['count'].sum())
        except Exception:
            pass

    monthly = df.groupby('year_month').size().sort_index()
    now = pd.Timestamp.now()
    # Use the last 6 complete months ending with the previous month to avoid partial current month bias
    end_period = (now - pd.offsets.MonthBegin(1)).to_period('M')
    months = pd.period_range(end=end_period, periods=6)
    months_str = [m.strftime('%Y-%m') for m in months]
    monthly_recent = monthly.reindex(months_str, fill_value=0)

    # keep averages as floats for more accurate estimation
    avg_month_3 = monthly_recent[-3:].mean() if len(monthly_recent) >= 3 else monthly_recent.mean()
    avg_month_6 = monthly_recent.mean()

    remaining = max(0, TOTAL_TARGET - total_completed)

    per_month = (remaining + 6 - 1) // 6
    per_week = (remaining + 26 - 1) // 26
    per_workday = (remaining + 126 - 1) // 126

    # more robust current rate handling
    current_rate = avg_month_3 if avg_month_3 > 0 else avg_month_6
    if current_rate is None or current_rate <= 0:
        months_needed = None
    else:
        months_needed = math.ceil(remaining / current_rate)

    summary = {
        'total_completed': total_completed,
        'monthly_recent': monthly_recent.to_dict(),
        'avg_month_3': float(avg_month_3) if avg_month_3 is not None else 0.0,
        'avg_month_6': float(avg_month_6) if avg_month_6 is not None else 0.0,
        'remaining': remaining,
        'per_month': per_month,
        'per_week': per_week,
        'per_workday': per_workday,
        'months_needed_at_current_rate': months_needed,
    }
    return summary


def render_report(summary):
    lines = []
    lines.append(f"Analysis run: {datetime.utcnow().isoformat()}Z")
    lines.append(f"TOTAL_TARGET = {TOTAL_TARGET}")
    lines.append(f"Total completed (from data): {summary['total_completed']}")
    lines.append(f"Remaining to reach target: {summary['remaining']}")
    lines.append("")
    lines.append("Recent monthly completion (last 6 months):")
    for m, v in summary['monthly_recent'].items():
        lines.append(f"  {m}: {v}")
    lines.append("")
    lines.append(f"Average per month (last 3 months): {round(summary['avg_month_3'], 2)}")
    lines.append(f"Average per month (last 6 months): {round(summary['avg_month_6'], 2)}")
    lines.append("")
    lines.append("Suggested cadence to clear remaining (evenly over 6 months):")
    lines.append(f"  Per month: {summary['per_month']}")
    lines.append(f"  Per week: {summary['per_week']}")
    lines.append(f"  Per workday: {summary['per_workday']}")
    lines.append("")
    if summary['months_needed_at_current_rate'] is None:
        lines.append("At current pace, unable to produce a reliable months-to-clear estimate (insufficient or zero recent throughput).")
    else:
        lines.append(f"At current pace (avg of last 3 months), estimated months to clear remaining: {summary['months_needed_at_current_rate']}")
    text = "\n".join(lines)
    with open(os.path.join(REPORT_DIR, 'report.txt'), 'w', encoding='utf-8') as f:
        f.write(text)
    print(text)


def plot_monthly(summary):
    monthly = summary['monthly_recent']
    months = list(monthly.keys())
    vals = list(monthly.values())
    plt.figure(figsize=(8, 4))
    plt.bar(months, vals, color='#2b81e6')
    plt.title('Completed tasks by month (most recent 6 months)')
    plt.xlabel('Month')
    plt.ylabel('Completed tasks')
    plt.tight_layout()
    plt.savefig(os.path.join(REPORT_DIR, 'plot.png'))
    plt.close()


def main():
    data_file = find_data_file()
    if data_file is None or os.path.basename(data_file) == 'tasks_sample.csv':
        print('No data/tasks.csv found. Will create an issue to request the data.')
        if GITHUB_REPOSITORY and GITHUB_TOKEN:
            create_issue_missing_data(GITHUB_REPOSITORY, GITHUB_TOKEN)
        else:
            print('GITHUB_REPOSITORY or GITHUB_TOKEN not available in environment; cannot create issue automatically.')
        with open(os.path.join(REPORT_DIR, 'report.txt'), 'w', encoding='utf-8') as f:
            f.write('No data/tasks.csv found. Please add data and re-run.\n')
        return

    try:
        df, date_col = load_data(data_file)
    except Exception as e:
        print('Failed to load data:', e)
        with open(os.path.join(REPORT_DIR, 'report.txt'), 'w', encoding='utf-8') as f:
            f.write('Failed to load data: ' + str(e) + '\n')
        return

    summary = summarize(df, date_col)
    render_report(summary)
    plot_monthly(summary)


if __name__ == '__main__':
    main()
