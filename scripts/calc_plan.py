#!/usr/bin/env python3
import os
import textwrap
from datetime import datetime
import math
import io

import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import requests

# Optional decryption for password-protected Excel
try:
    import msoffcrypto
except Exception:
    msoffcrypto = None

GITHUB_REPOSITORY = os.environ.get('GITHUB_REPOSITORY')  # owner/repo
GITHUB_TOKEN = os.environ.get('GITHUB_TOKEN')
TOTAL_TARGET = int(os.environ.get('TOTAL_TARGET', '4313'))
EXCEL_PASSWORD = os.environ.get('EXCEL_PASSWORD')  # set this as a repo secret if your Excel is protected

DATA_DIR = 'data'
DATA_PATHS = [os.path.join(DATA_DIR, "tasks.csv"), os.path.join(DATA_DIR, "tasks_sample.csv")]

SAMPLE_CSV = "id,completed_date,assignee,type,priority\n1,2026-01-05,alice,normal,medium\n2,2026-01-12,bob,normal,low\n3,2026-02-03,carol,normal,high\n"

REPORT_DIR = "reports"
os.makedirs(REPORT_DIR, exist_ok=True)

DATE_CANDIDATES = ['completed_date', 'date', '完成日期', 'completed_at', '签约日期', '签约时间', '下单日期', '成交日期']


def create_issue_missing_data(repo, token):
    if not repo or not token:
        print("Missing GITHUB_REPOSITORY or GITHUB_TOKEN; cannot create issue.")
        return
    owner_repo = repo
    url = f"https://api.github.com/repos/{owner_repo}/issues"
    title = "[action required] Add data/tasks.csv or data/*.xlsx for analysis"
    body = textwrap.dedent(f"""
    The scheduled analysis workflow couldn't find a data file for analysis.

    Please add either a CSV file at `data/tasks.csv` or an Excel file (e.g. `data/已签预订单.xlsx`) with at least the following columns: `id`, `completed_date` (YYYY-MM-DD). Optional columns: `assignee`, `type`, `priority`, `count`.

    If your Excel is password-protected, add the password as a repository secret named `EXCEL_PASSWORD`.

    Example (CSV):

    ```csv
    {SAMPLE_CSV}
    ```

    After you add the file, re-run the workflow (Actions -> analysis -> Run workflow).
    """)

    payload = {"title": title, "body": body}
    headers = {"Authorization": f"token {token}", "Accept": "application/vnd.github.v3+json"}
    r = requests.post(url, json=payload, headers=headers)
    if r.status_code in (200, 201):
        print("Created issue to request data file")
    else:
        print("Failed to create issue", r.status_code, r.text)


def find_data_file():
    # Prefer the largest file in data/ among CSV and XLSX
    if not os.path.isdir(DATA_DIR):
        return None
    candidates = []
    for fn in os.listdir(DATA_DIR):
        lower = fn.lower()
        if lower.endswith('.csv') or lower.endswith('.xlsx') or lower.endswith('.xls'):
            path = os.path.join(DATA_DIR, fn)
            try:
                size = os.path.getsize(path)
            except Exception:
                size = 0
            candidates.append((size, path))
    if not candidates:
        return None
    # pick the largest file (most likely the full export)
    candidates.sort(reverse=True)
    return candidates[0][1]


def load_excel_with_password(path, password=None):
    # Try to open normally first
    try:
        return pd.read_excel(path, sheet_name=None)
    except Exception:
        # If msoffcrypto is available, try to decrypt using provided password
        if msoffcrypto is None:
            raise
        try:
            with open(path, 'rb') as f:
                file = msoffcrypto.OfficeFile(f)
                if password is not None:
                    file.load_key(password=password)
                else:
                    try:
                        file.load_key(password='')
                    except Exception:
                        raise
                bio = io.BytesIO()
                file.decrypt(bio)
                bio.seek(0)
                return pd.read_excel(bio, sheet_name=None)
        except Exception:
            raise


def load_data(path):
    lower = path.lower()
    if lower.endswith('.csv'):
        df = pd.read_csv(path)
        df.columns = [c.lower() for c in df.columns]
        date_col = None
        for candidate in DATE_CANDIDATES:
            if candidate in df.columns:
                date_col = candidate
                break
        if date_col is None:
            # try to find the column that parses best as dates
            best_col = None
            best_count = 0
            for col in df.columns:
                parsed = pd.to_datetime(df[col], errors='coerce')
                c = parsed.notna().sum()
                if c > best_count:
                    best_count = c
                    best_col = col
            if best_count >= max(1, int(0.05 * len(df))):
                date_col = best_col
        if date_col is None:
            raise ValueError('No date column found in CSV. Expected one of: ' + ','.join(DATE_CANDIDATES))
        df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
        df = df.dropna(subset=[date_col]).copy()
        df['year_month'] = df[date_col].dt.to_period('M').astype(str)
        data_source = os.path.basename(path)
        return df, date_col, data_source

    elif lower.endswith('.xlsx') or lower.endswith('.xls'):
        # read all sheets and pick the most promising one
        try:
            sheets = load_excel_with_password(path, EXCEL_PASSWORD)
        except Exception as e:
            raise ValueError(f'Failed to read Excel file {path}: {e}')

        best_sheet = None
        best_df = None
        best_rows = 0
        best_date_col = None

        for sheet_name, sheet_df in sheets.items():
            if sheet_df is None or sheet_df.empty:
                continue
            df_sheet = sheet_df.copy()
            # normalize column names
            df_sheet.columns = [str(c).lower().strip() for c in df_sheet.columns]
            # try to find a date column by name first
            date_col = None
            for candidate in DATE_CANDIDATES:
                if candidate in df_sheet.columns:
                    date_col = candidate
                    break
            # if not found, try to auto-detect by parseable date counts
            if date_col is None:
                best_col = None
                best_count = 0
                for col in df_sheet.columns:
                    try:
                        parsed = pd.to_datetime(df_sheet[col], errors='coerce')
                        c = int(parsed.notna().sum())
                    except Exception:
                        c = 0
                    if c > best_count:
                        best_count = c
                        best_col = col
                # choose if there are at least a few parseable dates or >=5% of rows
                if best_count >= max(5, int(0.05 * len(df_sheet))):
                    date_col = best_col

            if date_col is not None:
                # compute number of valid date rows
                parsed = pd.to_datetime(df_sheet[date_col], errors='coerce')
                valid = int(parsed.notna().sum())
                rows = len(df_sheet)
                # prefer sheet with more valid rows
                if valid > best_rows:
                    best_rows = valid
                    best_sheet = sheet_name
                    best_df = df_sheet
                    best_date_col = date_col

        if best_sheet is None:
            raise ValueError('No suitable sheet with parseable date column found in Excel file')

        # finalize
        df_final = best_df.copy()
        df_final[best_date_col] = pd.to_datetime(df_final[best_date_col], errors='coerce')
        df_final = df_final.dropna(subset=[best_date_col]).copy()
        df_final['year_month'] = df_final[best_date_col].dt.to_period('M').astype(str)
        data_source = os.path.basename(path) + '::' + str(best_sheet)
        return df_final, best_date_col, data_source

    else:
        raise ValueError('Unsupported data file format')


def summarize(df, date_col, data_source=None):
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
        'data_source': data_source,
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
    if summary.get('data_source'):
        lines.append(f"Data source: {summary['data_source']}")
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
        mn = summary['months_needed_at_current_rate']
        if mn > 60:
            lines.append("At current pace (avg of last 3 months), estimated months to clear remaining: >60 months (unrealistic). Please increase throughput or add resources.")
        else:
            lines.append(f"At current pace (avg of last 3 months), estimated months to clear remaining: {mn}")

    # add H2 monthly breakdown (even split)
    h2_total = TOTAL_TARGET
    baseline_h2 = None
    shortfall = None
    # if analysis_config.yaml exists, try to read values
    try:
        import yaml
        cfg_path = 'analysis_config.yaml'
        if os.path.exists(cfg_path):
            with open(cfg_path, 'r', encoding='utf-8') as cf:
                cfg = yaml.safe_load(cf)
            baseline_h2 = cfg.get('baseline_h2')
            shortfall = cfg.get('shortfall')
            h2_total = cfg.get('total_target', TOTAL_TARGET)
    except Exception:
        pass

    lines.append("")
    lines.append("H2 planning (7-12 month) — even distribution")
    lines.append(f"  H2 total target: {h2_total}")
    if baseline_h2 is not None:
        lines.append(f"  Baseline H2: {baseline_h2}")
    if shortfall is not None:
        lines.append(f"  Shortfall to cover: {shortfall}")
    # distribute remaining across 6 months
    remaining = summary['remaining']
    per_month = summary['per_month']
    months = ['2026-07','2026-08','2026-09','2026-10','2026-11','2026-12']
    # allocate evenly and adjust last month
    alloc = [per_month] * 6
    total_alloc = sum(alloc)
    diff = remaining - total_alloc
    alloc[-1] += diff
    lines.append("  Monthly breakdown:")
    for m, a in zip(months, alloc):
        lines.append(f"    {m}: {a}")

    text = "\n".join(lines)
    with open(os.path.join(REPORT_DIR, 'report.txt'), 'w', encoding='utf-8') as f:
        f.write(text)
    print(text)


def plot_monthly(summary):
    monthly = summary['monthly_recent']
    months = list(monthly.keys())
    vals = list(monthly.values())
    plt.figure(figsize=(10, 4))
    plt.bar(months, vals, color='#2b81e6')
    plt.title('Completed tasks by month (most recent 6 months)')
    plt.xlabel('Month')
    plt.ylabel('Completed tasks')
    # add a cumulative actual line
    try:
        actual_cum = [sum(vals[:i+1]) for i in range(len(vals))]
        plt.plot(months, actual_cum, color='black', marker='o', label='Actual cumulative')
        plt.legend()
    except Exception:
        pass
    plt.tight_layout()
    plt.savefig(os.path.join(REPORT_DIR, 'plot.png'))
    plt.close()


def main():
    data_file = find_data_file()
    if data_file is None:
        print('No data file found in data/. Will create an issue to request the data.')
        if GITHUB_REPOSITORY and GITHUB_TOKEN:
            create_issue_missing_data(GITHUB_REPOSITORY, GITHUB_TOKEN)
        else:
            print('GITHUB_REPOSITORY or GITHUB_TOKEN not available in environment; cannot create issue automatically.')
        with open(os.path.join(REPORT_DIR, 'report.txt'), 'w', encoding='utf-8') as f:
            f.write('No data file found. Please add data and re-run.\n')
        return

    print(f'Using data file: {data_file}')

    try:
        df, date_col, data_source = load_data(data_file)
    except Exception as e:
        print('Failed to load data:', e)
        with open(os.path.join(REPORT_DIR, 'report.txt'), 'w', encoding='utf-8') as f:
            f.write('Failed to load data: ' + str(e) + '\n')
        return

    summary = summarize(df, date_col, data_source=data_source)
    render_report(summary)
    plot_monthly(summary)


if __name__ == '__main__':
    main()
