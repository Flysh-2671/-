""" parse_and_plan.py

Parses the workbook '已签预订单6.30.xlsx', extracts sheet previews and basic stats, computes H2 planning (target = 3714 + 599 = 4313 万) and writes outputs to /outputs/. """

import sys, os import pandas as pd from datetime import datetime

TARGET_H2_BASE = 3714.0 BACKFILL = 599.0 R_TOTAL = TARGET_H2_BASE + BACKFILL SIGN_RATE = 0.50 VISIT2LEAD_DEFAULTS = [0.20, 0.10] PREVIEW_ROWS = 50

INPUT_FILENAME = sys.argv[1] if len(sys.argv) > 1 else '已签预订单6.30.xlsx' OUTPUT_DIR = 'outputs' OUTPUT_XLSX = os.path.join(OUTPUT_DIR, '已签预订单6.30.planned.xlsx') REPORT_TXT = os.path.join(OUTPUT_DIR, 'report.txt')

os.makedirs(OUTPUT_DIR, exist_ok=True)

try: all_sheets = pd.read_excel(INPUT_FILENAME, sheet_name=None, engine='openpyxl') except Exception as e: print('Error reading workbook:', e) raise

def find_amount_col(cols): keys = ['金额','total','amount','price','合同','实收','成交额','订单金额'] for c in cols: low = str(c).lower() for k in keys: if k in low: return c return None

def find_date_col(cols): keys = ['日期','时间','date','month','月份'] for c in cols: low = str(c).lower() for k in keys: if k in low: return c return None

sheet_stats = {} for name, df in all_sheets.items(): cols = list(df.columns) amt_col = find_amount_col(cols) date_col = find_date_col(cols) stat = {'rows': len(df), 'cols': cols, 'amt_col': amt_col, 'date_col': date_col} if amt_col: ser = pd.to_numeric(df[amt_col].astype(str).str.replace(',','').str.replace('￥','').str.replace('¥',''), errors='coerce') stat['total'] = ser.sum(skipna=True) stat['mean'] = ser[ser>0].mean() if ser.notna().any() else None else: stat['total'] = None stat['mean'] = None if date_col and stat.get('total'): try: dates = pd.to_datetime(df[date_col], errors='coerce') df['_month'] = dates.dt.to_period('M') monthly = df.groupby('_month')[amt_col].apply(lambda s: pd.to_numeric(s.astype(str).str.replace(',','').str.replace('￥','').str.replace('¥',''), errors='coerce').sum()) stat['monthly'] = monthly.to_dict() except Exception: stat['monthly'] = None else: stat['monthly'] = None sheet_stats[name] = stat

best_mean = None best_sheet = None for s, st in sheet_stats.items(): if st.get('mean') is not None: if best_mean is None or (st.get('total') or 0) > (sheet_stats.get(best_sheet,{}).get('total') or 0): best_mean = st['mean'] best_sheet = s

if best_mean is None: best_mean = 5.0 detect_note = '未检测到金额列，使用默认平均合同额 5 万/单。' else: detect_note = f'使用 sheet "{best_sheet}" 的平均合同额 {best_mean:.2f} 万/单。'

A = best_mean R = R_TOTAL S_needed = R / A if A and A>0 else None L_effective = S_needed / SIGN_RATE if S_needed and SIGN_RATE>0 else None

months_labels = [] today = pd.Timestamp.today() for i in range(1,7): months_labels.append((today + pd.offsets.MonthBegin(i)).strftime('%Y-%m'))

plan_even = [R/6.0]*6

import json with pd.ExcelWriter(OUTPUT_XLSX, engine='openpyxl') as writer: for name, df in all_sheets.items(): try: df.to_excel(writer, sheet_name=name[:31], index=False) except Exception: df.head(PREVIEW_ROWS).to_excel(writer, sheet_name=(name[:27] + '_prv'), index=False) pd.DataFrame({'month': months_labels, 'target_万_even': plan_even}).to_excel(writer, sheet_name='H2_按月_均分', index=False) pd.DataFrame([{'指标':'R_总目标(万)','数值':R},{'指标':'平均合同额_A(万/单)','数值':A},{'指标':'需签单数_S(单)','数值':S_needed},{'指标':'签单率(%)','数值':SIGN_RATE*100},{'指标':'需有效商机_L(个)','数值':L_effective}]).to_excel(writer, sheet_name='下半年任务规划', index=False)

with open(REPORT_TXT, 'w', encoding='utf-8') as f: f.write('自动解析报告\n') f.write(detect_note + '\n') for s, st in sheet_stats.items(): f.write(f"Sheet: {s} rows={st['rows']} amt_col={st['amt_col']} total={(st['total'] or 0)}\n") f.write(f"\nA={A} R={R} S_needed={S_needed} L_effective={L_effective}\n")

print('Wrote outputs:', OUTPUT_XLSX, REPORT_TXT)
