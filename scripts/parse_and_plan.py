#!/usr/bin/env python3
"""Parse encrypted Excel and generate monthly analysis report."""
import sys, os, io, json
from datetime import datetime
from collections import defaultdict
import msoffcrypto, openpyxl


def load_workbook(excel_path, password=None):
    with open(excel_path, 'rb') as f:
        office_file = msoffcrypto.OfficeFile(f)
        if office_file.is_encrypted():
            if not password:
                raise ValueError("File encrypted, set EXCEL_PASSWORD env var")
            office_file.load_key(password=password)
            decrypted = io.BytesIO()
            office_file.decrypt(decrypted)
            return openpyxl.load_workbook(decrypted, data_only=True)
        return openpyxl.load_workbook(excel_path, data_only=True)


def get_amount(row, ac=18, pc=16, tc=17):
    amt = row[ac]
    if amt is not None:
        try:
            v = float(amt)
            if v > 0: return v
        except: pass
    try:
        p, t = float(row[pc]), float(row[tc])
        if p > 0 and t > 0: return p * t
    except: pass
    return None


def parse_excel(excel_path, password=None):
    print(f"Loading: {excel_path}")
    wb = load_workbook(excel_path, password)
    sheet = next((n for n in wb.sheetnames if '年' in n), wb.sheetnames[0])
    ws = wb[sheet]
    print(f"Sheet: {sheet} ({ws.max_row} rows)")
    monthly, skip, refund = defaultdict(float), 0, 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[24] is not None:
            refund += 1; continue
        if not row[0] or not isinstance(row[0], datetime):
            skip += 1; continue
        amt = get_amount(row)
        if amt is None or amt <= 0:
            skip += 1; continue
        monthly[(row[0].year, row[0].month)] += amt
    print(f"Valid: {len(monthly)} months | Skip: {skip} | Refund: {refund}")
    return monthly


def generate_report(monthly, od="outputs"):
    os.makedirs(od, exist_ok=True)
    years = sorted(set(k[0] for k in monthly))
    hm = list(range(7, 13))
    r = ["# 签单金额月度分析报告\n",
         f"> 生成时间: {datetime.now():%Y-%m-%d %H:%M}\n",
         "\n## 各年月度签单金额\n"]
    for y in years:
        t = sum(v for k, v in monthly.items() if k[0] == y)
        if not t: continue
        r.append(f"\n### {y}年 合计{t/10000:.1f}万\n\n|月份|金额|占比|\n|---|---|---|\n")
        for m in range(1, 13):
            v = monthly.get((y, m), 0)
            r.append(f"|{m}月|{v/10000:.1f}万|{v/t*100:.1f}%|\n")
    r.append("\n## H2 (7-12月) 历史分析\n")
    hp = defaultdict(list)
    for y in years:
        ht = sum(monthly.get((y, m), 0) for m in hm)
        if not ht: continue
        r.append(f"\n### {y}年 H2合计{ht/10000:.1f}万\n\n|月份|金额|占比|\n|---|---|---|\n")
        for m in hm:
            v = monthly.get((y, m), 0)
            r.append(f"|{m}月|{v/10000:.1f}万|{v/ht*100:.1f}%|\n")
            hp[m].append(v / ht * 100)
    r.append("\n## H2月度平均占比\n\n|月份|占比|\n|---|---|\n")
    ap = {}
    for m in hm:
        ap[m] = sum(hp[m]) / len(hp[m]) if hp[m] else 0
        r.append(f"|{m}月|{ap[m]:.1f}%|\n")
    rp = f"{od}/h2_monthly_report.md"
    with open(rp, "w", encoding="utf-8") as f:
        f.write("".join(r))
    jd = {"generated_at": datetime.now().isoformat(),
          "monthly_sign": {f"{k[0]}-{k[1]:02d}": round(v, 2) for k, v in monthly.items()},
          "h2_proportions": {str(m): round(v, 2) for m, v in ap.items()}}
    jp = f"{od}/h2_monthly_data.json"
    with open(jp, "w", encoding="utf-8") as f:
        json.dump(jd, f, ensure_ascii=False, indent=2)
    print(f"\n✅ {rp}\n✅ {jp}")


def main():
    if len(sys.argv) < 2:
        print("Usage: python scripts/parse_and_plan.py <excel_file>")
        sys.exit(1)
    pw = os.environ.get("EXCEL_PASSWORD", "")
    if not os.path.exists(sys.argv[1]):
        print(f"Error: File not found: {sys.argv[1]}"); sys.exit(1)
    try:
        generate_report(parse_excel(sys.argv[1], pw))
    except Exception as e:
        print(f"\n❌ {e}"); sys.exit(1)


if __name__ == "__main__":
    main()
