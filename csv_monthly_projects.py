# csv_projects_to_excel.py
import os
import re
import pandas as pd
from openpyxl import Workbook

ROOT = os.getcwd()  # run the script from C:\Users\Dell\clockify_reports
OUTPUT = os.path.join(ROOT, "clockify_all_projects.xlsx")

# map month names to numbers for sorting
MONTH_MAP = {
    'january':1,'february':2,'march':3,'april':4,'may':5,'june':6,
    'july':7,'august':8,'september':9,'october':10,'november':11,'december':12
}

def normalize_col(c):
    return re.sub(r'[\s\(\)\-\.]', '', str(c).strip().lower())

def find_duration_column(cols):
    """Find the column that likely contains duration/time."""
    norm_to_orig = {normalize_col(c): c for c in cols}
    for candidate in ("durationdecimal", "durationh", "duration"):
        if candidate in norm_to_orig:
            return norm_to_orig[candidate]
    for n, orig in norm_to_orig.items():
        if 'duration' in n or 'time' in n or 'hours' in n:
            return orig
    return None

def parse_duration_value(v):
    """Return decimal hours (float) from various input formats."""
    if pd.isna(v):
        return 0.0
    try:
        if isinstance(v, (int, float)):
            return float(v)
        s = str(v).strip().replace(',', '.')
        if re.fullmatch(r'[+-]?\d+(\.\d+)?', s):
            return float(s)
        if ':' in s:
            parts = s.split(':')
            nums = []
            for p in parts[:3]:
                p = re.sub(r'[^\d]', '', p.strip())
                nums.append(int(p) if p else 0)
            if len(nums) == 2:
                h, m = nums
                sec = 0
            else:
                h, m, sec = nums
            return h + m/60.0 + sec/3600.0
        return 0.0
    except Exception:
        return 0.0

def month_sort_key(name):
    """Return a sortable key for month labels."""
    if not name:
        return (9999, name)
    n = name.strip()
    m = re.match(r'(\d{4})[-_](\d{2})$', n)
    if m:
        y = int(m.group(1)); mo = int(m.group(2))
        return (y*100 + mo, n)
    parts = re.split(r'[_\-\s]', n.lower())
    for p in parts:
        if p in MONTH_MAP:
            year = None
            for q in parts:
                if re.fullmatch(r'\d{4}', q):
                    year = int(q)
                    break
            if year:
                return (year*100 + MONTH_MAP[p], n)
            return (2000 + MONTH_MAP[p], n)
    return (9999, n)

def process_project_folder(project_folder):
    """Return a dict month_label -> dict(total, entries[]) for this project."""
    path = os.path.join(ROOT, project_folder)
    if not os.path.isdir(path):
        return None
    month_totals = {}

    for fname in sorted(os.listdir(path)):
        if not fname.lower().endswith('.csv'):
            continue
        fpath = os.path.join(path, fname)
        month_label = os.path.splitext(fname)[0]
        try:
            df = pd.read_csv(fpath)
        except Exception as e:
            print(f"   ❌ Could not read {fname}: {e}")
            continue
        if df.shape[0] == 0:
            print(f"   ⚠ {fname}: empty")
            continue

        # find duration column
        duration_col = find_duration_column(df.columns)
        if duration_col is None:
            print(f"   ⚠ {fname}: no duration column found (columns: {list(df.columns)})")
            continue

        # find description column
        desc_col = None
        for c in df.columns:
            n = normalize_col(c)
            if 'description' in n or 'details' in n or 'note' in n:
                desc_col = c
                break

        # parse durations and collect entries
        hours_series = df[duration_col].apply(parse_duration_value)
        total_hours = round(float(hours_series.sum()), 2)
        entries = []
        for _, row in df.iterrows():
            desc = row[desc_col] if desc_col and desc_col in row else ''
            hrs = parse_duration_value(row[duration_col])
            if hrs > 0:
                entries.append((desc, hrs))

        month_totals[month_label] = {
            'total': total_hours,
            'entries': entries
        }

        print(f"   ✓ {fname} -> {total_hours} hrs (using '{duration_col}')")

    return month_totals

def main():
    print(f"Scanning root folder: {ROOT}")
    wb = Workbook()
    default_sheet = wb.active
    added = 0

    for item in sorted(os.listdir(ROOT)):
        folder = os.path.join(ROOT, item)
        if not os.path.isdir(folder):
            continue
        if item.lower() in ('venv', '__pycache__') or item.startswith('.'):
            continue

        print(f"\nProcessing project folder: {item}")
        month_totals = process_project_folder(item)
        if not month_totals:
            print(f"  (no valid CSV months found for {item})")
            continue

        ws = wb.create_sheet(title=item[:31])
        ws.append(["Month", "Total Hours"])

        for mn in sorted(month_totals.keys(), key=month_sort_key):
            ws.append([mn, month_totals[mn]['total']])
            ws.append(["Description", "Hours"])
            for desc, hrs in month_totals[mn]['entries']:
                ws.append([desc, hrs])
            ws.append([])

        grand = round(sum(m['total'] for m in month_totals.values()), 2)
        ws.append([])
        ws.append(["TOTAL", grand])
        added += 1

    if added > 0:
        try:
            wb.remove(default_sheet)
        except Exception:
            pass
    else:
        wb.active.title = "NoData"

    wb.save(OUTPUT)
    print(f"\n✅ Workbook saved as: {OUTPUT}")

if __name__ == "__main__":
    main()
