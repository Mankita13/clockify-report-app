# app.py
import os
import re
import io
import shutil
import tempfile
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from datetime import datetime

st.set_page_config(page_title="Clockify → Excel (folder mode)", layout="centered")
st.title("🕒 Clockify Projects → Excel (Folder mode)")

# -------------------------
# Helpers (kept from your original script)
# -------------------------
MONTH_MAP = {
    'january':1,'february':2,'march':3,'april':4,'may':5,'june':6,
    'july':7,'august':8,'september':9,'october':10,'november':11,'december':12
}

def normalize_col(c):
    return re.sub(r'[\s\(\)\-\.]', '', str(c).strip().lower())

def find_duration_column(cols):
    norm_to_orig = {normalize_col(c): c for c in cols}
    for candidate in ("durationdecimal", "durationh", "duration"):
        if candidate in norm_to_orig:
            return norm_to_orig[candidate]
    for n, orig in norm_to_orig.items():
        if 'duration' in n or 'time' in n or 'hours' in n:
            return orig
    return None

def parse_duration_value(v):
    if pd.isna(v):
        return 0.0
    try:
        if isinstance(v, (int, float)):
            return float(v)
        s = str(v).strip().replace(',', '.')
        # pure decimal number
        if re.fullmatch(r'[+-]?\d+(\.\d+)?', s):
            return float(s)
        # hh:mm[:ss] style
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

# -------------------------
# Processing (keeps your original per-project/month logic)
# -------------------------
def process_project_folder(root_path, project_folder, logs):
    path = os.path.join(root_path, project_folder)
    if not os.path.isdir(path):
        return None
    month_totals = {}
    for fname in sorted(os.listdir(path)):
        if not fname.lower().endswith('.csv'):
            continue
        fpath = os.path.join(path, fname)
        month_label = os.path.splitext(fname)[0]
        try:
            try:
                df = pd.read_csv(fpath)
            except Exception:
                df = pd.read_csv(fpath, encoding='latin1')
        except Exception as e:
            logs.append(f"   ❌ Could not read {project_folder}/{fname}: {e}")
            continue
        if df.shape[0] == 0:
            logs.append(f"   ⚠ {project_folder}/{fname}: empty")
            continue

        duration_col = find_duration_column(df.columns)
        if duration_col is None:
            logs.append(f"   ⚠ {project_folder}/{fname}: no duration column found (columns: {list(df.columns)})")
            continue

        desc_col = None
        for c in df.columns:
            n = normalize_col(c)
            if 'description' in n or 'details' in n or 'note' in n:
                desc_col = c
                break

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

        logs.append(f"   ✓ {project_folder}/{fname} -> {total_hours} hrs (using '{duration_col}')")

    return month_totals

def generate_workbook_bytes(root):
    logs = []
    wb = Workbook()
    default_sheet = wb.active
    added = 0
    project_summaries = []

    for item in sorted(os.listdir(root)):
        folder = os.path.join(root, item)
        if not os.path.isdir(folder):
            continue
        if item.lower() in ('venv', '__pycache__') or item.startswith('.'):
            continue

        logs.append(f"\nProcessing project folder: {item}")
        month_totals = process_project_folder(root, item, logs)
        if not month_totals:
            logs.append(f"  (no valid CSV months found for {item})")
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
        project_summaries.append({'project': item, 'total_hours': grand})
        added += 1

    if added > 0:
        try:
            wb.remove(default_sheet)
        except Exception:
            pass
    else:
        wb.active.title = "NoData"

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"clockify_all_projects_{timestamp}.xlsx"
    return bio.getvalue(), filename, project_summaries, logs

# -------------------------
# Streamlit UI (folder-only mode)
# -------------------------
st.markdown("**Mode:** Read directly from a local or network folder (no upload).")

root = st.text_input("Root folder path containing project folders (e.g., C:\\Users\\Dell\\clockify_reports)", value=os.getcwd())
save_into_folder = st.checkbox("Save generated Excel into the selected folder", value=True)

if st.button("Generate report"):
    if not os.path.isdir(root):
        st.error("❌ Invalid folder path. Paste the full path to the folder that contains project subfolders.")
    else:
        with st.spinner("Processing CSV files..."):
            bytes_xlsx, filename, summaries, logs = generate_workbook_bytes(root)
        st.success("✅ Report generated")
        st.download_button("⬇️ Download Excel", data=bytes_xlsx, file_name=filename,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        if save_into_folder:
            outpath = os.path.join(root, filename)
            try:
                with open(outpath, "wb") as f:
                    f.write(bytes_xlsx)
                st.info(f"Saved copy to: {outpath}")
            except Exception as e:
                st.warning(f"Could not save into folder: {e}")

        if summaries:
            st.write("Project totals:")
            st.table(pd.DataFrame(summaries).sort_values(by='project').reset_index(drop=True))

        with st.expander("Processing log"):
            for line in logs:
                st.text(line)
