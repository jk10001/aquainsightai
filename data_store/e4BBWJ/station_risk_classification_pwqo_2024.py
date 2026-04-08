# filename: station_risk_classification_pwqo_2024.py
import pandas as pd
import numpy as np
import re
from datetime import datetime
import math
import os

# --- Settings / filenames
xlsx_file = "Ontario_PWQMN_2024.xlsx"
out_base = "station_risk_classification_pwqo_2024"
csv_out = out_base + ".csv"
xlsx_out = out_base + ".xlsx"
py_file = os.path.basename(__file__)

# --- Helper functions
def parse_date_mixed(s):
    # Parse mixed date formats: 'MM/DD/YYYY' and 'YYYY-MM-DD 00:00:00'
    if pd.isna(s):
        return pd.NaT
    if isinstance(s, (pd.Timestamp, datetime)):
        return pd.to_datetime(s).date()
    s = str(s).strip()
    # Try direct pandas parse
    try:
        dt = pd.to_datetime(s, errors='coerce', dayfirst=False)
        if pd.isna(dt):
            return pd.NaT
        return dt.date()
    except Exception:
        return pd.NaT

def parse_result_value(result_str, detection_limit):
    # Returns numeric value (float) and a boolean is_nd
    if pd.isna(result_str):
        return (np.nan, False)
    rs = str(result_str).strip()
    if rs.startswith("<"):
        # non-detect
        # try to parse number after '<'
        m = re.search(r"<\s*([0-9.+eE-]+)", rs)
        if m:
            try:
                return (float(m.group(1)), True)
            except:
                pass
        # fallback to provided detection limit
        if not pd.isna(detection_limit):
            return (float(detection_limit), True)
        # if nothing else, NaN but mark as ND
        return (np.nan, True)
    else:
        # try numeric parse
        try:
            return (float(rs), False)
        except:
            # sometimes results may include non-numeric; return NaN
            return (np.nan, False)

def percent_or_blank(n_exceed, n_total):
    if n_total == 0 or pd.isna(n_total):
        return ""
    return 100.0 * n_exceed / n_total

def assign_tier(pct):
    # pct is percent (0-100)
    if pd.isna(pct):
        return "Tier 1"
    if pct == 0:
        return "Tier 1"
    if 0 < pct <= 15:
        return "Tier 2"
    if 15 < pct < 50:
        return "Tier 3"
    if pct >= 50:
        return "Tier 4"
    return "Tier 1"

# Hardness-adjusted objectives
def cadmium_objective(hardness):
    # hardness in mg/L as CaCO3
    # 0.1 µg/L if hardness <100; else 0.5
    if pd.isna(hardness):
        return None
    try:
        if hardness < 100:
            return 0.1
        else:
            return 0.5
    except:
        return None

def copper_objective(hardness):
    # 1 µg/L if hardness 0–20; else 5 µg/L
    if pd.isna(hardness):
        return None
    try:
        if hardness <= 20:
            return 1.0
        else:
            return 5.0
    except:
        return None

def lead_objective(hardness):
    # 1 µg/L if hardness <30; 3 µg/L if 30–80; 5 µg/L if >80
    if pd.isna(hardness):
        return None
    try:
        if hardness < 30:
            return 1.0
        elif 30 <= hardness <= 80:
            return 3.0
        else:
            return 5.0
    except:
        return None

# --- Read data
xls = pd.ExcelFile(xlsx_file)
df = pd.read_excel(xls, sheet_name="Data", dtype={"Collection Site": object})
stations = pd.read_excel(xls, sheet_name="Stations", dtype={"STATION": object})

# Standardize station id fields as strings (to preserve leading zeros if any)
df['Collection Site'] = df['Collection Site'].astype(str).str.strip()
stations['STATION'] = stations['STATION'].astype(str).str.strip()

# Parse collected date to date object
df['collected_date'] = df['Collected'].apply(parse_date_mixed)

# Convert Detection Limit to numeric where possible
df['Detection Limit numeric'] = pd.to_numeric(df['Detection Limit'], errors='coerce')

# Parse Results to numeric and ND flag
parsed = df.apply(lambda row: parse_result_value(row['Results'], row['Detection Limit numeric']), axis=1)
df['result_numeric'] = parsed.apply(lambda x: x[0])
df['is_nd'] = parsed.apply(lambda x: x[1])
# is_detected from Result Call column (Detected vs BDL)
df['is_detected'] = df['Result Call'].astype(str).str.strip().str.lower() == 'detected'

# For safety, standardize Analyte names
df['Analyte'] = df['Analyte'].astype(str).str.strip()

# Extract hardness and pH same-day by station+date (aggregate mean if multiple)
hardness_df = df[df['Analyte'].str.lower().str.strip() == 'hardness'].copy()
hardness_df = hardness_df.dropna(subset=['collected_date'])
hardness_agg = hardness_df.groupby(['Collection Site', 'collected_date']).agg(
    hardness_mean=('result_numeric', 'mean')
).reset_index()

ph_df = df[df['Analyte'].str.lower().str.strip() == 'pH'.lower()].copy()
# Note: 'pH' likely exactly 'pH', ensure matching
ph_df = df[df['Analyte'].str.strip().str.lower() == 'pH'.lower()].copy()
ph_df = ph_df.dropna(subset=['collected_date'])
ph_agg = ph_df.groupby(['Collection Site', 'collected_date']).agg(
    pH_mean=('result_numeric', 'mean')
).reset_index()

# Prepare lookup dictionaries for same-day hardness and pH
# Use tuple (site,date) as key, date is datetime.date
hardness_lookup = {(row['Collection Site'], row['collected_date']): row['hardness_mean'] for idx,row in hardness_agg.iterrows()}
ph_lookup = {(row['Collection Site'], row['collected_date']): row['pH_mean'] for idx,row in ph_agg.iterrows()}

# List of PWQO analytes and mapping to dataset names
analytes_info = [
    ("Aluminum", "Aluminum", "ug/L"),
    ("Arsenic", "Arsenic", "ug/L"),
    ("Boron", "Boron", "ug/L"),
    ("Cadmium", "Cadmium", "ug/L"),
    ("Chromium", "Chromium", "ug/L"),
    ("Copper", "Copper", "ug/L"),
    ("E. coli", "E. coli count per 100 mL", "MPN / 100mL"),
    ("Iron", "Iron", "ug/L"),
    ("Lead", "Lead", "ug/L"),
    ("Mercury", "Mercury", "ng/L"),
    ("Nickel", "Nickel", "ug/L"),
    ("Total Phosphorus", "Phosphorus; total", "ug/L"),
    ("Selenium", "Selenium", "ug/L"),
    ("Silver", "Silver", "ug/L"),
    ("Zinc", "Zinc", "ug/L"),
    ("pH", "pH", "SU")
]

# Objectives (in analyte units or indicated). For hardness-adjusted/pH-dependent define special handling.
objectives = {
    "Aluminum": None,  # pH-dependent
    "Arsenic": 5.0,
    "Boron": 200.0,
    "Cadmium": None,  # hardness-adjusted
    "Chromium": 1.0,
    "Copper": None,   # hardness-adjusted
    "E. coli": 100.0,
    "Iron": 300.0,
    "Lead": None,     # hardness-adjusted
    # Mercury: objective 0.2 µg/L = 200 ng/L
    "Mercury": 200.0,  # we'll compare in ng/L
    "Nickel": 25.0,
    "Total Phosphorus": 30.0,
    "Selenium": 100.0,
    "Silver": 0.1,
    "Zinc": 20.0,
    "pH": (6.5, 8.5)
}

# We'll prepare a per-sample boolean whether that sample is "usable" for the analyte (after exclusions)
# and whether it is an exceedance (Result Call == Detected and numeric result > objective)

# Add date key for lookups
df['date_key'] = list(zip(df['Collection Site'], df['collected_date']))

# Function to evaluate single sample for a given analyte row
def evaluate_sample(row, analyte_display, analyte_dataset_name):
    # Returns (usable_flag, is_exceed_flag)
    a_name = analyte_dataset_name
    if row['Analyte'] != a_name:
        return (False, False)
    # Must have a collected date
    if pd.isna(row['collected_date']):
        return (False, False)
    # For non-detect samples: usable but cannot be exceedance per instructions.
    # However some analytes require same-day pH/hardness; those samples may be excluded entirely if no same-day measurement.
    # pH-dependent Aluminum
    if analyte_display == "Aluminum":
        # require same-day pH
        pkey = (row['Collection Site'], row['collected_date'])
        pval = ph_lookup.get(pkey, np.nan)
        if pd.isna(pval):
            return (False, False)  # exclude from both N and exceed
        # Only screen when pH >6.5 or pH <5.5; exclude if 5.5 <= pH <= 6.5
        try:
            p_val = float(pval)
        except:
            return (False, False)
        if 5.5 <= p_val <= 6.5:
            return (False, False)
        # Determine objective based on pH
        if p_val > 6.5:
            obj = 75.0
        elif p_val < 5.5:
            obj = 15.0
        else:
            return (False, False)
        # Now evaluate exceedance: must be Detected and numeric > obj
        if (row['is_detected']) and (not pd.isna(row['result_numeric'])) and (row['result_numeric'] > obj):
            return (True, True)
        else:
            # even if ND, it's counted as usable (per earlier note: Aluminum screening attempted only where same-day pH available)
            return (True, False)
    # Hardness-adjusted cadmium/copper/lead
    if analyte_display in ("Cadmium","Copper","Lead"):
        hkey = (row['Collection Site'], row['collected_date'])
        hardness = hardness_lookup.get(hkey, np.nan)
        if pd.isna(hardness):
            return (False, False)  # exclude sample if no same-day hardness
        # Determine objective
        if analyte_display == "Cadmium":
            obj = cadmium_objective(hardness)
        elif analyte_display == "Copper":
            obj = copper_objective(hardness)
        else:  # Lead
            obj = lead_objective(hardness)
        if obj is None:
            return (False, False)
        # Evaluate exceedance: must be Detected and numeric > obj
        if (row['is_detected']) and (not pd.isna(row['result_numeric'])) and (row['result_numeric'] > obj):
            return (True, True)
        else:
            return (True, False)
    # Mercury: ensure unit conversion (objective given in ng/L)
    if analyte_display == "Mercury":
        # data units might be 'ng/L' or 'ug/L'
        units = str(row.get('Units', '')).lower()
        # We'll convert measured numeric to ng/L if in ug/L multiply by 1000
        val = row['result_numeric']
        if pd.isna(val):
            # ND counts as usable sample (unless detection missing?), but per instruction ND is not an exceedance
            return (True, False)
        numeric_ng = val
        if 'ug' in units:
            numeric_ng = float(val) * 1000.0
        # Also if units contain 'ng' assume ng/L
        obj_ng = objectives['Mercury']  # already 200 ng/L
        if (row['is_detected']) and (not pd.isna(numeric_ng)) and (numeric_ng > obj_ng):
            return (True, True)
        else:
            return (True, False)
    # pH
    if analyte_display == "pH":
        # pH sample itself; usable if numeric; exceed if <6.5 or >8.5
        val = row['result_numeric']
        if pd.isna(val):
            return (True, False)
        try:
            v = float(val)
        except:
            return (True, False)
        if (v < objectives['pH'][0]) or (v > objectives['pH'][1]):
            return (True, True)
        else:
            return (True, False)
    # E. coli: simple numeric compare; ND is not exceeding
    if analyte_display == "E. coli":
        val = row['result_numeric']
        if pd.isna(val):
            return (True, False)
        try:
            v = float(val)
        except:
            return (True, False)
        if (row['is_detected']) and (v > objectives['E. coli']):
            return (True, True)
        else:
            return (True, False)
    # For remaining analytes with fixed objectives (units typically ug/L)
    if analyte_display in objectives and objectives[analyte_display] is not None:
        obj = objectives[analyte_display]
        val = row['result_numeric']
        if pd.isna(val):
            return (True, False)
        # Units handling: most are ug/L and objective is in ug/L except Silver objective 0.1 ug/L
        # If units are in different scale (e.g., ng/L), attempt conversion:
        units = str(row.get('Units', '')).lower()
        numeric_val_for_compare = val
        # If units contain 'ng' and objective in ug/L -> convert val ng->ug by /1000
        if 'ng' in units and obj > 1e-6:
            # obj in ug/L -> convert val ng/L to ug/L
            numeric_val_for_compare = float(val) / 1000.0
        # If units contains 'mg' convert to ug (mg/L to ug/L multiply 1e6) - unlikely but safe
        if 'mg' in units and 'mg/l' in units:
            numeric_val_for_compare = float(val) * 1e6
            # convert objective to ug/L if objective is in ug/L; but this is complex; assume data units align
        if (row['is_detected']) and (not pd.isna(numeric_val_for_compare)) and (numeric_val_for_compare > obj):
            return (True, True)
        else:
            return (True, False)
    # If analyte objective not specified and not special-case, exclude
    return (False, False)

# For performance, filter dataset to only relevant analytes + hardness + pH
relevant_dataset_analytes = set([info[1] for info in analytes_info] + ['Hardness','hardness','pH','Hardness '])
df_relevant = df[df['Analyte'].isin([info[1] for info in analytes_info] + ['Hardness'])].copy()

# Create per-station, per-analyte aggregation
# We'll iterate stations present in df_relevant (only stations with >=1 sample for any PWQO analyte)
stations_with_samples = df_relevant['Collection Site'].unique().tolist()

# Create a mapping from station id to station name using Stations sheet
station_name_map = stations.set_index('STATION')['NAME'].to_dict()

# Build results rows
rows = []
for station in stations_with_samples:
    station_rows = df_relevant[df_relevant['Collection Site'] == station]
    # We'll compute per-analyte Ns and exceeds
    per_analyte_N = {}
    per_analyte_exceed = {}
    total_N = 0
    total_exceed = 0
    for display_name, dataset_name, unit in analytes_info:
        per_analyte_N[display_name] = 0
        per_analyte_exceed[display_name] = 0
        # iterate over station_rows where Analyte==dataset_name
        subset = station_rows[station_rows['Analyte'] == dataset_name]
        if subset.empty:
            continue
        # For each sample, evaluate
        for idx, sample in subset.iterrows():
            usable, exceed = evaluate_sample(sample, display_name, dataset_name)
            if usable:
                per_analyte_N[display_name] += 1
                total_N += 1
                if exceed:
                    per_analyte_exceed[display_name] += 1
                    total_exceed += 1
    # Only include station rows that had at least one usable sample across PWQO analytes
    if total_N == 0:
        continue
    # Compute % total exceedances
    pct_total_exceed = 100.0 * total_exceed / total_N if total_N > 0 else 0.0
    tier = assign_tier(pct_total_exceed)
    # Build row
    station_name = station_name_map.get(station, "")
    row = {
        "Station name": station_name,
        "Station number": station,
        "Total samples (PWQO analytes)": total_N,
        "% total exceedances (PWQO analytes)": round(pct_total_exceed, 2),
        "Risk tier": tier
    }
    # Add analyte-specific columns: N <analyte> and % exceed <analyte>
    for display_name, dataset_name, unit in analytes_info:
        n = per_analyte_N.get(display_name, 0)
        ne = per_analyte_exceed.get(display_name, 0)
        pct = percent_or_blank(ne, n)
        # Round percent to 2 decimals if present
        if pct != "":
            pct = round(pct, 2)
        # Column naming with units where appropriate
        col_n = f"N {display_name} [{unit}]"
        col_pct = f"% exceed {display_name}"
        row[col_n] = n
        row[col_pct] = pct
    rows.append(row)

# Create DataFrame
result_df = pd.DataFrame(rows)

# Ensure column order: Station name, Station number, Total samples, % total exceedances, Risk tier, then per-analyte columns in given order
base_cols = ["Station name", "Station number", "Total samples (PWQO analytes)", "% total exceedances (PWQO analytes)", "Risk tier"]
analyte_cols = []
for display_name, _, unit in analytes_info:
    analyte_cols.append(f"N {display_name} [{unit}]")
    analyte_cols.append(f"% exceed {display_name}")
cols_order = base_cols + analyte_cols
result_df = result_df[cols_order]

# Sorting: by Risk tier (Tier 4 → Tier 1), then by % total exceed descending, then by Total samples descending.
# Map tier to order key
tier_order = {"Tier 4": 4, "Tier 3": 3, "Tier 2": 2, "Tier 1": 1}
result_df['tier_sort'] = result_df['Risk tier'].map(tier_order).fillna(0).astype(int)
result_df['pct_total_exceed_sort'] = pd.to_numeric(result_df["% total exceedances (PWQO analytes)"], errors='coerce').fillna(0)
result_df['total_samples_sort'] = pd.to_numeric(result_df["Total samples (PWQO analytes)"], errors='coerce').fillna(0)
result_df = result_df.sort_values(by=['tier_sort','pct_total_exceed_sort','total_samples_sort'], ascending=[False, False, False])
result_df = result_df.drop(columns=['tier_sort','pct_total_exceed_sort','total_samples_sort'])

# For % columns, ensure blank when N==0 (we stored blank already via percent_or_blank)
# But ensure numeric zeros are kept as blank for analyte % if N==0
for display_name, _, unit in analytes_info:
    col_n = f"N {display_name} [{unit}]"
    col_pct = f"% exceed {display_name}"
    # If N==0 and pct is not blank, set to blank
    mask = result_df[col_n] == 0
    result_df.loc[mask, col_pct] = ""

# Reorder Risk tiers to text exactly "Tier X" already set
# Ensure columns with numeric types are appropriate
# Write CSV (final processed table only)
result_df.to_csv(csv_out, index=False)

# Write Excel with frozen header and auto column widths
try:
    with pd.ExcelWriter(xlsx_out, engine='xlsxwriter') as writer:
        result_df.to_excel(writer, sheet_name='Station Risk', index=False)
        workbook = writer.book
        worksheet = writer.sheets['Station Risk']
        # Freeze header row
        worksheet.freeze_panes(1, 0)
        # Autofit columns: compute max length per column
        for i, col in enumerate(result_df.columns):
            # Determine max width between column name and content
            series = result_df[col].astype(str).fillna("")
            max_len = max(series.map(len).max(), len(col)) + 2
            # Limit max width to reasonable value
            max_len = min(max_len, 60)
            worksheet.set_column(i, i, max_len)
except Exception as e:
    # fallback: write with pandas default engine
    result_df.to_excel(xlsx_out, index=False)

# Print created file names and python filename as required
print(py_file)
print(csv_out)
print(xlsx_out)