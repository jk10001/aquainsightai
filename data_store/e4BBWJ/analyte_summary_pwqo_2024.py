# filename: analyte_summary_pwqo_2024.py
import pandas as pd
import numpy as np
import re
from datetime import datetime

# --- Config / filenames ---
excel_file = "Ontario_PWQMN_2024.xlsx"
sheet_data = "Data"
base_name = "analyte_summary_pwqo_2024"
csv_out = f"{base_name}.csv"
py_name = f"{base_name}.py"

# --- Helper functions ---
def parse_numeric_from_string(s):
    if pd.isna(s):
        return np.nan
    s = str(s).strip()
    # If begins with '<' or '>' capture number
    m = re.search(r"[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?", s)
    if m:
        try:
            return float(m.group())
        except:
            return np.nan
    return np.nan

def common_unit(series):
    vals = series.dropna().astype(str).str.strip()
    if len(vals) == 0:
        return ""
    # Return the modal (most common) unit
    try:
        return vals.mode().iat[0]
    except:
        return vals.iloc[0]

def to_ugL(value, unit):
    """
    Convert a numeric value in 'unit' to ug/L.
    Supports: ug/L, ng/L, mg/L (including 'mg/L as CaCO3'), and returns np.nan for non-mass conc units.
    """
    if pd.isna(value):
        return np.nan
    if unit is None:
        return np.nan
    u = str(unit).lower().strip()
    try:
        # Normalize common spellings
        if u in ("ug/l", "µg/l", "ug/l)", "µg/l)"):
            return value
        if u in ("ng/l", "ng/l)", "ng/l " , "ng/l)"):
            # ng/L -> ug/L: divide by 1000
            return value / 1000.0
        if "mg/l" in u:
            # mg/L -> ug/L
            return value * 1000.0
        # if unit contains 'mpn' or 'uS/cm' etc -> not convertible
        return np.nan
    except:
        return np.nan

def to_native(value_ugL, native_unit):
    """
    Convert ug/L value back to native unit for reporting when native unit is mass concentration.
    Supports ug/L, mg/L, ng/L.
    """
    if pd.isna(value_ugL):
        return np.nan
    if native_unit is None:
        return np.nan
    u = str(native_unit).lower().strip()
    try:
        if u in ("ug/l", "µg/l"):
            return value_ugL
        if "mg/l" in u:
            return value_ugL / 1000.0
        if "ng/l" in u:
            return value_ugL * 1000.0
        return np.nan
    except:
        return np.nan

# --- PWQO mapping and rules ---
pwqo_map = {
    "pH": {
        "objective": (6.5, 8.5),
        "objective_unit": "SU",
        "note": "Range 6.5-8.5 (exceed if <6.5 or >8.5)."
    },
    "E. coli count per 100 mL": {
        "objective": 100.0,
        "objective_unit": "MPN/100mL",
        "note": "Objective = 100 MPN/100 mL; geometric mean of >=5 samples required for official assessment; per-sample >100 flagged."
    },
    "Phosphorus; total": {
        "objective": 30.0,
        "objective_unit": "ug/L",
        "note": "River objective only (30 µg/L). Dataset contains rivers only per instructions."
    },
    "Aluminum": {
        "objective": None,
        "objective_unit": "ug/L",
        "note": "pH-dependent: use 75 µg/L for pH >6.5–9.0; 15 µg/L for pH 4.5–5.5. For pH 5.5–6.5: background-based / 10% increase; flagged as not directly screenable with single number.",
        "pH_dependent": True
    },
    "Arsenic": {
        "objective": 5.0,
        "objective_unit": "ug/L",
        "note": "Interim revised PWQO 5 µg/L (more stringent)."
    },
    "Cadmium": {
        "objective": None,
        "objective_unit": "ug/L",
        "note": "Hardness-adjusted (Interim revised): 0.1 µg/L if hardness <100 mg/L as CaCO3; else 0.5 µg/L.",
        "hardness_adjusted": True
    },
    "Chromium": {
        "objective": 1.0,
        "objective_unit": "ug/L",
        "note": "Speciation not provided in dataset; conservative screening = Cr VI objective 1 µg/L."
    },
    "Copper": {
        "objective": None,
        "objective_unit": "ug/L",
        "note": "Hardness-adjusted (Interim revised): 1 µg/L if hardness 0–20 mg/L as CaCO3; else 5 µg/L.",
        "hardness_adjusted": True
    },
    "Iron": {
        "objective": 300.0,
        "objective_unit": "ug/L",
        "note": ""
    },
    "Lead": {
        "objective": None,
        "objective_unit": "ug/L",
        "note": "Hardness-adjusted (Interim revised): 1 µg/L if hardness <30; 3 µg/L if 30–80; 5 µg/L if >80 mg/L as CaCO3.",
        "hardness_adjusted": True
    },
    "Mercury": {
        "objective": 0.2,
        "objective_unit": "ug/L",
        "note": "Objective 0.2 µg/L; applies to filtered/dissolved samples — dataset may not indicate filter status."
    },
    "Nickel": {
        "objective": 25.0,
        "objective_unit": "ug/L",
        "note": ""
    },
    "Selenium": {
        "objective": 100.0,
        "objective_unit": "ug/L",
        "note": ""
    },
    "Silver": {
        "objective": 0.1,
        "objective_unit": "ug/L",
        "note": ""
    },
    "Zinc": {
        "objective": 20.0,
        "objective_unit": "ug/L",
        "note": "Interim revised = 20 µg/L (more stringent than PWQO 30 µg/L)."
    },
    "Boron": {
        "objective": 200.0,
        "objective_unit": "ug/L",
        "note": "Interim 200 µg/L."
    }
}

# --- Read data ---
df = pd.read_excel(excel_file, sheet_name=sheet_data, engine="openpyxl", dtype=object)
df.columns = [c.strip() for c in df.columns]

# Parse dates robustly
df['Collected_dt'] = pd.to_datetime(df['Collected'], errors='coerce', infer_datetime_format=True)
mask_nat = df['Collected_dt'].isna()
if mask_nat.any():
    try:
        df.loc[mask_nat, 'Collected_dt'] = pd.to_datetime(df.loc[mask_nat, 'Collected'], format='%m/%d/%Y', errors='coerce')
    except:
        pass
df['Collected_date'] = df['Collected_dt'].dt.date

# Ensure Units column exists and normalized strings
if 'Units' not in df.columns:
    df['Units'] = np.nan
df['Units'] = df['Units'].astype(str).replace({'nan': np.nan})

# Numeric parsing of Results and Detection Limit
df['Result_str'] = df['Results'].astype(str).str.strip()
df['numeric_result'] = df['Result_str'].apply(parse_numeric_from_string)
if 'Detection Limit' in df.columns:
    df['MDL'] = pd.to_numeric(df['Detection Limit'], errors='coerce')
else:
    df['MDL'] = np.nan

# Identify non-detects: Result Call == 'BDL' or Value Qualifier contains '<' or Results startswith '<'
df['Value Qualifier'] = df.get('Value Qualifier', pd.Series([np.nan] * len(df))).astype(str)
df['Result Call'] = df.get('Result Call', pd.Series([np.nan] * len(df))).astype(str)
df['is_nd'] = False
df.loc[df['Result Call'].str.strip().str.upper() == 'BDL', 'is_nd'] = True
df.loc[df['Value Qualifier'].str.contains('<', na=False), 'is_nd'] = True
df.loc[df['Result_str'].str.startswith('<'), 'is_nd'] = True

# For ND, if MDL missing try to parse from Results string '<x'
mask_nd_mdl_missing = df['is_nd'] & df['MDL'].isna()
if mask_nd_mdl_missing.any():
    df.loc[mask_nd_mdl_missing, 'MDL'] = df.loc[mask_nd_mdl_missing, 'Result_str'].str.replace('<', '').apply(parse_numeric_from_string)

# Prepare analyte field
df['Analyte'] = df['Analyte'].astype(str).str.strip()

# --- Prepare hardness lookup: mean hardness by station and date ---
hardness_df = df[df['Analyte'].str.lower() == 'hardness'].copy()
hardness_df['hardness_native'] = hardness_df['numeric_result']
hardness_df['hardness_unit'] = hardness_df['Units'].astype(str)

def hardness_to_mgL(x, unit):
    if pd.isna(x):
        return np.nan
    u = str(unit).lower()
    if 'mg' in u:
        return x
    if 'ug' in u or 'µg' in u:
        return x / 1000.0
    # fallback: assume mg/L if ambiguous
    return x

hardness_df['hardness_mgL'] = hardness_df.apply(lambda row: hardness_to_mgL(row['hardness_native'], row['hardness_unit']), axis=1)
hard_by_site_date = hardness_df.groupby(['Collection Site', 'Collected_date'], dropna=False)['hardness_mgL'].mean().reset_index()
hard_by_site_date = hard_by_site_date.rename(columns={'hardness_mgL': 'mean_hardness_mgL'})

def compute_hardness_objective(analyte_name, hardness_mgL):
    if pd.isna(hardness_mgL):
        return np.nan
    if analyte_name == 'Cadmium':
        return 0.1 if hardness_mgL < 100.0 else 0.5
    if analyte_name == 'Copper':
        return 1.0 if hardness_mgL <= 20.0 else 5.0
    if analyte_name == 'Lead':
        if hardness_mgL < 30.0:
            return 1.0
        elif 30.0 <= hardness_mgL <= 80.0:
            return 3.0
        else:
            return 5.0
    return np.nan

# --- Identify analytes to include ---
present_analytes = sorted(df['Analyte'].unique())
target_analytes = [a for a in present_analytes if a in pwqo_map.keys()]

# Map variant E. coli names if present
if 'E. coli count per 100 mL' not in target_analytes:
    for a in present_analytes:
        if 'e. coli' in a.lower() or 'ecoli' in a.lower():
            if a not in target_analytes:
                # Map the dataset name to reference PWQO entry (keep original name in table)
                pwqo_map.setdefault(a, pwqo_map['E. coli count per 100 mL'])
                target_analytes.append(a)

# --- Build summary per analyte ---
summary_rows = []
for analyte in sorted(target_analytes):
    sub = df[df['Analyte'] == analyte].copy()
    N_samples = len(sub)
    if N_samples == 0:
        continue

    data_unit = common_unit(sub['Units'])
    sub['numeric_native'] = sub['numeric_result']
    sub['MDL_native'] = sub['MDL']

    # Build stat_value_native: detected numeric or 0.5*MDL for ND (in native units)
    def stat_value_row(r):
        if r['is_nd']:
            if pd.isna(r['MDL_native']):
                return np.nan
            return 0.5 * r['MDL_native']
        else:
            return r['numeric_native']
    sub['stat_value_native'] = sub.apply(stat_value_row, axis=1)

    pw = pwqo_map.get(analyte, {})
    obj = pw.get('objective', None)
    obj_unit = pw.get('objective_unit', None)
    pH_dependent = pw.get('pH_dependent', False)
    hardness_adjusted = pw.get('hardness_adjusted', False)

    # Prepare numeric_for_obj: convert detected numeric to objective unit (ug/L) when needed
    def conv_to_obj_unit(row):
        # pH and E. coli leave as native
        if pH_dependent or obj_unit in ("SU", "MPN/100mL"):
            return row['numeric_native']
        if obj_unit == 'ug/L':
            return to_ugL(row['numeric_native'], row['Units'])
        return row['numeric_native']

    sub['numeric_for_obj'] = sub.apply(conv_to_obj_unit, axis=1)

    # MDL in objective unit for ND comparisons
    def mdl_to_obj_unit(row):
        if pH_dependent or obj_unit in ("SU", "MPN/100mL"):
            return row['MDL_native']
        if obj_unit == 'ug/L':
            return to_ugL(row['MDL_native'], row['Units'])
        return row['MDL_native']
    sub['MDL_for_obj'] = sub.apply(mdl_to_obj_unit, axis=1)

    # For hardness-adjusted analytes, attach same-day mean hardness (mg/L)
    missing_hardness_count = 0
    if hardness_adjusted:
        merged = pd.merge(sub, hard_by_site_date, how='left', left_on=['Collection Site', 'Collected_date'], right_on=['Collection Site', 'Collected_date'])
        sub = merged.copy()
        sub['sample_obj_ugL'] = sub['mean_hardness_mgL'].apply(lambda h: compute_hardness_objective(analyte, h))
        missing_hardness_count = int(sub['mean_hardness_mgL'].isna().sum())
    else:
        if isinstance(obj, (int, float)):
            # static numeric objective in obj_unit (often ug/L)
            sub['sample_obj_ugL'] = obj
        else:
            sub['sample_obj_ugL'] = np.nan

    # Determine exceedance per rules
    def determine_exceed(row):
        if row['is_nd']:
            return False
        # pH
        if analyte == 'pH':
            val = row['numeric_native']
            if pd.isna(val):
                return False
            return (val < 6.5) or (val > 8.5)
        # Aluminum pH-dependent: require same-day pH at site/date
        if analyte == 'Aluminum':
            mask = (df['Collection Site'] == row['Collection Site']) & (df['Collected_date'] == row['Collected_date']) & (df['Analyte'].str.lower() == 'ph')
            phvals = df.loc[mask, 'numeric_result'].dropna().values
            if len(phvals) >= 1:
                ph = float(np.nanmean(phvals))
                if ph > 6.5:
                    threshold = 75.0
                elif 4.5 <= ph <= 5.5:
                    threshold = 15.0
                else:
                    # ambiguous range 5.5-6.5 or outside ranges: cannot reliably screen
                    return False
                num_ugL = to_ugL(row['numeric_native'], row['Units'])
                if pd.isna(num_ugL):
                    return False
                return num_ugL > threshold
            else:
                return False
        # Hardness-adjusted analytes
        if hardness_adjusted:
            obj_here = row.get('sample_obj_ugL', np.nan)
            if pd.isna(obj_here):
                # exclude from exceedance determination if no same-day hardness
                return False
            # numeric_for_obj should be in ug/L for comparison (or already converted)
            num = row['numeric_for_obj']
            if pd.isna(num):
                return False
            return num > obj_here
        # Other numeric objectives
        if isinstance(obj, (int, float)):
            num = row['numeric_for_obj']
            if pd.isna(num):
                return False
            return num > obj
        return False

    sub['is_exceedance'] = sub.apply(determine_exceed, axis=1)

    # N exceedances and station counts
    N_exceed = int(sub['is_exceedance'].sum())
    stations_with_exceed = sub.loc[sub['is_exceedance'], 'Collection Site'].dropna().unique()
    N_stations_exceed = int(len(stations_with_exceed))
    pct_exceed = round(100.0 * N_exceed / N_samples, 1) if N_samples > 0 else 0.0

    # % non-detect
    N_nd = int(sub['is_nd'].sum())
    pct_nd = round(100.0 * N_nd / N_samples, 1) if N_samples > 0 else 0.0

    # % non-detect where MDL > objective
    # For non-numeric objectives (hardness-adjusted or pH-dependent) set blank
    if hardness_adjusted or pH_dependent:
        pct_nd_mdl_gt_obj = ""
    else:
        count_nd_mdl_gt_obj = 0
        nd_with_mdl = sub[sub['is_nd'] & sub['MDL_for_obj'].notna()].copy()
        if nd_with_mdl.shape[0] > 0 and isinstance(obj, (int, float)):
            if obj_unit == 'ug/L':
                # MDL_for_obj already converted to ug/L (including ng/L -> ug/L handling)
                count_nd_mdl_gt_obj = int((nd_with_mdl['MDL_for_obj'] > obj).sum())
            else:
                count_nd_mdl_gt_obj = int((nd_with_mdl['MDL_for_obj'] > obj).sum())
        pct_nd_mdl_gt_obj = round(100.0 * count_nd_mdl_gt_obj / N_samples, 1) if N_samples > 0 else 0.0

    # Statistics in native data units (stat_value_native already in native units)
    stat_vals = pd.to_numeric(sub['stat_value_native'], errors='coerce')
    stat_vals_valid = stat_vals.dropna().values
    if len(stat_vals_valid) == 0:
        mean_val = np.nan
        median_val = np.nan
        p95_val = np.nan
    else:
        mean_val = float(np.nanmean(stat_vals_valid))
        median_val = float(np.nanmedian(stat_vals_valid))
        p95_val = float(np.nanpercentile(stat_vals_valid, 95))

    # Display objective
    if analyte == 'pH':
        obj_display = "6.5-8.5"
        obj_unit_display = "SU"
    else:
        if isinstance(pw.get('objective', None), (int, float)):
            obj_display = pw['objective']
            obj_unit_display = pw.get('objective_unit', '')
        else:
            if hardness_adjusted:
                obj_display = "hardness-adjusted"
                obj_unit_display = pw.get('objective_unit', '')
            elif pH_dependent:
                obj_display = "pH-dependent"
                obj_unit_display = pw.get('objective_unit', '')
            else:
                obj_display = ""
                obj_unit_display = pw.get('objective_unit', '')

    notes = pw.get('note', '')
    if hardness_adjusted:
        notes = notes + f" {missing_hardness_count} sample(s) lacked same-day hardness and were excluded from exceedance screening."
    if analyte == 'Aluminum':
        notes = notes + " Aluminum screening attempted only where same-day pH available."

    # Build row
    summary_rows.append({
        "Analyte": analyte,
        "Units (data)": data_unit if data_unit is not None else "",
        "PWQO objective": obj_display,
        "Objective units": obj_unit_display,
        "PWQO notes / requirement": notes.strip(),
        "N samples": N_samples,
        "N exceedances": N_exceed,
        "% exceedances": pct_exceed,
        "N stations with ≥1 exceedance": N_stations_exceed,
        "% non-detect": pct_nd,
        "% non-detect MDL>objective": pct_nd_mdl_gt_obj,
        "Mean (½MDL for ND)": mean_val,
        "Median (½MDL for ND)": median_val,
        "95th percentile (½MDL for ND)": p95_val
    })

# --- Create final DataFrame ---
summary_df = pd.DataFrame(summary_rows, columns=[
    "Analyte",
    "Units (data)",
    "PWQO objective",
    "Objective units",
    "PWQO notes / requirement",
    "N samples",
    "N exceedances",
    "% exceedances",
    "N stations with ≥1 exceedance",
    "% non-detect",
    "% non-detect MDL>objective",
    "Mean (½MDL for ND)",
    "Median (½MDL for ND)",
    "95th percentile (½MDL for ND)"
])

# Round numeric statistic columns sensibly; keep stats in native units
def fmt_stat(x):
    if pd.isna(x):
        return ""
    try:
        # Use 3 decimal places except when values are large; still safe to round 3
        return round(float(x), 3)
    except:
        return x

summary_df["Mean (½MDL for ND)"] = summary_df["Mean (½MDL for ND)"].apply(fmt_stat)
summary_df["Median (½MDL for ND)"] = summary_df["Median (½MDL for ND)"].apply(fmt_stat)
summary_df["95th percentile (½MDL for ND)"] = summary_df["95th percentile (½MDL for ND)"].apply(fmt_stat)

# For "% non-detect MDL>objective", ensure blank strings for analytes where set as ""
summary_df["% non-detect MDL>objective"] = summary_df["% non-detect MDL>objective"].apply(lambda x: "" if (x == "" or pd.isna(x)) else x)

# Save CSV (overwrite)
summary_df.to_csv(csv_out, index=False)

# Print created files and python script name per instructions
print(py_name)
print(csv_out)