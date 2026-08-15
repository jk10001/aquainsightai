# filename: muscatine_wrrf_digesters_2022_event_chronology.py
import pandas as pd
import numpy as np

file_name = "muscatine_wrrf_digesters_2022.xlsx"
base_name = "muscatine_wrrf_digesters_2022_event_chronology"

xls = pd.ExcelFile(file_name)
if "daily_data" not in xls.sheet_names:
    raise ValueError("Sheet 'daily_data' not found in workbook.")

df = pd.read_excel(xls, sheet_name="daily_data")

required_cols = [
    "Year", "Month", "Day",
    "V-TWAS_m3d", "V-PS_m3d", "V-HSW_m3d",
    "Dig1-VFA-Alkalinity_ratio", "Dig2-VFA-Alkalinity_ratio",
    "Dig1-pH", "Dig2-pH",
    "Dig1-alk_mgL", "Dig2-alk_mgL",
    "TWAS-VS_percent", "PS-VS_percent", "HSW-VS_percent",
]
missing = [c for c in required_cols if c not in df.columns]
if missing:
    raise ValueError(f"Missing required columns: {missing}")

hsw_d1_col = "HSW-Dig1-portion_pct"
hsw_d2_col = "HSW2-Dig2-portion_pct"
has_hsw_d1 = hsw_d1_col in df.columns
has_hsw_d2 = hsw_d2_col in df.columns

df["Date"] = pd.to_datetime(dict(year=df["Year"], month=df["Month"], day=df["Day"]), errors="coerce")
df = df.sort_values("Date").reset_index(drop=True)

def safe_num(x):
    return np.nan if pd.isna(x) else float(x)

def compute_hsw_fractions(row):
    p1 = row[hsw_d1_col] if has_hsw_d1 else np.nan
    p2 = row[hsw_d2_col] if has_hsw_d2 else np.nan

    if pd.isna(p1) and pd.isna(p2):
        return 0.5, 0.5
    if pd.isna(p1) and not pd.isna(p2):
        f2 = float(p2) / 100.0
        return 1.0 - f2, f2
    if not pd.isna(p1) and pd.isna(p2):
        f1 = float(p1) / 100.0
        return f1, 1.0 - f1

    f1 = float(p1) / 100.0
    f2 = float(p2) / 100.0
    s = f1 + f2
    if pd.isna(s) or s <= 0:
        return 0.5, 0.5
    if abs(s - 1.0) <= 0.01:
        return f1, f2
    return f1 / s, f2 / s

fractions = df.apply(compute_hsw_fractions, axis=1, result_type="expand")
df["HSW fraction to Digester 1 [-]"] = fractions[0]
df["HSW fraction to Digester 2 [-]"] = fractions[1]

# 7-day trailing rolling average VS loading rate, same assumptions as prior chart
def compute_daily_olr(row):
    twas = 0.0 if pd.isna(row["V-TWAS_m3d"]) else float(row["V-TWAS_m3d"])
    ps = 0.0 if pd.isna(row["V-PS_m3d"]) else float(row["V-PS_m3d"])
    hsw = 0.0 if pd.isna(row["V-HSW_m3d"]) else float(row["V-HSW_m3d"])

    twas_vs = 0.0 if pd.isna(row["TWAS-VS_percent"]) else float(row["TWAS-VS_percent"]) / 100.0
    ps_vs = 0.0 if pd.isna(row["PS-VS_percent"]) else float(row["PS-VS_percent"]) / 100.0
    hsw_vs = 0.0 if pd.isna(row["HSW-VS_percent"]) else float(row["HSW-VS_percent"]) / 100.0

    f1 = float(row["HSW fraction to Digester 1 [-]"])
    f2 = float(row["HSW fraction to Digester 2 [-]"])

    twas_to_d1 = 0.5 * twas
    twas_to_d2 = 0.5 * twas
    ps_to_d1 = 0.5 * ps
    ps_to_d2 = 0.5 * ps
    hsw_to_d1 = hsw * f1
    hsw_to_d2 = hsw * f2

    vs_d1 = (twas_to_d1 * 1000.0 * twas_vs) + (ps_to_d1 * 1000.0 * ps_vs) + (hsw_to_d1 * 1000.0 * hsw_vs)
    vs_d2 = (twas_to_d2 * 1000.0 * twas_vs) + (ps_to_d2 * 1000.0 * ps_vs) + (hsw_to_d2 * 1000.0 * hsw_vs)

    return vs_d1 / 1625.0 if vs_d1 > 0 else np.nan, vs_d2 / 1625.0 if vs_d2 > 0 else np.nan

olr_pairs = df.apply(compute_daily_olr, axis=1, result_type="expand")
df["Digester 1 daily VS loading rate [kgVS/m3-d]"] = olr_pairs[0]
df["Digester 2 daily VS loading rate [kgVS/m3-d]"] = olr_pairs[1]
df["Digester 1 7d avg VS loading rate [kgVS/m3-d]"] = (
    df["Digester 1 daily VS loading rate [kgVS/m3-d]"].rolling(window=7, min_periods=4).mean()
)
df["Digester 2 7d avg VS loading rate [kgVS/m3-d]"] = (
    df["Digester 2 daily VS loading rate [kgVS/m3-d]"].rolling(window=7, min_periods=4).mean()
)

def build_event_periods(flag_series, dates):
    """Return merged periods where gaps of <=1 day are merged."""
    periods = []
    in_event = False
    start_idx = None
    prev_flag_idx = None

    flag_indices = np.where(flag_series.to_numpy())[0].tolist()

    if not flag_indices:
        return periods

    # Build initial contiguous periods, allowing one-day gaps
    current_start = flag_indices[0]
    current_end = flag_indices[0]

    for idx in flag_indices[1:]:
        gap_days = (dates.iloc[idx] - dates.iloc[current_end]).days
        if gap_days <= 2:  # one-day gap means consecutive flagged days are 2 days apart
            current_end = idx
        else:
            periods.append((current_start, current_end))
            current_start = idx
            current_end = idx
    periods.append((current_start, current_end))
    return periods

def summarize_event(digester, event_start_idx, event_end_idx):
    if digester == 1:
        vfa_col = "Dig1-VFA-Alkalinity_ratio"
        ph_col = "Dig1-pH"
        alk_col = "Dig1-alk_mgL"
        olr_col = "Digester 1 7d avg VS loading rate [kgVS/m3-d]"
        hsw_frac_col = "HSW fraction to Digester 1 [-]"
    else:
        vfa_col = "Dig2-VFA-Alkalinity_ratio"
        ph_col = "Dig2-pH"
        alk_col = "Dig2-alk_mgL"
        olr_col = "Digester 2 7d avg VS loading rate [kgVS/m3-d]"
        hsw_frac_col = "HSW fraction to Digester 2 [-]"

    event = df.iloc[event_start_idx:event_end_idx + 1].copy()

    vfa = event[vfa_col]
    ph = event[ph_col]
    alk = event[alk_col]
    olr = event[olr_col]
    hsw_frac = event[hsw_frac_col]

    any_vfa_upset = (vfa > 0.40).fillna(False)
    any_ph_low = (ph < 6.8).fillna(False)
    any_alk_low = (alk < 3000).fillna(False)
    any_olr_high = (olr > 3.2).fillna(False)
    any_hsw_mostly = (hsw_frac >= 0.90).fillna(False)

    notes = []
    if any_vfa_upset.any():
        notes.append("VFA/Alk upset")
    if any_ph_low.any():
        notes.append("low pH")
    if any_alk_low.any():
        notes.append("low alkalinity")
    if any_olr_high.any():
        notes.append("high OLR")
    if any_hsw_mostly.any():
        notes.append("HSW mostly to this digester")

    return {
        "Digester [1/2]": digester,
        "Event start date [YYYY-MM-DD]": event["Date"].iloc[0].strftime("%Y-%m-%d"),
        "Event end date [YYYY-MM-DD]": event["Date"].iloc[-1].strftime("%Y-%m-%d"),
        "Event duration [days]": int((event["Date"].iloc[-1] - event["Date"].iloc[0]).days + 1),
        "Max VFA/Alk ratio [-]": float(vfa.max(skipna=True)) if vfa.notna().any() else np.nan,
        "Days VFA/Alk >0.40 [d]": int((vfa > 0.40).sum()),
        "Min pH [s.u.]": float(ph.min(skipna=True)) if ph.notna().any() else np.nan,
        "Days pH <6.8 [d]": int((ph < 6.8).sum()),
        "Min alkalinity [mg/L as CaCO3]": float(alk.min(skipna=True)) if alk.notna().any() else np.nan,
        "Days alkalinity <3000 [d]": int((alk < 3000).sum()),
        "Max 7d avg VS OLR [kgVS/m3-d]": float(olr.max(skipna=True)) if olr.notna().any() else np.nan,
        "Days 7d avg VS OLR >3.2 [d]": int((olr > 3.2).sum()),
        "Days HSW fraction >=0.90 to this digester [d]": int((hsw_frac >= 0.90).sum()),
        "Notes (auto)": "; ".join(notes),
    }

records = []
for digester in [1, 2]:
    if digester == 1:
        flag = (
            (df["Dig1-VFA-Alkalinity_ratio"] > 0.40) |
            (df["Dig1-pH"] < 6.8) |
            (df["Dig1-alk_mgL"] < 3000) |
            (df["Digester 1 7d avg VS loading rate [kgVS/m3-d]"] > 3.2)
        ).fillna(False)
    else:
        flag = (
            (df["Dig2-VFA-Alkalinity_ratio"] > 0.40) |
            (df["Dig2-pH"] < 6.8) |
            (df["Dig2-alk_mgL"] < 3000) |
            (df["Digester 2 7d avg VS loading rate [kgVS/m3-d]"] > 3.2)
        ).fillna(False)

    periods = build_event_periods(flag, df["Date"])
    for start_idx, end_idx in periods:
        records.append(summarize_event(digester, start_idx, end_idx))

out_df = pd.DataFrame(records)

# Stable ordering
if not out_df.empty:
    out_df["__sortdate"] = pd.to_datetime(out_df["Event start date [YYYY-MM-DD]"], errors="coerce")
    out_df = out_df.sort_values(["Digester [1/2]", "__sortdate"]).drop(columns="__sortdate")
else:
    out_df = pd.DataFrame(columns=[
        "Digester [1/2]",
        "Event start date [YYYY-MM-DD]",
        "Event end date [YYYY-MM-DD]",
        "Event duration [days]",
        "Max VFA/Alk ratio [-]",
        "Days VFA/Alk >0.40 [d]",
        "Min pH [s.u.]",
        "Days pH <6.8 [d]",
        "Min alkalinity [mg/L as CaCO3]",
        "Days alkalinity <3000 [d]",
        "Max 7d avg VS OLR [kgVS/m3-d]",
        "Days 7d avg VS OLR >3.2 [d]",
        "Days HSW fraction >=0.90 to this digester [d]",
        "Notes (auto)",
    ])

csv_file = f"{base_name}.csv"
out_df.to_csv(csv_file, index=False)

print(f"Created file: {csv_file}")
print("Created file: muscatine_wrrf_digesters_2022_event_chronology.py")