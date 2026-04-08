# filename: top10_tp_exceedance_2024_v2.py
import os
import sys
import pandas as pd
import numpy as np

def parse_result_to_numeric(result_raw, result_call, value_qualifier, detection_limit):
    """
    Parse a Results cell for phosphorus into a numeric TP_value_ugL per rules:
    - If Results begins with '<' OR Value Qualifier is '<' OR Result Call == 'BDL':
        TP_value_ugL = 0.5 * Detection Limit
    - Else:
        TP_value_ugL = float(Results)
    Returns float or np.nan if cannot determine.
    """
    # Normalize inputs
    if pd.isna(result_raw):
        result_raw = ""
    result_str = str(result_raw).strip()
    result_call = "" if pd.isna(result_call) else str(result_call).strip()
    value_qualifier = "" if pd.isna(value_qualifier) else str(value_qualifier).strip()
    # Determine if censored
    censored = False
    if result_str.startswith("<"):
        censored = True
    if value_qualifier == "<":
        censored = True
    if str(result_call).upper() == "BDL":
        censored = True
    # Try to get numeric detection limit
    dl = None
    try:
        if pd.notna(detection_limit):
            # detection_limit may be numeric or string
            dl = float(str(detection_limit).strip())
    except Exception:
        dl = None
    # If censored -> use 0.5 * DL if DL available; else try to parse number after '<' and use half of that
    if censored:
        if dl is not None and not np.isnan(dl):
            return 0.5 * dl
        # fallback: try extract number from result_str like "<2" -> 2
        if result_str.startswith("<"):
            numpart = result_str.lstrip("<").strip()
            # Remove any non-numeric trailing chars
            # Keep characters that are digits, decimal, or leading minus
            numpart_clean = ""
            for ch in numpart:
                if ch.isdigit() or ch in ".-+eE":
                    numpart_clean += ch
                else:
                    break
            try:
                val = float(numpart_clean)
                return 0.5 * val
            except Exception:
                return np.nan
        # If cannot parse, return nan
        return np.nan
    # Not censored: try to parse numeric value from result_str (may include commas)
    try:
        clean = result_str.replace(",", "").strip()
        # Some results may include qualifiers after a space; take first token that can parse
        tokens = clean.split()
        for tok in tokens:
            try:
                val = float(tok)
                return val
            except Exception:
                continue
        # If none parsed, raise
        return np.nan
    except Exception:
        return np.nan

def safe_numeric(x):
    try:
        if pd.isna(x):
            return np.nan
        sx = str(x).strip().replace(",", "")
        return float(sx)
    except Exception:
        return np.nan

def station_display_from_padded(padded):
    """
    Return unpadded numeric-looking string (leading zeros removed).
    If padded is empty or NaN, return empty string.
    If padded contains non-digits, return original stripped string.
    """
    if padded is None:
        return ""
    s = str(padded).strip()
    if s == "" or s.lower() == "nan":
        return ""
    # If all digits, strip leading zeros but keep single '0' if all zeros
    if s.isdigit():
        s2 = s.lstrip("0")
        return s2 if s2 != "" else "0"
    # else return as-is
    return s

def main():
    py_filename = os.path.basename(__file__)
    excel_file = "Ontario_PWQMN_2024.xlsx"
    out_csv = "top10_tp_exceedance_2024_v2.csv"

    # Read Data and Stations sheets
    try:
        df = pd.read_excel(excel_file, sheet_name="Data", engine="openpyxl", dtype=object)
    except Exception as e:
        print(f"ERROR reading 'Data' sheet from '{excel_file}': {e}", file=sys.stderr)
        raise

    try:
        df_stations = pd.read_excel(excel_file, sheet_name="Stations", engine="openpyxl", dtype=object)
    except Exception as e:
        print(f"ERROR reading 'Stations' sheet from '{excel_file}': {e}", file=sys.stderr)
        raise

    # Validate expected columns
    expected_data_cols = {"Collection Site", "Analyte", "Results", "Result Call", "Detection Limit", "Value Qualifier"}
    missing = expected_data_cols - set(df.columns)
    if missing:
        raise KeyError(f"Missing expected columns in Data sheet: {missing}")

    expected_st_cols = {"STATION", "NAME"}
    missing2 = expected_st_cols - set(df_stations.columns)
    if missing2:
        raise KeyError(f"Missing expected columns in Stations sheet: {missing2}")

    # Filter to Analyte = 'Phosphorus; total'
    df_tp = df[df["Analyte"].astype(str).str.strip() == "Phosphorus; total"].copy()

    # If no TP rows, create empty output
    if df_tp.shape[0] == 0:
        df_empty = pd.DataFrame(columns=[
            "Station number",
            "Station name",
            "N samples",
            "N exceedances",
            "Exceedance percent [%]",
            "Median total phosphorus [ug/L]",
            "95th percentile total phosphorus [ug/L]",
            "Max total phosphorus [ug/L]"
        ])
        df_empty.to_csv(out_csv, index=False)
        print(py_filename)
        print(out_csv)
        return

    # Create padded join keys (zfill to 11) for both sheets
    def make_padded(val):
        if pd.isna(val):
            return ""
        s = str(val).strip()
        if s.lower() == "nan":
            return ""
        # If the value contains decimal like 1.0 from Excel, remove .0 if integer-like
        if s.replace(".", "", 1).isdigit() and "." in s:
            # try convert to int if appropriate
            try:
                f = float(s)
                if f.is_integer():
                    s = str(int(f))
            except Exception:
                pass
        # Strip any whitespace then pad to 11 digits if digits-only; else pad the string as-is to 11 chars
        # We will pad the string (zfill) regardless, but ensure it's digits-first when possible.
        return s.zfill(11)

    # Add padded id to Data
    df_tp["station_padded_id"] = df_tp["Collection Site"].apply(make_padded)

    # Add padded id to Stations
    df_stations["station_padded_id"] = df_stations["STATION"].apply(make_padded)

    # Create station name map using padded id
    # If multiple identical padded ids exist in Stations, last occurrence will be used
    station_name_map = pd.Series(df_stations["NAME"].values, index=df_stations["station_padded_id"]).to_dict()

    # Prepare detection limit numeric
    df_tp["Detection Limit numeric"] = df_tp["Detection Limit"].apply(safe_numeric)

    # Compute TP_value_ugL per rules using padded-aware rows
    df_tp["TP_value_ugL"] = df_tp.apply(
        lambda r: parse_result_to_numeric(
            r.get("Results", ""),
            r.get("Result Call", ""),
            r.get("Value Qualifier", ""),
            r.get("Detection Limit numeric", np.nan)
        ), axis=1
    )

    # Drop rows where TP_value_ugL is NaN (unable to compute)
    df_tp_valid = df_tp[~df_tp["TP_value_ugL"].isna()].copy()

    # Exceedance flag
    df_tp_valid["Exceeds_PWQO"] = df_tp_valid["TP_value_ugL"] > 30.0

    # Aggregate per padded station id
    grouped = df_tp_valid.groupby("station_padded_id", dropna=False)
    agg_list = []
    for padded_id, g in grouped:
        n_samples = int(len(g))
        n_exceed = int(g["Exceeds_PWQO"].sum())
        exceed_pct = 100.0 * n_exceed / n_samples if n_samples > 0 else 0.0
        median = float(g["TP_value_ugL"].median(skipna=True)) if n_samples > 0 else np.nan
        # pandas quantile interpolation parameter: use interpolation="linear" for compatibility
        try:
            p95 = float(g["TP_value_ugL"].quantile(0.95, interpolation="linear"))
        except TypeError:
            # newer pandas uses method parameter name 'method' instead of 'interpolation'
            p95 = float(g["TP_value_ugL"].quantile(0.95, method="linear"))
        mx = float(g["TP_value_ugL"].max(skipna=True)) if n_samples > 0 else np.nan

        # Station number display (un-padded numeric-looking string)
        station_number_display = station_display_from_padded(padded_id)

        # Station name by padded join; QA fallback if unmatched
        station_name = station_name_map.get(padded_id, "")
        if station_name is None or str(station_name).strip() == "":
            station_name = "Unknown (not in Stations sheet)"

        agg_list.append({
            "station_padded_id": padded_id,
            "Station number": station_number_display,
            "Station name": station_name,
            "N samples": n_samples,
            "N exceedances": n_exceed,
            "Exceedance percent [%]": exceed_pct,
            "Median total phosphorus [ug/L]": median,
            "95th percentile total phosphorus [ug/L]": p95,
            "Max total phosphorus [ug/L]": mx
        })

    df_agg = pd.DataFrame(agg_list)

    # Sorting: Exceedance percent (desc), Median (desc), N samples (desc), Station number (asc)
    df_agg["Exceedance_sort"] = df_agg["Exceedance percent [%]"].fillna(-1)
    df_agg["Median_sort"] = df_agg["Median total phosphorus [ug/L]"].fillna(-np.inf)
    df_agg["N_samples_sort"] = df_agg["N samples"].fillna(0).astype(int)

    # Station sort key: numeric if possible else string
    def station_sort_key_wrap(s):
        if s == "" or s is None:
            return ""
        try:
            return int(s)
        except Exception:
            return str(s)

    df_agg["Station_sort_key"] = df_agg["Station number"].apply(station_sort_key_wrap)

    df_agg_sorted = df_agg.sort_values(
        by=["Exceedance_sort", "Median_sort", "N_samples_sort", "Station_sort_key"],
        ascending=[False, False, False, True]
    ).reset_index(drop=True)

    # Select top 10
    df_top10 = df_agg_sorted.head(10).copy()

    # Round numeric output columns: percent one decimal, med/p95/max two decimals
    df_top10["Exceedance percent [%]"] = df_top10["Exceedance percent [%]"].round(1)
    df_top10["Median total phosphorus [ug/L]"] = df_top10["Median total phosphorus [ug/L]"].round(2)
    df_top10["95th percentile total phosphorus [ug/L]"] = df_top10["95th percentile total phosphorus [ug/L]"].round(2)
    df_top10["Max total phosphorus [ug/L]"] = df_top10["Max total phosphorus [ug/L]"].round(2)

    # Prepare final CSV with required column order and exactly top rows
    final_cols = [
        "Station number",
        "Station name",
        "N samples",
        "N exceedances",
        "Exceedance percent [%]",
        "Median total phosphorus [ug/L]",
        "95th percentile total phosphorus [ug/L]",
        "Max total phosphorus [ug/L]"
    ]
    df_final_out = df_top10[final_cols].copy()

    # Save to CSV (utf-8, no index)
    df_final_out.to_csv(out_csv, index=False, encoding="utf-8")

    # Print created filenames to terminal: python file and csv file
    print(py_filename)
    print(out_csv)

if __name__ == "__main__":
    main()