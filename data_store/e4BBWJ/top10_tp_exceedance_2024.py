# filename: top10_tp_exceedance_2024.py
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
    if result_call.upper() == "BDL":
        censored = True
    # Try to get numeric detection limit
    dl = None
    try:
        if pd.notna(detection_limit):
            dl = float(detection_limit)
    except Exception:
        dl = None
    # If censored -> use 0.5 * DL if DL available; else try to parse number after '<' and use half of that
    if censored:
        if dl is not None:
            return 0.5 * dl
        # fallback: try extract number from result_str like "<2" -> 2
        if result_str.startswith("<"):
            numpart = result_str.lstrip("<").strip()
            # Remove any non-numeric trailing chars
            try:
                val = float(numpart)
                return 0.5 * val
            except Exception:
                return np.nan
        # If cannot parse, return nan
        return np.nan
    # Not censored: try to parse numeric value from result_str (may include commas)
    try:
        clean = result_str.replace(",", "").strip()
        # Some results may be like "1 " or "1.0" or have qualifiers; attempt to parse initial numeric token
        # Split on whitespace and take first token
        first_tok = clean.split()[0] if len(clean.split())>0 else clean
        val = float(first_tok)
        return val
    except Exception:
        # final fallback: if DL exists, maybe treat as 0.5*DL? But spec says only when censored -> so return nan
        return np.nan

def main():
    py_filename = os.path.basename(__file__)
    excel_file = "Ontario_PWQMN_2024.xlsx"
    out_csv = "top10_tp_exceedance_2024.csv"

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

    # Ensure expected columns exist
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

    # Apply parsing to compute TP_value_ugL
    # Ensure Detection Limit is passed appropriately: may be numeric or string
    # We'll create a numeric detection limit column
    def safe_numeric(x):
        try:
            if pd.isna(x):
                return np.nan
            # strip possible strings
            sx = str(x).strip().replace(",", "")
            return float(sx)
        except Exception:
            return np.nan

    df_tp["Detection Limit numeric"] = df_tp["Detection Limit"].apply(safe_numeric)

    # Compute TP_value_ugL per rules
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

    # Normalize station ID keys for joining: use exact equality as strings (do not zero-pad)
    # Some Collection Site values may be numeric; convert to string without leading/trailing spaces
    df_tp_valid["Station_key"] = df_tp_valid["Collection Site"].apply(lambda x: "" if pd.isna(x) else str(x).strip())

    df_stations["STATION_key"] = df_stations["STATION"].apply(lambda x: "" if pd.isna(x) else str(x).strip())
    # Map station name by exact key match
    station_name_map = pd.Series(df_stations["NAME"].values, index=df_stations["STATION_key"]).to_dict()

    # Aggregate per station
    agg_list = []
    grouped = df_tp_valid.groupby("Station_key")
    for station_key, g in grouped:
        n_samples = int(len(g))
        n_exceed = int(g["Exceeds_PWQO"].sum())
        exceed_pct = 100.0 * n_exceed / n_samples if n_samples > 0 else 0.0
        median = float(g["TP_value_ugL"].median(skipna=True)) if n_samples>0 else np.nan
        p95 = float(g["TP_value_ugL"].quantile(0.95, interpolation="linear")) if n_samples>0 else np.nan
        mx = float(g["TP_value_ugL"].max(skipna=True)) if n_samples>0 else np.nan
        # Station number display without leading zeros (as per user instruction)
        station_display = station_key
        # If numeric-like and has leading zeros, converting to int will drop zeros
        # But requirement: "Don’t worry about including leading zeros" which means display as-is without padding.
        # So we'll try to preserve original but if it's a numeric string, we can strip leading zeros only if they exist and are not supposed to be kept.
        # To follow instruction strictly: display without leading zeros -> attempt to lstrip zeros but keep "0" if that's the value.
        if station_display is None or station_display == "":
            station_display_clean = ""
        else:
            s = station_display
            # If s is all digits, strip leading zeros but keep at least one digit
            if s.isdigit():
                s2 = s.lstrip("0")
                station_display_clean = s2 if s2 != "" else "0"
            else:
                station_display_clean = s
        station_name = station_name_map.get(station_key, "")
        agg_list.append({
            "Station number": station_display_clean,
            "Station_key": station_key,
            "Station name": station_name,
            "N samples": n_samples,
            "N exceedances": n_exceed,
            "Exceedance percent [%]": exceed_pct,
            "Median total phosphorus [ug/L]": median,
            "95th percentile total phosphorus [ug/L]": p95,
            "Max total phosphorus [ug/L]": mx
        })

    df_agg = pd.DataFrame(agg_list)

    # If some stations had empty station key, include them but their Station number will be blank string.
    # Sorting: Exceedance percent (desc), Median (desc), N samples (desc), Station number (asc)
    # For numeric sorts, ensure NaNs handled (treat NaN median as -inf so they go last)
    df_agg["Exceedance_sort"] = df_agg["Exceedance percent [%]"].fillna(-1)
    df_agg["Median_sort"] = df_agg["Median total phosphorus [ug/L]"].fillna(-np.inf)
    df_agg["N_samples_sort"] = df_agg["N samples"].fillna(0).astype(int)
    # For station number asc, convert to a comparable key: if numeric string, integer; else string
    def station_sort_key(s):
        if s == "" or s is None:
            return ""
        try:
            return int(s)
        except Exception:
            return str(s)
    df_agg["Station_sort_key"] = df_agg["Station number"].apply(station_sort_key)

    df_agg_sorted = df_agg.sort_values(
        by=["Exceedance_sort", "Median_sort", "N_samples_sort", "Station_sort_key"],
        ascending=[False, False, False, True]
    ).reset_index(drop=True)

    # Select top 10
    df_top10 = df_agg_sorted.head(10).copy()

    # Round numeric output columns to reasonable precision: percent one decimal, med/p95/max two decimals
    df_top10["Exceedance percent [%]"] = df_top10["Exceedance percent [%]"].round(1)
    df_top10["Median total phosphorus [ug/L]"] = df_top10["Median total phosphorus [ug/L]"].round(2)
    df_top10["95th percentile total phosphorus [ug/L]"] = df_top10["95th percentile total phosphorus [ug/L]"].round(2)
    df_top10["Max total phosphorus [ug/L]"] = df_top10["Max total phosphorus [ug/L]"].round(2)

    # Prepare final CSV with required column order and exactly 10 rows (or fewer if fewer stations)
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