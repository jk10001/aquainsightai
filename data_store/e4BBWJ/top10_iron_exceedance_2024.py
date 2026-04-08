# filename: top10_iron_exceedance_2024.py
import os
import sys
import pandas as pd
import numpy as np

def safe_float(x):
    try:
        if pd.isna(x):
            return np.nan
        s = str(x).strip().replace(",", "")
        return float(s)
    except Exception:
        return np.nan

def make_padded(val):
    if pd.isna(val):
        return ""
    s = str(val).strip()
    if s.lower() == "nan" or s == "":
        return ""
    # If Excel wrote floats like 1.0, convert to int string when appropriate
    try:
        if "." in s:
            f = float(s)
            if f.is_integer():
                s = str(int(f))
    except Exception:
        pass
    return s.zfill(11)

def station_display_from_padded(padded):
    """
    Return unpadded numeric-looking string (leading zeros removed).
    If padded is empty, return empty string.
    """
    if padded is None:
        return ""
    s = str(padded).strip()
    if s == "" or s.lower() == "nan":
        return ""
    if s.isdigit():
        s2 = s.lstrip("0")
        return s2 if s2 != "" else "0"
    return s

def parse_iron_value(result_raw, result_call, value_qualifier, detection_limit):
    """
    Parse Results into numeric Iron [ug/L] per rules:
    - If Results begins with '<' OR Value Qualifier == '<' OR Result Call == 'BDL' -> use 0.5 * Detection Limit (DL expected in ug/L)
    - Else parse numeric value from Results
    Returns float or np.nan
    """
    result = "" if pd.isna(result_raw) else str(result_raw).strip()
    rc = "" if pd.isna(result_call) else str(result_call).strip()
    vq = "" if pd.isna(value_qualifier) else str(value_qualifier).strip()

    censored = False
    if isinstance(result, str) and result.startswith("<"):
        censored = True
    if vq == "<":
        censored = True
    if str(rc).upper() == "BDL":
        censored = True

    dl_num = safe_float(detection_limit)

    if censored:
        if not np.isnan(dl_num):
            return 0.5 * dl_num
        # fallback: attempt to parse number after '<' (e.g., "<2")
        if isinstance(result, str) and result.startswith("<"):
            numpart = result.lstrip("<").strip()
            # keep leading numeric chars
            token = ""
            for ch in numpart:
                if ch.isdigit() or ch in ".-+eE":
                    token += ch
                else:
                    break
            try:
                val = float(token)
                return 0.5 * val
            except Exception:
                return np.nan
        return np.nan
    else:
        # not censored: parse numeric token from result
        s = str(result).replace(",", "")
        tokens = s.split()
        for tok in tokens:
            try:
                return float(tok)
            except Exception:
                continue
        return safe_float(s)

def main():
    py_filename = os.path.basename(__file__)
    excel_file = "Ontario_PWQMN_2024.xlsx"
    out_csv = "top10_iron_exceedance_2024.csv"

    # Read sheets
    try:
        df = pd.read_excel(excel_file, sheet_name="Data", engine="openpyxl", dtype=object)
    except Exception as e:
        print(f"ERROR reading 'Data' sheet from '{excel_file}': {e}", file=sys.stderr)
        raise

    try:
        df_st = pd.read_excel(excel_file, sheet_name="Stations", engine="openpyxl", dtype=object)
    except Exception as e:
        print(f"ERROR reading 'Stations' sheet from '{excel_file}': {e}", file=sys.stderr)
        raise

    # Validate required columns
    required_data_cols = {"Collection Site", "Analyte", "Results", "Result Call", "Detection Limit", "Value Qualifier"}
    missing = required_data_cols - set(df.columns)
    if missing:
        raise KeyError(f"Missing expected columns in Data sheet: {missing}")

    required_st_cols = {"STATION", "NAME"}
    missing2 = required_st_cols - set(df_st.columns)
    if missing2:
        raise KeyError(f"Missing expected columns in Stations sheet: {missing2}")

    # Filter to Iron analyte
    df_iron = df[df["Analyte"].astype(str).str.strip() == "Iron"].copy()

    # If no iron rows, create empty CSV and exit
    if df_iron.shape[0] == 0:
        df_empty = pd.DataFrame(columns=[
            "Station number",
            "Station name",
            "N samples",
            "N exceedances",
            "% exceedances [%]",
            "Median iron [µg/L]",
            "95th percentile iron [µg/L]",
            "Max iron [µg/L]"
        ])
        df_empty.to_csv(out_csv, index=False)
        print(py_filename)
        print(out_csv)
        return

    # Add padded id to Data and Stations for joining
    df_iron["station_padded_id"] = df_iron["Collection Site"].apply(make_padded)
    df_st["station_padded_id"] = df_st["STATION"].apply(make_padded)

    # Map station padded -> name
    station_name_map = pd.Series(df_st["NAME"].values, index=df_st["station_padded_id"]).to_dict()

    # Ensure Detection Limit numeric
    df_iron["Detection Limit numeric"] = df_iron["Detection Limit"].apply(safe_float)

    # Parse numeric Iron [µg/L]
    df_iron["Iron_ugL"] = df_iron.apply(
        lambda r: parse_iron_value(
            r.get("Results", ""),
            r.get("Result Call", ""),
            r.get("Value Qualifier", ""),
            r.get("Detection Limit numeric", np.nan)
        ), axis=1
    )

    # Keep only rows with numeric Iron_ugL
    df_valid = df_iron[~df_iron["Iron_ugL"].isna()].copy()

    # If none valid, output empty CSV
    if df_valid.shape[0] == 0:
        df_empty = pd.DataFrame(columns=[
            "Station number",
            "Station name",
            "N samples",
            "N exceedances",
            "% exceedances [%]",
            "Median iron [µg/L]",
            "95th percentile iron [µg/L]",
            "Max iron [µg/L]"
        ])
        df_empty.to_csv(out_csv, index=False)
        print(py_filename)
        print(out_csv)
        return

    # Exceedance flag: Iron > 300 ug/L
    df_valid["Exceeds"] = df_valid["Iron_ugL"].astype(float) > 300.0

    # Aggregate per padded station
    grouped = df_valid.groupby("station_padded_id", dropna=False)
    agg_list = []
    for padded_id, g in grouped:
        n_samples = int(len(g))
        n_exceed = int(g["Exceeds"].sum())
        exceed_pct = 100.0 * n_exceed / n_samples if n_samples > 0 else 0.0
        median = float(g["Iron_ugL"].median(skipna=True)) if n_samples > 0 else np.nan
        # compute 95th percentile with interpolation; handle pandas version differences
        try:
            p95 = float(g["Iron_ugL"].quantile(0.95, interpolation="linear"))
        except TypeError:
            p95 = float(g["Iron_ugL"].quantile(0.95, method="linear"))
        mx = float(g["Iron_ugL"].max(skipna=True)) if n_samples > 0 else np.nan

        station_number_display = station_display_from_padded(padded_id)
        station_name = station_name_map.get(padded_id, "")
        if station_name is None or str(station_name).strip() == "":
            station_name = "Unknown (not in Stations sheet)"

        agg_list.append({
            "station_padded_id": padded_id,
            "Station number": station_number_display,
            "Station name": station_name,
            "N samples": n_samples,
            "N exceedances": n_exceed,
            "% exceedances [%]": exceed_pct,
            "Median iron [µg/L]": median,
            "95th percentile iron [µg/L]": p95,
            "Max iron [µg/L]": mx
        })

    df_agg = pd.DataFrame(agg_list)

    # Sorting: % exceedances desc, median desc, N samples desc, Station number asc
    df_agg["Exceed_sort"] = df_agg["% exceedances [%]"].fillna(-1)
    df_agg["Median_sort"] = df_agg["Median iron [µg/L]"].fillna(-np.inf)
    df_agg["N_samples_sort"] = df_agg["N samples"].fillna(0).astype(int)

    # Station sort key numeric if possible else string
    def station_sort_key_wrap(s):
        if s == "" or s is None:
            return ""
        try:
            return int(s)
        except Exception:
            return str(s)

    df_agg["Station_sort_key"] = df_agg["Station number"].apply(station_sort_key_wrap)

    df_sorted = df_agg.sort_values(
        by=["Exceed_sort", "Median_sort", "N_samples_sort", "Station_sort_key"],
        ascending=[False, False, False, True]
    ).reset_index(drop=True)

    # Take top 10
    df_top10 = df_sorted.head(10).copy()

    # Round numeric columns: percent 1 decimal, median/p95/max 2 decimals
    df_top10["% exceedances [%]"] = df_top10["% exceedances [%]"].round(1)
    df_top10["Median iron [µg/L]"] = df_top10["Median iron [µg/L]"].round(2)
    df_top10["95th percentile iron [µg/L]"] = df_top10["95th percentile iron [µg/L]"].round(2)
    df_top10["Max iron [µg/L]"] = df_top10["Max iron [µg/L]"].round(2)

    # Final columns order and exact headings as required
    final_cols = [
        "Station number",
        "Station name",
        "N samples",
        "N exceedances",
        "% exceedances [%]",
        "Median iron [µg/L]",
        "95th percentile iron [µg/L]",
        "Max iron [µg/L]"
    ]
    df_out = df_top10[final_cols].copy()

    # Save CSV
    df_out.to_csv(out_csv, index=False, encoding="utf-8")

    # Print created filenames per instructions
    print(py_filename)
    print(out_csv)

if __name__ == "__main__":
    main()