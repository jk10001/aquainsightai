# filename: top10_stations_total_phosphorus_exceedance_2024.py
import os
import sys
import pandas as pd
import numpy as np

def parse_result_to_numeric(res_str):
    """
    Parse Results text to numeric. If it starts with '<', return ('censored', detection_limit_float)
    Else return ('measured', value_float)
    """
    if pd.isna(res_str):
        return (None, np.nan)
    s = str(res_str).strip()
    if s == "":
        return (None, np.nan)
    # Some values may include qualifier symbols like '<' or '<='; treat any leading '<' (possibly with '=') as censored
    if s.startswith("<"):
        # remove leading '<' and any '=' or spaces
        numpart = s.lstrip("<=").strip()
        try:
            val = float(numpart)
            return ("censored", val)
        except:
            return ("censored", np.nan)
    else:
        # plain numeric, possibly with commas
        try:
            return ("measured", float(s.replace(",", "")))
        except:
            # try to remove any non-numeric trailing characters
            import re
            m = re.search(r"[-+]?\d*\.?\d+([eE][-+]?\d+)?", s)
            if m:
                try:
                    return ("measured", float(m.group(0)))
                except:
                    return (None, np.nan)
            return (None, np.nan)

def main():
    py_filename = os.path.basename(__file__)
    excel_file = "Ontario_PWQMN_2024.xlsx"
    out_csv = "top10_stations_total_phosphorus_exceedance_2024.csv"

    # Read sheets
    try:
        df = pd.read_excel(excel_file, sheet_name="Data", dtype=str, engine="openpyxl")
    except Exception as e:
        print(f"ERROR reading 'Data' sheet from '{excel_file}': {e}", file=sys.stderr)
        raise

    try:
        df_stations = pd.read_excel(excel_file, sheet_name="Stations", dtype=str, engine="openpyxl")
    except Exception as e:
        print(f"ERROR reading 'Stations' sheet from '{excel_file}': {e}", file=sys.stderr)
        raise

    # Normalize column names
    df.columns = [c.strip() for c in df.columns]
    df_stations.columns = [c.strip() for c in df_stations.columns]

    # Ensure necessary columns exist
    for col in ["Collection Site", "Analyte", "Results", "Units"]:
        if col not in df.columns:
            raise KeyError(f"Expected column '{col}' in Data sheet.")

    for col in ["STATION", "NAME"]:
        if col not in df_stations.columns:
            raise KeyError(f"Expected column '{col}' in Stations sheet.")

    # Filter to Total Phosphorus analyte
    df_tp = df[df["Analyte"].astype(str).str.strip().str.lower() == "phosphorus; total".lower()].copy()

    # Ensure Collection Site is string and zero-pad to 11 characters
    df_tp["Collection Site"] = df_tp["Collection Site"].fillna("").astype(str).str.strip()
    # Replace 'nan' strings if present
    df_tp.loc[df_tp["Collection Site"].str.lower() == "nan", "Collection Site"] = ""
    df_tp["station_padded_id"] = df_tp["Collection Site"].apply(lambda x: x.zfill(11) if x != "" else "")

    # Prepare Stations padded id and name mapping
    df_stations["STATION"] = df_stations["STATION"].fillna("").astype(str).str.strip()
    df_stations.loc[df_stations["STATION"].str.lower() == "nan", "STATION"] = ""
    df_stations["station_padded_id"] = df_stations["STATION"].apply(lambda x: x.zfill(11) if x != "" else "")
    df_stations["NAME"] = df_stations["NAME"].fillna("").astype(str)

    stations_map = df_stations.set_index("station_padded_id")["NAME"].to_dict()

    # Parse Results into numeric TP_ugL following rules:
    # - If Results begins with '<', parse DL and set TP_ugL = 0.5 * DL
    # - Else parse as float
    # - Convert mg/L to ug/L if Units indicates mg/L
    results_parsed = []
    tp_ugL_list = []
    parsed_flags = []
    for idx, row in df_tp.iterrows():
        res = row.get("Results", "")
        units = row.get("Units", "")
        kind, val = parse_result_to_numeric(res)
        # Determine units: treat None/NaN as assume ug/L if units missing? But we will attempt to detect mg/L explicitly
        units_str = "" if pd.isna(units) else str(units).strip().lower()
        # Normalize some unit variants
        units_str = units_str.replace("µ", "u").replace("μ", "u")
        # handle mg/L and ug/L detection
        is_mg_per_l = False
        if "mg/l" in units_str or units_str.startswith("mg/"):
            is_mg_per_l = True
        # Some dataset uses 'mg/L as CaCO3' etc; handle by checking 'mg/l' substring
        if "mg/l" in units_str:
            is_mg_per_l = True
        # Also handle uppercase variants
        if units_str == "mg/l" or units_str == "mg/l as caco3" or units_str.startswith("mg/"):
            is_mg_per_l = True

        tp_ugL = np.nan
        parsed_note = ""
        if kind == "censored":
            # val is detection limit in the units reported
            if np.isnan(val):
                tp_ugL = np.nan
                parsed_note = "censored_missing_DL"
            else:
                # DL is in units reported; convert to ug/L if mg/L
                dl = float(val)
                if is_mg_per_l:
                    dl_ugL = dl * 1000.0
                else:
                    # assume ug/L if units indicate ug/L or units blank
                    dl_ugL = dl
                tp_ugL = 0.5 * dl_ugL
                parsed_note = "censored_0.5DL"
        elif kind == "measured":
            if np.isnan(val):
                tp_ugL = np.nan
                parsed_note = "measured_missing"
            else:
                v = float(val)
                if is_mg_per_l:
                    tp_ugL = v * 1000.0
                else:
                    tp_ugL = v
                parsed_note = "measured"
        else:
            tp_ugL = np.nan
            parsed_note = "unparsed"

        results_parsed.append((kind, val, units_str, parsed_note))
        tp_ugL_list.append(tp_ugL)
        parsed_flags.append(parsed_note)

    df_tp["TP_ugL"] = tp_ugL_list
    df_tp["parse_note"] = parsed_flags

    # Drop rows with TP_ugL NaN (unable to parse) from analysis
    df_tp_clean = df_tp[~df_tp["TP_ugL"].isna()].copy()

    # For safety, convert station padded id blank to explicit empty string
    df_tp_clean["station_padded_id"] = df_tp_clean["station_padded_id"].fillna("").astype(str)

    # Aggregate per station
    def compute_stats(group):
        vals = group["TP_ugL"].values.astype(float)
        n = int(len(vals))
        n_exceed = int((vals > 30.0).sum())
        pct_exceed = 100.0 * n_exceed / n if n > 0 else 0.0
        median = float(np.nanmedian(vals)) if n > 0 else np.nan
        p95 = float(np.nanpercentile(vals, 95)) if n > 0 else np.nan
        mx = float(np.nanmax(vals)) if n > 0 else np.nan
        return pd.Series({
            "n_samples": n,
            "n_exceed": n_exceed,
            "%_exceed": pct_exceed,
            "median_ugL": median,
            "p95_ugL": p95,
            "max_ugL": mx
        })

    grouped = df_tp_clean.groupby("station_padded_id", dropna=False).apply(compute_stats).reset_index()

    # Attach station names; default to "Unknown (not in Stations sheet)"
    def get_station_name(padded_id):
        if padded_id in stations_map and str(stations_map[padded_id]).strip() != "":
            return stations_map[padded_id]
        else:
            return "Unknown (not in Stations sheet)"

    grouped["Station number"] = grouped["station_padded_id"].apply(lambda x: x if (isinstance(x, str) and x.strip() != "") else "")
    grouped["Station name"] = grouped["station_padded_id"].apply(get_station_name)

    # Reorder and format columns; include units in headings where appropriate
    final_df = grouped[[
        "Station number",
        "Station name",
        "n_samples",
        "n_exceed",
        "%_exceed",
        "median_ugL",
        "p95_ugL",
        "max_ugL"
    ]].copy()

    # Round numeric columns sensibly
    final_df["%_exceed"] = final_df["%_exceed"].round(2)
    final_df["median_ugL"] = final_df["median_ugL"].round(3)
    final_df["p95_ugL"] = final_df["p95_ugL"].round(3)
    final_df["max_ugL"] = final_df["max_ugL"].round(3)

    # Ranking: primary by % exceed descending, secondary by median descending
    final_df_sorted = final_df.sort_values(by=["%_exceed", "median_ugL"], ascending=[False, False]).reset_index(drop=True)

    # Select top 10
    top10 = final_df_sorted.head(10).copy()

    # Prepare CSV headings with units in english
    csv_df = top10.rename(columns={
        "Station number": "Station number",
        "Station name": "Station name",
        "n_samples": "N samples",
        "n_exceed": "N exceedances",
        "%_exceed": "% exceedances [%]",
        "median_ugL": "Median TP [µg/L]",
        "p95_ugL": "95th percentile TP [µg/L]",
        "max_ugL": "Max TP [µg/L]"
    })

    # Save only the final processed table to CSV (no index)
    csv_df.to_csv(out_csv, index=False, encoding="utf-8")

    # Print created filenames and python filename to terminal
    print(py_filename)
    print(out_csv)

if __name__ == "__main__":
    main()