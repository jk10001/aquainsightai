# filename: top10_tp_exceedance_by_station_2024.py
import os
import sys
import re
import pandas as pd
import numpy as np
import tempfile
import shutil

def parse_result_to_numeric(res):
    """
    Parse a Results field into numeric value (float) and a flag if it was BDL.
    BDL format like '<2' or '< 2' -> return 0.5 * 2 and bdl=True
    Otherwise try float conversion -> bdl=False
    If cannot parse, return np.nan, False
    """
    if pd.isna(res):
        return np.nan, False
    s = str(res).strip()
    # Detect leading '<' (maybe with spaces)
    m = re.match(r'^\s*<\s*([0-9,+\-eE.]+)\s*$', s)
    if m:
        try:
            x = float(m.group(1).replace(',', ''))
            return 0.5 * x, True
        except:
            return np.nan, True
    # Otherwise try plain numeric
    try:
        s2 = s.replace(',', '')
        val = float(s2)
        return val, False
    except:
        return np.nan, False

def safe_write_csv(df, path, fallback_suffix="_fallback"):
    """
    Try to write df to 'path'. If PermissionError occurs, attempt to write to a temporary file
    and then move it into place. If still not permitted, write to a fallback file in the current
    directory with suffix and return the actual path written.
    Returns the actual path written.
    """
    # First try direct write
    try:
        df.to_csv(path, index=False, encoding="utf-8")
        return path
    except PermissionError:
        # Try writing to a tempfile in the same directory
        dirpath = os.path.dirname(os.path.abspath(path)) or "."
        try:
            fd, tmpname = tempfile.mkstemp(prefix=".tmp_top10_", dir=dirpath, suffix=".csv")
            os.close(fd)
            df.to_csv(tmpname, index=False, encoding="utf-8")
            # Attempt to replace target (may still fail)
            try:
                os.replace(tmpname, path)
                return path
            except PermissionError:
                # cannot replace; move to fallback
                fallback = os.path.splitext(path)[0] + fallback_suffix + ".csv"
                try:
                    shutil.move(tmpname, fallback)
                    return fallback
                except Exception:
                    # last resort: write to /tmp
                    fallback2 = os.path.join(tempfile.gettempdir(), os.path.basename(path))
                    try:
                        df.to_csv(fallback2, index=False, encoding="utf-8")
                        return fallback2
                    except Exception as e:
                        raise PermissionError(f"Unable to write CSV to {path} or fallback locations: {e}") from e
        except Exception:
            # fallback to writing in current directory with suffix
            fallback = os.path.splitext(path)[0] + fallback_suffix + ".csv"
            try:
                df.to_csv(fallback, index=False, encoding="utf-8")
                return fallback
            except Exception as e:
                raise PermissionError(f"Unable to write CSV to {path} or fallback '{fallback}': {e}") from e
    except Exception as e:
        # Other exceptions propagate
        raise

def pad_station_id_raw(s):
    """
    Given a raw station identifier (from Collection Site or STATION),
    extract digits and return an 11-character zero-padded string.
    If digits cannot be found, return the original stripped string zfilled/truncated to 11.
    Always returns a string of length 11 (digits + leading zeros or fallback).
    """
    if pd.isna(s):
        return "".zfill(11)
    s_str = str(s).strip()
    if s_str == "" or s_str.lower() == "nan":
        return "".zfill(11)
    # Extract digits
    digits = re.sub(r'\D', '', s_str)
    if digits:
        return digits.zfill(11)
    # Fallback: if contains letters, take ascii codes? Simpler: right-pad/truncate the cleaned string
    clean = re.sub(r'\s+', '', s_str)
    # Convert each char to its ordinal modulo 10 to form digits? Too complex; instead use numeric hash fallback
    # But requirement is to zero-pad numeric station IDs; non-numeric should be rare; we'll right-pad/truncate
    if len(clean) >= 11:
        return clean[:11]
    else:
        return clean.rjust(11, "0")

def finalize_station_num(s):
    """
    Ensure final Station number column is an 11-character zero-padded string.
    If input already 11 digit-like, keep it. If blank, return 11 zeros string (but user expects padded IDs; blank
    should be all zeros to enforce length). We will produce 11-digit zero-padded strings for any numeric IDs.
    """
    if pd.isna(s):
        return "".zfill(11)
    s_str = str(s).strip()
    if s_str == "" or s_str.lower() == "nan":
        return "".zfill(11)
    digits = re.sub(r'\D', '', s_str)
    if digits:
        return digits.zfill(11)
    # otherwise, pad/truncate
    if len(s_str) >= 11:
        return s_str[:11]
    return s_str.zfill(11)

def main():
    py_filename = os.path.basename(__file__)
    excel_file = "Ontario_PWQMN_2024.xlsx"
    desired_out = "top10_tp_exceedance_by_station_2024.csv"

    # Read sheets
    try:
        df = pd.read_excel(excel_file, sheet_name="Data", engine="openpyxl", dtype=str)
    except Exception as e:
        print(f"ERROR reading Data sheet from '{excel_file}': {e}", file=sys.stderr)
        raise

    try:
        df_stations = pd.read_excel(excel_file, sheet_name="Stations", engine="openpyxl", dtype=str)
    except Exception as e:
        print(f"ERROR reading Stations sheet from '{excel_file}': {e}", file=sys.stderr)
        raise

    # Normalize column names
    df.columns = [c.strip() for c in df.columns]
    df_stations.columns = [c.strip() for c in df_stations.columns]

    # Validate required columns
    for c in ["Year", "Collection Site", "Analyte", "Results", "Units"]:
        if c not in df.columns:
            raise KeyError(f"Expected column '{c}' in Data sheet.")
    for c in ["STATION", "NAME"]:
        if c not in df_stations.columns:
            raise KeyError(f"Expected column '{c}' in Stations sheet.")

    # Filter to Total Phosphorus and Year 2024
    df = df[df["Analyte"].astype(str).str.strip() == "Phosphorus; total"].copy()
    df = df[df["Year"].astype(str).str.strip() == "2024"].copy()

    # Normalize Units and keep only ug/L or mg/L (normalize variants)
    df["Units_norm"] = df["Units"].fillna("").astype(str).str.strip().str.lower()
    df["Units_norm"] = df["Units_norm"].replace({
        "µg/l": "ug/l",
        "μg/l": "ug/l",
        "ug/l.": "ug/l",
        "ug/l ": "ug/l",
        "μg/l ": "ug/l",
        "µg/l ": "ug/l"
    })
    df = df[df["Units_norm"].isin(["ug/l", "mg/l"])].copy()

    # Parse Results to numeric (apply BDL handling)
    parsed = df["Results"].apply(lambda x: parse_result_to_numeric(x))
    df["value_raw"] = parsed.apply(lambda t: t[0])
    df["was_bdl"] = parsed.apply(lambda t: t[1])

    # Convert mg/L to ug/L when needed
    mg_mask = df["Units_norm"] == "mg/l"
    if mg_mask.any():
        df.loc[mg_mask, "value_raw"] = pd.to_numeric(df.loc[mg_mask, "value_raw"], errors="coerce") * 1000.0

    df["value_numeric"] = pd.to_numeric(df["value_raw"], errors="coerce")

    # Drop rows with NaN numeric values (cannot be used)
    df = df[~df["value_numeric"].isna()].copy()

    # Exceedance flag (>30 µg/L)
    df["exceed"] = df["value_numeric"] > 30.0

    # Normalize station ID to 11-digit zero-padded string for Collection Site
    df["Collection Site"] = df["Collection Site"].fillna("").astype(str).str.strip()
    df["station_padded_id"] = df["Collection Site"].apply(pad_station_id_raw)

    # Also ensure Stations sheet station IDs are padded to 11 for reliable join
    df_stations["STATION"] = df_stations["STATION"].fillna("").astype(str).str.strip()
    df_stations["station_padded_id"] = df_stations["STATION"].apply(lambda x: pad_station_id_raw(x))

    # Group by station and compute stats
    group = df.groupby("station_padded_id", dropna=False)

    summary = group["value_numeric"].agg(
        N_samples = "count",
        Median = lambda x: float(np.nanmedian(x)) if x.size>0 else np.nan,
        P95 = lambda x: float(np.nanquantile(x, 0.95)) if x.size>0 else np.nan,
        Max = lambda x: float(np.nanmax(x)) if x.size>0 else np.nan
    ).reset_index()

    exceed_counts = group["exceed"].sum().reset_index().rename(columns={"exceed":"N_exceedances"})
    summary = summary.merge(exceed_counts, on="station_padded_id", how="left")

    summary["% exceedances"] = (summary["N_exceedances"] / summary["N_samples"]) * 100.0
    summary["% exceedances"] = summary["% exceedances"].fillna(0.0)

    # Prepare Stations mapping (ensure padded IDs)
    stations_map = df_stations[["station_padded_id", "NAME"]].rename(columns={"NAME":"Station name"})

    summary = summary.merge(stations_map, on="station_padded_id", how="left")
    missing_mask = summary["Station name"].isna() | (summary["Station name"].astype(str).str.strip() == "")
    summary.loc[missing_mask, "Station name"] = "Unknown (not in Stations sheet)"

    # Finalize Station number string - ensure 11-digit zero-padded for every row
    summary["Station number"] = summary["station_padded_id"].apply(finalize_station_num)

    # Build output dataframe with required columns/order and formatting
    out_df = pd.DataFrame({
        "Station number": summary["Station number"].astype(str),
        "Station name": summary["Station name"].astype(str),
        "N samples": summary["N_samples"].astype(int),
        "N exceedances": summary["N_exceedances"].astype(int),
        "% exceedances": summary["% exceedances"].round(2),
        "Median (µg/L)": summary["Median"].round(3),
        "95th percentile (µg/L)": summary["P95"].round(3),
        "Max (µg/L)": summary["Max"].round(3)
    })

    # Sort and pick top 10 (ranking: % exceed desc, Median desc, N samples desc)
    out_df_sorted = out_df.sort_values(
        by=["% exceedances", "Median (µg/L)", "N samples"],
        ascending=[False, False, False]
    ).reset_index(drop=True)

    top10 = out_df_sorted.head(10).copy()

    # Ensure Station number column is exactly 11 characters for every row
    top10["Station number"] = top10["Station number"].apply(lambda x: finalize_station_num(x))

    # Try to write CSV; if permission problems occur, safe_write_csv will attempt fallbacks
    try:
        actual_written = safe_write_csv(top10, desired_out)
    except Exception as e:
        print(f"ERROR writing output CSV: {e}", file=sys.stderr)
        raise

    # Print created filenames as required
    print(py_filename)
    print(actual_written)

if __name__ == "__main__":
    main()