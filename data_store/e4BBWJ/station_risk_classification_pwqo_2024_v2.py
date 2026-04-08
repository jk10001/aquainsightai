# filename: station_risk_classification_pwqo_2024_v2.py
import pandas as pd
import os
import sys

def main():
    # Input filenames (must exist in current directory)
    excel_file = "Ontario_PWQMN_2024.xlsx"
    orig_csv = "station_risk_classification_pwqo_2024.csv"

    # Output filename (incremented version)
    out_csv = "station_risk_classification_pwqo_2024_v2.csv"
    py_filename = os.path.basename(__file__)

    # 1) Read Excel sheets with station IDs as strings
    # Use dtype=str to keep station IDs as strings (prevent integer coercion)
    try:
        df_data = pd.read_excel(excel_file, sheet_name="Data", dtype=str, engine="openpyxl")
    except Exception as e:
        print(f"ERROR reading Data sheet from '{excel_file}': {e}", file=sys.stderr)
        raise

    try:
        df_stations = pd.read_excel(excel_file, sheet_name="Stations", dtype=str, engine="openpyxl")
    except Exception as e:
        print(f"ERROR reading Stations sheet from '{excel_file}': {e}", file=sys.stderr)
        raise

    # Ensure the expected columns exist
    if "Collection Site" not in df_data.columns:
        raise KeyError("Expected 'Collection Site' column not found in Data sheet.")
    if "STATION" not in df_stations.columns:
        raise KeyError("Expected 'STATION' column not found in Stations sheet.")
    if "NAME" not in df_stations.columns:
        raise KeyError("Expected 'NAME' column not found in Stations sheet.")

    # 2) Convert station IDs to strings and left-pad to 11 characters
    # For Data sheet Collection Site
    df_data["Collection Site"] = df_data["Collection Site"].fillna("").astype(str).str.strip()
    # Replace any occurrence of 'nan' string (if created) with empty string
    df_data.loc[df_data["Collection Site"].str.lower() == "nan", "Collection Site"] = ""
    df_data["data_padded_id"] = df_data["Collection Site"].apply(lambda x: x.zfill(11) if x != "" else "")

    # For Stations sheet STATION
    df_stations["STATION"] = df_stations["STATION"].fillna("").astype(str).str.strip()
    df_stations.loc[df_stations["STATION"].str.lower() == "nan", "STATION"] = ""
    df_stations["station_padded_id"] = df_stations["STATION"].apply(lambda x: x.zfill(11) if x != "" else "")

    # 3) Read original station-level CSV (table) and treat Station number as string; pad to 11 characters
    try:
        df_orig = pd.read_csv(orig_csv, dtype=str)
    except Exception as e:
        print(f"ERROR reading original CSV '{orig_csv}': {e}", file=sys.stderr)
        raise

    # Ensure 'Station number' column exists
    if "Station number" not in df_orig.columns:
        raise KeyError("Expected 'Station number' column not found in original CSV.")

    # Normalize Station number to string and pad
    df_orig["Station number"] = df_orig["Station number"].fillna("").astype(str).str.strip()
    df_orig.loc[df_orig["Station number"].str.lower() == "nan", "Station number"] = ""
    df_orig["station_padded_id"] = df_orig["Station number"].apply(lambda x: x.zfill(11) if x != "" else "")

    # 4) Merge to bring in station NAME from Stations via padded key
    # Use left join so we preserve all rows from the original table
    df_merged = df_orig.merge(
        df_stations[["station_padded_id", "NAME"]],
        left_on="station_padded_id",
        right_on="station_padded_id",
        how="left",
        suffixes=("", "_stations")
    )

    # 5) Populate Station name column (replace or fill original 'Station name' if present)
    # If original has 'Station name' column, overwrite it with matched NAME; otherwise create it.
    if "Station name" in df_merged.columns:
        # Replace blanks or existing values with NAME where available
        df_merged["Station name"] = df_merged["NAME"].where(df_merged["NAME"].notna(), df_merged["Station name"])
    else:
        df_merged["Station name"] = df_merged["NAME"]

    # 6) For any padded station ID with no match in Stations, set explicit Unknown text
    mask_no_match = df_merged["Station name"].isna() | (df_merged["Station name"].astype(str).str.strip() == "")
    df_merged.loc[mask_no_match, "Station name"] = "Unknown (not in Stations sheet)"

    # 7) Add QA boolean column 'Station name matched' = True/False
    # Consider matched if the padded id exists in Stations and NAME is non-empty (i.e., not the Unknown string)
    df_merged["Station name matched"] = (~mask_no_match)

    # 8) Enforce Station number displayed as the 11-character zero-padded ID
    # Replace Station number with padded key (but keep blanks as-is)
    df_merged["Station number"] = df_merged["station_padded_id"].replace("", pd.NA)
    # For any NaN (originally blank), leave as blank string to mimic original formatting
    df_merged["Station number"] = df_merged["Station number"].fillna("")

    # 9) Prepare final dataframe columns ordering: keep original order, but ensure Station name and Station number are correct and add QA column
    final_cols = list(df_orig.columns)  # original columns as read
    # If original columns included 'Station name' ensure its in final_cols; otherwise prepend
    if "Station name" not in final_cols:
        # original CSV probably had Station name first; we'll insert it at the front
        final_cols.insert(0, "Station name")
    # Ensure 'Station number' is present and reflect updated padded numbers
    if "Station number" not in final_cols:
        final_cols.insert(1, "Station number")
    # Add 'Station name matched' as last column for QA visibility (if not already present)
    if "Station name matched" in df_merged.columns:
        if "Station name matched" in final_cols:
            pass
        else:
            final_cols.append("Station name matched")

    # Some columns may not be present in df_merged (e.g., if original CSV had unexpected formatting)
    # Build final frame using intersection while preserving order
    final_cols_in_df = [c for c in final_cols if c in df_merged.columns]

    df_final = df_merged[final_cols_in_df].copy()

    # 10) Save revised CSV
    df_final.to_csv(out_csv, index=False, encoding="utf-8")

    # Print created filenames and python filename to the terminal
    print(py_filename)
    print(out_csv)

if __name__ == "__main__":
    main()