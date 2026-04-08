# filename: top10_tp_exceedance_pwqo_2024.py
import os
import sys
import pandas as pd
import numpy as np
import re

def parse_result_value(result_str, detection_limit):
    """
    Parse Results string and return numeric value per rules:
    - If result like '<X' or Value Qualifier indicates '<' treat as non-detect:
        return 0.5 * DL if DL available, else 0.5 * parsed X from '<X'
    - Else parse as float
    Returns (numeric_value, is_nd)
    """
    if pd.isna(result_str):
        return (np.nan, False)
    s = str(result_str).strip()
    # Common non-detect formats start with '<'
    if s.startswith('<'):
        # Try detection limit first
        if not pd.isna(detection_limit):
            try:
                dl = float(detection_limit)
                return (0.5 * dl, True)
            except Exception:
                pass
        # else parse the number after '<'
        m = re.search(r'<\s*([0-9.+-eE]+)', s)
        if m:
            try:
                val = float(m.group(1))
                return (0.5 * val, True)
            except Exception:
                return (np.nan, True)
        return (np.nan, True)
    # Sometimes results contain inequality in Value Qualifier rather than Results string,
    # but this function only handles Results text. Non-numeric characters removed if possible.
    # Try direct float conversion
    # Remove commas and any spaces
    s_clean = s.replace(',', '').strip()
    # If contains spaces with units appended, remove trailing non-numeric
    m = re.match(r'^([0-9.+-eE]+)', s_clean)
    if m:
        try:
            return (float(m.group(1)), False)
        except Exception:
            return (np.nan, False)
    # fallback: try to extract number anywhere
    m2 = re.search(r'([0-9]+(?:\.[0-9]*)?(?:[eE][+-]?\d+)?)', s_clean)
    if m2:
        try:
            return (float(m2.group(1)), False)
        except Exception:
            return (np.nan, False)
    return (np.nan, False)

def main():
    excel_file = "Ontario_PWQMN_2024.xlsx"
    out_csv = "top10_tp_exceedance_pwqo_2024.csv"
    py_filename = os.path.basename(__file__)

    # 1) Read sheets
    try:
        df = pd.read_excel(excel_file, sheet_name="Data", engine="openpyxl", dtype=object)
    except Exception as e:
        print(f"ERROR reading Data sheet from '{excel_file}': {e}", file=sys.stderr)
        raise

    try:
        df_stations = pd.read_excel(excel_file, sheet_name="Stations", engine="openpyxl", dtype=object)
    except Exception as e:
        print(f"ERROR reading Stations sheet from '{excel_file}': {e}", file=sys.stderr)
        raise

    # Standardize column names
    df.columns = [c.strip() for c in df.columns]
    df_stations.columns = [c.strip() for c in df_stations.columns]

    # Validate required columns
    required_cols = ["Year", "Collection Site", "Analyte", "Results", "Units", "Result Call", "Detection Limit", "Value Qualifier"]
    for c in required_cols:
        if c not in df.columns:
            raise KeyError(f"Expected column '{c}' not found in Data sheet.")

    for c in ["STATION", "NAME"]:
        if c not in df_stations.columns:
            raise KeyError(f"Expected column '{c}' not found in Stations sheet.")

    # 2) Filter for Year==2024 and Analyte == 'Phosphorus; total'
    df_filtered = df.copy()
    # Ensure Year numeric if possible
    df_filtered['Year'] = pd.to_numeric(df_filtered['Year'], errors='coerce')
    df_filtered = df_filtered[df_filtered['Year'] == 2024]
    df_filtered = df_filtered[df_filtered['Analyte'].astype(str).str.strip() == "Phosphorus; total"].copy()

    # If no rows, exit with empty CSV
    if df_filtered.shape[0] == 0:
        df_out_empty = pd.DataFrame(columns=[
            "Station number", "Station name", "Station name matched",
            "N samples", "N exceedances", "% exceedances [%]",
            "Median [ug/L]", "95th percentile [ug/L]", "Max [ug/L]"
        ])
        df_out_empty.to_csv(out_csv, index=False)
        print(py_filename)
        print(out_csv)
        return

    # 3) Normalize station IDs: treat Collection Site as string and zero-pad to 11 chars
    df_filtered['Collection Site'] = df_filtered['Collection Site'].astype(str).str.strip()
    # Convert possible NaN string to empty
    df_filtered.loc[df_filtered['Collection Site'].str.lower() == 'nan', 'Collection Site'] = ''
    df_filtered['station_padded_id'] = df_filtered['Collection Site'].apply(lambda x: x.zfill(11) if x != '' else '')

    # 4) Units handling: expected µg/L; convert mg/L -> µg/L (multiply by 1000)
    # Normalize Units strings
    df_filtered['Units'] = df_filtered['Units'].astype(str).fillna('').str.strip()
    # Map common variants to canonical
    def canonical_unit(u):
        u = str(u).strip()
        if u in ['ug/L', 'µg/L', 'ug/L ', 'ug/L']:
            return 'ug/L'
        if u.lower() in ['mg/l', 'mg/l ', 'mg/l as caco3', 'mg/l as caco3', 'mg/l as o2', 'mg/l as caCO3'.lower()]:
            # We'll handle mg/L generically
            return 'mg/L'
        if 'mg' in u and '/l' in u.lower():
            return 'mg/L'
        if 'ug' in u and '/l' in u.lower():
            return 'ug/L'
        # blank or unknown
        return u

    df_filtered['Units_canonical'] = df_filtered['Units'].apply(canonical_unit)

    # 5) Parse Detection Limit as numeric where present
    df_filtered['Detection Limit numeric'] = pd.to_numeric(df_filtered['Detection Limit'], errors='coerce')

    # 6) Parse Results per rules into TP_result_ugL
    results_parsed = []
    nd_flags = []
    for idx, row in df_filtered.iterrows():
        res = row.get('Results', '')
        dl = row.get('Detection Limit numeric', np.nan)
        val, is_nd = parse_result_value(res, dl)
        # If the row has Result Call 'BDL' or Value Qualifier contains '<', treat as non-detect.
        result_call = str(row.get('Result Call', '')).strip().upper()
        val_qual = str(row.get('Value Qualifier', '')).strip()
        if result_call == 'BDL' and not is_nd:
            # mark as non-detect: if DL available use 0.5*DL, else try to parse number in Results
            if not pd.isna(dl):
                val = 0.5 * float(dl)
            else:
                # try to parse numeric in Results or set NaN
                val2, _ = parse_result_value(res, np.nan)
                if not pd.isna(val2):
                    val = 0.5 * val2
                else:
                    val = np.nan
            is_nd = True
        if val_qual == '<' and not is_nd:
            # Value Qualifier indicates '<' even if Results didn't show it
            if not pd.isna(dl):
                val = 0.5 * float(dl)
            else:
                # try to parse numeric in Results
                m = re.search(r'([0-9]+(?:\.[0-9]*)?)', str(res))
                if m:
                    try:
                        val = 0.5 * float(m.group(1))
                    except:
                        val = np.nan
                else:
                    val = np.nan
            is_nd = True

        # Units conversion to ug/L if necessary
        unit = row.get('Units_canonical', '')
        if unit == 'mg/L':
            # convert mg/L to ug/L
            if not pd.isna(val):
                val = float(val) * 1000.0
        # If unit is not ug/L or mg/L, we will exclude later (set NaN) — but many rows may have missing units; assume ug/L if units blank?
        # Per instructions: ensure expected units µg/L; if other unit appears, convert mg/L->µg/L or exclude.
        if unit not in ('ug/L', 'mg/L', '') and unit is not None:
            # Unknown unit; set NaN to exclude
            val = np.nan

        results_parsed.append(val)
        nd_flags.append(bool(is_nd))

    df_filtered['TP_result_ugL'] = results_parsed
    df_filtered['TP_non_detect'] = nd_flags

    # 7) For any rows where Units were blank, check if numeric Results likely in mg/L or ug/L
    # Heuristic: if TP_result_ugL is NaN but Results string contains decimal and small numbers, assume ug/L.
    # However, to be conservative, if Units blank we will assume ug/L if value parsed < 1000; if >1000 maybe mg/L? This is risky.
    # Instead, only fill when parse succeeded but unit blank: treat as ug/L (most TP entries expected ug/L).
    mask_unit_blank = (df_filtered['Units_canonical'] == '') & (~df_filtered['TP_result_ugL'].isna())
    # treat as ug/L (no conversion)
    # Nothing to do since parsed value assumed numeric in Results was already in whatever units; we assume ug/L here.

    # 8) Drop rows with TP_result_ugL NaN (unable to parse or unknown units)
    before_drop = df_filtered.shape[0]
    df_filtered = df_filtered[~df_filtered['TP_result_ugL'].isna()].copy()
    after_drop = df_filtered.shape[0]
    # (We will not print these counts per instructions, but kept variables for debugging if needed)

    # 9) Exceedance flag per sample: > 30 ug/L
    df_filtered['TP_exceed'] = df_filtered['TP_result_ugL'].astype(float) > 30.0

    # 10) Aggregate per station (use padded station id as Station number)
    agg = df_filtered.groupby('station_padded_id').agg(
        N_samples = ('TP_result_ugL', 'count'),
        N_exceedances = ('TP_exceed', 'sum'),
        Median_ugL = ('TP_result_ugL', lambda x: float(np.nanmedian(x.values)) if x.notna().any() else np.nan),
        P95_ugL = ('TP_result_ugL', lambda x: float(np.nanpercentile(x.values, 95)) if x.notna().any() else np.nan),
        Max_ugL = ('TP_result_ugL', lambda x: float(np.nanmax(x.values)) if x.notna().any() else np.nan)
    ).reset_index().rename(columns={'station_padded_id':'Station number'})

    # Compute percent exceedances
    agg['% exceedances [%]'] = 100.0 * agg['N_exceedances'] / agg['N_samples']

    # 11) Join station names from Stations sheet
    # Prepare stations padded IDs
    df_stations['STATION'] = df_stations['STATION'].astype(str).str.strip()
    df_stations.loc[df_stations['STATION'].str.lower() == 'nan', 'STATION'] = ''
    df_stations['station_padded_id'] = df_stations['STATION'].apply(lambda x: x.zfill(11) if x != '' else '')
    # Keep mapping
    station_map = df_stations.set_index('station_padded_id')['NAME'].to_dict()

    def lookup_name(padded_id):
        if padded_id in station_map and pd.notna(station_map[padded_id]) and str(station_map[padded_id]).strip() != '':
            return (str(station_map[padded_id]), True)
        else:
            return ("Unknown (not in Stations sheet)", False)

    names = [lookup_name(s) for s in agg['Station number'].astype(str).tolist()]
    agg['Station name'] = [n[0] for n in names]
    agg['Station name matched'] = [n[1] for n in names]

    # 12) Prepare final columns and apply ranking
    # Required output columns order:
    # 1. Station number
    # 2. Station name
    # 3. Station name matched
    # 4. N samples
    # 5. N exceedances
    # 6. % exceedances [%]
    # 7. Median [ug/L]
    # 8. 95th percentile [ug/L]
    # 9. Max [ug/L]
    agg_final = agg[[
        'Station number', 'Station name', 'Station name matched',
        'N_samples', 'N_exceedances', '% exceedances [%]',
        'Median_ugL', 'P95_ugL', 'Max_ugL'
    ]].copy()

    # Rename to match exact heading labels
    agg_final = agg_final.rename(columns={
        'N_samples': 'N samples',
        'N_exceedances': 'N exceedances',
        'Median_ugL': 'Median [ug/L]',
        'P95_ugL': '95th percentile [ug/L]',
        'Max_ugL': 'Max [ug/L]'
    })

    # Ensure Station number is string zero-padded 11 chars; some may be empty strings — keep as ''
    agg_final['Station number'] = agg_final['Station number'].astype(str).apply(lambda x: x if x == '' else x.zfill(11))

    # Sorting:
    # - primary: descending % exceedances
    # - secondary: descending Median [ug/L]
    # - tie: descending N samples
    # - tie: ascending station number (string)
    agg_final['% exceedances [%]'] = pd.to_numeric(agg_final['% exceedances [%]'], errors='coerce').fillna(0)
    agg_final['Median [ug/L]'] = pd.to_numeric(agg_final['Median [ug/L]'], errors='coerce').fillna(-np.inf)
    agg_final['N samples'] = pd.to_numeric(agg_final['N samples'], errors='coerce').fillna(0).astype(int)

    agg_final_sorted = agg_final.sort_values(
        by=['% exceedances [%]', 'Median [ug/L]', 'N samples', 'Station number'],
        ascending=[False, False, False, True],
        kind='mergesort'  # stable
    ).reset_index(drop=True)

    # 13) Select top 10 stations
    top10 = agg_final_sorted.head(10).copy()

    # 14) Formatting numeric columns to reasonable precision
    top10['% exceedances [%]'] = top10['% exceedances [%]'].round(2)
    top10['Median [ug/L]'] = top10['Median [ug/L]'].round(3)
    top10['95th percentile [ug/L]'] = top10['95th percentile [ug/L]'].round(3)
    top10['Max [ug/L]'] = top10['Max [ug/L]'].round(3)

    # 15) Ensure columns order exactly as requested
    final_columns = [
        "Station number",
        "Station name",
        "Station name matched",
        "N samples",
        "N exceedances",
        "% exceedances [%]",
        "Median [ug/L]",
        "95th percentile [ug/L]",
        "Max [ug/L]"
    ]
    top10 = top10[final_columns]

    # 16) Save CSV (only these 10 rows)
    top10.to_csv(out_csv, index=False, encoding='utf-8')

    # Print created filenames and python filename to the terminal
    print(py_filename)
    print(out_csv)

if __name__ == "__main__":
    main()