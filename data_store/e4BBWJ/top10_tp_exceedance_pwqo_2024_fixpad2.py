# filename: top10_tp_exceedance_pwqo_2024_fixpad2.py
import os
import sys
import pandas as pd
import re

def normalize_to_11_digits(s):
    """
    Given an input (possibly containing digits and other chars), return:
    - '' if empty or all whitespace
    - an 11-digit string composed only of digits (zero-padded on the left) if digits can be found
    Rules:
    - Extract all digits from the input in order.
    - If no digits found -> return '' (to avoid non-digit station numbers).
    - If extracted digits length > 11 -> keep the RIGHTMOST 11 digits (to preserve trailing station id digits)
    - If extracted digits length < 11 -> left-pad with zeros to 11.
    """
    if s is None:
        return ""
    s_str = str(s).strip()
    if s_str == "":
        return ""
    # Remove trailing .0 if present due to float-to-string conversions
    if s_str.endswith(".0"):
        s_str = s_str[:-2]
    # Extract digits
    digits = "".join(re.findall(r"\d+", s_str))
    if digits == "":
        # No digits found -> return empty string to avoid non-digit station numbers
        return ""
    if len(digits) > 11:
        # Keep rightmost 11 digits
        digits = digits[-11:]
    if len(digits) < 11:
        digits = digits.zfill(11)
    return digits

def main():
    py_filename = os.path.basename(__file__)
    in_csv = "top10_tp_exceedance_pwqo_2024.csv"
    out_csv = in_csv  # overwrite as requested

    # 1) Read input CSV
    try:
        # Keep original formatting for non-numeric fields; do not convert empty strings to NaN
        df = pd.read_csv(in_csv, dtype=str, keep_default_na=False)
    except FileNotFoundError:
        print(f"ERROR: Input CSV '{in_csv}' not found in current directory.", file=sys.stderr)
        raise
    except Exception as e:
        print(f"ERROR reading '{in_csv}': {e}", file=sys.stderr)
        raise

    # 2) Validate expected column present
    expected_col = "Station number"
    if expected_col not in df.columns:
        raise KeyError(f"Expected column '{expected_col}' not found in '{in_csv}'.")

    # 3) Normalize/pad every Station number to exactly 11 digits (digits-only) or blank
    df_corrected = df.copy()
    df_corrected[expected_col] = df_corrected[expected_col].apply(normalize_to_11_digits)

    # 4) Verify all station numbers are either blank or exactly 11 digits
    def check_station_format(s):
        if s is None:
            return False
        s_str = str(s)
        if s_str == "":
            return True
        return bool(re.fullmatch(r"\d{11}", s_str))

    bad_mask = ~df_corrected[expected_col].apply(check_station_format)
    if bad_mask.any():
        # Should not happen, but if it does, raise an error listing offending rows
        bad_rows = df_corrected.loc[bad_mask, expected_col].tolist()
        raise ValueError(f"Post-processing validation failed for Station number values: {bad_rows}")

    # 5) Overwrite CSV with corrected station numbers
    try:
        df_corrected.to_csv(out_csv, index=False, encoding="utf-8")
    except Exception as e:
        print(f"ERROR writing '{out_csv}': {e}", file=sys.stderr)
        raise

    # 6) Print required filenames and preview of updated CSV (all rows)
    print(py_filename)
    print(out_csv)
    # Print CSV preview - show full content to confirm station numbers are padded
    # Read back (safe) and print as CSV-style lines
    try:
        df_check = pd.read_csv(out_csv, dtype=str, keep_default_na=False)
    except Exception as e:
        print(f"ERROR re-reading '{out_csv}' for preview: {e}", file=sys.stderr)
        raise

    # Print header
    header = list(df_check.columns)
    print(",".join(header))
    # Print all rows
    for _, row in df_check.iterrows():
        # ensure commas in fields don't break simple preview: escape by quoting fields with commas
        out_fields = []
        for col in header:
            val = "" if pd.isna(row[col]) else str(row[col])
            if "," in val or "\n" in val:
                # simple quoting
                val = f'"{val}"'
            out_fields.append(val)
        print(",".join(out_fields))

if __name__ == "__main__":
    main()