# filename: top10_tp_exceedance_pwqo_2024_fixpad.py
import os
import sys
import pandas as pd

def main():
    py_filename = os.path.basename(__file__)
    in_csv = "top10_tp_exceedance_pwqo_2024.csv"
    out_csv = in_csv  # overwrite as requested

    # 1) Read existing CSV
    try:
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

    # 3) Zero-pad station numbers to 11 characters.
    # Keep empty strings as-is (if any)
    def pad_station(s):
        if s is None:
            return ""
        s_str = str(s).strip()
        if s_str == "":
            return ""
        # Remove accidental decimal markers if present from numeric read (e.g., '4001303302.0')
        if s_str.endswith(".0"):
            s_str = s_str[:-2]
        # If contains non-digit characters, preserve them but try to extract digits for padding
        if not s_str.isdigit():
            # If string contains digits, extract contiguous digits and pad that; else leave as-is
            import re
            m = re.search(r'(\d+)', s_str)
            if m:
                digits = m.group(1).zfill(11)
                # Replace first occurrence of the digits with padded digits
                s_str = re.sub(r'\d+', digits, s_str, count=1)
                return s_str
            else:
                return s_str
        return s_str.zfill(11)

    df[expected_col] = df[expected_col].apply(pad_station)

    # 4) Save corrected CSV (overwrite)
    try:
        df.to_csv(out_csv, index=False, encoding="utf-8")
    except Exception as e:
        print(f"ERROR writing '{out_csv}': {e}", file=sys.stderr)
        raise

    # 5) Print filenames as required
    print(py_filename)
    print(out_csv)

if __name__ == "__main__":
    main()