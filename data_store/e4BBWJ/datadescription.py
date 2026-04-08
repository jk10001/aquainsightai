# filename: datadescription.py
import pandas as pd
import io
import os
import sys

file_name = "Ontario_PWQMN_2024.xlsx"
output_md = "datadescription.md"

def main():
    if not os.path.exists(file_name):
        print(f"Error: Source file not found: {file_name}", file=sys.stderr)
        sys.exit(1)

    ext = file_name.split('.')[-1].lower()

    # Build a dict of {sheet_name: DataFrame}
    if ext in {'xlsx', 'xls'}:
        xls = pd.ExcelFile(file_name)
        try:
            data_frames = {sheet: pd.read_excel(xls, sheet_name=sheet) for sheet in xls.sheet_names}
        finally:
            try:
                xls.close()
            except Exception:
                pass
    elif ext == 'csv':
        df_csv = pd.read_csv(file_name)
        sheet_name = 'CSV Sheet'
        data_frames = {sheet_name: df_csv}
    else:
        raise ValueError(f"Unsupported file extension '{ext}'.")

    with open(output_md, "w", encoding="utf-8") as f:
        f.write(f"# Data description for {file_name}\n\n")

        if ext == "csv":
            f.write(f"## CSV Encoding: The expected CSV encoding to use with this CSV is \n\n")

        if ext in {"xlsx", "xls"}:
            sheet_count = len(data_frames)
            f.write(f"## Number of sheets: \n{sheet_count}\n\n")

        for sheet_name, df in data_frames.items():
            f.write(f"## Sheet name: {sheet_name}\n\n")

            # Column headings
            f.write("### Column Headings\n")
            f.write(", ".join([f"'{col}'" for col in df.columns]) + "\n\n")

            # First 10 rows
            f.write("### First 10 Rows\n")
            try:
                f.write(df.head(10).to_markdown() + "\n\n")
            except Exception:
                # Fallback if to_markdown not available
                f.write(df.head(10).to_csv(index=False) + "\n\n")

            # Bottom 10 rows
            f.write("### Bottom 10 Rows\n")
            try:
                f.write(df.tail(10).to_markdown() + "\n\n")
            except Exception:
                f.write(df.tail(10).to_csv(index=False) + "\n\n")

            # Describe function include='all'
            f.write("### Describe (include='all')\n")
            try:
                desc = df.describe(include="all")
                try:
                    f.write(desc.to_markdown() + "\n\n")
                except Exception:
                    f.write(desc.to_csv() + "\n\n")
            except Exception as e:
                f.write(f"Error running describe(include='all'): {e}\n\n")

            # Info function verbose=True
            f.write("### Info (verbose=True)\n")
            buffer = io.StringIO()
            try:
                df.info(buf=buffer, verbose=True)
            except TypeError:
                # Older pandas may not accept verbose param
                df.info(buf=buffer)
            info_output = buffer.getvalue()
            f.write("```\n" + info_output + "\n```\n\n")

            # Shape function
            f.write("### Shape\n")
            f.write(str(df.shape) + "\n\n")

            # Text columns unique values (excluding numeric values, dates, and columns with 'date' in name)
            # Identify object dtype and also string dtype columns
            candidate_text_cols = []
            for col in df.columns:
                if "date" in str(col).lower():
                    continue
                # Consider columns with object or string dtype as textual
                if pd.api.types.is_object_dtype(df[col].dtype) or pd.api.types.is_string_dtype(df[col].dtype):
                    candidate_text_cols.append(col)

            if candidate_text_cols:
                f.write("### Unique values in text columns (non-numeric only, up to 100 values)\n")
                for col in candidate_text_cols:
                    series = df[col]

                    # Work on non-null values; keep empty-string and NaN flags separately
                    non_null = series.dropna()

                    # Identify empty strings (after stripping)
                    # Convert to string for stripping safely
                    stripped = non_null.astype(str).str.strip()
                    has_empty = (stripped == "").any()

                    # Keep only non-empty strings
                    non_empty_mask = stripped != ""
                    non_empty = non_null[non_empty_mask]

                    # Remove numeric-like values: if entire value can be parsed as number, treat as numeric-like
                    as_str = non_empty.astype(str).str.strip()
                    numeric_mask = as_str.apply(lambda x: pd.to_numeric(x, errors="coerce")).notna()
                    non_numeric = non_empty[~numeric_mask]

                    # Collect unique values (preserve original representation where possible)
                    uniques = pd.Series(non_numeric.unique()).tolist()

                    if has_empty:
                        uniques.append("<empty string>")
                    if series.isna().any():
                        uniques.append("<NaN>")

                    # Limit to first 100 unique values
                    unique_vals_quoted = [f"'{str(val)}'" for val in uniques[:100]]

                    f.write(f"#### Column: {col}\n")
                    if unique_vals_quoted:
                        f.write(", ".join(unique_vals_quoted) + "\n\n")
                    else:
                        f.write("(no non-numeric textual unique values found)\n\n")

    # Print created file names to terminal as required
    # Print the python filename and any created files
    print(os.path.basename(__file__))
    print(output_md)

if __name__ == "__main__":
    main()