# filename: datadescription.py
import pandas as pd
import io
import sys
import traceback

file_name = "digester_data_3.csv"
output_md = "datadescription.md"

def safe_to_markdown(df):
    try:
        return df.to_markdown()
    except Exception:
        # Fallback to CSV representation if tabulate not available or error occurs
        return df.to_csv(index=True)

def quote_vals(vals):
    return ", ".join([f"'{str(v)}'" for v in vals])

def first_n_unique_non_numeric(series, n=100):
    # Work on non-null values
    non_null = series.dropna().astype(str)
    # Identify empty strings (after stripping)
    stripped = non_null.str.strip()
    has_empty = (stripped == "").any()
    non_empty = non_null[stripped != ""]
    # Remove numeric-like values
    as_str = non_empty.str.strip()
    numeric_mask = pd.to_numeric(as_str, errors="coerce").notna()
    non_numeric = non_empty[~numeric_mask]
    uniques = list(dict.fromkeys(non_numeric.tolist()))  # preserve order, unique
    extras = []
    if has_empty:
        extras.append("<empty string>")
    if series.isna().any():
        extras.append("<NaN>")
    combined = uniques[:n]
    combined.extend(extras)
    return combined

def main():
    try:
        ext = file_name.split('.')[-1].lower()

        # Build dict of {sheet_name: DataFrame}
        if ext in {'xlsx', 'xls'}:
            xls = pd.ExcelFile(file_name)
            data_frames = {sheet: pd.read_excel(xls, sheet_name=sheet) for sheet in xls.sheet_names}
        elif ext == 'csv':
            df_csv = pd.read_csv(file_name, encoding="utf-8")
            sheet_name = 'CSV Sheet'
            data_frames = {sheet_name: df_csv}
        else:
            raise ValueError(f"Unsupported file extension '{ext}'.")

        with open(output_md, "w", encoding="utf-8") as f:
            f.write(f"# Data description for {file_name}\n\n")

            if ext == "csv":
                f.write(f"## CSV Encoding: The expected CSV encoding to use with this CSV is encoding=\"utf-8\"\n\n")

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
                    f.write(safe_to_markdown(df.head(10)) + "\n\n")
                except Exception as e:
                    f.write("Could not render first 10 rows: " + str(e) + "\n\n")

                # Bottom 10 rows
                f.write("### Bottom 10 Rows\n")
                try:
                    f.write(safe_to_markdown(df.tail(10)) + "\n\n")
                except Exception as e:
                    f.write("Could not render bottom 10 rows: " + str(e) + "\n\n")

                # Describe function include='all'
                f.write("### Describe (include='all')\n")
                try:
                    desc = df.describe(include="all")
                    f.write(safe_to_markdown(desc) + "\n\n")
                except Exception as e:
                    f.write("Describe failed: " + str(e) + "\n\n")

                # Info function verbose=True
                f.write("### Info (verbose=True)\n")
                try:
                    buffer = io.StringIO()
                    df.info(buf=buffer, verbose=True)
                    info_output = buffer.getvalue()
                    f.write("```\n" + info_output + "\n```\n\n")
                except Exception as e:
                    f.write("Info failed: " + str(e) + "\n\n")

                # Shape function
                f.write("### Shape\n")
                try:
                    f.write(str(df.shape) + "\n\n")
                except Exception as e:
                    f.write("Shape retrieval failed: " + str(e) + "\n\n")

                # Text columns unique values (excluding numeric values, dates, and columns with 'date' in name)
                try:
                    # Consider object dtype or string dtype as text columns; exclude any column with 'date' in name (case-insensitive)
                    text_columns = [col for col in df.columns if ("date" not in col.lower()) and (df[col].dtype == object or pd.api.types.is_string_dtype(df[col]))]
                    if text_columns:
                        f.write("### Unique values in text columns (non-numeric only, up to 100 values)\n")
                        for col in text_columns:
                            uniques = first_n_unique_non_numeric(df[col], n=100)
                            quoted = quote_vals(uniques)
                            f.write(f"#### Column: {col}\n")
                            if quoted.strip() == "":
                                f.write("(no non-numeric values found)\n\n")
                            else:
                                f.write(quoted + "\n\n")
                    else:
                        f.write("### Unique values in text columns\nNo text columns found (excluding columns with 'date' in name).\n\n")
                except Exception as e:
                    f.write("Failed to compute unique text values: " + str(e) + "\n\n")

        # Print created filenames to terminal as required
        print(output_md)
        print("datadescription.py")

    except Exception as e:
        # Write traceback to md file as well for debugging
        try:
            with open(output_md, "a", encoding="utf-8") as f:
                f.write("\n\n## Error during processing\n")
                f.write(traceback.format_exc())
        except Exception:
            pass
        print("An error occurred during execution:", file=sys.stderr)
        traceback.print_exc()
        # Still print python filename as required
        print("datadescription.py")

if __name__ == "__main__":
    main()