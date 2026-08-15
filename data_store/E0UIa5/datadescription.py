# filename: datadescription.py
import pandas as pd
import io
import os

file_name = "muscatine_wrrf_digesters_2022.xlsx"
output_md = "datadescription.md"

ext = file_name.split(".")[-1].lower()

# Build a dict of {sheet_name: DataFrame}
if ext in {"xlsx", "xls"}:
    xls = pd.ExcelFile(file_name)
    data_frames = {sheet: pd.read_excel(xls, sheet_name=sheet) for sheet in xls.sheet_names}
elif ext == "csv":
    df_csv = pd.read_csv(file_name)
    sheet_name = "CSV Sheet"
    data_frames = {sheet_name: df_csv}
else:
    raise ValueError(f"Unsupported file extension '{ext}'.")

def is_date_like_series(series: pd.Series) -> bool:
    name = str(series.name).lower()
    if "date" in name:
        return True
    if pd.api.types.is_datetime64_any_dtype(series):
        return True
    if pd.api.types.is_object_dtype(series) or pd.api.types.is_string_dtype(series):
        non_null = series.dropna()
        if non_null.empty:
            return False
        parsed = pd.to_datetime(non_null, errors="coerce", infer_datetime_format=True)
        return parsed.notna().mean() >= 0.8
    return False

def is_numeric_like_value(x) -> bool:
    if pd.isna(x):
        return False
    if isinstance(x, (int, float)) and not isinstance(x, bool):
        return True
    s = str(x).strip()
    if s == "":
        return False
    return pd.to_numeric(s, errors="coerce") is not None and pd.notna(pd.to_numeric(s, errors="coerce"))

with open(output_md, "w", encoding="utf-8") as f:
    f.write(f"# Data description for {file_name}\n\n")

    if ext in {"xlsx", "xls"}:
        f.write(f"## Number of sheets: {len(data_frames)}\n\n")

    for sheet_name, df in data_frames.items():
        f.write(f"## Sheet name: {sheet_name}\n\n")

        # Column headings
        f.write("### Column Headings\n")
        f.write(", ".join([f"'{col}'" for col in df.columns]) + "\n\n")

        # First 10 rows
        f.write("### First 10 Rows\n")
        first_10 = df.head(10)
        f.write(first_10.to_markdown(index=True) + "\n\n")

        # Bottom 10 rows
        f.write("### Bottom 10 Rows\n")
        last_10 = df.tail(10)
        f.write(last_10.to_markdown(index=True) + "\n\n")

        # Describe include='all'
        f.write("### Describe (include='all')\n")
        try:
            desc = df.describe(include="all")
            f.write(desc.to_markdown() + "\n\n")
        except Exception as e:
            f.write(f"Describe failed: {e}\n\n")

        # Info verbose=True
        f.write("### Info (verbose=True)\n")
        buffer = io.StringIO()
        df.info(buf=buffer, verbose=True)
        info_output = buffer.getvalue()
        f.write("```\n" + info_output + "```\n\n")

        # Shape
        f.write("### Shape\n")
        f.write(f"{df.shape}\n\n")

        # Text columns unique values
        eligible_text_columns = []
        for col in df.columns:
            series = df[col]
            if "date" in str(col).lower():
                continue
            if is_date_like_series(series):
                continue
            if pd.api.types.is_numeric_dtype(series) or pd.api.types.is_datetime64_any_dtype(series):
                continue
            # object/string-like columns only
            if pd.api.types.is_object_dtype(series) or pd.api.types.is_string_dtype(series):
                eligible_text_columns.append(col)

        if eligible_text_columns:
            f.write("### Unique values in text columns (non-numeric only, up to 100 values)\n")
            for col in eligible_text_columns:
                series = df[col]
                non_null = series.dropna()

                unique_vals = []
                seen = set()

                for val in non_null.tolist():
                    sval = str(val).strip()
                    if sval == "":
                        token = "<empty string>"
                    elif is_numeric_like_value(sval):
                        continue
                    else:
                        token = sval

                    if token not in seen:
                        seen.add(token)
                        unique_vals.append(token)
                    if len(unique_vals) >= 100:
                        break

                if series.isna().any() and "<NaN>" not in seen and len(unique_vals) < 100:
                    unique_vals.append("<NaN>")

                unique_vals_quoted = [f"'{str(val)}'" for val in unique_vals[:100]]
                f.write(f"#### Column: {col}\n")
                if unique_vals_quoted:
                    f.write(", ".join(unique_vals_quoted) + "\n\n")
                else:
                    f.write("(No eligible unique text values found)\n\n")
        else:
            f.write("### Unique values in text columns (non-numeric only, up to 100 values)\n")
            f.write("(No eligible text columns found)\n\n")

print(os.path.basename(__file__))
print(output_md)