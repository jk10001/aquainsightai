# filename: datadescription.py
import pandas as pd
import io

file_name = "Anglian_Water_Domestic_Water_Quality.csv"

ext = file_name.split('.')[-1]

# Build a dict of {sheet_name: DataFrame}
if ext in {'xlsx', 'xls'}:
    xls = pd.ExcelFile(file_name)
    data_frames = {sheet: pd.read_excel(xls, sheet_name=sheet) for sheet in xls.sheet_names}
elif ext == 'csv':
    df_csv = pd.read_csv(file_name, encoding="utf-8")
    sheet_name = 'CSV Sheet'
    data_frames = {sheet_name: df_csv}
else:
    raise ValueError(f"Unsupported file extension '{ext}'.")

with open(f'datadescription.md', 'w', encoding='utf-8') as f:
    f.write(f"# Data description for {file_name}\n\n")

    if ext == 'csv':
        f.write(f"## CSV Encoding: The expected CSV encoding to use with this CSV is encoding=\"utf-8\"\n\n")

    if ext in {'xlsx', 'xls'}:
        sheet_count = len(data_frames)
        f.write(f"## Number of sheets: \n{sheet_count}\n\n")

    for sheet_name, df in data_frames.items():
        
        f.write(f"## Sheet name: {sheet_name}\n\n")
        
        # Column headings
        f.write("### Column Headings\n")
        f.write(", ".join([f"'{col}'" for col in df.columns]) + "\n\n")
        
        # First 10 rows
        f.write("### First 10 Rows\n")
        f.write(df.head(10).to_markdown() + "\n\n")
        
        # Bottom 10 rows
        f.write("### Bottom 10 Rows\n")
        f.write(df.tail(10).to_markdown() + "\n\n")
        
        # Describe function include='all'
        f.write("### Describe (include='all')\n")
        f.write(df.describe(include='all').to_markdown() + "\n\n")
        
        # Info function verbose=True
        f.write("### Info (verbose=True)\n")
        buffer = io.StringIO()
        df.info(buf=buffer, verbose=True)
        info_output = buffer.getvalue()
        f.write("```\n" + info_output + "\n```\n\n")
        
        # Shape function
        f.write("### Shape\n")
        f.write(str(df.shape) + "\n\n")
        
        # Text columns unique values (excluding numeric, dates, and columns with 'date' in name)
        text_columns = [col for col in df.columns if df[col].dtype == object and 'date' not in col.lower()]
        if text_columns:
            f.write("### Unique values in text columns (up to 100 values)\n")
            for col in text_columns:
                series = df[col]
                uniques = series.dropna()
                uniques = uniques[uniques != ''].unique().tolist()
                if (series == '').any():
                    uniques.append("<empty string>")
                if series.isna().any():
                    uniques.append("<NaN>")
                unique_vals_quoted = [f"'{str(val)}'" for val in uniques[:100]]
                f.write(f"#### Column: {col}\n")
                f.write(", ".join(unique_vals_quoted) + "\n\n")

print(f"datadescription.md")
print("datadescription.py")