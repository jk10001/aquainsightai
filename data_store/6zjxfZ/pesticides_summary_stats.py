# filename: pesticides_summary_stats.py
import pandas as pd
import numpy as np

file_name = "Anglian_Water_Domestic_Water_Quality.csv"

# Load data
df = pd.read_csv(file_name, encoding="utf-8")

# Filter pesticide determinands:
# Determinand starts with "Pesticides " OR equals "Pesticides (Total by Calculation)"
mask_pesticides = df['Determinand'].str.startswith("Pesticides ", na=False) | (df['Determinand'] == "Pesticides (Total by Calculation)")
df_pesticides = df.loc[mask_pesticides].copy()

# Treat Operator:
# Operator == '<' means non-detect; exclude from median and max calculations
df_pesticides['Is_Detect'] = ~df_pesticides['Operator'].eq('<')

# Calculate counts
agg = df_pesticides.groupby(['Determinand', 'DWI_Code', 'Units']).agg(
    n_results=('Result', 'size'),
    n_detects=('Is_Detect', 'sum'),
    max_detect=('Result', lambda x: x[df_pesticides.loc[x.index, 'Is_Detect']].max() if x[df_pesticides.loc[x.index, 'Is_Detect']].size > 0 else np.nan),
    median_detect=('Result', lambda x: x[df_pesticides.loc[x.index, 'Is_Detect']].median() if x[df_pesticides.loc[x.index, 'Is_Detect']].size > 0 else np.nan)
).reset_index()

# Function to normalize units for comparison to PCV in ug/L
def convert_to_ugL(row):
    units = str(row['Units']).lower()
    val_max = row['max_detect']
    val_median = row['median_detect']
    if pd.isnull(val_max) or pd.isnull(val_median):
        return pd.Series([np.nan, np.nan])
    if any(u in units for u in ['µg/l', 'μg/l', 'ug/l', 'î¼g/l', 'ï¿¼g/l', 'î»g/l', 'µg/l']):
        # Already micrograms per litre, no conversion
        return pd.Series([val_max, val_median])
    elif 'mg/l' in units:
        # Convert mg/L to ug/L
        return pd.Series([val_max * 1000, val_median * 1000])
    else:
        # Unknown units conversion, use original values (warning: may be invalid)
        return pd.Series([val_max, val_median])

# Apply conversion
agg[['max_detect_ugL', 'median_detect_ugL']] = agg.apply(convert_to_ugL, axis=1)

# Add PCV columns according to rules
def assign_pcvs(determin):
    pcv_individual = np.nan
    pcv_total = np.nan
    if determin == "Pesticides Aldrin":
        pcv_individual = 0.03
    elif determin == "Pesticides Dieldrin (Total)":
        pcv_individual = 0.03
    elif determin == "Pesticides Heptachlor (Total)":
        pcv_individual = 0.03
    elif determin == "Pesticides Heptachlor Epoxide - Total (Trans, CIS) (Heptachlor Epoxide)":
        pcv_individual = 0.03
    elif determin == "Pesticides (Total by Calculation)":
        pcv_total = 0.5
    else:
        # For all other pesticides except total pesticides (and not one of the above)
        if determin.startswith("Pesticides "):
            pcv_individual = 0.1
    return pd.Series([pcv_individual, pcv_total])

agg[['PCV_ugL_individual', 'PCV_ugL_total']] = agg['Determinand'].apply(assign_pcvs)

# Save to CSV
output_file = "pesticides_summary_stats.csv"
agg.to_csv(output_file, index=False)

print(f"pesticides_summary_stats.csv")
print(f"pesticides_summary_stats.py")