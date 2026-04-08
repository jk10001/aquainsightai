# filename: WWTP_Counts_by_Country_and_Treatment_Level.py
import pandas as pd

file_name = "HydroWASTE_v10_-_UTF8.csv"
df = pd.read_csv(file_name)

# Filter out rows with missing COUNTRY or LEVEL to avoid missing entries in the pivot
df_filtered = df[df['COUNTRY'].notna() & df['LEVEL'].notna()]

# Define treatment levels to maintain consistent order, only include those present in data
valid_levels = ['Primary', 'Secondary', 'Advanced']
df_filtered = df_filtered[df_filtered['LEVEL'].isin(valid_levels)]

# Create pivot table counting number of WWTPs by country and treatment level
pivot = pd.pivot_table(df_filtered, index='COUNTRY', columns='LEVEL', values='WASTE_ID',
                       aggfunc='count', fill_value=0)

# Ensure all levels columns exist, if missing add with zeros
for lvl in valid_levels:
    if lvl not in pivot.columns:
        pivot[lvl] = 0

# Add Total column
pivot['Total'] = pivot[valid_levels].sum(axis=1)

# Sort by Total descending
pivot_sorted = pivot.sort_values(by='Total', ascending=False)

# Reorder columns to have Primary, Secondary, Advanced, Total (alphabetical except Total last)
pivot_sorted = pivot_sorted[valid_levels + ['Total']]

# Save to CSV
output_csv = "WWTP_Counts_by_Country_and_Treatment_Level.csv"
pivot_sorted.to_csv(output_csv)

# Print table in a formatted way
print(pivot_sorted.head(50).to_markdown())

# Print filename for confirmation
print(output_csv)