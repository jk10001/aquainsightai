# filename: WWTP_Counts_by_Country_and_Level.py
import pandas as pd

# Load data
file_name = "HydroWASTE_v10_-_UTF8.csv"
df = pd.read_csv(file_name)

# Filter to rows with non-missing COUNTRY and LEVEL
df_filtered = df[df['COUNTRY'].notna() & df['LEVEL'].notna()]

# Define known levels for consistency and ordering
level_order = ['Primary', 'Secondary', 'Advanced']

# Create pivot table counting number of WWTPs by COUNTRY and LEVEL
pivot = pd.pivot_table(
    df_filtered,
    index='COUNTRY',
    columns='LEVEL',
    values='WASTE_ID',
    aggfunc='count',
    fill_value=0
)

# Ensure all expected LEVEL columns are present even if missing in data
for lvl in level_order:
    if lvl not in pivot.columns:
        pivot[lvl] = 0

# Reorder columns according to level_order
pivot = pivot[level_order]

# Add Total column
pivot['Total'] = pivot.sum(axis=1)

# Sort by Total descending
pivot_sorted = pivot.sort_values(by='Total', ascending=False)

# Save to CSV
out_filename = "WWTP_Counts_by_Country_and_Level.csv"
pivot_sorted.to_csv(out_filename)

# Display top 50 countries nicely formatted
print(pivot_sorted.head(50).to_string())

# Print output filename for record
print(out_filename)