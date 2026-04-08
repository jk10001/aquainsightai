# filename: api_calibration_summary.py
import pandas as pd
import numpy as np

# Load API calibration results from prior calculations
# File saved by the previous analysis step
csv_filename = "api_calibration_spearman.csv"
df_api = pd.read_csv(csv_filename)

# For each dam, find the k with max |Spearman rho|
def get_optimal_row(sub_df):
    # Calculate absolute rho
    sub_df['abs_rho'] = sub_df['spearman_rho'].abs()
    max_abs_rho = sub_df['abs_rho'].max()
    # Filter to rows with max absolute rho
    candidates = sub_df[sub_df['abs_rho'] == max_abs_rho]
    # If tie, take smallest k
    optimal = candidates.loc[candidates['k'].idxmin()]
    return optimal

optimal_rows = df_api.groupby('dam').apply(get_optimal_row).reset_index(drop=True)

# Compute memory_time_days = -1 / ln(k_opt)
optimal_rows['memory_time_days'] = -1 / np.log(optimal_rows['k'])

# Rename/organize columns as required
summary_df = optimal_rows.rename(columns={
    'k': 'k_opt',
    'spearman_rho': 'rho_at_k_opt',
    'abs_rho': 'abs_rho_at_k_opt',
    'n_pairs': 'n_pairs',
    'dam': 'dam'
})[['dam', 'k_opt', 'rho_at_k_opt', 'abs_rho_at_k_opt', 'n_pairs', 'memory_time_days']]

# Save CSV
output_csv = "api_calibration_summary.csv"
summary_df.to_csv(output_csv, index=False)

print(output_csv)