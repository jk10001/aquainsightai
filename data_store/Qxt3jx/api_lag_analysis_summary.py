# filename: api_lag_analysis_summary.py
import pandas as pd
import numpy as np

# Load api lag analysis data
df = pd.read_csv("api_lag_analysis.csv")

def summary_for_dam(df_dam):
    # Determine lag at max absolute rho (tie -> smallest lag)
    df_dam['abs_rho'] = df_dam['spearman_rho'].abs()
    max_abs_rho = df_dam['abs_rho'].max()
    candidates = df_dam[df_dam['abs_rho'] == max_abs_rho]
    lag_at_max = candidates['lag_days'].min()
    row_max = candidates[candidates['lag_days'] == lag_at_max].iloc[0]

    # Extract rhos at specified lags
    rho_lag0 = df_dam.loc[df_dam['lag_days']==0, 'spearman_rho'].values[0]
    rho_lag1 = df_dam.loc[df_dam['lag_days']==1, 'spearman_rho'].values[0]
    rho_lag3 = df_dam.loc[df_dam['lag_days']==3, 'spearman_rho'].values[0]
    rho_lag7 = df_dam.loc[df_dam['lag_days']==7, 'spearman_rho'].values[0]
    rho_lag14 = df_dam.loc[df_dam['lag_days']==14, 'spearman_rho'].values[0]

    # Compute largest lag with rho >= 0.75
    lags_ge_075 = df_dam[df_dam['spearman_rho'] >= 0.75]['lag_days']
    if len(lags_ge_075) == 0:
        forecast_horizon = np.nan
    else:
        forecast_horizon = lags_ge_075.max()

    summary = {
        'dam': df_dam['dam'].iloc[0],
        'k_opt': df_dam['k_opt'].iloc[0],
        'lag_at_max_abs_rho': lag_at_max,
        'rho_at_max_abs_rho': row_max['spearman_rho'],
        'abs_rho_at_max_abs_rho': row_max['abs_rho'],
        'n_pairs_at_max_abs_rho': row_max['n_pairs'],
        'rho_lag0': rho_lag0,
        'rho_lag1': rho_lag1,
        'rho_lag3': rho_lag3,
        'rho_lag7': rho_lag7,
        'rho_lag14': rho_lag14,
        'forecast_horizon_days_at_rho_ge_0_75': forecast_horizon
    }
    return summary

summary_rows = []
for dam, group in df.groupby('dam'):
    summary_rows.append(summary_for_dam(group))

df_summary = pd.DataFrame(summary_rows)

# Save summary CSV
output_csv = "api_lag_analysis_summary.csv"
df_summary.to_csv(output_csv, index=False)

print(output_csv)