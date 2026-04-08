# filename: api_lag_analysis.py
import pandas as pd
import numpy as np
from scipy.stats import spearmanr
import plotly.graph_objects as go
import plotly.io as pio

file_name = "MWC_raw_water_reservoirs_2015_2019.xlsx"

# Load necessary datasets
xls = pd.ExcelFile(file_name)
df_rainfall = pd.read_excel(xls, sheet_name="MajorCatchmentRainfal")
df_inflow = pd.read_excel(xls, sheet_name="MajorStorageStreamflow")
df_summary = pd.read_csv("api_calibration_summary.csv")

# Convert dates
df_rainfall['recorddate'] = pd.to_datetime(df_rainfall['recorddate'])
df_inflow['recorddate'] = pd.to_datetime(df_inflow['recorddate'])

# Rename inflow column for convenience
df_inflow = df_inflow.rename(columns={"Streamflow into reservoir (ML/d)":"inflow_ML_d"})

# Reservoirs with api summary and both rainfall and inflow data
common_dams = sorted(df_summary['dam'].unique())

# Define analysis full period
start_date = pd.Timestamp("2015-01-01")
end_date = pd.Timestamp("2019-12-31")

# Filter data and subset dams
df_rainfall = df_rainfall[(df_rainfall['recorddate'] >= start_date) & (df_rainfall['recorddate'] <= end_date) & (df_rainfall['dam'].isin(common_dams))]
df_inflow = df_inflow[(df_inflow['recorddate'] >= start_date) & (df_inflow['recorddate'] <= end_date) & (df_inflow['dam'].isin(common_dams))]

# Pivot rainfall and inflow to wide format aligned by date and dam
pivot_rainfall = df_rainfall.pivot(index='recorddate', columns='dam', values='rainfall_mm')
pivot_inflow = df_inflow.pivot(index='recorddate', columns='dam', values='inflow_ML_d')

# Reindex rainfall to full date range and fill missing rainfall with 0
full_dates = pd.date_range(start_date, end_date, freq='D')
pivot_rainfall = pivot_rainfall.reindex(full_dates).fillna(0)
pivot_inflow = pivot_inflow.reindex(full_dates)  # Do not fill inflow

lags = range(0,15)  # 0 to 14 days lag

results = []

for dam in common_dams:
    k_opt = df_summary.loc[df_summary['dam']==dam, 'k_opt'].values[0]
    rain_series = pivot_rainfall[dam]
    inflow_series = pivot_inflow[dam]
    
    # Compute API for entire period using optimal k
    api_vals = np.zeros(len(rain_series))
    api_vals[0] = rain_series.iat[0]
    for t in range(1,len(rain_series)):
        api_vals[t] = rain_series.iat[t] + k_opt * api_vals[t-1]
    api_series = pd.Series(api_vals, index=rain_series.index)

    for lag in lags:
        # Correlate API[t] with inflow[t+lag]
        inflow_shifted = inflow_series.shift(-lag)  # inflow moves backward so inflow[t+lag] aligns to API[t]
        combined = pd.DataFrame({'API': api_series, 'inflow': inflow_shifted}).dropna()
        n_pairs = combined.shape[0]
        if n_pairs == 0:
            rho = np.nan
        else:
            rho, _ = spearmanr(combined['API'], combined['inflow'])
        results.append({'dam': dam, 'k_opt': k_opt, 'lag_days': lag, 'spearman_rho': rho, 'n_pairs': n_pairs})

df_results = pd.DataFrame(results)

# Plot
fig = go.Figure()

colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728']
color_map = dict(zip(common_dams, colors))

for dam in common_dams:
    df_dam = df_results[df_results['dam'] == dam]
    fig.add_trace(go.Scatter(
        x=df_dam['lag_days'],
        y=df_dam['spearman_rho'],
        mode='lines+markers',
        name=dam,
        line=dict(color=color_map[dam]),
        marker=dict(size=8)
    ))

# Horizontal zero line
fig.add_hline(y=0, line=dict(color='black', width=1, dash='dash'))

# Layout
fig.update_layout(
    xaxis=dict(
        title='API lead (days)',
        dtick=1,
        range=[-0.1,14.1],
        showgrid=True,
        zeroline=False,
        linecolor='black',
        ticks='outside',
        mirror=True,
    ),
    yaxis=dict(
        title='Spearman ρ (API vs inflow)',
        range=[0,1],  # Changed from [-1, 1] to [0,1] as requested
        showgrid=True,
        zeroline=False,
        linecolor='black',
        ticks='outside',
        mirror=True,
    ),
    legend=dict(
        bordercolor='black',
        borderwidth=1,
    ),
    template='plotly_white',
    font=dict(size=16),
    margin=dict(t=40,b=40,l=60,r=40),
    hovermode='x unified'
)

# Save files
base_name = "api_lag_analysis"
csv_filename = f"{base_name}.csv"
html_filename = f"{base_name}.html"
png_filename = f"{base_name}.png"
py_filename = f"{base_name}.py"

df_results.to_csv(csv_filename, index=False)

pio.write_html(fig, file=html_filename, include_plotlyjs='cdn', full_html=True, config={'displayModeBar': True})
fig.write_image(png_filename, width=1000, height=600, scale=1, engine='kaleido')

print(py_filename)
print(csv_filename)
print(html_filename)
print(png_filename)