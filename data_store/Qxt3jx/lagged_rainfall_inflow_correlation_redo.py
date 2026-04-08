# filename: lagged_rainfall_inflow_correlation_redo.py
import pandas as pd
import numpy as np
from scipy.stats import spearmanr
import plotly.graph_objects as go
import plotly.io as pio

file_name = "MWC_raw_water_reservoirs_2015_2019.xlsx"

# Load datasets
xls = pd.ExcelFile(file_name)
df_rainfall = pd.read_excel(xls, sheet_name="MajorCatchmentRainfal")
df_inflow = pd.read_excel(xls, sheet_name="MajorStorageStreamflow")

# Convert dates to datetime
df_rainfall['recorddate'] = pd.to_datetime(df_rainfall['recorddate'])
df_inflow['recorddate'] = pd.to_datetime(df_inflow['recorddate'])

# Rename inflow column for convenience
df_inflow.rename(columns={"Streamflow into reservoir (ML/d)": "inflow_ML_d"}, inplace=True)

# Find reservoirs common to both datasets
rainfall_dams = set(df_rainfall['dam'].unique())
inflow_dams = set(df_inflow['dam'].unique())
common_dams = sorted(list(rainfall_dams.intersection(inflow_dams)))

# Define full analysis period from data range
start_date = pd.Timestamp("2015-01-01")
end_date = pd.Timestamp("2019-12-31")

# Filter data to full period and only common dams
df_rainfall = df_rainfall[(df_rainfall['recorddate'] >= start_date) & (df_rainfall['recorddate'] <= end_date) & (df_rainfall['dam'].isin(common_dams))]
df_inflow = df_inflow[(df_inflow['recorddate'] >= start_date) & (df_inflow['recorddate'] <= end_date) & (df_inflow['dam'].isin(common_dams))]

# Pivot rainfall and inflow to wide format (dates x dams)
pivot_rainfall = df_rainfall.pivot(index='recorddate', columns='dam', values='rainfall_mm')
pivot_inflow = df_inflow.pivot(index='recorddate', columns='dam', values='inflow_ML_d')

# Reindex rainfall to full date range, fill missing with 0 (no rainfall)
full_dates = pd.date_range(start_date, end_date, freq='D')
pivot_rainfall = pivot_rainfall.reindex(full_dates).fillna(0)
pivot_inflow = pivot_inflow.reindex(full_dates)  # inflow not filled to keep NaNs

lags = range(0, 15)  # 0 to 14 days lag (rainfall leading inflow)

results = []

for dam in common_dams:
    rain_series = pivot_rainfall[dam]
    inflow_series = pivot_inflow[dam]
    for lag in lags:
        # inflow shifted backward so inflow[t+lag] aligns with rainfall[t]
        inflow_shifted = inflow_series.shift(-lag)
        combined_df = pd.DataFrame({'rainfall': rain_series, 'inflow': inflow_shifted}).dropna()
        if combined_df.empty:
            rho = np.nan
        else:
            rho, _ = spearmanr(combined_df['rainfall'], combined_df['inflow'])
        results.append({'dam': dam, 'lag_days': lag, 'spearman_rho': rho})

df_results = pd.DataFrame(results)

# Plot lagged rainfall-inflow Spearman correlation for all dams on same chart
fig = go.Figure()

colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728',
          '#9467bd', '#8c564b', '#e377c2', '#7f7f7f',
          '#bcbd22', '#17becf']  # sufficient colors for up to 10 dams
color_map = dict(zip(common_dams, colors))

for dam in common_dams:
    df_dam = df_results[df_results['dam'] == dam]
    # Clip correlations below 0 to 0 because y-axis is 0 to 1 as requested
    # But we keep original values to plot properly with lines and markers
    y_vals = df_dam['spearman_rho'].clip(lower=0)
    fig.add_trace(go.Scatter(
        x=df_dam['lag_days'],
        y=y_vals,
        mode='lines+markers',
        name=dam,
        line=dict(color=color_map[dam]),
        marker=dict(size=8)
    ))

# Add horizontal line at y=0
fig.add_hline(y=0, line=dict(color='black', width=1, dash='dash'))

# Update layout per instructions
fig.update_layout(
    xaxis=dict(
        title='Days lag (rainfall leading inflow)',
        dtick=1,
        range=[0, 14],
        showgrid=True,
        zeroline=False,
        linecolor='black',
        ticks='outside',
        mirror=True,
    ),
    yaxis=dict(
        title='Spearman ρ (Rainfall vs Inflow)',
        range=[0, 1],  # Set y-axis range to 0 to 1 as requested
        showgrid=True,
        zeroline=False,
        linecolor='black',
        ticks='outside',
        mirror=True,
    ),
    template='plotly_white',
    font=dict(size=16),
    legend=dict(
        bordercolor='black',
        borderwidth=1,
    ),
    margin=dict(t=40, b=40, l=60, r=40),
    hovermode='x unified'
)

# Save output files using consistent basename
base_name = "lagged_rainfall_inflow_correlation_redo"
csv_filename = f"{base_name}.csv"
html_filename = f"{base_name}.html"
png_filename = f"{base_name}.png"
py_filename = f"{base_name}.py"

# Save CSV with data used for plotting (exact filtered/calculated data)
df_results.to_csv(csv_filename, index=False)

# Save HTML chart (responsive, plotly_white, font size 16, black axis lines and ticks, include_plotlyjs='cdn')
pio.write_html(fig, file=html_filename, include_plotlyjs='cdn', full_html=True,
               config={'displayModeBar': True})

# Save PNG chart 1000x600 px with kaleido
fig.write_image(png_filename, width=1000, height=600, scale=1, engine='kaleido')

print(py_filename)
print(csv_filename)
print(html_filename)
print(png_filename)