# filename: api_calibration_spearman.py
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.io as pio
from scipy.stats import spearmanr

file_name = "MWC_raw_water_reservoirs_2015_2019.xlsx"

# Load relevant sheets
xls = pd.ExcelFile(file_name)
df_rainfall = pd.read_excel(xls, sheet_name="MajorCatchmentRainfal")
df_inflow = pd.read_excel(xls, sheet_name="MajorStorageStreamflow")

# Convert recorddate to datetime if not already
df_rainfall['recorddate'] = pd.to_datetime(df_rainfall['recorddate'])
df_inflow['recorddate'] = pd.to_datetime(df_inflow['recorddate'])

# Rename inflow column for convenience
df_inflow = df_inflow.rename(columns={"Streamflow into reservoir (ML/d)":"inflow_ML_d"})

# Define reservoirs with both rainfall and inflow data
rainfall_dams = set(df_rainfall['dam'].unique())
inflow_dams = set(df_inflow['dam'].unique())
common_dams = sorted(list(rainfall_dams.intersection(inflow_dams)))

# Define analysis full period
start_date = pd.Timestamp("2015-01-01")
end_date = pd.Timestamp("2019-12-31")

# Filter full period data and subset dams
df_rainfall = df_rainfall[(df_rainfall['recorddate'] >= start_date) & (df_rainfall['recorddate'] <= end_date) & (df_rainfall['dam'].isin(common_dams))]
df_inflow = df_inflow[(df_inflow['recorddate'] >= start_date) & (df_inflow['recorddate'] <= end_date) & (df_inflow['dam'].isin(common_dams))]

# Pivot rainfall and inflow to wide format to align date/dam
pivot_rainfall = df_rainfall.pivot(index='recorddate', columns='dam', values='rainfall_mm')
pivot_inflow = df_inflow.pivot(index='recorddate', columns='dam', values='inflow_ML_d')

# Fill missing rainfall within date range with 0 for each dam
pivot_rainfall = pivot_rainfall.reindex(pd.date_range(start_date, end_date, freq='D'))
pivot_rainfall = pivot_rainfall.fillna(0)

results = []
k_values = np.round(np.arange(0.85, 0.981, 0.01), 3)  # k from 0.85 to 0.98 inclusive, step 0.01

for dam in common_dams:
    rain_series = pivot_rainfall[dam]
    inflow_series = pivot_inflow[dam]
    
    # Align inflow series index to rainfall index, drop missing inflow after reindexing
    # Reindex inflow to full date range as well (missing means no inflow data)
    inflow_series = inflow_series.reindex(rain_series.index)
    
    # Initialize API array
    api_series = pd.Series(index=rain_series.index, dtype=float)
    
    for k in k_values:
        # Compute API recursively: API_t = P_t + k * API_{t-1}
        api_vals = np.zeros(len(rain_series))
        # Initialize first API value at day 0 as rainfall day 0
        if len(rain_series) == 0:
            results.append({'dam': dam, 'k': k, 'spearman_rho': np.nan, 'n_pairs': 0})
            continue
        api_vals[0] = rain_series.iat[0]
        for t in range(1, len(rain_series)):
            api_vals[t] = rain_series.iat[t] + k * api_vals[t-1]
        
        api_series.loc[rain_series.index] = api_vals
        
        # Create DataFrame and drop missing inflow or API if any (API has no NaNs here)
        combined = pd.DataFrame({'API': api_series, 'inflow': inflow_series}).dropna()
        n_pairs = combined.shape[0]
        if n_pairs == 0:
            rho = np.nan
        else:
            rho, _ = spearmanr(combined['API'], combined['inflow'])
        results.append({'dam': dam, 'k': k, 'spearman_rho': rho, 'n_pairs': n_pairs})

# Assemble results DataFrame
df_results = pd.DataFrame(results)

# Plot
fig = go.Figure()

colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728']  # Distinct default plotly colors, 4 reservoirs
color_map = dict(zip(common_dams, colors))

for dam in common_dams:
    df_dam = df_results[df_results['dam'] == dam]
    fig.add_trace(go.Scatter(
        x=df_dam['k'],
        y=df_dam['spearman_rho'],
        mode='lines+markers',
        name=dam,
        line=dict(color=color_map[dam]),
        marker=dict(size=8)
    ))

# Add horizontal zero line
fig.add_hline(y=0, line=dict(color='black', width=1, dash='dash'))

# Update layout
fig.update_layout(
    xaxis=dict(
        title='API decay factor k',
        dtick=0.01,
        range=[0.84,0.99],
        showgrid=True,
        zeroline=False,
        linecolor='black',
        ticks='outside',
        mirror=True,
    ),
    yaxis=dict(
        title='Spearman ρ (API vs inflow, lag=0)',
        range=[0,1],  # Modified to 0 to 1 as requested
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

# Save outputs
base_name = "api_calibration_spearman"
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