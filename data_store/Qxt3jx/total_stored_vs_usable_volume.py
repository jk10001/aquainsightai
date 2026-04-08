# filename: total_stored_vs_usable_volume.py
import pandas as pd
import plotly.graph_objects as go
import plotly.io as pio

file_name = "MWC_raw_water_reservoirs_2015_2019.xlsx"

# Load relevant sheets
xls = pd.ExcelFile(file_name)
df_storage = pd.read_excel(xls, sheet_name="StorageVolume")
df_assets = pd.read_excel(xls, sheet_name="WS_Storage_Dam_OpenData")

# Convert recorddate to datetime if not already
df_storage['recorddate'] = pd.to_datetime(df_storage['recorddate'])

# Filter to full period explicitly just to ensure
start_date = pd.Timestamp("2015-01-01")
end_date = pd.Timestamp("2019-12-31")
mask = (df_storage['recorddate'] >= start_date) & (df_storage['recorddate'] <= end_date)
df_storage = df_storage.loc[mask]

# Pivot storage volumes to wide format: index date, columns dams, values volume (ML)
pivot_storage = df_storage.pivot(index='recorddate', columns='dam', values='volume (ML)')

# Sort columns alphabetically (to ensure consistent stacking order)
pivot_storage = pivot_storage.sort_index(axis=1)

# Calculate total stored volume daily (sum across dams axis=1)
pivot_storage['total_stored_ML'] = pivot_storage.sum(axis=1)

# Calculate total usable volume (static sum of USABLE_VOLUME in assets sheet)
# Sum all reservoirs in WS_Storage_Dam_OpenData regardless of whether they appear in storage volume timeseries
total_usable_volume_ML = df_assets['USABLE_VOLUME'].sum()

# Create dataframe for plotting and output csv with:
# daily volume per dam, total_stored_ML, total_usable_volume_ML (repeated for each day)
pivot_storage['total_usable_volume_ML'] = total_usable_volume_ML

# Prepare DataFrame for CSV output
output_df = pivot_storage.reset_index()

# Plot stacked area chart with line overlay
fig = go.Figure()

# Add stacked area traces by dam (exclude the total columns)
for dam in pivot_storage.columns.drop(['total_stored_ML', 'total_usable_volume_ML']):
    fig.add_trace(go.Scatter(
        x=output_df['recorddate'],
        y=output_df[dam],
        mode='lines',
        stackgroup='one',
        name=dam,
        line=dict(width=0.5),
        hoverinfo='x+y+name'
    ))

# Add the horizontal usable volume line
fig.add_trace(go.Scatter(
    x=output_df['recorddate'],
    y=output_df['total_usable_volume_ML'],
    mode='lines',
    name='Total usable volume (ML)',
    line=dict(color='black', width=3, dash='dash'),
    hoverinfo='x+y+name'
))

# Update layout per requirements: no title, axis labels, legend
fig.update_layout(
    xaxis_title="Date",
    yaxis_title="Stored volume (ML)",
    template='plotly_white',
    font=dict(size=16),
    xaxis=dict(showgrid=True, linecolor='black', mirror=True, ticks='outside', showline=True),
    yaxis=dict(showgrid=True, linecolor='black', mirror=True, ticks='outside', showline=True),
    legend=dict(
        bordercolor='black',
        borderwidth=1,
    )
)

# Save files: use base name
base_name = "total_stored_vs_usable_volume"
csv_filename = f"{base_name}.csv"
html_filename = f"{base_name}.html"
png_filename = f"{base_name}.png"
py_filename = f"{base_name}.py"

# Save CSV with the plotted data
output_df.to_csv(csv_filename, index=False)

# Save HTML file (responsive, no rangeslider, plotly_white style, font 16, black axis lines)
pio.write_html(
    fig, file=html_filename,
    include_plotlyjs='cdn',
    full_html=True,
    config={'displayModeBar': True},
)

# Save PNG at 1000x600px with kaleido
fig.write_image(png_filename, width=1000, height=600, scale=1, engine='kaleido')

print(py_filename)
print(csv_filename)
print(html_filename)
print(png_filename)