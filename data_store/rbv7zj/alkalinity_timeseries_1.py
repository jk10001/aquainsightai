# filename: alkalinity_timeseries_1.py
import pandas as pd
import plotly.graph_objects as go
import plotly.io as pio
import os
from datetime import datetime

# Input filename
file_name = "digester_data_3.csv"

# Output base name
base = "alkalinity_timeseries_1"
py_fname = os.path.basename(__file__) if '__file__' in globals() else "alkalinity_timeseries_1.py"
csv_out = f"{base}.csv"
html_out = f"{base}.html"
png_out = f"{base}.png"

# Read CSV, parse Date as day-first
df = pd.read_csv(file_name, encoding="utf-8", dayfirst=True)

# Parse Date column to datetime (day-first), coerce errors
df['Date'] = pd.to_datetime(df['Date'], dayfirst=True, errors='coerce')

# Select required columns and rename to requested headings
plot_df = df[['Date',
              'Digester 1 Alkalinity (mg/L)',
              'Digester 2 Alkalinity (mg/L)',
              'Digester 3 Alkalinity (mg/L)']].copy()

plot_df.rename(columns={
    'Digester 1 Alkalinity (mg/L)': 'Digester 1 Alkalinity [mg/L]',
    'Digester 2 Alkalinity (mg/L)': 'Digester 2 Alkalinity [mg/L]',
    'Digester 3 Alkalinity (mg/L)': 'Digester 3 Alkalinity [mg/L]'
}, inplace=True)

# Sort by Date to ensure time series order
plot_df.sort_values('Date', inplace=True)

# Save the CSV containing exactly the data plotted (Date as ISO YYYY-MM-DD)
out_csv_df = plot_df.copy()
out_csv_df['Date'] = out_csv_df['Date'].dt.strftime('%Y-%m-%d')
out_csv_df.to_csv(csv_out, index=False, encoding='utf-8')

# Build Plotly figure
pio.templates.default = "plotly_white"
fig = go.Figure()

# Line settings
line_width = 2
marker_settings = dict(size=6)

# Add traces for each digester
fig.add_trace(go.Scatter(
    x=plot_df['Date'],
    y=plot_df['Digester 1 Alkalinity [mg/L]'],
    mode='lines+markers',
    name='Digester 1',
    line=dict(width=line_width, color='#1f77b4'),
    marker=marker_settings,
    connectgaps=False
))
fig.add_trace(go.Scatter(
    x=plot_df['Date'],
    y=plot_df['Digester 2 Alkalinity [mg/L]'],
    mode='lines+markers',
    name='Digester 2',
    line=dict(width=line_width, color='#ff7f0e'),
    marker=marker_settings,
    connectgaps=False
))
fig.add_trace(go.Scatter(
    x=plot_df['Date'],
    y=plot_df['Digester 3 Alkalinity [mg/L]'],
    mode='lines+markers',
    name='Digester 3',
    line=dict(width=line_width, color='#2ca02c'),
    marker=marker_settings,
    connectgaps=False
))

# Add horizontal reference lines as separate traces so they appear in the legend
ref_lines = [
    (2500, '2500 mg/L (lower buffering target)', 'black', 'dash'),
    (4000, '4000 mg/L (upper buffering target)', 'red', 'dashdot')
]
# Determine x-range for lines; handle NaT
xmin = plot_df['Date'].min()
xmax = plot_df['Date'].max()
if pd.isna(xmin) or pd.isna(xmax):
    xmin = datetime.now()
    xmax = datetime.now()

for y, label, color, dash in ref_lines:
    fig.add_trace(go.Scatter(
        x=[xmin, xmax],
        y=[y, y],
        mode='lines',
        name=label,
        line=dict(color=color, width=2, dash=dash),
        hoverinfo='skip',
        showlegend=True
    ))

# Update axes: Y-axis starts at 0
fig.update_yaxes(title_text='Alkalinity [mg/L as CaCO3]', rangemode='tozero', zeroline=True)
fig.update_xaxes(title_text='Date', showgrid=True)

# Update layout to match style requirements: plotly_white, black axis lines and ticks, global font size 16, grid visible
fig.update_layout(
    template='plotly_white',
    font=dict(size=16),
    xaxis=dict(showline=True, linewidth=1, linecolor='black', mirror=True),
    yaxis=dict(showline=True, linewidth=1, linecolor='black', mirror=True),
    legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1),
    margin=dict(l=60, r=20, t=20, b=60)
)

# Ensure grid is visible
fig.update_xaxes(showgrid=True, gridcolor='#e6e6e6')
fig.update_yaxes(showgrid=True, gridcolor='#e6e6e6')

# Save HTML (interactive). Requirements: responsive, no rangeslider, no title, plotly_white style, black axis lines and tick marks,
# include_plotlyjs='cdn'
fig.write_html(html_out, include_plotlyjs='cdn', full_html=True, auto_open=False)

# Save PNG using kaleido at 1000x600
fig.write_image(png_out, format='png', width=1000, height=600, scale=1)

# Print created filenames and the python file name to the terminal
print(py_fname)
print(csv_out)
print(html_out)
print(png_out)