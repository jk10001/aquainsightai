# filename: vs_destruction_1.py
import pandas as pd
import plotly.graph_objects as go
import plotly.io as pio
import os
from datetime import datetime

# Input filename
file_name = "digester_data_3.csv"

# Output base name
base = "vs_destruction_1"
py_fname = os.path.basename(__file__) if '__file__' in globals() else "vs_destruction_1.py"
csv_out = f"{base}.csv"
html_out = f"{base}.html"
png_out = f"{base}.png"
txt_out = f"{base}_reference.txt"

# Read CSV, parse Date as day-first
df = pd.read_csv(file_name, encoding="utf-8", dayfirst=True)

# Parse Date column to datetime (day-first), coerce errors
df['Date'] = pd.to_datetime(df['Date'], dayfirst=True, errors='coerce')

# Select required columns
cols = ['Date',
        'Digester 1 volatile solids destruction (%)',
        'Digester 2 volatile solids destruction (%)',
        'Digester 3 volatile solids destruction (%)']
plot_df = df[cols].copy()

# Sort by Date ascending
plot_df.sort_values('Date', inplace=True)

# Ensure numeric types (coerce if needed)
for c in cols[1:]:
    plot_df[c] = pd.to_numeric(plot_df[c], errors='coerce')

# Prepare CSV output: Date as ISO YYYY-MM-DD and clear headings with units
out_csv_df = plot_df.copy()
out_csv_df['Date'] = out_csv_df['Date'].dt.strftime('%Y-%m-%d')
out_csv_df.rename(columns={
    'Digester 1 volatile solids destruction (%)': 'Digester 1 VS destruction [%]',
    'Digester 2 volatile solids destruction (%)': 'Digester 2 VS destruction [%]',
    'Digester 3 volatile solids destruction (%)': 'Digester 3 VS destruction [%]'
}, inplace=True)

# Save the CSV containing exactly the data plotted
out_csv_df.to_csv(csv_out, index=False, encoding='utf-8')

# Build Plotly figure
pio.templates.default = "plotly_white"
fig = go.Figure()

line_width = 2
marker_settings = dict(size=6)

# Add traces for each digester (lines + markers), keep gaps for NaNs
fig.add_trace(go.Scatter(
    x=plot_df['Date'],
    y=plot_df['Digester 1 volatile solids destruction (%)'],
    mode='lines+markers',
    name='Digester 1',
    line=dict(width=line_width, color='#1f77b4'),
    marker=marker_settings,
    connectgaps=False
))
fig.add_trace(go.Scatter(
    x=plot_df['Date'],
    y=plot_df['Digester 2 volatile solids destruction (%)'],
    mode='lines+markers',
    name='Digester 2',
    line=dict(width=line_width, color='#ff7f0e'),
    marker=marker_settings,
    connectgaps=False
))
fig.add_trace(go.Scatter(
    x=plot_df['Date'],
    y=plot_df['Digester 3 volatile solids destruction (%)'],
    mode='lines+markers',
    name='Digester 3',
    line=dict(width=line_width, color='#2ca02c'),
    marker=marker_settings,
    connectgaps=False
))

# Add horizontal reference lines for 45% and 60%
xmin = plot_df['Date'].min()
xmax = plot_df['Date'].max()
if pd.isna(xmin) or pd.isna(xmax):
    xmin = datetime.now()
    xmax = datetime.now()

ref_lines = [
    (45, '45% (typical lower bound)', 'black', 'dash'),
    (60, '60% (typical upper bound)', 'red', 'dashdot')
]
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

# Update axes and layout to match requirements
fig.update_yaxes(title_text='Volatile solids destruction [%]', rangemode='tozero')
fig.update_xaxes(title_text='Date', showgrid=True)

fig.update_layout(
    template='plotly_white',
    font=dict(size=16),
    xaxis=dict(showline=True, linewidth=1, linecolor='black', mirror=True),
    yaxis=dict(showline=True, linewidth=1, linecolor='black', mirror=True),
    legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1),
    margin=dict(l=60, r=20, t=20, b=60)
)

fig.update_xaxes(showgrid=True, gridcolor='#e6e6e6')
fig.update_yaxes(showgrid=True, gridcolor='#e6e6e6')

# Save HTML (interactive). Requirements: responsive, no rangeslider, no title, plotly_white style, black axis lines and tick marks,
# include_plotlyjs='cdn'
fig.write_html(html_out, include_plotlyjs='cdn', full_html=True, auto_open=False)

# Save PNG using kaleido at 1000x600
fig.write_image(png_out, format='png', width=1000, height=600, scale=1)

# Save short TXT reference file with a credible source for typical VS destruction ranges
reference_text = (
    "Typical volatile solids (VS) destruction ranges for municipal anaerobic digesters:\n\n"
    "Many wastewater engineering references report typical VS destruction for anaerobic digestion in the\n"
    "range of ~45% to 60% for well-operated municipal full-scale digesters. Example source:\n\n"
    "Metcalf & Eddy (4th Ed.) / 'Wastewater Engineering: Treatment and Resource Recovery' — Section on\n"
    "Anaerobic Digestion and Sludge Stabilization. This textbook discusses expected VS destruction\n"
    "in the typical range of about 40–60% depending on digester type and operating conditions.\n\n"
    "Another practical guidance: U.S. EPA and technical manuals on biosolids and anaerobic digestion\n"
    "commonly reference similar ranges for mesophilic digesters under normal loading and retention times.\n\n"
    "Please consult the cited textbook or local design guidance for precise design/operational targets."
)
with open(txt_out, "w", encoding="utf-8") as f:
    f.write(reference_text)

# Print created filenames and the python file name to the terminal
print(py_fname)
print(csv_out)
print(html_out)
print(png_out)
print(txt_out)