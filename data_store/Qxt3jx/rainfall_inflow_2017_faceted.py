# filename: rainfall_inflow_2017_faceted.py
import pandas as pd
import plotly.graph_objects as go
import plotly.subplots as sp
import plotly.io as pio

file_name = "MWC_raw_water_reservoirs_2015_2019.xlsx"

# Load relevant sheets
xls = pd.ExcelFile(file_name)
df_rainfall = pd.read_excel(xls, sheet_name="MajorCatchmentRainfal")
df_inflow = pd.read_excel(xls, sheet_name="MajorStorageStreamflow")

# Convert recorddate to datetime if not already
df_rainfall['recorddate'] = pd.to_datetime(df_rainfall['recorddate'])
df_inflow['recorddate'] = pd.to_datetime(df_inflow['recorddate'])

# Define target year range
start_date = pd.Timestamp("2017-01-01")
end_date = pd.Timestamp("2017-12-31")

# Filter to 2017 only
df_rainfall_2017 = df_rainfall[(df_rainfall['recorddate'] >= start_date) & (df_rainfall['recorddate'] <= end_date)]
df_inflow_2017 = df_inflow[(df_inflow['recorddate'] >= start_date) & (df_inflow['recorddate'] <= end_date)]

# Reservoirs of interest based on presence in both rainfall and inflow datasets
# Check dams in each
rainfall_dams = set(df_rainfall_2017['dam'].unique())
inflow_dams = set(df_inflow_2017['dam'].unique())
common_dams = sorted(list(rainfall_dams.intersection(inflow_dams)))

# Inner join rainfall and inflow on (recorddate, dam) to align daily data per reservoir
# Prepare inflow with simpler name for plotting
df_inflow_2017 = df_inflow_2017.rename(columns={"Streamflow into reservoir (ML/d)": "inflow_ML_d"})

merged_frames = []
for dam in common_dams:
    df_r = df_rainfall_2017[df_rainfall_2017['dam'] == dam].copy()
    df_i = df_inflow_2017[df_inflow_2017['dam'] == dam].copy()
    merged = pd.merge(df_r[['recorddate','dam','rainfall_mm']],
                      df_i[['recorddate','dam','inflow_ML_d']],
                      on=['recorddate','dam'], how='inner')
    merged_frames.append(merged)

df_plot = pd.concat(merged_frames, ignore_index=True)

# Prepare for faceted plot: one row per dam, shared x-axis
rows = len(common_dams)

# Create figure with subplots: vertically stacked with independent y-axes on left and right per subplot
fig = sp.make_subplots(
    rows=rows, cols=1, shared_xaxes=True, vertical_spacing=0.05,
    specs=[[{"secondary_y": True}] for _ in range(rows)],
    subplot_titles=common_dams
)

rainfall_color = 'lightblue'
inflow_color = 'darkblue'

for i, dam in enumerate(common_dams, start=1):
    df_sub = df_plot[df_plot['dam'] == dam]

    # Bar trace for rainfall (left y axis)
    fig.add_trace(go.Bar(
        x=df_sub['recorddate'],
        y=df_sub['rainfall_mm'],
        name='Rainfall',
        marker_color=rainfall_color,
        showlegend=(i==1),
        hovertemplate='%{x|%Y-%m-%d}<br>Rainfall: %{y} mm/day<br>',
    ), row=i, col=1, secondary_y=False)

    # Line trace for inflow (right y axis)
    fig.add_trace(go.Scatter(
        x=df_sub['recorddate'],
        y=df_sub['inflow_ML_d'],
        mode='lines',
        name='Inflow',
        line=dict(color=inflow_color, width=2),
        showlegend=(i==1),
        hovertemplate='%{x|%Y-%m-%d}<br>Inflow: %{y} ML/day<br>'
    ), row=i, col=1, secondary_y=True)

# Update axes labels
for i in range(1, rows+1):
    fig.update_yaxes(title_text="Rainfall (mm/day)", row=i, col=1, secondary_y=False,
                     showgrid=True, zeroline=False, linecolor='black', ticks='outside', mirror=True)
    fig.update_yaxes(title_text="Inflow (ML/d)", row=i, col=1, secondary_y=True,
                     showgrid=False, zeroline=False, linecolor='black', ticks='outside', mirror=True)

# X axis label on last subplot only
fig.update_xaxes(title_text="Date", row=rows, col=1,
                 showgrid=True, linecolor='black', ticks='outside', mirror=True)

# Layout updates: no overall title, font size 16, white background, black axis lines
fig.update_layout(
    height=250*rows,  # reasonable height per subplot
    template='plotly_white',
    font=dict(size=16),
    legend=dict(
        orientation='h',
        yanchor='bottom',
        y=1.02,
        xanchor='right',
        x=1,
        bordercolor='black',
        borderwidth=1,
    ),
    margin=dict(t=40,b=40,l=60,r=60)
)

# Save CSV containing date, dam, rainfall_mm, inflow_ML_d
csv_filename = "rainfall_inflow_2017_faceted.csv"
df_plot.to_csv(csv_filename, index=False)

# Save HTML & PNG
html_filename = "rainfall_inflow_2017_faceted.html"
png_filename = "rainfall_inflow_2017_faceted.png"
py_filename = "rainfall_inflow_2017_faceted.py"

pio.write_html(
    fig, file=html_filename,
    include_plotlyjs='cdn',
    full_html=True,
    config={'displayModeBar': True},
)

fig.write_image(png_filename, width=1000, height=250*rows, scale=1, engine='kaleido')

print(py_filename)
print(csv_filename)
print(html_filename)
print(png_filename)