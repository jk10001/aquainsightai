# filename: annual_bod_90th_percentile_vs_limit.py
import pandas as pd
import plotly.graph_objects as go

# Load the data
file_csv = "annual_discharge_compliance_summary_2014_2018.csv"
df = pd.read_csv(file_csv)

# Prepare the bar chart for BOD 90th percentile
fig = go.Figure()

# Blue bars for BOD 90th percentile
fig.add_trace(go.Bar(
    x=df['Year'].astype(str),
    y=df['BOD - 90th Percentile (mg/L)'],
    name='BOD - 90th Percentile (mg/L)',
    marker_color='blue'
))

# Red horizontal line for 90th percentile discharge limit
bod_90th_limit = df['BOD 90th Percentile Limit (mg/L)'].iloc[0]  # limit is same for all years
fig.add_trace(go.Scatter(
    x=df['Year'].astype(str),
    y=[bod_90th_limit]*len(df),
    mode='lines',
    name='90th Percentile Limit (10 mg/L)',
    line=dict(color='red', dash='dash')
))

# Update layout
ymax = max(df['BOD - 90th Percentile (mg/L)'].max(), bod_90th_limit) * 1.1

fig.update_layout(
    yaxis=dict(title='BOD Concentration (mg/L)', range=[0, ymax],
               gridcolor='lightgrey', zeroline=True, zerolinecolor='black'),
    xaxis=dict(title='Year'),
    legend=dict(x=0.01, y=0.99),
    template='plotly_white',
    font=dict(size=16),
    margin=dict(l=60, r=20, t=20, b=60),
)

# Save the figure to files
html_file = "annual_bod_90th_percentile_vs_limit.html"
png_file = "annual_bod_90th_percentile_vs_limit.png"
csv_file = "annual_bod_90th_percentile_vs_limit_data.csv"

fig.write_html(html_file, include_plotlyjs='cdn')
fig.write_image(png_file, width=1000, height=600)
df.to_csv(csv_file, index=False)

print(html_file)
print(png_file)
print(csv_file)