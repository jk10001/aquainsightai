# filename: annual_ammonia_mean_90th_percentile_vs_limits.py
import pandas as pd
import plotly.graph_objects as go

# Load the data
file_csv = "annual_discharge_compliance_summary_2014_2018.csv"
df = pd.read_csv(file_csv)

# Prepare the bar chart for annual mean and 90th percentile ammonia
fig = go.Figure()

# Blue bars for annual mean ammonia
fig.add_trace(go.Bar(
    x=df['Year'].astype(str),
    y=df['Ammonia - Annual Mean (mg/L)'],
    name='Ammonia - Annual Mean (mg/L)',
    marker_color='blue'
))

# Green bars for ammonia 90th percentile
fig.add_trace(go.Bar(
    x=df['Year'].astype(str),
    y=df['Ammonia - 90th Percentile (mg/L)'],
    name='Ammonia - 90th Percentile (mg/L)',
    marker_color='green'
))

# Red horizontal line for annual mean discharge limit
amu_mean_limit = df['Ammonia Annual Mean Limit (mg/L)'].iloc[0]  # limit is same for all years
fig.add_trace(go.Scatter(
    x=df['Year'].astype(str),
    y=[amu_mean_limit]*len(df),
    mode='lines',
    name='Annual Mean Limit (0.5 mg/L)',
    line=dict(color='red', dash='dash')
))

# Orange horizontal line for 90th percentile discharge limit
amu_90th_limit = df['Ammonia 90th Percentile Limit (mg/L)'].iloc[0]
fig.add_trace(go.Scatter(
    x=df['Year'].astype(str),
    y=[amu_90th_limit]*len(df),
    mode='lines',
    name='90th Percentile Limit (2.0 mg/L)',
    line=dict(color='orange', dash='dash')
))

# Update layout
ymax = max(df['Ammonia - Annual Mean (mg/L)'].max(), df['Ammonia - 90th Percentile (mg/L)'].max(), amu_90th_limit)*1.1

fig.update_layout(
    barmode='group',
    yaxis=dict(title='Ammonia Concentration (mg/L)', range=[0, ymax],
               gridcolor='lightgrey', zeroline=True, zerolinecolor='black'),
    xaxis=dict(title='Year'),
    legend=dict(x=0.01, y=0.99),
    template='plotly_white',
    font=dict(size=16),
    margin=dict(l=60, r=20, t=20, b=60),
)

# Save the figure to files
html_file = "annual_ammonia_mean_90th_percentile_vs_limits.html"
png_file = "annual_ammonia_mean_90th_percentile_vs_limits.png"
csv_file = "annual_ammonia_mean_90th_percentile_vs_limits_data.csv"

fig.write_html(html_file, include_plotlyjs='cdn')
fig.write_image(png_file, width=1000, height=600)
df.to_csv(csv_file, index=False)

print(html_file)
print(png_file)
print(csv_file)