# filename: annual_mean_flow_rate_vs_limit.py
import pandas as pd
import plotly.graph_objects as go

# Load the data
file_csv = "annual_discharge_compliance_summary_2014_2018.csv"
df = pd.read_csv(file_csv)

# Prepare the bar chart with annual mean flow rate
fig = go.Figure()

# Blue bars for annual mean flow rate
fig.add_trace(go.Bar(
    x=df['Year'].astype(str),
    y=df['Flow Rate - Annual Mean (ML/d)'],
    name='Annual Mean Flow Rate (ML/d)',
    marker_color='blue'
))

# Red horizontal line for discharge limit
flow_rate_limit = df['Flow Rate Limit (ML/d)'].iloc[0]  # limit is same for all years
fig.add_trace(go.Scatter(
    x=df['Year'].astype(str),
    y=[flow_rate_limit]*len(df),
    mode='lines',
    name='Discharge Limit (540 ML/d)',
    line=dict(color='red', dash='dash')
))

# Update layout
fig.update_layout(
    yaxis=dict(title='Flow Rate (ML/d)', range=[0, max(df['Flow Rate - Annual Mean (ML/d)'].max()*1.1, flow_rate_limit*1.1)],
               gridcolor='lightgrey', zeroline=True, zerolinecolor='black'),
    xaxis=dict(title='Year'),
    legend=dict(x=0.01, y=0.99),
    template='plotly_white',
    font=dict(size=16),
    margin=dict(l=60, r=20, t=20, b=60),
)

# Save the figure to files
html_file = "annual_mean_flow_rate_vs_limit.html"
png_file = "annual_mean_flow_rate_vs_limit.png"
csv_file = "annual_mean_flow_rate_vs_limit_data.csv"

fig.write_html(html_file, include_plotlyjs='cdn')
fig.write_image(png_file, width=1000, height=600)
df.to_csv(csv_file, index=False)

print(html_file)
print(png_file)
print(csv_file)