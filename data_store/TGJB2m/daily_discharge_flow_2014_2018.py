# filename: daily_discharge_flow_2014_2018.py
import pandas as pd
import plotly.graph_objects as go

# Load the ETP Influent Flow data
file_excel = "MWC_ETP_Data.xlsx"
df_flow = pd.read_excel(file_excel, sheet_name="ETP Influent Flow")

# Ensure Date is datetime and sort
df_flow['Date'] = pd.to_datetime(df_flow['Date'])
df_flow = df_flow.sort_values('Date')

# Filter data for 2014-01-01 to 2018-12-31
start_date = pd.Timestamp("2014-01-01")
end_date = pd.Timestamp("2018-12-31")
df_flow = df_flow[(df_flow['Date'] >= start_date) & (df_flow['Date'] <= end_date)]

# Discharge limit
discharge_limit = 540  # ML/d

# Create line chart without markers for daily influent flow (proxy discharge flow)
fig = go.Figure()

fig.add_trace(go.Scatter(
    x=df_flow['Date'],
    y=df_flow['Influent Flow (ML/d)'],
    mode='lines',
    name='Daily Influent Flow (ML/d)',
    line=dict(color='blue')
))

# Add horizontal line for discharge limit
fig.add_trace(go.Scatter(
    x=[df_flow['Date'].min(), df_flow['Date'].max()],
    y=[discharge_limit, discharge_limit],
    mode='lines',
    name='Discharge Limit (540 ML/d)',
    line=dict(color='red', dash='dash')
))

# Layout update
fig.update_layout(
    yaxis=dict(
        title='Flow Rate (ML/d)',
        range=[0, max(df_flow['Influent Flow (ML/d)'].max()*1.1, discharge_limit*1.1)],
        gridcolor='lightgrey',
        zeroline=True,
        zerolinecolor='black'
    ),
    xaxis=dict(title='Date'),
    template='plotly_white',
    font=dict(size=16),
    legend=dict(x=0.01, y=0.99),
    margin=dict(l=60, r=20, t=20, b=60),
)

# Save files
html_file = "daily_discharge_flow_2014_2018.html"
png_file = "daily_discharge_flow_2014_2018.png"
csv_file = "daily_discharge_flow_2014_2018_data.csv"

fig.write_html(html_file, include_plotlyjs='cdn')
fig.write_image(png_file, width=1000, height=600)
df_flow.to_csv(csv_file, index=False)

print(html_file)
print(png_file)
print(csv_file)