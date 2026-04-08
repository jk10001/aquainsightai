# filename: daily_effluent_ammonia_2014_2018.py
import pandas as pd
import plotly.graph_objects as go

# Load the ETP Effluent Quality data
file_excel = "MWC_ETP_Data.xlsx"
df_effluent = pd.read_excel(file_excel, sheet_name="ETP Effluent Quality")

# Ensure Date is datetime and sort
df_effluent['Date'] = pd.to_datetime(df_effluent['Date'])
df_effluent = df_effluent.sort_values('Date')

# Filter data for 2014-01-01 to 2018-12-31
start_date = pd.Timestamp("2014-01-01")
end_date = pd.Timestamp("2018-12-31")
df_effluent = df_effluent[(df_effluent['Date'] >= start_date) & (df_effluent['Date'] <= end_date)]

# Discharge limits
annual_mean_limit = 0.5  # mg/L
percentile_90_limit = 2.0  # mg/L

# Create line chart with markers for Daily Effluent Ammonia
fig = go.Figure()

fig.add_trace(go.Scatter(
    x=df_effluent['Date'],
    y=df_effluent['Ammonia (mg/L)'],
    mode='lines+markers',
    name='Daily Effluent Ammonia (mg/L)',
    line=dict(color='blue'),
    marker=dict(size=6)
))

# Add horizontal line for annual mean discharge limit
fig.add_trace(go.Scatter(
    x=[df_effluent['Date'].min(), df_effluent['Date'].max()],
    y=[annual_mean_limit, annual_mean_limit],
    mode='lines',
    name='Annual Mean Limit (0.5 mg/L)',
    line=dict(color='red', dash='dash')
))

# Add horizontal line for 90th percentile discharge limit
fig.add_trace(go.Scatter(
    x=[df_effluent['Date'].min(), df_effluent['Date'].max()],
    y=[percentile_90_limit, percentile_90_limit],
    mode='lines',
    name='90th Percentile Limit (2.0 mg/L)',
    line=dict(color='orange', dash='dash')
))

# Layout update
fig.update_layout(
    yaxis=dict(
        title='Ammonia Concentration (mg/L)',
        range=[0, max(df_effluent['Ammonia (mg/L)'].max()*1.1, percentile_90_limit*1.1)],
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
html_file = "daily_effluent_ammonia_2014_2018.html"
png_file = "daily_effluent_ammonia_2014_2018.png"
csv_file = "daily_effluent_ammonia_2014_2018_data.csv"

fig.write_html(html_file, include_plotlyjs='cdn')
fig.write_image(png_file, width=1000, height=600)
df_effluent[['Date', 'Ammonia (mg/L)']].to_csv(csv_file, index=False)

print(html_file)
print(png_file)
print(csv_file)