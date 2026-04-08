# filename: nickel_timeseries.py
import pandas as pd
import plotly.graph_objects as go

# Load and preprocess data
file_name = "Anglian_Water_Domestic_Water_Quality.csv"
df = pd.read_csv(file_name, encoding="utf-8")
df['Sample_Date'] = pd.to_datetime(df['Sample_Date'], format='%d/%m/%Y %H:%M', errors='coerce')

# Filter for Nickel (Total) determinand and non-null Result
df_filtered = df[(df['Determinand'] == 'Nickel (Total)') & df['Result'].notna()].copy()
df_filtered.sort_values('Sample_Date', inplace=True)

# Save filtered data for chart plotting
output_csv = "nickel_timeseries_data.csv"
df_filtered[['Sample_Date', 'Result', 'Units', 'LSOA', 'Sample_Id', 'ObjectId']].to_csv(output_csv, index=False)

# Confirm units for label (use the most common unit)
units_series = df_filtered['Units'].mode()
units_label = units_series[0] if not units_series.empty else 'µg/L'
yaxis_title = f'Nickel ({units_label})'

# UK PCV limit for Nickel (µg/L)
pcv_value = 20

# Create plotly figure
fig = go.Figure()

# Add scatter markers for nickel results
fig.add_trace(go.Scatter(
    x=df_filtered['Sample_Date'],
    y=df_filtered['Result'],
    mode='markers',
    marker=dict(
        size=6,
        symbol='circle',
        color='blue',
        opacity=0.7
    ),
    name='Sample results'
))

# Add horizontal dashed red line for PCV 20 µg/L
fig.add_trace(go.Scatter(
    x=[df_filtered['Sample_Date'].min(), df_filtered['Sample_Date'].max()],
    y=[pcv_value, pcv_value],
    mode='lines',
    line=dict(
        color='red',
        width=2,
        dash='dash'
    ),
    name='PCV 20 µg/L'
))

# Update layout with axis labels, gridlines, axis lines and styling
fig.update_layout(
    template='plotly_white',
    font=dict(size=16),
    margin=dict(l=60, r=40, t=40, b=80),
    xaxis=dict(
        title='Sample Date',
        showgrid=True,
        gridcolor='lightgrey',
        tickformat='%b %Y',
        tickangle=45,
        dtick="M1",
        ticks="outside",
        showline=True,
        linecolor='black',
        mirror=True,
        zeroline=False
    ),
    yaxis=dict(
        title=yaxis_title,
        showgrid=True,
        gridcolor='lightgrey',
        showline=True,
        linecolor='black',
        mirror=True,
        zeroline=False,
        rangemode='tozero'
    ),
    legend=dict(
        bgcolor='rgba(255,255,255,0)',
        bordercolor='rgba(0,0,0,0)',
        borderwidth=0
    )
)

# Save the plot to files
html_file = "nickel_timeseries.html"
png_file = "nickel_timeseries.png"
fig.write_html(html_file, include_plotlyjs='cdn', full_html=True)
fig.write_image(png_file, width=1000, height=600)

# Print the output filenames for terminal confirmation
print(png_file)
print(html_file)
print(output_csv)
print("nickel_timeseries.py")