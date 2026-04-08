# filename: enterococci_timeseries.py
import pandas as pd
import plotly.graph_objects as go

# Load and preprocess data
file_name = "Anglian_Water_Domestic_Water_Quality.csv"
df = pd.read_csv(file_name, encoding="utf-8")
df['Sample_Date'] = pd.to_datetime(df['Sample_Date'], format='%d/%m/%Y %H:%M', errors='coerce')

# Filter to Enterococci (Confirmed) with non-null Result
df_filtered = df[(df['Determinand'] == 'Enterococci (Confirmed)') & df['Result'].notna()].copy()

# Sort by Sample_Date for time series plotting
df_filtered.sort_values('Sample_Date', inplace=True)

# Save filtered data for chart plotting
output_csv = "enterococci_timeseries_data.csv"
df_filtered[['Sample_Date', 'Result', 'LSOA', 'Sample_Id', 'ObjectId']].to_csv(output_csv, index=False)

# Create Plotly figure for time series scatter plot
fig = go.Figure()

# Add markers for filtered Enterococci results
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

# Layout settings: titles, axis labels, gridlines, ticks formatting
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
        zeroline=False,
    ),
    yaxis=dict(
        title='Enterococci (number/100 ml)',
        showgrid=True,
        gridcolor='lightgrey',
        showline=True,
        linecolor='black',
        mirror=True,
        zeroline=False,
    ),
    legend=dict(
        bgcolor='rgba(255,255,255,0)',
        bordercolor='rgba(0,0,0,0)',
        borderwidth=0
    )
)

# Save chart as HTML and PNG with proper size and plotly settings
html_file = "enterococci_timeseries.html"
png_file = "enterococci_timeseries.png"
fig.write_html(html_file, include_plotlyjs='cdn', full_html=True)
fig.write_image(png_file, width=1000, height=600)

# Output filenames to console
print(png_file)
print(html_file)
print(output_csv)
print("enterococci_timeseries.py")