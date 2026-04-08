# filename: pesticide_detections_timeseries.py
import pandas as pd
import plotly.express as px

# Load data
file_name = "Anglian_Water_Domestic_Water_Quality.csv"
df = pd.read_csv(file_name, encoding="utf-8")

# Convert Sample_Date to datetime
df['Sample_Date'] = pd.to_datetime(df['Sample_Date'], format="%d/%m/%Y %H:%M", errors='coerce')

# Filter pesticide determinands: only those starting with "Pesticides " to exclude total by calculation
mask_pesticides = df['Determinand'].str.startswith("Pesticides ", na=False)
# Remove the "Pesticides (Total by Calculation)" series explicitly
df_pesticides = df.loc[mask_pesticides & (df['Determinand'] != "Pesticides (Total by Calculation)")].copy()

# Define detections: Operator not '<'
df_pesticides['Is_Detect'] = ~df_pesticides['Operator'].eq('<')

# Find determinands with at least one detection (n_detects > 0)
detection_counts = df_pesticides.groupby('Determinand')['Is_Detect'].sum()
valid_determinands = detection_counts[detection_counts > 0].index

# Filter to only include detected pesticide determinands
df_detections = df_pesticides[df_pesticides['Determinand'].isin(valid_determinands) & df_pesticides['Is_Detect']]

# Keep only needed columns for plotting
df_plot = df_detections[['Sample_Date', 'Determinand', 'Result', 'Units', 'DWI_Code', 'LSOA']].copy()

# Save CSV file containing data plotted
csv_output = "pesticide_detections_timeseries.csv"
df_plot.to_csv(csv_output, index=False)

# Create Plotly scatter plot with markers only, color by Determinand
fig = px.scatter(
    df_plot,
    x='Sample_Date',
    y='Result',
    color='Determinand',
    labels={'Sample_Date': 'Sample Date', 'Result': 'Result'},
    hover_data={'Units': True, 'DWI_Code': True, 'LSOA': True},
    title='',  # no title as requested
    template='plotly_white',
)

# Update layout for gridlines, font, and black axis lines/tick marks
fig.update_layout(
    font=dict(size=16),
    xaxis=dict(showgrid=True, gridcolor='lightgrey', linecolor='black', mirror=True, ticks='outside'),
    yaxis=dict(showgrid=True, gridcolor='lightgrey', linecolor='black', mirror=True, ticks='outside'),
    legend_title_text='Determinand',
)

# Resize plotly image dimensions when saved png
img_output = "pesticide_detections_timeseries.png"
html_output = "pesticide_detections_timeseries.html"

# Save figure html with plotlyjs='cdn' and responsive
fig.write_html(html_output, include_plotlyjs='cdn', full_html=True)

# Save figure to png with size 1000x600
fig.write_image(img_output, width=1000, height=600)

print(csv_output)
print(html_output)
print(img_output)
print("pesticide_detections_timeseries.py")