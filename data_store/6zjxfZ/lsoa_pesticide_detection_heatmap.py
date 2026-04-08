# filename: lsoa_pesticide_detection_heatmap.py
import pandas as pd
import numpy as np
import plotly.express as px

# Main water quality data filename
water_quality_file = "Anglian_Water_Domestic_Water_Quality.csv"

# Load main data
df = pd.read_csv(water_quality_file, encoding="utf-8")

# Convert Sample_Date to datetime
df['Sample_Date'] = pd.to_datetime(df['Sample_Date'], format="%d/%m/%Y %H:%M", errors='coerce')

# Filter pesticide determinands:
mask_pesticides = df['Determinand'].str.startswith("Pesticides ", na=False) | (df['Determinand'] == "Pesticides (Total by Calculation)")
df_pesticides = df.loc[mask_pesticides].copy()

# Define detections: Operator not '<'
df_pesticides['Is_Detect'] = ~df_pesticides['Operator'].eq('<')

# Calculate detection counts per determinand
detection_counts = df_pesticides.groupby('Determinand')['Is_Detect'].sum()
valid_determinands = detection_counts[detection_counts > 0].index

# Filter only to valid detected pesticide determinands and only detects
df_valid_detections = df_pesticides[(df_pesticides['Determinand'].isin(valid_determinands)) & (df_pesticides['Is_Detect'])]

# Group by LSOA to count detections
lsoa_counts = df_valid_detections.groupby('LSOA').size().reset_index(name='detection_count')

# Load ONS LSOA centroid CSV
ons_lsoa_url = "https://open-geography-portalx-ons.hub.arcgis.com/api/download/v1/items/b8263c2364e9452483a0e5783c6fdb53/csv?layers=0"
df_lsoa = pd.read_csv(ons_lsoa_url, encoding='utf-8')

# Identify latitude and longitude column names (try common variants)
lat_cols = [col for col in df_lsoa.columns if col.lower() in {'lat', 'latitude', 'y'}]
lon_cols = [col for col in df_lsoa.columns if col.lower() in {'lon', 'long', 'longitude', 'x'}]
if not lat_cols or not lon_cols:
    raise ValueError("Latitude and/or Longitude columns not found in LSOA centroids data.")
lat_col = lat_cols[0]
lon_col = lon_cols[0]

# Select and rename to common names
df_lsoa_selected = df_lsoa[['LSOA21CD', 'LSOA21NM', lat_col, lon_col]].rename(
    columns={lat_col: 'lat', lon_col: 'lon'}
)

# Inner join detection counts with LSOA centroids on code
df_map = pd.merge(lsoa_counts, df_lsoa_selected, how='inner', left_on='LSOA', right_on='LSOA21CD')

# Sort by detection_count ascending so higher counts are plotted last (on top)
df_map = df_map.sort_values('detection_count')

# Marker size scale: detection_count mapped from [min, max] -> [5, 25]
min_size = 5
max_size = 25
min_count = df_map['detection_count'].min()
max_count = df_map['detection_count'].max()

# Avoid division by zero if all counts are same
if max_count == min_count:
    df_map['marker_size'] = (min_size + max_size) / 2
else:
    df_map['marker_size'] = min_size + (df_map['detection_count'] - min_count) / (max_count - min_count) * (max_size - min_size)

# Compute map center as the midpoint of lat and lon
center_lat = df_map['lat'].mean()
center_lon = df_map['lon'].mean()

# Compute latitude and longitude ranges to estimate zoom level
lat_range = df_map['lat'].max() - df_map['lat'].min()
lon_range = df_map['lon'].max() - df_map['lon'].min()

# Approximate zoom level based on max range - empirical formula for Mapbox zoom levels
# Larger range means smaller zoom
max_range = max(lat_range, lon_range)
if max_range == 0:
    zoom_level = 10  # arbitrary default for single point
else:
    # zoom_level decreases as max_range increases
    # Reference: zoom level approximately = 8 - log2(range), adjusted for suitable appearance
    import math
    zoom_level = 8 - np.log2(max_range)  
    if zoom_level < 3:
        zoom_level = 3  # minimum zoom to avoid zooming out too far
    elif zoom_level > 12:
        zoom_level = 12  # maximum zoom to avoid zooming in too far

# Use Plotly Mapbox scatter plot for heat map style
fig = px.scatter_mapbox(
    df_map,
    lat='lat',
    lon='lon',
    size='marker_size',
    size_max=max_size,
    color='detection_count',
    color_continuous_scale='YlOrRd',
    hover_name='LSOA21NM',
    hover_data={'LSOA21CD': True, 'detection_count': True, 'lat': False, 'lon': False, 'marker_size': False},
    zoom=zoom_level,
    height=600,
    width=1000,
    labels={'detection_count': 'Detection Count'},
    mapbox_style='open-street-map'
)

# Update layout to remove title and add black lines (axes lines are not shown on mapbox plots)
fig.update_layout(
    font=dict(size=16),
    margin={"r":0,"t":0,"l":0,"b":0},
    coloraxis_colorbar=dict(title="Detections"),
    mapbox_center=dict(lat=center_lat, lon=center_lon)
)

# Save outputs
csv_file = "lsoa_pesticide_detection_heatmap_data.csv"
html_file = "lsoa_pesticide_detection_heatmap.html"
png_file = "lsoa_pesticide_detection_heatmap.png"

df_map[['LSOA21CD', 'LSOA21NM', 'detection_count', 'lat', 'lon']].to_csv(csv_file, index=False)
fig.write_html(html_file, include_plotlyjs='cdn', full_html=True)
fig.write_image(png_file, width=1000, height=600)

print(csv_file)
print(html_file)
print(png_file)
print("lsoa_pesticide_detection_heatmap.py")