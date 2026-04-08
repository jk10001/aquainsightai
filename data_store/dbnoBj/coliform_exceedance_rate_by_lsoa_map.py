# filename: coliform_exceedance_rate_by_lsoa_map.py
import pandas as pd
import numpy as np
import plotly.graph_objects as go

# Step 1: Load and process Anglian Water data for Coliform Bacteria exceedances
aw_file = "Anglian_Water_Domestic_Water_Quality.csv"
df_aw = pd.read_csv(aw_file, encoding="utf-8")

# Filter for Coliform Bacteria (Indicator)
df_coliform = df_aw[df_aw['Determinand'] == 'Coliform Bacteria (Indicator)'].copy()

# Define exceedance as Result > 0 given PCV=0
df_coliform['Exceedance'] = df_coliform['Result'] > 0

# Group by LSOA: total samples, count exceedances
grouped = df_coliform.groupby('LSOA').agg(
    Total_Samples=('Exceedance', 'size'),
    Num_Exceed=('Exceedance', 'sum')
).reset_index()

# Calculate exceedance rate per LSOA
grouped['Exceedance_Rate'] = grouped['Num_Exceed'] / grouped['Total_Samples']

# Keep all LSOAs with at least one sample
grouped = grouped[grouped['Total_Samples'] > 0]

# For max Result > 0, calculate max result for exceedances
max_result = df_coliform[df_coliform['Result'] > 0].groupby('LSOA')['Result'].max().reset_index()
max_result.rename(columns={'Result': 'Max_Result'}, inplace=True)

# Merge max_result into grouped
grouped = pd.merge(grouped, max_result, on='LSOA', how='left')

# Step 2: Load LSOA geography CSV with centroid coordinates
geo_csv_url = "https://open-geography-portalx-ons.hub.arcgis.com/api/download/v1/items/b8263c2364e9452483a0e5783c6fdb53/csv?layers=0"
df_geo = pd.read_csv(geo_csv_url, encoding="utf-8")

# Identify LSOA code column in geo data
lsoa_code_col = None
for col in df_geo.columns:
    if col.upper().startswith("LSOA") and ('CD' in col.upper() or 'CODE' in col.upper()):
        lsoa_code_col = col
        break

if lsoa_code_col is None and 'LSOA21CD' in df_geo.columns:
    lsoa_code_col = 'LSOA21CD'

lsoa_name_col = 'LSOA21NM'

latitude_col = None
longitude_col = None
for col in df_geo.columns:
    col_lower = col.lower()
    if 'lat' in col_lower and latitude_col is None:
        latitude_col = col
    if ('long' in col_lower or 'lon' in col_lower) and longitude_col is None:
        longitude_col = col

if latitude_col is None or longitude_col is None:
    float_cols = df_geo.select_dtypes(include='number').columns
    possible_lat = []
    possible_lon = []
    for col in float_cols:
        if df_geo[col].between(49, 61).any():
            possible_lat.append(col)
        if df_geo[col].between(-8, 2).any():
            possible_lon.append(col)
    if latitude_col is None and possible_lat:
        latitude_col = possible_lat[0]
    if longitude_col is None and possible_lon:
        longitude_col = possible_lon[0]

if lsoa_code_col is None or lsoa_name_col not in df_geo.columns or latitude_col is None or longitude_col is None:
    raise ValueError("Required columns for LSOA code, name, latitude or longitude not found in geography CSV")

# Step 3: Merge data with geography on LSOA code
df_merged = pd.merge(
    grouped,
    df_geo[[lsoa_code_col, lsoa_name_col, latitude_col, longitude_col]],
    left_on='LSOA', right_on=lsoa_code_col, how='left'
)

df_merged = df_merged.dropna(subset=[latitude_col, longitude_col])

# Step 4: Create map visualization
# Sort by exceedance rate ascending to plot higher on top
df_merged = df_merged.sort_values('Exceedance_Rate', ascending=True)

# Marker size based on total samples or constant if preferred
marker_base_size = 8
marker_size_factor = 0.1
marker_sizes = marker_base_size + marker_size_factor * df_merged['Total_Samples']
# Limit marker size to reasonable max
marker_sizes = marker_sizes.clip(upper=20)

color_scale = "YlOrRd"

lat_min = df_merged[latitude_col].min()
lat_max = df_merged[latitude_col].max()
lon_min = df_merged[longitude_col].min()
lon_max = df_merged[longitude_col].max()
lat_center = (lat_min + lat_max) / 2
lon_center = (lon_min + lon_max) / 2
lat_range = lat_max - lat_min
lon_range = lon_max - lon_min
max_range = max(lat_range, lon_range)

if max_range < 0.1:
    zoom_level = 10
elif max_range < 0.3:
    zoom_level = 9
elif max_range < 0.7:
    zoom_level = 8
elif max_range < 1.5:
    zoom_level = 7
elif max_range < 3:
    zoom_level = 6
else:
    zoom_level = 5

hover_text = (
    "LSOA Name: " + df_merged[lsoa_name_col].astype(str) + "<br>" +
    "LSOA Code: " + df_merged['LSOA'].astype(str) + "<br>" +
    "Total Samples: " + df_merged['Total_Samples'].astype(str) + "<br>" +
    "Number of exceedances: " + df_merged['Num_Exceed'].astype(str) + "<br>" +
    "Exceedance Rate: " + (df_merged['Exceedance_Rate'] * 100).map(lambda x: f"{x:.1f}%") + "<br>" +
    "Max result (number/100 ml): " + df_merged['Max_Result'].map(
        lambda x: f"{x:.1f}" if pd.notna(x) else "No Data"
    )
)

fig = go.Figure(go.Scattermapbox(
    lat=df_merged[latitude_col],
    lon=df_merged[longitude_col],
    mode='markers',
    marker=go.scattermapbox.Marker(
        size=marker_sizes,
        color=df_merged['Exceedance_Rate'],
        colorscale=color_scale,
        cmin=0,
        cmax=0.5,
        colorbar=dict(
            title="Coliform exceedance rate<br>(failures / total samples)",
            titleside="top",
            outlinewidth=0,
            ticks="outside",
            tickfont=dict(size=14),
            titlefont=dict(size=16)
        ),
        showscale=True,
        opacity=0.7
    ),
    hoverinfo='text',
    hovertext=hover_text
))

fig.update_layout(
    mapbox=dict(
        style="open-street-map",
        center=dict(lat=lat_center, lon=lon_center),
        zoom=zoom_level
    ),
    margin=dict(l=0, r=0, t=0, b=0),
    template='plotly_white',
    font=dict(size=16)
)

# Step 5: Save output files
base_name = "coliform_exceedance_rate_by_lsoa"

csv_output = base_name + ".csv"
html_output = base_name + "_map.html"
png_output = base_name + "_map.png"
py_output = base_name + ".py"

df_out = df_merged[['LSOA', lsoa_name_col, 'Total_Samples', 'Num_Exceed', 'Exceedance_Rate', latitude_col, longitude_col]].copy()
df_out.rename(columns={
    'LSOA': 'LSOA Code',
    lsoa_name_col: 'LSOA Name',
    'Total_Samples': 'Total samples',
    'Num_Exceed': 'Number of exceedances',
    'Exceedance_Rate': 'Exceedance rate (failures/total samples)',
    latitude_col: 'Latitude',
    longitude_col: 'Longitude'
}, inplace=True)

df_out.to_csv(csv_output, index=False)
fig.write_html(html_output, include_plotlyjs='cdn', full_html=True)
fig.write_image(png_output, width=1000, height=600)

print(png_output)
print(html_output)
print(csv_output)
print(py_output)