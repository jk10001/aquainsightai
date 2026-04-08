# filename: reservoir_locations_map.py
import pandas as pd
import plotly.express as px
import plotly.io as pio
import math

file_name = "MWC_raw_water_reservoirs_2015_2019.xlsx"

# Read reservoir asset sheet
df_assets = pd.read_excel(file_name, sheet_name="WS_Storage_Dam_OpenData")

# UTM to Latitude/Longitude conversion for Southern Hemisphere based on Snyder’s formulas
# Using GRS80 ellipsoid parameters (very close to WGS84)
def utm_to_latlon(easting, northing, zone=55, southern_hemisphere=True):
    # Constants for GRS80/WGS84 ellipsoid
    a = 6378137.0  # Semi-major axis
    f = 1 / 298.257222101  # Flattening
    k0 = 0.9996  # Scale factor
    e = math.sqrt(2*f - f*f)  # Eccentricity
    
    # Remove false easting
    x = easting - 500000.0  
    # Remove false northing if southern hemisphere
    y = northing
    if southern_hemisphere:
        y -= 10000000.0

    # Calculate longitude of central meridian of zone
    lon_origin = (zone - 1)*6 - 180 + 3
    #print(f"Lon central meridian: {lon_origin}")

    # Meridional arc
    m = y / k0

    # Calculate mu
    mu = m / (a * (1 - e*e/4 - 3*e**4/64 - 5*e**6/256))

    # Calculate footprint latitude (phi1) using series expansions
    e1 = (1 - math.sqrt(1 - e*e)) / (1 + math.sqrt(1 - e*e))

    j1 = 3*e1/2 - 27*e1**3/32
    j2 = 21*e1**2/16 - 55*e1**4/32
    j3 = 151*e1**3/96
    j4 = 1097*e1**4/512

    phi1 = mu + j1*math.sin(2*mu) + j2*math.sin(4*mu) + j3*math.sin(6*mu) + j4*math.sin(8*mu)

    # Calculate parameters for latitude and longitude
    c1 = (e*e / (1 - e*e)) * (math.cos(phi1)**2)
    t1 = math.tan(phi1)**2
    n1 = a / math.sqrt(1 - e*e * math.sin(phi1)**2)
    r1 = a*(1 - e*e) / (1 - e*e * math.sin(phi1)**2)**1.5
    d = x / (n1 * k0)

    # Calculate latitude in radians
    lat_rad = (phi1 - (n1 * math.tan(phi1) / r1) * (d*d/2 - (5 + 3*t1 + 10*c1 - 4*c1*c1 - 9*(e*e/(1 - e*e))) * d**4 / 24
                + (61 + 90*t1 + 298*c1 + 45*t1*t1 - 252*(e*e/(1 - e*e)) - 3*c1*c1) * d**6 / 720))

    # Calculate longitude in radians
    lon_rad = (d - (1 + 2*t1 + c1) * d**3 / 6 
               + (5 - 2*c1 + 28*t1 - 3*c1*c1 + 8*(e*e/(1 - e*e)) + 24*t1*t1) * d**5 / 120) / math.cos(phi1)
    lon_deg = lon_origin + math.degrees(lon_rad)
    lat_deg = math.degrees(lat_rad)

    # Return lat_deg (negative for southern hemisphere), lon_deg
    return lat_deg, lon_deg

# Apply conversion to dataframe rows
def convert_row(row):
    lat, lon = utm_to_latlon(row['EASTING'], row['NORTHING'], zone=55, southern_hemisphere=True)
    return pd.Series([lat, lon])

df_assets[['latitude', 'longitude']] = df_assets.apply(convert_row, axis=1)

# Sanity check: latitudes must be negative for Victoria area
if (df_assets['latitude'] > 0).any():
    raise ValueError("Latitude > 0 detected, conversion likely incorrect for Southern Hemisphere UTM.")

# Calculate marker size, scaling by USABLE_VOLUME between 8 and 20 pts
min_size = 8
max_size = 20
min_vol = df_assets['USABLE_VOLUME'].min()
max_vol = df_assets['USABLE_VOLUME'].max()
df_assets['marker_size'] = df_assets['USABLE_VOLUME'].apply(
    lambda v: min_size + (v - min_vol) / (max_vol - min_vol) * (max_size - min_size)
)

# Prepare hover data columns
hover_cols = ['ASSET_ID', 'USABLE_VOLUME', 'GAUGE_TWL']

# Plot map using Plotly Express
fig = px.scatter_mapbox(
    df_assets,
    lat='latitude',
    lon='longitude',
    hover_name='ASSET_NAME',
    hover_data={col: True for col in hover_cols},
    size='marker_size',
    size_max=20,
    zoom=7,
    height=600,
    mapbox_style="open-street-map"
)

# Add text labels as additional scatter trace
fig.add_trace(
    px.scatter_mapbox(
        df_assets,
        lat='latitude',
        lon='longitude',
        text='ASSET_NAME'
    ).data[0]
)
fig.data[-1].update(
    mode="markers+text",
    marker=dict(size=0),
    textposition="top right",
    showlegend=False,
    hoverinfo='skip'
)

# Adjust map zoom and center to bounds of points
lat_margin = 0.05
lon_margin = 0.05
lat_min = df_assets['latitude'].min() - lat_margin
lat_max = df_assets['latitude'].max() + lat_margin
lon_min = df_assets['longitude'].min() - lon_margin
lon_max = df_assets['longitude'].max() + lon_margin
center_lat = (lat_min + lat_max) / 2
center_lon = (lon_min + lon_max) / 2

fig.update_layout(
    mapbox=dict(
        center=dict(lat=center_lat, lon=center_lon),
        zoom=7,
        style='open-street-map'
    ),
    margin=dict(l=0, r=0, t=0, b=0),
    font=dict(size=16),
    legend=dict(bordercolor='black', borderwidth=1)
)

# Save output files
base_name = "reservoir_locations_map"
csv_filename = f"{base_name}.csv"
html_filename = f"{base_name}.html"
png_filename = f"{base_name}.png"
py_filename = f"{base_name}.py"

# Save CSV with required columns plus lat/lon
df_assets_output = df_assets[['ASSET_NAME', 'ASSET_ID', 'USABLE_VOLUME', 'GAUGE_TWL', 'EASTING', 'NORTHING', 'latitude', 'longitude']]
df_assets_output.to_csv(csv_filename, index=False)

pio.write_html(fig, file=html_filename, include_plotlyjs='cdn', full_html=True, config={'displayModeBar': True})
fig.write_image(png_filename, width=1000, height=600, scale=1, engine='kaleido')

print(py_filename)
print(csv_filename)
print(html_filename)
print(png_filename)