# filename: lsoa_exceedance_heatmap.py
import pandas as pd
import numpy as np
import plotly.graph_objects as go

def map_pesticide(determinand):
    pesticides_specific = {
        'Pesticides Aldrin': ('Aldrin', 0.030, 'µg/l', True),
        'Pesticides Dieldrin (Total)': ('Dieldrin', 0.030, 'µg/l', True),
        'Pesticides Heptachlor (Total)': ('Heptachlor', None, 'µg/l', True),
        'Pesticides Heptachlor Epoxide - Total (Trans, CIS) (Heptachlor Epoxide)': ('Heptachlor epoxide', None, 'µg/l', True)
    }
    if determinand in pesticides_specific:
        return pesticides_specific[determinand]
    if determinand and determinand.startswith('Pesticides '):
        return ('Other pesticides', 0.10, 'µg/l', True)
    return None

def get_mapping(determinand):
    mapping = {
        'Coliform Bacteria (Indicator)': ('Coliform bacteria', 0, 'number/100ml', True),
        'E.Coli (faecal coliforms Confirmed)': ('Escherichia coli', 0, 'number/100ml', True),
        'Enterococci (Confirmed)': ('Enterococci', 0, 'number/100ml', True),
        'Clostridum Perfringens (Sulphite-reducing Clostridia) (Confirmed)': ('Clostridium perfringens', 0, 'Number/100ml', True),

        'Aluminium (Total)': ('Aluminium', 200, 'µg/l', True),
        'Iron (Total)': ('Iron', 200, 'µg/l', True),
        'Manganese (Total)': ('Manganese', 50, 'µg/l', True),
        'Nitrate (Total)': ('Nitrate', 50, 'mgNO3/l', True),
        "Nitrite - Consumer's Taps": ('Nitrite', 0.50, 'mgNO2/l', True),
        'Nitrite/Nitrate formula': ('Nitrate+Nitrite formula', 1.0, '', True),
        'Conductivity (Electrical Conductivity)': ('Conductivity', 2500, 'µS/cm @ 20°C', True),
        'Sodium (Total)': ('Sodium', 200, 'mg/l', True),
        'Chloride': ('Chloride', 250, 'mg/l', True),
        'Sulphate': ('Sulphate', 250, 'mg/l', True),
        'Boron': ('Boron', 1.0, 'mg/l', True),
        'Fluoride (Total)': ('Fluoride', 1.5, 'mg/l', True),
        'Cyanide (Total)': ('Cyanide', 50, 'µg/l', True),
        'Mercury (Total)': ('Mercury', 1.0, 'µg/l', True),
        'Nickel (Total)': ('Nickel', 20, 'µg/l', True),
        'Lead (10 - will apply 25.12.2013)': ('Lead', 10, 'µg/l', True),
        'Antimony': ('Antimony', 5.0, 'µg/l', True),
        'Selenium (Total)': ('Selenium', 10, 'µg/l', True),
        'Copper (Total)': ('Copper', 2.0, 'mg/l', True),
        'Chromium (Total)': ('Chromium', 50, 'µg/l', True),

        'Bromate': ('Bromate', 10, 'µg/l', True),
        'Trihalomethanes (Total by Calculation)': ('THM (total)', 100, 'µg/l', True),

        'Gross Alpha': ('gross alpha', 0.1, 'Bq/l', True),
        'Gross Beta': ('gross beta', 1, 'Bq/l', True),
        'Tritium': ('tritium', 100, 'Bq/l', True),

        'Polycyclic Aromatic Hydrocarbons (Total by Calculation)': ('PAH', 0.10, 'µg/l', True),
        'Benzo[a]Pyrene (Total)': ('benzo[a]pyrene', 0.010, 'µg/l', True),

        'Pesticides (Total by Calculation)': ('Pesticides: total', 0.50, 'µg/l', True),
    }
    if determinand in mapping:
        return mapping[determinand]
    pesticide_map = map_pesticide(determinand)
    if pesticide_map is not None:
        return pesticide_map
    return (None, None, None, False)

file_name = "Anglian_Water_Domestic_Water_Quality.csv"
df = pd.read_csv(file_name, encoding="utf-8")

df = df[df['Determinand'].notna() & df['Result'].notna() & df['LSOA'].notna()].copy()

df[['Reg_Determinand_Name', 'Reg_PCV_Value', 'Reg_PCV_Units', 'Is_Numeric_PCV']] = df['Determinand'].apply(
    lambda x: pd.Series(get_mapping(x))
)

df_numeric = df[df['Is_Numeric_PCV'] == True].copy()

def is_exceedance(row):
    if not row['Is_Numeric_PCV']:
        return False
    pcv = row['Reg_PCV_Value']
    result = row['Result']
    if pcv == 0:
        return result > 0
    if pcv is None:
        return False
    return result > pcv

df_numeric['Is_Exceedance'] = df_numeric.apply(is_exceedance, axis=1)

agg = df_numeric.groupby('LSOA').agg(
    Total_Exceedances = ('Is_Exceedance', 'sum'),
    Total_Samples_Assessed = ('Is_Exceedance', 'count')
).reset_index()

agg_exceed = agg[agg['Total_Exceedances'] > 0].copy()

exceed_csv = "exceedances_by_lsoa.csv"
agg.to_csv(exceed_csv, index=False)

geo_url = "https://open-geography-portalx-ons.hub.arcgis.com/api/download/v1/items/b8263c2364e9452483a0e5783c6fdb53/csv?layers=0"
geo_df = pd.read_csv(geo_url, encoding="utf-8")

lon_candidates = ['LONG', 'LONGITUDE', 'Longitude', 'Easting', 'x', 'X', 'long']
lat_candidates = ['LAT', 'LATITUDE', 'Latitude', 'Northing', 'y', 'Y', 'lat']

def find_column(candidates, df):
    for c in candidates:
        if c in df.columns:
            return c
    return None

lon_col = find_column(lon_candidates, geo_df)
lat_col = find_column(lat_candidates, geo_df)
if lon_col is None or lat_col is None:
    raise ValueError("Longitude or Latitude column not found in LSOA geography data.")

joined = agg_exceed.merge(geo_df, left_on='LSOA', right_on='LSOA21CD', how='left')
joined = joined[joined[lon_col].notnull() & joined[lat_col].notnull()]

plot_df = pd.DataFrame({
    'LSOA_code': joined['LSOA21CD'],
    'LSOA_name': joined['LSOA21NM'],
    'Longitude': joined[lon_col],
    'Latitude': joined[lat_col],
    'Total_Exceedances': joined['Total_Exceedances'],
    'Total_Samples_Assessed': joined['Total_Samples_Assessed'],
})

mapdata_csv = "lsoa_exceedance_map_data.csv"
plot_df.to_csv(mapdata_csv, index=False)

plot_df_sorted = plot_df.sort_values('Total_Exceedances')

def scale_marker_size(exceedances):
    base = 8
    scale = 5
    return base + scale*np.sqrt(exceedances)

plot_df_sorted['Marker_Size'] = plot_df_sorted['Total_Exceedances'].apply(scale_marker_size)

center_lon = plot_df_sorted['Longitude'].mean()
center_lat = plot_df_sorted['Latitude'].mean()

lon_span = plot_df_sorted['Longitude'].max() - plot_df_sorted['Longitude'].min()
lat_span = plot_df_sorted['Latitude'].max() - plot_df_sorted['Latitude'].min()

zoom = 7
if lon_span > 2.5 or lat_span > 2.5:
    zoom = 6
elif lon_span < 1.0 and lat_span < 1.0:
    zoom = 8

fig = go.Figure()

fig.add_trace(go.Scattermapbox(
    lon=plot_df_sorted['Longitude'],
    lat=plot_df_sorted['Latitude'],
    mode='markers',
    marker=go.scattermapbox.Marker(
        size=plot_df_sorted['Marker_Size'],
        color=plot_df_sorted['Total_Exceedances'],
        colorscale='Viridis',
        cmin=plot_df_sorted['Total_Exceedances'].min(),
        cmax=plot_df_sorted['Total_Exceedances'].max(),
        colorbar=dict(
            title='Total exceedances',
            titleside='right',
            outlinecolor='black',
            outlinewidth=1,
            ticks='outside',
            ticklen=5,
            nticks=6,
            thickness=15,
            len=0.7,
        ),
        sizemode='area',
        sizemin=6,
        showscale=True,
        opacity=0.75
    ),
    text=(
        "LSOA: " + plot_df_sorted['LSOA_name'] + "<br>" +
        "Code: " + plot_df_sorted['LSOA_code'] + "<br>" +
        "Exceedances: " + plot_df_sorted['Total_Exceedances'].astype(str) + "<br>" +
        "Samples: " + plot_df_sorted['Total_Samples_Assessed'].astype(str) + "<br>" +
        "Pct: " + (100*plot_df_sorted['Total_Exceedances']/plot_df_sorted['Total_Samples_Assessed']).round(1).astype(str) + "%"
    ),
    hoverinfo='text',
    name='Exceedances',
))

fig.update_layout(
    mapbox=dict(
        style='carto-positron',
        center=dict(lon=center_lon, lat=center_lat),
        zoom=zoom,
    ),
    margin=dict(l=0, r=0, t=0, b=0),
    template='plotly_white',
    font=dict(size=16),
    legend=dict(
        bgcolor='rgba(255,255,255,0)',
        bordercolor='rgba(0,0,0,0)',
        borderwidth=0
    ),
)

html_file = "lsoa_exceedance_heatmap.html"
png_file = "lsoa_exceedance_heatmap.png"

fig.write_html(html_file, include_plotlyjs='cdn', full_html=True)
fig.write_image(png_file, width=1000, height=800)

print(png_file)
print(html_file)
print(mapdata_csv)
print(exceed_csv)
print("lsoa_exceedance_heatmap.py")