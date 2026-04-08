# filename: WWTP_Locations_Map_with_Treatment_Level.py
import pandas as pd
import plotly.express as px
import numpy as np

file_name = "HydroWASTE_v10_-_UTF8.csv"
df = pd.read_csv(file_name)

# Prepare data: clean DESIGN_CAP, POP_SERVED for display and handle missing DESIGN_CAP
def format_num(x):
    if pd.isna(x):
        return "N/A"
    try:
        return f"{int(x):,}"
    except Exception:
        return "N/A"

df['DESIGN_CAP_DISPLAY'] = df['DESIGN_CAP'].apply(format_num)
df['POP_SERVED_DISPLAY'] = df['POP_SERVED'].apply(lambda x: f"{int(x):,}" if not pd.isna(x) else "N/A")

# Popup info creation as html strings for hover data
df['Popup'] = (
    "WWTP: " + df['WWTP_NAME'].fillna("N/A") + "<br>" +
    "Country: " + df['COUNTRY'].fillna("N/A") + "<br>" +
    "Population served: " + df['POP_SERVED_DISPLAY'] + "<br>" +
    "Design capacity: " + df['DESIGN_CAP_DISPLAY'] + "<br>" +
    "Level: " + df['LEVEL'].fillna("N/A")
)

# Choose color map for levels
level_order = ['Primary', 'Secondary', 'Advanced']
color_discrete_map = {
    'Primary': 'blue',
    'Secondary': 'orange',
    'Advanced': 'green'
}

# Filter to only rows with valid lat/lon and level in our defined levels
df_map = df[
    df['LAT_WWTP'].notna() &
    df['LON_WWTP'].notna() &
    df['LEVEL'].isin(level_order)
].copy()

fig = px.scatter_mapbox(
    df_map,
    lat="LAT_WWTP",
    lon="LON_WWTP",
    color="LEVEL",
    color_discrete_map=color_discrete_map,
    category_orders={"LEVEL": level_order},
    hover_name="WWTP_NAME",
    hover_data={
        "COUNTRY": True,
        "POP_SERVED_DISPLAY": False,  # we add in custom tooltip
        "DESIGN_CAP_DISPLAY": False,
        "LEVEL": False,
        "LAT_WWTP": False,
        "LON_WWTP": False,
    },
    custom_data=['Popup'],
    size_max=7,
    zoom=1,
    height=600,
    width=1000,
)

# Use open source style with plotly white base and black axis lines (though Mapbox has no axis lines)
fig.update_layout(
    mapbox_style="open-street-map",
    margin={"r":0,"t":0,"l":0,"b":0},
    font=dict(size=16),
    legend_title_text="Treatment Level",
    showlegend=True,
    hovermode='closest',
)

# Override hover template to show our popup html on click
fig.update_traces(
    marker=dict(size=7, opacity=0.7),
    hovertemplate=None,
)

# Add click interaction with popups - plotly does not have native click-to-popup on mapbox,
# so use customdata in hovertemplate for the info display instead:
fig.update_traces(
    hovertemplate="%{customdata[0]}<extra></extra>",
)

# Save files
base_filename = "WWTP_Locations_Map_with_Treatment_Level"
fig.write_html(f"{base_filename}.html", include_plotlyjs='cdn')
fig.write_image(f"{base_filename}.png", scale=1, width=1000, height=600)

# Save CSV of plotted data with relevant fields
out_csv = df_map[['WWTP_NAME', 'COUNTRY', 'POP_SERVED', 'DESIGN_CAP', 'LEVEL', 'LAT_WWTP', 'LON_WWTP']]
out_csv.to_csv(f"{base_filename}.csv", index=False)

# Print filenames to terminal
print(f"{base_filename}.py")
print(f"{base_filename}.html")
print(f"{base_filename}.png")
print(f"{base_filename}.csv")