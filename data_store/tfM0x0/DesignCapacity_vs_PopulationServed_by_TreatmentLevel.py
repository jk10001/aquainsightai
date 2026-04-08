# filename: DesignCapacity_vs_PopulationServed_by_TreatmentLevel.py
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np

# Load data
file_name = "HydroWASTE_v10_-_UTF8.csv"
df = pd.read_csv(file_name)

# Filter data for required columns and drop rows with missing POP_SERVED or DESIGN_CAP or LEVEL
df_plot = df[['POP_SERVED', 'DESIGN_CAP', 'LEVEL']].dropna(subset=['POP_SERVED', 'DESIGN_CAP', 'LEVEL']).copy()

# Keep only accepted treatment levels
level_order = ['Primary', 'Secondary', 'Advanced']
df_plot = df_plot[df_plot['LEVEL'].isin(level_order)]

# Ensure values are positive for log scale
df_plot = df_plot[(df_plot['POP_SERVED'] > 0) & (df_plot['DESIGN_CAP'] > 0)]

# Define colors consistent with prior map
color_discrete_map = {
    'Primary': 'blue',
    'Secondary': 'orange',
    'Advanced': 'green'
}

# Create scatter plot with log scale axes
fig = px.scatter(
    df_plot,
    x='POP_SERVED',
    y='DESIGN_CAP',
    color='LEVEL',
    category_orders={"LEVEL": level_order},
    color_discrete_map=color_discrete_map,
    opacity=0.6,
    labels={
        'POP_SERVED': "Population Served (persons)",
        'DESIGN_CAP': "Design Capacity (persons)",
        'LEVEL': "Treatment Level"
    },
    height=600,
    width=1000,
)

fig.update_xaxes(type="log", showgrid=True, gridwidth=1, gridcolor='lightgrey')
fig.update_yaxes(type="log", showgrid=True, gridwidth=1, gridcolor='lightgrey')

# Add diagonal 1:1 reference line (capacity = population served)
min_val = min(df_plot['POP_SERVED'].min(), df_plot['DESIGN_CAP'].min())
max_val = max(df_plot['POP_SERVED'].max(), df_plot['DESIGN_CAP'].max())
line_values = np.logspace(np.log10(min_val), np.log10(max_val), num=500)

fig.add_trace(go.Scatter(
    x=line_values,
    y=line_values,
    mode='lines',
    line=dict(color='black', dash='dash'),
    name='1:1 Capacity = Population'
))

# Layout adjustments for clarity
fig.update_layout(
    font=dict(size=16),
    legend_title_text="Treatment Level",
    margin=dict(l=60, r=20, t=20, b=60),
    hovermode='closest',
)

# Save files
base_filename = "DesignCapacity_vs_PopulationServed_by_TreatmentLevel"
fig.write_html(f"{base_filename}.html", include_plotlyjs='cdn')
fig.write_image(f"{base_filename}.png", width=1000, height=600, scale=1)

# Save the data used in chart to CSV
df_plot.to_csv(f"{base_filename}.csv", index=False)

# Print filenames for terminal
print(f"{base_filename}.py")
print(f"{base_filename}.html")
print(f"{base_filename}.png")
print(f"{base_filename}.csv")