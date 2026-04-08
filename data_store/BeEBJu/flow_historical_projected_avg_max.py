# filename: flow_historical_projected_avg_max.py
import pandas as pd
import plotly.graph_objects as go

# Load population projection data
pop_proj_file = "population_historical_projected.csv"
df_pop_proj = pd.read_csv(pop_proj_file)

# Load per capita historical averages
per_capita_avg_file = "per_capita_historical_averages.csv"
df_per_capita_avg = pd.read_csv(per_capita_avg_file)

# Extract per capita average and max day flow (L/person/day)
Q_avg_pc_L_p_d = float(df_per_capita_avg.loc[df_per_capita_avg["Parameter"] == "Average annual daily flow per capita", "Value"])
Q_max_pc_L_p_d = float(df_per_capita_avg.loc[df_per_capita_avg["Parameter"] == "Maximum day flow per capita", "Value"])

# Calculate average and max day flow (m3/d) using projected population
df_pop_proj["Average_day_flow_m3_d"] = df_pop_proj["Population_projected"] * Q_avg_pc_L_p_d / 1000.0
df_pop_proj["Maximum_day_flow_m3_d"] = df_pop_proj["Population_projected"] * Q_max_pc_L_p_d / 1000.0

# Separate historical and projected years
historical_years = range(2014, 2018)
df_pop_proj["Is_historical"] = df_pop_proj["Year"].isin(historical_years)

# Create figure
fig = go.Figure()

# Average day flow traces
fig.add_trace(go.Scatter(
    x=df_pop_proj[df_pop_proj["Is_historical"]]["Year"],
    y=df_pop_proj[df_pop_proj["Is_historical"]]["Average_day_flow_m3_d"],
    mode="lines+markers",
    name="Average day flow (historical)",
    line=dict(color="blue", width=2, dash="solid"),
    marker=dict(symbol="circle", size=8)
))
fig.add_trace(go.Scatter(
    x=df_pop_proj[~df_pop_proj["Is_historical"]]["Year"],
    y=df_pop_proj[~df_pop_proj["Is_historical"]]["Average_day_flow_m3_d"],
    mode="lines+markers",
    name="Average day flow (projected)",
    line=dict(color="blue", width=2, dash="dash"),
    marker=dict(symbol="circle-open", size=8)
))

# Maximum day flow traces
fig.add_trace(go.Scatter(
    x=df_pop_proj[df_pop_proj["Is_historical"]]["Year"],
    y=df_pop_proj[df_pop_proj["Is_historical"]]["Maximum_day_flow_m3_d"],
    mode="lines+markers",
    name="Maximum day flow (historical)",
    line=dict(color="darkorange", width=2, dash="solid"),
    marker=dict(symbol="square", size=8)
))
fig.add_trace(go.Scatter(
    x=df_pop_proj[~df_pop_proj["Is_historical"]]["Year"],
    y=df_pop_proj[~df_pop_proj["Is_historical"]]["Maximum_day_flow_m3_d"],
    mode="lines+markers",
    name="Maximum day flow (projected)",
    line=dict(color="darkorange", width=2, dash="dash"),
    marker=dict(symbol="square-open", size=8)
))

# Update layout
fig.update_layout(
    xaxis_title="Year",
    yaxis_title="Flow (m³/d)",
    yaxis=dict(showgrid=True, zeroline=False),
    font=dict(size=16),
    template="plotly_white",
    legend=dict(
        bordercolor="black",
        borderwidth=1,
        bgcolor="white"
    )
)

# Save outputs
base_filename = "flow_historical_projected_avg_max"
csv_filename = base_filename + ".csv"
html_filename = base_filename + ".html"
png_filename = base_filename + ".png"

# Save CSV with required columns
df_pop_proj[["Year", "Population_projected", "Average_day_flow_m3_d", "Maximum_day_flow_m3_d"]].to_csv(csv_filename, index=False)

# Save plotly outputs
fig.write_html(html_filename, include_plotlyjs='cdn', full_html=True)
fig.write_image(png_filename, width=1000, height=600)

# Print filenames to terminal
print(base_filename + ".py")
print(csv_filename)
print(html_filename)
print(png_filename)