# filename: historical_projected_avg_loads_BOD_TSS_TP_20yr.py
import pandas as pd
import plotly.graph_objects as go

# File and sheet names
influent_file = "WWTP_medium.xlsx"
influent_sheet = "WWTP Influent Data"
pop_proj_file = "population_projection_20yr.csv"
per_capita_avg_file = "per_capita_historical_averages.csv"

# Load influent data with parsing dates
df_influent = pd.read_excel(influent_file, sheet_name=influent_sheet, parse_dates=["Date"])

# Restrict to historical years 2014-2017
df_influent = df_influent[(df_influent["Date"].dt.year >= 2014) & (df_influent["Date"].dt.year <= 2017)].copy()
df_influent["Year"] = df_influent["Date"].dt.year

# Define flow and concentration columns
flow_col = "Influent flow rate (m^3/d)"
params = {
    "BOD": "Influent BOD (mg/L)",
    "TSS": "Influent TSS (mg/L)",
    "TP": "Influent total phosphorous (mg/L)"
}

# Calculate daily loads (kg/d) only where flow and concentration exist
for key, conc_col in params.items():
    load_col = f"{key}_daily_load_kg_d"
    valid_mask = df_influent[flow_col].notna() & df_influent[conc_col].notna()
    df_influent[load_col] = pd.NA
    df_influent.loc[valid_mask, load_col] = df_influent.loc[valid_mask, flow_col] * df_influent.loc[valid_mask, conc_col] / 1000.0
    df_influent[load_col] = pd.to_numeric(df_influent[load_col], errors='coerce')

# Compute historical average annual loads (kg/d) per parameter, per year
historic_avg_loads = {}
for key in params.keys():
    load_col = f"{key}_daily_load_kg_d"
    annual_avg = df_influent.groupby("Year")[load_col].mean()
    historic_avg_loads[key] = annual_avg

df_hist_avgs = pd.DataFrame(historic_avg_loads)  # index = Year

# Read population projections 2014-2037
df_pop_proj = pd.read_csv(pop_proj_file)
df_pop_proj = df_pop_proj.sort_values("Year").reset_index(drop=True)
# Filter for years 2014-2037 (the file may already be this range)
df_pop_proj = df_pop_proj[(df_pop_proj["Year"] >= 2014) & (df_pop_proj["Year"] <= 2037)].copy()

# Read per capita historical averages
df_percap = pd.read_csv(per_capita_avg_file)

# Extract g/person/day values for average day loads for BOD, TSS, TP
def get_percap_value(param_name):
    row = df_percap[df_percap["Parameter"] == f"Average day {param_name} load per capita"]
    if row.empty:
        row = df_percap[df_percap["Parameter"].str.contains(param_name) & df_percap["Parameter"].str.contains("average day", case=False, na=False)]
        if row.empty:
            raise ValueError(f"Per capita average day load for '{param_name}' not found in {per_capita_avg_file}.")
    return float(row.iloc[0]["Value"])

bod_percap_gpd = get_percap_value("BOD")
tss_percap_gpd = get_percap_value("TSS")
tp_percap_gpd  = get_percap_value("Total P")

# Prepare projected loads (kg/d) for 2018-2037 only (use per capita * population / 1000)
df_pop_proj["BOD_avg_load_kg_d_projected"] = pd.NA
df_pop_proj["TSS_avg_load_kg_d_projected"] = pd.NA
df_pop_proj["TP_avg_load_kg_d_projected"] = pd.NA

for idx, row in df_pop_proj.iterrows():
    year = row["Year"]
    pop = row["Serviced Population"]
    if year >= 2018:
        df_pop_proj.at[idx, "BOD_avg_load_kg_d_projected"] = pop * bod_percap_gpd / 1000.0
        df_pop_proj.at[idx, "TSS_avg_load_kg_d_projected"] = pop * tss_percap_gpd / 1000.0
        df_pop_proj.at[idx, "TP_avg_load_kg_d_projected"]  = pop * tp_percap_gpd  / 1000.0

# Prepare combined DataFrame for years 2014-2037
years = list(range(2014, 2038))
df_combined = pd.DataFrame({"Year": years})

# Add historical columns for 2014-2017 (NaN for others)
df_hist_avgs_reset = df_hist_avgs.reset_index()
df_combined = df_combined.merge(df_hist_avgs_reset.rename(columns={
    "BOD": "BOD_avg_load_kg_d_historical",
    "TSS": "TSS_avg_load_kg_d_historical",
    "TP": "TP_avg_load_kg_d_historical"
}), on="Year", how="left")

# Add projected columns for 2018-2037 (NaN for others)
proj_cols = ["Year", "BOD_avg_load_kg_d_projected", "TSS_avg_load_kg_d_projected", "TP_avg_load_kg_d_projected"]
df_combined = df_combined.merge(df_pop_proj[proj_cols], on="Year", how="left")

# For historical years (2014-2017), set projected load columns to NaN to avoid mixing
df_combined.loc[df_combined["Year"] <= 2017, ["BOD_avg_load_kg_d_projected", "TSS_avg_load_kg_d_projected", "TP_avg_load_kg_d_projected"]] = pd.NA
# For projected years (2018+), set historical load columns to NaN
df_combined.loc[df_combined["Year"] >= 2018, ["BOD_avg_load_kg_d_historical", "TSS_avg_load_kg_d_historical", "TP_avg_load_kg_d_historical"]] = pd.NA

# Save combined data to CSV
base_filename = "historical_projected_avg_loads_BOD_TSS_TP_20yr"
csv_filename = base_filename + ".csv"
df_combined.to_csv(csv_filename, index=False)

# Create Plotly figure
fig = go.Figure()

# Define colors for params
colors = {"BOD": "blue", "TSS": "orange", "TP": "green"}
params_list = ["BOD", "TSS", "TP"]

# Plot historical and projected lines for each param
for key in params_list:
    # Historical trace, solid line
    fig.add_trace(go.Scatter(
        x=df_combined["Year"],
        y=df_combined[f"{key}_avg_load_kg_d_historical"],
        mode="lines+markers",
        name=f"Historical {key}",
        line=dict(color=colors[key], dash="solid"),
        marker=dict(size=8)
    ))
    # Projected trace, dashed line
    fig.add_trace(go.Scatter(
        x=df_combined["Year"],
        y=df_combined[f"{key}_avg_load_kg_d_projected"],
        mode="lines+markers",
        name=f"Projected {key}",
        line=dict(color=colors[key], dash="dash"),
        marker=dict(size=8)
    ))

fig.update_layout(
    xaxis_title="Year",
    yaxis_title="Average annual load (kg/d)",
    font=dict(size=16),
    template="plotly_white",
    yaxis=dict(showgrid=True, zeroline=False),
    legend=dict(bordercolor="black", borderwidth=1)
)

# Save outputs
html_filename = base_filename + ".html"
png_filename = base_filename + ".png"
py_filename = base_filename + ".py"

fig.write_html(html_filename, include_plotlyjs='cdn', full_html=True)
fig.write_image(png_filename, width=1000, height=600)

# Print filenames to terminal
print(py_filename)
print(csv_filename)
print(html_filename)
print(png_filename)