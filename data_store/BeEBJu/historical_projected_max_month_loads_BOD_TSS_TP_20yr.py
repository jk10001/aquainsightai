# filename: historical_projected_max_month_loads_BOD_TSS_TP_20yr.py
import pandas as pd
import plotly.graph_objects as go

# Filenames
base_filename = "historical_projected_max_month_loads_BOD_TSS_TP_20yr"
csv_filename = base_filename + ".csv"
html_filename = base_filename + ".html"
png_filename = base_filename + ".png"

# Load historical max-month loads
# Prefer reading from CSV if available, else fallback to calculation (not requested here)
try:
    df_hist_max = pd.read_csv("max_month_loads_BOD_TSS_TP.csv")
    # Filter to years 2014-2017 and Year_or_Group as int
    df_hist_max = df_hist_max[df_hist_max["Year_or_Group"].apply(lambda x: str(x).isdigit())].copy()
    df_hist_max["Year"] = df_hist_max["Year_or_Group"].astype(int)
    df_hist_max = df_hist_max[(df_hist_max["Year"] >= 2014) & (df_hist_max["Year"] <= 2017)]
    hist_df = df_hist_max[[
        "Year",
        "BOD_max_month_daily_load_kg_d",
        "TSS_max_month_daily_load_kg_d",
        "TP_max_month_daily_load_kg_d"
    ]].rename(columns={
        "BOD_max_month_daily_load_kg_d": "BOD_max_month_load_kg_d_historical",
        "TSS_max_month_daily_load_kg_d": "TSS_max_month_load_kg_d_historical",
        "TP_max_month_daily_load_kg_d": "TP_max_month_load_kg_d_historical"
    })
except Exception:
    # Fallback to calculating from raw data not implemented in this code (per instructions)
    raise RuntimeError("max_month_loads_BOD_TSS_TP.csv file not found or could not be loaded.")

# Load per capita historical averages: maximum month loads (g/person/day)
df_percap = pd.read_csv("per_capita_historical_averages.csv")

def get_percap_value(parameter_name):
    row = df_percap[df_percap["Parameter"] == parameter_name]
    if row.empty:
        raise ValueError(f"Parameter '{parameter_name}' not found in per_capita_historical_averages.csv")
    return float(row["Value"].values[0])

bod_percap = get_percap_value("Maximum month BOD load per capita (30-day avg)")
tss_percap = get_percap_value("Maximum month TSS load per capita (30-day avg)")
tp_percap = get_percap_value("Maximum month Total P load per capita (30-day avg)")

# Load population projection (historical + projected)
df_pop_proj = pd.read_csv("population_projection_20yr.csv")
df_pop_proj = df_pop_proj[(df_pop_proj["Year"] >= 2014) & (df_pop_proj["Year"] <= 2037)]

# Compute projected loads (kg/d) for years >= 2018
df_pop_proj["BOD_max_month_load_kg_d_projected"] = pd.NA
df_pop_proj["TSS_max_month_load_kg_d_projected"] = pd.NA
df_pop_proj["TP_max_month_load_kg_d_projected"] = pd.NA

for idx, row in df_pop_proj.iterrows():
    year = row["Year"]
    pop = row["Serviced Population"]
    if year >= 2018:
        df_pop_proj.at[idx, "BOD_max_month_load_kg_d_projected"] = pop * bod_percap / 1000.0
        df_pop_proj.at[idx, "TSS_max_month_load_kg_d_projected"] = pop * tss_percap / 1000.0
        df_pop_proj.at[idx, "TP_max_month_load_kg_d_projected"] = pop * tp_percap / 1000.0

# Merge historical and projected into one DataFrame with all years 2014-2037
years_all = pd.DataFrame({"Year": range(2014, 2038)})

df_combined = years_all.merge(hist_df, on="Year", how="left")
proj_cols = [
    "Year",
    "BOD_max_month_load_kg_d_projected",
    "TSS_max_month_load_kg_d_projected",
    "TP_max_month_load_kg_d_projected"
]
df_combined = df_combined.merge(df_pop_proj[proj_cols], on="Year", how="left")

# Ensure historical projected columns NaN for <=2017
df_combined.loc[df_combined["Year"] <= 2017, [
    "BOD_max_month_load_kg_d_projected",
    "TSS_max_month_load_kg_d_projected",
    "TP_max_month_load_kg_d_projected"
]] = pd.NA

# Ensure historical columns NaN for >=2018
df_combined.loc[df_combined["Year"] >= 2018, [
    "BOD_max_month_load_kg_d_historical",
    "TSS_max_month_load_kg_d_historical",
    "TP_max_month_load_kg_d_historical"
]] = pd.NA

# Plotting
fig = go.Figure()

colors = {"BOD": "blue", "TSS": "orange", "TP": "green"}

# Add historical lines (solid)
fig.add_trace(go.Scatter(
    x=df_combined["Year"],
    y=df_combined["BOD_max_month_load_kg_d_historical"],
    mode="lines+markers",
    name="Historical BOD",
    line=dict(color=colors["BOD"], dash="solid"),
    marker=dict(size=8)
))
fig.add_trace(go.Scatter(
    x=df_combined["Year"],
    y=df_combined["TSS_max_month_load_kg_d_historical"],
    mode="lines+markers",
    name="Historical TSS",
    line=dict(color=colors["TSS"], dash="solid"),
    marker=dict(size=8)
))
fig.add_trace(go.Scatter(
    x=df_combined["Year"],
    y=df_combined["TP_max_month_load_kg_d_historical"],
    mode="lines+markers",
    name="Historical Total P",
    line=dict(color=colors["TP"], dash="solid"),
    marker=dict(size=8)
))

# Add projected lines (dashed)
fig.add_trace(go.Scatter(
    x=df_combined["Year"],
    y=df_combined["BOD_max_month_load_kg_d_projected"],
    mode="lines+markers",
    name="Projected BOD",
    line=dict(color=colors["BOD"], dash="dash"),
    marker=dict(size=8)
))
fig.add_trace(go.Scatter(
    x=df_combined["Year"],
    y=df_combined["TSS_max_month_load_kg_d_projected"],
    mode="lines+markers",
    name="Projected TSS",
    line=dict(color=colors["TSS"], dash="dash"),
    marker=dict(size=8)
))
fig.add_trace(go.Scatter(
    x=df_combined["Year"],
    y=df_combined["TP_max_month_load_kg_d_projected"],
    mode="lines+markers",
    name="Projected Total P",
    line=dict(color=colors["TP"], dash="dash"),
    marker=dict(size=8)
))

fig.update_layout(
    xaxis_title="Year",
    yaxis_title="Maximum month load (kg/d)",
    font=dict(size=16),
    template="plotly_white",
    yaxis=dict(showgrid=True, zeroline=False),
    legend=dict(bordercolor="black", borderwidth=1)
)

# Save outputs
df_combined.to_csv(csv_filename, index=False)
fig.write_html(html_filename, include_plotlyjs='cdn', full_html=True)
fig.write_image(png_filename, width=1000, height=600)

# Print filenames to terminal
print(base_filename + ".py")
print(csv_filename)
print(html_filename)
print(png_filename)