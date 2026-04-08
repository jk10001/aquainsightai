# filename: historical_projected_flows_20yr.py
import pandas as pd
import plotly.graph_objects as go

# Load data from WWTP_medium.xlsx
file_name = "WWTP_medium.xlsx"
influent_sheet = "WWTP Influent Data"
pop_sheet = "population"

# Load influent data with date parsing
df_influent = pd.read_excel(file_name, sheet_name=influent_sheet, parse_dates=["Date"])
df_influent = df_influent[(df_influent["Date"].dt.year >= 2014) & (df_influent["Date"].dt.year <= 2017)]
df_influent["Year"] = df_influent["Date"].dt.year

# Group by year: compute actual historical average and max day flow (2014-2017)
grouped = df_influent.groupby("Year")["Influent flow rate (m^3/d)"].agg(
    Avg_flow_historical_m3_d="mean",
    Max_flow_historical_m3_d="max"
).reset_index()

# Load historical-average per capita flow rates (single values)
df_percap = pd.read_csv("per_capita_historical_averages.csv")
# Extract single average per capita flow rates for average day and max day flows
avg_percap_rate = df_percap.loc[df_percap["Parameter"] == "Average annual daily flow per capita", "Value"].values[0]
max_percap_rate = df_percap.loc[df_percap["Parameter"] == "Maximum day flow per capita", "Value"].values[0]
# Units are L/person/day for both

# Load population projection data (2014-2037)
df_pop_proj = pd.read_csv("population_projection_20yr.csv")

# Prepare projected flows for years 2018-2037 only based on per capita averages
df_pop_proj["Avg_flow_projected_m3_d"] = pd.NA
df_pop_proj["Max_flow_projected_m3_d"] = pd.NA

for idx, row in df_pop_proj.iterrows():
    year = row["Year"]
    pop = row["Serviced Population"]
    if year >= 2018:
        df_pop_proj.at[idx, "Avg_flow_projected_m3_d"] = pop * avg_percap_rate / 1000.0
        df_pop_proj.at[idx, "Max_flow_projected_m3_d"] = pop * max_percap_rate / 1000.0

# Merge historical actual flow data with population projections on Year
# First create a dataframe with all years of interest 2014-2037
years_all = pd.DataFrame({"Year": range(2014, 2038)})

# Left join historical flows
df_all = years_all.merge(grouped, on="Year", how="left")

# Join population projections with projected flows
df_all = df_all.merge(df_pop_proj[["Year", "Avg_flow_projected_m3_d", "Max_flow_projected_m3_d"]], on="Year", how="left")

# Historical projected fields should be NaN for 2014-2017, so ensure that explicitly
df_all.loc[df_all["Year"] <= 2017, ["Avg_flow_projected_m3_d", "Max_flow_projected_m3_d"]] = pd.NA
# Historical actual fields should be NaN for 2018+ (no actuals)
df_all.loc[df_all["Year"] >= 2018, ["Avg_flow_historical_m3_d", "Max_flow_historical_m3_d"]] = pd.NA

# Prepare plot
fig = go.Figure()

# Historical average day flow trace (2014-2017)
hist_avg = df_all.dropna(subset=["Avg_flow_historical_m3_d"])
fig.add_trace(go.Scatter(
    x=hist_avg["Year"],
    y=hist_avg["Avg_flow_historical_m3_d"],
    mode="lines+markers",
    name="Historical average day flow",
    line=dict(color="blue", dash="solid"),
    marker=dict(symbol="circle", size=8)
))

# Projected average day flow trace (2018-2037)
proj_avg = df_all.dropna(subset=["Avg_flow_projected_m3_d"])
fig.add_trace(go.Scatter(
    x=proj_avg["Year"],
    y=proj_avg["Avg_flow_projected_m3_d"],
    mode="lines+markers",
    name="Projected average day flow",
    line=dict(color="blue", dash="dash"),
    marker=dict(symbol="circle", size=8)
))

# Historical max day flow trace (2014-2017)
hist_max = df_all.dropna(subset=["Max_flow_historical_m3_d"])
fig.add_trace(go.Scatter(
    x=hist_max["Year"],
    y=hist_max["Max_flow_historical_m3_d"],
    mode="lines+markers",
    name="Historical maximum day flow",
    line=dict(color="red", dash="solid"),
    marker=dict(symbol="circle", size=8)
))

# Projected max day flow trace (2018-2037)
proj_max = df_all.dropna(subset=["Max_flow_projected_m3_d"])
fig.add_trace(go.Scatter(
    x=proj_max["Year"],
    y=proj_max["Max_flow_projected_m3_d"],
    mode="lines+markers",
    name="Projected maximum day flow",
    line=dict(color="red", dash="dash"),
    marker=dict(symbol="circle", size=8)
))

# Layout updates
fig.update_layout(
    xaxis_title="Year",
    yaxis_title="Flow (m³/d)",
    template="plotly_white",
    font=dict(size=16),
    yaxis=dict(showgrid=True, zeroline=False),
    legend=dict(bordercolor="black", borderwidth=1)
)

# Save outputs
base_filename = "historical_projected_flows_20yr"
csv_filename = base_filename + ".csv"
html_filename = base_filename + ".html"
png_filename = base_filename + ".png"

# Save CSV with all relevant data columns
df_all.to_csv(csv_filename, index=False)

# Save plotly HTML (with Plotly JS CDN, full html)
fig.write_html(html_filename, include_plotlyjs="cdn", full_html=True)

# Save PNG with kaleido (1000 x 600)
fig.write_image(png_filename, width=1000, height=600)

# Print filenames to terminal
print(base_filename + ".py")
print(csv_filename)
print(html_filename)
print(png_filename)