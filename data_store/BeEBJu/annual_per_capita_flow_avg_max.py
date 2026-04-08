# filename: annual_per_capita_flow_avg_max.py
import pandas as pd
import plotly.graph_objects as go

# Load data
file_name = "WWTP_medium.xlsx"

# Load influent data
influent_sheet = "WWTP Influent Data"
df_influent = pd.read_excel(file_name, sheet_name=influent_sheet, parse_dates=["Date"])
df_influent = df_influent[(df_influent["Date"].dt.year >= 2014) & (df_influent["Date"].dt.year <= 2017)]
df_influent["Year"] = df_influent["Date"].dt.year

# Compute annual average and maximum daily flow (m3/d)
annual_flows = df_influent.groupby("Year")["Influent flow rate (m^3/d)"].agg(
    Average_annual_daily_flow_m3_d="mean",
    Maximum_day_flow_m3_d="max"
).reset_index()

# Load population data
pop_sheet = "population"
df_pop = pd.read_excel(file_name, sheet_name=pop_sheet)

# Restrict population data to matching years 2014-2017
df_pop = df_pop[(df_pop["Year"] >= 2014) & (df_pop["Year"] <= 2017)]

# Merge flows and population on Year
merged = pd.merge(annual_flows, df_pop, on="Year", how="inner")

# Calculate per capita flows in L/person/day
merged["Average_annual_daily_flow_L_per_p_d"] = merged["Average_annual_daily_flow_m3_d"] * 1000 / merged["Serviced Population"]
merged["Maximum_day_flow_L_per_p_d"] = merged["Maximum_day_flow_m3_d"] * 1000 / merged["Serviced Population"]

# Calculate historical averages for the per capita flows
avg_annual_per_capita = merged["Average_annual_daily_flow_L_per_p_d"].mean()
max_annual_per_capita = merged["Maximum_day_flow_L_per_p_d"].mean()

# Append Average row
avg_row = pd.DataFrame({
    "Year": ["Average"],
    "Average_annual_daily_flow_m3_d": [None],
    "Maximum_day_flow_m3_d": [None],
    "Serviced Population": [None],
    "Average_annual_daily_flow_L_per_p_d": [avg_annual_per_capita],
    "Maximum_day_flow_L_per_p_d": [max_annual_per_capita]
})
final_df = pd.concat([merged, avg_row], ignore_index=True)

# Prepare x axis categories as strings (years plus Average)
x_categories = final_df["Year"].astype(str).tolist()

# Create bar chart traces
trace_avg = go.Bar(
    name="Average annual daily flow per person (L/p/d)",
    x=x_categories,
    y=final_df["Average_annual_daily_flow_L_per_p_d"],
    marker_color="royalblue",
    offsetgroup=0
)
trace_max = go.Bar(
    name="Maximum day flow per person (L/p/d)",
    x=x_categories,
    y=final_df["Maximum_day_flow_L_per_p_d"],
    marker_color="darkorange",
    offsetgroup=1
)

# Create figure
fig = go.Figure(data=[trace_avg, trace_max])
fig.update_layout(
    barmode='group',
    yaxis_title="Flow per capita (L/person/day)",
    xaxis_title="Year or Group",
    yaxis=dict(showgrid=True, zeroline=False),
    font=dict(size=16),
    template='plotly_white',
    legend=dict(
        bordercolor='black',
        borderwidth=1
    )
)

# Output filenames
base_filename = "annual_per_capita_flow_avg_max"
csv_filename = base_filename + ".csv"
html_filename = base_filename + ".html"
png_filename = base_filename + ".png"

# Save CSV with relevant columns
csv_df = final_df[["Year", "Average_annual_daily_flow_L_per_p_d", "Maximum_day_flow_L_per_p_d"]].rename(columns={"Year": "Year_or_Group"})
csv_df.to_csv(csv_filename, index=False)

# Save plot as HTML and PNG
fig.write_html(html_filename, include_plotlyjs='cdn', full_html=True)
fig.write_image(png_filename, width=1000, height=600)

# Print filenames
print(base_filename + ".py")
print(csv_filename)
print(html_filename)
print(png_filename)