# filename: annual_per_capita_avg_loads_BOD_TSS_TP.py
import pandas as pd
import plotly.graph_objects as go

# Load data
file_name = "WWTP_medium.xlsx"
influent_sheet = "WWTP Influent Data"
pop_sheet = "population"

# Load influent data
df_influent = pd.read_excel(file_name, sheet_name=influent_sheet, parse_dates=["Date"])
df_influent = df_influent[(df_influent["Date"].dt.year >= 2014) & (df_influent["Date"].dt.year <= 2017)]
df_influent["Year"] = df_influent["Date"].dt.year

# Load population data and restrict to years 2014-2017
df_pop = pd.read_excel(file_name, sheet_name=pop_sheet)
df_pop = df_pop[(df_pop["Year"] >= 2014) & (df_pop["Year"] <= 2017)]

# Define columns for parameters
flow_col = "Influent flow rate (m^3/d)"
params = {
    "BOD": "Influent BOD (mg/L)",
    "TSS": "Influent TSS (mg/L)",
    "TP": "Influent total phosphorous (mg/L)"
}

# Calculate daily load (kg/d) for each parameter - only where flow and concentration are both non-null
for key, conc_col in params.items():
    load_col = f"{key}_daily_load_kg_d"
    valid_mask = df_influent[flow_col].notna() & df_influent[conc_col].notna()
    df_influent[load_col] = pd.NA
    df_influent.loc[valid_mask, load_col] = df_influent.loc[valid_mask, flow_col] * df_influent.loc[valid_mask, conc_col] / 1000.0
    df_influent[load_col] = pd.to_numeric(df_influent[load_col], errors='coerce')

# Calculate annual average daily loads (kg/d) for each parameter
annual_avg_loads = {}
for key in params.keys():
    load_col = f"{key}_daily_load_kg_d"
    annual_means = df_influent.groupby("Year")[load_col].mean()
    annual_avg_loads[key] = annual_means

annual_avg_loads_df = pd.DataFrame(annual_avg_loads)

# Merge with population data to calculate per capita loads (g/person/day)
merged = annual_avg_loads_df.merge(df_pop, left_index=True, right_on="Year")

# Convert loads from kg/d to g/p/d by dividing by population and multiplying by 1000
for key in params.keys():
    merged[f"{key}_avg_daily_load_g_p_d"] = (merged[key] * 1000) / merged["Serviced Population"]

# Calculate historical averages for each parameter
historical_avg = {
    f"{key}_avg_daily_load_g_p_d": merged[f"{key}_avg_daily_load_g_p_d"].mean()
    for key in params.keys()
}

# Append "Average" row
avg_row = pd.DataFrame({
    "Year": ["Average"],
    "BOD": [None],
    "TSS": [None],
    "TP": [None],
    "Serviced Population": [None],
    "BOD_avg_daily_load_g_p_d": [historical_avg["BOD_avg_daily_load_g_p_d"]],
    "TSS_avg_daily_load_g_p_d": [historical_avg["TSS_avg_daily_load_g_p_d"]],
    "TP_avg_daily_load_g_p_d": [historical_avg["TP_avg_daily_load_g_p_d"]]
})

final_df = pd.concat([merged, avg_row], ignore_index=True)

# Prepare x axis categories as strings
x_categories = final_df["Year"].astype(str).tolist()

# Create bar chart traces for BOD, TSS, TP
trace_bod = go.Bar(
    name="BOD",
    x=x_categories,
    y=final_df["BOD_avg_daily_load_g_p_d"],
    marker_color="royalblue",
    offsetgroup=0
)
trace_tss = go.Bar(
    name="TSS",
    x=x_categories,
    y=final_df["TSS_avg_daily_load_g_p_d"],
    marker_color="darkorange",
    offsetgroup=1
)
trace_tp = go.Bar(
    name="Total P",
    x=x_categories,
    y=final_df["TP_avg_daily_load_g_p_d"],
    marker_color="green",
    offsetgroup=2
)

# Create figure
fig = go.Figure(data=[trace_bod, trace_tss, trace_tp])
fig.update_layout(
    barmode='group',
    yaxis_title="Average annual daily load per capita (g/person/day)",
    xaxis_title="Year or Group",
    yaxis=dict(showgrid=True, zeroline=False),
    font=dict(size=16),
    template='plotly_white',
    legend=dict(
        bordercolor='black',
        borderwidth=1
    )
)

# Save outputs
base_filename = "annual_per_capita_avg_loads_BOD_TSS_TP"
csv_filename = base_filename + ".csv"
html_filename = base_filename + ".html"
png_filename = base_filename + ".png"

# Save CSV of underlying data with renamed columns
csv_df = final_df[["Year", "BOD_avg_daily_load_g_p_d", "TSS_avg_daily_load_g_p_d", "TP_avg_daily_load_g_p_d"]].rename(columns={"Year": "Year_or_Group"})
csv_df.to_csv(csv_filename, index=False)

# Save plotly HTML with requirements
fig.write_html(html_filename, include_plotlyjs='cdn', full_html=True)

# Save PNG using kaleido
fig.write_image(png_filename, width=1000, height=600)

# Print filenames to terminal
print(base_filename + ".py")
print(csv_filename)
print(html_filename)
print(png_filename)