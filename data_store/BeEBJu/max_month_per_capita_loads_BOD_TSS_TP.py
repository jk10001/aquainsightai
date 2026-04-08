# filename: max_month_per_capita_loads_BOD_TSS_TP.py
import pandas as pd
import plotly.graph_objects as go

# Load data
file_name = "WWTP_medium.xlsx"
influent_sheet = "WWTP Influent Data"
pop_sheet = "population"

# Load influent data with datetime parsing
df = pd.read_excel(file_name, sheet_name=influent_sheet, parse_dates=["Date"])
df = df[(df["Date"].dt.year >= 2014) & (df["Date"].dt.year <= 2017)].sort_values("Date")
df["Year"] = df["Date"].dt.year

# Load population data and restrict
df_pop = pd.read_excel(file_name, sheet_name=pop_sheet)
df_pop = df_pop[(df_pop["Year"] >= 2014) & (df_pop["Year"] <= 2017)]

# Parameters and columns
flow_col = "Influent flow rate (m^3/d)"
params = {
    "BOD": "Influent BOD (mg/L)",
    "TSS": "Influent TSS (mg/L)",
    "TP": "Influent total phosphorous (mg/L)"
}

# Calculate daily load (kg/d) for each parameter where both flow and concentration are present
for key, conc_col in params.items():
    load_col = f"{key}_daily_load_kg_d"
    valid_mask = df[flow_col].notna() & df[conc_col].notna()
    df.loc[:, load_col] = pd.NA
    df.loc[valid_mask, load_col] = df.loc[valid_mask, flow_col] * df.loc[valid_mask, conc_col] / 1000.0
    df[load_col] = pd.to_numeric(df[load_col], errors='coerce')

# For each year and parameter compute max 30-day rolling average load (kg/d)
max_month_loads = {key: [] for key in params}
years = sorted(df["Year"].unique())

for year in years:
    df_year = df[df["Year"] == year]
    for key in params.keys():
        load_col = f"{key}_daily_load_kg_d"
        # 30-day rolling average with min_periods=1 to allow sparse data in window
        rolling_avg = df_year[load_col].rolling(window=30, min_periods=1).mean()
        max_30d_avg = rolling_avg.max(skipna=True)
        max_month_loads[key].append(max_30d_avg)

# Convert max month loads from kg/d to g/person/day using population
pop_dict = df_pop.set_index("Year")["Serviced Population"].to_dict()

max_month_percapita = {key: [] for key in params}
for year_index, year in enumerate(years):
    pop = pop_dict.get(year)
    for key in params.keys():
        max_load = max_month_loads[key][year_index]
        if pd.isna(max_load) or pop is None or pop == 0:
            max_month_percapita[key].append(pd.NA)
        else:
            # Convert kg/d to g/p/d
            max_month_percapita[key].append(max_load * 1000 / pop)

# Calculate historical averages for each parameter
avg_percapita = {key: pd.Series(vals).mean() for key, vals in max_month_percapita.items()}

# Construct result DataFrame
result_df = pd.DataFrame({
    "Year_or_Group": years + ["Average"],
    "BOD_max_month_daily_load_g_p_d": max_month_percapita["BOD"] + [avg_percapita["BOD"]],
    "TSS_max_month_daily_load_g_p_d": max_month_percapita["TSS"] + [avg_percapita["TSS"]],
    "TP_max_month_daily_load_g_p_d": max_month_percapita["TP"] + [avg_percapita["TP"]],
})

# Prepare x axis categories
x_categories = result_df["Year_or_Group"].astype(str).tolist()

# Plot bars
trace_bod = go.Bar(
    name="BOD",
    x=x_categories,
    y=result_df["BOD_max_month_daily_load_g_p_d"],
    marker_color="royalblue",
    offsetgroup=0
)
trace_tss = go.Bar(
    name="TSS",
    x=x_categories,
    y=result_df["TSS_max_month_daily_load_g_p_d"],
    marker_color="darkorange",
    offsetgroup=1
)
trace_tp = go.Bar(
    name="Total P",
    x=x_categories,
    y=result_df["TP_max_month_daily_load_g_p_d"],
    marker_color="green",
    offsetgroup=2
)

fig = go.Figure(data=[trace_bod, trace_tss, trace_tp])
fig.update_layout(
    barmode='group',
    yaxis_title="Maximum month daily load per capita (g/person/day)",
    xaxis_title="Year or Group",
    yaxis=dict(showgrid=True, zeroline=False),
    font=dict(size=16),
    template='plotly_white',
    legend=dict(
        bordercolor='black',
        borderwidth=1
    )
)

# Save results
base_filename = "max_month_per_capita_loads_BOD_TSS_TP"
csv_filename = base_filename + ".csv"
html_filename = base_filename + ".html"
png_filename = base_filename + ".png"

result_df.to_csv(csv_filename, index=False)
fig.write_html(html_filename, include_plotlyjs='cdn', full_html=True)
fig.write_image(png_filename, width=1000, height=600)

print(base_filename + ".py")
print(csv_filename)
print(html_filename)
print(png_filename)