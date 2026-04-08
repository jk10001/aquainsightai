# filename: max_month_loads_BOD_TSS_TP.py
import pandas as pd
import plotly.graph_objects as go

# Load data
file_name = "WWTP_medium.xlsx"
sheet_name = "WWTP Influent Data"
df = pd.read_excel(file_name, sheet_name=sheet_name, parse_dates=["Date"])

# Restrict to calendar years 2014-2017 and sort by Date
df = df[(df["Date"].dt.year >= 2014) & (df["Date"].dt.year <= 2017)].sort_values("Date")
df["Year"] = df["Date"].dt.year

# Columns for parameters
flow_col = "Influent flow rate (m^3/d)"
params = {
    "BOD": "Influent BOD (mg/L)",
    "TSS": "Influent TSS (mg/L)",
    "TP": "Influent total phosphorous (mg/L)"
}

# Calculate the daily load kg/d for each valid flow and concentration
for key, col in params.items():
    load_col = f"{key}_daily_load_kg_d"
    valid_mask = df[flow_col].notna() & df[col].notna()
    df.loc[:, load_col] = pd.NA  # Initialize column with NA
    df.loc[valid_mask, load_col] = df.loc[valid_mask, flow_col] * df.loc[valid_mask, col] / 1000.0
    # Convert column dtype to float to enable rolling operations (convert NA to np.nan)
    df[load_col] = pd.to_numeric(df[load_col], errors='coerce')

# Calculate max 30-day rolling average loads for each year and parameter with min_periods=1 to allow sparse data
max_month_loads = {key: [] for key in params}
years = sorted(df["Year"].unique())

for year in years:
    df_year = df[df["Year"] == year]
    for key in params.keys():
        load_col = f"{key}_daily_load_kg_d"
        # Rolling window of 30 calendar days, min_periods=1 means average over available days in window
        rolling_avg = df_year[load_col].rolling(window=30, min_periods=1).mean()
        max_30d_avg = rolling_avg.max(skipna=True)
        max_month_loads[key].append(max_30d_avg)

# Calculate historical average of max month loads for each parameter
avg_max_month_loads = {
    key: pd.Series(values).mean() for key, values in max_month_loads.items()
}

# Create DataFrame with results for years and Average
result_df = pd.DataFrame({
    "Year_or_Group": years + ["Average"],
    "BOD_max_month_daily_load_kg_d": max_month_loads["BOD"] + [avg_max_month_loads["BOD"]],
    "TSS_max_month_daily_load_kg_d": max_month_loads["TSS"] + [avg_max_month_loads["TSS"]],
    "TP_max_month_daily_load_kg_d": max_month_loads["TP"] + [avg_max_month_loads["TP"]],
})

# Prepare x axis categories as strings (years and Average)
x_categories = result_df["Year_or_Group"].astype(str).tolist()

# Create bar chart traces
trace_bod = go.Bar(
    name="BOD",
    x=x_categories,
    y=result_df["BOD_max_month_daily_load_kg_d"],
    marker_color="royalblue",
    offsetgroup=0
)
trace_tss = go.Bar(
    name="TSS",
    x=x_categories,
    y=result_df["TSS_max_month_daily_load_kg_d"],
    marker_color="darkorange",
    offsetgroup=1
)
trace_tp = go.Bar(
    name="Total P",
    x=x_categories,
    y=result_df["TP_max_month_daily_load_kg_d"],
    marker_color="green",
    offsetgroup=2
)

# Create figure
fig = go.Figure(data=[trace_bod, trace_tss, trace_tp])
fig.update_layout(
    barmode='group',
    yaxis_title="Maximum month daily load (kg/d)",
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
base_filename = "max_month_loads_BOD_TSS_TP"
csv_filename = base_filename + ".csv"
html_filename = base_filename + ".html"
png_filename = base_filename + ".png"

# Save CSV of underlying data
result_df.to_csv(csv_filename, index=False)

# Save plotly HTML with requirements
fig.write_html(html_filename, include_plotlyjs='cdn', full_html=True)

# Save PNG using kaleido
fig.write_image(png_filename, width=1000, height=600)

# Output filenames
print(base_filename + ".py")
print(csv_filename)
print(html_filename)
print(png_filename)