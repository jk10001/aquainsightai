# filename: annual_avg_loads_BOD_TSS_TP.py
import pandas as pd
import plotly.graph_objects as go

# Load data
file_name = "WWTP_medium.xlsx"
sheet_name = "WWTP Influent Data"
df = pd.read_excel(file_name, sheet_name=sheet_name, parse_dates=["Date"])

# Restrict to full calendar years 2014-2017
df = df[(df["Date"].dt.year >= 2014) & (df["Date"].dt.year <= 2017)]

# Extract year for grouping
df["Year"] = df["Date"].dt.year

# Parameters
flow_col = "Influent flow rate (m^3/d)"
params = {
    "BOD": "Influent BOD (mg/L)",
    "TSS": "Influent TSS (mg/L)",
    "TP": "Influent total phosphorous (mg/L)"
}

# Calculate daily loads kg/d for each parameter with valid flow and concentration for that day
for key, col in params.items():
    # Only compute load when both flow and concentration are not null
    load_col = f"{key}_daily_load_kg_d"
    valid_mask = df[flow_col].notna() & df[col].notna()
    df[load_col] = pd.NA
    df.loc[valid_mask, load_col] = df.loc[valid_mask, flow_col] * df.loc[valid_mask, col] / 1000.0

# Compute annual mean daily load per parameter based on valid daily loads
annual_mean_loads = {}
for key in params.keys():
    load_col = f"{key}_daily_load_kg_d"
    annual_means = df.groupby("Year")[load_col].mean()
    annual_mean_loads[key] = annual_means

# Convert to DataFrame
annual_mean_loads_df = pd.DataFrame(annual_mean_loads)

# Calculate historical average (mean of annual means for 2014-2017) for each parameter
hist_avg = annual_mean_loads_df.mean().to_frame().T
hist_avg.index = ["Average"]

# Append historical average as last row
result_df = pd.concat([annual_mean_loads_df, hist_avg])

# Reset index to have Year_or_Group column
result_df = result_df.reset_index().rename(columns={
    "index": "Year_or_Group",
    "BOD": "BOD_avg_daily_load_kg_d",
    "TSS": "TSS_avg_daily_load_kg_d",
    "TP": "TP_avg_daily_load_kg_d"
})

# Prepare x axis categories as strings (years and Average)
x_categories = result_df["Year_or_Group"].astype(str).tolist()

# Create bar chart traces
trace_bod = go.Bar(
    name="BOD",
    x=x_categories,
    y=result_df["BOD_avg_daily_load_kg_d"],
    marker_color="royalblue",
    offsetgroup=0
)
trace_tss = go.Bar(
    name="TSS",
    x=x_categories,
    y=result_df["TSS_avg_daily_load_kg_d"],
    marker_color="darkorange",
    offsetgroup=1
)
trace_tp = go.Bar(
    name="Total P",
    x=x_categories,
    y=result_df["TP_avg_daily_load_kg_d"],
    marker_color="green",
    offsetgroup=2
)

# Create figure
fig = go.Figure(data=[trace_bod, trace_tss, trace_tp])
fig.update_layout(
    barmode='group',
    yaxis_title="Average annual daily load (kg/d)",
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
base_filename = "annual_avg_loads_BOD_TSS_TP"
csv_filename = base_filename + ".csv"
html_filename = base_filename + ".html"
png_filename = base_filename + ".png"

# Save CSV of underlying data
result_df.to_csv(csv_filename, index=False)

# Save plotly HTML with requirements
fig.write_html(html_filename, include_plotlyjs='cdn', full_html=True)

# Save PNG using kaleido
fig.write_image(png_filename, width=1000, height=600)

# Print filenames to terminal
print(base_filename + ".py")
print(csv_filename)
print(html_filename)
print(png_filename)