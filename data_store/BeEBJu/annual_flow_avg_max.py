# filename: annual_flow_avg_max.py
import pandas as pd
import plotly.graph_objects as go

# Load data
file_name = "WWTP_medium.xlsx"
sheet_name = "WWTP Influent Data"
df = pd.read_excel(file_name, sheet_name=sheet_name, parse_dates=["Date"])

# Restrict to complete calendar years 2014-2017
df = df[(df["Date"].dt.year >= 2014) & (df["Date"].dt.year <= 2017)]

# Extract year for grouping
df["Year"] = df["Date"].dt.year

# Group by year: calculate average daily flow and max daily flow
grouped = df.groupby("Year")["Influent flow rate (m^3/d)"].agg(
    Average_annual_daily_flow_m3_per_d="mean",
    Maximum_day_flow_m3_per_d="max"
).reset_index()

# Calculate historical averages of each metric
avg_annual_flow_hist = grouped["Average_annual_daily_flow_m3_per_d"].mean()
avg_max_day_flow_hist = grouped["Maximum_day_flow_m3_per_d"].mean()

# Append "Average" row
avg_row = pd.DataFrame({
    "Year": ["Average"],
    "Average_annual_daily_flow_m3_per_d": [avg_annual_flow_hist],
    "Maximum_day_flow_m3_per_d": [avg_max_day_flow_hist]
})
result_df = pd.concat([grouped, avg_row], ignore_index=True)

# Prepare x axis categories as strings (years and Average)
x_categories = result_df["Year"].astype(str).tolist()

# Prepare bar chart traces for average and max flows
trace_avg = go.Bar(
    name="Average annual daily flow (m³/d)",
    x=x_categories,
    y=result_df["Average_annual_daily_flow_m3_per_d"],
    marker_color="royalblue",
    offsetgroup=0
)
trace_max = go.Bar(
    name="Maximum day flow (m³/d)",
    x=x_categories,
    y=result_df["Maximum_day_flow_m3_per_d"],
    marker_color="darkorange",
    offsetgroup=1
)

# Create figure
fig = go.Figure(data=[trace_avg, trace_max])
fig.update_layout(
    barmode='group',
    yaxis_title="Flow (m³/d)",
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
base_filename = "annual_flow_avg_max"
csv_filename = base_filename + ".csv"
html_filename = base_filename + ".html"
png_filename = base_filename + ".png"

# Save CSV of underlying data
csv_df = result_df.rename(columns={"Year": "Year_or_Group"})
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