# filename: muscatine_wrrf_digesters_2022_biogas_time_series.py
import pandas as pd
import plotly.graph_objects as go

file_name = "muscatine_wrrf_digesters_2022.xlsx"
base_name = "muscatine_wrrf_digesters_2022_biogas_time_series"

# Load workbook
xls = pd.ExcelFile(file_name)
if "daily_data" not in xls.sheet_names:
    raise ValueError("Sheet 'daily_data' not found in workbook.")

df = pd.read_excel(xls, sheet_name="daily_data")

required_cols = ["Year", "Month", "Day", "Biogas_m3d"]
missing = [c for c in required_cols if c not in df.columns]
if missing:
    raise ValueError(f"Missing required columns: {missing}")

# Build proper datetime
df["Date"] = pd.to_datetime(
    dict(year=df["Year"], month=df["Month"], day=df["Day"]),
    errors="coerce"
)

plot_df = df.loc[:, ["Date", "Biogas_m3d"]].copy()
plot_df = plot_df.sort_values("Date").reset_index(drop=True)

# Save CSV with exactly requested columns
csv_df = plot_df.rename(columns={
    "Date": "Date [YYYY-MM-DD]",
    "Biogas_m3d": "Biogas production [m3/d]",
})
csv_df["Date [YYYY-MM-DD]"] = csv_df["Date [YYYY-MM-DD]"].dt.strftime("%Y-%m-%d")
csv_file = f"{base_name}.csv"
csv_df.to_csv(csv_file, index=False)

# Reference line value from describe output mean
mean_biogas = 4716

# Create chart
fig = go.Figure()

fig.add_trace(go.Scatter(
    x=plot_df["Date"],
    y=plot_df["Biogas_m3d"],
    mode="lines+markers",
    name="Biogas production",
    line=dict(width=2),
    marker=dict(size=5)
))

x0 = plot_df["Date"].min()
x1 = plot_df["Date"].max()

fig.add_trace(go.Scatter(
    x=[x0, x1],
    y=[mean_biogas, mean_biogas],
    mode="lines",
    name="2022 mean biogas (4716 m3/d)",
    line=dict(color="firebrick", width=2, dash="dash"),
    hoverinfo="skip"
))

fig.update_layout(
    template="plotly_white",
    title=None,
    font=dict(size=16, color="black"),
    legend=dict(
        orientation="h",
        yanchor="bottom",
        y=1.02,
        xanchor="left",
        x=0
    ),
    margin=dict(l=85, r=30, t=30, b=70),
    width=1000,
    height=600
)

fig.update_xaxes(
    title_text="Date",
    showgrid=True,
    gridcolor="rgba(0,0,0,0.12)",
    showline=True,
    linecolor="black",
    mirror=True,
    ticks="outside",
    tickcolor="black",
    rangeslider_visible=False
)

fig.update_yaxes(
    title_text="Biogas production [m³/d]",
    showgrid=True,
    gridcolor="rgba(0,0,0,0.12)",
    showline=True,
    linecolor="black",
    mirror=True,
    ticks="outside",
    tickcolor="black"
)

html_file = f"{base_name}.html"
png_file = f"{base_name}.png"

fig.write_html(html_file, include_plotlyjs="cdn", full_html=True, config={"responsive": True})
fig.write_image(png_file, width=1000, height=600, engine="kaleido")

print(f"Created file: {csv_file}")
print(f"Created file: {html_file}")
print(f"Created file: {png_file}")
print("Created file: muscatine_wrrf_digesters_2022_biogas_time_series.py")