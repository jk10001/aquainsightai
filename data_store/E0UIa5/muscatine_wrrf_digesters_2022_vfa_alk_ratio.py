# filename: muscatine_wrrf_digesters_2022_vfa_alk_ratio.py
import pandas as pd
import plotly.graph_objects as go

file_name = "muscatine_wrrf_digesters_2022.xlsx"
base_name = "muscatine_wrrf_digesters_2022_vfa_alk_ratio"

# Load workbook
xls = pd.ExcelFile(file_name)
if "daily_data" not in xls.sheet_names:
    raise ValueError("Sheet 'daily_data' not found in workbook.")

df = pd.read_excel(xls, sheet_name="daily_data")

required_cols = [
    "Year",
    "Month",
    "Day",
    "Dig1-VFA-Alkalinity_ratio",
    "Dig2-VFA-Alkalinity_ratio",
]
missing = [c for c in required_cols if c not in df.columns]
if missing:
    raise ValueError(f"Missing required columns: {missing}")

# Build proper datetime
df["Date"] = pd.to_datetime(
    dict(year=df["Year"], month=df["Month"], day=df["Day"]),
    errors="coerce"
)

plot_df = df.loc[:, ["Date", "Dig1-VFA-Alkalinity_ratio", "Dig2-VFA-Alkalinity_ratio"]].copy()
plot_df = plot_df.sort_values("Date").reset_index(drop=True)

# Save CSV with exactly requested columns
csv_df = plot_df.rename(columns={
    "Date": "Date [YYYY-MM-DD]",
    "Dig1-VFA-Alkalinity_ratio": "Digester 1 VFA/Alk ratio [-]",
    "Dig2-VFA-Alkalinity_ratio": "Digester 2 VFA/Alk ratio [-]",
})
csv_df["Date [YYYY-MM-DD]"] = csv_df["Date [YYYY-MM-DD]"].dt.strftime("%Y-%m-%d")
csv_file = f"{base_name}.csv"
csv_df.to_csv(csv_file, index=False)

# Create chart
fig = go.Figure()

fig.add_trace(go.Scatter(
    x=plot_df["Date"],
    y=plot_df["Dig1-VFA-Alkalinity_ratio"],
    mode="lines+markers",
    name="Digester 1 VFA/Alk ratio",
    line=dict(width=2),
    marker=dict(size=5)
))

fig.add_trace(go.Scatter(
    x=plot_df["Date"],
    y=plot_df["Dig2-VFA-Alkalinity_ratio"],
    mode="lines+markers",
    name="Digester 2 VFA/Alk ratio",
    line=dict(width=2),
    marker=dict(size=5)
))

# Thresholds as legend items using constant y Scatter traces
x0 = plot_df["Date"].min()
x1 = plot_df["Date"].max()

fig.add_trace(go.Scatter(
    x=[x0, x1],
    y=[0.30, 0.30],
    mode="lines",
    name="0.30 typical caution threshold",
    line=dict(color="firebrick", width=2, dash="dash"),
    hoverinfo="skip"
))

fig.add_trace(go.Scatter(
    x=[x0, x1],
    y=[0.40, 0.40],
    mode="lines",
    name="0.40 typical instability/upset threshold",
    line=dict(color="darkorange", width=2, dash="dash"),
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
    margin=dict(l=80, r=30, t=30, b=70),
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
    title_text="VFA/Alkalinity ratio [–]",
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
print("Created file: muscatine_wrrf_digesters_2022_vfa_alk_ratio.py")