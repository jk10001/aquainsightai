# filename: muscatine_wrrf_digesters_2022_hrt_time_series.py
import pandas as pd
import plotly.graph_objects as go

file_name = "muscatine_wrrf_digesters_2022.xlsx"
base_name = "muscatine_wrrf_digesters_2022_hrt_time_series"

# Load workbook
xls = pd.ExcelFile(file_name)
if "daily_data" not in xls.sheet_names:
    raise ValueError("Sheet 'daily_data' not found in workbook.")

df = pd.read_excel(xls, sheet_name="daily_data")

required_cols = ["Year", "Month", "Day", "V-TWAS_m3d", "V-PS_m3d", "V-HSW_m3d"]
missing = [c for c in required_cols if c not in df.columns]
if missing:
    raise ValueError(f"Missing required columns: {missing}")

hsw_d1_col = "HSW-Dig1-portion_pct"
hsw_d2_col = "HSW2-Dig2-portion_pct"
has_hsw_d1 = hsw_d1_col in df.columns
has_hsw_d2 = hsw_d2_col in df.columns

# Build proper datetime
df["Date"] = pd.to_datetime(
    dict(year=df["Year"], month=df["Month"], day=df["Day"]),
    errors="coerce"
)

# Helper to compute HSW split fractions
def compute_hsw_fractions(row):
    hsw = row["V-HSW_m3d"]
    p1 = row[hsw_d1_col] if has_hsw_d1 else float("nan")
    p2 = row[hsw_d2_col] if has_hsw_d2 else float("nan")

    # Missing portion(s): assume 50/50
    if pd.isna(p1) and pd.isna(p2):
        return 0.5, 0.5
    if pd.isna(p1) and not pd.isna(p2):
        f2 = float(p2) / 100.0
        f1 = 1.0 - f2
        return f1, f2
    if not pd.isna(p1) and pd.isna(p2):
        f1 = float(p1) / 100.0
        f2 = 1.0 - f1
        return f1, f2

    # Both present: normalize if needed
    f1 = float(p1) / 100.0
    f2 = float(p2) / 100.0
    s = f1 + f2
    if pd.isna(s) or s <= 0:
        return 0.5, 0.5
    if abs(s - 1.0) <= 0.01:
        return f1, f2
    return f1 / s, f2 / s

# Compute HRT values
d1_hrt = []
d2_hrt = []

for _, row in df.iterrows():
    twas = row["V-TWAS_m3d"]
    ps = row["V-PS_m3d"]
    hsw = row["V-HSW_m3d"]

    if pd.isna(twas):
        twas = 0.0
    if pd.isna(ps):
        ps = 0.0
    if pd.isna(hsw):
        hsw = 0.0

    f1, f2 = compute_hsw_fractions(row)

    twas_to_d1 = 0.5 * twas
    twas_to_d2 = 0.5 * twas
    ps_to_d1 = 0.5 * ps
    ps_to_d2 = 0.5 * ps

    hsw_to_d1 = hsw * f1
    hsw_to_d2 = hsw * f2

    q_d1 = twas_to_d1 + ps_to_d1 + hsw_to_d1
    q_d2 = twas_to_d2 + ps_to_d2 + hsw_to_d2

    hrt_d1 = 1625.0 / q_d1 if q_d1 > 0 else float("nan")
    hrt_d2 = 1625.0 / q_d2 if q_d2 > 0 else float("nan")

    d1_hrt.append(hrt_d1)
    d2_hrt.append(hrt_d2)

df["Digester 1 calculated HRT [days]"] = d1_hrt
df["Digester 2 calculated HRT [days]"] = d2_hrt

plot_df = df.loc[:, ["Date", "Digester 1 calculated HRT [days]", "Digester 2 calculated HRT [days]"]].copy()
plot_df = plot_df.sort_values("Date").reset_index(drop=True)

# Save CSV with exactly requested columns
csv_df = plot_df.rename(columns={
    "Date": "Date [YYYY-MM-DD]",
})
csv_df["Date [YYYY-MM-DD]"] = csv_df["Date [YYYY-MM-DD]"].dt.strftime("%Y-%m-%d")
csv_file = f"{base_name}.csv"
csv_df.to_csv(csv_file, index=False)

# Create chart
fig = go.Figure()

fig.add_trace(go.Scatter(
    x=plot_df["Date"],
    y=plot_df["Digester 1 calculated HRT [days]"],
    mode="lines+markers",
    name="Digester 1 calculated HRT",
    line=dict(width=2),
    marker=dict(size=5)
))

fig.add_trace(go.Scatter(
    x=plot_df["Date"],
    y=plot_df["Digester 2 calculated HRT [days]"],
    mode="lines+markers",
    name="Digester 2 calculated HRT",
    line=dict(width=2),
    marker=dict(size=5)
))

# Reference lines as legend items
x0 = plot_df["Date"].min()
x1 = plot_df["Date"].max()

fig.add_trace(go.Scatter(
    x=[x0, x1],
    y=[15, 15],
    mode="lines",
    name="15 d typical minimum",
    line=dict(color="firebrick", width=2, dash="dash"),
    hoverinfo="skip"
))

fig.add_trace(go.Scatter(
    x=[x0, x1],
    y=[25, 25],
    mode="lines",
    name="25 d typical target",
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
    title_text="Calculated HRT [days]",
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
print("Created file: muscatine_wrrf_digesters_2022_hrt_time_series.py")