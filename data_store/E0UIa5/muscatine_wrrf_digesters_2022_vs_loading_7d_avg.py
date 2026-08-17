# filename: muscatine_wrrf_digesters_2022_vs_loading_7d_avg.py
import pandas as pd
import numpy as np
import plotly.graph_objects as go

file_name = "muscatine_wrrf_digesters_2022.xlsx"
base_name = "muscatine_wrrf_digesters_2022_vs_loading_7d_avg"

# Load workbook
xls = pd.ExcelFile(file_name)
if "daily_data" not in xls.sheet_names:
    raise ValueError("Sheet 'daily_data' not found in workbook.")

df = pd.read_excel(xls, sheet_name="daily_data")

required_cols = [
    "Year", "Month", "Day",
    "V-TWAS_m3d", "V-PS_m3d", "V-HSW_m3d",
    "TWAS-VS_percent", "PS-VS_percent", "HSW-VS_percent"
]
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

def compute_hsw_fractions(row):
    p1 = row[hsw_d1_col] if has_hsw_d1 else np.nan
    p2 = row[hsw_d2_col] if has_hsw_d2 else np.nan

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

    f1 = float(p1) / 100.0
    f2 = float(p2) / 100.0
    s = f1 + f2
    if pd.isna(s) or s <= 0:
        return 0.5, 0.5
    if abs(s - 1.0) <= 0.01:
        return f1, f2
    return f1 / s, f2 / s

def safe_num(x):
    return 0.0 if pd.isna(x) else float(x)

# Calculate daily VS loading rate to each digester
vs_d1_list = []
vs_d2_list = []

for _, row in df.iterrows():
    twas_m3d = safe_num(row["V-TWAS_m3d"])
    ps_m3d = safe_num(row["V-PS_m3d"])
    hsw_m3d = safe_num(row["V-HSW_m3d"])

    twas_vs = safe_num(row["TWAS-VS_percent"]) / 100.0
    ps_vs = safe_num(row["PS-VS_percent"]) / 100.0
    hsw_vs = safe_num(row["HSW-VS_percent"]) / 100.0

    f1, f2 = compute_hsw_fractions(row)

    # Split flows
    twas_to_d1 = 0.5 * twas_m3d
    twas_to_d2 = 0.5 * twas_m3d
    ps_to_d1 = 0.5 * ps_m3d
    ps_to_d2 = 0.5 * ps_m3d
    hsw_to_d1 = hsw_m3d * f1
    hsw_to_d2 = hsw_m3d * f2

    # kg VS/d = m3/d * 1000 kg/m3 * VS fraction
    vs_twas_d1 = twas_to_d1 * 1000.0 * twas_vs
    vs_twas_d2 = twas_to_d2 * 1000.0 * twas_vs
    vs_ps_d1 = ps_to_d1 * 1000.0 * ps_vs
    vs_ps_d2 = ps_to_d2 * 1000.0 * ps_vs
    vs_hsw_d1 = hsw_to_d1 * 1000.0 * hsw_vs
    vs_hsw_d2 = hsw_to_d2 * 1000.0 * hsw_vs

    vs_d1 = vs_twas_d1 + vs_ps_d1 + vs_hsw_d1
    vs_d2 = vs_twas_d2 + vs_ps_d2 + vs_hsw_d2

    olr_d1 = vs_d1 / 1625.0
    olr_d2 = vs_d2 / 1625.0

    vs_d1_list.append(olr_d1)
    vs_d2_list.append(olr_d2)

df["Digester 1 daily VS loading rate [kgVS/m3-d]"] = vs_d1_list
df["Digester 2 daily VS loading rate [kgVS/m3-d]"] = vs_d2_list

plot_df = df.loc[
    :, ["Date", "Digester 1 daily VS loading rate [kgVS/m3-d]", "Digester 2 daily VS loading rate [kgVS/m3-d]"]
].copy()
plot_df = plot_df.sort_values("Date").reset_index(drop=True)

# 7-day trailing rolling average (window=7, min_periods=4)
plot_df["Digester 1 7d avg VS loading rate [kgVS/m3-d]"] = (
    plot_df["Digester 1 daily VS loading rate [kgVS/m3-d]"]
    .rolling(window=7, min_periods=4)
    .mean()
)
plot_df["Digester 2 7d avg VS loading rate [kgVS/m3-d]"] = (
    plot_df["Digester 2 daily VS loading rate [kgVS/m3-d]"]
    .rolling(window=7, min_periods=4)
    .mean()
)

# Save CSV with exactly requested columns
csv_df = plot_df.rename(columns={"Date": "Date [YYYY-MM-DD]"})
csv_df["Date [YYYY-MM-DD]"] = csv_df["Date [YYYY-MM-DD]"].dt.strftime("%Y-%m-%d")
csv_df = csv_df[
    [
        "Date [YYYY-MM-DD]",
        "Digester 1 7d avg VS loading rate [kgVS/m3-d]",
        "Digester 2 7d avg VS loading rate [kgVS/m3-d]",
    ]
]
csv_file = f"{base_name}.csv"
csv_df.to_csv(csv_file, index=False)

# Create chart
fig = go.Figure()

fig.add_trace(go.Scatter(
    x=plot_df["Date"],
    y=plot_df["Digester 1 7d avg VS loading rate [kgVS/m3-d]"],
    mode="lines+markers",
    name="Digester 1 7d avg VS loading rate",
    line=dict(width=2),
    marker=dict(size=5)
))

fig.add_trace(go.Scatter(
    x=plot_df["Date"],
    y=plot_df["Digester 2 7d avg VS loading rate [kgVS/m3-d]"],
    mode="lines+markers",
    name="Digester 2 7d avg VS loading rate",
    line=dict(width=2),
    marker=dict(size=5)
))

x0 = plot_df["Date"].min()
x1 = plot_df["Date"].max()

fig.add_trace(go.Scatter(
    x=[x0, x1],
    y=[1.6, 1.6],
    mode="lines",
    name="1.6 typical lower OLR",
    line=dict(color="firebrick", width=2, dash="dash"),
    hoverinfo="skip"
))

fig.add_trace(go.Scatter(
    x=[x0, x1],
    y=[3.2, 3.2],
    mode="lines",
    name="3.2 typical upper OLR",
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
    margin=dict(l=90, r=30, t=30, b=70),
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
    title_text="7-d avg VS loading rate [kg VS/m³-d]",
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
print("Created file: muscatine_wrrf_digesters_2022_vs_loading_7d_avg.py")