# filename: muscatine_wrrf_digesters_2022_hsw_split_fraction.py
import pandas as pd
import numpy as np
import plotly.graph_objects as go

file_name = "muscatine_wrrf_digesters_2022.xlsx"
base_name = "muscatine_wrrf_digesters_2022_hsw_split_fraction"

# Load workbook
xls = pd.ExcelFile(file_name)
if "daily_data" not in xls.sheet_names:
    raise ValueError("Sheet 'daily_data' not found in workbook.")

df = pd.read_excel(xls, sheet_name="daily_data")

required_cols = ["Year", "Month", "Day", "V-HSW_m3d"]
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
    """
    Return (fraction_to_dig1, fraction_to_dig2) using the requested rules:
    - If both missing: 0.50 / 0.50
    - If only one present: other is remainder
    - If both present but do not sum to 100% within ±1%, normalize to sum to 1.0
    """
    p1 = row[hsw_d1_col] if has_hsw_d1 else np.nan
    p2 = row[hsw_d2_col] if has_hsw_d2 else np.nan

    # Both missing
    if pd.isna(p1) and pd.isna(p2):
        return 0.5, 0.5

    # Only one present
    if pd.isna(p1) and not pd.isna(p2):
        f2 = float(p2) / 100.0
        f1 = 1.0 - f2
        return f1, f2

    if not pd.isna(p1) and pd.isna(p2):
        f1 = float(p1) / 100.0
        f2 = 1.0 - f1
        return f1, f2

    # Both present
    f1 = float(p1) / 100.0
    f2 = float(p2) / 100.0
    s = f1 + f2
    if pd.isna(s) or s <= 0:
        return 0.5, 0.5

    # If they are close enough to 1.0, keep as-is; otherwise normalize
    if abs(s - 1.0) <= 0.01:
        return f1, f2

    return f1 / s, f2 / s

fractions = df.apply(compute_hsw_fractions, axis=1, result_type="expand")
df["HSW fraction to Digester 1 [-]"] = fractions[0]
df["HSW fraction to Digester 2 [-]"] = fractions[1]

plot_df = df.loc[
    :, ["Date", "HSW fraction to Digester 1 [-]", "HSW fraction to Digester 2 [-]"]
].copy()
plot_df = plot_df.sort_values("Date").reset_index(drop=True)

# Save CSV with exactly requested columns
csv_df = plot_df.rename(columns={"Date": "Date [YYYY-MM-DD]"})
csv_df["Date [YYYY-MM-DD]"] = csv_df["Date [YYYY-MM-DD]"].dt.strftime("%Y-%m-%d")
csv_df = csv_df[
    [
        "Date [YYYY-MM-DD]",
        "HSW fraction to Digester 1 [-]",
        "HSW fraction to Digester 2 [-]",
    ]
]
csv_file = f"{base_name}.csv"
csv_df.to_csv(csv_file, index=False)

# Create 100% stacked area chart
fig = go.Figure()

fig.add_trace(go.Scatter(
    x=plot_df["Date"],
    y=plot_df["HSW fraction to Digester 1 [-]"],
    mode="lines",
    name="HSW fraction to Digester 1",
    line=dict(width=2),
    stackgroup="one",
    groupnorm="fraction",
    hovertemplate="Date=%{x|%Y-%m-%d}<br>Digester 1=%{y:.3f}<extra></extra>"
))

fig.add_trace(go.Scatter(
    x=plot_df["Date"],
    y=plot_df["HSW fraction to Digester 2 [-]"],
    mode="lines",
    name="HSW fraction to Digester 2",
    line=dict(width=2),
    stackgroup="one",
    groupnorm="fraction",
    hovertemplate="Date=%{x|%Y-%m-%d}<br>Digester 2=%{y:.3f}<extra></extra>"
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
    title_text="HSW split fraction [-]",
    range=[0, 1],
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
print("Created file: muscatine_wrrf_digesters_2022_hsw_split_fraction.py")