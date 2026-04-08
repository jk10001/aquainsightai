# filename: monthly_boxplot_total_phosphorus_2024_v2.py
import os
import sys
import pandas as pd
import numpy as np
import plotly.graph_objects as go

def safe_float(x):
    try:
        if pd.isna(x):
            return np.nan
        s = str(x).strip().replace(",", "")
        return float(s)
    except Exception:
        return np.nan

def to_ug_per_l(value, units):
    """
    Convert a numeric value to ug/L based on units string.
    - mg/L -> multiply by 1e3
    - ug/L or µg/L or ug -> leave as-is
    If units missing or unrecognized, assume same units as provided (no silent conversion),
    but for our dataset common units are 'mg/L' or 'ug/L'.
    """
    if pd.isna(value):
        return np.nan
    units_str = "" if pd.isna(units) else str(units).lower()
    try:
        val = float(value)
    except Exception:
        return np.nan
    if "mg" in units_str:
        return val * 1000.0
    # treat 'ug' or 'µg' or 'ug/l' as ug/L
    return val

def parse_tp_value(row):
    """
    Parse TP numeric value in ug/L according to rules:
    - If Results starts with '<' OR Value Qualifier == '<' OR Result Call == 'BDL':
        use 0.5 * Detection Limit (converted to ug/L)
    - Else parse numeric from Results and convert units to ug/L
    Returns float (ug/L) or np.nan
    """
    results_raw = "" if pd.isna(row.get("Results", "")) else str(row.get("Results", "")).strip()
    result_call = "" if pd.isna(row.get("Result Call", "")) else str(row.get("Result Call", "")).strip()
    value_qual = "" if pd.isna(row.get("Value Qualifier", "")) else str(row.get("Value Qualifier", "")).strip()
    units = row.get("Units", "")
    dl_raw = row.get("Detection Limit", np.nan)

    censored = False
    if isinstance(results_raw, str) and results_raw.startswith("<"):
        censored = True
    if value_qual == "<":
        censored = True
    if str(result_call).upper() == "BDL":
        censored = True

    # Try numeric detection limit
    dl_num = safe_float(dl_raw)
    dl_ugL = to_ug_per_l(dl_num, units) if not np.isnan(dl_num) else np.nan

    if censored:
        if not np.isnan(dl_ugL):
            return 0.5 * dl_ugL
        # fallback: try parse numeric from results like '<2' -> 2, then half and convert units
        if isinstance(results_raw, str) and results_raw.startswith("<"):
            tail = results_raw.lstrip("<").strip()
            token = ""
            for ch in tail:
                if ch.isdigit() or ch in ".-+eE":
                    token += ch
                else:
                    break
            try:
                parsed = float(token)
                return 0.5 * to_ug_per_l(parsed, units)
            except Exception:
                return np.nan
        return np.nan
    else:
        # not censored: parse numeric value and convert to ug/L
        s = str(results_raw).replace(",", "")
        tokens = s.split()
        for tok in tokens:
            try:
                val = float(tok)
                return to_ug_per_l(val, units)
            except Exception:
                continue
        return np.nan

def main():
    py_filename = os.path.basename(__file__)
    excel_file = "Ontario_PWQMN_2024.xlsx"
    sheet_name = "Data"

    base = "monthly_boxplot_total_phosphorus_2024"
    out_py = py_filename
    out_csv = f"{base}.csv"
    out_html = f"{base}.html"
    out_png = f"{base}.png"

    # Read Data sheet
    try:
        df = pd.read_excel(excel_file, sheet_name=sheet_name, engine="openpyxl", dtype=object)
    except Exception as e:
        print(f"ERROR reading '{sheet_name}' from '{excel_file}': {e}", file=sys.stderr)
        raise

    # Validate expected columns
    expected_cols = {"Collection Site", "Analyte", "Collected", "Results", "Units", "Result Call", "Detection Limit", "Value Qualifier"}
    missing = expected_cols - set(df.columns)
    if missing:
        raise KeyError(f"Missing expected columns in Data sheet: {missing}")

    # Filter to total phosphorus analyte
    df_tp = df[df["Analyte"].astype(str).str.strip() == "Phosphorus; total"].copy()

    # Parse Collected to datetime (handle mixed formats)
    # Use errors='coerce' to drop unparsable rows later
    # Some entries may be like 'MM/DD/YYYY' or 'YYYY-MM-DD 00:00:00' etc.
    df_tp["Collected date"] = pd.to_datetime(df_tp["Collected"], errors="coerce", infer_datetime_format=True)

    # Derive Month category as Jan–Dec ordered
    month_order = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
    df_tp["Month_num"] = df_tp["Collected date"].dt.month
    df_tp["Month"] = df_tp["Month_num"].apply(lambda m: month_order[int(m)-1] if not pd.isna(m) else np.nan)

    # Compute numeric Total phosphorus [ug/L]
    df_tp["Detection Limit numeric"] = df_tp["Detection Limit"].apply(safe_float)
    df_tp["Total phosphorus [ug/L]"] = df_tp.apply(parse_tp_value, axis=1)

    # Exclude rows where TP numeric cannot be computed or month missing or collected date missing
    df_plot = df_tp[
        (~df_tp["Total phosphorus [ug/L]"].isna()) &
        (~df_tp["Collected date"].isna()) &
        (~df_tp["Month"].isna())
    ].copy()

    # If nothing to plot, output empty CSV and exit gracefully
    if df_plot.shape[0] == 0:
        df_empty = pd.DataFrame(columns=["Collected date", "Month", "Total phosphorus [ug/L]", "Collection site"])
        df_empty.to_csv(out_csv, index=False)
        print(out_py)
        print(out_csv)
        print(out_html)
        print(out_png)
        print("No valid Total Phosphorus rows to plot after filtering/processing.")
        return

    # Prepare Month as categorical ordered
    df_plot["Month"] = pd.Categorical(df_plot["Month"], categories=month_order, ordered=True)

    # Create a single box trace with x=Month and y=TP, no outlier markers; whiskers are Tukey (1.5 IQR)
    box_trace = go.Box(
        x=df_plot["Month"].astype(str),
        y=df_plot["Total phosphorus [ug/L]"].astype(float),
        name="",                 # hide name so it does not clutter legend
        boxpoints=False,         # do not show sample points or outliers
        marker=dict(size=4),
        hovertemplate="Month: %{x}<br>Total phosphorus [ug/L]: %{y:.3f}<extra></extra>"
    )

    # PWQO reference line (single legend entry)
    # Create a line across x categories by using the month_order as x and constant 30 as y
    ref_line = go.Scatter(
        x=month_order,
        y=[30.0]*len(month_order),
        mode="lines",
        name="PWQO = 30 µg/L",
        line=dict(color="black", dash="dash"),
        hoverinfo="skip",
        showlegend=True
    )

    fig = go.Figure(data=[box_trace, ref_line])

    # Layout adjustments per requirements: plotly_white style, black axis lines and ticks, no title,
    # global font size 16, light gridlines, axes labels with units
    fig.update_layout(
        template="plotly_white",
        font=dict(size=16),
        xaxis=dict(title="Month", showgrid=True, gridcolor="lightgrey", linecolor="black", tickfont=dict(color="black")),
        yaxis=dict(title="Total phosphorus [µg/L]", showgrid=True, gridcolor="lightgrey", linecolor="black", tickfont=dict(color="black")),
        margin=dict(l=60, r=20, t=10, b=60),
        showlegend=True
    )

    # Save CSV containing only the exact plotted dataset with required columns
    out_df = df_plot.loc[:, ["Collected date", "Month", "Total phosphorus [ug/L]", "Collection Site"]].copy()
    out_df.rename(columns={"Collection Site": "Collection site"}, inplace=True)
    # Round TP values for export readability
    out_df["Total phosphorus [ug/L]"] = out_df["Total phosphorus [ug/L]"].round(6)
    out_df.to_csv(out_csv, index=False, encoding="utf-8")

    # Save HTML (interactive) - include_plotlyjs='cdn', full_html=True
    try:
        fig.write_html(out_html, include_plotlyjs="cdn", full_html=True)
    except Exception as e:
        print(f"ERROR saving HTML '{out_html}': {e}", file=sys.stderr)
        raise

    # Save PNG 1000x600 via kaleido
    try:
        fig.write_image(out_png, width=1000, height=600, scale=1)
    except Exception as e:
        print(f"ERROR saving PNG '{out_png}': {e}", file=sys.stderr)
        print("Ensure the 'kaleido' package is installed for static image export.", file=sys.stderr)
        raise

    # Print created filenames and python filename to terminal as required
    print(out_py)
    print(out_csv)
    print(out_html)
    print(out_png)

if __name__ == "__main__":
    main()