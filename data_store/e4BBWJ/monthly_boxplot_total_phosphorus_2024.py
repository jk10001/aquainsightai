# filename: monthly_boxplot_total_phosphorus_2024.py
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
    If units contains 'mg' -> multiply by 1e3.
    If units contains 'ug' or 'µg' -> keep as-is.
    If units is missing or unknown, assume ug/L (prefer not to silently convert).
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
        use 0.5 * Detection Limit (after converting DL to ug/L based on Units)
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
    # If detection limit present but units differ, convert DL to ug/L using Units field
    dl_ugL = to_ug_per_l(dl_num, units) if not np.isnan(dl_num) else np.nan

    if censored:
        if not np.isnan(dl_ugL):
            return 0.5 * dl_ugL
        # fallback: try parse numeric from results like '<2' -> 2, then half
        if isinstance(results_raw, str) and results_raw.startswith("<"):
            tail = results_raw.lstrip("<").strip()
            # extract leading numeric token
            token = ""
            for ch in tail:
                if ch.isdigit() or ch in ".-+eE":
                    token += ch
                else:
                    break
            try:
                parsed = float(token)
                # assume same units as Units column -> convert to ug/L
                return 0.5 * to_ug_per_l(parsed, units)
            except Exception:
                return np.nan
        return np.nan
    else:
        # not censored: parse numeric value and convert to ug/L
        # Results may include trailing qualifiers/tokens; find first numeric token
        s = str(results_raw).replace(",", "")
        tokens = s.split()
        for tok in tokens:
            try:
                val = float(tok)
                return to_ug_per_l(val, units)
            except Exception:
                continue
        # final fallback
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
    df_tp["Collected date"] = pd.to_datetime(df_tp["Collected"], errors="coerce", infer_datetime_format=True)

    # Derive Month category as Jan–Dec ordered
    month_order = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
    # Create month name abbreviation
    df_tp["Month"] = df_tp["Collected date"].dt.month
    df_tp["Month"] = df_tp["Month"].apply(lambda m: month_order[int(m)-1] if not pd.isna(m) else np.nan)

    # Compute numeric Total phosphorus [ug/L]
    df_tp["Detection Limit numeric"] = df_tp["Detection Limit"].apply(safe_float)
    # Apply parsing function row-wise
    df_tp["Total phosphorus [ug/L]"] = df_tp.apply(parse_tp_value, axis=1)

    # Exclude rows where TP numeric cannot be computed or month missing or collected date missing
    df_plot = df_tp[ (~df_tp["Total phosphorus [ug/L]"].isna()) & (~df_tp["Collected date"].isna()) & (~df_tp["Month"].isna()) ].copy()

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

    # Build box traces per month using Plotly
    # We will create a single box trace with x=df_plot['Month'] and y values, using layout.boxmode default
    # But to ensure no outliers shown and whiskers = 1.5 IQR (Tukey), set boxpoints=False (no markers).
    # Plotly's default whisker length uses Tukey (1.5 IQR). So no extra parameter needed.
    # For explicit month order, we will aggregate month categories and create traces in order.

    traces = []
    for m in month_order:
        sub = df_plot[df_plot["Month"] == m]
        # If no data for month, create an empty trace to preserve month on x-axis (with no box)
        if sub.shape[0] == 0:
            # create an invisible empty trace to keep category
            trace = go.Box(
                x=[m],
                y=[],
                name=m,
                boxpoints=False,
                marker=dict(size=4),
                showlegend=False
            )
        else:
            trace = go.Box(
                x=sub["Month"].astype(str),
                y=sub["Total phosphorus [ug/L]"].astype(float),
                name=m,
                boxpoints=False,        # do not show sample points or outliers
                marker=dict(size=4),
                whiskerwidth=None,
                hovertemplate="Month: %{x}<br>Total phosphorus [ug/L]: %{y:.3f}<extra></extra>"
            )
        traces.append(trace)

    # PWQO reference line as separate scatter (for legend entry) spanning the months
    # Create an x sequence across month positions; using month_order as categories for plotting
    ref_line = go.Scatter(
        x=month_order,
        y=[30.0]*len(month_order),
        mode="lines",
        name="PWQO = 30 µg/L",
        line=dict(color="black", dash="dash"),
        hoverinfo="skip"
    )

    # Compose figure
    fig = go.Figure(data=traces + [ref_line])

    # Layout adjustments per requirements: plotly_white style, black axis lines and ticks, no title,
    # global font size 16, light gridlines, axes labels with units
    fig.update_layout(
        template="plotly_white",
        font=dict(size=16),
        xaxis=dict(title="Month", showgrid=True, gridcolor="lightgrey", linecolor="black", mirror=True),
        yaxis=dict(title="Total phosphorus [µg/L]", showgrid=True, gridcolor="lightgrey", linecolor="black", mirror=True),
        margin=dict(l=60, r=20, t=10, b=60),
        showlegend=True
    )

    # Ensure black axis ticks (Plotly default with plotly_white is fine, but enforce tick color)
    fig.update_xaxes(tickfont=dict(color="black"))
    fig.update_yaxes(tickfont=dict(color="black"))

    # Save CSV containing only the exact plotted dataset with required columns
    out_df = df_plot.loc[:, ["Collected date", "Month", "Total phosphorus [ug/L]", "Collection Site"]].copy()
    # Rename Collection Site to Collection site for requested column name casing
    out_df.rename(columns={"Collection Site": "Collection site"}, inplace=True)
    # Round TP values to reasonable precision for CSV
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