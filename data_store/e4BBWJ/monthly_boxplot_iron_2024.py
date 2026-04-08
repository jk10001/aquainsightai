# filename: monthly_boxplot_iron_2024.py
import os
import sys
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from datetime import datetime

def safe_float(x):
    try:
        if pd.isna(x):
            return np.nan
        s = str(x).strip().replace(",", "")
        return float(s)
    except Exception:
        return np.nan

def parse_iron_value(result_raw, result_call, value_qualifier, detection_limit):
    """
    Parse Results into numeric Iron [ug/L] per rules:
    - If Results begins with '<' OR Value Qualifier == '<' OR Result Call == 'BDL':
        use 0.5 * Detection Limit (DL expected in ug/L)
    - Else parse numeric value from Results
    Returns float or np.nan
    """
    result = "" if pd.isna(result_raw) else str(result_raw).strip()
    rc = "" if pd.isna(result_call) else str(result_call).strip()
    vq = "" if pd.isna(value_qualifier) else str(value_qualifier).strip()
    dl = detection_limit

    censored = False
    if isinstance(result, str) and result.startswith("<"):
        censored = True
    if vq == "<":
        censored = True
    if str(rc).upper() == "BDL":
        censored = True

    dl_num = safe_float(dl)

    if censored:
        if not np.isnan(dl_num):
            return 0.5 * dl_num
        # fallback: attempt to parse numeric after '<'
        if isinstance(result, str) and result.startswith("<"):
            numpart = result.lstrip("<").strip()
            token = ""
            for ch in numpart:
                if ch.isdigit() or ch in ".-+eE":
                    token += ch
                else:
                    break
            try:
                val = float(token)
                return 0.5 * val
            except Exception:
                return np.nan
        return np.nan
    else:
        # not censored: parse numeric token from result
        s = str(result).replace(",", "")
        tokens = s.split()
        for tok in tokens:
            try:
                return float(tok)
            except Exception:
                continue
        return safe_float(s)

def compute_tukey_whiskers(values):
    """
    Given a 1D array-like of numeric values (no NaNs), compute Tukey stats:
    q1, median, q3, iqr, lower_fence, upper_fence, lower_whisker, upper_whisker
    """
    arr = np.array(values, dtype=float)
    if arr.size == 0:
        return None
    q1 = np.percentile(arr, 25)
    median = np.percentile(arr, 50)
    q3 = np.percentile(arr, 75)
    iqr = q3 - q1
    lower_fence = q1 - 1.5 * iqr
    upper_fence = q3 + 1.5 * iqr
    within = arr[(arr >= lower_fence) & (arr <= upper_fence)]
    if within.size > 0:
        lower_whisker = float(within.min())
        upper_whisker = float(within.max())
    else:
        lower_whisker = float(q1)
        upper_whisker = float(q3)
    return {
        "q1": float(q1),
        "median": float(median),
        "q3": float(q3),
        "iqr": float(iqr),
        "lower_fence": float(lower_fence),
        "upper_fence": float(upper_fence),
        "lower_whisker": lower_whisker,
        "upper_whisker": upper_whisker
    }

def make_month_name(dt):
    if pd.isna(dt):
        return None
    month_order = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
    try:
        m = int(dt.month)
        return month_order[m-1]
    except Exception:
        return None

def main():
    py_filename = os.path.basename(__file__)
    excel_file = "Ontario_PWQMN_2024.xlsx"
    sheet_name = "Data"
    base = "monthly_boxplot_iron_2024"
    out_py = py_filename
    out_csv = f"{base}.csv"
    out_html = f"{base}.html"
    out_png = f"{base}.png"

    # Read Excel Data sheet
    try:
        df = pd.read_excel(excel_file, sheet_name=sheet_name, engine="openpyxl", dtype=object)
    except Exception as e:
        print(f"ERROR reading '{sheet_name}' from '{excel_file}': {e}", file=sys.stderr)
        raise

    required_cols = {"Collection Site", "Analyte", "Collected", "Results", "Result Call", "Detection Limit", "Value Qualifier", "Units"}
    missing = required_cols - set(df.columns)
    if missing:
        raise KeyError(f"Missing expected columns in Data sheet: {missing}")

    # Filter to Iron analyte
    df_iron = df[df["Analyte"].astype(str).str.strip() == "Iron"].copy()

    # Parse Collected to datetime (mixed formats)
    df_iron["Collected date"] = pd.to_datetime(df_iron["Collected"], errors="coerce", infer_datetime_format=True)

    # Compute numeric Iron [ug/L]
    # Detection limit expected in same units as Results (dataset shows ug/L for Iron); but we'll assume DL is in ug/L as per description.
    df_iron["Detection Limit numeric"] = df_iron["Detection Limit"].apply(safe_float)
    df_iron["Iron [ug/L]"] = df_iron.apply(
        lambda r: parse_iron_value(
            r.get("Results", ""),
            r.get("Result Call", ""),
            r.get("Value Qualifier", ""),
            r.get("Detection Limit numeric", np.nan)
        ), axis=1
    )

    # Create Month name
    df_iron["Month"] = df_iron["Collected date"].apply(make_month_name)

    # Keep only rows with valid date, month, and numeric Iron
    df_plot = df_iron[
        (~df_iron["Collected date"].isna()) &
        (~df_iron["Month"].isna()) &
        (~df_iron["Iron [ug/L]"].isna())
    ].copy()

    # If no data to plot, create empty CSV/HTML/PNG placeholders and exit gracefully (but still print names)
    if df_plot.shape[0] == 0:
        # write empty CSV with requested headers
        df_empty = pd.DataFrame(columns=["Collected date [YYYY-MM-DD]", "Month", "Iron [ug/L]", "Collection site"])
        df_empty.to_csv(out_csv, index=False, encoding="utf-8")
        # minimal HTML
        with open(out_html, "w", encoding="utf-8") as f:
            f.write("<html><body><p>No Iron data available for plotting.</p></body></html>")
        # blank PNG
        try:
            from PIL import Image
            img = Image.new("RGB", (1000, 600), color=(255,255,255))
            img.save(out_png, format="PNG")
        except Exception:
            # if PIL not available, skip PNG creation
            pass
        print(out_py)
        print(out_csv)
        print(out_html)
        print(out_png)
        return

    # Format CSV output: Collected date [YYYY-MM-DD], Month, Iron [ug/L], Collection site
    df_csv_out = pd.DataFrame({
        "Collected date [YYYY-MM-DD]": df_plot["Collected date"].dt.strftime("%Y-%m-%d"),
        "Month": df_plot["Month"],
        "Iron [ug/L]": df_plot["Iron [ug/L]"].astype(float).round(6),
        "Collection site": df_plot["Collection Site"].astype(str)
    })
    df_csv_out.to_csv(out_csv, index=False, encoding="utf-8")

    # Prepare month ordering and per-month values
    month_order = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
    df_plot["Month"] = pd.Categorical(df_plot["Month"], categories=month_order, ordered=True)

    month_values = {m: df_plot.loc[df_plot["Month"] == m, "Iron [ug/L]"].dropna().astype(float).values for m in month_order}

    # Compute Tukey summary stats per month
    month_stats = {}
    for m in month_order:
        vals = month_values.get(m, np.array([], dtype=float))
        stats = compute_tukey_whiskers(vals)
        month_stats[m] = {
            "n": int(len(vals)),
            "values": vals,
            "stats": stats
        }

    # Build Plotly traces: one box per month using summary stats (box-from-stats)
    box_traces = []
    for m in month_order:
        s = month_stats[m]["stats"]
        n = month_stats[m]["n"]
        if s is None or n == 0:
            # add an invisible empty box to keep month positions consistent
            trace_empty = go.Box(
                x=[m],
                y=[],
                name=m,
                boxpoints=False,
                showlegend=False,
                marker=dict(color="lightblue"),
                hoverinfo="skip",
                visible="legendonly"
            )
            box_traces.append(trace_empty)
            continue
        # Create a box-from-stats trace specifying quartiles and whiskers
        trace = go.Box(
            x=[m],
            y=[],  # leave raw y empty since we're providing stats
            name=m,
            boxpoints=False,  # no outlier/points
            showlegend=False,
            marker=dict(color="lightblue"),
            q1=[s["q1"]],
            median=[s["median"]],
            q3=[s["q3"]],
            lowerfence=[s["lower_whisker"]],
            upperfence=[s["upper_whisker"]],
            quartilemethod="linear"
        )
        box_traces.append(trace)

    # Reference line for PWQO = 300 ug/L as dashed and included in legend
    ref_line = go.Scatter(
        x=month_order,
        y=[300.0]*len(month_order),
        mode="lines",
        name="PWQO = 300 µg/L",
        line=dict(color="black", dash="dash"),
        hoverinfo="skip",
        showlegend=True
    )

    # Construct figure
    fig = go.Figure(data=box_traces + [ref_line])

    # Layout: plotly_white, black axis lines and ticks, font size 16, no title, no rangeslider
    fig.update_layout(
        template="plotly_white",
        font=dict(size=16),
        xaxis=dict(title="Month", showgrid=True, gridcolor="lightgrey", linecolor="black", tickfont=dict(color="black")),
        yaxis=dict(title="Iron [µg/L]", showgrid=True, gridcolor="lightgrey", linecolor="black", tickfont=dict(color="black")),
        margin=dict(l=60, r=20, t=10, b=60),
        showlegend=True
    )

    # Save HTML (responsive, include_plotlyjs='cdn', no title)
    try:
        fig.write_html(out_html, include_plotlyjs="cdn", full_html=True)
    except Exception as e:
        print(f"ERROR saving HTML '{out_html}': {e}", file=sys.stderr)
        raise

    # Save PNG via kaleido at 1000x600
    try:
        fig.write_image(out_png, width=1000, height=600, scale=1)
    except Exception as e:
        print(f"ERROR saving PNG '{out_png}': {e}", file=sys.stderr)
        print("Make sure the 'kaleido' package is installed.", file=sys.stderr)
        raise

    # Print created filenames and python filename to terminal
    print(out_py)
    print(out_csv)
    print(out_html)
    print(out_png)

if __name__ == "__main__":
    main()