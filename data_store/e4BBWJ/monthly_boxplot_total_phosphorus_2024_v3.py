# filename: monthly_boxplot_total_phosphorus_2024_v3.py
import os
import sys
import math
import numpy as np
import pandas as pd

# Visualization imports
import matplotlib.pyplot as plt
import seaborn as sns
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
    Convert numeric value to ug/L based on units string.
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
    return val

def parse_tp_value(row):
    """
    Parse TP numeric value in ug/L according to original rules:
    - If Results starts with '<' OR Value Qualifier == '<' OR Result Call == 'BDL':
        use 0.5 * Detection Limit (converted to ug/L)
    - Else parse numeric from Results and convert units to ug/L
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

    dl_num = safe_float(dl_raw)
    dl_ugL = to_ug_per_l(dl_num, units) if not np.isnan(dl_num) else np.nan

    if censored:
        if not np.isnan(dl_ugL):
            return 0.5 * dl_ugL
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
        s = str(results_raw).replace(",", "")
        tokens = s.split()
        for tok in tokens:
            try:
                val = float(tok)
                return to_ug_per_l(val, units)
            except Exception:
                continue
        return np.nan

def compute_tukey_whiskers(values):
    """
    Given a 1D array-like of numeric values (no NaNs), compute:
     - q1, median, q3
     - iqr
     - lower_fence = q1 - 1.5*iqr
     - upper_fence = q3 + 1.5*iqr
     - lower_whisker = min(values >= lower_fence) else q1
     - upper_whisker = max(values <= upper_fence) else q3
    Returns dict with q1, median, q3, iqr, lower_whisker, upper_whisker, lower_fence, upper_fence
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
    # Values within fences
    within_lower = arr[arr >= lower_fence]
    within_upper = arr[arr <= upper_fence]
    # For Tukey whisker we take the most extreme observed points within the fences
    if within_lower.size > 0:
        lower_whisker = float(within_lower.min())
    else:
        lower_whisker = float(q1)
    if within_upper.size > 0:
        upper_whisker = float(within_upper.max())
    else:
        upper_whisker = float(q3)
    return {
        "q1": float(q1),
        "median": float(median),
        "q3": float(q3),
        "iqr": float(iqr),
        "lower_fence": float(lower_fence),
        "upper_fence": float(upper_fence),
        "lower_whisker": float(lower_whisker),
        "upper_whisker": float(upper_whisker)
    }

def main():
    py_filename = os.path.basename(__file__)
    excel_file = "Ontario_PWQMN_2024.xlsx"
    sheet_name = "Data"

    base = "monthly_boxplot_total_phosphorus_2024_v3"
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

    expected_cols = {"Collection Site", "Analyte", "Collected", "Results", "Units", "Result Call", "Detection Limit", "Value Qualifier"}
    missing = expected_cols - set(df.columns)
    if missing:
        raise KeyError(f"Missing expected columns in Data sheet: {missing}")

    # Filter to total phosphorus analyte
    df_tp = df[df["Analyte"].astype(str).str.strip() == "Phosphorus; total"].copy()

    # Parse Collected to datetime
    df_tp["Collected date"] = pd.to_datetime(df_tp["Collected"], errors="coerce", infer_datetime_format=True)

    # Month ordering
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

    if df_plot.shape[0] == 0:
        df_empty = pd.DataFrame(columns=["Collected date", "Month", "Total phosphorus [ug/L]", "Collection Site"])
        df_empty.to_csv(out_csv, index=False)
        print(out_py)
        print(out_csv)
        print(out_html)
        print(out_png)
        print("No valid Total Phosphorus rows to plot after filtering/processing.")
        return

    # Prepare Month categorical ordered
    df_plot["Month"] = pd.Categorical(df_plot["Month"], categories=month_order, ordered=True)

    # Save plotted data CSV with exact fields used
    out_df = df_plot.loc[:, ["Collected date", "Month", "Total phosphorus [ug/L]", "Collection Site"]].copy()
    out_df.rename(columns={"Collection Site": "Collection site"}, inplace=True)
    out_df["Total phosphorus [ug/L]"] = out_df["Total phosphorus [ug/L]"].round(6)
    out_df.to_csv(out_csv, index=False, encoding="utf-8")

    # Compute per-month summary stats using Tukey rule
    month_stats = {}
    for m in month_order:
        vals = df_plot.loc[df_plot["Month"] == m, "Total phosphorus [ug/L]"].dropna().astype(float).values
        stats = compute_tukey_whiskers(vals)
        month_stats[m] = {
            "n": int(len(vals)),
            "values": vals,
            "stats": stats
        }

    # MATPLOTLIB / SEABORN PNG creation with whis=1.5, showfliers=False to ensure Tukey whiskers
    # Prepare a list of arrays ordered by month_order (empty arrays for months with no data)
    data_for_plot = [month_stats[m]["values"] for m in month_order]

    # Create the figure
    plt.rcParams.update({'font.size': 16})
    fig, ax = plt.subplots(figsize=(1000/100, 600/100))  # dpi scaling: figsize in inches, with default dpi=100 gives desired px
    # Use seaborn boxplot wrapper to ensure consistent styling
    sns.set_style("whitegrid")
    # For seaborn, provide data as list; set whis=1.5 showfliers=False
    # Create positions for months that have data; seaborn handles empty lists by leaving gaps — we'll pass all months and rely on matplotlib
    # However seaborn.boxplot accepts a list of arrays via pd.DataFrame or list; using list of arrays with labels
    box = ax.boxplot(
        data_for_plot,
        positions=list(range(1, len(month_order)+1)),
        widths=0.6,
        patch_artist=True,
        showfliers=False,
        whis=1.5,
        medianprops=dict(color="black"),
        boxprops=dict(facecolor="lightblue", edgecolor="black"),
        whiskerprops=dict(color="black"),
        capprops=dict(color="black")
    )

    # Set x-axis ticks/labels
    ax.set_xticks(list(range(1, len(month_order)+1)))
    ax.set_xticklabels(month_order)
    ax.set_xlabel("Month")
    ax.set_ylabel("Total phosphorus [µg/L]")
    # PWQO reference line at 30 µg/L dashed black
    ax.axhline(30.0, color="black", linestyle="--", linewidth=1.5)

    # Tight layout and save PNG at 1000x600
    plt.tight_layout()
    # Save PNG with specified size and DPI to reach 1000x600 px
    fig.set_size_inches(1000/100, 600/100)  # inches at dpi=100 -> 1000x600
    try:
        plt.savefig(out_png, dpi=100, bbox_inches="tight")
    except Exception as e:
        print(f"ERROR saving PNG '{out_png}': {e}", file=sys.stderr)
        raise
    plt.close(fig)

    # Build Plotly figure from computed month_stats so interactive HTML shows exact Tukey whiskers
    # For each month, if there are data points, create a box trace using summary stats
    plotly_traces = []
    for m_idx, m in enumerate(month_order):
        s = month_stats[m]["stats"]
        n = month_stats[m]["n"]
        # If no data, produce an empty invisible trace to keep axis categories consistent
        if s is None or n == 0:
            # create an empty invisible trace for the month so categories align
            trace_empty = go.Box(
                x=[m],
                y=[],
                name=m,
                boxpoints=False,
                marker=dict(size=4),
                showlegend=False,
                hoverinfo="skip",
                visible="legendonly"
            )
            plotly_traces.append(trace_empty)
            continue
        # Use the summary-stat fields supported by Plotly to draw box from provided quartiles and whiskers:
        # q1, median, q3, lowerfence, upperfence correspond to the Tukey-derived whiskers
        # Provide a single-box trace for this month using the stats; specify y as empty and supply summary fields.
        trace = go.Box(
            x=[m],
            y=month_stats[m]["values"].tolist(),  # keep original points hidden (Plotly will compute by default but we supply summary fields too)
            boxpoints=False,
            name=m,
            marker=dict(size=4),
            hovertemplate=(
                f"Month: {m}<br>N = {n}<br>"
                "Q1: %{customdata[0]:.3f}<br>Median: %{customdata[1]:.3f}<br>Q3: %{customdata[2]:.3f}<br>"
                "Lower whisker: %{customdata[3]:.3f}<br>Upper whisker: %{customdata[4]:.3f}<extra></extra>"
            ),
            customdata=[[
                s["q1"], s["median"], s["q3"], s["lower_whisker"], s["upper_whisker"]
            ]],
            quartilemethod="linear"  # consistent quartile computation
        )
        # Plotly by default computes its own quartiles/whiskers from y; to ensure explicit whisker positions we will
        # replace the computed stats by adding a second trace of type 'box' specifying the summary stat properties.
        # However Plotly's Box supports specifying q1/median/q3/lowerfence/upperfence directly as attributes.
        # Create a "box-from-stats" using those attributes (y left empty)
        trace_from_stats = go.Box(
            x=[m],
            y=[],  # empty raw y
            name=m,
            showlegend=False,
            hoverinfo="skip",
            boxpoints=False,
            marker=dict(size=4),
            q1=[s["q1"]],
            median=[s["median"]],
            q3=[s["q3"]],
            lowerfence=[s["lower_whisker"]],
            upperfence=[s["upper_whisker"]],
        )
        # We will add only the trace_from_stats to the figure (so hover uses customdata from earlier is not used)
        plotly_traces.append(trace_from_stats)

    # Add PWQO line across months (dashed)
    ref_line = go.Scatter(
        x=month_order,
        y=[30.0]*len(month_order),
        mode="lines",
        name="PWQO = 30 µg/L",
        line=dict(color="black", dash="dash"),
        hoverinfo="skip",
        showlegend=True
    )

    # Combine traces (boxes then ref line)
    fig_plotly = go.Figure(data=plotly_traces + [ref_line])

    # Layout adjustments: plotly_white style, black axis lines, font size 16, no title
    fig_plotly.update_layout(
        template="plotly_white",
        font=dict(size=16),
        xaxis=dict(title="Month", showgrid=True, gridcolor="lightgrey", linecolor="black", tickfont=dict(color="black")),
        yaxis=dict(title="Total phosphorus [µg/L]", showgrid=True, gridcolor="lightgrey", linecolor="black", tickfont=dict(color="black")),
        margin=dict(l=60, r=20, t=10, b=60),
        showlegend=True
    )

    # Save interactive HTML (include_plotlyjs='cdn')
    try:
        fig_plotly.write_html(out_html, include_plotlyjs="cdn", full_html=True)
    except Exception as e:
        print(f"ERROR saving HTML '{out_html}': {e}", file=sys.stderr)
        raise

    # Also save the Plotly-based PNG at 1000x600 via kaleido
    try:
        fig_plotly.write_image(out_png, width=1000, height=600, scale=1)
    except Exception as e:
        # If plotly PNG fails, note but still proceed (PNG was already created by matplotlib)
        print(f"WARNING: writing Plotly PNG '{out_png}' failed: {e}", file=sys.stderr)
        print("Matplotlib PNG was created; ensure 'kaleido' is installed if Plotly PNG is required.", file=sys.stderr)

    # Print created filenames and python filename to terminal as required
    print(out_py)
    print(out_csv)
    print(out_html)
    print(out_png)

if __name__ == "__main__":
    main()