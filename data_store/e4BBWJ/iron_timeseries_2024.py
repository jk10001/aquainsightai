# filename: iron_timeseries_2024.py
import os
import sys
from datetime import datetime
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

def parse_iron_result(result_raw, result_call, value_qualifier, detection_limit):
    """
    Return tuple (iron_ugL (float or np.nan), censored_flag (bool))
    Rules:
      - If Results starts with '<' OR Value Qualifier == '<' OR Result Call == 'BDL':
          iron = 0.5 * detection_limit (assume DL in ug/L). censored_flag = True
      - Else parse numeric from Results. censored_flag = False
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
            return (0.5 * dl_num, True)
        # fallback: try parse numeric after '<'
        if isinstance(result, str) and result.startswith("<"):
            tail = result.lstrip("<").strip()
            token = ""
            for ch in tail:
                if ch.isdigit() or ch in ".-+eE":
                    token += ch
                else:
                    break
            try:
                val = float(token)
                return (0.5 * val, True)
            except Exception:
                return (np.nan, True)
        return (np.nan, True)

    # not censored -> parse numeric token
    if isinstance(result, str):
        s = result.replace(",", "")
        tokens = s.split()
        for tok in tokens:
            try:
                val = float(tok)
                return (val, False)
            except Exception:
                continue
    # final fallback
    return (safe_float(result), False)

def main():
    py_filename = os.path.basename(__file__)
    excel_file = "Ontario_PWQMN_2024.xlsx"
    sheet_name = "Data"

    base = "iron_timeseries_2024"
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

    # Validate required columns
    required = {"Collected", "Analyte", "Results", "Result Call", "Value Qualifier", "Detection Limit", "Collection Site"}
    missing = required - set(df.columns)
    if missing:
        raise KeyError(f"Missing required columns in Data sheet: {missing}")

    # Filter to Analyte == "Iron"
    df_iron = df[df["Analyte"].astype(str).str.strip() == "Iron"].copy()

    # Parse Collected to datetime (mixed formats); coerce errors to NaT and drop them
    # Try infer_datetime_format and also attempt forgiving parsing for strings
    df_iron["Collected date"] = pd.to_datetime(df_iron["Collected"], errors="coerce", infer_datetime_format=True)

    # For rows where parsing failed, try a secondary attempt by stripping time and common formats
    mask_na_dates = df_iron["Collected date"].isna()
    if mask_na_dates.any():
        def try_parse_mixed(x):
            if pd.isna(x):
                return pd.NaT
            s = str(x).strip()
            # Attempt common alternative formats
            fmts = ["%m/%d/%Y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y"]
            for f in fmts:
                try:
                    return pd.to_datetime(datetime.strptime(s, f))
                except Exception:
                    continue
            # Last resort: pandas parse
            try:
                return pd.to_datetime(s, errors="coerce", infer_datetime_format=True)
            except Exception:
                return pd.NaT
        df_iron.loc[mask_na_dates, "Collected date"] = df_iron.loc[mask_na_dates, "Collected"].apply(try_parse_mixed)

    # Drop rows with invalid/unparsed dates
    df_iron = df_iron[~df_iron["Collected date"].isna()].copy()

    # Compute Iron [µg/L] and censored flag
    parsed = df_iron.apply(
        lambda r: parse_iron_result(
            r.get("Results", ""),
            r.get("Result Call", ""),
            r.get("Value Qualifier", ""),
            r.get("Detection Limit", np.nan)
        ), axis=1
    )
    # parsed is a Series of tuples; expand
    df_iron["Iron_value"], df_iron["Censored_flag"] = zip(*parsed)

    # Drop rows where Iron_value is NaN or <= 0 (required for log scale)
    df_iron = df_iron[df_iron["Iron_value"].notna()].copy()
    # Convert to numeric float
    df_iron["Iron_value"] = df_iron["Iron_value"].astype(float)
    df_iron = df_iron[df_iron["Iron_value"] > 0].copy()

    # Prepare plotted dataframe with required CSV columns:
    # Collected date [YYYY-MM-DD], Collection site, Iron [µg/L], Censored flag [true/false], Detection limit [µg/L], Results (raw)
    df_plot = pd.DataFrame({
        "Collected date [YYYY-MM-DD]": df_iron["Collected date"].dt.strftime("%Y-%m-%d"),
        "Collection site": df_iron["Collection Site"].astype(str).fillna(""),
        "Iron [µg/L]": df_iron["Iron_value"].astype(float),
        "Censored flag [true/false]": df_iron["Censored_flag"].apply(lambda x: "true" if bool(x) else "false"),
        "Detection limit [µg/L]": df_iron["Detection Limit"].apply(lambda x: "" if pd.isna(x) else str(x)),
        "Results (raw)": df_iron["Results"].astype(str).fillna("")
    })

    # Sort by date for plotting convenience (not required but helpful)
    df_plot["Collected date_dt"] = pd.to_datetime(df_plot["Collected date [YYYY-MM-DD]"], format="%Y-%m-%d", errors="coerce")
    df_plot = df_plot.sort_values("Collected date_dt").reset_index(drop=True)

    # Save CSV of exact plotted dataset (only the columns requested, in that order)
    df_csv_out = df_plot.loc[:, [
        "Collected date [YYYY-MM-DD]",
        "Collection site",
        "Iron [µg/L]",
        "Censored flag [true/false]",
        "Detection limit [µg/L]",
        "Results (raw)"
    ]].copy()
    df_csv_out.to_csv(out_csv, index=False, encoding="utf-8")

    # Build Plotly scatter (time series) with log y-axis
    # Marker specs: small size (6), high transparency (opacity 0.2)
    trace_points = go.Scatter(
        x=df_plot["Collected date_dt"],
        y=df_plot["Iron [µg/L]"],
        mode="markers",
        marker=dict(
            size=6,
            color="steelblue",
            opacity=0.22,
            line=dict(width=0)
        ),
        name="Iron samples",
        hovertemplate=(
            "Date: %{x|%Y-%m-%d}<br>"
            "Iron [µg/L]: %{y:.3f}<br>"
            "Collection site: %{customdata[0]}<br>"
            "Censored: %{customdata[1]}<br>"
            "Results (raw): %{customdata[2]}<extra></extra>"
        ),
        customdata=np.stack([
            df_csv_out["Collection site"].astype(str),
            df_csv_out["Censored flag [true/false]"].astype(str),
            df_csv_out["Results (raw)"].astype(str)
        ], axis=-1)
    )

    # PWQO horizontal line at 300 µg/L
    trace_pwqo = go.Scatter(
        x=[df_plot["Collected date_dt"].min(), df_plot["Collected date_dt"].max()],
        y=[300.0, 300.0],
        mode="lines",
        line=dict(color="black", dash="dash"),
        name="PWQO = 300 µg/L",
        hoverinfo="skip"
    )

    # Compose figure
    fig = go.Figure(data=[trace_points, trace_pwqo])

    # Layout per requirements: plotly_white, black axis lines and ticks, font size 16, no title
    fig.update_layout(
        template="plotly_white",
        font=dict(size=16),
        xaxis=dict(
            title="Collected date",
            showgrid=True,
            gridcolor="lightgrey",
            linecolor="black",
            tickformat="%Y-%m-%d"
        ),
        yaxis=dict(
            title="Iron [µg/L]",
            type="log",
            showgrid=True,
            gridcolor="lightgrey",
            linecolor="black",
            # Ensure ticks are shown in reasonable places; allow automatic tick spacing
        ),
        margin=dict(l=60, r=20, t=10, b=60),
        showlegend=True
    )

    # Ensure gridlines are visible for log scale: set minor/major grid visible by enabling autorange ticks
    # Plotly will handle log ticks; we keep default ticks for readability.

    # Save HTML (responsive, include_plotlyjs='cdn')
    try:
        fig.write_html(out_html, include_plotlyjs="cdn", full_html=True)
    except Exception as e:
        print(f"ERROR saving HTML '{out_html}': {e}", file=sys.stderr)
        raise

    # Save PNG via kaleido 1000x600
    try:
        fig.write_image(out_png, width=1000, height=600, scale=1)
    except Exception as e:
        print(f"ERROR saving PNG '{out_png}': {e}", file=sys.stderr)
        print("Make sure the 'kaleido' package is installed.", file=sys.stderr)
        raise

    # Print created filenames and python filename to terminal as required
    print(out_py)
    print(out_csv)
    print(out_html)
    print(out_png)

if __name__ == "__main__":
    # local import for numpy used in customdata stack
    import numpy as np
    main()