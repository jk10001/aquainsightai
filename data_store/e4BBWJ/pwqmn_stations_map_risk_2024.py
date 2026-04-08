# filename: pwqmn_stations_map_risk_2024.py
import os
import sys
import pandas as pd
import numpy as np
import plotly.graph_objects as go

def canonicalize_tier(raw):
    """
    Map various possible tier strings to canonical labels:
    "Tier 1", "Tier 2", "Tier 3", "Tier 4".
    Return None if unrecognized.
    """
    if pd.isna(raw):
        return None
    s = str(raw).strip().lower()
    if s == "":
        return None
    # Normalize separators and common words
    s = s.replace("-", " ").replace("_", " ").replace(".", " ").replace(",", " ")
    s = s.replace("tier", "").replace("t", " ").strip()
    s = " ".join(s.split())
    # Try to extract numeric token
    tokens = s.split()
    num = None
    for tok in tokens:
        tok_clean = tok.lstrip("0") or "0"
        if tok_clean.isdigit():
            try:
                num = int(tok_clean)
                break
            except:
                continue
    if num is None and tokens:
        spelled_map = {"one":1, "two":2, "three":3, "four":4}
        tok0 = tokens[0]
        if tok0 in spelled_map:
            num = spelled_map[tok0]
    if num is not None:
        if num == 1:
            return "Tier 1"
        if num == 2:
            return "Tier 2"
        if num == 3:
            return "Tier 3"
        if num == 4:
            return "Tier 4"
    # fallback contains checks
    if "1" == s or "tier 1" in s or "tier1" in s:
        return "Tier 1"
    if "2" == s or "tier 2" in s or "tier2" in s:
        return "Tier 2"
    if "3" == s or "tier 3" in s or "tier3" in s:
        return "Tier 3"
    if "4" == s or "tier 4" in s or "tier4" in s:
        return "Tier 4"
    return None

def estimate_zoom_from_bbox(lat_min, lat_max, lon_min, lon_max):
    """
    Heuristic to pick zoom from lat/lon span.
    """
    lat_span = max(0.0001, lat_max - lat_min)
    lon_span = max(0.0001, lon_max - lon_min)
    max_span = max(lat_span, lon_span)
    if max_span > 20:
        return 4
    if max_span > 10:
        return 5
    if max_span > 5:
        return 6
    if max_span > 2:
        return 7
    if max_span > 1:
        return 8
    if max_span > 0.5:
        return 9
    return 10

def main():
    base_name = "pwqmn_stations_map_risk_2024"
    py_filename = os.path.basename(__file__)
    excel_file = "Ontario_PWQMN_2024.xlsx"
    risk_csv = "station_risk_classification_pwqo_2024_v2.csv"
    out_csv = f"{base_name}.csv"
    out_html = f"{base_name}.html"
    out_png = f"{base_name}.png"

    # Read inputs
    try:
        df_risk = pd.read_csv(risk_csv, dtype=str, na_values=["", "nan", "NaN"])
    except Exception as e:
        print(f"ERROR reading '{risk_csv}': {e}", file=sys.stderr)
        raise

    try:
        df_stations = pd.read_excel(excel_file, sheet_name="Stations", dtype=str, engine="openpyxl")
    except Exception as e:
        print(f"ERROR reading sheet 'Stations' from '{excel_file}': {e}", file=sys.stderr)
        raise

    # Normalize column names
    df_risk.columns = [c.strip() for c in df_risk.columns]
    df_stations.columns = [c.strip() for c in df_stations.columns]

    # Validate required columns
    if "Station number" not in df_risk.columns:
        raise KeyError("Expected 'Station number' column not found in risk CSV.")
    if "Risk tier" not in df_risk.columns:
        raise KeyError("Expected 'Risk tier' column not found in risk CSV.")
    if "STATION" not in df_stations.columns:
        raise KeyError("Expected 'STATION' column not found in Stations sheet.")
    if "LATITUDE" not in df_stations.columns or "LONGITUDE" not in df_stations.columns:
        raise KeyError("Expected 'LATITUDE' and 'LONGITUDE' columns in Stations sheet.")

    # Standardize station IDs to 11-digit zero-padded strings
    df_risk["Station number"] = df_risk["Station number"].fillna("").astype(str).str.strip()
    df_risk["station_padded_id"] = df_risk["Station number"].apply(lambda x: x.zfill(11) if x not in ("", "nan", "NaN") else "")

    df_stations["STATION"] = df_stations["STATION"].fillna("").astype(str).str.strip()
    df_stations["station_padded_id"] = df_stations["STATION"].apply(lambda x: x.zfill(11) if x not in ("", "nan", "NaN") else "")

    # Merge risk table with station coordinates (left join keep all risk rows)
    df_merged = df_risk.merge(
        df_stations[["station_padded_id", "NAME", "LATITUDE", "LONGITUDE"]],
        left_on="station_padded_id",
        right_on="station_padded_id",
        how="left",
        suffixes=("", "_stations")
    )

    # Populate Station name: prefer NAME from Stations if available; else use 'Station name' from risk CSV if present; else blank
    if "Station name" in df_merged.columns:
        df_merged["Station name"] = df_merged["NAME"].where(df_merged["NAME"].notna(), df_merged["Station name"])
    else:
        df_merged["Station name"] = df_merged["NAME"]
    df_merged["Station name"] = df_merged["Station name"].fillna("")

    # Keep original raw LAT/LON strings for potential diagnostics, and create numeric columns for plotting
    df_merged["LATITUDE_raw"] = df_merged["LATITUDE"]
    df_merged["LONGITUDE_raw"] = df_merged["LONGITUDE"]
    df_merged["LATITUDE"] = pd.to_numeric(df_merged["LATITUDE"], errors="coerce")
    df_merged["LONGITUDE"] = pd.to_numeric(df_merged["LONGITUDE"], errors="coerce")

    # Canonicalize Risk tier into a new column 'Risk tier (clean)'
    df_merged["Risk tier (original)"] = df_merged["Risk tier"].fillna("").astype(str)
    df_merged["Risk tier (clean)"] = df_merged["Risk tier (original)"].apply(canonicalize_tier)

    # Determine plotting eligibility and reasons
    reasons = []
    plotted_flags = []
    for _, row in df_merged.iterrows():
        lat = row["LATITUDE"]
        lon = row["LONGITUDE"]
        padded = row.get("station_padded_id", "")
        reason = ""
        plotted = True
        if not padded or str(padded).strip() == "":
            reason = "Missing station ID"
            plotted = False
        elif pd.isna(lat) or pd.isna(lon):
            reason = "Missing coordinates"
            plotted = False
        elif not (-90 <= lat <= 90 and -180 <= lon <= 180):
            reason = "Coordinates out of range"
            plotted = False
        # Also require a valid cleaned tier
        if plotted and (pd.isna(row["Risk tier (clean)"]) or row["Risk tier (clean)"] is None):
            reason = f"Unrecognized risk tier: '{row['Risk tier (original)']}'"
            plotted = False
        reasons.append(reason)
        plotted_flags.append(plotted)

    df_merged["Plotted"] = plotted_flags
    df_merged["Reason"] = reasons

    total_count = int(len(df_merged))
    excluded_count = int((df_merged["Plotted"] == False).sum())

    # DataFrame for plotting: only rows with Plotted True
    df_plot = df_merged[df_merged["Plotted"] == True].copy()

    # If nothing to plot, save empty CSV and exit with message
    if df_plot.shape[0] == 0:
        print(py_filename)
        df_plot_out = pd.DataFrame(columns=["Station number", "Station name", "Risk tier (original)", "Risk tier (clean)", "LATITUDE", "LONGITUDE"])
        df_plot_out.to_csv(out_csv, index=False, encoding="utf-8")
        print(out_csv)
        print("No stations had valid coordinates and recognized tiers to plot.")
        print(f"Total rows in risk file: {total_count}. Excluded: {excluded_count}")
        return

    # Force a single, deterministic discrete colour mapping
    categories = ["Tier 1", "Tier 2", "Tier 3", "Tier 4"]
    color_discrete_map = {
        "Tier 1": "green",
        "Tier 2": "yellow",
        "Tier 3": "orange",
        "Tier 4": "red"
    }

    # Ensure the cleaned tier column is categorical with explicit order
    df_plot["Risk tier (clean)"] = pd.Categorical(df_plot["Risk tier (clean)"], categories=categories, ordered=True)

    # Prepare percent exceed column for hover if present (try to find % total exceedances column)
    percent_col = None
    for c in df_plot.columns:
        if "% total exceedances" in c.lower() or "total exceedances" in c.lower():
            percent_col = c
            break
    if percent_col and percent_col in df_plot.columns:
        df_plot["pct_exceed"] = df_plot[percent_col].fillna("").astype(str)
    else:
        df_plot["pct_exceed"] = df_plot.get("% total exceedances (PWQO analytes)", "")
        df_plot["pct_exceed"] = df_plot["pct_exceed"].fillna("").astype(str)

    # Marker sizes and outline
    marker_size = 8
    outline_extra = 4  # black outline size addition

    # Build a base black marker trace for outlines (single trace)
    base_trace = go.Scattermapbox(
        lat=df_plot["LATITUDE"],
        lon=df_plot["LONGITUDE"],
        mode="markers",
        marker=go.scattermapbox.Marker(size=marker_size + outline_extra, color="black", opacity=1.0),
        hoverinfo="skip",
        showlegend=False,
    )

    # Build one colored trace per canonical tier (only if there are points in that tier)
    colored_traces = []
    for tier in categories:
        sub = df_plot[df_plot["Risk tier (clean)"] == tier]
        if sub.shape[0] == 0:
            continue
        customdata = np.stack([
            sub["station_padded_id"].astype(str).fillna(""),
            sub["Station name"].astype(str).fillna(""),
            sub["Risk tier (original)"].astype(str).fillna(""),
            sub["Risk tier (clean)"].astype(str).fillna(""),
            sub["pct_exceed"].astype(str).fillna("")
        ], axis=-1)
        trace = go.Scattermapbox(
            lat=sub["LATITUDE"],
            lon=sub["LONGITUDE"],
            mode="markers",
            marker=go.scattermapbox.Marker(size=marker_size, color=color_discrete_map[tier], opacity=0.95),
            name=tier,
            customdata=customdata.tolist(),
            hovertemplate=(
                "Station number: %{customdata[0]}<br>"
                "Station name: %{customdata[1]}<br>"
                "Risk tier (original): %{customdata[2]}<br>"
                "Risk tier (clean): %{customdata[3]}<br>"
                "% total exceedances: %{customdata[4]}<extra></extra>"
            ),
            showlegend=True
        )
        colored_traces.append(trace)

    # Combine traces: base (outline) first then colored traces
    all_traces = [base_trace] + colored_traces

    # Create figure
    fig = go.Figure(data=all_traces)

    # Layout adjustments
    fig.update_layout(
        template="plotly_white",
        margin=dict(l=10, r=10, t=10, b=10),
        legend=dict(title="", traceorder="normal"),
        font=dict(size=16),
        title_text="",
    )

    # Compute raw bounding box from plotted points
    lat_min = float(df_plot["LATITUDE"].min())
    lat_max = float(df_plot["LATITUDE"].max())
    lon_min = float(df_plot["LONGITUDE"].min())
    lon_max = float(df_plot["LONGITUDE"].max())

    # Add explicit padding to the bbox before calculating center/zoom
    # Minimum fixed geographic pad
    min_lat_pad = 0.5   # degrees latitude
    min_lon_pad = 0.75  # degrees longitude

    # Relative pad (5% of span)
    lat_span = max(0.0001, lat_max - lat_min)
    lon_span = max(0.0001, lon_max - lon_min)
    rel_lat_pad = 0.05 * lat_span
    rel_lon_pad = 0.05 * lon_span

    # Choose the larger of min fixed pad and relative pad
    lat_pad = max(min_lat_pad, rel_lat_pad)
    lon_pad = max(min_lon_pad, rel_lon_pad)

    # Padded bounds
    lat_min_p = lat_min - lat_pad
    lat_max_p = lat_max + lat_pad
    lon_min_p = lon_min - lon_pad
    lon_max_p = lon_max + lon_pad

    # Compute center and zoom using padded bounds
    center_lat = (lat_min_p + lat_max_p) / 2.0
    center_lon = (lon_min_p + lon_max_p) / 2.0
    zoom = estimate_zoom_from_bbox(lat_min_p, lat_max_p, lon_min_p, lon_max_p)

    # Apply mapbox center and zoom unconditionally (do NOT use fitbounds)
    fig.update_layout(mapbox=dict(style="open-street-map", center=dict(lat=center_lat, lon=center_lon), zoom=zoom))

    # Save HTML (responsive, include_plotlyjs='cdn', full_html=True)
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

    # Prepare output CSV listing ONLY plotted stations (exact data visualized)
    df_out = pd.DataFrame({
        "Station number": df_plot["station_padded_id"].replace("", pd.NA).fillna(""),
        "Station name": df_plot["Station name"].fillna(""),
        "Risk tier (original)": df_plot["Risk tier (original)"].fillna(""),
        # Convert categorical to string for safe CSV writing
        "Risk tier (clean)": df_plot["Risk tier (clean)"].astype(str).fillna(""),
        "LATITUDE": df_plot["LATITUDE"].astype(float),
        "LONGITUDE": df_plot["LONGITUDE"].astype(float)
    })
    df_out.to_csv(out_csv, index=False, encoding="utf-8")

    # Print created filenames and QA summary
    print(py_filename)
    print(out_csv)
    print(out_html)
    print(out_png)
    print(f"Total stations in risk CSV: {total_count}")
    print(f"Stations plotted: {int(df_plot.shape[0])}")
    print(f"Stations excluded (not plotted): {excluded_count}")
    # Breakdown of exclusion reasons
    reason_counts = df_merged[df_merged["Plotted"] == False]["Reason"].value_counts(dropna=False)
    print("Excluded breakdown (reason: count):")
    for reason, cnt in reason_counts.items():
        print(f"{reason}: {cnt}")
    # Risk tier counts (clean) for plotted points
    tier_counts = df_plot["Risk tier (clean)"].value_counts().reindex(categories).fillna(0).astype(int)
    print("Risk tier counts (clean) for plotted points:")
    for tier in categories:
        print(f"{tier}: {int(tier_counts.get(tier, 0))}")
    # Lat/lon ranges
    print(f"Plotted lat range: {lat_min} to {lat_max}; lon range: {lon_min} to {lon_max}")
    # Padded bounds QA
    print(f"Padded lat bounds: {lat_min_p} to {lat_max_p}; padded lon bounds: {lon_min_p} to {lon_max_p}")
    print(f"Map center: {center_lat}, {center_lon}; zoom: {zoom}")

if __name__ == "__main__":
    main()