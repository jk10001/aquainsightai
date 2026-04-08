# filename: tp_exceedance_map_2024.py
import os
import sys
import pandas as pd
import numpy as np

# Visualization imports
import plotly.graph_objects as go

def safe_float(x):
    try:
        if pd.isna(x):
            return np.nan
        s = str(x).strip().replace(",", "")
        return float(s)
    except Exception:
        return np.nan

def parse_result_value(result_raw, result_call, value_qualifier, detection_limit):
    """
    Parse a Results cell for TP into numeric ug/L based on rules:
    - If Results starts with '<' OR Value Qualifier == '<' OR Result Call == 'BDL':
        -> use 0.5 * Detection Limit (if available)
        -> else try parse number after '<' and use 0.5 * that
        -> else return np.nan (exclude)
    - Else parse numeric value from Results.
    """
    result = "" if pd.isna(result_raw) else str(result_raw).strip()
    rc = "" if pd.isna(result_call) else str(result_call).strip()
    vq = "" if pd.isna(value_qualifier) else str(value_qualifier).strip()
    dl = detection_limit

    censored = False
    if result.startswith("<"):
        censored = True
    if vq == "<":
        censored = True
    if str(rc).upper() == "BDL":
        censored = True

    if censored:
        # prefer detection limit numeric
        dl_num = safe_float(dl)
        if not np.isnan(dl_num):
            return 0.5 * dl_num
        # fallback: parse number after '<' if present
        if result.startswith("<"):
            numpart = result.lstrip("<").strip()
            # extract leading numeric token
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
        # not censored: parse numeric from result (may include qualifiers)
        s = result.replace(",", "")
        # pick first token that parses to float
        tokens = s.split()
        for tok in tokens:
            try:
                return float(tok)
            except Exception:
                continue
        # if none parsed, try safe_float overall
        return safe_float(s)

def make_padded_id(val):
    if pd.isna(val):
        return ""
    s = str(val).strip()
    if s == "" or s.lower() == "nan":
        return ""
    # if looks like a float with .0, reduce to int string
    try:
        if "." in s:
            f = float(s)
            if f.is_integer():
                s = str(int(f))
    except Exception:
        pass
    # zero-pad to 11 chars
    # if non-digit characters present, still zfill to length
    return s.zfill(11)

def estimate_zoom_from_bbox(lat_min, lat_max, lon_min, lon_max):
    # simple heuristic based on maximum span in degrees
    lat_span = max(0.0001, lat_max - lat_min)
    lon_span = max(0.0001, lon_max - lon_min)
    max_span = max(lat_span, lon_span)
    if max_span > 30:
        return 3
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
    py_filename = os.path.basename(__file__)
    excel_file = "Ontario_PWQMN_2024.xlsx"
    sheet_data = "Data"
    sheet_stations = "Stations"

    base = "tp_exceedance_map_2024"
    out_csv = f"{base}.csv"
    out_html = f"{base}.html"
    out_png = f"{base}.png"

    # Read sheets
    try:
        df = pd.read_excel(excel_file, sheet_name=sheet_data, engine="openpyxl", dtype=object)
    except Exception as e:
        print(f"ERROR reading '{sheet_data}' from '{excel_file}': {e}", file=sys.stderr)
        raise

    try:
        df_st = pd.read_excel(excel_file, sheet_name=sheet_stations, engine="openpyxl", dtype=object)
    except Exception as e:
        print(f"ERROR reading '{sheet_stations}' from '{excel_file}': {e}", file=sys.stderr)
        raise

    # Validate columns
    required_data_cols = {"Collection Site", "Analyte", "Results", "Result Call", "Detection Limit", "Value Qualifier"}
    missing = required_data_cols - set(df.columns)
    if missing:
        raise KeyError(f"Missing columns in Data sheet: {missing}")

    required_st_cols = {"STATION", "NAME", "LATITUDE", "LONGITUDE"}
    missing2 = required_st_cols - set(df_st.columns)
    if missing2:
        raise KeyError(f"Missing columns in Stations sheet: {missing2}")

    # Filter to Phosphorus; total
    df_tp = df[df["Analyte"].astype(str).str.strip() == "Phosphorus; total"].copy()

    # No TP rows: produce empty outputs
    if df_tp.shape[0] == 0:
        df_empty = pd.DataFrame(columns=[
            "Station number", "Station name", "N samples", "N exceedances", "Percent exceedance",
            "LATITUDE", "LONGITUDE"
        ])
        df_empty.to_csv(out_csv, index=False)
        print(py_filename)
        print(out_csv)
        print(out_html)
        print(out_png)
        return

    # Create padded station id in Data
    df_tp["station_padded_id"] = df_tp["Collection Site"].apply(make_padded_id)

    # Ensure detection limit numeric
    df_tp["Detection Limit numeric"] = df_tp["Detection Limit"].apply(safe_float)

    # Parse TP numeric values
    df_tp["TP_value_ugL"] = df_tp.apply(
        lambda r: parse_result_value(
            r.get("Results", ""),
            r.get("Result Call", ""),
            r.get("Value Qualifier", ""),
            r.get("Detection Limit numeric", np.nan)
        ), axis=1
    )

    # Keep only rows where we obtained a numeric TP_value_ugL
    df_tp_valid = df_tp[~df_tp["TP_value_ugL"].isna()].copy()

    # Exceedance flag > 30 ug/L
    df_tp_valid["Exceed"] = df_tp_valid["TP_value_ugL"].astype(float) > 30.0

    # Aggregate per station_padded_id
    agg = df_tp_valid.groupby("station_padded_id", dropna=False).agg(
        N_samples=("TP_value_ugL", "count"),
        N_exceedances=("Exceed", "sum")
    ).reset_index()

    agg["Percent_exceedance"] = 100.0 * agg["N_exceedances"] / agg["N_samples"]

    # Prepare station padded ids and display station numbers
    def display_station_number(padded):
        if not padded or pd.isna(padded):
            return ""
        s = str(padded).strip()
        if s.isdigit():
            s2 = s.lstrip("0")
            return s2 if s2 != "" else "0"
        return s

    agg["Station number"] = agg["station_padded_id"].apply(display_station_number)

    # Prepare station name map from Stations sheet
    df_st["station_padded_id"] = df_st["STATION"].apply(make_padded_id)
    st_name_map = pd.Series(df_st["NAME"].values, index=df_st["station_padded_id"]).to_dict()

    # Attach station name, lat, lon by merging
    df_st_coords = df_st[["station_padded_id", "NAME", "LATITUDE", "LONGITUDE"]].copy()
    df_st_coords.rename(columns={"NAME": "Station name"}, inplace=True)

    df_out = agg.merge(df_st_coords, on="station_padded_id", how="left")

    # Drop rows without valid numeric coordinates
    df_out["LATITUDE_num"] = pd.to_numeric(df_out["LATITUDE"], errors="coerce")
    df_out["LONGITUDE_num"] = pd.to_numeric(df_out["LONGITUDE"], errors="coerce")

    # Exclude invalid coordinates (NaN or out of range)
    mask_valid_coords = df_out["LATITUDE_num"].notna() & df_out["LONGITUDE_num"].notna() & \
        (df_out["LATITUDE_num"].between(-90, 90)) & (df_out["LONGITUDE_num"].between(-180, 180))

    df_plot = df_out[mask_valid_coords].copy()

    # If no stations to plot, save empty CSV and exit
    if df_plot.shape[0] == 0:
        df_empty = pd.DataFrame(columns=[
            "Station number", "Station name", "N samples", "N exceedances", "Percent exceedance",
            "LATITUDE", "LONGITUDE"
        ])
        df_empty.to_csv(out_csv, index=False)
        print(py_filename)
        print(out_csv)
        print(out_html)
        print(out_png)
        print("No stations with valid coordinates to plot.")
        return

    # For any missing Station name, fill with Unknown
    df_plot["Station name"] = df_plot["Station name"].fillna("Unknown (not in Stations sheet)")

    # Limit percent exceedance to 0-100 and fill NaN with 0
    df_plot["Percent_exceedance"] = df_plot["Percent_exceedance"].fillna(0.0).clip(lower=0.0, upper=100.0)

    # Prepare data for color scaling: min 0, max = observed max
    vmin = 0.0
    vmax = float(df_plot["Percent_exceedance"].max())

    # If vmax == vmin (e.g., all zeros), set vmax to 1 to allow colorbar rendering
    if vmax == vmin:
        vmax = 1.0

    # Build Plotly figure using Scattermapbox with continuous color scale green->yellow->red
    # Prepare marker outlines by drawing a larger black marker below colored markers
    lats = df_plot["LATITUDE_num"].astype(float)
    lons = df_plot["LONGITUDE_num"].astype(float)
    perc = df_plot["Percent_exceedance"].astype(float)

    # Base outline trace (black slightly larger)
    outline_trace = go.Scattermapbox(
        lat=lats,
        lon=lons,
        mode="markers",
        marker=go.scattermapbox.Marker(
            size=12,
            color="black",
            opacity=1.0
        ),
        hoverinfo="skip",
        showlegend=False
    )

    # Colored trace
    colorscale = [
        [0.0, "green"],
        [0.5, "yellow"],
        [1.0, "red"]
    ]

    colored_trace = go.Scattermapbox(
        lat=lats,
        lon=lons,
        mode="markers",
        marker=go.scattermapbox.Marker(
            size=8,
            color=perc,
            colorscale=colorscale,
            cmin=vmin,
            cmax=vmax,
            showscale=True,
            colorbar=dict(
                title="Total Phosphorus exceedance [%] (TP > 30 µg/L)",
                titleside="right",
                ticks="outside",
                showticklabels=True
            ),
            opacity=0.95
        ),
        customdata=np.stack([
            df_plot["Station number"].astype(str),
            df_plot["Station name"].astype(str),
            df_plot["N_samples"].astype(int),
            df_plot["N_exceedances"].astype(int),
            df_plot["Percent_exceedance"].round(1).astype(str)
        ], axis=-1).tolist(),
        hovertemplate=(
            "Station number: %{customdata[0]}<br>"
            "Station name: %{customdata[1]}<br>"
            "N samples: %{customdata[2]}<br>"
            "N exceedances: %{customdata[3]}<br>"
            "Percent exceedance: %{customdata[4]}%<extra></extra>"
        ),
        showlegend=False
    )

    fig = go.Figure(data=[outline_trace, colored_trace])

    # Determine bounding box and center/zoom
    lat_min = float(lats.min())
    lat_max = float(lats.max())
    lon_min = float(lons.min())
    lon_max = float(lons.max())

    # Add padding: minimum fixed pad and relative pad
    min_lat_pad = 0.5
    min_lon_pad = 0.75
    lat_span = max(0.0001, lat_max - lat_min)
    lon_span = max(0.0001, lon_max - lon_min)
    rel_lat_pad = 0.05 * lat_span
    rel_lon_pad = 0.05 * lon_span
    lat_pad = max(min_lat_pad, rel_lat_pad)
    lon_pad = max(min_lon_pad, rel_lon_pad)

    lat_min_p = lat_min - lat_pad
    lat_max_p = lat_max + lat_pad
    lon_min_p = lon_min - lon_pad
    lon_max_p = lon_max + lon_pad

    center_lat = (lat_min_p + lat_max_p) / 2.0
    center_lon = (lon_min_p + lon_max_p) / 2.0
    zoom = estimate_zoom_from_bbox(lat_min_p, lat_max_p, lon_min_p, lon_max_p)

    # Layout per requirements: plotly_white style with font size 16, no title
    fig.update_layout(
        template="plotly_white",
        mapbox=dict(
            style="open-street-map",
            center=dict(lat=center_lat, lon=center_lon),
            zoom=zoom
        ),
        margin=dict(l=10, r=10, t=10, b=10),
        font=dict(size=16),
        showlegend=False
    )

    # Save CSV of plotted data (only exact fields visualized)
    df_csv_out = pd.DataFrame({
        "Station number": df_plot["Station number"].astype(str),
        "Station name": df_plot["Station name"].astype(str),
        "N samples": df_plot["N_samples"].astype(int),
        "N exceedances": df_plot["N_exceedances"].astype(int),
        "Percent exceedance": df_plot["Percent_exceedance"].round(3),
        "LATITUDE": df_plot["LATITUDE_num"].astype(float),
        "LONGITUDE": df_plot["LONGITUDE_num"].astype(float)
    })
    df_csv_out.to_csv(out_csv, index=False, encoding="utf-8")

    # Save HTML (include_plotlyjs='cdn', responsive)
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

    # Print created filenames and python filename to terminal as required
    print(py_filename)
    print(out_csv)
    print(out_html)
    print(out_png)

if __name__ == "__main__":
    main()