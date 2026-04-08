# filename: iron_exceedance_map_2024.py
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

def make_padded_id(val):
    if pd.isna(val):
        return ""
    s = str(val).strip()
    if s == "" or s.lower() == "nan":
        return ""
    # If looks like float with .0 convert to int-like string
    try:
        if "." in s:
            f = float(s)
            if f.is_integer():
                s = str(int(f))
    except Exception:
        pass
    return s.zfill(11)

def station_display_from_padded(padded):
    if padded is None:
        return ""
    s = str(padded).strip()
    if s == "" or s.lower() == "nan":
        return ""
    if s.isdigit():
        s2 = s.lstrip("0")
        return s2 if s2 != "" else "0"
    return s

def parse_iron_value(result_raw, result_call, value_qualifier, detection_limit):
    """
    Parse Iron numeric (ug/L) according to rules:
    - If Results starts with '<' OR Value Qualifier == '<' OR Result Call == 'BDL':
        use 0.5 * Detection Limit (DL assumed ug/L)
        if DL missing, try parse number after '<' and take half
        else return np.nan (row excluded)
    - Else parse numeric from Results
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
        # fallback: parse number after '<'
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
                return 0.5 * val
            except Exception:
                return np.nan
        return np.nan
    else:
        # not censored: parse first numeric token
        s = str(result).replace(",", "")
        tokens = s.split()
        for tok in tokens:
            try:
                return float(tok)
            except Exception:
                continue
        return safe_float(s)

def estimate_zoom_from_bbox(lat_min, lat_max, lon_min, lon_max):
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

    base = "iron_exceedance_map_2024"
    out_csv = f"{base}.csv"
    out_html = f"{base}.html"
    out_png = f"{base}.png"

    # Read Data and Stations
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

    # Validate required columns
    required_data_cols = {"Collection Site", "Analyte", "Results", "Result Call", "Detection Limit", "Value Qualifier"}
    missing = required_data_cols - set(df.columns)
    if missing:
        raise KeyError(f"Missing required columns in Data sheet: {missing}")

    required_st_cols = {"STATION", "NAME", "LATITUDE", "LONGITUDE"}
    missing2 = required_st_cols - set(df_st.columns)
    if missing2:
        raise KeyError(f"Missing required columns in Stations sheet: {missing2}")

    # Filter to Iron analyte
    df_iron = df[df["Analyte"].astype(str).str.strip() == "Iron"].copy()

    # If no iron rows, create empty outputs and print filenames
    if df_iron.shape[0] == 0:
        df_empty = pd.DataFrame(columns=[
            "station_padded_id", "Station number", "Station name", "N samples", "N exceedances", "Percent exceedance [%]",
            "LATITUDE [deg]", "LONGITUDE [deg]"
        ])
        df_empty.to_csv(out_csv, index=False, encoding="utf-8")
        # minimal HTML and PNG placeholders
        with open(out_html, "w", encoding="utf-8") as f:
            f.write("<html><body>No Iron data to plot.</body></html>")
        try:
            from PIL import Image
            img = Image.new("RGB", (1000, 600), color=(255,255,255))
            img.save(out_png, format="PNG")
        except Exception:
            pass
        print(py_filename)
        print(out_csv)
        print(out_html)
        print(out_png)
        return

    # Create padded ids for data and stations
    df_iron["station_padded_id"] = df_iron["Collection Site"].apply(make_padded_id)
    df_st["station_padded_id"] = df_st["STATION"].apply(make_padded_id)

    # Parse detection limit numeric and iron numeric values
    df_iron["Detection Limit numeric"] = df_iron["Detection Limit"].apply(safe_float)
    df_iron["Iron_ugL"] = df_iron.apply(
        lambda r: parse_iron_value(
            r.get("Results", ""),
            r.get("Result Call", ""),
            r.get("Value Qualifier", ""),
            r.get("Detection Limit numeric", np.nan)
        ), axis=1
    )

    # Keep only rows with numeric iron values
    df_valid = df_iron[~df_iron["Iron_ugL"].isna()].copy()

    # If no valid parsed numeric rows, produce empty outputs
    if df_valid.shape[0] == 0:
        df_empty = pd.DataFrame(columns=[
            "station_padded_id", "Station number", "Station name", "N samples", "N exceedances", "Percent exceedance [%]",
            "LATITUDE [deg]", "LONGITUDE [deg]"
        ])
        df_empty.to_csv(out_csv, index=False, encoding="utf-8")
        with open(out_html, "w", encoding="utf-8") as f:
            f.write("<html><body>No valid Iron numeric values to plot.</body></html>")
        try:
            from PIL import Image
            img = Image.new("RGB", (1000, 600), color=(255,255,255))
            img.save(out_png, format="PNG")
        except Exception:
            pass
        print(py_filename)
        print(out_csv)
        print(out_html)
        print(out_png)
        return

    # Aggregate per station
    agg = df_valid.groupby("station_padded_id", dropna=False).agg(
        N_samples=("Iron_ugL", "count"),
        N_exceedances=("Iron_ugL", lambda s: (s.astype(float) > 300.0).sum())
    ).reset_index()

    # Compute percent exceedance and cap at 100
    agg["Percent_exceedance"] = 100.0 * agg["N_exceedances"] / agg["N_samples"]
    agg["Percent_exceedance"] = agg["Percent_exceedance"].clip(upper=100.0)

    # Prepare display Station number
    agg["Station number"] = agg["station_padded_id"].apply(station_display_from_padded)

    # Prepare station coords/name lookup, deduplicate stations by padded id (keep first)
    df_st_coords = df_st[["station_padded_id", "NAME", "LATITUDE", "LONGITUDE"]].copy()
    # Deduplicate as requested
    before_dups = df_st_coords.shape[0]
    df_st_coords = df_st_coords.drop_duplicates(subset=["station_padded_id"], keep="first").reset_index(drop=True)
    after_dups = df_st_coords.shape[0]
    # Merge
    df_merged = agg.merge(df_st_coords, on="station_padded_id", how="left")

    # QA: enforce uniqueness of station_padded_id after merge; if duplicates remain, drop duplicates keeping first
    if not df_merged["station_padded_id"].is_unique:
        # count duplicates
        dup_counts = df_merged.groupby("station_padded_id").size()
        # For reproducibility keep first occurrence
        df_merged = df_merged.drop_duplicates(subset=["station_padded_id"], keep="first").reset_index(drop=True)

    # Fill missing station names
    df_merged["Station name"] = df_merged["NAME"].fillna("Unknown (not in Stations sheet)")

    # Convert coords to numeric and drop invalid coords
    df_merged["LATITUDE_num"] = pd.to_numeric(df_merged["LATITUDE"], errors="coerce")
    df_merged["LONGITUDE_num"] = pd.to_numeric(df_merged["LONGITUDE"], errors="coerce")

    mask_valid_coords = df_merged["LATITUDE_num"].notna() & df_merged["LONGITUDE_num"].notna() & \
        df_merged["LATITUDE_num"].between(-90, 90) & df_merged["LONGITUDE_num"].between(-180, 180)

    df_plot = df_merged[mask_valid_coords].copy()

    # If no stations with valid coords, create empty outputs
    if df_plot.shape[0] == 0:
        df_out_empty = pd.DataFrame(columns=[
            "station_padded_id", "Station number", "Station name", "N samples", "N exceedances", "Percent exceedance [%]",
            "LATITUDE [deg]", "LONGITUDE [deg]"
        ])
        df_out_empty.to_csv(out_csv, index=False, encoding="utf-8")
        with open(out_html, "w", encoding="utf-8") as f:
            f.write("<html><body>No stations with valid coordinates to plot.</body></html>")
        try:
            from PIL import Image
            img = Image.new("RGB", (1000, 600), color=(255,255,255))
            img.save(out_png, format="PNG")
        except Exception:
            pass
        print(py_filename)
        print(out_csv)
        print(out_html)
        print(out_png)
        return

    # Prepare CSV output exactly as plotted with required headings and units; include station_padded_id for QA
    df_csv_out = pd.DataFrame({
        "station_padded_id": df_plot["station_padded_id"].astype(str),
        "Station number": df_plot["Station number"].astype(str),
        "Station name": df_plot["Station name"].astype(str),
        "N samples": df_plot["N_samples"].astype(int),
        "N exceedances": df_plot["N_exceedances"].astype(int),
        "Percent exceedance [%]": df_plot["Percent_exceedance"].round(3),
        "LATITUDE [deg]": df_plot["LATITUDE_num"].astype(float),
        "LONGITUDE [deg]": df_plot["LONGITUDE_num"].astype(float)
    })
    # Ensure CSV shows only one table (single CSV)
    df_csv_out.to_csv(out_csv, index=False, encoding="utf-8")

    # Prepare map plotting
    lats = df_plot["LATITUDE_num"].astype(float)
    lons = df_plot["LONGITUDE_num"].astype(float)
    perc = df_plot["Percent_exceedance"].astype(float)

    vmin = 0.0
    vmax = float(perc.max())
    if vmax == vmin:
        vmax = 1.0  # ensure colorbar renders

    # Black outline trace (slightly larger markers)
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

    # Continuous green->yellow->red colorscale anchored 0->max
    colorscale = [
        [0.0, "green"],
        [0.5, "yellow"],
        [1.0, "red"]
    ]

    customdata = np.stack([
        df_plot["station_padded_id"].astype(str),
        df_plot["Station number"].astype(str),
        df_plot["Station name"].astype(str),
        df_plot["N_samples"].astype(int),
        df_plot["N_exceedances"].astype(int),
        df_plot["Percent_exceedance"].round(3).astype(str)
    ], axis=-1)

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
                title="Iron exceedance [%] (Fe > 300 µg/L)",
                titleside="right",
                ticks="outside",
                showticklabels=True
            ),
            opacity=0.95
        ),
        customdata=customdata.tolist(),
        hovertemplate=(
            "Station id: %{customdata[0]}<br>"
            "Station number: %{customdata[1]}<br>"
            "Station name: %{customdata[2]}<br>"
            "N samples: %{customdata[3]}<br>"
            "N exceedances: %{customdata[4]}<br>"
            "Percent exceedance: %{customdata[5]}%<extra></extra>"
        ),
        showlegend=False
    )

    fig = go.Figure(data=[outline_trace, colored_trace])

    # Determine bounding box and padded center/zoom
    lat_min = float(lats.min())
    lat_max = float(lats.max())
    lon_min = float(lons.min())
    lon_max = float(lons.max())

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

    fig.update_layout(
        template="plotly_white",
        mapbox=dict(style="open-street-map", center=dict(lat=center_lat, lon=center_lon), zoom=zoom),
        margin=dict(l=10, r=10, t=10, b=10),
        font=dict(size=16),
        showlegend=False
    )

    # Save HTML responsive with include_plotlyjs='cdn' (no title)
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
        print("Make sure 'kaleido' is installed.", file=sys.stderr)
        raise

    # Print created filenames and python filename as required
    print(py_filename)
    print(out_csv)
    print(out_html)
    print(out_png)
    # QA prints
    print(f"Stations in Stations sheet before dedup: {before_dups}; after dedup: {after_dups}")
    print(f"Stations aggregated (unique padded ids): {int(agg.shape[0])}; stations plotted (valid coords): {int(df_plot.shape[0])}")

if __name__ == "__main__":
    main()