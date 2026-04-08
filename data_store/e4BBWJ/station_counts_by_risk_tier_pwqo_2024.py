# filename: station_counts_by_risk_tier_pwqo_2024.py
import os
import sys
import pandas as pd

def canonicalize_tier(raw):
    """
    Map various possible tier strings to canonical labels:
    "Tier 1", "Tier 2", "Tier 3", "Tier 4".
    Return None if unrecognized.
    """
    if pd.isna(raw):
        return None
    s = str(raw).strip()
    if s == "":
        return None
    s_low = s.lower()
    # direct matches
    if "tier 1" in s_low or s_low == "tier1" or s_low == "1" or s_low.endswith(" 1") or s_low.startswith("1"):
        return "Tier 1"
    if "tier 2" in s_low or s_low == "tier2" or s_low == "2" or s_low.endswith(" 2") or s_low.startswith("2"):
        return "Tier 2"
    if "tier 3" in s_low or s_low == "tier3" or s_low == "3" or s_low.endswith(" 3") or s_low.startswith("3"):
        return "Tier 3"
    if "tier 4" in s_low or s_low == "tier4" or s_low == "4" or s_low.endswith(" 4") or s_low.startswith("4"):
        return "Tier 4"
    # try to find digit in string
    for ch in s_low:
        if ch.isdigit():
            if ch == "1":
                return "Tier 1"
            if ch == "2":
                return "Tier 2"
            if ch == "3":
                return "Tier 3"
            if ch == "4":
                return "Tier 4"
    return None

def main():
    py_filename = os.path.basename(__file__)
    input_csv = "station_risk_classification_pwqo_2024_v2.csv"
    out_csv = "station_counts_by_risk_tier_pwqo_2024.csv"

    # Read input CSV
    try:
        df = pd.read_csv(input_csv, dtype=str)
    except Exception as e:
        print(f"ERROR reading '{input_csv}': {e}", file=sys.stderr)
        raise

    # Validate required columns
    required = {"Risk tier", "Station number"}
    missing = required - set(df.columns)
    if missing:
        raise KeyError(f"Missing required column(s) in '{input_csv}': {missing}")

    # Ensure Station number is string and strip whitespace
    df["Station number"] = df["Station number"].fillna("").astype(str).str.strip()

    # Canonicalize Risk tier
    df["Risk tier canonical"] = df["Risk tier"].apply(canonicalize_tier)

    # For any rows where canonicalization failed, keep original trimmed value as fallback if it already matches expected patterns
    fallback_mask = df["Risk tier canonical"].isna() & df["Risk tier"].notna()
    if fallback_mask.any():
        # attempt simple strip and title-case check
        def fallback(x):
            s = str(x).strip()
            s_title = s.title()
            if s_title in {"Tier 1","Tier 2","Tier 3","Tier 4"}:
                return s_title
            return None
        df.loc[fallback_mask, "Risk tier canonical"] = df.loc[fallback_mask, "Risk tier"].apply(fallback)

    # Now group and count unique Station number per canonical tier
    # Treat empty Station number as not counted (drop)
    df_counts = (
        df[df["Station number"].astype(str).str.strip() != ""]
        .groupby("Risk tier canonical", dropna=False)["Station number"]
        .nunique()
        .rename("Number of stations [-]")
        .reset_index()
    )

    # Prepare full tiers list to ensure presence of all four tiers
    tiers = ["Tier 1", "Tier 2", "Tier 3", "Tier 4"]
    # Build a DataFrame with all tiers and join counts (fill zeros)
    df_all = pd.DataFrame({"Risk tier": tiers})
    df_counts = df_counts.rename(columns={"Risk tier canonical": "Risk tier"})
    df_merged = df_all.merge(df_counts, on="Risk tier", how="left")
    df_merged["Number of stations [-]"] = df_merged["Number of stations [-]"].fillna(0).astype(int)

    # Compute total across all tiers (unique Station numbers across the whole file)
    # Use non-empty Station numbers only
    unique_stations_all = df.loc[df["Station number"].astype(str).str.strip() != "", "Station number"].nunique()
    total_row = pd.DataFrame({
        "Risk tier": ["Total (all tiers)"],
        "Number of stations [-]": [int(unique_stations_all)]
    })

    # Final output: tiers rows followed by total row
    df_out = pd.concat([df_merged, total_row], ignore_index=True)

    # Save CSV with required single header row
    df_out.to_csv(out_csv, index=False, encoding="utf-8")

    # Print created filenames (python file and output CSV) to terminal
    print(py_filename)
    print(out_csv)

if __name__ == "__main__":
    main()