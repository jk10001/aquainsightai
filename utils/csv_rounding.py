"""Utilities for rounding CSV numeric data to significant figures."""

from __future__ import annotations

from pathlib import Path

import numpy as np
import pandas as pd


EXCLUDE_KEYWORDS: tuple[str, ...] = (
    "date",
    "time",
    "year",
    "month",
    "day",
    "datetime",
    "timestamp",
    "epoch",
    "tz",
    "index",
    "idx",
    "record",
    "id_",
    "_id",
    "uid",
    "uuid",
    "key",
    "code",
    "no.",
    "num",
    "number",
    "seq",
    "sequence",
    "order",
    "count",
    "rank",
    "label",
    "desc",
    "description",
    "flag",
    "status",
    "class",
    "type",
    "category",
    "group",
    "kind",
    "postcode",
    "zip",
    "batch",
    "lot",
    "run",
    "northing",
    "easting",
    "station",
    "lat",
    "lon",
    "latitude",
    "longitude",
)

OVERRIDE_INCLUDE_KEYWORDS: tuple[str, ...] = (
    "m3",
    # Explicit time-rate words
    "/day",
    "per day",
    "per_d",
    "max day",
    "average day",
    "mean day",
    "maximum day",
    "max month",
    "average month",
    "mean month",
    "maximum month",
    "/week",
    "per week",
    "/d",
    "per d",
    "per-day",
    "/hour",
    "per hour",
    "/h",
    "per h",
    "/hr",
    "per hr",
    "/min",
    "per min",
    "per minute",
    "/s",
    "per s",
    "per sec",
    "per second",
    # Flow units
    "m3/d",
    "m³/d",
    "l/d",
    "ml/d",
    "kl/d",
    "m3/s",
    "m³/s",
    "cms",
    "cmd",
    "mld",
    "gpm",
    "cfs",
    "l/s",
    "L/s",
    "lps",
    # Mass-rate
    "kg/d",
    "kg/day",
    "g/d",
    "mg/d",
    "t/d",
    # Loads and rates
    "load",
    "mass rate",
    # General rate structures
    "per unit",
    "per run",
    "per cycle",
    "per event",
    "per sample",
    # Numeric units that could look categorical
    "count per",
    "/100ml",
    "/l",
    "index",
    "ratio",
)


def csv_sigfig_to_string(
    csv_path: str | Path,
    sig_figs: int = 5,
    *,
    exclude_keywords: tuple[str, ...] = EXCLUDE_KEYWORDS,
    override_include_keywords: tuple[str, ...] = OVERRIDE_INCLUDE_KEYWORDS,
) -> str:
    """
    Return a CSV string where **numeric columns** are rounded to `sig_figs`
    significant figures *in place* (column names stay exactly the same).

    A column is **left untouched** if
    • pandas recognises it as datetime, **or**
    • its name contains any of `exclude_keywords` **and does NOT contain any of
      `override_include_keywords`**.

    Original column order is preserved.
    """

    csv_path = Path(csv_path)
    if csv_path.suffix.lower() != ".csv":
        raise ValueError("Input file must be a .csv")

    def _round_sig(x: float):
        """Round x to the required significant figures."""
        if pd.isna(x) or x == 0:
            return x
        decimals = sig_figs - int(np.floor(np.log10(abs(x)))) - 1
        y = round(x, decimals)
        # Drop ".0" if the rounded value is an integer
        if decimals <= 0:
            y = int(y)
        return y

    # Read CSV
    df = pd.read_csv(csv_path)

    for col in df.columns:
        series = df[col]

        # Skip datetime-like columns
        if pd.api.types.is_datetime64_any_dtype(series):
            continue

        # Only operate on numeric columns
        if not pd.api.types.is_numeric_dtype(series):
            continue

        col_name_lower = str(col).lower()

        has_override = any(
            override_kw in col_name_lower for override_kw in override_include_keywords
        )
        has_exclude = any(
            exclude_kw in col_name_lower for exclude_kw in exclude_keywords
        )

        # If it matches an exclude keyword AND does not match any override,
        # skip rounding this column.
        if has_exclude and not has_override:
            continue

        # Otherwise, round numeric values in this column
        df[col] = df[col].apply(_round_sig)

    # Return CSV as string (no index)
    return df.to_csv(index=False)
