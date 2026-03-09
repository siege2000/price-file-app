from typing import List

import pandas as pd

from column_helpers import safe_str


def combine_columns(df: pd.DataFrame, columns: List[str], sep: str = " ") -> pd.Series:
    if not columns:
        return pd.Series([""] * len(df), index=df.index)

    parts = [safe_str(df[c]) for c in columns]
    combined = parts[0]
    for p in parts[1:]:
        combined = combined + sep + p

    combined = combined.str.replace(r"\s{2,}", " ", regex=True).str.strip()
    return combined


def normalize_units(tradename: pd.Series) -> pd.Series:
    """
    Remove space between number and unit:
    '75 g' -> '75g', '0.5 ml' -> '0.5ml'
    """
    s = pd.Series(tradename).astype("string").fillna("")
    pattern = r"(?i)(\d+(?:\.\d+)?)\s+(mg|g|kg|ml|l|mcg)"
    s = s.str.replace(pattern, r"\1\2", regex=True)
    s = s.str.replace(r"\s{2,}", " ", regex=True).str.strip()
    return s


def clean_description(series: pd.Series) -> pd.Series:
    s = series.astype("string").fillna("")
    s = s.str.replace(r"[™®©]", "", regex=True)
    s = s.str.replace(r"[^\w\s\-\.,%/()]", "", regex=True)
    s = s.str.replace(r"\s{2,}", " ", regex=True).str.strip()
    return s


def desc_len(series) -> pd.Series:
    return safe_str(series).str.len()
