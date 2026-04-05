from typing import Any, Dict, List, Optional

import pandas as pd


def safe_str(series) -> pd.Series:
    """Convert a series to stripped strings, replacing NaN with empty string.

    Integer-valued floats (e.g. 1263310.0 read from Excel) are converted
    without the decimal suffix so barcodes/codes match correctly.
    Leading zeros are preserved — e.g. '074469513951' stays '074469513951'.
    """
    s = pd.Series(series)
    numeric = pd.to_numeric(s, errors="coerce")
    is_whole_float = numeric.notna() & (numeric % 1 == 0)
    result = s.astype("string").fillna("").str.strip()
    if is_whole_float.any():
        as_int = numeric[is_whole_float].astype("Int64").astype("string")
        # Only apply the int conversion where the string didn't already have a
        # leading zero — preserves barcodes like '074469513951' as-is.
        original = result[is_whole_float]
        keep_original = original.str.startswith("0")
        result[is_whole_float] = original.where(keep_original, as_int)
    return result


def idx_for(cols_with_none: List[str], colname: Optional[str]) -> int:
    """Return the index of colname in cols_with_none, or 0 if not found."""
    if not colname:
        return 0
    try:
        return cols_with_none.index(colname)
    except ValueError:
        return 0


def tpl_field_default(df: pd.DataFrame, field_cfg: Dict[str, Any]) -> Optional[str]:
    """
    Prefer exact 'column' if present. Fall back to fuzzy 'match_terms' if provided.
    """
    exact = field_cfg.get("column")
    if exact and exact in df.columns:
        return exact

    terms = field_cfg.get("match_terms", []) or []
    if terms:
        for term in terms:
            t = str(term).lower()
            for col in df.columns:
                if t in col.lower():
                    return col

    return None


def tpl_desc_defaults(df: pd.DataFrame, desc_cfg: Dict[str, Any]) -> List[str]:
    """
    Prefer exact 'columns' list. Fall back to 'columns_match_terms'.
    """
    exact_cols = desc_cfg.get("columns")
    if exact_cols:
        return [c for c in exact_cols if c in df.columns]

    terms = desc_cfg.get("columns_match_terms", []) or []
    if terms:
        out = []
        for col in df.columns:
            cl = col.lower()
            if any(str(t).lower() in cl for t in terms):
                out.append(col)
        return out

    return []
