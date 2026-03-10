import re
from decimal import Decimal, ROUND_HALF_UP
from typing import Tuple

import pandas as pd


def parse_money(series) -> pd.Series:
    s = pd.Series(series).astype("string").fillna("")
    s = s.str.replace(r"[^\d,.\-]", "", regex=True)

    # Handle comma decimals (EU-ish format)
    if s.str.contains(",", regex=False).mean() > 0.5 and s.str.contains(".", regex=False).mean() < 0.5:
        s = s.str.replace(".", "", regex=False)
        s = s.str.replace(",", ".", regex=False)

    nums = pd.to_numeric(s, errors="coerce")

    def round_money(x):
        if pd.isna(x):
            return x
        return float(Decimal(str(x)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))

    return nums.apply(round_money)


def load_replacements_csv(file) -> pd.DataFrame:
    rules = pd.read_csv(file)
    rules.columns = [c.strip().lower() for c in rules.columns]
    if "from" not in rules.columns or "to" not in rules.columns:
        raise ValueError("Replacement file must have columns: from,to")
    rules = rules[["from", "to"]].dropna()
    rules["from"] = rules["from"].astype("string").str.strip()
    rules["to"] = rules["to"].astype("string").str.strip()
    rules = rules[rules["from"] != ""]
    return rules


def apply_replacements(
    description_series: pd.Series,
    rules_df: pd.DataFrame,
    whole_word: bool = True,
    case_insensitive: bool = True,
) -> Tuple[pd.Series, int]:
    """
    Supports:
      - Literal rules: from="Nexcare", to="N/C"
      - Regex rules:   from="regex:(...pattern...)", to="...replacement..."
    """
    s = pd.Series(description_series).astype("string").fillna("")
    flags = re.IGNORECASE if case_insensitive else 0
    total = 0

    for _, row in rules_df.iterrows():
        src = str(row["from"])
        dst = str(row["to"])

        is_regex = src.lower().startswith("regex:")
        pattern = src[6:] if is_regex else re.escape(src)

        if not is_regex and whole_word:
            pattern = r"\b" + pattern + r"\b"

        matches = s.str.count(pattern, flags=flags).sum()
        if matches:
            total += int(matches)
            s = s.str.replace(pattern, dst, regex=True, flags=flags)

    s = s.str.replace(r"\s{2,}", " ", regex=True).str.strip()
    return s, total
