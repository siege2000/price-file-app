# helpers.py
import json
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
import os
import pyodbc

import pandas as pd
from decimal import Decimal, ROUND_HALF_UP 
from datetime import datetime
import uuid

MAX_DESC_LEN = 40


#----------------------------
#Convert $ to cents and vice versa to keep consisten with RxOne currency
#----------------------------
def dollars_to_cents(x) -> int:
    # Safe conversion: 12.34 -> 1234
    try:
        return int(round(float(x) * 100))
    except Exception:
        return 0

def cents_to_dollars(x) -> float:
    # Safe conversion: 1234 -> 12.34
    try:
        return float(x) / 100.0
    except Exception:
        return 0.0

# ----------------------------
# File + template loading
# ----------------------------
def load_file(file) -> pd.DataFrame:
    name = file.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(file)
    if name.endswith(".xlsx") or name.endswith(".xls"):
        return pd.read_excel(file)
    raise ValueError("Unsupported file type (must be .csv or .xlsx)")


def load_templates(path: str = "templates.json") -> Dict[str, Any]:
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"Template file not found: {path}")
    data = json.loads(p.read_text(encoding="utf-8"))
    if "templates" not in data:
        raise ValueError("templates.json must contain a top-level key: 'templates'")
    return data["templates"]


# ----------------------------
# Column helpers
# ----------------------------
def safe_str(series) -> pd.Series:
    return pd.Series(series).astype("string").fillna("").str.strip()


def idx_for(cols_with_none: List[str], colname: Optional[str]) -> int:
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


# ----------------------------
# Description building / normalization
# ----------------------------
def combine_columns(df: pd.DataFrame, columns: List[str], sep: str = " ") -> pd.Series:
    if not columns:
        return pd.Series([""] * len(df), index=df.index)

    parts = [safe_str(df[c]) for c in columns]
    combined = parts[0]
    for p in parts[1:]:
        combined = combined + sep + p

    # clean up extra whitespace
    combined = combined.str.replace(r"\s{2,}", " ", regex=True).str.strip()
    return combined


def normalize_units(tradename: pd.Series) -> pd.Series:
    """
    Remove space between number and unit:
    '75 g' -> '75g', '0.5 ml' -> '0.5ml'
    """
    s = pd.Series(tradename).astype("string").fillna("")
    # Inline (?i) flag is the most reliable across pandas versions.
    pattern = r"(?i)(\d+(?:\.\d+)?)\s+(mg|g|kg|ml|l|mcg)"
    s = s.str.replace(pattern, r"\1\2", regex=True)
    s = s.str.replace(r"\s{2,}", " ", regex=True).str.strip()
    return s
def clean_description(series: pd.Series) -> pd.Series:
    s = series.astype("string").fillna("")

    # Remove trademark + similar symbols
    s = s.str.replace(r"[™®©]", "", regex=True)

    # Remove any other weird unicode symbols but keep letters/numbers/basic punctuation
    s = s.str.replace(r"[^\w\s\-\.,%/()]", "", regex=True)

    # Collapse double spaces
    s = s.str.replace(r"\s{2,}", " ", regex=True).str.strip()

    return s

# ----------------------------
# Money parsing + validation
# ----------------------------
def parse_money(series) -> pd.Series:
    s = pd.Series(series).astype("string").fillna("")
    s = s.str.replace(r"[^\d,.\-]", "", regex=True)

    # Handle comma decimals (EU-ish format)
    if s.str.contains(",", regex=False).mean() > 0.5 and s.str.contains(".", regex=False).mean() < 0.5:
        s = s.str.replace(".", "", regex=False)
        s = s.str.replace(",", ".", regex=False)

    nums =  pd.to_numeric(s, errors="coerce")

    def round_money(x):
        if pd.isna(x):
            return x
        return float(Decimal(str(x)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))
    return nums.apply(round_money)



def desc_len(series) -> pd.Series:
    return safe_str(series).str.len()


# ----------------------------
# Replacement rules (CSV)
# ----------------------------
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

# ----------------------------
# Access DB helpers
# ----------------------------
ACCESS_PASSWORD = "LOCKIE MONDAY"
ACCESS_FILE = "suppliers.mdb"   # <-- CHANGE to your real filename: supplier.mdb or suppliers.mdb

def get_access_conn():
    db_path = os.path.join(os.getcwd(), ACCESS_FILE)
    conn_str = (
        r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
        rf"DBQ={db_path};"
        rf"PWD={ACCESS_PASSWORD};"
    )
    return pyodbc.connect(conn_str)

def load_suppliers():
    with get_access_conn() as conn:
        sql = """
        SELECT SupplierID, SupplierName, SupplierCode
        FROM Suppliers
        ORDER BY SupplierName
        """
        return pd.read_sql(sql, conn)

def new_guid_32() -> str:
    # 32 uppercase hex, no dashes (matches your system)
    return uuid.uuid4().hex.upper()

def load_details_for_supplier(supplier_id: int) -> pd.DataFrame:
    with get_access_conn() as conn:
        sql = """
        SELECT SupplierID,
               Barcode,
               Code,
               Description,
               Pharmacode,
               Cost,
               Retail,
               PartCodeGuid,
               StockcardGuid
        FROM Details
        WHERE SupplierID = ?
        """
        return pd.read_sql(sql, conn, params=[supplier_id])

def upsert_details(
    export_df: pd.DataFrame,
    supplier_id: int,
    mark_updated: bool = True,
) -> tuple[int, int]:
    """
    Upsert into Details.

    Match order:
      1) SupplierID + Barcode  (if barcode present)
      2) SupplierID + Code     (supplier code)

    Update fields:
      Code, Description, Pharmacode, Cost, Retail, Barcode (where appropriate), Updated, LastUpdated

    Insert fields:
      PartCodeGuid, StockcardGuid, SupplierID, Code, Description, Pharmacode, Cost, Retail, Barcode, Updated, LastUpdated
    """
    df = export_df.copy()

    # --- Normalize incoming values ---
    df["SupplierID"] = int(supplier_id)
    df["Code"] = df.get("Supplier_Code", "").fillna("").astype(str).str.strip()          # supplier code
    df["Barcode"] = df.get("Barcode", "").fillna("").astype(str).str.strip()
    df["Description"] = df.get("TradeName", "").fillna("").astype(str).str.slice(0, MAX_DESC_LEN)
    df["Pharmacode"] = df.get("Pharmacode", "").fillna("").astype(str).str.strip()
    df["Cost"] = pd.to_numeric(df.get("Cost", 0), errors="coerce").fillna(0.0)
    df["Retail"] = pd.to_numeric(df.get("Retail", 0), errors="coerce").fillna(0.0)

    now = datetime.now()
    updated_total = 0
    inserted_total = 0

    # --- SQL ---
    update_by_barcode_sql = """
        UPDATE Details
        SET [Code]=?,
            [Description]=?,
            [Pharmacode]=?,
            [Cost]=?,
            [Retail]=?,
            [Updated]=?,
            [LastUpdated]=?
        WHERE [SupplierID]=? AND [Barcode]=?
    """

    update_by_code_sql = """
        UPDATE Details
        SET [Barcode]=?,
            [Description]=?,
            [Pharmacode]=?,
            [Cost]=?,
            [Retail]=?,
            [Updated]=?,
            [LastUpdated]=?
        WHERE [SupplierID]=? AND [Code]=?
    """

    insert_sql = """
        INSERT INTO Details
            ([PartCodeGuid],
             [StockcardGuid],
             [SupplierID],
             [Code],
             [Description],
             [Pharmacode],
             [Cost],
             [Retail],
             [Barcode],
             [Updated],
             [LastUpdated])
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """

    with get_access_conn() as conn:
        cur = conn.cursor()

        for r in df.itertuples(index=False):
            supplier = int(r.SupplierID)
            code = (r.Code or "").strip()
            barcode = (r.Barcode or "").strip()

            desc = r.Description
            pharmacode = (r.Pharmacode or "").strip()
            cost_cents = dollars_to_cents(r.Cost)
            retail_cents = dollars_to_cents(r.Retail)

            did_update = False

            # 1) Try barcode match first
            if barcode:
                cur.execute(
                    update_by_barcode_sql,
                    (code, desc, pharmacode, cost_cents, retail_cents, bool(mark_updated), now, supplier, barcode)
                )
                if cur.rowcount and cur.rowcount > 0:
                    updated_total += int(cur.rowcount)
                    did_update = True

            # 2) Fallback: try supplier code match
            if (not did_update) and code:
                cur.execute(
                    update_by_code_sql,
                    (barcode, desc, pharmacode, cost_cents, retail_cents, bool(mark_updated), now, supplier, code)
                )
                if cur.rowcount and cur.rowcount > 0:
                    updated_total += int(cur.rowcount)
                    did_update = True

            # 3) Insert new
            if not did_update:
                part_guid = new_guid_32()
                stock_guid = new_guid_32()
                cur.execute(
                    insert_sql,
                    (part_guid, stock_guid, supplier, code, desc, pharmacode, cost_cents, retail_cents, barcode, bool(mark_updated), now)
                )
                inserted_total += 1

        conn.commit()

    return updated_total, inserted_total