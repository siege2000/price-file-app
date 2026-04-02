# app_state.py — Session-state helpers and JSON persistence for the Streamlit app
from __future__ import annotations

import json
from pathlib import Path

import pandas as pd
import streamlit as st

SUPPLIER_SETTINGS_FILE = "data/supplier_settings.json"
SUPPLIER_RULES_FILE = "data/supplier_rules.json"


# ── Supplier settings JSON (supplier_settings.json) ───────────────────────────

def read_all_supplier_settings() -> dict:
    try:
        p = Path(SUPPLIER_SETTINGS_FILE)
        if p.exists():
            return json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        pass
    return {}


def save_supplier_settings(supplier_id: int, sheet_key: str, sheet_settings: dict):
    """Save settings for one supplier+sheet, enforcing max-6 sheet layouts."""
    all_s = read_all_supplier_settings()
    sup_key = str(supplier_id)
    entry = all_s.get(sup_key, {})
    # Migrate old flat format → sheet-keyed
    if "sheets" not in entry and ("template" in entry or "columns" in entry):
        entry = {"sheets": {"_default": entry}}
    if "sheets" not in entry:
        entry = {"sheets": {}}
    sheets = entry["sheets"]
    sheets[sheet_key] = sheet_settings
    if len(sheets) > 6:
        del sheets[next(iter(sheets))]
    entry["sheets"] = sheets
    all_s[sup_key] = entry
    Path(SUPPLIER_SETTINGS_FILE).write_text(
        json.dumps(all_s, indent=2, ensure_ascii=False), encoding="utf-8"
    )


def get_supplier_sheet_layouts(supplier_id: int) -> dict:
    """Return the sheet layouts dict for a supplier, handling flat-format migration."""
    all_s = read_all_supplier_settings()
    entry = all_s.get(str(supplier_id), {})
    if not entry:
        return {}
    if "sheets" not in entry and ("template" in entry or "columns" in entry):
        return {"_default": entry}
    return entry.get("sheets", {})


def get_sheet_settings(supplier_id: int, sheet_key: str) -> dict | None:
    """Return settings for a specific (supplier_id, sheet_key), or None."""
    sheets = get_supplier_sheet_layouts(supplier_id)
    return sheets.get(sheet_key) or (next(iter(sheets.values()), None) if sheets else None)


def apply_supplier_settings_to_state(
    supplier_id: int,
    sheet_key: str,
    df_columns: list[str],
):
    """
    Push saved supplier+sheet settings into session_state widget keys.
    Call this BEFORE rendering the mapping widgets so Streamlit picks up the values.
    Tracks (supplier_id, sheet_key) in 'price_settings_loaded_for' to avoid re-applying
    on every rerun.
    """
    loaded_for = st.session_state.get("price_settings_loaded_for")
    if loaded_for == (supplier_id, sheet_key):
        return  # already applied this run

    settings = get_sheet_settings(supplier_id, sheet_key)
    if not settings:
        st.session_state["price_settings_loaded_for"] = (supplier_id, sheet_key)
        return

    # Template
    tpl = settings.get("template")
    if tpl:
        st.session_state["price_template"] = tpl

    # Skip rows
    sr = settings.get("skip_rows")
    if sr is not None:
        st.session_state["price_skip_rows"] = int(sr)

    # Column mappings
    key_map = {
        "plu": "price_col_plu",
        "pharmacode": "price_col_pharmacode",
        "retail": "price_col_retail",
        "cost": "price_col_cost",
        "barcode": "price_col_barcode",
        "sc": "price_col_sc",
        "outer": "price_col_outer",
    }
    cols_with_none = ["(None)"] + df_columns
    for json_key, state_key in key_map.items():
        saved = settings.get("columns", {}).get(json_key)
        if saved and saved in cols_with_none:
            st.session_state[state_key] = saved

    # TradeName settings
    desc = settings.get("description", {})
    saved_cols = desc.get("columns", [])
    if saved_cols:
        existing = [c for c in saved_cols if c in df_columns]
        remaining = [c for c in df_columns if c not in existing]
        st.session_state["price_desc_cols_ordered"] = existing + remaining
        st.session_state["price_desc_cols_selected"] = existing
    prefix = desc.get("prefix")
    if prefix is not None:
        st.session_state["price_prefix"] = prefix
    sep = desc.get("separator")
    if sep is not None:
        st.session_state["price_sep"] = sep
    tc = desc.get("title_case")
    if tc is not None:
        st.session_state["price_title_case"] = bool(tc)
    nu = desc.get("normalize_units")
    if nu is not None:
        st.session_state["price_norm_units"] = bool(nu)

    st.session_state["price_settings_loaded_for"] = (supplier_id, sheet_key)


# ── Supplier rules JSON (supplier_rules.json) ─────────────────────────────────

def read_all_supplier_rules() -> dict:
    try:
        p = Path(SUPPLIER_RULES_FILE)
        if p.exists():
            return json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        pass
    return {}


def get_supplier_rules(supplier_id: int) -> pd.DataFrame | None:
    """Return rules DataFrame for supplier, or None if none defined."""
    if supplier_id is None:
        return None
    try:
        data = read_all_supplier_rules()
        raw = data.get(str(supplier_id), [])
        if not raw:
            return None
        rules = pd.DataFrame(raw)
        rules.columns = [c.strip().lower() for c in rules.columns]
        if "from" not in rules.columns or "to" not in rules.columns:
            return None
        rules = rules[["from", "to"]].dropna()
        rules["from"] = rules["from"].astype("string").str.strip()
        rules["to"] = rules["to"].astype("string").str.strip()
        return rules[rules["from"] != ""].reset_index(drop=True)
    except Exception:
        return None


def save_supplier_rules(supplier_id: int, rules: list[dict]):
    """Save rules list for supplier_id back to supplier_rules.json."""
    all_rules = read_all_supplier_rules()
    all_rules[str(supplier_id)] = [r for r in rules if r.get("from", "").strip()]
    Path(SUPPLIER_RULES_FILE).write_text(
        json.dumps(all_rules, indent=2, ensure_ascii=False), encoding="utf-8"
    )


# ── File loading helpers ───────────────────────────────────────────────────────

def get_excel_sheets(uploaded_file) -> list[str]:
    """Return sheet names for an Excel upload (reads the file object)."""
    try:
        xf = pd.ExcelFile(uploaded_file)
        sheets = xf.sheet_names
        uploaded_file.seek(0)
        return sheets
    except Exception:
        return []


def load_uploaded_file(uploaded_file, sheet_name=None, skip_rows: int = 0) -> pd.DataFrame:
    """Load a Streamlit UploadedFile into a DataFrame."""
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        for enc in ("utf-8-sig", "cp1252", "latin-1"):
            try:
                uploaded_file.seek(0)
                df = pd.read_csv(uploaded_file, encoding=enc, dtype=str)
                break
            except UnicodeDecodeError:
                continue
        else:
            uploaded_file.seek(0)
            df = pd.read_csv(uploaded_file, encoding="latin-1", dtype=str)
    elif name.endswith((".xlsx", ".xls")):
        uploaded_file.seek(0)
        df = pd.read_excel(
            uploaded_file,
            sheet_name=sheet_name if sheet_name is not None else 0,
            dtype=str,
        )
    else:
        raise ValueError(f"Unsupported file type: {uploaded_file.name}")

    if skip_rows > 0:
        df = df.iloc[skip_rows:].reset_index(drop=True)

    # Coerce ID-like columns to string (strip .0 suffix)
    for c in df.columns:
        if c.strip().lower() in {"pharmacode", "pharma code", "barcode",
                                  "barcode/ean", "ean", "sap code", "plu"}:
            df[c] = (
                df[c].fillna("").astype("string")
                .str.replace(r"\.0$", "", regex=True)
                .str.strip()
            )

    return df.fillna("")
