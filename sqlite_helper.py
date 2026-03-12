"""sqlite_helper.py — SQLite equivalent of access_helpers.py.

Same DB schema as the Access Suppliers.mdb, using Python's built-in sqlite3.
Database file: suppliers.sqlite (created automatically on first use).
"""
from __future__ import annotations

import os
import sqlite3
import uuid
from datetime import datetime

import pandas as pd

import config

MAX_DESC_LEN = 40


# ── Connection ────────────────────────────────────────────────────────────────

def get_sqlite_conn() -> sqlite3.Connection:
    db_path = config.get_sqlite_path()
    if not os.path.isabs(db_path):
        db_path = os.path.join(os.getcwd(), db_path)
    conn = sqlite3.connect(db_path, detect_types=sqlite3.PARSE_DECLTYPES)
    conn.row_factory = sqlite3.Row
    return conn


# ── Schema initialisation ─────────────────────────────────────────────────────

def init_db():
    """Create tables if they don't exist yet. Safe to call on every startup."""
    with get_sqlite_conn() as conn:
        conn.executescript("""
            CREATE TABLE IF NOT EXISTS Suppliers (
                SupplierID   INTEGER PRIMARY KEY AUTOINCREMENT,
                SupplierName TEXT    NOT NULL,
                SupplierCode TEXT    DEFAULT '',
                LastUpdated  TEXT
            );

            CREATE TABLE IF NOT EXISTS Details (
                PartCodeGuid  TEXT PRIMARY KEY,
                SupplierID    INTEGER NOT NULL,
                Code          TEXT    DEFAULT '',
                Description   TEXT    DEFAULT '',
                Pharmacode    TEXT    DEFAULT '',
                Cost          INTEGER DEFAULT 0,
                Retail        INTEGER DEFAULT 0,
                Barcode       TEXT    DEFAULT '',
                Outers        INTEGER DEFAULT 1,
                Updated       INTEGER DEFAULT 0,
                LastUpdated   TEXT,
                DateCreated   TEXT,
                LastChange    TEXT,
                StockcardGuid TEXT    DEFAULT '',
                FOREIGN KEY (SupplierID) REFERENCES Suppliers(SupplierID)
            );

            CREATE INDEX IF NOT EXISTS idx_details_supplier
                ON Details (SupplierID);
            CREATE INDEX IF NOT EXISTS idx_details_code
                ON Details (SupplierID, Code);
            CREATE INDEX IF NOT EXISTS idx_details_barcode
                ON Details (SupplierID, Barcode);
        """)


# ── Helpers ───────────────────────────────────────────────────────────────────

def new_guid() -> str:
    return uuid.uuid4().hex.upper()


def _dt(dt: datetime) -> str:
    """SQLite stores datetimes as ISO text."""
    return dt.isoformat(sep=" ", timespec="seconds")


# ── Supplier queries ──────────────────────────────────────────────────────────

def load_suppliers() -> pd.DataFrame:
    init_db()
    with get_sqlite_conn() as conn:
        return pd.read_sql(
            "SELECT SupplierID, SupplierName, SupplierCode FROM Suppliers ORDER BY SupplierName",
            conn,
        )


def ensure_supplier(supplier_name: str, supplier_code: str = "") -> int:
    """Return existing SupplierID or insert a new Supplier row."""
    init_db()
    with get_sqlite_conn() as conn:
        row = conn.execute(
            "SELECT SupplierID FROM Suppliers WHERE SupplierName = ?",
            (supplier_name,),
        ).fetchone()
        if row:
            return int(row["SupplierID"])
        cur = conn.execute(
            "INSERT INTO Suppliers (SupplierName, SupplierCode, LastUpdated) VALUES (?, ?, ?)",
            (supplier_name, supplier_code, _dt(datetime.now())),
        )
        conn.commit()
        return cur.lastrowid


# ── Data prep (shared by all write functions) ─────────────────────────────────

def _prepare_df(export_df: pd.DataFrame, supplier_id: int, mark_updated: bool) -> pd.DataFrame:
    df = export_df.copy()
    df["SupplierID"]   = int(supplier_id)
    df["Code"]         = df.get("Supplier_Code", "").fillna("").astype(str)
    df["Description"]  = df.get("TradeName", "").fillna("").astype(str).str.slice(0, MAX_DESC_LEN)
    df["Pharmacode"]   = df.get("Pharmacode", "").fillna("").astype(str)
    df["Cost"]         = (pd.to_numeric(df.get("Cost", 0), errors="coerce").fillna(0.0) * 100).round().astype(int)
    df["Retail"]       = (pd.to_numeric(df.get("Retail", 0), errors="coerce").fillna(0.0) * 100).round().astype(int)
    df["Barcode"]      = df.get("Barcode", "").fillna("").astype(str)
    if "Outers" in df.columns:
        outers = pd.to_numeric(df["Outers"], errors="coerce").fillna(0).round().astype(int)
        df["Outers"] = outers.where(outers > 0, 1)
    else:
        df["Outers"] = 1
    df["Updated"] = 1 if mark_updated else 0
    return df


# ── Write operations ──────────────────────────────────────────────────────────

def save_to_details(export_df: pd.DataFrame, supplier_id: int, mark_updated: bool = True):
    """Insert all rows from export_df as new Detail records (no duplicate check)."""
    init_db()
    df = _prepare_df(export_df, supplier_id, mark_updated)
    now = _dt(datetime.now())

    rows = [
        (
            new_guid(),
            int(r.SupplierID), r.Code, r.Description, r.Pharmacode,
            int(r.Cost), int(r.Retail), r.Barcode, int(r.Outers),
            int(r.Updated), now, now, now, new_guid(),
        )
        for r in df[["SupplierID", "Code", "Description", "Pharmacode",
                      "Cost", "Retail", "Barcode", "Outers", "Updated"]].itertuples(index=False)
    ]

    sql = """
        INSERT INTO Details
            (PartCodeGuid, SupplierID, Code, Description, Pharmacode,
             Cost, Retail, Barcode, Outers, Updated, LastUpdated,
             DateCreated, LastChange, StockcardGuid)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """
    with get_sqlite_conn() as conn:
        conn.executemany(sql, rows)
        conn.commit()


def upsert_details(
    export_df: pd.DataFrame,
    supplier_id: int,
    mark_updated: bool = True,
) -> tuple[int, int]:
    """
    Upsert into SQLite Details table.

    Match order: SupplierID + Code first, then SupplierID + Barcode fallback.
    Updates existing record; inserts new record if no match found.
    Returns (updated_count, inserted_count).
    """
    init_db()
    df = _prepare_df(export_df, supplier_id, mark_updated)
    now = _dt(datetime.now())
    sid = int(supplier_id)

    update_by_code_sql = """
        UPDATE Details
        SET Description = ?, Cost = ?, Retail = ?, Pharmacode = ?, Outers = ?,
            Updated = ?, LastUpdated = ?, LastChange = ?, PartCodeGuid = ?
        WHERE SupplierID = ? AND Code = ? AND Code != ''
    """
    update_by_barcode_sql = """
        UPDATE Details
        SET Description = ?, Cost = ?, Retail = ?, Pharmacode = ?, Outers = ?,
            Updated = ?, LastUpdated = ?, LastChange = ?, PartCodeGuid = ?
        WHERE SupplierID = ? AND Barcode = ? AND Barcode != ''
    """
    insert_sql = """
        INSERT INTO Details
            (PartCodeGuid, SupplierID, Code, Description, Pharmacode,
             Cost, Retail, Barcode, Outers, Updated, LastUpdated,
             DateCreated, LastChange, StockcardGuid)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """

    updated_count = 0
    inserted_count = 0

    with get_sqlite_conn() as conn:
        cur = conn.cursor()
        for r in df.itertuples(index=False):
            sup_code = (str(r.Code) or "").strip()
            barcode  = (str(r.Barcode) or "").strip()
            desc     = r.Description
            pharma   = r.Pharmacode
            cost     = int(r.Cost)
            retail   = int(r.Retail)
            outers   = int(r.Outers)
            updated  = int(r.Updated)

            did_update = False

            if sup_code:
                cur.execute(
                    update_by_code_sql,
                    (desc, cost, retail, pharma, outers, updated, now, now, new_guid(), sid, sup_code),
                )
                if cur.rowcount > 0:
                    did_update = True
                    updated_count += 1

            if not did_update and barcode:
                cur.execute(
                    update_by_barcode_sql,
                    (desc, cost, retail, pharma, outers, updated, now, now, new_guid(), sid, barcode),
                )
                if cur.rowcount > 0:
                    did_update = True
                    updated_count += 1

            if not did_update:
                cur.execute(
                    insert_sql,
                    (new_guid(), sid, sup_code, desc, pharma, cost, retail, barcode,
                     outers, updated, now, now, now, new_guid()),
                )
                inserted_count += 1

        conn.commit()

    return updated_count, inserted_count


def replace_all_details(
    export_df: pd.DataFrame,
    supplier_id: int,
    mark_updated: bool = True,
) -> int:
    """Delete ALL existing Details rows for supplier_id, then insert export_df fresh.
    Returns the number of rows inserted."""
    init_db()
    df = _prepare_df(export_df, supplier_id, mark_updated)
    now = _dt(datetime.now())
    sid = int(supplier_id)

    insert_sql = """
        INSERT INTO Details
            (PartCodeGuid, SupplierID, Code, Description, Pharmacode,
             Cost, Retail, Barcode, Outers, Updated, LastUpdated,
             DateCreated, LastChange, StockcardGuid)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """
    rows = [
        (
            new_guid(),
            sid, r.Code, r.Description, r.Pharmacode,
            int(r.Cost), int(r.Retail), r.Barcode, int(r.Outers),
            int(r.Updated), now, now, now, new_guid(),
        )
        for r in df[["Code", "Description", "Pharmacode", "Cost", "Retail",
                      "Barcode", "Outers", "Updated"]].itertuples(index=False)
    ]

    with get_sqlite_conn() as conn:
        conn.execute("DELETE FROM Details WHERE SupplierID = ?", (sid,))
        conn.executemany(insert_sql, rows)
        conn.commit()

    return len(rows)


# ── Preview (mirrors access_helpers.build_upsert_preview) ────────────────────

def load_supplier_details(supplier_id: int) -> pd.DataFrame:
    """Load all Details records for a supplier from SQLite."""
    init_db()
    with get_sqlite_conn() as conn:
        return pd.read_sql(
            "SELECT Code, Barcode, Description, Cost, Retail, Pharmacode "
            "FROM Details WHERE SupplierID = ?",
            conn,
            params=(int(supplier_id),),
        )


def build_upsert_preview(
    export_df: pd.DataFrame,
    supplier_id: int,
) -> tuple[pd.DataFrame, dict]:
    """
    Compare export_df against existing Details for supplier_id.
    Returns (preview_df, stats).  New entries sorted to the top.
    """
    existing = load_supplier_details(supplier_id)

    by_code: dict = {}
    by_barcode: dict = {}
    for _, row in existing.iterrows():
        code    = str(row["Code"]    or "").strip()
        barcode = str(row["Barcode"] or "").strip()
        if code:
            by_code[code] = row
        if barcode:
            by_barcode[barcode] = row

    df = export_df.copy()
    df["_Code"]        = df.get("Supplier_Code", "").fillna("").astype(str).str.strip()
    df["_Barcode"]     = df.get("Barcode", "").fillna("").astype(str).str.strip()
    df["_PLU"]         = df.get("PLU", "").fillna("").astype(str).str.strip()
    df["_Pharmacode"]  = df.get("Pharmacode", "").fillna("").astype(str).str.strip()
    df["_Description"] = df.get("TradeName", "").fillna("").astype(str)
    df["_Cost"]        = pd.to_numeric(df.get("Cost", 0), errors="coerce").fillna(0.0)
    df["_Retail"]      = pd.to_numeric(df.get("Retail", 0), errors="coerce").fillna(0.0)
    if "Outers" in df.columns:
        _outers = pd.to_numeric(df["Outers"], errors="coerce").fillna(0).round().astype(int)
        df["_Outers"] = _outers.where(_outers > 0, 1)
    else:
        df["_Outers"] = 1

    rows = []
    matched_by_code    = 0
    matched_by_barcode = 0

    for _, r in df.iterrows():
        code     = r["_Code"]
        barcode  = r["_Barcode"]
        plu      = r["_PLU"]
        pharma   = r["_Pharmacode"]
        desc     = r["_Description"]
        cost     = float(r["_Cost"])
        retail   = float(r["_Retail"])
        outers   = int(r["_Outers"])

        existing_row = None
        if code and code in by_code:
            existing_row = by_code[code]
            matched_by_code += 1
        elif barcode and barcode in by_barcode:
            existing_row = by_barcode[barcode]
            matched_by_barcode += 1

        if existing_row is None:
            status    = "NEW"
            old_desc  = ""
            old_cost  = None
            old_retail = None
        else:
            old_desc   = str(existing_row["Description"] or "")
            old_cost   = float(existing_row["Cost"]   or 0) / 100
            old_retail = float(existing_row["Retail"] or 0) / 100
            status     = "UPDATE"

        rows.append({
            "Status":          status,
            "Supplier_Code":   code,
            "Barcode":         barcode,
            "PLU":             plu,
            "Pharmacode":      pharma,
            "Outers":          outers,
            "Old_Description": old_desc,
            "New_Description": desc,
            "Old_Cost":        old_cost,
            "New_Cost":        cost,
            "Old_Retail":      old_retail,
            "New_Retail":      retail,
        })

    preview = pd.DataFrame(rows)
    preview["_sort"] = preview["Status"].map({"NEW": 0, "UPDATE": 1})
    preview = preview.sort_values("_sort").drop(columns=["_sort"]).reset_index(drop=True)

    def _sample(values):
        seen = []
        for v in values:
            if v and v not in seen:
                seen.append(v)
            if len(seen) >= 8:
                break
        return seen

    stats = {
        "existing_count":    len(existing),
        "sqlite_codes":      _sample(list(by_code.keys())),
        "sqlite_barcodes":   _sample(list(by_barcode.keys())),
        "export_codes":      _sample(df["_Code"].tolist()),
        "export_barcodes":   _sample(df["_Barcode"].tolist()),
        "matched_by_code":   matched_by_code,
        "matched_by_barcode": matched_by_barcode,
        "new_count":         len(preview) - matched_by_code - matched_by_barcode,
    }

    return preview, stats


# ── Browser helpers ────────────────────────────────────────────────────────────

# Columns shown in the browser grid (Cost/Retail converted to dollars for display)
BROWSER_DISPLAY_COLS = ["Code", "Barcode", "Description", "Pharmacode",
                        "Cost", "Retail", "Outers", "Updated"]
BROWSER_EDITABLE_COLS = ["Description", "Cost", "Retail", "Pharmacode", "Outers"]


def load_details_for_browser(supplier_id: int) -> pd.DataFrame:
    """Load Details for supplier and convert Cost/Retail to dollars for display."""
    init_db()
    with get_sqlite_conn() as conn:
        cur = conn.cursor()
        cur.execute(
            "SELECT Code, Barcode, Description, Pharmacode, Cost, Retail, Outers, Updated "
            "FROM Details WHERE SupplierID = ? ORDER BY Description",
            (int(supplier_id),),
        )
        cols = [d[0] for d in cur.description]
        rows = cur.fetchall()
    df = pd.DataFrame([list(r) for r in rows], columns=cols)
    df["Cost"] = (pd.to_numeric(df["Cost"], errors="coerce").fillna(0) / 100).round(2)
    df["Retail"] = (pd.to_numeric(df["Retail"], errors="coerce").fillna(0) / 100).round(2)
    return df


def save_details_from_browser(supplier_id: int, df: pd.DataFrame) -> int:
    """Replace all Details for supplier with the edited browser DataFrame.

    Cost/Retail are expected as dollars; converts back to integer cents for storage.
    Returns number of rows saved.
    """
    init_db()
    now = _dt(datetime.now())
    sid = int(supplier_id)
    df = df.copy()
    df["Cost"] = (pd.to_numeric(df["Cost"], errors="coerce").fillna(0) * 100).round().astype(int)
    df["Retail"] = (pd.to_numeric(df["Retail"], errors="coerce").fillna(0) * 100).round().astype(int)
    df["Outers"] = pd.to_numeric(df["Outers"], errors="coerce").fillna(1).round().astype(int)
    df["Outers"] = df["Outers"].where(df["Outers"] > 0, 1)
    df["Description"] = df["Description"].fillna("").astype(str).str.strip().str.slice(0, 40)
    df["Pharmacode"] = df["Pharmacode"].fillna("").astype(str)
    df["Updated"] = pd.to_numeric(df.get("Updated", 1), errors="coerce").fillna(1).astype(int)

    rows = [
        (
            new_guid(), sid,
            str(r.Code or ""), str(r.Description or ""), str(r.Pharmacode or ""),
            int(r.Cost), int(r.Retail), str(r.Barcode or ""),
            int(r.Outers), int(r.Updated), now, now, now, new_guid(),
        )
        for r in df.itertuples(index=False)
    ]
    sql = """
        INSERT INTO Details
            (PartCodeGuid, SupplierID, Code, Description, Pharmacode,
             Cost, Retail, Barcode, Outers, Updated, LastUpdated,
             DateCreated, LastChange, StockcardGuid)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """
    with get_sqlite_conn() as conn:
        conn.execute("DELETE FROM Details WHERE SupplierID = ?", (sid,))
        conn.executemany(sql, rows)
        conn.commit()
    return len(rows)


def get_table_names() -> list[str]:
    """Return all user table names in the SQLite database."""
    init_db()
    with get_sqlite_conn() as conn:
        cur = conn.execute(
            "SELECT name FROM sqlite_master WHERE type='table' ORDER BY name"
        )
        return [r[0] for r in cur.fetchall()]


def load_full_table(table_name: str) -> pd.DataFrame:
    """Load an entire table as a DataFrame (for non-Details tables in the browser)."""
    init_db()
    with get_sqlite_conn() as conn:
        cur = conn.cursor()
        cur.execute(f"SELECT * FROM [{table_name}]")  # noqa: S608 — internal use only
        cols = [d[0] for d in cur.description]
        return pd.DataFrame([list(r) for r in cur.fetchall()], columns=cols)


def delete_details_by_keys(supplier_id: int, codes: list[str]) -> int:
    """Delete Details rows for supplier where Code OR Barcode is in the given list.

    Returns the number of rows deleted.
    """
    init_db()
    if not codes:
        return 0
    placeholders = ",".join("?" * len(codes))
    sid = int(supplier_id)
    with get_sqlite_conn() as conn:
        cur = conn.execute(
            f"DELETE FROM Details WHERE SupplierID = ? "
            f"AND (Code IN ({placeholders}) OR Barcode IN ({placeholders}))",
            [sid] + codes + codes,
        )
        conn.commit()
        return cur.rowcount
