import os
import uuid
from datetime import datetime

import pandas as pd
import pyodbc

import config

MAX_DESC_LEN = 40


def get_access_conn():
    """Open and return a pyodbc connection to the Access MDB using settings from config."""
    mdb_path = config.get_mdb_path()
    if not os.path.isabs(mdb_path):
        mdb_path = os.path.join(os.getcwd(), mdb_path)
    password = config.get_access_password()
    conn_str = (
        r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
        rf"DBQ={mdb_path};"
        rf"PWD={password};"
    )
    return pyodbc.connect(conn_str)


def load_suppliers():
    """Load all suppliers (ID, Name, Code) from the Access Suppliers table."""
    with get_access_conn() as conn:
        sql = """
        SELECT SupplierID, SupplierName, SupplierCode
        FROM Suppliers
        ORDER BY SupplierName
        """
        return pd.read_sql(sql, conn)


def new_guid():
    """Generate a new uppercase hex GUID string."""
    return uuid.uuid4().hex.upper()


def save_to_details(export_df: pd.DataFrame, supplier_id: int, mark_updated: bool = True):
    """Insert all rows from export_df into the Access Details table for the given supplier."""
    df = export_df.copy()

    df["SupplierID"] = int(supplier_id)
    df["Code"] = df.get("Supplier_Code", "").fillna("").astype(str)
    df["Description"] = df.get("TradeName", "").fillna("").astype(str).str.slice(0, MAX_DESC_LEN)
    df["Pharmacode"] = df.get("Pharmacode", "").fillna("").astype(str)
    df["Cost"] = (pd.to_numeric(df.get("Cost", 0), errors="coerce").fillna(0.0) * 100).round().astype(int)
    df["Retail"] = (pd.to_numeric(df.get("Retail", 0), errors="coerce").fillna(0.0) * 100).round().astype(int)
    df["Barcode"] = df.get("Barcode", "").fillna("").astype(str)
    if "Outers" in df.columns:
        df["Outers"] = pd.to_numeric(df["Outers"], errors="coerce").fillna(0).round().astype(int)
        df["Outers"] = df["Outers"].where(df["Outers"] > 0, 1)
    else:
        df["Outers"] = 1
    now = datetime.now()
    df["Updated"] = bool(mark_updated)
    df["LastUpdated"] = now
    df["DateCreated"] = now
    df["LastChange"] = now
    df["PartCodeGuid"] = [new_guid() for _ in range(len(df))]
    df["StockcardGuid"] = [new_guid() for _ in range(len(df))]

    insert_sql = """
        INSERT INTO Details
            ([SupplierID], [Code], [Description], [Pharmacode], [Cost], [Retail], [Barcode],
             [Outers], [Updated], [LastUpdated], [DateCreated], [LastChange],
             [PartCodeGuid], [StockcardGuid])
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """

    rows = df[["SupplierID", "Code", "Description", "Pharmacode", "Cost", "Retail", "Barcode",
               "Outers", "Updated", "LastUpdated", "DateCreated", "LastChange",
               "PartCodeGuid", "StockcardGuid"]].itertuples(index=False, name=None)

    with get_access_conn() as conn:
        cur = conn.cursor()
        cur.fast_executemany = False
        cur.executemany(insert_sql, list(rows))
        conn.commit()


def load_supplier_details(supplier_id: int) -> pd.DataFrame:
    """Load all Details records for a supplier from Access."""
    with get_access_conn() as conn:
        sql = """
        SELECT Code, Barcode, Description, Cost, Retail, Pharmacode
        FROM Details
        WHERE SupplierID = ?
        """
        return pd.read_sql(sql, conn, params=[int(supplier_id)])


def build_upsert_preview(
    export_df: pd.DataFrame,
    supplier_id: int,
    existing_df: pd.DataFrame | None = None,
) -> tuple[pd.DataFrame, dict]:
    """
    Compare export_df against existing Details records for the supplier.

    Pass existing_df to use a pre-fetched copy and skip the Access round-trip.
    Returns (preview_df, stats) where stats contains match diagnostics.
    Match order: Supplier+Code first, then Barcode fallback.
    New entries are sorted to the top.
    """
    existing = existing_df if existing_df is not None else load_supplier_details(supplier_id)

    # Build lookup dicts keyed by normalised string
    by_code: dict = {}
    by_barcode: dict = {}
    for _, row in existing.iterrows():
        code = str(row["Code"] or "").strip()
        barcode = str(row["Barcode"] or "").strip()
        if code:
            by_code[code] = row
        if barcode:
            by_barcode[barcode] = row

    df = export_df.copy()
    df["_Code"] = df.get("Supplier_Code", "").fillna("").astype(str).str.strip()
    df["_Barcode"] = df.get("Barcode", "").fillna("").astype(str).str.strip()
    df["_PLU"] = df.get("PLU", "").fillna("").astype(str).str.strip()
    df["_Pharmacode"] = df.get("Pharmacode", "").fillna("").astype(str).str.strip()
    df["_Description"] = df.get("TradeName", "").fillna("").astype(str)
    df["_Cost"] = pd.to_numeric(df.get("Cost", 0), errors="coerce").fillna(0.0)
    df["_Retail"] = pd.to_numeric(df.get("Retail", 0), errors="coerce").fillna(0.0)
    if "Outers" in df.columns:
        _Outers = pd.to_numeric(df["Outers"], errors="coerce").fillna(0).round().astype(int)
        df["_Outers"] = _Outers.where(_Outers > 0, 1)
    else:
        df["_Outers"] = 1

    rows = []
    matched_by_code = 0
    matched_by_barcode = 0

    for _, r in df.iterrows():
        code = r["_Code"]
        barcode = r["_Barcode"]
        plu = r["_PLU"]
        pharmacode = r["_Pharmacode"]
        desc = r["_Description"]
        cost = float(r["_Cost"])
        retail = float(r["_Retail"])
        Outers = int(r["_Outers"])

        existing_row = None
        # Match by supplier code first, then barcode
        if code and code in by_code:
            existing_row = by_code[code]
            matched_by_code += 1
        elif barcode and barcode in by_barcode:
            existing_row = by_barcode[barcode]
            matched_by_barcode += 1

        if existing_row is None:
            status = "NEW"
            old_desc = ""
            old_cost = None
            old_retail = None
        else:
            old_desc = str(existing_row["Description"] or "")
            old_cost = float(existing_row["Cost"] or 0) / 100
            old_retail = float(existing_row["Retail"] or 0) / 100
            old_pharmacode = str(existing_row["Pharmacode"] or "")
            # Detect whether any field actually changed
            cost_changed = round(cost, 2) != round(old_cost, 2)
            retail_changed = round(retail, 2) != round(old_retail, 2)
            desc_changed = desc.strip() != old_desc.strip()
            pharmacode_changed = pharmacode.strip() != old_pharmacode.strip()
            status = "UPDATE" if (cost_changed or retail_changed or desc_changed or pharmacode_changed) else "UNCHANGED"

        rows.append({
            "Status": status,
            "Supplier_Code": code,
            "Barcode": barcode,
            "PLU": plu,
            "Pharmacode": pharmacode,
            "Outers": Outers,
            "Old_Description": old_desc,
            "New_Description": desc,
            "Old_Cost": old_cost,
            "New_Cost": cost,
            "Old_Retail": old_retail,
            "New_Retail": retail,
        })

    preview = pd.DataFrame(rows)
    # Sort: NEW first, then UPDATE, then UNCHANGED
    preview["_sort"] = preview["Status"].map({"NEW": 0, "UPDATE": 1, "UNCHANGED": 2})
    preview = preview.sort_values("_sort").drop(columns=["_sort"]).reset_index(drop=True)

    def _sample(values):
        """Collect up to 8 unique non-empty values for diagnostic display."""
        seen = []
        for v in values:
            if v and v not in seen:
                seen.append(v)
            if len(seen) >= 8:
                break
        return seen

    stats = {
        "existing_count": len(existing),
        "access_codes": _sample(by_code.keys()),
        "access_barcodes": _sample(by_barcode.keys()),
        "export_codes": _sample(df["_Code"].tolist()),
        "export_barcodes": _sample(df["_Barcode"].tolist()),
        "matched_by_code": matched_by_code,
        "matched_by_barcode": matched_by_barcode,
        "new_count": (preview["Status"] == "NEW").sum(),
        "update_count": (preview["Status"] == "UPDATE").sum(),
        "unchanged_count": (preview["Status"] == "UNCHANGED").sum(),
    }

    return preview, stats


def _row_changed(existing_row: pd.Series, desc: str, cost: float, retail: float, pharmacode: str) -> bool:
    """Return True if any tracked field differs from the existing Access record."""
    old_desc = str(existing_row["Description"] or "").strip()
    old_cost = round(float(existing_row["Cost"] or 0) / 100, 2)
    old_retail = round(float(existing_row["Retail"] or 0) / 100, 2)
    old_pharmacode = str(existing_row["Pharmacode"] or "").strip()
    return (
        desc.strip() != old_desc
        or round(cost, 2) != old_cost
        or round(retail, 2) != old_retail
        or pharmacode.strip() != old_pharmacode
    )


def upsert_details(
    export_df: pd.DataFrame,
    supplier_id: int,
    mark_updated: bool = True,
    guid_field: str = "PartCodeGuid",
    existing_df: pd.DataFrame | None = None,
) -> tuple[int, int]:
    """
    Upsert into Access Details table.

    Fetches existing records once (or uses existing_df if pre-fetched), classifies
    each row as an update-by-code, update-by-barcode, or insert, then executes all
    three groups as batched executemany calls — reducing N round trips to 3.
    Match order: SupplierID + Code first, then SupplierID + Barcode fallback.
    """
    df = export_df.copy()
    df["SupplierID"] = int(supplier_id)
    df["Code"] = df.get("Supplier_Code", "").fillna("").astype(str)
    df["Description"] = df.get("TradeName", "").fillna("").astype(str).str.slice(0, MAX_DESC_LEN)
    df["Pharmacode"] = df.get("Pharmacode", "").fillna("").astype(str)
    df["Cost"] = (pd.to_numeric(df.get("Cost", 0), errors="coerce").fillna(0.0) * 100).round().astype(int)
    df["Retail"] = (pd.to_numeric(df.get("Retail", 0), errors="coerce").fillna(0.0) * 100).round().astype(int)
    df["Barcode"] = df.get("Barcode", "").fillna("").astype(str)
    if "Outers" in df.columns:
        df["Outers"] = pd.to_numeric(df["Outers"], errors="coerce").fillna(0).round().astype(int)
        df["Outers"] = df["Outers"].where(df["Outers"] > 0, 1)
    else:
        df["Outers"] = 1

    now = datetime.now()

    # Build lookup dicts from existing records (one fetch, or use pre-fetched copy)
    existing = existing_df if existing_df is not None else load_supplier_details(supplier_id)
    by_code: dict[str, pd.Series] = {}
    by_barcode: dict[str, pd.Series] = {}
    for _, row in existing.iterrows():
        code = str(row["Code"] or "").strip()
        barcode = str(row["Barcode"] or "").strip()
        if code:
            by_code[code] = row
        if barcode:
            by_barcode[barcode] = row

    # Classify each export row — skip rows where nothing has changed
    update_code_rows: list[tuple] = []
    update_barcode_rows: list[tuple] = []
    insert_rows: list[tuple] = []

    for r in df.itertuples(index=False):
        sup_code = (r.Code or "").strip()
        barcode = (r.Barcode or "").strip()
        desc = r.Description
        pharmacode = r.Pharmacode
        cost = float(r.Cost)
        retail = float(r.Retail)
        outers = int(r.Outers)
        upd_params = (desc, cost, retail, pharmacode, outers, bool(mark_updated), now, now, new_guid())

        if sup_code and sup_code in by_code:
            ex = by_code[sup_code]
            if _row_changed(ex, desc, cost, retail, pharmacode):
                update_code_rows.append(upd_params + (int(supplier_id), sup_code))
        elif barcode and barcode in by_barcode:
            ex = by_barcode[barcode]
            if _row_changed(ex, desc, cost, retail, pharmacode):
                update_barcode_rows.append(upd_params + (int(supplier_id), barcode))
        else:
            insert_rows.append((
                new_guid(), int(supplier_id), sup_code, desc, pharmacode,
                cost, retail, barcode, outers,
                bool(mark_updated), now, now, now, new_guid(),
            ))

    update_by_partcode_sql = """
        UPDATE Details
        SET [Description] = ?, [Cost] = ?, [Retail] = ?, [Pharmacode] = ?, [Outers] = ?,
            [Updated] = ?, [LastUpdated] = ?, [LastChange] = ?, [PartCodeGuid] = ?
        WHERE SupplierID = ? AND Code = ?
    """
    update_by_barcode_sql = """
        UPDATE Details
        SET [Description] = ?, [Cost] = ?, [Retail] = ?, [Pharmacode] = ?, [Outers] = ?,
            [Updated] = ?, [LastUpdated] = ?, [LastChange] = ?, [PartCodeGuid] = ?
        WHERE SupplierID = ? AND Barcode = ?
    """
    insert_sql = f"""
        INSERT INTO Details
        ({guid_field}, SupplierID, Code, Description, Pharmacode, Cost, Retail, Barcode,
         Outers, Updated, LastUpdated, DateCreated, LastChange, StockcardGuid)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """

    with get_access_conn() as conn:
        cur = conn.cursor()
        cur.fast_executemany = False
        if update_code_rows:
            cur.executemany(update_by_partcode_sql, update_code_rows)
        if update_barcode_rows:
            cur.executemany(update_by_barcode_sql, update_barcode_rows)
        if insert_rows:
            cur.executemany(insert_sql, insert_rows)
        cur.execute(
            "UPDATE Suppliers SET LastUpdated = ? WHERE SupplierID = ?",
            (now, int(supplier_id)),
        )
        conn.commit()

    return len(update_code_rows) + len(update_barcode_rows), len(insert_rows)


def replace_all_details(
    export_df: pd.DataFrame,
    supplier_id: int,
    mark_updated: bool = True,
) -> int:
    """
    Delete ALL existing Details records for supplier_id, then insert export_df fresh.
    Returns the number of rows inserted.
    """
    df = export_df.copy()
    df["SupplierID"] = int(supplier_id)
    df["Code"] = df.get("Supplier_Code", "").fillna("").astype(str)
    df["Description"] = df.get("TradeName", "").fillna("").astype(str).str.slice(0, MAX_DESC_LEN)
    df["Pharmacode"] = df.get("Pharmacode", "").fillna("").astype(str)
    df["Cost"] = (pd.to_numeric(df.get("Cost", 0), errors="coerce").fillna(0.0) * 100).round().astype(int)
    df["Retail"] = (pd.to_numeric(df.get("Retail", 0), errors="coerce").fillna(0.0) * 100).round().astype(int)
    df["Barcode"] = df.get("Barcode", "").fillna("").astype(str)
    if "Outers" in df.columns:
        df["Outers"] = pd.to_numeric(df["Outers"], errors="coerce").fillna(0).round().astype(int)
        df["Outers"] = df["Outers"].where(df["Outers"] > 0, 1)
    else:
        df["Outers"] = 1
    now = datetime.now()

    insert_sql = """
        INSERT INTO Details
            ([SupplierID], [Code], [Description], [Pharmacode], [Cost], [Retail],
             [Barcode], [Outers], [Updated], [LastUpdated], [DateCreated], [LastChange],
             [PartCodeGuid], [StockcardGuid])
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """

    rows = [
        (
            int(r.SupplierID), r.Code, r.Description, r.Pharmacode,
            float(r.Cost), float(r.Retail), r.Barcode, int(r.Outers),
            bool(mark_updated), now, now, now, new_guid(), new_guid(),
        )
        for r in df[["SupplierID", "Code", "Description", "Pharmacode",
                      "Cost", "Retail", "Barcode", "Outers"]].itertuples(index=False)
    ]

    with get_access_conn() as conn:
        cur = conn.cursor()
        cur.execute("DELETE FROM Details WHERE SupplierID = ?", (int(supplier_id),))
        cur.fast_executemany = False
        cur.executemany(insert_sql, rows)
        cur.execute(
            "UPDATE Suppliers SET LastUpdated = ? WHERE SupplierID = ?",
            (now, int(supplier_id)),
        )
        conn.commit()

    return len(rows)
