import os
import uuid
from datetime import datetime

import pandas as pd
import pyodbc

MAX_DESC_LEN = 40

ACCESS_PASSWORD = "LOCKIE MONDAY"
ACCESS_FILE = "suppliers.mdb"


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


def new_guid():
    return uuid.uuid4().hex.upper()


def save_to_details(export_df: pd.DataFrame, supplier_id: int, mark_updated: bool = True):
    df = export_df.copy()

    df["SupplierID"] = int(supplier_id)
    df["Code"] = df.get("Supplier_Code", "").fillna("").astype(str)
    df["Description"] = df.get("TradeName", "").fillna("").astype(str).str.slice(0, MAX_DESC_LEN)
    df["Pharmacode"] = df.get("Pharmacode", "").fillna("").astype(str)
    df["Cost"] = pd.to_numeric(df.get("Cost", 0), errors="coerce").fillna(0.0)
    df["Retail"] = pd.to_numeric(df.get("Retail", 0), errors="coerce").fillna(0.0)
    df["Barcode"] = df.get("Barcode", "").fillna("").astype(str)
    df["Updated"] = bool(mark_updated)
    df["LastUpdated"] = datetime.now()
    df["PartCodeGuid"] = new_guid()
    df["StockCodeGuid"] = new_guid()

    insert_sql = """
        INSERT INTO Details
            ([SupplierID], [Code], [Description], [Pharmacode], [Cost], [Retail], [Barcode], [Updated], [LastUpdated],[PartCodeGuid],[StockcardGuid])
        VALUES (?, ?, ?, ?, ?, ?, ?, ?,?,?,?)
    """

    rows = df[["SupplierID", "Code", "Description", "Pharmacode", "Cost", "Retail", "Barcode", "Updated", "LastUpdated", "PartCodeGuid", "StockCodeGuid"]].itertuples(index=False, name=None)

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


def build_upsert_preview(export_df: pd.DataFrame, supplier_id: int) -> pd.DataFrame:
    """
    Compare export_df against existing Details records for the supplier.

    Returns a preview DataFrame with Status (NEW/UPDATE), old and new values
    for Description, Cost and Retail only.

    Match order: Supplier+Code first, then Barcode fallback.
    New entries are sorted to the top.
    """
    existing = load_supplier_details(supplier_id)

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
    df["_Description"] = df.get("TradeName", "").fillna("").astype(str).str.slice(0, MAX_DESC_LEN)
    df["_Cost"] = pd.to_numeric(df.get("Cost", 0), errors="coerce").fillna(0.0)
    df["_Retail"] = pd.to_numeric(df.get("Retail", 0), errors="coerce").fillna(0.0)

    rows = []
    for _, r in df.iterrows():
        code = r["_Code"]
        barcode = r["_Barcode"]
        desc = r["_Description"]
        cost = float(r["_Cost"])
        retail = float(r["_Retail"])

        existing_row = None
        # Match by supplier code first, then barcode
        if code and code in by_code:
            existing_row = by_code[code]
        elif barcode and barcode in by_barcode:
            existing_row = by_barcode[barcode]

        if existing_row is None:
            status = "NEW"
            old_desc = ""
            old_cost = None
            old_retail = None
        else:
            old_desc = str(existing_row["Description"] or "")
            old_cost = float(existing_row["Cost"] or 0)
            old_retail = float(existing_row["Retail"] or 0)
            status = "UPDATE"

        rows.append({
            "Status": status,
            "Supplier_Code": code,
            "Barcode": barcode,
            "Old_Description": old_desc,
            "New_Description": desc,
            "Old_Cost": old_cost,
            "New_Cost": cost,
            "Old_Retail": old_retail,
            "New_Retail": retail,
        })

    preview = pd.DataFrame(rows)
    # NEW entries first
    preview["_sort"] = preview["Status"].map({"NEW": 0, "UPDATE": 1})
    preview = preview.sort_values("_sort").drop(columns=["_sort"]).reset_index(drop=True)
    return preview


def upsert_details(
    export_df: pd.DataFrame,
    supplier_id: int,
    mark_updated: bool = True,
    guid_field: str = "PartCodeGuid",
) -> tuple[int, int]:
    """
    Upsert into Access Details table.

    Match order: SupplierID + Code first, then SupplierID + Barcode fallback.
    Updates Description, Cost, Retail, Pharmacode and updated timestamp if record exists.
    Inserts new record if no match found.
    """
    df = export_df.copy()

    df["SupplierID"] = int(supplier_id)
    df["Code"] = df.get("Supplier_Code", "").fillna("").astype(str)
    df["Description"] = df.get("TradeName", "").fillna("").astype(str).str.slice(0, MAX_DESC_LEN)
    df["Pharmacode"] = df.get("Pharmacode", "").fillna("").astype(str)
    df["Cost"] = pd.to_numeric(df.get("Cost", 0), errors="coerce").fillna(0.0)
    df["Retail"] = pd.to_numeric(df.get("Retail", 0), errors="coerce").fillna(0.0)
    df["Barcode"] = df.get("Barcode", "").fillna("").astype(str)

    now = datetime.now()
    updated_count = 0
    inserted_count = 0

    update_by_partcode_sql = """
        UPDATE Details
        SET [Description] = ?, [Cost] = ?, [Retail] = ?, [Pharmacode] = ?, [Updated] = ?, [LastUpdated] = ?, [PartCodeGuid] = ?
        WHERE SupplierID = ? AND Code = ?
    """

    update_by_barcode_sql = """
        UPDATE Details
        SET [Description] = ?, [Cost] = ?, [Retail] = ?, [Pharmacode] = ?, [Updated] = ?, [LastUpdated] = ?, [PartCodeGuid] = ?
        WHERE SupplierID = ? AND Barcode = ?
    """

    insert_sql = f"""
        INSERT INTO Details
        ({guid_field}, SupplierID, Code, Description, Pharmacode, Cost, Retail, Barcode, Updated, LastUpdated)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """

    with get_access_conn() as conn:
        cur = conn.cursor()

        for r in df.itertuples(index=False):
            supplier = r.SupplierID
            sup_code = (r.Code or "").strip()
            barcode = (r.Barcode or "").strip()
            desc = r.Description
            pharmacode = r.Pharmacode
            cost = float(r.Cost)
            retail = float(r.Retail)

            did_update = False

            # Match by supplier code first
            if sup_code:
                cur.execute(
                    update_by_partcode_sql,
                    (desc, cost, retail, pharmacode, bool(mark_updated), now, new_guid(), supplier, sup_code),
                )
                if cur.rowcount > 0:
                    did_update = True
                    updated_count += 1

            # Fallback: match by barcode
            if not did_update and barcode:
                cur.execute(
                    update_by_barcode_sql,
                    (desc, cost, retail, pharmacode, bool(mark_updated), now, new_guid(), supplier, barcode),
                )
                if cur.rowcount > 0:
                    did_update = True
                    updated_count += 1

            # Insert if no match
            if not did_update:
                cur.execute(
                    insert_sql,
                    (new_guid(), supplier, sup_code, desc, pharmacode, cost, retail, barcode, bool(mark_updated), now),
                )
                inserted_count += 1

        conn.commit()

    return updated_count, inserted_count
