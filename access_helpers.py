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


def upsert_details(
    export_df: pd.DataFrame,
    supplier_id: int,
    mark_updated: bool = True,
    guid_field: str = "PartCodeGuid",
) -> tuple[int, int]:
    """
    Upsert into Access Details table.

    Match the following to check for existing record:
      - SupplierID + Barcode
      - SupplierID + Code (fallback)

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

    update_by_barcode_sql = """
        UPDATE Details
        SET [Description] = ?, [Cost] = ?, [Retail] = ?, [Pharmacode] = ?, [Updated] = ?, [LastUpdated] = ?, [PartCodeGuid] = ?
        WHERE SupplierID = ? AND Barcode = ?
    """

    update_by_partcode_sql = """
        UPDATE Details
        SET [Description] = ?, [Cost] = ?, [Retail] = ?, [Pharmacode] = ?, [Updated] = ?, [LastUpdated] = ?, [PartCodeGuid] = ?
        WHERE SupplierID = ? AND Code = ?
    """

    insert_sql = f"""
        INSERT INTO Details
        ({guid_field}, SupplierID, Code, Description, Pharmacode, Cost, Retail, Barcode, Updated, LastUpdated)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """

    with get_access_conn() as conn:
        cur = conn.cursor()

        for r in df.itertuples(index=False):  # fixed: was df.intertuples (typo)
            supplier = r.SupplierID
            sup_code = (r.Code or "").strip()
            barcode = (r.Barcode or "").strip()
            desc = r.Description
            pharmacode = r.Pharmacode
            cost = float(r.Cost)
            retail = float(r.Retail)

            did_update = False

            # Match by barcode first
            if barcode:
                cur.execute(
                    update_by_barcode_sql,
                    (desc, cost, retail, pharmacode, bool(mark_updated), now, new_guid(), supplier, barcode),
                )
                if cur.rowcount > 0:
                    did_update = True
                    updated_count += 1

            # Fallback: match by supplier code
            if not did_update and sup_code:
                cur.execute(
                    update_by_partcode_sql,
                    (desc, cost, retail, pharmacode, bool(mark_updated), now, new_guid(), supplier, sup_code),
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
