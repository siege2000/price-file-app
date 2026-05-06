import os
import uuid
from datetime import datetime

import pandas as pd
import pyodbc

import config

MAX_DESC_LEN = 40


def _first_barcode(series: "pd.Series") -> "pd.Series":
    """Return only the first barcode from a potentially comma-separated value.

    The Access Details table stores a single barcode per row (matching legacy
    VB6 behaviour).  Incoming price files sometimes carry multiple barcodes as
    a comma-separated string; we keep only the first to stay compatible.
    """
    return series.str.split(",").str[0].str.strip()


def _access_conn(mdb_path: str) -> "pyodbc.Connection":
    if not os.path.isabs(mdb_path):
        mdb_path = os.path.join(os.getcwd(), mdb_path)
    password = config.get_access_password()
    conn_str = (
        r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
        rf"DBQ={mdb_path};"
        rf"PWD={password};"
    )
    return pyodbc.connect(conn_str)


def get_access_conn():
    """Open and return a pyodbc connection to Suppliers.mdb."""
    return _access_conn(config.get_mdb_path())


def get_settings1_conn():
    """Open and return a pyodbc connection to Settings1.mdb."""
    return _access_conn(config.get_settings1_mdb_path())


def load_find_replace_rules() -> pd.DataFrame:
    """Load all rows from FindAndReplace in Settings1.mdb.

    Returns a DataFrame with columns: SupplierID, Find, Replace.
    """
    with get_settings1_conn() as conn:
        return pd.read_sql(
            "SELECT SupplierID, Find, Replace FROM FindAndReplace ORDER BY SupplierID",
            conn,
        )


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
    df["Barcode"] = _first_barcode(df.get("Barcode", "").fillna("").astype(str))
    if "Outers" in df.columns:
        df["Outers"] = pd.to_numeric(df["Outers"], errors="coerce").fillna(0).round().astype(int)
        df["Outers"] = df["Outers"].where(df["Outers"] > 0, 1)
    else:
        df["Outers"] = 1
    now = datetime.now().date()
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
    """Load all Details records for a supplier from Access.

    Uses cursor.fetchall() rather than pd.read_sql to avoid pyodbc row-by-row
    overhead, which is significantly faster over a network drive.
    """
    sql = """
        SELECT Code, Barcode, Description, Cost, Retail, Pharmacode
        FROM Details
        WHERE SupplierID = ?
    """
    with get_access_conn() as conn:
        cur = conn.cursor()
        cur.execute(sql, [int(supplier_id)])
        cols = [d[0] for d in cur.description]
        rows = cur.fetchall()
    return pd.DataFrame([list(r) for r in rows], columns=cols)


def build_upsert_preview(
    export_df: pd.DataFrame,
    supplier_id: int,
    existing_df: pd.DataFrame | None = None,
) -> tuple[pd.DataFrame, dict]:
    """
    Compare export_df against existing Details records for the supplier.

    Pass existing_df to use a pre-fetched copy and skip the Access round-trip.
    Uses vectorized pandas merge (no iterrows) for fast comparison on large files.
    Match order: Supplier+Code first, then Barcode fallback.
    Returns (preview_df, stats). Sort order: NEW → UPDATE → UNCHANGED.
    """
    existing = (existing_df if existing_df is not None else load_supplier_details(supplier_id)).copy()

    # Normalise existing lookup columns
    existing["_ecode"] = existing["Code"].fillna("").astype(str).str.strip()
    existing["_ebarcode"] = existing["Barcode"].fillna("").astype(str).str.strip()
    existing["_edesc"] = existing["Description"].fillna("").astype(str).str.strip()
    existing["_ecost"] = (pd.to_numeric(existing["Cost"], errors="coerce").fillna(0) / 100).round(2)
    existing["_eretail"] = (pd.to_numeric(existing["Retail"], errors="coerce").fillna(0) / 100).round(2)
    existing["_epharm"] = existing["Pharmacode"].fillna("").astype(str).str.strip()

    ex_cols = ["_edesc", "_ecost", "_eretail", "_epharm"]

    # Build deduplicated lookup tables indexed by code / barcode
    ex_by_code = (
        existing[existing["_ecode"] != ""]
        .drop_duplicates("_ecode", keep="last")
        .set_index("_ecode")[ex_cols]
    )
    # Explode any comma-separated barcodes stored in Access so either individual
    # barcode can be found in the index.
    _ex_bc = existing[existing["_ebarcode"] != ""].copy()
    _ex_bc = _ex_bc.assign(
        _ebarcode=_ex_bc["_ebarcode"].str.split(r"\s*,\s*")
    ).explode("_ebarcode")
    _ex_bc["_ebarcode"] = _ex_bc["_ebarcode"].str.strip()
    ex_by_barcode = (
        _ex_bc[_ex_bc["_ebarcode"] != ""]
        .drop_duplicates("_ebarcode", keep="last")
        .set_index("_ebarcode")[ex_cols]
    )
    # Pharmacode fallback — only index pharmacodes that are unique for this supplier.
    # _epharm is the index so it must be excluded from the column list; it is
    # restored as _epharm = _pharm after the merge (they are equal by definition).
    _ex_ph = existing[existing["_epharm"] != ""].copy()
    _ph_counts = _ex_ph["_epharm"].value_counts()
    _ex_pharm_cols = ["_edesc", "_ecost", "_eretail"]  # _epharm is the index, not a column
    ex_by_pharm = (
        _ex_ph[_ex_ph["_epharm"].isin(_ph_counts[_ph_counts == 1].index)]
        .drop_duplicates("_epharm", keep="last")
        .set_index("_epharm")[_ex_pharm_cols]
    )

    # Normalise export columns
    df = export_df.copy().reset_index(drop=True)
    df["_code"] = df.get("Supplier_Code", "").fillna("").astype(str).str.strip()
    df["_barcode"] = df.get("Barcode", "").fillna("").astype(str).str.strip()
    df["_plu"] = df.get("PLU", "").fillna("").astype(str).str.strip()
    df["_pharm"] = df.get("Pharmacode", "").fillna("").astype(str).str.strip()
    df["_desc"] = df.get("TradeName", "").fillna("").astype(str).str.strip()
    df["_cost"] = pd.to_numeric(df.get("Cost", 0), errors="coerce").fillna(0.0).round(2)
    df["_retail"] = pd.to_numeric(df.get("Retail", 0), errors="coerce").fillna(0.0).round(2)
    if "Outers" in df.columns:
        _o = pd.to_numeric(df["Outers"], errors="coerce").fillna(0).round().astype(int)
        df["_outers"] = _o.where(_o > 0, 1)
    else:
        df["_outers"] = 1

    # --- Step 1: match by code ---
    has_code = df["_code"] != ""
    matched_code = df[has_code].merge(
        ex_by_code, left_on="_code", right_index=True, how="left"
    )
    # Only rows with no supplier code fall through to barcode matching.
    # Rows that have a code (whether found in Access or not) are fully handled
    # by the code path above — including new items whose code isn't in Access yet
    # (those appear in code_new below). Without this restriction, new items that
    # have both a supplier code AND a barcode would appear twice in the preview:
    # once via code_new and once via the barcode/unmatched path.
    no_code_match_mask = ~has_code
    df_remaining = df[no_code_match_mask].copy()

    # --- Step 2: match remaining by barcode ---
    # Barcodes may be comma-separated in the price file; try each candidate.
    has_barcode = df_remaining["_barcode"] != ""
    barcode_candidates = df_remaining[has_barcode].copy().reset_index(drop=True)
    barcode_candidates["_orig_idx"] = barcode_candidates.index

    exploded = barcode_candidates.assign(
        _barcode_try=barcode_candidates["_barcode"].str.split(r"\s*,\s*")
    ).explode("_barcode_try")
    exploded["_barcode_try"] = exploded["_barcode_try"].str.strip()

    merged = exploded.merge(
        ex_by_barcode, left_on="_barcode_try", right_index=True, how="left"
    )
    first_hits = (
        merged[merged["_edesc"].notna()]
        .drop_duplicates(subset="_orig_idx", keep="first")
    )
    matched_idxs = set(first_hits["_orig_idx"])
    unmatched_barcode = barcode_candidates[~barcode_candidates["_orig_idx"].isin(matched_idxs)]

    matched_barcode = first_hits.drop(columns=["_barcode_try", "_orig_idx"])
    after_barcode = pd.concat([
        df_remaining[~has_barcode],
        unmatched_barcode.drop(columns=["_orig_idx"]),
    ], ignore_index=True)

    # --- Step 3: match remaining by pharmacode ---
    # Fallback when barcodes are absent or truncated in Access.
    has_pharm = after_barcode["_pharm"] != ""
    matched_pharm = after_barcode[has_pharm].merge(
        ex_by_pharm, left_on="_pharm", right_index=True, how="left"
    )
    unmatched = pd.concat([
        after_barcode[~has_pharm],
        matched_pharm[matched_pharm["_edesc"].isna()],
    ], ignore_index=True)
    matched_pharm = matched_pharm[matched_pharm["_edesc"].notna()].copy()
    # Restore _epharm — since we matched on pharmacode it equals _pharm
    matched_pharm["_epharm"] = matched_pharm["_pharm"]

    # --- Build preview rows ---
    def _make_rows(matched: pd.DataFrame, match_type: str) -> pd.DataFrame:
        """Classify matched rows as UPDATE or UNCHANGED."""
        m = matched.copy()
        changed = (
            (m["_desc"] != m["_edesc"]) |
            (m["_cost"] != m["_ecost"]) |
            (m["_retail"] != m["_eretail"]) |
            (m["_pharm"] != m["_epharm"])
        )
        m["Status"] = changed.map({True: "UPDATE", False: "UNCHANGED"})
        m["Old_Description"] = m["_edesc"].fillna("")
        m["Old_Cost"] = m["_ecost"]
        m["Old_Retail"] = m["_eretail"]
        return m

    code_rows = _make_rows(matched_code[matched_code["_edesc"].notna()], "code")
    # rows matched by code but not found (code present but not in Access = NEW)
    code_new = matched_code[matched_code["_edesc"].isna()].copy()
    code_new["Status"] = "NEW"
    code_new["Old_Description"] = ""
    code_new["Old_Cost"] = None
    code_new["Old_Retail"] = None

    barcode_rows = _make_rows(matched_barcode, "barcode")
    pharm_rows = _make_rows(matched_pharm, "pharm")

    unmatched["Status"] = "NEW"
    unmatched["Old_Description"] = ""
    unmatched["Old_Cost"] = None
    unmatched["Old_Retail"] = None

    all_rows = pd.concat([code_rows, code_new, barcode_rows, pharm_rows, unmatched], ignore_index=True)

    preview = pd.DataFrame({
        "Status": all_rows["Status"],
        "Supplier_Code": all_rows["_code"],
        "Barcode": all_rows["_barcode"],
        "PLU": all_rows["_plu"],
        "Pharmacode": all_rows["_pharm"],
        "Outers": all_rows["_outers"],
        "Old_Description": all_rows["Old_Description"],
        "New_Description": all_rows["_desc"],
        "Old_Cost": all_rows["Old_Cost"],
        "New_Cost": all_rows["_cost"],
        "Old_Retail": all_rows["Old_Retail"],
        "New_Retail": all_rows["_retail"],
    })

    # Sort: NEW → UPDATE → UNCHANGED
    preview["_sort"] = preview["Status"].map({"NEW": 0, "UPDATE": 1, "UNCHANGED": 2})
    preview = preview.sort_values("_sort").drop(columns=["_sort"]).reset_index(drop=True)

    matched_by_code = len(code_rows)
    matched_by_barcode = len(barcode_rows)

    def _sample(series: pd.Series) -> list:
        """Collect up to 8 unique non-empty values for diagnostic display."""
        seen: list = []
        for v in series:
            v = str(v).strip()
            if v and v not in seen:
                seen.append(v)
            if len(seen) >= 8:
                break
        return seen

    stats = {
        "existing_count": len(existing),
        "access_codes": _sample(ex_by_code.index.to_series()),
        "access_barcodes": _sample(ex_by_barcode.index.to_series()),
        "export_codes": _sample(df["_code"]),
        "export_barcodes": _sample(df["_barcode"]),
        "matched_by_code": matched_by_code,
        "matched_by_barcode": matched_by_barcode,
        "new_count": int((preview["Status"] == "NEW").sum()),
        "update_count": int((preview["Status"] == "UPDATE").sum()),
        "unchanged_count": int((preview["Status"] == "UNCHANGED").sum()),
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
    progress_callback=None,
    chunk_size: int = 50,
) -> tuple[int, int]:
    """
    Upsert into Access Details table.

    Fetches existing records once (or uses existing_df if pre-fetched), classifies
    each row as an update-by-code, update-by-barcode, or insert, then executes all
    three groups as chunked executemany calls — reducing N round trips to 3 batches.
    Match order: SupplierID + Code first, then SupplierID + Barcode fallback.

    progress_callback(rows_done, total) is called after each chunk if provided.
    """
    df = export_df.copy()
    df["SupplierID"] = int(supplier_id)
    df["Code"] = df.get("Supplier_Code", "").fillna("").astype(str)
    df["Description"] = df.get("TradeName", "").fillna("").astype(str).str.slice(0, MAX_DESC_LEN)
    df["Pharmacode"] = df.get("Pharmacode", "").fillna("").astype(str)
    df["Cost"] = (pd.to_numeric(df.get("Cost", 0), errors="coerce").fillna(0.0) * 100).round().astype(int)
    df["Retail"] = (pd.to_numeric(df.get("Retail", 0), errors="coerce").fillna(0.0) * 100).round().astype(int)
    df["Barcode"] = _first_barcode(df.get("Barcode", "").fillna("").astype(str))
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
    by_pharm: dict[str, pd.Series] = {}
    pharm_counts: dict[str, int] = {}
    for _, row in existing.iterrows():
        code = str(row["Code"] or "").strip()
        barcode = str(row["Barcode"] or "").strip()
        pharm = str(row["Pharmacode"] or "").strip()
        if code:
            by_code[code] = row
        if barcode:
            by_barcode[barcode] = row
        if pharm:
            pharm_counts[pharm] = pharm_counts.get(pharm, 0) + 1
            by_pharm[pharm] = row
    # Only use pharmacode as a fallback when it is unique for this supplier
    by_pharm = {k: v for k, v in by_pharm.items() if pharm_counts[k] == 1}

    # Classify each export row — skip rows where nothing has changed
    update_code_rows: list[tuple] = []
    update_barcode_rows: list[tuple] = []
    insert_rows: list[tuple] = []

    for r in df.itertuples(index=False):
        sup_code = (r.Code or "").strip()
        # r.Barcode is already the first barcode only (stripped by _first_barcode above)
        barcode = (r.Barcode or "").strip()
        desc = r.Description
        pharmacode = str(r.Pharmacode or "").strip()
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
        elif pharmacode and pharmacode in by_pharm:
            ex = by_pharm[pharmacode]
            # Update using the barcode stored in Access (the reliable single value)
            existing_barcode = str(ex["Barcode"] or "").strip()
            if _row_changed(ex, desc, cost, retail, pharmacode):
                update_barcode_rows.append(upd_params + (int(supplier_id), existing_barcode))
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

    total = len(update_code_rows) + len(update_barcode_rows) + len(insert_rows)
    done = 0

    def _run_chunks(cur, sql, rows):
        nonlocal done
        for i in range(0, len(rows), chunk_size):
            cur.executemany(sql, rows[i : i + chunk_size])
            done += len(rows[i : i + chunk_size])
            if progress_callback:
                progress_callback(done, total)

    with get_access_conn() as conn:
        cur = conn.cursor()
        cur.fast_executemany = False
        _run_chunks(cur, update_by_partcode_sql, update_code_rows)
        _run_chunks(cur, update_by_barcode_sql, update_barcode_rows)
        _run_chunks(cur, insert_sql, insert_rows)
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
    progress_callback=None,
    chunk_size: int = 100,
) -> int:
    """
    Delete ALL existing Details records for supplier_id, then insert export_df fresh.
    Returns the number of rows inserted.
    If progress_callback is provided it is called as callback(done: int, total: int)
    after each chunk is inserted.
    """
    df = export_df.copy()
    df["SupplierID"] = int(supplier_id)
    df["Code"] = df.get("Supplier_Code", "").fillna("").astype(str)
    df["Description"] = df.get("TradeName", "").fillna("").astype(str).str.slice(0, MAX_DESC_LEN)
    df["Pharmacode"] = df.get("Pharmacode", "").fillna("").astype(str)
    df["Cost"] = (pd.to_numeric(df.get("Cost", 0), errors="coerce").fillna(0.0) * 100).round().astype(int)
    df["Retail"] = (pd.to_numeric(df.get("Retail", 0), errors="coerce").fillna(0.0) * 100).round().astype(int)
    df["Barcode"] = _first_barcode(df.get("Barcode", "").fillna("").astype(str))
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

    total = len(rows)
    sid = int(supplier_id)
    with get_access_conn() as conn:
        cur = conn.cursor()
        # Delete in batches to stay under Access MaxLocksPerFile (~9500 default)
        while True:
            cur.execute(
                "DELETE FROM Details WHERE SupplierID = ? AND PartCodeGuid IN "
                "(SELECT TOP 1000 PartCodeGuid FROM Details WHERE SupplierID = ?)",
                (sid, sid),
            )
            conn.commit()
            if cur.rowcount == 0:
                break
        cur.fast_executemany = False
        for start in range(0, total, chunk_size):
            chunk = rows[start:start + chunk_size]
            cur.executemany(insert_sql, chunk)
            if progress_callback is not None:
                progress_callback(min(start + chunk_size, total), total)
        cur.execute(
            "UPDATE Suppliers SET LastUpdated = ? WHERE SupplierID = ?",
            (now, int(supplier_id)),
        )
        conn.commit()

    return total
