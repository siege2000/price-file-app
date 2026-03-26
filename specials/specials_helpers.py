# specials_helpers.py — Access DB operations for Import Specials
"""
Ported from VB6:
  - frmImportSpecials.frm  (cmdSave_Click, GetSpecialID, SaveSpecialSettings, LoadBrands)
  - basUpdateSupplier.bas  (ClearOldSpecials)

Key Access tables:
  Specials        – SpecialID (auto), SpecialName, BrandID, StartDate, FinishDate, Password
  SpecialItems    – SpecialID, Barcode, Description, MultiBuy, MultiRetail,
                    SpecialRetail, SpecialPercent, lowestFree, GroupID, SpecialNote,
                    GiftBarcode, GiftPharmacode, ComboPurchase, MessageIsQuestion,
                    DealName, SecondaryGroupID
  Brand           – BrandID, BrandName
  SuppUDSpecialsSettings – BrandID, Special, Description, Barcode,
                           FieldDelimiter, RecordDelimiter, IgnoreFirstLine, POSNote
"""
from __future__ import annotations

from datetime import date

import pandas as pd

from db.access_helpers import get_access_conn


# ── Brand ─────────────────────────────────────────────────────────────────────

def load_brands() -> pd.DataFrame:
    """Return DataFrame with columns BrandID, BrandName ordered by name."""
    with get_access_conn() as conn:
        return pd.read_sql(
            "SELECT BrandID, BrandName FROM Brand ORDER BY BrandName", conn
        )


# ── Special header (Specials table) ───────────────────────────────────────────

def get_special_id(
    special_name: str,
    brand_id: int,
    create_new: bool = False,
) -> int:
    """
    Find existing Specials record by name + brand.
    If not found and create_new=True, inserts a new record and returns its ID.
    Returns 0 if not found and create_new=False.
    """
    special_name = special_name[:50].strip()
    with get_access_conn() as conn:
        cur = conn.cursor()
        cur.execute(
            "SELECT SpecialID FROM Specials WHERE SpecialName = ? AND BrandID = ?",
            (special_name, int(brand_id)),
        )
        row = cur.fetchone()
        if row:
            return int(row[0])
        if not create_new:
            return 0
        cur.execute(
            "INSERT INTO Specials (SpecialName, BrandID) VALUES (?, ?)",
            (special_name, int(brand_id)),
        )
        conn.commit()
        cur.execute(
            "SELECT SpecialID FROM Specials WHERE SpecialName = ? AND BrandID = ?",
            (special_name, int(brand_id)),
        )
        row = cur.fetchone()
        return int(row[0]) if row else 0


# ── Save specials grid → SpecialItems ─────────────────────────────────────────

def _parse_special_price(price_raw: str) -> tuple[int, float]:
    """
    Convert the Special Price cell value.
    Returns (special_retail_cents, special_percent).
      "15%"   → (0, 15.0)
      ""      → (-1, 0.0)   # message-only / no price
      "12.99" → (1299, 0.0)
    """
    price_raw = price_raw.strip()
    if "%" in price_raw:
        try:
            pct = float(price_raw.replace("%", "").strip())
        except ValueError:
            pct = 0.0
        return 0, pct
    if price_raw == "":
        return -1, 0.0
    try:
        return round(float(price_raw) * 100), 0.0
    except ValueError:
        return -1, 0.0


def _parse_multibuy(multibuy_raw: str, is_combo: bool) -> str:
    """
    Determine MultiBuy field value.
      Combo=True         → "ALL"
      Qty <= 1, numeric  → ""   (no multi-buy)
      Otherwise          → raw value
    """
    if is_combo:
        return "ALL"
    mb = multibuy_raw.strip()
    if mb == "":
        return ""
    try:
        if float(mb) <= 1 and "," not in mb:
            return ""
    except ValueError:
        pass
    return mb


def _is_yes(val) -> bool:
    return str(val or "").upper().strip() in ("Y", "YES", "TRUE", "1")


def save_specials(
    grid_df: pd.DataFrame,
    special_id: int,
    start_date: date,
    finish_date: date,
    password: str = "",
    clear_existing: bool = True,
    progress_callback=None,
) -> int:
    """
    Write grid rows to SpecialItems for the given special_id.

    - Updates Specials.StartDate / FinishDate.
    - If clear_existing=True, deletes all existing SpecialItems for this special first.
    - Skips rows where Invalid == 'Y'.
    - Returns number of rows saved.
    """
    # Skip invalid rows
    inv_col = grid_df.get("Invalid", pd.Series(["N"] * len(grid_df)))
    df = grid_df[inv_col.fillna("N").str.upper() != "Y"].copy()
    total = len(df)

    with get_access_conn() as conn:
        cur = conn.cursor()

        # Update header dates
        cur.execute(
            "UPDATE Specials SET StartDate = ?, FinishDate = ? WHERE SpecialID = ?",
            (start_date, finish_date, int(special_id)),
        )

        # Optionally wipe existing items
        if clear_existing:
            cur.execute(
                "DELETE FROM SpecialItems WHERE SpecialID = ?", (int(special_id),)
            )

        saved = 0
        for _, r in df.iterrows():
            barcode     = str(r.get("Barcode", "") or "").strip()
            description = str(r.get("Description", "") or "").strip()[:40]
            combo       = _is_yes(r.get("Combo", ""))
            multibuy    = _parse_multibuy(str(r.get("MultiBuy Qty", "") or ""), combo)
            multi_retail = str(r.get("Multi Retail", "") or "").strip()
            special_retail, special_pct = _parse_special_price(
                str(r.get("Special Price", "") or "")
            )
            lowest_free    = _is_yes(r.get("Get Lowest Free", ""))
            group_id       = str(r.get("Group ID", "") or "").strip()
            pos_note       = str(r.get("POS Note", "") or "").strip()
            gift_barcode   = str(r.get("Gift Barcode", "") or "").strip()
            gift_pharmacode = str(r.get("Gift Pharmacode", "") or "").strip() or "False"
            msg_is_q       = _is_yes(r.get("Msg Is Question", ""))
            deal_name      = str(r.get("Deal Name", "") or "").strip()
            secondary_gid  = str(r.get("Secondary Group ID", "") or "").strip()

            # Check if row already exists (in case clear_existing=False)
            cur.execute(
                "SELECT COUNT(*) FROM SpecialItems WHERE SpecialID = ? AND Barcode = ?",
                (int(special_id), barcode),
            )
            exists = cur.fetchone()[0] > 0

            if exists:
                cur.execute(
                    """UPDATE SpecialItems SET
                        [Description]=?, [MultiBuy]=?, [MultiRetail]=?,
                        [SpecialRetail]=?, [SpecialPercent]=?, [lowestFree]=?,
                        [GroupID]=?, [SpecialNote]=?, [GiftBarcode]=?, [GiftPharmacode]=?,
                        [ComboPurchase]=?, [MessageIsQuestion]=?, [DealName]=?, [SecondaryGroupID]=?
                    WHERE SpecialID=? AND Barcode=?""",
                    (
                        description, multibuy, multi_retail,
                        special_retail, special_pct, lowest_free,
                        group_id, pos_note, gift_barcode, gift_pharmacode,
                        combo, msg_is_q, deal_name, secondary_gid,
                        int(special_id), barcode,
                    ),
                )
            else:
                cur.execute(
                    """INSERT INTO SpecialItems
                        ([SpecialID],[Barcode],[Description],[MultiBuy],[MultiRetail],
                         [SpecialRetail],[SpecialPercent],[lowestFree],[GroupID],
                         [SpecialNote],[GiftBarcode],[GiftPharmacode],
                         [ComboPurchase],[MessageIsQuestion],[DealName],[SecondaryGroupID])
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                    (
                        int(special_id), barcode, description, multibuy, multi_retail,
                        special_retail, special_pct, lowest_free, group_id,
                        pos_note, gift_barcode, gift_pharmacode,
                        combo, msg_is_q, deal_name, secondary_gid,
                    ),
                )
            saved += 1
            if progress_callback:
                progress_callback(saved, total)

        # Optional password update
        if password:
            cur.execute(
                "UPDATE Specials SET [Password] = ? WHERE SpecialID = ?",
                (password.upper().strip(), int(special_id)),
            )

        conn.commit()

    return saved


# ── Cleanup ───────────────────────────────────────────────────────────────────

def clear_old_specials():
    """
    Delete SpecialItems and Specials records whose FinishDate is older than 2 months.
    Called on form load (mirrors VB6 ClearOldSpecials in basUpdateSupplier).
    """
    with get_access_conn() as conn:
        cur = conn.cursor()
        cur.execute(
            "DELETE SpecialItems.* FROM Specials INNER JOIN SpecialItems "
            "ON Specials.SpecialID = SpecialItems.SpecialID "
            "WHERE FinishDate < dateadd('m',-2,now())"
        )
        cur.execute(
            "DELETE Specials.* FROM Specials "
            "WHERE FinishDate < dateadd('m',-2,now())"
        )
        conn.commit()


# ── Per-brand column-mapping settings (SuppUDSpecialsSettings) ────────────────

def load_special_settings(brand_id: int) -> dict:
    """
    Load saved column-mapping settings for brand_id.
    Returns a dict with keys matching save_special_settings, or {} if none saved.
    """
    with get_access_conn() as conn:
        df = pd.read_sql(
            "SELECT * FROM SuppUDSpecialsSettings WHERE BrandID = ?",
            conn,
            params=[int(brand_id)],
        )
    if df.empty:
        return {}
    r = df.iloc[0]
    return {
        "special_col":      int(r.get("Special", 0) or 0),
        "description_col":  int(r.get("Description", 0) or 0),
        "barcode_col":      int(r.get("Barcode", 0) or 0),
        "field_delimiter":  int(r.get("FieldDelimiter", 0) or 0),
        "record_delimiter": int(r.get("RecordDelimiter", 0) or 0),
        "ignore_first_line": str(r.get("IgnoreFirstLine", "FALSE")).upper() == "TRUE",
        "posnote_col":      int(r.get("POSNote", 0) or 0),
    }


def save_special_settings(brand_id: int, settings: dict):
    """Upsert column-mapping settings for brand_id into SuppUDSpecialsSettings."""
    with get_access_conn() as conn:
        cur = conn.cursor()
        cur.execute(
            "SELECT COUNT(*) FROM SuppUDSpecialsSettings WHERE BrandID = ?",
            (int(brand_id),),
        )
        exists = cur.fetchone()[0] > 0
        ignore = "TRUE" if settings.get("ignore_first_line") else "FALSE"
        vals = (
            settings.get("special_col", 0),
            settings.get("description_col", 0),
            settings.get("barcode_col", 0),
            settings.get("field_delimiter", 0),
            settings.get("record_delimiter", 0),
            ignore,
            settings.get("posnote_col", 0),
        )
        if exists:
            cur.execute(
                """UPDATE SuppUDSpecialsSettings SET
                    [Special]=?,[Description]=?,[Barcode]=?,
                    [FieldDelimiter]=?,[RecordDelimiter]=?,[IgnoreFirstLine]=?,[POSNote]=?
                WHERE BrandID=?""",
                (*vals, int(brand_id)),
            )
        else:
            cur.execute(
                """INSERT INTO SuppUDSpecialsSettings
                    (BrandID,Special,Description,Barcode,FieldDelimiter,RecordDelimiter,IgnoreFirstLine,POSNote)
                VALUES (?,?,?,?,?,?,?,?)""",
                (int(brand_id), *vals),
            )
        conn.commit()
