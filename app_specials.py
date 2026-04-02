# app_specials.py — Full specials workflow for the Streamlit app
from __future__ import annotations

import os
from datetime import date, datetime

import pandas as pd
import streamlit as st

from specials.specials_helpers import (
    load_brands,
    get_special_id,
    save_specials,
    clear_old_specials,
    load_special_settings,
    save_special_settings,
)
from specials.specials_widget import (
    SPECIALS_COLUMNS,
    _FIELD_SEP,
    _EXCEL_TO_COMBO,
    _EXCEL_TO_GRID,
    _BRAND_MAPPINGS,
)


# ── Cached loaders ────────────────────────────────────────────────────────────

@st.cache_data(ttl=300, show_spinner=False)
def _cached_load_brands() -> pd.DataFrame:
    try:
        clear_old_specials()
    except Exception:
        pass
    return load_brands()


# ── Grid building (pure function extracted from specials_widget._apply_mapping_to_grid) ───

def build_specials_grid(
    raw: pd.DataFrame,
    brand_name: str,
    col_combos: dict[str, int],  # key → 0-based field index (-1 = None)
) -> pd.DataFrame:
    """
    Map raw file columns → specials grid using brand mapping, named columns,
    or combo positions. Pure function — no Qt dependencies.
    """
    bm = _BRAND_MAPPINGS.get(brand_name, {})
    bm_cols = bm.get("columns", {})
    bm_rules = bm.get("rules", {})

    col_lower = {c.lower(): c for c in raw.columns}

    def by_name(*names) -> pd.Series | None:
        for name in names:
            actual = col_lower.get(name.lower())
            if actual:
                return raw[actual].astype(str).str.strip()
        return None

    def bm_col(role: str, *fallback_names) -> pd.Series | None:
        mapped = bm_cols.get(role)
        candidates = ([mapped] if mapped else []) + list(fallback_names)
        return by_name(*candidates)

    def combo_col(key: str) -> pd.Series:
        field_idx = col_combos.get(key, -1)
        if 0 <= field_idx < len(raw.columns):
            return raw.iloc[:, field_idx].astype(str).str.strip()
        return pd.Series([""] * len(raw))

    def first_not_none(*series) -> pd.Series:
        for s in series:
            if s is not None:
                return s
        return pd.Series([""] * len(raw))

    desc_s          = first_not_none(bm_col("description", "Product name"), combo_col("desc"))
    barcode_s       = first_not_none(bm_col("barcode", "Barcodes", "Barcode"), combo_col("barcode"))
    posnote_s       = first_not_none(bm_col("pos_note", "Promotion POS note", "POS Note"), combo_col("posnote"))
    product_price_s = bm_col("product_price", "Product price")
    deal_disc_type_s = bm_col("discount_type", "Deal discount type")
    deal_disc_val_s  = bm_col("discount_value", "Deal discount value")
    min_qty_s        = bm_col("multibuy", "Min qty")
    qty_breaks_s     = bm_col("multibuy_fallback", "Quantity breaks")
    deal_name_col    = bm_col("deal_name", "Deal name")
    deal_name_s      = first_not_none(deal_name_col, pd.Series([""] * len(raw)))

    # Also handle direct Excel→grid column mappings
    direct_cols: dict[str, pd.Series | None] = {}
    for excel_name, grid_col in _EXCEL_TO_GRID.items():
        s = by_name(excel_name)
        if s is not None:
            direct_cols[grid_col] = s

    def _v(series, i) -> str:
        return _clean(str(series.iloc[i])) if series is not None else ""

    def _clean(val: str) -> str:
        v = val.strip()
        return "" if v.lower() == "nan" else v

    # GroupID map: sequential number per unique Deal name
    deal_group_map: dict[str, int] = {}
    if deal_name_col is not None:
        next_gid = 1
        for dn in deal_name_col:
            key_dn = _clean(str(dn))
            if key_dn and key_dn not in deal_group_map:
                deal_group_map[key_dn] = next_gid
                next_gid += 1

    rows = []
    for i in range(len(raw)):
        barcode = _clean(str(barcode_s.iloc[i]))
        if "," in barcode:
            barcode = ",".join(
                b.replace('"', "").replace(" ", "")
                for b in barcode.split(",")
                if b.strip()
            )
        else:
            barcode = barcode.replace('"', "").replace(" ", "")

        desc = _clean(str(desc_s.iloc[i]))[:40]

        # Special price
        prod_price  = _v(product_price_s, i)
        disc_type   = _v(deal_disc_type_s, i)
        disc_val    = _v(deal_disc_val_s, i)
        if prod_price:
            sp = prod_price
        elif disc_type == "%" and disc_val:
            try:
                pct = float(disc_val)
                sp = (str(int(pct)) if pct == int(pct) else str(pct)) + "%"
            except ValueError:
                sp = disc_val + "%"
        else:
            sp = _v(combo_col("special") if col_combos.get("special", -1) >= 0 else None, i)

        # MultiBuy
        mb_raw = _v(min_qty_s, i) or _v(qty_breaks_s, i)
        mb = ""
        if mb_raw:
            try:
                mb = str(int(float(mb_raw)))
            except ValueError:
                mb = mb_raw

        # LowestFree
        lf_rule = bm_rules.get("lowest_free", "disc_pct_and_multibuy_gt1")
        if lf_rule == "always_false":
            lowest_free = "N"
        elif lf_rule.startswith("from_column:"):
            lf_col = lf_rule.split(":", 1)[1]
            lf_val = _v(by_name(lf_col), i) if by_name(lf_col) is not None else ""
            lowest_free = "Y" if lf_val.upper() in ("Y", "YES", "TRUE", "1") else "N"
        else:
            lowest_free = "N"
            if disc_type == "%" and disc_val and mb:
                try:
                    if float(disc_val) > 0 and float(mb) > 1:
                        lowest_free = "Y"
                except ValueError:
                    pass

        flybuys = "Y" if bm_rules.get("flybuys", False) else "N"

        deal_name_val = _clean(str(deal_name_s.iloc[i]))
        gid_rule = bm_rules.get("group_id", "sequential_by_deal_name")
        if gid_rule == "sequential_by_deal_name":
            group_id = str(deal_group_map.get(deal_name_val, "")) if deal_group_map else ""
        elif gid_rule.startswith("from_column:"):
            gid_col = gid_rule.split(":", 1)[1]
            group_id = _v(by_name(gid_col), i) if by_name(gid_col) is not None else ""
        else:
            group_id = ""

        # Invalid flag
        invalid = "Y" if not barcode and not desc else "N"

        row = {c: "" for c in SPECIALS_COLUMNS}
        row["Barcode"]         = barcode
        row["Description"]     = desc
        row["Chars"]           = str(len(desc))
        row["MultiBuy Qty"]    = mb
        row["Special Price"]   = sp
        row["Get Lowest Free"] = lowest_free
        row["New"]             = "Y"
        row["Changed"]         = "N"
        row["Invalid"]         = invalid
        row["POS Note"]        = _clean(str(posnote_s.iloc[i]))
        row["Flybuys"]         = flybuys
        row["Group ID"]        = group_id
        row["Deal Name"]       = deal_name_val

        # Overlay direct Excel→grid column mappings
        for grid_col, series in direct_cols.items():
            if grid_col in row:
                row[grid_col] = _clean(str(series.iloc[i]))

        rows.append(row)

    return pd.DataFrame(rows, columns=SPECIALS_COLUMNS)


def _auto_detect_from_headers(columns: list[str], combos: dict) -> dict:
    """Return updated combos dict with auto-detected column indices."""
    col_lower = {c.lower(): i for i, c in enumerate(columns)}
    for excel_name, combo_key in _EXCEL_TO_COMBO.items():
        if excel_name in col_lower:
            combos[combo_key] = col_lower[excel_name]
    return combos


def _autofill_header_fields(raw: pd.DataFrame, brand_name: str) -> dict:
    """Return dict with special_name, start_date, end_date from file data."""
    result = {}
    if len(raw) == 0:
        return result
    col_lower = {c.lower(): c for c in raw.columns}
    bm_cols = _BRAND_MAPPINGS.get(brand_name, {}).get("columns", {})

    def first_col_val(*candidates) -> str:
        for c in candidates:
            actual = col_lower.get(c.lower())
            if actual:
                v = str(raw[actual].iloc[0]).strip()
                return "" if v.lower() == "nan" else v
        return ""

    name_col = bm_cols.get("special_name", "")
    val = first_col_val(name_col, "Promotion name", "Special name", "SpecialName")
    if val:
        result["special_name"] = val

    start_col = bm_cols.get("start_date", "")
    val = first_col_val(start_col, "Promotion start date", "Start date", "StartDate")
    if val:
        try:
            dt = datetime.fromisoformat(val.split(" ")[0])
            result["start_date"] = dt.date()
        except Exception:
            pass

    end_col = bm_cols.get("end_date", "")
    val = first_col_val(end_col, "Promotion end date", "Finish date", "End date", "FinishDate")
    if val:
        try:
            dt = datetime.fromisoformat(val.split(" ")[0])
            result["end_date"] = dt.date()
        except Exception:
            pass

    return result


def _load_specials_file(uploaded, field_sep: str, ignore_first: bool) -> pd.DataFrame:
    name = uploaded.name.lower()
    ext = os.path.splitext(name)[1]
    if ext in (".xls", ".xlsx"):
        uploaded.seek(0)
        raw = pd.read_excel(
            uploaded,
            header=0 if not ignore_first else None,
            dtype=str,
        )
        if ignore_first:
            raw.columns = [f"Column {i+1}" for i in range(len(raw.columns))]
    else:
        for enc in ("utf-8-sig", "cp1252", "latin-1"):
            try:
                uploaded.seek(0)
                raw = pd.read_csv(
                    uploaded,
                    sep=field_sep,
                    header=0 if not ignore_first else None,
                    dtype=str,
                    encoding=enc,
                )
                break
            except UnicodeDecodeError:
                continue
        else:
            uploaded.seek(0)
            raw = pd.read_csv(uploaded, sep=field_sep, header=0 if not ignore_first else None,
                              dtype=str, encoding="latin-1")
        if ignore_first:
            raw.columns = [f"Column {i+1}" for i in range(len(raw.columns))]
    return raw.fillna("")


# ── Main render function ──────────────────────────────────────────────────────

def render_specials():
    # Load brands
    try:
        brands_df = _cached_load_brands()
    except Exception as e:
        st.error(f"Could not load brands from Access: {e}")
        brands_df = pd.DataFrame(columns=["BrandID", "BrandName"])

    brand_labels = ["(Select Brand)"] + brands_df["BrandName"].tolist()

    # ── Left (grid) | Right (settings) ───────────────────────────────────────
    left_col, right_col = st.columns([3, 2], gap="large")

    with right_col:
        # Brand selection
        st.subheader("Brand")
        prev_brand_label = st.session_state.get("spec_brand_label", brand_labels[0])
        brand_label = st.selectbox(
            "Brand",
            brand_labels,
            index=brand_labels.index(prev_brand_label) if prev_brand_label in brand_labels else 0,
            key="spec_brand_label",
        )
        brand_id = 0
        if brand_label != "(Select Brand)":
            row = brands_df[brands_df["BrandName"] == brand_label]
            brand_id = int(row["BrandID"].iloc[0]) if not row.empty else 0

        password = st.text_input("Password", type="password", key="spec_password")

        # Load brand settings when brand changes
        if brand_id and st.session_state.get("spec_settings_loaded_for") != brand_id:
            try:
                settings = load_special_settings(brand_id)
                st.session_state["spec_saved_settings"] = settings
                # Apply file format settings
                st.session_state["spec_field_delim_idx"] = min(
                    settings.get("field_delimiter", 0), 3
                )
                st.session_state["spec_ignore_first"] = bool(
                    settings.get("ignore_first_line", False)
                )
                # Column combos: VB6 0-based index
                for key in ("desc", "special", "barcode", "posnote"):
                    vb6_key = {
                        "desc": "description_col", "special": "special_col",
                        "barcode": "barcode_col", "posnote": "posnote_col",
                    }[key]
                    st.session_state[f"spec_col_{key}"] = max(
                        settings.get(vb6_key, -1), -1
                    )
                st.session_state["spec_settings_loaded_for"] = brand_id
            except Exception as e:
                st.warning(f"Could not load brand settings: {e}")

        st.divider()

        # File format
        st.subheader("File Format")
        delim_opts = ["Comma", "Tab", "Semicolon", "Pipe"]
        field_delim_idx = st.selectbox(
            "Field delimiter",
            delim_opts,
            index=st.session_state.get("spec_field_delim_idx", 0),
            key="spec_field_delim",
        )
        # Map label → index
        field_delim_sep = _FIELD_SEP.get(delim_opts.index(field_delim_idx), ",")

        ignore_first = st.checkbox(
            "Ignore first line of file",
            value=st.session_state.get("spec_ignore_first", False),
            key="spec_ignore_first_chk",
        )

        st.divider()

        # Column mapping (1-based display, 0-based storage as VB6 ListIndex)
        st.subheader("Column Mapping")
        st.caption("Set which column number in your file maps to each field (1 = first).")

        combo_opts = ["(None)"] + [str(n) for n in range(1, 21)]

        def _col_selectbox(label: str, key: str) -> int:
            """Render a column number selectbox. Returns 0-based field index (-1 if None)."""
            vb6_idx = st.session_state.get(f"spec_col_{key}", -1)
            combo_idx = vb6_idx + 1 if vb6_idx >= 0 else 0
            combo_idx = min(combo_idx, len(combo_opts) - 1)
            sel = st.selectbox(
                label,
                combo_opts,
                index=combo_idx,
                key=f"spec_col_sel_{key}",
            )
            return combo_opts.index(sel) - 1  # 0-based field index, -1 for None

        col_desc    = _col_selectbox("Description", "desc")
        col_special = _col_selectbox("Special Price", "special")
        col_barcode = _col_selectbox("Barcode / Pharmacode", "barcode")
        col_posnote = _col_selectbox("POS Note (optional)", "posnote")

        col_combos = {
            "desc": col_desc, "special": col_special,
            "barcode": col_barcode, "posnote": col_posnote,
        }

        if st.button("Save Column Settings", key="spec_save_col_settings"):
            if not brand_id:
                st.warning("Select a brand first.")
            else:
                try:
                    save_special_settings(brand_id, {
                        "description_col":  col_desc,
                        "special_col":      col_special,
                        "barcode_col":      col_barcode,
                        "posnote_col":      col_posnote,
                        "field_delimiter":  delim_opts.index(field_delim_idx),
                        "record_delimiter": 0,
                        "ignore_first_line": ignore_first,
                    })
                    st.success("Column settings saved.")
                except Exception as e:
                    st.error(f"Save error: {e}")

        st.divider()

        # Row actions
        st.subheader("Row Actions")
        ra1, ra2 = st.columns(2)
        with ra1:
            add_blank = st.button("Add Blank Row", key="spec_add_blank")
        with ra2:
            sort_invalid = st.button("Sort Invalid at Top", key="spec_sort_invalid")

        # Bulk Find & Replace
        with st.expander("Bulk Find & Replace"):
            fr_col = st.selectbox(
                "Column",
                ["Description", "MultiBuy Qty", "Special Price", "POS Note", "Multi Retail"],
                key="spec_fr_col",
            )
            fr_find = st.text_input("Find", key="spec_fr_find")
            fr_replace = st.text_input("Replace with", key="spec_fr_replace")
            fr_exact = st.checkbox("Exact match only", key="spec_fr_exact")
            if st.button("Replace All in Column", key="spec_fr_btn"):
                if fr_find and "spec_grid_df" in st.session_state:
                    gdf = st.session_state["spec_grid_df"].copy()
                    if fr_col in gdf.columns:
                        count = 0
                        for i in range(len(gdf)):
                            val = str(gdf.at[i, fr_col])
                            if fr_exact:
                                if val == fr_find:
                                    gdf.at[i, fr_col] = fr_replace
                                    count += 1
                            else:
                                if fr_find in val:
                                    gdf.at[i, fr_col] = val.replace(fr_find, fr_replace)
                                    count += 1
                        st.session_state["spec_grid_df"] = gdf
                        st.success(f"Replaced {count} cell(s).")
                        st.rerun()

        st.divider()

        # Log invalid items
        if st.button("Log Invalid Items (download CSV)", key="spec_log_invalid"):
            gdf = st.session_state.get("spec_grid_df", pd.DataFrame(columns=SPECIALS_COLUMNS))
            invalid_df = gdf[gdf.get("Invalid", pd.Series(["N"] * len(gdf))).str.upper() == "Y"].copy()
            if invalid_df.empty:
                st.info("No invalid items in the current grid.")
            else:
                def _reason(row) -> str:
                    reasons = []
                    if not str(row.get("Barcode", "")).strip():
                        reasons.append("No barcode")
                    if not str(row.get("Description", "")).strip():
                        reasons.append("No description")
                    return "; ".join(reasons) if reasons else "Unknown"
                invalid_df["Reason"] = invalid_df.apply(_reason, axis=1)
                invalid_df.insert(0, "Special Name", st.session_state.get("spec_special_name", ""))
                invalid_df.insert(0, "Brand", brand_label)
                invalid_df.insert(0, "Logged At", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
                report_cols = ["Logged At", "Brand", "Special Name", "Reason",
                               "Barcode", "Description", "Deal Name", "POS Note"]
                report_cols = [c for c in report_cols if c in invalid_df.columns]
                csv_bytes = invalid_df[report_cols].to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
                slug = st.session_state.get("spec_special_name", "specials").replace(" ", "_")
                fname = f"{slug}_invalid_{datetime.now().strftime('%Y%m%d')}.csv"
                st.download_button("Download Invalid Log", data=csv_bytes, file_name=fname, mime="text/csv",
                                   key="spec_log_download")

        # Stockcard search
        st.subheader("Find Barcode from Stock Database")
        with st.expander("Search stock database for barcode"):
            sc_search_term = st.text_input("Search description", key="spec_stockcard_term")
            if st.button("Search", key="spec_stockcard_search"):
                if not sc_search_term.strip():
                    st.warning("Enter a search term.")
                else:
                    try:
                        from specials.stockcard_dialog import search_stock
                        results = search_stock(sc_search_term)
                        st.session_state["spec_stockcard_results"] = results
                    except Exception as e:
                        st.error(f"Stock search failed: {e}")

            if "spec_stockcard_results" in st.session_state:
                results = st.session_state["spec_stockcard_results"]
                if results.empty:
                    st.info("No results found.")
                else:
                    st.dataframe(results, use_container_width=True, height=200)
                    apply_row = st.number_input(
                        "Apply barcode to grid row number (1-based)",
                        min_value=1,
                        max_value=max(1, len(st.session_state.get("spec_grid_df", pd.DataFrame()))),
                        step=1,
                        key="spec_stockcard_apply_row",
                    )
                    apply_barcode = st.text_input("Barcode to apply", key="spec_stockcard_apply_barcode")
                    if st.button("Apply to Row", key="spec_stockcard_apply"):
                        if "spec_grid_df" in st.session_state and apply_barcode:
                            gdf = st.session_state["spec_grid_df"].copy()
                            idx = int(apply_row) - 1
                            if 0 <= idx < len(gdf):
                                gdf.at[idx, "Barcode"] = apply_barcode.strip()
                                st.session_state["spec_grid_df"] = gdf
                                st.success(f"Row {apply_row}: barcode set to {apply_barcode}.")
                                st.rerun()

    # ── Left: toolbar + grid ──────────────────────────────────────────────────
    with left_col:
        # Toolbar
        tb1, tb2 = st.columns(2)
        with tb1:
            spec_uploader = st.file_uploader(
                "Import Specials File",
                type=["csv", "xls", "xlsx", "txt"],
                key="spec_uploader",
            )
        with tb2:
            grid_uploader = st.file_uploader(
                "Load Saved Grid",
                type=["csv", "xls", "xlsx"],
                key="spec_grid_loader",
            )

        # Import new file
        if spec_uploader is not None:
            if st.session_state.get("spec_upload_name") != spec_uploader.name:
                st.session_state["spec_upload_name"] = spec_uploader.name
                try:
                    raw = _load_specials_file(spec_uploader, field_delim_sep, ignore_first)
                    st.session_state["spec_raw_df"] = raw
                    # Auto-detect columns
                    new_combos = _auto_detect_from_headers(list(raw.columns), dict(col_combos))
                    for k, v in new_combos.items():
                        st.session_state[f"spec_col_{k}"] = v
                    # Auto-fill header fields
                    header = _autofill_header_fields(raw, brand_label)
                    if "special_name" in header:
                        st.session_state["spec_special_name_autofill"] = header["special_name"]
                    if "start_date" in header:
                        st.session_state["spec_start_date"] = header["start_date"]
                    if "end_date" in header:
                        st.session_state["spec_end_date"] = header["end_date"]
                    # Build grid
                    grid = build_specials_grid(raw, brand_label, col_combos)
                    st.session_state["spec_grid_df"] = grid
                    st.rerun()
                except Exception as e:
                    st.error(f"Could not load specials file: {e}")

        # Load saved grid
        if grid_uploader is not None:
            if st.session_state.get("spec_grid_load_name") != grid_uploader.name:
                st.session_state["spec_grid_load_name"] = grid_uploader.name
                try:
                    ext = os.path.splitext(grid_uploader.name)[1].lower()
                    if ext in (".xls", ".xlsx"):
                        gdf = pd.read_excel(grid_uploader, dtype=str).fillna("")
                    else:
                        gdf = pd.read_csv(grid_uploader, dtype=str).fillna("")
                    for col in SPECIALS_COLUMNS:
                        if col not in gdf.columns:
                            gdf[col] = ""
                    st.session_state["spec_grid_df"] = gdf[SPECIALS_COLUMNS]
                    st.rerun()
                except Exception as e:
                    st.error(f"Could not load grid: {e}")

        # Row operation handling
        gdf = st.session_state.get("spec_grid_df", pd.DataFrame(columns=SPECIALS_COLUMNS))

        if add_blank:
            blank = pd.DataFrame([{c: "" for c in SPECIALS_COLUMNS}])
            st.session_state["spec_grid_df"] = pd.concat([gdf, blank], ignore_index=True)
            st.rerun()

        if sort_invalid:
            if "Invalid" in gdf.columns:
                gdf["_sort"] = gdf["Invalid"].apply(lambda v: 0 if str(v) == "Y" else 1)
                gdf = gdf.sort_values("_sort").drop(columns=["_sort"]).reset_index(drop=True)
                st.session_state["spec_grid_df"] = gdf
                st.rerun()

        # Legend
        st.caption("🔵 New   🔴 Invalid   🟠 Changed")

        # Grid display
        if gdf.empty:
            st.info("Import a specials file or load a saved grid to get started.")
        else:
            invalid_count = (gdf.get("Invalid", pd.Series(["N"] * len(gdf))).str.upper() == "Y").sum()
            st.subheader(f"Import Grid — {len(gdf)} row(s), {invalid_count} invalid")

            editable_cols = {"Description", "MultiBuy Qty", "Special Price", "POS Note", "Multi Retail"}
            disabled_cols = [c for c in gdf.columns if c not in editable_cols]

            edited_grid = st.data_editor(
                gdf,
                use_container_width=True,
                height=500,
                num_rows="dynamic",
                key="spec_grid_editor",
                column_config={
                    "Chars": st.column_config.NumberColumn("Chars", format="%d", width="small"),
                    "Description": st.column_config.TextColumn("Description", width="large"),
                    "New": st.column_config.TextColumn("New", width="small"),
                    "Changed": st.column_config.TextColumn("Changed", width="small"),
                    "Invalid": st.column_config.TextColumn("Invalid", width="small"),
                },
                disabled=disabled_cols,
            )

            # Keep Chars in sync after edits
            edited_grid = edited_grid.copy()
            if "Description" in edited_grid.columns and "Chars" in edited_grid.columns:
                edited_grid["Chars"] = edited_grid["Description"].astype(str).str.len().astype(str)
            st.session_state["spec_grid_df"] = edited_grid

            # Save grid download
            csv_bytes = edited_grid.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
            st.download_button(
                "Save Grid (CSV)",
                data=csv_bytes,
                file_name="specials_grid.csv",
                mime="text/csv",
                key="spec_save_grid_dl",
            )

    st.divider()

    # ── Bottom: Special Name, Dates, Save ────────────────────────────────────
    bc1, bc2, bc3 = st.columns(3)
    with bc1:
        special_name = st.text_input(
            "Special Name",
            value=st.session_state.get("spec_special_name_autofill", ""),
            placeholder="e.g. April 2026 Catalogue",
            key="spec_special_name",
        )
    with bc2:
        start_date = st.date_input(
            "Start Date",
            value=st.session_state.get("spec_start_date", date.today()),
            key="spec_start_date_input",
        )
    with bc3:
        from datetime import timedelta
        end_date = st.date_input(
            "Finish Date",
            value=st.session_state.get("spec_end_date", date.today() + timedelta(days=30)),
            key="spec_end_date_input",
        )

    if st.button("Save to Database", type="primary", key="spec_save_db"):
        if not special_name.strip():
            st.warning("Enter a Special Name.")
        elif not brand_id:
            st.warning("Select a Brand.")
        elif st.session_state.get("spec_grid_df", pd.DataFrame()).empty:
            st.warning("Grid is empty — nothing to save.")
        else:
            grid_to_save = st.session_state["spec_grid_df"]
            with st.spinner("Saving to Access…"):
                try:
                    sid = get_special_id(special_name.strip(), brand_id, create_new=True)
                    saved = save_specials(
                        grid_to_save, sid,
                        start_date=start_date, finish_date=end_date,
                        password=password, clear_existing=True,
                    )
                    st.success(f"Saved {saved} row(s) to database.")
                    # Clear grid after successful save
                    st.session_state.pop("spec_grid_df", None)
                    st.session_state.pop("spec_raw_df", None)
                    st.session_state.pop("spec_upload_name", None)
                    st.rerun()
                except Exception as e:
                    st.error(f"Save error: {e}")
