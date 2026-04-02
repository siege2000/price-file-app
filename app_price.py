# app_price.py — Full price file workflow for the Streamlit app
from __future__ import annotations

import pandas as pd
import streamlit as st

from processing.file_helpers import load_templates
from processing.column_helpers import safe_str, tpl_field_default, tpl_desc_defaults
from processing.description_helpers import combine_columns, normalize_units, clean_description
from processing.editing_helpers import parse_money, apply_replacements
from db.access_helpers import (
    MAX_DESC_LEN, load_suppliers, load_supplier_details,
    build_upsert_preview, upsert_details, replace_all_details,
)
from db import sqlite_helper
from app_state import (
    apply_supplier_settings_to_state,
    get_supplier_sheet_layouts,
    get_supplier_rules,
    save_supplier_rules,
    save_supplier_settings,
    get_excel_sheets,
    load_uploaded_file,
    read_all_supplier_rules,
)


# ── Cached loaders ────────────────────────────────────────────────────────────

@st.cache_data(show_spinner=False)
def _cached_load_templates():
    return load_templates("data/templates.json")


@st.cache_data(ttl=300, show_spinner=False)
def _cached_load_suppliers():
    df = load_suppliers()
    df["Label"] = df.apply(
        lambda r: f'{r["SupplierName"]} ({r["SupplierCode"]})'
        if str(r["SupplierCode"] or "").strip()
        else str(r["SupplierName"]),
        axis=1,
    )
    return df


@st.cache_data(ttl=120, show_spinner=False)
def _cached_supplier_details(supplier_id: int) -> pd.DataFrame:
    return load_supplier_details(supplier_id)


# ── Helpers ───────────────────────────────────────────────────────────────────

def _preview_to_export_df(df: pd.DataFrame) -> pd.DataFrame:
    result = pd.DataFrame({
        "PLU": df.get("PLU", pd.Series([""] * len(df))),
        "Supplier_Code": df["Supplier_Code"],
        "Pharmacode": df.get("Pharmacode", pd.Series([""] * len(df))),
        "Retail": df["New_Retail"],
        "Cost": df["New_Cost"],
        "TradeName": df["New_Description"],
        "Barcode": df["Barcode"],
    })
    if "Outers" in df.columns:
        result["Outers"] = df["Outers"].values
    return result


def _get_supplier_id(suppliers_df: pd.DataFrame, label: str) -> int | None:
    row = suppliers_df[suppliers_df["Label"] == label]
    return int(row["SupplierID"].iloc[0]) if not row.empty else None


def _get_supplier_name_code(suppliers_df: pd.DataFrame, label: str) -> tuple[str, str]:
    row = suppliers_df[suppliers_df["Label"] == label]
    if row.empty:
        return ("", "")
    return (str(row["SupplierName"].iloc[0] or ""), str(row["SupplierCode"].iloc[0] or ""))


# ── Main render function ──────────────────────────────────────────────────────

def render_price_file():
    # ── Load static data ─────────────────────────────────────────────────────
    try:
        templates = _cached_load_templates()
    except Exception as e:
        st.error(f"Could not load templates: {e}")
        return

    try:
        suppliers_df = _cached_load_suppliers()
    except Exception as e:
        st.error(f"Could not load suppliers from Access: {e}")
        suppliers_df = pd.DataFrame(columns=["SupplierID", "SupplierName", "SupplierCode", "Label"])

    template_names = list(templates.keys())
    supplier_labels = ["— select supplier —"] + suppliers_df["Label"].tolist()

    # ── Top row: supplier, template, skip rows, clear ────────────────────────
    col_sup, col_tpl, col_skip, col_clr = st.columns([3, 2, 1, 1])

    with col_sup:
        prev_sup = st.session_state.get("price_supplier_label", supplier_labels[0])
        supplier_label = st.selectbox(
            "Supplier",
            supplier_labels,
            index=supplier_labels.index(prev_sup) if prev_sup in supplier_labels else 0,
            key="price_supplier_label",
        )

    supplier_id = _get_supplier_id(suppliers_df, supplier_label)

    # Prefetch supplier details when supplier changes
    if supplier_id is not None:
        prefetch_sid = st.session_state.get("price_prefetch_sid")
        if prefetch_sid != supplier_id:
            with st.spinner("Fetching supplier data from Access…"):
                try:
                    df_prefetch = _cached_supplier_details(supplier_id)
                    st.session_state["price_prefetch_df"] = df_prefetch
                    st.session_state["price_prefetch_sid"] = supplier_id
                    st.session_state.pop("price_settings_loaded_for", None)
                except Exception as e:
                    st.warning(f"Could not prefetch supplier data: {e}")
                    st.session_state["price_prefetch_df"] = None
                    st.session_state["price_prefetch_sid"] = supplier_id

    # Sheet layout combo — shown when supplier has >1 saved layout
    sheet_layouts = get_supplier_sheet_layouts(supplier_id) if supplier_id else {}
    layout_keys = list(sheet_layouts.keys())
    if len(layout_keys) > 1:
        layout_labels = ["(default)" if k == "_default" else k for k in layout_keys]
        sel_layout_label = st.selectbox(
            "Sheet layout",
            layout_labels,
            key="price_sheet_layout_label",
        )
        sel_layout_idx = layout_labels.index(sel_layout_label)
        sheet_key = layout_keys[sel_layout_idx]
    else:
        sheet_key = "_default"

    with col_tpl:
        tpl_default = st.session_state.get("price_template", template_names[0])
        tpl_idx = template_names.index(tpl_default) if tpl_default in template_names else 0
        template_name = st.selectbox(
            "Template",
            template_names,
            index=tpl_idx,
            key="price_template",
        )

    tpl = templates[template_name]
    opts = tpl.get("options", {})

    with col_skip:
        skip_rows = st.number_input(
            "Skip rows",
            min_value=0, max_value=1000,
            value=int(st.session_state.get("price_skip_rows", int(opts.get("skip_rows_default", 0)))),
            step=1,
            key="price_skip_rows",
        )

    with col_clr:
        st.write("")
        if st.button("Clear / Reset", key="price_clear"):
            for k in list(st.session_state.keys()):
                if k.startswith("price_") and k not in (
                    "price_supplier_label", "price_template",
                ):
                    del st.session_state[k]
            st.rerun()

    st.divider()

    # ── File upload ───────────────────────────────────────────────────────────
    uploaded = st.file_uploader(
        "Open price file (CSV or Excel)",
        type=["csv", "xlsx", "xls"],
        key="price_uploader",
    )

    if uploaded is None:
        st.info("Upload a price file to get started.")
        return

    # Detect file change → reset derived state
    if st.session_state.get("price_file_name") != uploaded.name:
        for k in ["price_df_raw", "price_excel_sheets", "price_file_sheet",
                  "price_export_df", "price_preview_df", "price_preview_stats",
                  "price_settings_loaded_for"]:
            st.session_state.pop(k, None)
        st.session_state["price_file_name"] = uploaded.name

    # Excel sheet picker
    is_excel = uploaded.name.lower().endswith((".xlsx", ".xls"))
    if is_excel and "price_excel_sheets" not in st.session_state:
        sheets = get_excel_sheets(uploaded)
        st.session_state["price_excel_sheets"] = sheets

    excel_sheets = st.session_state.get("price_excel_sheets", [])
    if is_excel and len(excel_sheets) > 1:
        prev_sheet = st.session_state.get("price_file_sheet")
        sheet_choice = st.selectbox(
            "Select Excel sheet",
            excel_sheets,
            index=excel_sheets.index(prev_sheet) if prev_sheet in excel_sheets else 0,
            key="price_file_sheet",
        )
        if sheet_choice != prev_sheet:
            st.session_state.pop("price_df_raw", None)
            st.session_state.pop("price_export_df", None)
            st.session_state.pop("price_preview_df", None)
    else:
        sheet_choice = None

    # Load raw data
    if "price_df_raw" not in st.session_state or \
            st.session_state.get("price_df_skip") != skip_rows or \
            st.session_state.get("price_df_sheet") != sheet_choice:
        try:
            df = load_uploaded_file(uploaded, sheet_name=sheet_choice, skip_rows=skip_rows)
            st.session_state["price_df_raw"] = df
            st.session_state["price_df_skip"] = skip_rows
            st.session_state["price_df_sheet"] = sheet_choice
            # Reset build/preview on file reload
            st.session_state.pop("price_export_df", None)
            st.session_state.pop("price_preview_df", None)
            st.session_state.pop("price_settings_loaded_for", None)
        except Exception as e:
            st.error(f"Could not load file: {e}")
            return

    df = st.session_state["price_df_raw"]

    # Apply supplier settings to state (once per supplier+sheet change)
    if supplier_id is not None:
        apply_supplier_settings_to_state(supplier_id, sheet_key, df.columns.tolist())

    # ── Layout: left (source data + options) | right (mapping panel) ─────────
    left_col, right_col = st.columns([3, 2], gap="large")

    with left_col:
        st.subheader(f"Source Data Preview — {len(df)} rows")
        st.dataframe(df, use_container_width=True, height=300)

        # Export options
        with st.expander("Export Options", expanded=True):
            ec1, ec2, ec3 = st.columns(3)
            with ec1:
                delim_opts = [",", ";", "Tab"]
                delim_default = opts.get("csv_delimiter", ",")
                delim_idx = delim_opts.index(delim_default) if delim_default in delim_opts else 0
                out_delim = st.selectbox("CSV delimiter", delim_opts, index=delim_idx, key="price_delim")
            with ec2:
                dec_opts = [2, 3, 0]
                dec_default = int(opts.get("price_decimals", 2))
                dec_idx = dec_opts.index(dec_default) if dec_default in dec_opts else 0
                decimals = st.selectbox("Price decimals", dec_opts, index=dec_idx, key="price_decimals")
            with ec3:
                st.write("")
                allow_blank_plu = st.checkbox(
                    "Allow blank PLU",
                    value=bool(opts.get("allow_blank_plu", True)),
                    key="price_allow_blank_plu",
                )
            excl_no_id = st.checkbox(
                "Exclude rows with no Barcode, Supplier Code and Pharmacode",
                value=True,
                key="price_excl_no_id",
            )

    # ── Right: column mapping + TradeName settings ────────────────────────────
    with right_col:
        cols_with_none = ["(None)"] + df.columns.tolist()
        fields = tpl.get("fields", {})

        def _default_idx(key: str) -> int:
            state_key = f"price_col_{key}"
            saved = st.session_state.get(state_key)
            if saved and saved in cols_with_none:
                return cols_with_none.index(saved)
            d = tpl_field_default(df, fields.get(key, {}))
            return cols_with_none.index(d) if d and d in cols_with_none else 0

        st.subheader("Column Mapping")
        col_plu = st.selectbox("PLU", cols_with_none, index=_default_idx("plu"), key="price_col_plu")
        col_pharmacode = st.selectbox("Pharmacode", cols_with_none, index=_default_idx("pharmacode"), key="price_col_pharmacode")
        col_retail = st.selectbox("Retail *", cols_with_none, index=_default_idx("retail"), key="price_col_retail")
        col_cost = st.selectbox("Cost *", cols_with_none, index=_default_idx("cost"), key="price_col_cost")
        col_barcode = st.selectbox("Barcode/EAN (optional)", cols_with_none, index=_default_idx("barcode"), key="price_col_barcode")
        col_sc = st.selectbox("Supplier Code (optional)", cols_with_none, index=_default_idx("sc"), key="price_col_sc")
        col_outer = st.selectbox("Outer (optional)", cols_with_none, index=_default_idx("outer"), key="price_col_outer")

        st.subheader("TradeName Settings")
        desc_cfg = tpl.get("description", tpl.get("tradename", {}))

        # Ordered column selection with move-up/down
        if "price_desc_cols_ordered" not in st.session_state:
            st.session_state["price_desc_cols_ordered"] = df.columns.tolist()
        if "price_desc_cols_selected" not in st.session_state:
            defaults = tpl_desc_defaults(df, desc_cfg)
            st.session_state["price_desc_cols_selected"] = defaults

        # Filter ordered list to only include columns in current df
        ordered = [c for c in st.session_state["price_desc_cols_ordered"] if c in df.columns]
        remaining_new = [c for c in df.columns if c not in ordered]
        st.session_state["price_desc_cols_ordered"] = ordered + remaining_new

        selected_set = set(st.session_state.get("price_desc_cols_selected", []))

        desc_cols_sel = st.multiselect(
            "TradeName columns (selected = included)",
            options=st.session_state["price_desc_cols_ordered"],
            default=[c for c in st.session_state["price_desc_cols_ordered"] if c in selected_set],
            key="price_desc_cols_multiselect",
        )

        # Reorder: keep order from price_desc_cols_ordered, filtered to selected
        desc_cols = [c for c in st.session_state["price_desc_cols_ordered"] if c in desc_cols_sel]
        st.session_state["price_desc_cols_selected"] = desc_cols

        # Move-up / move-down
        if desc_cols:
            st.caption("Order: " + " → ".join(desc_cols))
            mu_col, md_col, _ = st.columns([1, 1, 3])
            with mu_col:
                if st.button("▲ Up", key="price_desc_up") and len(desc_cols) > 1:
                    # Move last-selected up one position in the ordered list
                    last = desc_cols[-1]
                    o = st.session_state["price_desc_cols_ordered"]
                    idx = o.index(last)
                    if idx > 0:
                        o[idx], o[idx - 1] = o[idx - 1], o[idx]
                        st.session_state["price_desc_cols_ordered"] = o
                        st.rerun()
            with md_col:
                if st.button("▼ Down", key="price_desc_down") and len(desc_cols) > 1:
                    last = desc_cols[-1]
                    o = st.session_state["price_desc_cols_ordered"]
                    idx = o.index(last)
                    if idx < len(o) - 1:
                        o[idx], o[idx + 1] = o[idx + 1], o[idx]
                        st.session_state["price_desc_cols_ordered"] = o
                        st.rerun()

        sep_opts = [" ", " - ", "-", " | ", ", "]
        sep_default = st.session_state.get("price_sep", desc_cfg.get("separator", " "))
        sep_idx = sep_opts.index(sep_default) if sep_default in sep_opts else 0
        sep = st.selectbox("Separator", sep_opts, index=sep_idx, key="price_sep")

        prefix = st.text_input(
            "Prefix (e.g. HH — leave blank for none)",
            value=st.session_state.get("price_prefix", ""),
            key="price_prefix",
        )

        tc_default = st.session_state.get("price_title_case", bool(desc_cfg.get("title_case", False)))
        nu_default = st.session_state.get("price_norm_units", bool(desc_cfg.get("normalize_units", True)))
        title_case = st.checkbox("Title Case", value=tc_default, key="price_title_case")
        norm_units = st.checkbox("Normalize units (75 g → 75g)", value=nu_default, key="price_norm_units")

        # Live sample preview
        if desc_cols and len(df) > 0:
            try:
                sample = combine_columns(df.head(1), desc_cols, sep=sep).iloc[0]
                sample = clean_description(pd.Series([sample])).iloc[0]
                if title_case:
                    sample = sample.title()
                if norm_units:
                    sample = normalize_units(pd.Series([sample])).iloc[0]
                if prefix:
                    sample = prefix + " " + sample
                ch = len(sample)
                color = "red" if ch > 40 else "green"
                st.markdown(
                    f'<span style="color:{color}">Sample: {sample[:55]}{"…" if len(sample)>55 else ""} ({ch} chars)</span>',
                    unsafe_allow_html=True,
                )
            except Exception:
                pass

    # ── Supplier rules expander ───────────────────────────────────────────────
    if supplier_id is not None:
        with st.expander("Supplier Shortening Rules (auto-apply during Build Output)", expanded=False):
            existing_rules = get_supplier_rules(supplier_id)
            n_rules = len(existing_rules) if existing_rules is not None else 0

            if n_rules:
                st.success(f"⚡ {n_rules} rule(s) will auto-apply during Build Output.")
            else:
                st.caption("No shortening rules defined for this supplier.")

            # Load from supplier_rules.json as list of dicts for the editor
            all_rules_data = read_all_supplier_rules()
            current_rules_list = [
                r for r in all_rules_data.get(str(supplier_id), [])
                if isinstance(r, dict) and "from" in r
            ]
            rules_edit_df = pd.DataFrame(
                current_rules_list if current_rules_list else [{"from": "", "to": ""}]
            )
            if "from" not in rules_edit_df.columns:
                rules_edit_df["from"] = ""
            if "to" not in rules_edit_df.columns:
                rules_edit_df["to"] = ""

            edited_rules_df = st.data_editor(
                rules_edit_df[["from", "to"]],
                num_rows="dynamic",
                use_container_width=True,
                key=f"price_rules_editor_{supplier_id}",
                column_config={
                    "from": st.column_config.TextColumn("Find (prefix 'regex:' for regex)"),
                    "to": st.column_config.TextColumn("Replace with"),
                },
            )
            if st.button("Save Rules", key="price_rules_save"):
                try:
                    new_rules = [
                        {"from": str(r["from"]).strip(), "to": str(r["to"]).strip()}
                        for _, r in edited_rules_df.iterrows()
                        if str(r.get("from", "")).strip()
                    ]
                    save_supplier_rules(supplier_id, new_rules)
                    st.success(f"Saved {len(new_rules)} rule(s).")
                    st.rerun()
                except Exception as e:
                    st.error(f"Could not save rules: {e}")

    st.divider()

    # ── Action buttons: Build, Truncate, Rules, Save Settings ────────────────
    ab1, ab2, ab3, ab4 = st.columns(4)

    with ab1:
        build_clicked = st.button("Build Output", type="primary", key="price_build_btn")

    with ab2:
        truncate_clicked = st.button("Bulk Truncate (40 chars)", key="price_truncate_btn")

    with ab3:
        mark_updated = st.checkbox("Mark rows as Updated", value=True, key="price_mark_updated")

    with ab4:
        save_settings_clicked = st.button("Save Settings", key="price_save_settings")

    # Build Output
    if build_clicked:
        if supplier_id is None:
            st.warning("Select a supplier first.")
        else:
            try:
                out = pd.DataFrame(index=df.index)
                out["PLU"] = safe_str(df[col_plu]) if col_plu != "(None)" else ""
                out["Supplier_Code"] = safe_str(df[col_sc]) if col_sc != "(None)" else ""
                out["Pharmacode"] = safe_str(df[col_pharmacode]) if col_pharmacode != "(None)" else ""
                out["Retail"] = parse_money(df[col_retail]).round(decimals) if col_retail != "(None)" else 0.0
                out["Cost"] = parse_money(df[col_cost]).round(decimals) if col_cost != "(None)" else 0.0

                trade = (
                    combine_columns(df, desc_cols, sep=sep)
                    if desc_cols
                    else pd.Series([""] * len(df), index=df.index)
                )
                trade = clean_description(trade)
                if title_case:
                    trade = trade.str.title()
                if norm_units:
                    trade = normalize_units(trade)
                if prefix:
                    trade = prefix + " " + trade
                out["TradeName"] = trade.fillna("").astype(str)

                # Auto-apply per-supplier shortening rules
                rules_df = get_supplier_rules(supplier_id)
                if rules_df is not None and not rules_df.empty:
                    out["TradeName"], n_repl = apply_replacements(
                        out["TradeName"], rules_df, whole_word=True, case_insensitive=True
                    )
                    if n_repl:
                        st.info(f"Supplier rules: {n_repl} replacement(s) applied.")

                out["Barcode"] = safe_str(df[col_barcode]) if col_barcode != "(None)" else ""

                if col_outer != "(None)":
                    outer_vals = pd.to_numeric(df[col_outer], errors="coerce").fillna(0).round().astype(int)
                    out["Outers"] = outer_vals.where(outer_vals > 0, 1)
                else:
                    out["Outers"] = 1

                out = out.dropna(how="all")

                # De-dupe non-blank PLUs
                if col_plu != "(None)":
                    non_blank = out["PLU"].astype("string").str.strip() != ""
                    out = pd.concat([
                        out[~non_blank],
                        out[non_blank].drop_duplicates(subset=["PLU"], keep="last"),
                    ]).sort_index()

                # Exclude rows with no identifying key
                if st.session_state.get("price_excl_no_id", True):
                    has_id = (
                        out["Barcode"].astype(str).str.strip().ne("") |
                        out["Supplier_Code"].astype(str).str.strip().ne("") |
                        out["Pharmacode"].astype(str).str.strip().ne("")
                    )
                    excluded = (~has_id).sum()
                    out = out[has_id]
                    if excluded:
                        st.info(f"Excluded {excluded} row(s) with no Barcode, Supplier Code or Pharmacode.")

                st.session_state["price_export_df"] = out.reset_index(drop=True)
                st.session_state.pop("price_preview_df", None)
                st.session_state.pop("price_preview_stats", None)
                st.success(f"Built {len(out)} rows — click **Preview Changes** to compare with Access.")
            except Exception as e:
                st.error(f"Build error: {e}")

    # Save Settings
    if save_settings_clicked:
        if supplier_id is None:
            st.warning("Select a supplier first.")
        else:
            try:
                sheet_settings = {
                    "template": template_name,
                    "skip_rows": skip_rows,
                    "columns": {
                        "plu": col_plu, "pharmacode": col_pharmacode,
                        "retail": col_retail, "cost": col_cost,
                        "barcode": col_barcode, "sc": col_sc, "outer": col_outer,
                    },
                    "description": {
                        "columns": desc_cols,
                        "prefix": prefix,
                        "separator": sep,
                        "title_case": title_case,
                        "normalize_units": norm_units,
                    },
                }
                save_supplier_settings(supplier_id, sheet_key, sheet_settings)
                st.success(f"Settings saved for {supplier_label} [{sheet_key}]")
                # Reset so the new settings are picked up next time
                st.session_state["price_settings_loaded_for"] = (supplier_id, sheet_key)
            except Exception as e:
                st.error(f"Could not save settings: {e}")

    # ── Output Editor ─────────────────────────────────────────────────────────
    export_df = st.session_state.get("price_export_df")
    if export_df is None:
        return

    # Add CharCount
    export_df = export_df.copy()
    export_df["CharCount"] = export_df["TradeName"].str.len()

    # Bulk Truncate
    if truncate_clicked:
        export_df["TradeName"] = export_df["TradeName"].str.slice(0, 40)
        export_df["CharCount"] = export_df["TradeName"].str.len()
        st.session_state["price_export_df"] = export_df.drop(columns=["CharCount"])
        st.rerun()

    long_count = (export_df["CharCount"] > 40).sum()
    if long_count:
        st.warning(f"⚠ {long_count} row(s) exceed 40 chars — use Bulk Truncate or edit below.")

    st.subheader("Output Preview & Editor")
    st.caption("Edit TradeName cells directly. Rows with CharCount > 40 should be fixed before upsert.")

    edited_out = st.data_editor(
        export_df,
        use_container_width=True,
        height=300,
        key="price_output_editor",
        column_config={
            "TradeName": st.column_config.TextColumn("TradeName", width="large"),
            "CharCount": st.column_config.NumberColumn("Chars", format="%d"),
            "Retail": st.column_config.NumberColumn(format="$%.2f"),
            "Cost": st.column_config.NumberColumn(format="$%.2f"),
        },
        disabled=[c for c in export_df.columns if c != "TradeName"],
    )

    # Keep export_df in sync with edits (drop CharCount for storage)
    if "CharCount" in edited_out.columns:
        st.session_state["price_export_df"] = edited_out.drop(columns=["CharCount"])
    else:
        st.session_state["price_export_df"] = edited_out

    export_df = st.session_state["price_export_df"]

    # Download CSV
    delim_map = {",": ",", ";": ";", "Tab": "\t"}
    delim_char = delim_map.get(st.session_state.get("price_delim", ","), ",")
    csv_bytes = export_df.to_csv(index=False, sep=delim_char).encode("utf-8")
    st.download_button(
        "Download CSV",
        data=csv_bytes,
        file_name=f"price_file_{template_name}.csv",
        mime="text/csv",
        key="price_download_csv",
    )

    st.divider()

    # ── Preview / Upsert bar ──────────────────────────────────────────────────
    st.subheader("Preview & Write to Access")

    pc1, pc2, pc3, pc4, pc5 = st.columns(5)
    with pc1:
        preview_clicked = st.button("Preview Changes", key="price_preview_btn")
    with pc2:
        upsert_clicked = st.button("Execute Upsert", type="primary", key="price_upsert_btn")
    with pc3:
        replace_clicked = st.button("Replace All in Access", key="price_replace_btn")
    with pc4:
        sqlite_upsert_clicked = st.button("Upsert → SQLite", key="price_sqlite_upsert_btn")
    with pc5:
        sqlite_replace_clicked = st.button("Replace All → SQLite", key="price_sqlite_replace_btn")

    # Preview Changes
    if preview_clicked:
        if supplier_id is None:
            st.warning("Select a supplier first.")
        else:
            existing = st.session_state.get("price_prefetch_df") \
                if st.session_state.get("price_prefetch_sid") == supplier_id else None
            with st.spinner("Comparing with Access…"):
                try:
                    preview_df, stats = build_upsert_preview(
                        export_df, supplier_id, existing_df=existing
                    )
                    preview_df["Chars"] = preview_df["New_Description"].str.len()
                    preview_df["Valid"] = preview_df["Chars"].apply(lambda x: "✓" if x <= 40 else "⚠")
                    # Reorder Chars/Valid after New_Description
                    cols_list = list(preview_df.columns)
                    for extra in ["Valid", "Chars"]:
                        if extra in cols_list:
                            cols_list.remove(extra)
                            cols_list.insert(cols_list.index("New_Description") + 1, extra)
                    preview_df = preview_df[cols_list]
                    st.session_state["price_preview_df"] = preview_df
                    st.session_state["price_preview_stats"] = stats
                except Exception as e:
                    st.error(f"Preview failed: {e}")

    # Display preview
    preview_df = st.session_state.get("price_preview_df")
    stats = st.session_state.get("price_preview_stats")

    if preview_df is not None and stats is not None:
        new_c = (preview_df["Status"] == "NEW").sum()
        upd_c = (preview_df["Status"] == "UPDATE").sum()
        unch_c = (preview_df["Status"] == "UNCHANGED").sum()
        long_c = (preview_df.get("Chars", 0) > 40).sum()

        st.caption(
            f"{stats['existing_count']} in Access  |  "
            f"matched by code: {stats['matched_by_code']}  |  "
            f"matched by barcode: {stats['matched_by_barcode']}  |  "
            f"**{new_c} new**  |  **{upd_c} updates**  |  {unch_c} unchanged"
            + (f"  |  ⚠ {long_c} exceed 40 chars" if long_c else "")
        )

        # Sort controls
        sc1, sc2, sc3 = st.columns([1, 1, 5])
        with sc1:
            if st.button("Sort New at Top", key="price_sort_new"):
                df_s = st.session_state["price_preview_df"].copy()
                df_s["_s"] = df_s["Status"].map({"NEW": 0, "UPDATE": 1}).fillna(2)
                st.session_state["price_preview_df"] = df_s.sort_values("_s").drop(columns=["_s"]).reset_index(drop=True)
                st.rerun()
        with sc2:
            if st.button("Sort Invalid at Top", key="price_sort_invalid"):
                df_s = st.session_state["price_preview_df"].copy()
                df_s["_s"] = df_s.get("Valid", "✓").apply(lambda v: 0 if str(v) == "⚠" else 1)
                st.session_state["price_preview_df"] = df_s.sort_values("_s").drop(columns=["_s"]).reset_index(drop=True)
                st.rerun()

        # Diagnostics expander
        with st.expander("Match Diagnostics"):
            def _fmt(lst):
                return ", ".join(lst) if lst else "(none)"
            st.text(
                f"Access codes (sample):    {_fmt(stats['access_codes'])}\n"
                f"Export codes (sample):    {_fmt(stats['export_codes'])}\n"
                f"Access barcodes (sample): {_fmt(stats['access_barcodes'])}\n"
                f"Export barcodes (sample): {_fmt(stats['export_barcodes'])}\n\n"
                f"If all rows show as NEW, check that codes/barcodes match in format "
                f"(leading zeros, spaces, etc.)."
            )

        # Preview table (editable New_Description, deletable rows)
        edited_preview = st.data_editor(
            preview_df,
            use_container_width=True,
            height=400,
            num_rows="dynamic",
            key="price_preview_editor",
            column_config={
                "Status": st.column_config.TextColumn("Status", width="small"),
                "Valid": st.column_config.TextColumn("Valid", width="small"),
                "Chars": st.column_config.NumberColumn("Chars", format="%d", width="small"),
                "New_Description": st.column_config.TextColumn("New Description", width="large"),
                "Old_Description": st.column_config.TextColumn("Old Description", width="large"),
                "Old_Cost": st.column_config.NumberColumn("Old Cost", format="$%.2f"),
                "New_Cost": st.column_config.NumberColumn("New Cost", format="$%.2f"),
                "Old_Retail": st.column_config.NumberColumn("Old Retail", format="$%.2f"),
                "New_Retail": st.column_config.NumberColumn("New Retail", format="$%.2f"),
            },
            disabled=[c for c in preview_df.columns if c not in ("New_Description",)],
        )

        # Sync edited preview back to state, updating Chars/Valid
        edited_preview = edited_preview.copy()
        edited_preview["Chars"] = edited_preview["New_Description"].astype(str).str.len()
        edited_preview["Valid"] = edited_preview["Chars"].apply(lambda x: "✓" if x <= 40 else "⚠")
        st.session_state["price_preview_df"] = edited_preview

    # Execute Upsert
    if upsert_clicked:
        if supplier_id is None:
            st.warning("Select a supplier first.")
        else:
            cur_preview = st.session_state.get("price_preview_df")
            if cur_preview is None:
                st.warning("Run Preview Changes first.")
            else:
                cur_preview = cur_preview.copy()
                invalid_mask = cur_preview.get("Chars", pd.Series([0] * len(cur_preview))) > 40
                fallback_mask = invalid_mask & cur_preview.get("Status", pd.Series([""] * len(cur_preview))).eq("UPDATE")
                skip_mask = invalid_mask & ~fallback_mask

                # Fallback UPDATE rows with > 40 chars to old description
                if fallback_mask.any():
                    cur_preview.loc[fallback_mask, "New_Description"] = (
                        cur_preview.loc[fallback_mask, "Old_Description"].fillna("").astype(str)
                    )
                    st.info(f"{int(fallback_mask.sum())} UPDATE row(s) with >40 chars will keep their existing description.")

                valid_rows = cur_preview[~skip_mask].copy()
                skipped_rows = cur_preview[skip_mask].copy()

                if valid_rows.empty:
                    st.warning("All rows have descriptions > 40 chars. Fix them before upserting.")
                else:
                    if not skipped_rows.empty:
                        st.warning(f"{len(skipped_rows)} NEW row(s) with >40 chars will be skipped.")

                    exp = _preview_to_export_df(valid_rows)
                    existing = st.session_state.get("price_prefetch_df") \
                        if st.session_state.get("price_prefetch_sid") == supplier_id else None
                    with st.spinner("Writing to Access…"):
                        try:
                            updated, inserted = upsert_details(
                                exp, supplier_id, mark_updated=mark_updated, existing_df=existing
                            )
                            st.success(f"Done — {inserted} inserted, {updated} updated for {supplier_label}")
                            _cached_supplier_details.clear()
                            st.session_state.pop("price_prefetch_df", None)
                            st.session_state.pop("price_prefetch_sid", None)
                            if not skipped_rows.empty:
                                st.session_state["price_preview_df"] = skipped_rows
                            else:
                                st.session_state.pop("price_preview_df", None)
                                st.session_state.pop("price_preview_stats", None)
                        except Exception as e:
                            st.error(f"Upsert failed: {e}")

    # Replace All in Access
    if replace_clicked:
        if supplier_id is None:
            st.warning("Select a supplier first.")
        else:
            cur_preview = st.session_state.get("price_preview_df")
            use_df = _preview_to_export_df(cur_preview) if cur_preview is not None else export_df
            long_cnt = 0
            if cur_preview is not None:
                long_cnt = int((cur_preview.get("Chars", 0) > 40).sum())
            if long_cnt:
                st.warning(f"{long_cnt} row(s) have descriptions > 40 chars. Fix before replacing.")
            else:
                st.session_state["price_replace_confirm"] = True

    if st.session_state.get("price_replace_confirm"):
        cur_preview = st.session_state.get("price_preview_df")
        use_df = _preview_to_export_df(cur_preview) if cur_preview is not None else export_df
        st.error(
            f"⚠ **REPLACE ALL** will permanently delete every existing record for **{supplier_label}** "
            f"and insert {len(use_df)} row(s) fresh. This cannot be undone."
        )
        conf1, conf2 = st.columns(2)
        with conf1:
            if st.button("Yes, Replace All", key="price_replace_confirm_yes"):
                with st.spinner("Deleting and re-inserting…"):
                    try:
                        inserted = replace_all_details(use_df, supplier_id, mark_updated=mark_updated)
                        st.success(f"Done — all previous records deleted, {inserted} row(s) inserted for {supplier_label}.")
                        _cached_supplier_details.clear()
                        st.session_state.pop("price_prefetch_df", None)
                        st.session_state.pop("price_prefetch_sid", None)
                        st.session_state.pop("price_preview_df", None)
                        st.session_state.pop("price_preview_stats", None)
                        st.session_state.pop("price_replace_confirm", None)
                    except Exception as e:
                        st.error(f"Replace failed: {e}")
        with conf2:
            if st.button("Cancel", key="price_replace_confirm_no"):
                st.session_state.pop("price_replace_confirm", None)
                st.rerun()

    # SQLite Upsert
    if sqlite_upsert_clicked:
        if supplier_id is None:
            st.warning("Select a supplier first.")
        else:
            cur_preview = st.session_state.get("price_preview_df")
            use_df = _preview_to_export_df(cur_preview) if cur_preview is not None else export_df
            name, code = _get_supplier_name_code(suppliers_df, supplier_label)
            with st.spinner("Writing to SQLite…"):
                try:
                    sqlite_helper.ensure_supplier_by_id(supplier_id, name, code)
                    updated, inserted = sqlite_helper.upsert_details(
                        use_df, supplier_id, mark_updated=mark_updated
                    )
                    st.success(f"SQLite done — {inserted} inserted, {updated} updated for {supplier_label}.")
                except Exception as e:
                    st.error(f"SQLite upsert failed: {e}")

    # SQLite Replace All
    if sqlite_replace_clicked:
        if supplier_id is None:
            st.warning("Select a supplier first.")
        else:
            st.session_state["price_sqlite_replace_confirm"] = True

    if st.session_state.get("price_sqlite_replace_confirm"):
        cur_preview = st.session_state.get("price_preview_df")
        use_df = _preview_to_export_df(cur_preview) if cur_preview is not None else export_df
        st.error(
            f"⚠ **REPLACE ALL → SQLite** will permanently delete every SQLite record for **{supplier_label}** "
            f"and insert {len(use_df)} row(s) fresh."
        )
        sc1, sc2 = st.columns(2)
        with sc1:
            if st.button("Yes, Replace SQLite", key="price_sqlite_replace_yes"):
                name, code = _get_supplier_name_code(suppliers_df, supplier_label)
                with st.spinner("Deleting and re-inserting in SQLite…"):
                    try:
                        sqlite_helper.ensure_supplier_by_id(supplier_id, name, code)
                        inserted = sqlite_helper.replace_all_details(
                            use_df, supplier_id, mark_updated=mark_updated
                        )
                        st.success(f"SQLite done — {inserted} row(s) inserted for {supplier_label}.")
                        st.session_state.pop("price_sqlite_replace_confirm", None)
                    except Exception as e:
                        st.error(f"SQLite replace failed: {e}")
        with sc2:
            if st.button("Cancel SQLite", key="price_sqlite_replace_no"):
                st.session_state.pop("price_sqlite_replace_confirm", None)
                st.rerun()
