# app.py
import pandas as pd
import streamlit as st

from helpers import (
    MAX_DESC_LEN,
    apply_replacements,
    combine_columns,
    idx_for,
    load_file,
    load_replacements_csv,
    load_templates,
    normalize_units,
    parse_money,
    safe_str,
    tpl_desc_defaults,
    tpl_field_default,
    load_suppliers,
    clean_description,
    save_to_details,
    build_upsert_preview,
    upsert_details,
)

st.set_page_config(layout="wide")
st.title("Price File Builder")

mode = st.radio(
    "What do you want to work on?",
    ["Price File", "Specials File"],
    horizontal=True
)

# -------------------------
# PRICE FILE MODE
# -------------------------
if mode == "Price File":
    uploaded_file = st.file_uploader(
        "Upload your CSV or Excel file",
        type=["csv", "xlsx"],
        key="price_uploader"
    )

    if uploaded_file is None:
        st.stop()

    try:
        # ---- Templates
        templates = load_templates("templates.json")
        template_name = st.selectbox("Template", list(templates.keys()), index=0)
        tpl = templates[template_name]

        # IMPORTANT: tag keys by template+file so defaults re-apply when these change
        state_tag = f"{template_name}__{uploaded_file.name}"

        # ---- Skip rows
        skip_default = int(tpl.get("options", {}).get("skip_rows_default", 0))
        skip_rows = st.number_input(
            "Exclude first N rows (supplier headers, notes, etc.)",
            min_value=0, max_value=1000, value=skip_default, step=1
        )

        # ---- Load data
        df = load_file(uploaded_file)
        if skip_rows > 0:
            df = df.iloc[skip_rows:].reset_index(drop=True)

        # Force ID-like columns to string to avoid 123456.0
        for c in df.columns:
            cl = c.strip().lower()
            if cl in {"pharmacode", "pharma code", "barcode", "barcode/ean", "ean", "sap code", "plu"}:
                df[c] = (
                    df[c].astype("string")
                    .str.replace(r"\.0$", "", regex=True)
                    .str.strip()
                )

        left, right = st.columns([3, 1], gap="large")

        # -------- LEFT: preview + export options + rules
        with left:
            st.subheader("Data Preview")
            st.dataframe(df, use_container_width=True, height=600)

            st.divider()
            st.subheader("Export options")
            opts = tpl.get("options", {})
            delim_default = opts.get("csv_delimiter", ",")
            out_delim = st.selectbox(
                "CSV delimiter",
                [",", ";", "\t"],
                index=[",", ";", "\t"].index(delim_default) if delim_default in [",", ";", "\t"] else 0,
                key=f"delim_{state_tag}"
            )

            decimals_default = int(opts.get("price_decimals", 2))
            decimals = int(st.selectbox(
                "Price decimals",
                [2, 3, 0],
                index=[2, 3, 0].index(decimals_default) if decimals_default in [2, 3, 0] else 0,
                key=f"decimals_{state_tag}"
            ))

            allow_blank_plu = st.checkbox(
                "Allow blank PLU",
                value=bool(opts.get("allow_blank_plu", True)),
                key=f"allow_blank_plu_{state_tag}"
            )

            st.divider()
            st.subheader("TradeName shortening rules (optional)")
            rules_file = st.file_uploader(
                "Upload replacement rules CSV (from,to)",
                type=["csv"],
                key=f"rules_uploader_{state_tag}"
            )
            whole_word = st.checkbox("Match whole words only", value=True, key=f"whole_word_{state_tag}")
            case_insensitive = st.checkbox("Case-insensitive match", value=True, key=f"case_insensitive_{state_tag}")
            apply_rules = st.button("Apply shortening rules to TradeName", type="primary", key=f"apply_rules_{state_tag}")

        # -------- RIGHT: mapping + tradename settings + supplier select
        with right:
            st.subheader("Supplier")
            suppliers_df = load_suppliers()
            suppliers_df["Label"] = suppliers_df.apply(
                lambda r: f'{r["SupplierName"]} ({r["SupplierCode"]})'
                if str(r["SupplierCode"] or "").strip()
                else f'{r["SupplierName"]}',
                axis=1
            )
            supplier_label = st.selectbox(
                "Supplier",
                suppliers_df["Label"].tolist(),
                key=f"supplier_select_{state_tag}"
            )
            supplier_id = int(suppliers_df.loc[suppliers_df["Label"] == supplier_label, "SupplierID"].iloc[0])

            st.divider()
            st.subheader("Map columns → Price File Fields")

            cols = ["(None)"] + df.columns.tolist()

            plu_default = tpl_field_default(df, tpl["fields"]["plu"])
            pharmacode_default = tpl_field_default(df, tpl["fields"].get("pharmacode", {}))
            retail_default = tpl_field_default(df, tpl["fields"]["retail"])
            cost_default = tpl_field_default(df, tpl["fields"]["cost"])

            sc_cfg = tpl.get("fields", {}).get("sc", None)
            sc_default = tpl_field_default(df, sc_cfg) if isinstance(sc_cfg, dict) else None

            barcode_default = tpl_field_default(df, tpl["fields"].get("barcode", {}))

            # Optional debug
            # st.write("DEBUG defaults:", dict(plu=plu_default, pharmacode=pharmacode_default, retail=retail_default, cost=cost_default, barcode=barcode_default, sc=sc_default))

            plu_col = st.selectbox("PLU", cols, index=idx_for(cols, plu_default), key=f"plu_{state_tag}")
            pharmacode_col = st.selectbox("Pharmacode", cols, index=idx_for(cols, pharmacode_default), key=f"pharmacode_{state_tag}")
            retail_col = st.selectbox("Retail", cols, index=idx_for(cols, retail_default), key=f"retail_{state_tag}")
            cost_col = st.selectbox("Cost", cols, index=idx_for(cols, cost_default), key=f"cost_{state_tag}")
            barcode_col = st.selectbox("Barcode/EAN (optional)", cols, index=idx_for(cols, barcode_default), key=f"barcode_{state_tag}")
            supplier_code_col = st.selectbox("Supplier Code (optional)", cols, index=idx_for(cols, sc_default), key=f"sc_{state_tag}")

            st.divider()
            st.subheader("TradeName settings")

            desc_cfg = tpl.get("description", tpl.get("tradename", {}))
            desc_defaults = tpl_desc_defaults(df, desc_cfg)

            desc_cols = st.multiselect(
                "TradeName columns (in order)",
                options=df.columns.tolist(),
                default=desc_defaults,
                key=f"desc_cols_{state_tag}"
            )

            sep_options = [" ", " - ", "-", " | ", ", "]
            sep_default = desc_cfg.get("separator", " ")
            desc_sep = st.selectbox(
                "TradeName separator",
                sep_options,
                index=sep_options.index(sep_default) if sep_default in sep_options else 0,
                key=f"desc_sep_{state_tag}"
            )

            title_case = st.checkbox("Title case TradeName", value=bool(desc_cfg.get("title_case", False)), key=f"title_case_{state_tag}")
            normalize_units_flag = st.checkbox("Normalize units (75 g → 75g)", value=bool(desc_cfg.get("normalize_units", True)), key=f"normalize_units_{state_tag}")

            st.divider()
            if st.button("Reset mappings to template defaults", key=f"reset_defaults_{state_tag}"):
                # delete only keys for THIS file+template
                for k in list(st.session_state.keys()):
                    if k.endswith(f"_{state_tag}"):
                        del st.session_state[k]
                st.rerun()


        # ---- Build output
       # ---- Build output
        out = pd.DataFrame(index=df.index)

        out["PLU"] = safe_str(df[plu_col]) if plu_col != "(None)" else ""
        out["Supplier_Code"] = safe_str(df[supplier_code_col]) if supplier_code_col != "(None)" else ""
        out["Pharmacode"] = safe_str(df[pharmacode_col]) if pharmacode_col != "(None)" else ""
        out["Retail"] = parse_money(df[retail_col]).round(decimals) if retail_col != "(None)" else 0.0
        out["Cost"] = parse_money(df[cost_col]).round(decimals) if cost_col != "(None)" else 0.0

        trade = combine_columns(df, desc_cols, sep=desc_sep) if desc_cols else pd.Series([""] * len(df), index=df.index)
        trade = clean_description(trade)
        if title_case:
            trade = trade.str.title()
        if normalize_units_flag:
            trade = normalize_units(trade)
        
        # Initialize TradeName and ensure it's string type
        out["TradeName"] = trade.fillna("").astype(str)

        if barcode_col != "(None)":
            out["Barcode"] = safe_str(df[barcode_col])

        out = out.dropna(how="all")

        # De-dupe only non-blank PLUs
        if plu_col != "(None)":
            non_blank = out["PLU"].astype("string").str.strip() != ""
            out = pd.concat([out[~non_blank], out[non_blank].drop_duplicates(subset=["PLU"], keep="last")]).sort_index()

        # Apply shortening rules
        if apply_rules:
            if rules_file is None:
                st.warning("Upload a replacement rules CSV first.")
            else:
                rules_df = load_replacements_csv(rules_file)
                out["TradeName"], change_count = apply_replacements(
                    out["TradeName"],
                    rules_df,
                    whole_word=whole_word,
                    case_insensitive=case_insensitive
                )
                st.success(f"Applied rules. Total replacements made: {change_count}")

        # ---- PREVIEW & MANUAL EDITING ----
        st.subheader("Output Preview & Editor")
        st.info("💡 You can edit the **TradeName** cells directly in the table below.")

        # Add character count column for visibility
        out["CharCount"] = out["TradeName"].str.len()
        
        # Reorder columns to put CharCount next to TradeName
        cols = list(out.columns)
        if "TradeName" in cols:
            idx = cols.index("TradeName")
            cols.insert(idx + 1, cols.pop(cols.index("CharCount")))
            out = out[cols]

        # Bulk Truncate Button
        if st.button("Bulk Truncate all TradeNames to 40 chars"):
            out["TradeName"] = out["TradeName"].str.slice(0, 40)
            out["CharCount"] = out["TradeName"].str.len()
            st.rerun()

        # Use data_editor to allow manual overrides
        edited_df = st.data_editor(
            out,
            use_container_width=True,
            height=400,
            column_config={
                "CharCount": st.column_config.NumberColumn("Chars", help="Length of TradeName", format="%d"),
                "TradeName": st.column_config.TextColumn("TradeName", width="large", help="Max 40 chars recommended"),
                "Retail": st.column_config.NumberColumn(format="$%.2f"),
                "Cost": st.column_config.NumberColumn(format="$%.2f"),
            },
            disabled=[c for c in out.columns if c != "TradeName"], # Only allow editing TradeName
            key=f"editor_{state_tag}"
        )

        # Update the export_df with edits and recalculate count
        export_df = edited_df.copy()
        export_df["CharCount"] = export_df["TradeName"].str.len()
        
        # Final safety clip for export (just in case)
        # export_df["TradeName"] = export_df["TradeName"].str.slice(0, MAX_DESC_LEN)

        # Show warning if any names are still too long
        long_names = export_df[export_df["CharCount"] > 40]
        if not long_names.empty:
            st.warning(f"⚠️ {len(long_names)} items still exceed 40 characters.")

        # ---- Export Actions
        csv_bytes = export_df.to_csv(index=False, sep=out_delim).encode("utf-8")
        st.download_button(
            "Download Price File CSV",
            data=csv_bytes,
            file_name=f"price_file_{template_name}.csv",
            mime="text/csv",
            key=f"download_csv_{state_tag}"
        )

        # ---- Upsert to Access (with preview)
        st.divider()
        st.subheader("Upsert to Access (supplier.mdb)")

        mark_updated = st.checkbox("Mark rows as Updated", value=True, key=f"mark_updated_{state_tag}")

        preview_key = f"upsert_preview_{state_tag}"

        col_prev, col_exec = st.columns([1, 1])
        with col_prev:
            if st.button("Preview Changes", key=f"preview_upsert_{state_tag}"):
                try:
                    preview_df, _ = build_upsert_preview(export_df, supplier_id)
                    st.session_state[preview_key] = preview_df
                except Exception as e:
                    st.error(f"Preview failed: {e}")
                    st.exception(e)

        with col_exec:
            if st.button("Execute Upsert", type="primary", key=f"exec_upsert_{state_tag}"):
                try:
                    updated, inserted = upsert_details(export_df, supplier_id, mark_updated=mark_updated)
                    st.success(f"Done — {inserted} inserted, {updated} updated for {supplier_label}")
                    # Clear preview so it refreshes on next preview click
                    st.session_state.pop(preview_key, None)
                except Exception as e:
                    st.error(f"Upsert failed: {e}")
                    st.exception(e)

        if preview_key in st.session_state:
            preview_df = st.session_state[preview_key]
            new_count = (preview_df["Status"] == "NEW").sum()
            upd_count = (preview_df["Status"] == "UPDATE").sum()
            st.caption(f"{new_count} new  |  {upd_count} updates")

            st.dataframe(
                preview_df,
                use_container_width=True,
                height=400,
                column_config={
                    "Status": st.column_config.TextColumn("Status", width="small"),
                    "Supplier_Code": st.column_config.TextColumn("Supplier Code", width="small"),
                    "Barcode": st.column_config.TextColumn("Barcode", width="small"),
                    "Old_Description": st.column_config.TextColumn("Old Description"),
                    "New_Description": st.column_config.TextColumn("New Description"),
                    "Old_Cost": st.column_config.NumberColumn("Old Cost", format="$%.2f"),
                    "New_Cost": st.column_config.NumberColumn("New Cost", format="$%.2f"),
                    "Old_Retail": st.column_config.NumberColumn("Old Retail", format="$%.2f"),
                    "New_Retail": st.column_config.NumberColumn("New Retail", format="$%.2f"),
                },
            )

    except Exception as e:
        st.error(f"An error occurred: {e}")

# -------------------------
# SPECIALS FILE MODE
# -------------------------
elif mode == "Specials File":
    st.header("Specials File Editor (POC)")
    st.info("Your Specials mode code can stay as-is (not touched here).")
