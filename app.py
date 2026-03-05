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
    
    load_details_for_supplier,
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

        # Required mapping
        if retail_col == "(None)" or cost_col == "(None)":
            st.warning("Please map at least Retail and Cost to generate the export file.")
            st.stop()

        # ---- Build output
        out = pd.DataFrame(index=df.index)

        out["PLU"] = safe_str(df[plu_col]) if plu_col != "(None)" else ""
        out["Supplier_Code"] = safe_str(df[supplier_code_col]) if supplier_code_col != "(None)" else ""
        out["Pharmacode"] = safe_str(df[pharmacode_col]) if pharmacode_col != "(None)" else ""
        out["Retail"] = parse_money(df[retail_col]).round(decimals)
        out["Cost"] = parse_money(df[cost_col]).round(decimals)

        trade = combine_columns(df, desc_cols, sep=desc_sep) if desc_cols else pd.Series([""] * len(df), index=df.index)
        trade = clean_description(trade)
        if title_case:
            trade = trade.str.title()
        if normalize_units_flag:
            trade = normalize_units(trade)
        out["TradeName"] = trade

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

        # ---- Export df
        export_df = out.copy()
        export_df["TradeName"] = export_df["TradeName"].fillna("").astype("string").str.slice(0, MAX_DESC_LEN)

        st.subheader("Output Preview (Price File)")
        st.dataframe(export_df, width="stretch", height=350)

        csv_bytes = export_df.to_csv(index=False, sep=out_delim).encode("utf-8")
        st.download_button(
            "Download Price File CSV",
            data=csv_bytes,
            file_name=f"price_file_{template_name}.csv",
            mime="text/csv",
            key=f"download_csv_{state_tag}"
        )

        # -------------------------
        # REVIEW + UPSERT TO ACCESS
        # -------------------------
        st.divider()
        st.subheader(f"Save to Access ({'suppliers.mdb'})")

        mark_updated = st.checkbox("Mark rows as Updated", value=True, key=f"mark_updated_{state_tag}")

        # Review mode state
        if "review_mode" not in st.session_state:
            st.session_state.review_mode = False

        c1, c2, c3 = st.columns([1, 1, 2])
        with c1:
            if st.button("Review changes", type="primary", key=f"review_btn_{state_tag}"):
                st.session_state.review_mode = True
        with c2:
            if st.session_state.review_mode:
                if st.button("Back", key=f"review_back_{state_tag}"):
                    st.session_state.review_mode = False

        # ----- Review screen -----
        if st.session_state.review_mode:
            old_df = load_details_for_supplier(supplier_id)

            # NEW fields
            new_df = export_df.copy()
            new_df["SupplierCode"] = new_df.get("Supplier_Code", "").fillna("").astype(str).str.strip()
            new_df["Barcode"] = new_df.get("Barcode", "").fillna("").astype(str).str.strip()
            new_df["New_TradeName"] = new_df.get("TradeName", "").fillna("").astype(str)
            new_df["New_Cost"] = pd.to_numeric(new_df.get("Cost", 0), errors="coerce").fillna(0)
            new_df["New_Retail"] = pd.to_numeric(new_df.get("Retail", 0), errors="coerce").fillna(0)

            # OLD fields (IMPORTANT: supplier code is Details.Code, tradename is Details.Description)
            old_df = old_df.copy()
            old_df["SupplierCode"] = old_df.get("Code", "").fillna("").astype(str).str.strip()
            old_df["Barcode"] = old_df.get("Barcode", "").fillna("").astype(str).str.strip()
            old_df["Old_TradeName"] = old_df.get("Description", "").fillna("").astype(str)
            old_df["Old_Cost"] = old_df["Cost"].fillna(0).astype(int)/100
            old_df["Old_Retail"] = old_df["Retail"].fillna(0).astype(int)/100

            # --- OR match logic: barcode first, then supplier code for remainder ---
            new_has_barcode = new_df[new_df["Barcode"] != ""].copy()
            merged_barcode = new_has_barcode.merge(
                old_df[["Barcode", "Old_TradeName", "Old_Cost", "Old_Retail"]],
                on="Barcode",
                how="left"
            )

            matched_barcode = merged_barcode["Old_TradeName"].notna() | merged_barcode["Old_Cost"].notna() | merged_barcode["Old_Retail"].notna()

            new_need_supcode = pd.concat([
                merged_barcode.loc[~matched_barcode, new_df.columns],
                new_df[new_df["Barcode"] == ""],
            ], ignore_index=True).drop_duplicates()

            merged_supcode = new_need_supcode.merge(
                old_df[["SupplierCode", "Old_TradeName", "Old_Cost", "Old_Retail"]],
                on="SupplierCode",
                how="left"
            )

            merged = pd.concat([merged_barcode.loc[matched_barcode], merged_supcode], ignore_index=True)

            merged["Changed"] = (
                (merged["Old_TradeName"].fillna("") != merged["New_TradeName"].fillna("")) |
                (merged["Old_Cost"].fillna(-999999) != merged["New_Cost"].fillna(-999999)) |
                (merged["Old_Retail"].fillna(-999999) != merged["New_Retail"].fillna(-999999))
            )

            st.subheader("Review changes (old vs new)")
            st.caption("Blank Old_* means no match found → will be inserted as new.")

            show_only_changes = st.checkbox("Show only changed rows", value=True, key=f"only_changes_{state_tag}")
            view = merged[merged["Changed"]] if show_only_changes else merged

            # Comparison first, then the rest of the new fields
            comparison_cols = [
                "SupplierCode",
                "Old_TradeName", "New_TradeName",
                "Old_Cost", "New_Cost",
                "Old_Retail", "New_Retail",
            ]
            other_new_cols = [c for c in export_df.columns if c not in ["TradeName", "Cost", "Retail"]]
            display_cols = [c for c in (comparison_cols + other_new_cols) if c in view.columns]

            st.dataframe(view[display_cols], width="stretch", height=600)

            st.divider()
            if st.button("Confirm + Save to Access", type="primary", key=f"confirm_save_{state_tag}"):
                try:
                    upd, ins = upsert_details(export_df, supplier_id, mark_updated=mark_updated)
                    st.success(f"Saved. Updated: {upd} | Inserted: {ins}")
                    st.session_state.review_mode = False
                except Exception as e:
                    st.exception(e)

            st.stop()

        # ----- No review: allow direct save (optional) -----
        st.caption("Tip: click “Review changes” first to confirm old vs new before saving.")
        if st.button("Save to Access now (no review)", type="secondary", key=f"save_now_{state_tag}"):
            try:
                upd, ins = upsert_details(export_df, supplier_id, mark_updated=mark_updated)
                st.success(f"Saved. Updated: {upd} | Inserted: {ins}")
            except Exception as e:
                st.exception(e)
    except Exception as e:
            st.error(f"An error has occurred: {e}")

# -------------------------
# SPECIALS FILE MODE
# -------------------------
elif mode == "Specials File":
    st.header("Specials File Editor (POC)")
    st.info("Your Specials mode code can stay as-is (not touched here).")
