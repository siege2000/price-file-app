# app.py
import pandas as pd
import streamlit as st

from helpers import (
    MAX_DESC_LEN,
    apply_replacements,
    combine_columns,
    desc_len,
    idx_for,
    load_file,
    load_replacements_csv,
    load_templates,
    normalize_units,
    parse_money,
    safe_str,
    tpl_desc_defaults,
    tpl_field_default,
)

st.set_page_config(layout="wide")
st.title("Price File Builder")

mode = st.radio(
    "What do you want to work on?",
    ["Price File", "Specials File"],
    horizontal=True
)

def highlight_invalid_tradename(df_):
    def style_row(row):
        is_bad = int(row.get("DescLen", 0)) > MAX_DESC_LEN
        return ["background-color: #ffe6e6" if (is_bad and c == "TradeName") else "" for c in row.index]
    return df_.style.apply(style_row, axis=1)

# -------------------------
# PRICE FILE MODE
# -------------------------
if mode == "Price File":
    uploaded_file = st.file_uploader("Upload your CSV or Excel file", type=["csv", "xlsx"], key="price_uploader")

    if uploaded_file is None:
        st.stop()

    try:
        templates = load_templates("templates.json")
        template_name = st.selectbox("Template", list(templates.keys()), index=0)
        tpl = templates[template_name]

        skip_default = int(tpl.get("options", {}).get("skip_rows_default", 0))
        skip_rows = st.number_input(
            "Exclude first N rows (supplier headers, notes, etc.)",
            min_value=0, max_value=1000, value=skip_default, step=1
        )

        df = load_file(uploaded_file)
        if skip_rows > 0:
            df = df.iloc[skip_rows:].reset_index(drop=True)

        left, right = st.columns([3, 1], gap="large")

        # -------- LEFT: preview + export options + rules
        with left:
            st.subheader("Data Preview")
            st.dataframe(df, use_container_width=True, height=600)

            st.divider()
            st.subheader("Export options")
            opts = tpl.get("options", {})
            delim_default = opts.get("csv_delimiter", ",")
            out_delim = st.selectbox("CSV delimiter", [",", ";", "\t"],
                                     index=[",", ";", "\t"].index(delim_default) if delim_default in [",", ";", "\t"] else 0)

            decimals_default = int(opts.get("price_decimals", 2))
            decimals = int(st.selectbox("Price decimals", [2, 3, 0],
                                        index=[2, 3, 0].index(decimals_default) if decimals_default in [2, 3, 0] else 0))

            allow_blank_plu = st.checkbox("Allow blank PLU", value=bool(opts.get("allow_blank_plu", True)))

            st.divider()
            st.subheader("TradeName shortening rules (optional)")
            rules_file = st.file_uploader("Upload replacement rules CSV (from,to)", type=["csv"], key="rules_uploader")
            whole_word = st.checkbox("Match whole words only", value=True)
            case_insensitive = st.checkbox("Case-insensitive match", value=True)
            apply_rules = st.button("Apply shortening rules to TradeName", type="primary")

        # -------- RIGHT: mapping + tradename settings
        with right:
            st.subheader("Map columns → Price File Fields")
            cols = ["(None)"] + df.columns.tolist()

            plu_default = tpl_field_default(df, tpl["fields"]["plu"])
            retail_default = tpl_field_default(df, tpl["fields"]["retail"])
            cost_default = tpl_field_default(df, tpl["fields"]["cost"])

            # sc is optional in template
            sc_cfg = tpl.get("fields", {}).get("sc", None)
            sc_default = tpl_field_default(df, sc_cfg) if isinstance(sc_cfg, dict) else None

            barcode_default = tpl_field_default(df, tpl["fields"].get("barcode", {}))

            plu_col = st.selectbox("PLU", cols, index=idx_for(cols, plu_default))
            retail_col = st.selectbox("Retail", cols, index=idx_for(cols, retail_default))
            cost_col = st.selectbox("Cost", cols, index=idx_for(cols, cost_default))
            barcode_col = st.selectbox("Barcode/EAN (optional)", cols, index=idx_for(cols, barcode_default))
            supplier_code_col = st.selectbox("Supplier Code (optional)", cols, index=idx_for(cols, sc_default))

            st.divider()
            st.subheader("TradeName settings")

            # template key: prefer "description", but allow "tradename" if you changed JSON
            desc_cfg = tpl.get("description", tpl.get("tradename", {}))
            desc_defaults = tpl_desc_defaults(df, desc_cfg)

            desc_cols = st.multiselect(
                "TradeName columns (in order)",
                options=df.columns.tolist(),
                default=desc_defaults,
            )

            sep_options = [" ", " - ", "-", " | ", ", "]
            sep_default = desc_cfg.get("separator", " ")
            desc_sep = st.selectbox(
                "TradeName separator",
                sep_options,
                index=sep_options.index(sep_default) if sep_default in sep_options else 0,
            )

            title_case = st.checkbox("Title case TradeName", value=bool(desc_cfg.get("title_case", False)))
            normalize_units_flag = st.checkbox("Normalize units (75 g → 75g)", value=bool(desc_cfg.get("normalize_units", True)))

        # Required mapping
        if retail_col == "(None)" or cost_col == "(None)":
            st.warning("Please map at least Retail and Cost to generate the export file.")
            st.stop()

        # Build output
        out = pd.DataFrame(index=df.index)

        out["PLU"] = safe_str(df[plu_col]) if plu_col != "(None)" else ""
        out["Supplier_Code"] = safe_str(df[supplier_code_col]) if supplier_code_col != "(None)" else ""
        out["Retail"] = parse_money(df[retail_col]).round(decimals)
        out["Cost"] = parse_money(df[cost_col]).round(decimals)

        trade = combine_columns(df, desc_cols, sep=desc_sep) if desc_cols else pd.Series([""] * len(df), index=df.index)
        if title_case:
            trade = trade.str.title()
        if normalize_units_flag:
            trade = normalize_units(trade)
        out["TradeName"] = trade

        if barcode_col != "(None)":
            out["Barcode"] = safe_str(df[barcode_col])

        # Keep blank PLUs if allowed; just drop fully empty rows
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
                out["TradeName"], change_count = apply_replacements(out["TradeName"], rules_df, whole_word=whole_word, case_insensitive=case_insensitive)
                st.success(f"Applied rules. Total replacements made: {change_count}")

        # Validate TradeName length
        out["DescLen"] = desc_len(out["TradeName"])
        out["InvalidDesc"] = out["DescLen"] > MAX_DESC_LEN

        invalid = out[out["InvalidDesc"]].copy()

        st.subheader("TradeNames over 40 characters (fix these first)")
        if len(invalid) == 0:
            st.success("All TradeNames are 40 characters or less ✅")
        else:
            st.dataframe(highlight_invalid_tradename(invalid[["PLU", "TradeName", "DescLen"]]), use_container_width=True, height=250)
            edited_invalid = st.data_editor(
                invalid[["PLU", "TradeName"]],
                use_container_width=True,
                height=350,
                num_rows="fixed",
                column_config={
                    "PLU": st.column_config.TextColumn("PLU", disabled=True),
                    "TradeName": st.column_config.TextColumn("TradeName", help=f"Max {MAX_DESC_LEN} characters"),
                },
                key="invalid_desc_editor",
            )
            out.loc[invalid.index, "TradeName"] = safe_str(edited_invalid["TradeName"])
            out["DescLen"] = desc_len(out["TradeName"])
            out["InvalidDesc"] = out["DescLen"] > MAX_DESC_LEN

        export_df = out.drop(columns=["DescLen", "InvalidDesc"], errors="ignore")

        st.subheader("Output Preview (Price File)")
        st.dataframe(export_df, use_container_width=True, height=350)

        csv_bytes = export_df.to_csv(index=False, sep=out_delim).encode("utf-8")
        st.download_button(
            "Download Price File CSV",
            data=csv_bytes,
            file_name=f"price_file_{template_name}.csv",
            mime="text/csv",
        )

    except Exception as e:
        st.error(f"An error occurred: {e}")

# -------------------------
# SPECIALS FILE MODE
# -------------------------
elif mode == "Specials File":
    st.header("Specials File Editor (POC)")

    uploaded_specials = st.file_uploader("Upload Specials CSV/XLSX", type=["csv", "xlsx"], key="specials_uploader")
    if uploaded_specials is None:
        st.stop()

    if uploaded_specials.name.lower().endswith(".csv"):
        specials_df = pd.read_csv(uploaded_specials, sep=None, engine="python")
    else:
        specials_df = pd.read_excel(uploaded_specials)

    skip_rows = st.number_input("Exclude first N rows", 0, 1000, 0, key="specials_skip")
    if skip_rows > 0:
        specials_df = specials_df.iloc[skip_rows:].reset_index(drop=True)

    st.subheader("Imported Specials Preview")
    st.dataframe(specials_df, use_container_width=True, height=350)

    # (keep your specials editor code here — your block was fine, the indentation wasn’t)
      # ---- Expected column names (from your sample) ----
    # If your file is tab-separated, pandas still reads it fine; the headers are what matter.
    DATE_COLS = [
        "Promotion start date", "Promotion end date",
        "Deal start date", "Deal end date"
    ]

    ID_COLS = ["Product GUID", "Barcodes", "PLU code", "Manufacturers product code"]

    EDITABLE_COLS = [
        # Keep this small for POC, you can expand later
        "Status",
        "Promotion POS note",
        "Prompt read promo POS note",
        "Promo receipt note",
        "Product price",
        "Deal price $",
        "Min qty",
        "Max qty",
        "Min spend $",
        "Deal discount type",
        "Deal discount value",
        "Loyalty customers only",
        "Free product flag"
    ]

    # Show quick “schema” check
    missing_cols = [c for c in (DATE_COLS + ID_COLS) if c not in specials_df.columns]
    if missing_cols:
        st.warning("Some expected columns are missing (may be OK if your supplier format differs):")
        st.write(missing_cols)

    # Parse dates (NZ suppliers often use dd/mm/yyyy hh:mm)
    parsed = specials_df.copy()

    for c in DATE_COLS:
        if c in parsed.columns:
            parsed[c] = pd.to_datetime(parsed[c], errors="coerce", dayfirst=True)

    # Numeric parsing for key numeric fields if present
    NUM_COLS = ["Product price", "Deal price $", "Min qty", "Max qty", "Min spend $", "Deal discount value"]
    for c in NUM_COLS:
        if c in parsed.columns:
            parsed[c] = pd.to_numeric(
                parsed[c].astype("string").str.replace(r"[^\d.\-]", "", regex=True),
                errors="coerce"
            )

    # Validation flags
    issues = pd.DataFrame(index=parsed.index)
    if "Promotion start date" in parsed.columns and "Promotion end date" in parsed.columns:
        issues["Bad promo dates"] = parsed["Promotion start date"].isna() | parsed["Promotion end date"].isna() | (
            parsed["Promotion end date"] < parsed["Promotion start date"]
        )
    if "Deal start date" in parsed.columns and "Deal end date" in parsed.columns:
        issues["Bad deal dates"] = parsed["Deal start date"].isna() | parsed["Deal end date"].isna() | (
            parsed["Deal end date"] < parsed["Deal start date"]
        )

    # Missing identifiers: require at least one of GUID/Barcode/PLU/Mfr code
    present_id_cols = [c for c in ID_COLS if c in parsed.columns]
    if present_id_cols:
        has_any_id = False
        # build a boolean Series that is True when any ID col has content
        mask = pd.Series(False, index=parsed.index)
        for c in present_id_cols:
            mask = mask | (parsed[c].astype("string").fillna("").str.strip() != "")
        issues["Missing all IDs"] = ~mask

    # Rows with any issue
    any_issue = issues.any(axis=1) if len(issues.columns) else pd.Series(False, index=parsed.index)
    bad_rows = parsed.loc[any_issue].copy()

    st.subheader("Rows needing attention (invalid dates / missing IDs)")
    if bad_rows.empty:
        st.success("No obvious issues found ✅")
    else:
        preview_cols = [c for c in ["Promotion name", "Deal name", "Product name"] if c in bad_rows.columns]
        show_cols = preview_cols + [c for c in DATE_COLS if c in bad_rows.columns] + present_id_cols
        st.dataframe(bad_rows[show_cols].head(200), use_container_width=True, height=250)
        st.caption(f"{len(bad_rows)} rows flagged. (Shown first 200.)")

    # ---- Editable editor (POC) ----
    st.subheader("Edit specials (POC)")
    st.caption("Editing limited to selected fields for now. We’ll automate more later.")

    cols_to_edit = [c for c in EDITABLE_COLS if c in parsed.columns]
    cols_to_show = []
    for c in ["Promotion name", "Deal name", "Deal type", "Deal sub-type", "Product name"]:
        if c in parsed.columns:
            cols_to_show.append(c)
    cols_to_show += [c for c in DATE_COLS if c in parsed.columns]
    cols_to_show += [c for c in present_id_cols if c in parsed.columns]
    cols_to_show += cols_to_edit

    editable_view = parsed[cols_to_show].copy()

    edited = st.data_editor(
        editable_view,
        use_container_width=True,
        height=500,
        num_rows="dynamic",
        key="specials_editor"
    )

    # Stitch edits back into parsed (only editable columns + notes)
    for c in cols_to_edit:
        if c in edited.columns:
            parsed[c] = edited[c]

    # Export
    st.subheader("Export")
    out_name = st.text_input("Output file name", value="specials_cleaned.csv")
    csv_bytes = parsed.to_csv(index=False).encode("utf-8")
    st.download_button(
        "Download Specials CSV",
        data=csv_bytes,
        file_name=out_name,
        mime="text/csv"
    )