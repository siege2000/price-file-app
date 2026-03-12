# Price File Builder

A Windows desktop application for importing, processing, and exporting pharmacy supplier price files into the RxOne Access database (`suppliers.mdb`).

Suppliers send CSV and Excel price files to Healthsoft. This tool lets a user load those files, map the columns, clean up descriptions, preview changes against the live database, and upsert the data — all without writing any SQL.

---

## Features

### File Loading
- Open CSV or Excel (`.xlsx` / `.xls`) price files
- Multi-sheet Excel support: if the workbook has more than one sheet, a prompt appears to select which sheet to load
- Skip header rows (configurable per supplier)
- ID-like columns (Pharmacode, Barcode, SAP Code, PLU) are automatically coerced to strings to avoid `123456.0` formatting

### Column Mapping
- Map source file columns to price file fields: PLU, Pharmacode, Retail, Cost, Barcode/EAN, Supplier Code, Outer
- Templates (defined in `templates.json`) provide sensible column defaults for known suppliers (e.g. Toniq, RxOne)
- A **General** template uses fuzzy column-name matching for unknown suppliers

### TradeName / Description
- Combine multiple source columns into a single TradeName with a configurable separator
- Optional prefix (e.g. house-brand label)
- Title case conversion
- Unit normalisation (`75 g` → `75g`)
- 40-character limit enforced — rows exceeding this are highlighted; a bulk-truncate button is available
- Per-supplier shortening rules (find/replace, whole-word, case-insensitive, regex supported) stored in `supplier_rules.json` and editable from within the app

### Filtering
- De-duplicate rows by PLU (keeps last occurrence)
- **Exclude rows with no Barcode, Supplier Code and Pharmacode** checkbox (on by default) — removes lines that have no identifying key and cannot be matched in Access

### Supplier Settings
- Column mappings, TradeName settings, skip-rows, and template are saved per supplier in `supplier_settings.json`
- Up to **6 sheet layouts** can be saved per supplier (useful when a supplier sends a workbook with different product categories on separate sheets)
- A **Sheet layout** dropdown appears in the Supplier panel when more than one layout is saved

### Preview Changes
- Compares the built output against the existing Access records for the selected supplier
- Shows Status (`NEW` / `UPDATE`), old and new description, old and new cost/retail
- Rows with descriptions over 40 characters are flagged with ⚠
- **Background prefetch**: Access data for the selected supplier is fetched in the background as soon as the supplier is selected, so "Preview Changes" is near-instant
- Rows can be deleted from the preview grid by selecting them and pressing the **Delete** key

### Export / Upsert
- **Download CSV** — export the built output as a delimited file
- **Execute Upsert** — upsert into the Access `Details` table:
  - Match by Supplier Code first, Barcode as fallback
  - Insert new records; update existing ones
  - Batched SQL (`executemany`) — efficient even over a network drive
- **Replace All** — delete all existing records for the supplier and re-insert from scratch
- Optionally mark rows as `Updated` on write

### Database Support
- Primary target: Microsoft Access `suppliers.mdb` via ODBC
- SQLite alternative available for environments without Access ODBC drivers

---

## Requirements

### System
- **Windows** (required for Access ODBC driver)
- **Microsoft Access Database Engine** (32-bit or 64-bit, matching your Python install)
  - Download from Microsoft: [Access Database Engine](https://www.microsoft.com/en-us/download/details.aspx?id=54920)
- **Python 3.10+**

### Python packages
```
pandas
pyodbc
openpyxl
PyQt6
streamlit      # only needed for the Streamlit version
```

Install all dependencies:
```bash
pip install -r requirements.txt
pip install PyQt6
```

---

## Configuration

On first run, a `settings.ini` file is created automatically in the application directory:

```ini
[Access]
mdb_path = suppliers.mdb
password = YOUR_PASSWORD_HERE

[SQLite]
db_path = suppliers.sqlite
```

- **`mdb_path`** — path to `suppliers.mdb`. Can be absolute (`C:\Data\suppliers.mdb`) or a UNC path (`\\server\share\suppliers.mdb`).
- **`password`** — the Access database password.
- **`db_path`** — path to the SQLite fallback database (if used).

Edit `settings.ini` in any text editor, or use the Settings panel within the desktop app.

---

## Running the Desktop App

```bash
python qt_app.py
```

### Typical workflow

1. **Select a supplier** from the dropdown — the app immediately fetches their existing records from Access in the background.
2. **Open File** — browse to the supplier's CSV or Excel price file. If the workbook has multiple sheets, select which one to load.
3. **Choose a template** (or leave on General) — columns are mapped automatically where possible.
4. **Review the Source Data Preview** and adjust the column mappings and TradeName settings on the right panel.
5. **Build Output** — processes the file, applies TradeName rules, filters blank-ID rows, and shows the row count.
6. **Preview Changes** — compares the output against Access. NEW rows appear at the top in green; UPDATE rows in amber. Delete any unwanted rows with the Delete key.
7. **Fix descriptions** — edit `New_Description` cells directly in the grid, or use **Bulk Truncate** to clip everything to 40 characters.
8. **Save Settings** — saves the current column mappings for this supplier (and sheet) so they are restored next time.
9. **Execute Upsert** (or **Replace All**) — writes the data to Access.

---

## Running the Streamlit App (alternative)

A browser-based version is also available:

```bash
streamlit run app.py
```

This provides the same core price-file workflow in a web UI, accessible at `http://localhost:8501`.

---

## File Structure

| File | Purpose |
|------|---------|
| `qt_app.py` | PyQt6 desktop application (main entry point) |
| `app.py` | Streamlit web application |
| `access_helpers.py` | Access database read/write operations |
| `sqlite_helper.py` | SQLite equivalent of access_helpers |
| `file_helpers.py` | CSV/Excel loading and template parsing |
| `column_helpers.py` | Column mapping and default detection |
| `description_helpers.py` | TradeName building, cleaning, unit normalisation |
| `editing_helpers.py` | Money parsing, find/replace rule application |
| `config.py` | Read/write `settings.ini` |
| `specials_widget.py` | Specials file editor (separate workflow) |
| `templates.json` | Column mapping templates per supplier type |
| `supplier_settings.json` | Saved per-supplier column/layout settings |
| `supplier_rules.json` | Per-supplier TradeName shortening rules |
| `settings.ini` | Local database path and credentials (not committed) |
