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
settings1_mdb_path = Settings1.mdb

[SQLite]
db_path = suppliers.sqlite

[MSSQL]
server = SUPPORT0630\HEALTHSOFTLOTS
database = lotssql
```

- **`mdb_path`** — path to `suppliers.mdb`. Can be absolute (`C:\Data\suppliers.mdb`) or a UNC path (`\\server\share\suppliers.mdb`).
- **`password`** — the Access database password.
- **`settings1_mdb_path`** — path to `Settings1.mdb` (find/replace rules source).
- **`db_path`** — path to the SQLite fallback database (if used).
- **`server`** — SQL Server instance for the Healthsoft stock database (used by the Specials barcode lookup).
- **`database`** — database name on that SQL Server instance.

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
| `config.py` | Read/write `settings.ini` |
| `price_file_app.spec` | PyInstaller build spec (see Distribution below) |
| **`db/`** | |
| `db/access_helpers.py` | Access database read/write operations |
| `db/sqlite_helper.py` | SQLite equivalent of access_helpers |
| **`processing/`** | |
| `processing/file_helpers.py` | CSV/Excel loading and template parsing |
| `processing/column_helpers.py` | Column mapping and default detection |
| `processing/description_helpers.py` | TradeName building, cleaning, unit normalisation |
| `processing/editing_helpers.py` | Money parsing, find/replace rule application |
| **`specials/`** | |
| `specials/specials_widget.py` | Specials file import UI (separate workflow) |
| `specials/specials_helpers.py` | Specials Access DB read/write operations |
| `specials/stockcard_dialog.py` | MSSQL stock barcode lookup dialog |
| `specials/brand_import_mappings.json` | Per-brand column mapping rules for specials import |
| **`data/`** | |
| `data/templates.json` | Column mapping templates per supplier type |
| `data/supplier_settings.json` | Saved per-supplier column/layout settings |
| `data/supplier_rules.json` | Per-supplier TradeName shortening rules |
| `data/replacements.csv` | Global description replacement reference |
| `settings.ini` | Local database paths and credentials (not committed) |

---

## Distribution — Building a Windows Executable

The app can be packaged into a self-contained Windows executable using [PyInstaller](https://pyinstaller.org). No Python installation is required on the target machine.

### Prerequisites (build machine only)

Activate the virtual environment, then install PyInstaller:

```bash
# Windows — activate venv first
venv\Scripts\activate

pip install pyinstaller
```

> **Note:** PyInstaller is not in `requirements.txt` because it is only needed on the build machine, not the target machine.

### Build

```bash
pyinstaller price_file_app.spec
```

Output is placed in `dist\PriceFileApp\`.

> **Tip:** Always use `pyinstaller price_file_app.spec` — do **not** run `pyinstaller qt_app.py` as that bypasses the bundled data files and hidden-import settings in the spec.

### Assemble the distribution folder

PyInstaller does not copy writable data files. After the build, copy these into `dist\PriceFileApp\`:

```
dist\PriceFileApp\
  PriceFileApp.exe               ← built by PyInstaller
  _internal\                     ← PyInstaller internals (do not modify)
  settings.ini                   ← copy from project root, edit for the target site
  brand_import_mappings.json     ← copy from specials\ — edit to add new brands without rebuilding
  templates.json                 ← copy from data\
  supplier_rules.json            ← copy from data\
  supplier_settings.json         ← copy from data\  (auto-created on first save if absent)
  replacements.csv               ← copy from data\  (optional — user can also load via file dialog)
```

Edit `settings.ini` for the target site before shipping:

```ini
[Access]
mdb_path = \\server\share\Suppliers.mdb
password = YOUR_PASSWORD

[MSSQL]
server = SUPPORT0630\HEALTHSOFTLOTS
database = lotssql
```

### Target machine requirements

| Requirement | Notes |
|---|---|
| Windows 10 / 11 | 64-bit |
| Microsoft Access Database Engine (64-bit) | Required for `.mdb` connectivity — download from Microsoft |
| SQL Server ODBC Driver | The `{SQL Server}` driver ships with Windows; usually already present |

### Notes

| | |
|---|---|
| **Startup time** | First launch may take 3–5 seconds while Windows loads the bundled DLLs — normal behaviour |
| **Antivirus** | Some AV software flags unsigned PyInstaller executables. Code-signing with a certificate resolves this |
| **Icon** | Uncomment the `icon=` line in `price_file_app.spec` and supply a `.ico` file to brand the executable |
| **Rebuilding** | Always use `pyinstaller price_file_app.spec` — do not run `pyinstaller qt_app.py` as that bypasses the spec settings |
