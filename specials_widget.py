# specials_widget.py — PyQt6 UI for Import Specials (ported from VB6 frmImportSpecials)
"""
Wired to specials_helpers.py for all Access DB operations.
Run via qt_app.py (Specials File mode).
"""
from __future__ import annotations

import json
import os
from datetime import date, datetime

import pandas as pd
from PyQt6.QtCore import Qt, QAbstractTableModel, QModelIndex, QThread, pyqtSignal, QDate
from PyQt6.QtGui import QColor, QBrush
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QPushButton, QComboBox, QCheckBox, QLineEdit,
    QFrame, QGroupBox, QSplitter, QTableView, QScrollArea,
    QAbstractItemView, QHeaderView, QSizePolicy, QDateEdit,
    QMessageBox, QFileDialog,
)

from specials_helpers import (
    load_brands,
    get_special_id,
    save_specials,
    clear_old_specials,
    load_special_settings,
    save_special_settings,
)


# ── Grid columns (mirrors VB6 vsImport column order) ─────────────────────────
SPECIALS_COLUMNS = [
    "Barcode",
    "Description",
    "Chars",
    "MultiBuy Qty",
    "Special Price",
    "Get Lowest Free",
    "New",
    "Changed",
    "Invalid",
    "POS Note",
    "Multi Retail",
    "Flybuys",
    "Group ID",
    "Gift Barcode",
    "Gift Pharmacode",
    "Combo",
    "Msg Is Question",
    "Deal Name",
    "Secondary Group ID",
]

_EDITABLE_COLS = {
    "Description", "MultiBuy Qty", "Special Price",
    "POS Note", "Multi Retail",
}

# Field delimiter combo index → pandas sep
_FIELD_SEP = {0: ",", 1: "\t", 2: ";", 3: "|"}

# Known Excel column names (lowercase) → column combo key
# Used to auto-set the combos when the file has named headers
_EXCEL_TO_COMBO: dict[str, str] = {
    "product name":       "desc",
    "description":        "desc",
    "barcodes":           "barcode",
    "barcode":            "barcode",
    "ean":                "barcode",
    "deal price $":       "special",
    "special price":      "special",
    "price":              "special",
    "promotion pos note": "posnote",
    "pos note":           "posnote",
    "posnote":            "posnote",
}

# Known Excel column names → SpecialItems grid column name (direct mapping)
_EXCEL_TO_GRID: dict[str, str] = {
    "deal name":                      "Deal Name",
    "min qty":                        "MultiBuy Qty",
    "free product flag":              "Get Lowest Free",
    "product a or b":                 "Group ID",
    "prompt read promo pos note":     "Msg Is Question",
    "loyalty customers only":         "Flybuys",
    "multi retail":                   "Multi Retail",
    "multiretail":                    "Multi Retail",
    "gift barcode":                   "Gift Barcode",
    "gift pharmacode":                "Gift Pharmacode",
}


# ── Brand import mappings (brand_import_mappings.json) ───────────────────────
_MAPPINGS_PATH = os.path.join(os.path.dirname(__file__), "brand_import_mappings.json")

def _load_brand_mappings() -> dict:
    try:
        with open(_MAPPINGS_PATH, encoding="utf-8") as f:
            data = json.load(f)
        return {k: v for k, v in data.items() if not k.startswith("_")}
    except Exception:
        return {}

_BRAND_MAPPINGS: dict = _load_brand_mappings()


# ── Generic QThread worker ────────────────────────────────────────────────────
class _Worker(QThread):
    finished = pyqtSignal(object)
    error    = pyqtSignal(str)

    def __init__(self, fn, *args, **kwargs):
        super().__init__()
        self._fn     = fn
        self._args   = args
        self._kwargs = kwargs

    def run(self):
        try:
            self.finished.emit(self._fn(*self._args, **self._kwargs))
        except Exception as exc:
            self.error.emit(str(exc))


# ── Table model ───────────────────────────────────────────────────────────────
class SpecialsModel(QAbstractTableModel):
    def __init__(self, df: pd.DataFrame, parent=None):
        super().__init__(parent)
        self._df = df.reset_index(drop=True)

    def rowCount(self, parent=QModelIndex()):
        return len(self._df)

    def columnCount(self, parent=QModelIndex()):
        return len(self._df.columns)

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if not index.isValid():
            return None
        col = self._df.columns[index.column()]
        val = self._df.iloc[index.row(), index.column()]

        if role == Qt.ItemDataRole.DisplayRole:
            return "" if pd.isna(val) else str(val)

        if role == Qt.ItemDataRole.EditRole:
            return "" if pd.isna(val) else str(val)

        if role == Qt.ItemDataRole.BackgroundRole:
            inv  = str(self._df.at[index.row(), "Invalid"]  if "Invalid"  in self._df.columns else "")
            new_ = str(self._df.at[index.row(), "New"]      if "New"      in self._df.columns else "")
            chg  = str(self._df.at[index.row(), "Changed"]  if "Changed"  in self._df.columns else "")
            if inv  == "Y": return QBrush(QColor("#ffcccc"))
            if new_ == "Y": return QBrush(QColor("#ccccff"))
            if chg  == "Y": return QBrush(QColor("#ffc0c0"))

        if role == Qt.ItemDataRole.ForegroundRole:
            if col == "Invalid" and str(val) == "Y":
                return QBrush(QColor("#842029"))
            if col == "New" and str(val) == "Y":
                return QBrush(QColor("#0d47a1"))

        return None

    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
        if role != Qt.ItemDataRole.DisplayRole:
            return None
        if orientation == Qt.Orientation.Horizontal:
            return str(self._df.columns[section])
        return str(section + 1)

    def flags(self, index):
        f = super().flags(index)
        col = self._df.columns[index.column()]
        if col in _EDITABLE_COLS:
            f |= Qt.ItemFlag.ItemIsEditable
        return f

    def setData(self, index, value, role=Qt.ItemDataRole.EditRole):
        if role == Qt.ItemDataRole.EditRole and index.isValid():
            col = self._df.columns[index.column()]
            if col in _EDITABLE_COLS:
                self._df.at[index.row(), col] = value
                if col == "Description" and "Chars" in self._df.columns:
                    ci = list(self._df.columns).index("Chars")
                    self._df.at[index.row(), "Chars"] = len(str(value))
                    self.dataChanged.emit(self.index(index.row(), ci),
                                         self.index(index.row(), ci))
                self.dataChanged.emit(index, index)
                return True
        return False

    def dataFrame(self) -> pd.DataFrame:
        return self._df.copy()

    def addBlankRow(self):
        self.beginInsertRows(QModelIndex(), len(self._df), len(self._df))
        blank = pd.DataFrame([{c: "" for c in self._df.columns}])
        self._df = pd.concat([self._df, blank], ignore_index=True)
        self.endInsertRows()

    def removeRow(self, row: int, parent=QModelIndex()):
        if 0 <= row < len(self._df):
            self.beginRemoveRows(parent, row, row)
            self._df = self._df.drop(index=row).reset_index(drop=True)
            self.endRemoveRows()
            return True
        return False


def _empty_model() -> SpecialsModel:
    df = pd.DataFrame(columns=SPECIALS_COLUMNS)
    return SpecialsModel(df)


def _make_table_view() -> QTableView:
    v = QTableView()
    v.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
    v.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
    v.horizontalHeader().setStretchLastSection(True)
    v.verticalHeader().setDefaultSectionSize(22)
    v.setAlternatingRowColors(True)
    v.setEditTriggers(QAbstractItemView.EditTrigger.DoubleClicked)
    return v


def _legend_swatch(color: str, label: str) -> QHBoxLayout:
    row = QHBoxLayout()
    swatch = QFrame()
    swatch.setFixedSize(16, 16)
    swatch.setStyleSheet(f"background:{color}; border:1px solid #aaa;")
    row.addWidget(swatch)
    row.addWidget(QLabel(label))
    row.addStretch()
    return row


# ── Main widget ───────────────────────────────────────────────────────────────
class SpecialsWidget(QWidget):
    status_message = pyqtSignal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._model          = _empty_model()
        self._saved_settings : dict = {}   # last-loaded brand settings
        self._raw_file_df   : pd.DataFrame | None = None   # parsed file before mapping
        self._file_columns  : list[str] = []               # column headers from imported file
        self._worker        = None         # keep reference to prevent GC

        self._build_ui()
        self._load_brands()

    # ─── UI construction ──────────────────────────────────────────────────────
    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(6, 6, 6, 6)
        root.setSpacing(6)

        # ── Toolbar row ──────────────────────────────────────────────────────
        tool_bar = QHBoxLayout()
        for label, slot in [
            ("Load Grid…",          self._do_load_grid),
            ("Save Grid…",          self._do_save_grid),
            ("Group Specials",      self._stub),
            ("Ungroup Specials",    self._stub),
            ("Bulk Find & Replace", self._bulk_find_replace),
            ("Sort Invalid at Top", self._sort_invalid_top),
        ]:
            btn = QPushButton(label)
            btn.clicked.connect(slot)
            tool_bar.addWidget(btn)

        self._btn_log_invalid = QPushButton("Log Invalid Items…")
        self._btn_log_invalid.setStyleSheet(
            "background-color: #dc3545; color: white; padding: 3px 8px;"
        )
        self._btn_log_invalid.clicked.connect(self._log_invalid_items)
        tool_bar.addWidget(self._btn_log_invalid)
        tool_bar.addStretch()
        root.addLayout(tool_bar)

        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.HLine)
        root.addWidget(sep)

        # ── Splitter: grid (left) | settings (right) ─────────────────────────
        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setChildrenCollapsible(False)

        # LEFT — main import grid
        left_w = QWidget()
        left_v = QVBoxLayout(left_w)
        left_v.setContentsMargins(0, 0, 4, 0)
        left_v.addWidget(QLabel("<b>Import Grid</b>"))
        self._grid = _make_table_view()
        self._grid.setModel(self._model)
        self._grid.resizeColumnsToContents()
        left_v.addWidget(self._grid)
        self._lbl_lines = QLabel("")
        self._lbl_lines.setStyleSheet("font-size: 11px; color: #6c757d;")
        left_v.addWidget(self._lbl_lines)
        splitter.addWidget(left_w)

        # RIGHT — settings panel (scrollable)
        right_scroll = QScrollArea()
        right_scroll.setWidgetResizable(True)
        right_scroll.setFrameShape(QFrame.Shape.NoFrame)
        right_scroll.setMinimumWidth(270)
        right_w = QWidget()
        right_v = QVBoxLayout(right_w)
        right_v.setContentsMargins(4, 0, 0, 0)
        right_v.setSpacing(8)

        # Legend
        legend_grp = QGroupBox("Legend")
        legend_v = QVBoxLayout(legend_grp)
        legend_v.setSpacing(3)
        legend_v.addLayout(_legend_swatch("#ccccff", "New Cards"))
        legend_v.addLayout(_legend_swatch("#ffc0c0", "Changed"))
        legend_v.addLayout(_legend_swatch("#ffcccc", "Invalid"))
        right_v.addWidget(legend_grp)

        # Brand / supplier
        brand_grp = QGroupBox("Brand")
        brand_form = QGridLayout(brand_grp)
        brand_form.addWidget(QLabel("Brand:"), 0, 0)
        self._cmb_brand = QComboBox()
        self._cmb_brand.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self._cmb_brand.currentIndexChanged.connect(self._on_brand_changed)
        brand_form.addWidget(self._cmb_brand, 0, 1)
        brand_form.addWidget(QLabel("Password:"), 1, 0)
        self._txt_password = QLineEdit()
        self._txt_password.setEchoMode(QLineEdit.EchoMode.Password)
        brand_form.addWidget(self._txt_password, 1, 1)
        self._lbl_code = QLabel("")
        self._lbl_code.setWordWrap(True)
        self._lbl_code.setStyleSheet("font-size: 11px; color: #155724; padding: 2px;")
        brand_form.addWidget(self._lbl_code, 2, 0, 1, 2)
        right_v.addWidget(brand_grp)

        # File format
        fmt_grp = QGroupBox("File Format")
        fmt_form = QGridLayout(fmt_grp)
        fmt_form.addWidget(QLabel("Field delimiter:"), 0, 0)
        self._cmb_field_delim = QComboBox()
        self._cmb_field_delim.addItems(["Comma", "Tab", "Semicolon", "Pipe"])
        fmt_form.addWidget(self._cmb_field_delim, 0, 1)
        fmt_form.addWidget(QLabel("Record delimiter:"), 1, 0)
        self._cmb_record_delim = QComboBox()
        self._cmb_record_delim.addItems(["CRLF", "CR", "LF"])
        fmt_form.addWidget(self._cmb_record_delim, 1, 1)
        self._chk_ignore_first = QCheckBox("Ignore first line of file")
        fmt_form.addWidget(self._chk_ignore_first, 2, 0, 1, 2)
        right_v.addWidget(fmt_grp)

        # Column mapping
        col_grp = QGroupBox("Columns")
        col_form = QGridLayout(col_grp)
        hint = QLabel("Set which column number in your\nfile maps to each field (1 = first).")
        hint.setStyleSheet("font-size: 11px; color: #6c757d;")
        col_form.addWidget(hint, 0, 0, 1, 2)
        self._col_combos: dict[str, QComboBox] = {}
        for row_idx, (label, key) in enumerate([
            ("Description",          "desc"),
            ("Special Price",        "special"),
            ("Barcode / Pharmacode", "barcode"),
            ("POS Note (optional)",  "posnote"),
        ]):
            col_form.addWidget(QLabel(label), row_idx + 1, 0)
            cb = QComboBox()
            cb.addItem("(None)")
            for n in range(1, 21):
                cb.addItem(str(n))
            self._col_combos[key] = cb
            col_form.addWidget(cb, row_idx + 1, 1)

        self._btn_save_settings = QPushButton("Save Column Settings")
        self._btn_save_settings.clicked.connect(self._save_col_settings)
        col_form.addWidget(self._btn_save_settings, len(self._col_combos), 0, 1, 2)
        right_v.addWidget(col_grp)

        # PBL note
        pbl_lbl = QLabel("Import barcodes first for PBL from separate file")
        pbl_lbl.setWordWrap(True)
        pbl_lbl.setStyleSheet("font-size: 11px; color: #6c757d; padding: 2px 4px;")
        right_v.addWidget(pbl_lbl)

        # Row actions
        row_act_grp = QGroupBox("Row Actions")
        row_act_v = QVBoxLayout(row_act_grp)
        for label, slot in [
            ("Add Blank Line",       self._add_blank_line),
            ("Delete Selected Line", self._delete_selected),
            ("Clear POS Note",       self._clear_pos_note),
            ("Bulk Find && Replace", self._bulk_find_replace),
        ]:
            btn = QPushButton(label)
            btn.clicked.connect(slot)
            row_act_v.addWidget(btn)
        right_v.addWidget(row_act_grp)

        right_v.addStretch()
        right_scroll.setWidget(right_w)
        splitter.addWidget(right_scroll)
        splitter.setSizes([750, 270])
        root.addWidget(splitter, 3)

        # ── Bottom bar ────────────────────────────────────────────────────────
        sep2 = QFrame()
        sep2.setFrameShape(QFrame.Shape.HLine)
        root.addWidget(sep2)

        bottom = QGridLayout()
        bottom.setSpacing(6)

        bottom.addWidget(QLabel("Special Name:"), 0, 0)
        self._txt_special_name = QLineEdit()
        self._txt_special_name.setPlaceholderText("e.g. April 2026 Catalogue")
        bottom.addWidget(self._txt_special_name, 0, 1, 1, 3)

        bottom.addWidget(QLabel("Start Date:"), 1, 0)
        self._dt_start = QDateEdit()
        self._dt_start.setCalendarPopup(True)
        self._dt_start.setDate(QDate.currentDate())
        self._dt_start.setDisplayFormat("dd/MM/yyyy")
        bottom.addWidget(self._dt_start, 1, 1)

        bottom.addWidget(QLabel("Finish Date:"), 1, 2)
        self._dt_finish = QDateEdit()
        self._dt_finish.setCalendarPopup(True)
        self._dt_finish.setDate(QDate.currentDate().addMonths(1))
        self._dt_finish.setDisplayFormat("dd/MM/yyyy")
        bottom.addWidget(self._dt_finish, 1, 3)

        root.addLayout(bottom)

        # Action buttons
        btn_bar = QHBoxLayout()
        self._btn_import = QPushButton("&Import…")
        self._btn_import.setStyleSheet(
            "background-color: #0d6efd; color: white; font-weight: bold; padding: 4px 12px;"
        )
        self._btn_import.clicked.connect(self._do_import)

        self._btn_save = QPushButton("&Save to Database")
        self._btn_save.setStyleSheet(
            "background-color: #198754; color: white; font-weight: bold; padding: 4px 12px;"
        )
        self._btn_save.clicked.connect(self._do_save)

        self._btn_find_barcode = QPushButton("Open Stockcards (get barcode/pharmacode)…")
        self._btn_find_barcode.clicked.connect(self._stub)

        btn_bar.addWidget(self._btn_import)
        btn_bar.addWidget(self._btn_save)
        btn_bar.addWidget(self._btn_find_barcode)
        btn_bar.addStretch()
        root.addLayout(btn_bar)

    # ─── Stubs ────────────────────────────────────────────────────────────────
    def _stub(self):
        self.status_message.emit("Not yet implemented.")

    # ─── Brand loading ────────────────────────────────────────────────────────
    def _load_brands(self):
        self._cmb_brand.setEnabled(False)
        self._cmb_brand.clear()
        self._cmb_brand.addItem("Loading…")

        def _do():
            try:
                clear_old_specials()
            except Exception:
                pass
            return load_brands()

        self._worker = _Worker(_do)
        self._worker.finished.connect(self._on_brands_loaded)
        self._worker.error.connect(
            lambda e: self.status_message.emit(f"Error loading brands: {e}")
        )
        self._worker.start()

    def _on_brands_loaded(self, df: pd.DataFrame):
        self._cmb_brand.blockSignals(True)
        self._cmb_brand.clear()
        self._cmb_brand.addItem("(Select Brand)", 0)
        for _, row in df.iterrows():
            self._cmb_brand.addItem(str(row["BrandName"]), int(row["BrandID"]))
        self._cmb_brand.setEnabled(True)
        self._cmb_brand.blockSignals(False)
        self.status_message.emit(f"{len(df)} brand(s) loaded.")

    # ─── Brand selection ──────────────────────────────────────────────────────
    def _on_brand_changed(self):
        brand_id = self._cmb_brand.currentData() or 0
        self._lbl_code.setText("")
        if not brand_id:
            self._update_line_count()
            return
        try:
            settings = load_special_settings(brand_id)
            self._saved_settings = settings
            # Apply file-format settings
            fd = settings.get("field_delimiter", 0)
            self._cmb_field_delim.setCurrentIndex(
                min(fd, self._cmb_field_delim.count() - 1)
            )
            rd = settings.get("record_delimiter", 0)
            self._cmb_record_delim.setCurrentIndex(
                min(rd, self._cmb_record_delim.count() - 1)
            )
            self._chk_ignore_first.setChecked(
                bool(settings.get("ignore_first_line", False))
            )
            self._apply_saved_col_indices(settings)
            self._lbl_code.setText(f"Brand ID: {brand_id}")
        except Exception as e:
            self.status_message.emit(f"Error loading brand settings: {e}")
        self._update_line_count()

    def _current_brand_id(self) -> int:
        return int(self._cmb_brand.currentData() or 0)

    # ─── Column mapping helpers ────────────────────────────────────────────────
    def _apply_saved_col_indices(self, settings: dict):
        """
        Restore saved column indices to the combos.
        VB6 saved ListIndex (0-based field index). Our combos have:
          index 0 = "(None)", index 1 = "1" (field 0), index 2 = "2" (field 1), ...
        So: combo_index = saved_vb6_listindex + 1
        """
        mapping = {
            "desc":    settings.get("description_col", -1),
            "special": settings.get("special_col", -1),
            "barcode": settings.get("barcode_col", -1),
            "posnote": settings.get("posnote_col", -1),
        }
        for key, vb6_idx in mapping.items():
            cb = self._col_combos[key]
            if vb6_idx >= 0:
                combo_idx = vb6_idx + 1  # shift past "(None)"
                cb.setCurrentIndex(min(combo_idx, cb.count() - 1))
            else:
                cb.setCurrentIndex(0)  # "(None)"

    def _save_col_settings(self):
        """Persist current column mapping to SuppUDSpecialsSettings."""
        brand_id = self._current_brand_id()
        if not brand_id:
            QMessageBox.warning(self, "No Brand", "Please select a brand first.")
            return
        settings = {
            # Save as VB6 0-based ListIndex (currentIndex - 1 removes the "(None)" offset)
            "description_col":  self._col_combos["desc"].currentIndex() - 1,
            "special_col":      self._col_combos["special"].currentIndex() - 1,
            "barcode_col":      self._col_combos["barcode"].currentIndex() - 1,
            "posnote_col":      self._col_combos["posnote"].currentIndex() - 1,
            "field_delimiter":  self._cmb_field_delim.currentIndex(),
            "record_delimiter": self._cmb_record_delim.currentIndex(),
            "ignore_first_line": self._chk_ignore_first.isChecked(),
        }
        try:
            save_special_settings(brand_id, settings)
            self._saved_settings = settings
            self.status_message.emit("Column settings saved.")
        except Exception as e:
            QMessageBox.critical(self, "Save Error", str(e))

    # ─── File import ──────────────────────────────────────────────────────────
    def _do_import(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Open Specials File",
            "",
            "Spreadsheet / CSV (*.csv *.xls *.xlsx *.txt);;All files (*)",
        )
        if not path:
            return

        sep = _FIELD_SEP.get(self._cmb_field_delim.currentIndex(), ",")
        ignore_first = self._chk_ignore_first.isChecked()

        try:
            ext = os.path.splitext(path)[1].lower()
            if ext in (".xls", ".xlsx"):
                raw = pd.read_excel(path, header=0 if not ignore_first else None,
                                    dtype=str)
                if ignore_first:
                    raw.columns = [f"Column {i+1}" for i in range(len(raw.columns))]
            else:
                for enc in ("utf-8-sig", "cp1252", "latin-1"):
                    try:
                        raw = pd.read_csv(
                            path,
                            sep=sep,
                            header=0 if not ignore_first else None,
                            dtype=str,
                            encoding=enc,
                        )
                        break
                    except UnicodeDecodeError:
                        continue
                else:
                    raw = pd.read_csv(
                        path, sep=sep,
                        header=0 if not ignore_first else None,
                        dtype=str, encoding="latin-1",
                    )
                if ignore_first:
                    raw.columns = [f"Column {i+1}" for i in range(len(raw.columns))]

            raw = raw.fillna("")
        except Exception as e:
            QMessageBox.critical(self, "Import Error", f"Could not read file:\n{e}")
            return

        self._raw_file_df = raw
        # Auto-detect column mappings from named headers
        self._auto_detect_from_headers(list(raw.columns))
        # Auto-fill Special Name and dates from file data
        self._autofill_header_fields(raw)
        self.status_message.emit(
            f"File loaded: {len(raw)} row(s), {len(raw.columns)} column(s). Applying mapping…"
        )
        self._apply_mapping_to_grid()

    def _auto_detect_from_headers(self, columns: list[str]):
        """Set column combos from known Excel column name patterns."""
        col_lower = {c.lower(): i for i, c in enumerate(columns)}
        for excel_name, combo_key in _EXCEL_TO_COMBO.items():
            if excel_name in col_lower:
                field_idx = col_lower[excel_name]   # 0-based
                combo_idx = field_idx + 1            # shift past "(None)"
                cb = self._col_combos[combo_key]
                if combo_idx < cb.count():
                    cb.setCurrentIndex(combo_idx)

    def _autofill_header_fields(self, raw: pd.DataFrame):
        """Auto-populate Special Name and date pickers from file data (first row)."""
        if len(raw) == 0:
            return
        col_lower = {c.lower(): c for c in raw.columns}

        brand_name = self._cmb_brand.currentText().strip()
        bm_cols = _BRAND_MAPPINGS.get(brand_name, {}).get("columns", {})

        def first_col_val(*candidates) -> str:
            for c in candidates:
                actual = col_lower.get(c.lower())
                if actual:
                    v = str(raw[actual].iloc[0]).strip()
                    return "" if v.lower() == "nan" else v
            return ""

        # Special Name
        name_col = bm_cols.get("special_name", "")
        val = first_col_val(name_col, "Promotion name", "Special name", "SpecialName")
        if val:
            self._txt_special_name.setText(val)

        # Start date
        start_col = bm_cols.get("start_date", "")
        val = first_col_val(start_col, "Promotion start date", "Start date", "StartDate")
        if val:
            try:
                dt = datetime.fromisoformat(val.split(" ")[0])
                self._dt_start.setDate(QDate(dt.year, dt.month, dt.day))
            except Exception:
                pass

        # Finish/End date
        end_col = bm_cols.get("end_date", "")
        val = first_col_val(end_col, "Promotion end date", "Finish date", "End date", "FinishDate")
        if val:
            try:
                dt = datetime.fromisoformat(val.split(" ")[0])
                self._dt_finish.setDate(QDate(dt.year, dt.month, dt.day))
            except Exception:
                pass

    def _apply_mapping_to_grid(self):
        """Map raw file columns → specials grid using brand mapping, named columns, or combo positions."""
        if self._raw_file_df is None:
            QMessageBox.information(self, "No File", "Please import a file first.")
            return
        raw = self._raw_file_df

        # Look up brand mapping from JSON (by brand name)
        brand_name = self._cmb_brand.currentText().strip()
        bm = _BRAND_MAPPINGS.get(brand_name, {})
        bm_cols  = bm.get("columns", {})
        bm_rules = bm.get("rules", {})

        # Build a lowercase → actual-name lookup for named-header detection
        col_lower = {c.lower(): c for c in raw.columns}

        def by_name(*names) -> pd.Series | None:
            """Return first matching column by name (case-insensitive), or None."""
            for name in names:
                actual = col_lower.get(name.lower())
                if actual:
                    return raw[actual].astype(str).str.strip()
            return None

        def bm_col(role: str, *fallback_names) -> pd.Series | None:
            """Get column using brand mapping role, then fallback names."""
            mapped = bm_cols.get(role)
            candidates = ([mapped] if mapped else []) + list(fallback_names)
            return by_name(*candidates)

        def combo_col(key: str) -> pd.Series:
            """Return column by combo positional selection."""
            field_idx = self._col_combos[key].currentIndex() - 1
            if 0 <= field_idx < len(raw.columns):
                return raw.iloc[:, field_idx].astype(str).str.strip()
            return pd.Series([""] * len(raw))

        def first_not_none(*series) -> pd.Series:
            for s in series:
                if s is not None:
                    return s
            return pd.Series([""] * len(raw))

        # ── Source columns (brand mapping → known header names → combo position) ──
        desc_s           = first_not_none(bm_col("description",    "Product name"),       combo_col("desc"))
        barcode_s        = first_not_none(bm_col("barcode",        "Barcodes", "Barcode"), combo_col("barcode"))
        posnote_s        = first_not_none(bm_col("pos_note",       "Promotion POS note", "POS Note"), combo_col("posnote"))
        product_price_s  = bm_col("product_price",   "Product price")
        deal_disc_type_s = bm_col("discount_type",   "Deal discount type")
        deal_disc_val_s  = bm_col("discount_value",  "Deal discount value")
        min_qty_s        = bm_col("multibuy",         "Min qty")
        qty_breaks_s     = bm_col("multibuy_fallback","Quantity breaks")
        deal_name_col    = bm_col("deal_name",        "Deal name")
        deal_name_s      = first_not_none(deal_name_col, pd.Series([""] * len(raw)))

        def _v(series, i) -> str:
            return clean(str(series.iloc[i])) if series is not None else ""

        def clean(val: str) -> str:
            v = val.strip()
            return "" if v.lower() == "nan" else v

        # ── GroupID map: sequential number per unique Deal name ───────────────
        deal_group_map: dict[str, int] = {}
        if deal_name_col is not None:
            next_gid = 1
            for dn in deal_name_col:
                key_dn = clean(str(dn))
                if key_dn and key_dn not in deal_group_map:
                    deal_group_map[key_dn] = next_gid
                    next_gid += 1

        rows = []
        for i in range(len(raw)):
            # ── Barcode: first value if comma-separated ───────────────────────
            barcode = clean(str(barcode_s.iloc[i]))
            if "," in barcode:
                barcode = barcode.split(",")[0].strip()
            barcode = barcode.replace('"', '').replace(' ', '')

            desc = clean(str(desc_s.iloc[i]))[:40]

            # ── SpecialRetail / SpecialPercent ────────────────────────────────
            # Rule: if Product price has a value → SpecialRetail (dollar string)
            #       if Deal discount type = % → SpecialPercent (e.g. "50%")
            prod_price = _v(product_price_s, i)
            disc_type  = _v(deal_disc_type_s, i)
            disc_val   = _v(deal_disc_val_s, i)
            if prod_price:
                sp = prod_price
            elif disc_type == "%" and disc_val:
                try:
                    pct = float(disc_val)
                    sp = (str(int(pct)) if pct == int(pct) else str(pct)) + "%"
                except ValueError:
                    sp = disc_val + "%"
            else:
                sp = ""

            # ── MultiBuy: Min qty; fallback to Quantity breaks for X-for-$y ──
            mb_raw = _v(min_qty_s, i) or _v(qty_breaks_s, i)
            mb = ""
            if mb_raw:
                try:
                    mb = str(int(float(mb_raw)))
                except ValueError:
                    mb = mb_raw

            # ── LowestFree ────────────────────────────────────────────────────
            lf_rule = bm_rules.get("lowest_free", "disc_pct_and_multibuy_gt1")
            if lf_rule == "always_false":
                lowest_free = "N"
            elif lf_rule.startswith("from_column:"):
                lf_col = lf_rule.split(":", 1)[1]
                lf_val = _v(by_name(lf_col), i) if by_name(lf_col) is not None else ""
                lowest_free = "Y" if lf_val.upper() in ("Y", "YES", "TRUE", "1") else "N"
            else:  # disc_pct_and_multibuy_gt1 (default)
                lowest_free = "N"
                if disc_type == "%" and disc_val and mb:
                    try:
                        if float(disc_val) > 0 and float(mb) > 1:
                            lowest_free = "Y"
                    except ValueError:
                        pass

            # ── Flybuys ───────────────────────────────────────────────────────
            flybuys = "Y" if bm_rules.get("flybuys", False) else "N"

            # ── DealName and GroupID ──────────────────────────────────────────
            deal_name_val = clean(str(deal_name_s.iloc[i]))
            gid_rule = bm_rules.get("group_id", "sequential_by_deal_name")
            if gid_rule == "sequential_by_deal_name":
                group_id = str(deal_group_map.get(deal_name_val, "")) if deal_group_map else ""
            elif gid_rule.startswith("from_column:"):
                gid_col = gid_rule.split(":", 1)[1]
                group_id = _v(by_name(gid_col), i) if by_name(gid_col) is not None else ""
            else:
                group_id = ""

            row = {c: "" for c in SPECIALS_COLUMNS}
            row["Barcode"]         = barcode
            row["Description"]     = desc
            row["Chars"]           = str(len(desc))
            row["MultiBuy Qty"]    = mb
            row["Special Price"]   = sp
            row["Get Lowest Free"] = lowest_free
            row["New"]             = "Y"
            row["Changed"]         = "N"
            row["Invalid"]         = "N" if barcode else "Y"
            row["POS Note"]        = clean(str(posnote_s.iloc[i]))
            row["Flybuys"]         = flybuys
            row["Group ID"]        = group_id
            row["Deal Name"]       = deal_name_val
            rows.append(row)

        df = pd.DataFrame(rows, columns=SPECIALS_COLUMNS)
        self._model = SpecialsModel(df)
        self._grid.setModel(self._model)
        self._grid.resizeColumnsToContents()
        self._update_line_count()
        invalid_count = (df["Invalid"] == "Y").sum()
        self.status_message.emit(
            f"Grid loaded: {len(df)} row(s), {invalid_count} invalid (no barcode)."
        )

    # ─── Save to Access ───────────────────────────────────────────────────────
    def _do_save(self):
        special_name = self._txt_special_name.text().strip()
        if not special_name:
            QMessageBox.warning(self, "Validation", "Please enter a Special Name.")
            return

        brand_id = self._current_brand_id()
        if not brand_id:
            QMessageBox.warning(self, "Validation", "Please select a Brand.")
            return

        grid_df = self._model.dataFrame()
        if grid_df.empty:
            QMessageBox.warning(self, "Validation", "Grid is empty — nothing to save.")
            return

        qstart  = self._dt_start.date()
        qfinish = self._dt_finish.date()
        start   = date(qstart.year(),  qstart.month(),  qstart.day())
        finish  = date(qfinish.year(), qfinish.month(), qfinish.day())
        password = self._txt_password.text().strip()

        self._btn_save.setEnabled(False)
        self.status_message.emit("Saving…")

        def _do():
            sid = get_special_id(special_name, brand_id, create_new=True)
            return save_specials(grid_df, sid, start, finish, password,
                                 clear_existing=True)

        self._worker = _Worker(_do)
        self._worker.finished.connect(self._on_save_done)
        self._worker.error.connect(self._on_save_error)
        self._worker.start()

    def _on_save_done(self, saved: int):
        self._btn_save.setEnabled(True)
        self.status_message.emit(f"Saved {saved} row(s) to database.")
        QMessageBox.information(self, "Saved", f"{saved} row(s) saved successfully.")
        # Clear the import grid after a successful save
        self._model = _empty_model()
        self._grid.setModel(self._model)
        self._raw_file_df = None

    def _on_save_error(self, msg: str):
        self._btn_save.setEnabled(True)
        self.status_message.emit(f"Save error: {msg}")
        QMessageBox.critical(self, "Save Error", msg)

    # ─── Load / Save grid CSV ─────────────────────────────────────────────────
    def _do_load_grid(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Load Grid", "",
            "Spreadsheet / CSV (*.csv *.xls *.xlsx);;All files (*)",
        )
        if not path:
            return
        try:
            ext = os.path.splitext(path)[1].lower()
            if ext in (".xls", ".xlsx"):
                df = pd.read_excel(path, dtype=str).fillna("")
            else:
                for enc in ("utf-8-sig", "cp1252", "latin-1"):
                    try:
                        df = pd.read_csv(path, dtype=str, encoding=enc).fillna("")
                        break
                    except UnicodeDecodeError:
                        continue
                else:
                    df = pd.read_csv(path, dtype=str, encoding="latin-1").fillna("")
            # Ensure all expected columns exist
            for col in SPECIALS_COLUMNS:
                if col not in df.columns:
                    df[col] = ""
            df = df[SPECIALS_COLUMNS]
            self._model = SpecialsModel(df)
            self._grid.setModel(self._model)
            self._grid.resizeColumnsToContents()
            self._update_line_count()
            self.status_message.emit(f"Grid loaded from {os.path.basename(path)}.")
        except Exception as e:
            QMessageBox.critical(self, "Load Error", str(e))

    def _do_save_grid(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Save Grid", "specials_grid.csv", "CSV files (*.csv)"
        )
        if not path:
            return
        try:
            self._model.dataFrame().to_csv(path, index=False)
            self.status_message.emit(f"Grid saved to {os.path.basename(path)}.")
        except Exception as e:
            QMessageBox.critical(self, "Save Error", str(e))

    # ─── Bulk Find & Replace ──────────────────────────────────────────────────
    def _bulk_find_replace(self):
        from PyQt6.QtWidgets import QDialog, QDialogButtonBox
        dlg = QDialog(self)
        dlg.setWindowTitle("Bulk Find & Replace")
        dlg.setMinimumWidth(360)
        layout = QGridLayout(dlg)
        layout.addWidget(QLabel("Column:"), 0, 0)
        cmb_col = QComboBox()
        cmb_col.addItems(list(_EDITABLE_COLS))
        layout.addWidget(cmb_col, 0, 1)
        layout.addWidget(QLabel("Find:"), 1, 0)
        txt_find = QLineEdit()
        layout.addWidget(txt_find, 1, 1)
        layout.addWidget(QLabel("Replace with:"), 2, 0)
        txt_replace = QLineEdit()
        layout.addWidget(txt_replace, 2, 1)
        chk_exact = QCheckBox("Exact match only")
        layout.addWidget(chk_exact, 3, 0, 1, 2)
        btns = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        btns.accepted.connect(dlg.accept)
        btns.rejected.connect(dlg.reject)
        layout.addWidget(btns, 4, 0, 1, 2)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        col_name   = cmb_col.currentText()
        find_text  = txt_find.text()
        repl_text  = txt_replace.text()
        exact      = chk_exact.isChecked()
        if not find_text:
            return
        df = self._model._df
        if col_name not in df.columns:
            return
        col_idx = list(df.columns).index(col_name)
        count = 0
        for r in range(len(df)):
            val = str(df.at[r, col_name])
            if exact:
                if val == find_text:
                    df.at[r, col_name] = repl_text
                    count += 1
            else:
                if find_text in val:
                    df.at[r, col_name] = val.replace(find_text, repl_text)
                    count += 1
        if count:
            top_left = self._model.index(0, col_idx)
            bot_right = self._model.index(len(df) - 1, col_idx)
            self._model.dataChanged.emit(top_left, bot_right)
        self.status_message.emit(f"Replaced {count} cell(s) in '{col_name}'.")

    # ─── Grid row operations ──────────────────────────────────────────────────
    def _add_blank_line(self):
        self._model.addBlankRow()
        self._update_line_count()

    def _delete_selected(self):
        rows = sorted(
            {idx.row() for idx in self._grid.selectedIndexes()},
            reverse=True,
        )
        for r in rows:
            self._model.removeRow(r)
        self._update_line_count()

    def _clear_pos_note(self):
        df = self._model.dataFrame()
        if "POS Note" not in df.columns:
            return
        selected_rows = {idx.row() for idx in self._grid.selectedIndexes()}
        if not selected_rows:
            return
        col_idx = list(df.columns).index("POS Note")
        for r in selected_rows:
            self._model._df.at[r, "POS Note"] = ""
            self._model.dataChanged.emit(
                self._model.index(r, col_idx),
                self._model.index(r, col_idx),
            )
        self.status_message.emit(f"Cleared POS Note on {len(selected_rows)} row(s).")

    def _sort_invalid_top(self):
        df = self._model.dataFrame()
        if "Invalid" not in df.columns:
            return
        df["_sort"] = df["Invalid"].apply(lambda v: 0 if str(v) == "Y" else 1)
        df = df.sort_values("_sort").drop(columns=["_sort"]).reset_index(drop=True)
        self._model = SpecialsModel(df)
        self._grid.setModel(self._model)
        self._update_line_count()

    # ─── Log invalid items ────────────────────────────────────────────────────
    def _log_invalid_items(self):
        df = self._model.dataFrame()
        if "Invalid" not in df.columns:
            self.status_message.emit("No grid data.")
            return

        invalid_df = df[df["Invalid"].str.upper() == "Y"].copy()
        if invalid_df.empty:
            QMessageBox.information(self, "Log Invalid Items", "No invalid items in the current grid.")
            return

        # Determine reason for each invalid row
        def _reason(row) -> str:
            reasons = []
            if not str(row.get("Barcode", "")).strip():
                reasons.append("No barcode")
            if not str(row.get("Description", "")).strip():
                reasons.append("No description")
            return "; ".join(reasons) if reasons else "Unknown"

        invalid_df = invalid_df.copy()
        invalid_df.insert(0, "Reason", invalid_df.apply(_reason, axis=1))
        invalid_df.insert(0, "Special Name", self._txt_special_name.text().strip())
        invalid_df.insert(0, "Brand", self._cmb_brand.currentText().strip())
        invalid_df.insert(0, "Logged At", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

        # Keep only the useful columns for the report
        report_cols = ["Logged At", "Brand", "Special Name", "Reason",
                       "Barcode", "Description", "Deal Name", "POS Note"]
        report_cols = [c for c in report_cols if c in invalid_df.columns]
        report_df = invalid_df[report_cols]

        # Default filename: <SpecialName>_invalid_<date>.csv
        special_slug = (self._txt_special_name.text().strip() or "specials").replace(" ", "_")
        default_name = f"{special_slug}_invalid_{datetime.now().strftime('%Y%m%d')}.csv"

        path, _ = QFileDialog.getSaveFileName(
            self, "Save Invalid Items Log", default_name,
            "CSV files (*.csv);;All files (*)"
        )
        if not path:
            return

        # Append if file already exists (running same import multiple times)
        write_header = not os.path.exists(path)
        report_df.to_csv(path, mode="a", index=False, header=write_header, encoding="utf-8-sig")

        self.status_message.emit(
            f"Logged {len(report_df)} invalid item(s) to {os.path.basename(path)}."
        )
        QMessageBox.information(
            self, "Log Saved",
            f"{len(report_df)} invalid item(s) written to:\n{path}"
        )

    def _update_line_count(self):
        count = self._model.rowCount()
        self._lbl_lines.setText(f"{count} line(s)")
        self.status_message.emit(f"{count} row(s) in grid.")
