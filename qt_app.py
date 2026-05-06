# qt_app.py — PyQt6 Windows UI for Price File Builder
"""
Windows-native equivalent of app.py (Streamlit).
Run with:  python qt_app.py
Requires:  pip install PyQt6
"""
from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Optional

import pandas as pd
from PyQt6.QtCore import (
    Qt, QAbstractTableModel, QModelIndex, QThread, pyqtSignal,
)
from PyQt6.QtGui import QColor, QBrush
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QGridLayout, QLabel, QPushButton, QComboBox, QCheckBox, QSpinBox,
    QFileDialog, QTableView, QTableWidget, QTableWidgetItem, QSplitter,
    QGroupBox, QScrollArea, QMessageBox, QStatusBar, QListWidget,
    QAbstractItemView, QHeaderView, QFrame, QRadioButton, QButtonGroup,
    QStackedWidget, QSizePolicy, QStyledItemDelegate, QLineEdit, QDialog,
)

# ── project helpers ──────────────────────────────────────────────────────────
from processing.file_helpers import load_templates
from processing.column_helpers import safe_str, tpl_field_default, tpl_desc_defaults
from processing.description_helpers import combine_columns, normalize_units, clean_description
from processing.editing_helpers import parse_money, apply_replacements
from db.access_helpers import (
    MAX_DESC_LEN, load_suppliers, load_supplier_details,
    build_upsert_preview, upsert_details, replace_all_details,
)
from db import sqlite_helper
import config
from specials.specials_widget import SpecialsWidget

import os as _os
SUPPLIER_SETTINGS_FILE = _os.path.join(config.base_path(), "supplier_settings.json")
SUPPLIER_RULES_FILE    = _os.path.join(config.base_path(), "supplier_rules.json")


# ── Pandas ↔ Qt model ────────────────────────────────────────────────────────
class PandasModel(QAbstractTableModel):
    """Read-only table model backed by a DataFrame."""

    def __init__(
        self,
        df: pd.DataFrame,
        currency_cols: list[str] | None = None,
        parent=None,
    ):
        super().__init__(parent)
        self._df = df.reset_index(drop=True)
        self._currency_cols = set(currency_cols or [])

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
            if col in self._currency_cols and pd.notna(val):
                try:
                    return f"${float(val):.2f}"
                except (ValueError, TypeError):
                    pass
            return "" if pd.isna(val) else str(val)

        if role == Qt.ItemDataRole.EditRole:
            return "" if pd.isna(val) else str(val)

        if role == Qt.ItemDataRole.ForegroundRole:
            if col == "Status":
                if str(val) == "NEW":
                    return QBrush(QColor("#155724"))
                if str(val) == "UPDATE":
                    return QBrush(QColor("#856404"))
                if str(val) == "UNCHANGED":
                    return QBrush(QColor("#6c757d"))
            if col == "Valid":
                return QBrush(QColor("#155724") if str(val) == "✓" else QColor("#842029"))

        if role == Qt.ItemDataRole.BackgroundRole:
            if col == "Status":
                if str(val) == "NEW":
                    return QBrush(QColor("#d4edda"))
                if str(val) == "UPDATE":
                    return QBrush(QColor("#fff3cd"))
                if str(val) == "UNCHANGED":
                    return QBrush(QColor("#f8f9fa"))
            if col == "Chars":
                try:
                    if int(val) > 40:
                        return QBrush(QColor("#f8d7da"))
                except (ValueError, TypeError):
                    pass

        return None

    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
        if role != Qt.ItemDataRole.DisplayRole:
            return None
        if orientation == Qt.Orientation.Horizontal:
            return str(self._df.columns[section])
        return str(section + 1)

    def sort(self, column: int, order=Qt.SortOrder.AscendingOrder):
        self.layoutAboutToBeChanged.emit()
        col_name = self._df.columns[column]
        ascending = (order == Qt.SortOrder.AscendingOrder)
        self._df = self._df.sort_values(by=col_name, ascending=ascending, kind="stable").reset_index(drop=True)
        self.layoutChanged.emit()

    def dataFrame(self) -> pd.DataFrame:
        return self._df.copy()

    def remove_rows(self, row_indices: list[int]):
        """Remove rows by integer position and notify views."""
        self._df = self._df.drop(self._df.index[row_indices]).reset_index(drop=True)
        self.layoutChanged.emit()


class EditablePandasModel(PandasModel):
    """PandasModel that allows editing specific columns."""

    def __init__(self, df, editable_cols: list[str], currency_cols=None, parent=None):
        super().__init__(df, currency_cols=currency_cols, parent=parent)
        self._editable_cols = set(editable_cols)

    def flags(self, index):
        f = super().flags(index)
        if self._df.columns[index.column()] in self._editable_cols:
            f |= Qt.ItemFlag.ItemIsEditable
        return f

    def setData(self, index, value, role=Qt.ItemDataRole.EditRole):
        if role == Qt.ItemDataRole.EditRole and index.isValid():
            col = self._df.columns[index.column()]
            if col in self._editable_cols:
                self._df.at[index.row(), col] = value
                cols_list = list(self._df.columns)
                new_len = len(str(value))
                # keep CharCount in sync (Output Editor legacy)
                if col == "TradeName" and "CharCount" in cols_list:
                    ci = cols_list.index("CharCount")
                    self._df.at[index.row(), "CharCount"] = new_len
                    self.dataChanged.emit(self.index(index.row(), ci), self.index(index.row(), ci))
                # keep Chars + Valid in sync (preview table)
                if col == "New_Description":
                    if "Chars" in cols_list:
                        ci = cols_list.index("Chars")
                        self._df.at[index.row(), "Chars"] = new_len
                        self.dataChanged.emit(self.index(index.row(), ci), self.index(index.row(), ci))
                    if "Valid" in cols_list:
                        vi = cols_list.index("Valid")
                        self._df.at[index.row(), "Valid"] = "✓" if new_len <= 40 else "⚠"
                        self.dataChanged.emit(self.index(index.row(), vi), self.index(index.row(), vi))
                self.dataChanged.emit(index, index)
                return True
        return False


# ── Delegate: cursor-in-place editing (no select-all on activation) ───────────
class CursorPlacingDelegate(QStyledItemDelegate):
    """On double-click, opens the editor with the existing text and cursor
    placed at the click position — no select-all wipe."""

    def setEditorData(self, editor, index):
        if isinstance(editor, QLineEdit):
            value = index.data(Qt.ItemDataRole.EditRole) or ""
            editor.setText(value)
            editor.deselect()
            editor.end(False)   # cursor to end, no selection

    def setModelData(self, editor, model, index):
        if isinstance(editor, QLineEdit):
            model.setData(index, editor.text(), Qt.ItemDataRole.EditRole)


# ── Worker threads ────────────────────────────────────────────────────────────
class PrefetchWorker(QThread):
    """Fetches existing Access Details for a supplier in the background."""
    done = pyqtSignal(int, object)  # (supplier_id, DataFrame)
    error = pyqtSignal(str)

    def __init__(self, supplier_id: int):
        super().__init__()
        self.supplier_id = supplier_id

    def run(self):
        try:
            df = load_supplier_details(self.supplier_id)
            self.done.emit(self.supplier_id, df)
        except Exception as e:
            self.error.emit(str(e))


class PreviewWorker(QThread):
    done = pyqtSignal(object, object)   # (preview_df, stats_dict)
    error = pyqtSignal(str)

    def __init__(self, export_df: pd.DataFrame, supplier_id: int,
                 existing_df: pd.DataFrame | None = None):
        super().__init__()
        self.export_df = export_df
        self.supplier_id = supplier_id
        self.existing_df = existing_df

    def run(self):
        try:
            preview, stats = build_upsert_preview(
                self.export_df, self.supplier_id, existing_df=self.existing_df
            )
            self.done.emit(preview, stats)
        except Exception as e:
            self.error.emit(str(e))


class UpsertWorker(QThread):
    done = pyqtSignal(int, int)
    progress = pyqtSignal(int, int)   # (rows_done, total)
    error = pyqtSignal(str)

    def __init__(self, export_df: pd.DataFrame, supplier_id: int, mark_updated: bool,
                 existing_df=None):
        super().__init__()
        self.export_df = export_df
        self.supplier_id = supplier_id
        self.mark_updated = mark_updated
        self.existing_df = existing_df

    def run(self):
        try:
            updated, inserted = upsert_details(
                self.export_df, self.supplier_id,
                mark_updated=self.mark_updated,
                existing_df=self.existing_df,
                progress_callback=lambda done, total: self.progress.emit(done, total),
            )
            self.done.emit(inserted, updated)
        except Exception as e:
            self.error.emit(str(e))


class ReplaceWorker(QThread):
    done = pyqtSignal(int)        # rows inserted
    error = pyqtSignal(str)
    progress = pyqtSignal(int, int)  # (done, total)

    def __init__(self, export_df: pd.DataFrame, supplier_id: int, mark_updated: bool):
        super().__init__()
        self.export_df = export_df
        self.supplier_id = supplier_id
        self.mark_updated = mark_updated

    def run(self):
        try:
            inserted = replace_all_details(
                self.export_df, self.supplier_id, mark_updated=self.mark_updated,
                progress_callback=self.progress.emit,
            )
            self.done.emit(inserted)
        except Exception as e:
            self.error.emit(str(e))


class SqliteUpsertWorker(QThread):
    done = pyqtSignal(int, int)   # (updated, inserted)
    error = pyqtSignal(str)

    def __init__(self, export_df: pd.DataFrame, supplier_id: int, mark_updated: bool,
                 supplier_name: str = "", supplier_code: str = ""):
        super().__init__()
        self.export_df = export_df
        self.supplier_id = supplier_id
        self.mark_updated = mark_updated
        self.supplier_name = supplier_name
        self.supplier_code = supplier_code

    def run(self):
        try:
            sqlite_helper.ensure_supplier_by_id(
                self.supplier_id, self.supplier_name, self.supplier_code
            )
            updated, inserted = sqlite_helper.upsert_details(
                self.export_df, self.supplier_id, mark_updated=self.mark_updated
            )
            self.done.emit(updated, inserted)
        except Exception as e:
            self.error.emit(str(e))


class SqliteReplaceWorker(QThread):
    done = pyqtSignal(int)   # rows inserted
    error = pyqtSignal(str)

    def __init__(self, export_df: pd.DataFrame, supplier_id: int, mark_updated: bool,
                 supplier_name: str = "", supplier_code: str = ""):
        super().__init__()
        self.export_df = export_df
        self.supplier_id = supplier_id
        self.mark_updated = mark_updated
        self.supplier_name = supplier_name
        self.supplier_code = supplier_code

    def run(self):
        try:
            sqlite_helper.ensure_supplier_by_id(
                self.supplier_id, self.supplier_name, self.supplier_code
            )
            inserted = sqlite_helper.replace_all_details(
                self.export_df, self.supplier_id, mark_updated=self.mark_updated
            )
            self.done.emit(inserted)
        except Exception as e:
            self.error.emit(str(e))


# ── Helpers ───────────────────────────────────────────────────────────────────
def load_file_by_path(path: str, sheet_name: Optional[str] = None) -> pd.DataFrame:
    """Load CSV or Excel by file-system path. Pass sheet_name to select a specific Excel sheet."""
    p = path.lower()
    if p.endswith(".csv"):
        return pd.read_csv(path)
    if p.endswith(".xlsx") or p.endswith(".xls"):
        return pd.read_excel(path, sheet_name=sheet_name if sheet_name is not None else 0)
    raise ValueError("Unsupported file type — must be .csv or .xlsx")


def get_excel_sheets(path: str) -> list[str]:
    """Return the list of sheet names for an Excel workbook."""
    try:
        return pd.ExcelFile(path).sheet_names
    except Exception:
        return []


def make_table_view() -> QTableView:
    v = QTableView()
    v.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
    v.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
    v.horizontalHeader().setStretchLastSection(True)
    v.verticalHeader().setDefaultSectionSize(22)
    v.setAlternatingRowColors(True)
    return v


class DeletableTableView(QTableView):
    """Preview table that removes selected rows when the Delete key is pressed."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.horizontalHeader().setStretchLastSection(True)
        self.verticalHeader().setDefaultSectionSize(22)
        self.setAlternatingRowColors(True)

    def keyPressEvent(self, event):
        if event.key() == Qt.Key.Key_Delete:
            model = self.model()
            if model is not None and hasattr(model, "remove_rows"):
                rows = sorted(
                    {idx.row() for idx in self.selectedIndexes()},
                    reverse=True,
                )
                if rows:
                    model.remove_rows(rows)
        else:
            super().keyPressEvent(event)


# ── Supplier Rules dialog ─────────────────────────────────────────────────────
class SupplierRulesDialog(QDialog):
    """Editable table of from→to shortening rules for a single supplier."""

    def __init__(self, supplier_name: str, rules: list[dict], parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Shortening Rules — {supplier_name}")
        self.resize(560, 400)
        self._accepted_rules: list[dict] = []

        layout = QVBoxLayout(self)
        layout.setSpacing(6)

        # Instructions
        info = QLabel(
            "Each row is a find-and-replace applied to the description during <b>Build Output</b>.<br>"
            "Matching is whole-word, case-insensitive. "
            "Prefix <b>From</b> with <code>regex:</code> to use a regular expression."
        )
        info.setWordWrap(True)
        info.setStyleSheet("color: #495057; font-size: 11px; padding: 4px;")
        layout.addWidget(info)

        # Table
        self._table = QTableWidget(0, 2)
        self._table.setHorizontalHeaderLabels(["From", "To"])
        self._table.horizontalHeader().setStretchLastSection(True)
        self._table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self._table.verticalHeader().setDefaultSectionSize(24)
        self._table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self._table.setAlternatingRowColors(True)
        layout.addWidget(self._table)

        # Row buttons
        row_btns = QHBoxLayout()
        btn_add = QPushButton("+ Add Row")
        btn_add.clicked.connect(self._add_row)
        btn_del = QPushButton("Delete Selected")
        btn_del.clicked.connect(self._delete_selected)
        row_btns.addWidget(btn_add)
        row_btns.addWidget(btn_del)
        row_btns.addStretch()
        layout.addLayout(row_btns)

        sep = QFrame(); sep.setFrameShape(QFrame.Shape.HLine)
        layout.addWidget(sep)

        # Dialog buttons
        dlg_btns = QHBoxLayout()
        self._btn_apply = QPushButton("Save && Apply to Preview")
        self._btn_apply.setStyleSheet(
            "background-color: #0d6efd; color: white; font-weight: bold; padding: 4px 14px;"
        )
        self._btn_apply.clicked.connect(self._on_apply)
        btn_cancel = QPushButton("Cancel")
        btn_cancel.clicked.connect(self.reject)
        dlg_btns.addStretch()
        dlg_btns.addWidget(btn_cancel)
        dlg_btns.addWidget(self._btn_apply)
        layout.addLayout(dlg_btns)

        # Populate with existing rules
        for r in rules:
            self._add_row(r.get("from", ""), r.get("to", ""))

    def _add_row(self, from_val: str = "", to_val: str = ""):
        row = self._table.rowCount()
        self._table.insertRow(row)
        self._table.setItem(row, 0, QTableWidgetItem(from_val))
        self._table.setItem(row, 1, QTableWidgetItem(to_val))
        if not from_val:
            self._table.editItem(self._table.item(row, 0))

    def _delete_selected(self):
        rows = sorted({idx.row() for idx in self._table.selectedIndexes()}, reverse=True)
        for r in rows:
            self._table.removeRow(r)

    def _on_apply(self):
        rules = []
        for r in range(self._table.rowCount()):
            from_item = self._table.item(r, 0)
            to_item = self._table.item(r, 1)
            f = (from_item.text() if from_item else "").strip()
            t = (to_item.text() if to_item else "").strip()
            if f:
                rules.append({"from": f, "to": t})
        self._accepted_rules = rules
        self.accept()

    def get_rules(self) -> list[dict]:
        return self._accepted_rules


# ── Sheet Picker dialog ───────────────────────────────────────────────────────
class SheetPickerDialog(QDialog):
    """Prompt the user to select an Excel sheet when a workbook has multiple sheets."""

    def __init__(self, sheet_names: list[str], parent=None):
        super().__init__(parent)
        self.setWindowTitle("Select Excel Sheet")
        self.resize(340, 160)
        self._selected: Optional[str] = None

        layout = QVBoxLayout(self)
        layout.setSpacing(8)
        layout.addWidget(QLabel("This workbook has multiple sheets.\nSelect the sheet to load:"))

        self._cmb = QComboBox()
        self._cmb.addItems(sheet_names)
        layout.addWidget(self._cmb)

        btns = QHBoxLayout()
        btn_ok = QPushButton("Load Sheet")
        btn_ok.setStyleSheet(
            "background-color: #0d6efd; color: white; font-weight: bold; padding: 4px 14px;"
        )
        btn_ok.clicked.connect(self._on_ok)
        btn_cancel = QPushButton("Cancel")
        btn_cancel.clicked.connect(self.reject)
        btns.addStretch()
        btns.addWidget(btn_cancel)
        btns.addWidget(btn_ok)
        layout.addLayout(btns)

    def _on_ok(self):
        self._selected = self._cmb.currentText()
        self.accept()

    def selected_sheet(self) -> Optional[str]:
        return self._selected


# ── Price File widget ─────────────────────────────────────────────────────────
class PriceFileWidget(QWidget):
    status_message = pyqtSignal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._file_path: Optional[str] = None
        self._rules_path: Optional[str] = None
        self._df: Optional[pd.DataFrame] = None
        self._export_df: Optional[pd.DataFrame] = None
        self._suppliers_df: Optional[pd.DataFrame] = None
        self._templates: dict = {}
        self._preview_model: Optional[EditablePandasModel] = None
        self._upsert_skipped_rows: Optional[pd.DataFrame] = None
        self._preview_worker: Optional[PreviewWorker] = None
        self._upsert_worker: Optional[UpsertWorker] = None
        self._last_preview_stats: Optional[dict] = None
        self._replace_worker: Optional[ReplaceWorker] = None
        self._sqlite_upsert_worker: Optional[SqliteUpsertWorker] = None
        self._sqlite_replace_worker: Optional[SqliteReplaceWorker] = None
        self._prefetch_worker: Optional[PrefetchWorker] = None
        self._prefetched_details: Optional[pd.DataFrame] = None
        self._prefetch_supplier_id: Optional[int] = None
        self._current_sheet: Optional[str] = None
        self._excel_sheets: list[str] = []

        self._build_ui()
        self._load_templates()
        self._load_suppliers()

    # ─── UI construction ─────────────────────────────────────────────────────
    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(6, 6, 6, 6)
        root.setSpacing(6)

        # Top bar: open file + template + skip rows
        top = QHBoxLayout()
        self._btn_open = QPushButton("Open File…")
        self._btn_open.clicked.connect(self._open_file)
        self._lbl_file = QLabel("No file loaded")
        self._lbl_file.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        top.addWidget(self._btn_open)
        top.addWidget(self._lbl_file)
        top.addWidget(QLabel("Template:"))
        self._cmb_template = QComboBox()
        self._cmb_template.setMinimumWidth(180)
        self._cmb_template.currentIndexChanged.connect(self._on_template_changed)
        top.addWidget(self._cmb_template)
        self._spin_skip = QSpinBox()
        self._spin_skip.setRange(0, 1000)
        self._spin_skip.setPrefix("Skip: ")
        self._spin_skip.setSuffix(" rows")
        self._spin_skip.valueChanged.connect(self._on_skip_changed)
        top.addWidget(self._spin_skip)
        self._btn_clear = QPushButton("Clear / Start Fresh")
        self._btn_clear.setToolTip("Reset loaded file and all derived state")
        self._btn_clear.clicked.connect(self._clear_state)
        top.addWidget(self._btn_clear)
        btn_exit = QPushButton("Exit")
        btn_exit.setStyleSheet("color: #842029; font-weight: bold;")
        btn_exit.clicked.connect(QApplication.instance().quit)
        top.addWidget(btn_exit)
        root.addLayout(top)

        # ── Main splitter: left (source + options) | right (mapping) ─────────
        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setChildrenCollapsible(False)

        # LEFT panel
        left_w = QWidget()
        left_v = QVBoxLayout(left_w)
        left_v.setContentsMargins(0, 0, 4, 0)

        left_v.addWidget(QLabel("<b>Source Data Preview</b>"))
        self._source_table = make_table_view()
        left_v.addWidget(self._source_table, 3)

        # Export options group
        exp_grp = QGroupBox("Export Options")
        exp_form = QGridLayout(exp_grp)
        exp_form.addWidget(QLabel("CSV delimiter:"), 0, 0)
        self._cmb_delim = QComboBox()
        self._cmb_delim.addItems([",", ";", "Tab"])
        exp_form.addWidget(self._cmb_delim, 0, 1)
        exp_form.addWidget(QLabel("Price decimals:"), 1, 0)
        self._cmb_decimals = QComboBox()
        self._cmb_decimals.addItems(["2", "3", "0"])
        exp_form.addWidget(self._cmb_decimals, 1, 1)
        self._chk_blank_plu = QCheckBox("Allow blank PLU")
        self._chk_blank_plu.setChecked(True)
        exp_form.addWidget(self._chk_blank_plu, 2, 0, 1, 2)
        self._chk_excl_no_id = QCheckBox("Exclude rows with no Barcode, Supplier Code and Pharmacode")
        self._chk_excl_no_id.setChecked(True)
        exp_form.addWidget(self._chk_excl_no_id, 3, 0, 1, 2)
        left_v.addWidget(exp_grp)

        # Rules group
        rules_grp = QGroupBox("TradeName Shortening Rules (optional)")
        rules_v = QVBoxLayout(rules_grp)
        rules_row = QHBoxLayout()
        self._btn_rules = QPushButton("Load rules CSV…")
        self._btn_rules.clicked.connect(self._open_rules)
        self._lbl_rules = QLabel("No rules file")
        rules_row.addWidget(self._btn_rules)
        rules_row.addWidget(self._lbl_rules, 1)
        rules_v.addLayout(rules_row)
        opts_row = QHBoxLayout()
        self._chk_whole_word = QCheckBox("Whole word")
        self._chk_whole_word.setChecked(True)
        self._chk_case = QCheckBox("Case-insensitive")
        self._chk_case.setChecked(True)
        opts_row.addWidget(self._chk_whole_word)
        opts_row.addWidget(self._chk_case)
        opts_row.addStretch()
        rules_v.addLayout(opts_row)
        rules_v.addWidget(QLabel("(Apply Rules button is in the action bar below)"))
        left_v.addWidget(rules_grp)

        splitter.addWidget(left_w)

        # RIGHT panel (scrollable)
        right_scroll = QScrollArea()
        right_scroll.setWidgetResizable(True)
        right_scroll.setFrameShape(QFrame.Shape.NoFrame)
        right_scroll.setMinimumWidth(260)
        right_w = QWidget()
        right_v = QVBoxLayout(right_w)
        right_v.setContentsMargins(4, 0, 0, 0)

        # Supplier
        sup_grp = QGroupBox("Supplier")
        sup_v = QVBoxLayout(sup_grp)
        self._cmb_supplier = QComboBox()
        self._cmb_supplier.currentIndexChanged.connect(self._on_supplier_changed)
        sup_v.addWidget(self._cmb_supplier)
        # Sheet layout selector — shown only when this supplier has >1 saved sheet layout
        sheet_cfg_row = QHBoxLayout()
        self._lbl_sheet_config = QLabel("Sheet layout:")
        sheet_cfg_row.addWidget(self._lbl_sheet_config)
        self._cmb_sheet_config = QComboBox()
        self._cmb_sheet_config.setMinimumWidth(130)
        self._cmb_sheet_config.currentIndexChanged.connect(self._on_sheet_config_changed)
        sheet_cfg_row.addWidget(self._cmb_sheet_config)
        sheet_cfg_row.addStretch()
        sup_v.addLayout(sheet_cfg_row)
        self._lbl_sheet_config.setVisible(False)
        self._cmb_sheet_config.setVisible(False)
        rules_row = QHBoxLayout()
        self._lbl_supplier_rules = QLabel("")
        self._lbl_supplier_rules.setStyleSheet("font-size: 11px; color: #6c757d; padding: 1px 2px;")
        self._lbl_supplier_rules.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        self._btn_edit_rules = QPushButton("Edit Rules…")
        self._btn_edit_rules.setMaximumWidth(90)
        self._btn_edit_rules.clicked.connect(self._open_rules_dialog)
        rules_row.addWidget(self._lbl_supplier_rules)
        rules_row.addWidget(self._btn_edit_rules)
        sup_v.addLayout(rules_row)
        right_v.addWidget(sup_grp)

        # Column mapping
        map_grp = QGroupBox("Map Columns → Price File Fields")
        map_form = QGridLayout(map_grp)
        self._col_combos: dict[str, QComboBox] = {}
        for row_idx, (label, key) in enumerate([
            ("PLU", "plu"),
            ("Pharmacode", "pharmacode"),
            ("Retail *", "retail"),
            ("Cost *", "cost"),
            ("Barcode/EAN (optional)", "barcode"),
            ("Supplier Code (optional)", "sc"),
            ("Outer (optional)", "outer"),
        ]):
            map_form.addWidget(QLabel(label), row_idx, 0)
            cb = QComboBox()
            cb.addItem("(None)")
            self._col_combos[key] = cb
            map_form.addWidget(cb, row_idx, 1)
        right_v.addWidget(map_grp)

        # TradeName settings
        tn_grp = QGroupBox("TradeName Settings")
        tn_v = QVBoxLayout(tn_grp)
        tn_v.addWidget(QLabel("Columns (Ctrl+click to select, ▲▼ to reorder):"))
        desc_row = QHBoxLayout()
        self._lst_desc_cols = QListWidget()
        self._lst_desc_cols.setSelectionMode(QAbstractItemView.SelectionMode.MultiSelection)
        self._lst_desc_cols.setMaximumHeight(130)
        desc_row.addWidget(self._lst_desc_cols)
        btn_col = QVBoxLayout()
        btn_col.setSpacing(2)
        self._btn_desc_up = QPushButton("▲")
        self._btn_desc_up.setMaximumWidth(28)
        self._btn_desc_up.setToolTip("Move selected column up")
        self._btn_desc_up.clicked.connect(self._move_desc_col_up)
        self._btn_desc_down = QPushButton("▼")
        self._btn_desc_down.setMaximumWidth(28)
        self._btn_desc_down.setToolTip("Move selected column down")
        self._btn_desc_down.clicked.connect(self._move_desc_col_down)
        btn_col.addWidget(self._btn_desc_up)
        btn_col.addWidget(self._btn_desc_down)
        btn_col.addStretch()
        desc_row.addLayout(btn_col)
        tn_v.addLayout(desc_row)
        sep_row = QHBoxLayout()
        sep_row.addWidget(QLabel("Separator:"))
        self._cmb_sep = QComboBox()
        self._cmb_sep.addItems([" ", " - ", "-", " | ", ", "])
        sep_row.addWidget(self._cmb_sep)
        sep_row.addStretch()
        tn_v.addLayout(sep_row)
        prefix_row = QHBoxLayout()
        prefix_row.addWidget(QLabel("Prefix:"))
        self._txt_prefix = QLineEdit()
        self._txt_prefix.setPlaceholderText("e.g. HH  (leave blank for none)")
        self._txt_prefix.setMaximumWidth(160)
        self._txt_prefix.textChanged.connect(self._update_desc_preview)
        prefix_row.addWidget(self._txt_prefix)
        prefix_row.addStretch()
        tn_v.addLayout(prefix_row)
        self._chk_title_case = QCheckBox("Title Case")
        self._chk_norm_units = QCheckBox("Normalize units (75 g → 75g)")
        self._chk_norm_units.setChecked(True)
        tn_v.addWidget(self._chk_title_case)
        tn_v.addWidget(self._chk_norm_units)

        self._lbl_desc_preview = QLabel("(load a file and select columns)")
        self._lbl_desc_preview.setWordWrap(True)
        self._lbl_desc_preview.setStyleSheet(
            "color: #495057; font-size: 11px; padding: 4px 6px; "
            "background: #f8f9fa; border: 1px solid #dee2e6; border-radius: 3px;"
        )
        tn_v.addWidget(self._lbl_desc_preview)

        # Connect signals for live preview update
        self._lst_desc_cols.itemSelectionChanged.connect(self._update_desc_preview)
        self._cmb_sep.currentIndexChanged.connect(self._update_desc_preview)
        self._chk_title_case.stateChanged.connect(self._update_desc_preview)
        self._chk_norm_units.stateChanged.connect(self._update_desc_preview)

        right_v.addWidget(tn_grp)
        right_v.addStretch()

        right_scroll.setWidget(right_w)
        splitter.addWidget(right_scroll)
        splitter.setSizes([650, 300])
        root.addWidget(splitter, 3)

        # ── Action bar + single working table ────────────────────────────────
        sep1 = QFrame(); sep1.setFrameShape(QFrame.Shape.HLine)
        root.addWidget(sep1)

        action_bar = QHBoxLayout()

        self._btn_build = QPushButton("Build Output")
        self._btn_build.setStyleSheet("font-weight: bold;")
        self._btn_build.clicked.connect(self._build_output)
        self._btn_truncate = QPushButton("Bulk Truncate (40 chars)")
        self._btn_truncate.clicked.connect(self._bulk_truncate)
        self._btn_apply_rules_inline = QPushButton("Apply Rules")
        self._btn_apply_rules_inline.clicked.connect(self._apply_rules)
        self._btn_export_csv = QPushButton("Save CSV…")
        self._btn_export_csv.clicked.connect(self._export_csv)
        action_bar.addWidget(self._btn_build)
        action_bar.addWidget(self._btn_truncate)
        action_bar.addWidget(self._btn_apply_rules_inline)
        action_bar.addWidget(self._btn_export_csv)

        div = QFrame(); div.setFrameShape(QFrame.Shape.VLine)
        div.setFrameShadow(QFrame.Shadow.Sunken)
        action_bar.addWidget(div)

        self._chk_mark_updated = QCheckBox("Mark rows as Updated")
        self._chk_mark_updated.setChecked(True)
        self._btn_preview_upsert = QPushButton("Preview Changes")
        self._btn_preview_upsert.clicked.connect(self._preview_upsert)
        self._btn_exec_upsert = QPushButton("Execute Upsert")
        self._btn_exec_upsert.setStyleSheet(
            "background-color: #0d6efd; color: white; font-weight: bold; padding: 4px 12px;"
        )
        self._btn_exec_upsert.clicked.connect(self._exec_upsert)
        self._btn_replace_all = QPushButton("Replace All in Access")
        self._btn_replace_all.setToolTip(
            "DELETE all existing records for this supplier, then INSERT all current rows fresh"
        )
        self._btn_replace_all.setStyleSheet(
            "background-color: #dc3545; color: white; font-weight: bold; padding: 4px 12px;"
        )
        self._btn_replace_all.clicked.connect(self._exec_replace)
        action_bar.addWidget(self._chk_mark_updated)
        action_bar.addWidget(self._btn_preview_upsert)
        action_bar.addWidget(self._btn_exec_upsert)
        action_bar.addWidget(self._btn_replace_all)

        div2 = QFrame(); div2.setFrameShape(QFrame.Shape.VLine)
        div2.setFrameShadow(QFrame.Shadow.Sunken)
        action_bar.addWidget(div2)

        self._btn_sqlite_upsert = QPushButton("Upsert → SQLite")
        self._btn_sqlite_upsert.setToolTip(
            "Upsert current rows into suppliers.sqlite (same logic as Access upsert)"
        )
        self._btn_sqlite_upsert.setStyleSheet(
            "background-color: #198754; color: white; font-weight: bold; padding: 4px 12px;"
        )
        self._btn_sqlite_upsert.clicked.connect(self._exec_sqlite_upsert)
        self._btn_sqlite_replace = QPushButton("Replace All → SQLite")
        self._btn_sqlite_replace.setToolTip(
            "DELETE all existing SQLite rows for this supplier, then INSERT current rows"
        )
        self._btn_sqlite_replace.setStyleSheet(
            "background-color: #6f42c1; color: white; font-weight: bold; padding: 4px 12px;"
        )
        self._btn_sqlite_replace.clicked.connect(self._exec_sqlite_replace)
        action_bar.addWidget(self._btn_sqlite_upsert)
        action_bar.addWidget(self._btn_sqlite_replace)
        btn_browse_sqlite = QPushButton("Browse SQLite…")
        btn_browse_sqlite.setToolTip("Open the SQLite database browser / editor")
        btn_browse_sqlite.setStyleSheet("padding: 4px 12px;")
        btn_browse_sqlite.clicked.connect(self._open_sqlite_browser)
        action_bar.addWidget(btn_browse_sqlite)
        root.addLayout(action_bar)

        self._lbl_preview_info = QLabel("")
        self._lbl_preview_info.setStyleSheet("padding: 2px 4px;")
        root.addWidget(self._lbl_preview_info)

        sort_bar = QHBoxLayout()
        self._btn_sort_new = QPushButton("Sort New at Top")
        self._btn_sort_new.clicked.connect(self._sort_new_top)
        self._btn_sort_invalid = QPushButton("Sort Invalid at Top")
        self._btn_sort_invalid.clicked.connect(self._sort_invalid_top)
        sort_bar.addWidget(self._btn_sort_new)
        sort_bar.addWidget(self._btn_sort_invalid)
        sort_bar.addStretch()
        self._btn_diagnostics = QPushButton("Match Diagnostics…")
        self._btn_diagnostics.setToolTip("Show how import codes/barcodes compare to Access records")
        self._btn_diagnostics.clicked.connect(self._show_diagnostics)
        sort_bar.addWidget(self._btn_diagnostics)
        root.addLayout(sort_bar)

        self._preview_table = DeletableTableView()
        root.addWidget(self._preview_table, 2)

    # ─── Template + supplier loading ──────────────────────────────────────────
    def _load_templates(self):
        try:
            self._templates = load_templates(_os.path.join(config.base_path(), "templates.json"))
            self._cmb_template.blockSignals(True)
            self._cmb_template.clear()
            self._cmb_template.addItems(list(self._templates.keys()))
            self._cmb_template.blockSignals(False)
        except Exception as e:
            QMessageBox.warning(self, "Templates", f"Could not load templates.json:\n{e}")

    def _load_suppliers(self):
        try:
            self._suppliers_df = load_suppliers()
            self._suppliers_df["Label"] = self._suppliers_df.apply(
                lambda r: (
                    f'{r["SupplierName"]} ({r["SupplierCode"]})'
                    if str(r["SupplierCode"] or "").strip()
                    else str(r["SupplierName"])
                ),
                axis=1,
            )
            self._cmb_supplier.clear()
            self._cmb_supplier.addItem("— select supplier —")
            self._cmb_supplier.addItems(self._suppliers_df["Label"].tolist())
        except Exception:
            btn = QMessageBox.warning(
                self,
                "Suppliers file not found",
                "Suppliers file not found.\n\nOpen Settings to set the correct path.",
                QMessageBox.StandardButton.Ok,
            )
            if btn == QMessageBox.StandardButton.Ok:
                dlg = SettingsDialog(self)
                dlg.exec()
                self._load_suppliers()

    # ─── File open / reload ───────────────────────────────────────────────────
    def _open_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Open Price File", "", "CSV / Excel (*.csv *.xlsx *.xls)"
        )
        if not path:
            return
        self._file_path = path
        self._current_sheet = None
        self._excel_sheets = []
        # Detect sheets for Excel files
        if path.lower().endswith((".xlsx", ".xls")):
            sheets = get_excel_sheets(path)
            if len(sheets) > 1:
                dlg = SheetPickerDialog(sheets, self)
                if dlg.exec() != QDialog.DialogCode.Accepted:
                    self._file_path = None
                    return
                self._current_sheet = dlg.selected_sheet()
            self._excel_sheets = sheets
        self._lbl_file.setText(Path(path).name)
        self._reload_data()
        # Invalidate the Access cache so the next Preview/Upsert fetches fresh data
        self._prefetched_details = None
        self._prefetch_supplier_id = None
        supplier_id = self._get_supplier_id()
        if supplier_id is not None:
            self._start_prefetch(supplier_id)

    def _on_skip_changed(self):
        if self._file_path:
            self._reload_data()

    def _on_template_changed(self):
        if self._df is not None:
            self._populate_column_combos()

    def _reload_data(self):
        try:
            df = load_file_by_path(self._file_path, sheet_name=self._current_sheet)
            skip = self._spin_skip.value()
            if skip > 0:
                df = df.iloc[skip:].reset_index(drop=True)
            # Coerce ID-like columns to string
            for c in df.columns:
                if c.strip().lower() in {
                    "pharmacode", "pharma code", "barcode",
                    "barcode/ean", "ean", "sap code", "plu",
                }:
                    df[c] = (
                        df[c].astype("string")
                        .str.replace(r"\.0$", "", regex=True)
                        .str.strip()
                    )
            self._df = df
            self._source_table.setModel(PandasModel(df))
            self._source_table.resizeColumnsToContents()
            self._populate_column_combos()
            self.status_message.emit(
                f"Loaded {len(df)} rows from {Path(self._file_path).name}"
            )
        except Exception as e:
            QMessageBox.critical(self, "Error loading file", str(e))

    # ─── Column combo population from template ────────────────────────────────
    def _populate_column_combos(self):
        if self._df is None:
            return
        tpl = self._templates.get(self._cmb_template.currentText(), {})
        fields = tpl.get("fields", {})
        cols = ["(None)"] + self._df.columns.tolist()

        defaults = {
            "plu": tpl_field_default(self._df, fields.get("plu", {})),
            "pharmacode": tpl_field_default(self._df, fields.get("pharmacode", {})),
            "retail": tpl_field_default(self._df, fields.get("retail", {})),
            "cost": tpl_field_default(self._df, fields.get("cost", {})),
            "barcode": tpl_field_default(self._df, fields.get("barcode", {})),
            "sc": tpl_field_default(self._df, fields.get("sc", {})),
            "outer": tpl_field_default(self._df, fields.get("outer", {})),
        }
        for key, cb in self._col_combos.items():
            cb.blockSignals(True)
            cb.clear()
            cb.addItems(cols)
            d = defaults.get(key)
            if d and d in cols:
                cb.setCurrentIndex(cols.index(d))
            cb.blockSignals(False)

        # TradeName columns list
        desc_cfg = tpl.get("description", tpl.get("tradename", {}))
        desc_defaults = tpl_desc_defaults(self._df, desc_cfg)
        self._lst_desc_cols.blockSignals(True)
        self._lst_desc_cols.clear()
        for col in self._df.columns.tolist():
            self._lst_desc_cols.addItem(col)
        for i in range(self._lst_desc_cols.count()):
            item = self._lst_desc_cols.item(i)
            if item and item.text() in desc_defaults:
                item.setSelected(True)
        self._lst_desc_cols.blockSignals(False)

        # Export options
        opts = tpl.get("options", {})
        delim_map = {",": 0, ";": 1, "\t": 2}
        self._cmb_delim.setCurrentIndex(delim_map.get(opts.get("csv_delimiter", ","), 0))
        dec_idx = self._cmb_decimals.findText(str(opts.get("price_decimals", 2)))
        if dec_idx >= 0:
            self._cmb_decimals.setCurrentIndex(dec_idx)
        self._chk_blank_plu.setChecked(bool(opts.get("allow_blank_plu", True)))

        # TradeName settings
        sep = desc_cfg.get("separator", " ")
        sep_idx = self._cmb_sep.findText(sep)
        if sep_idx >= 0:
            self._cmb_sep.setCurrentIndex(sep_idx)
        self._chk_title_case.setChecked(bool(desc_cfg.get("title_case", False)))
        self._chk_norm_units.setChecked(bool(desc_cfg.get("normalize_units", True)))

        # Overlay any saved supplier-specific settings on top of template defaults
        supplier_id = self._get_supplier_id()
        if supplier_id is not None:
            self._apply_supplier_settings(supplier_id)
        else:
            self._update_desc_preview()

    # ─── Description preview ──────────────────────────────────────────────────
    def _update_desc_preview(self):
        selected = [
            self._lst_desc_cols.item(i).text()
            for i in range(self._lst_desc_cols.count())
            if self._lst_desc_cols.item(i).isSelected()
        ]
        if not selected:
            self._lbl_desc_preview.setText("(no columns selected)")
            return

        sep = self._cmb_sep.currentText()
        prefix = self._txt_prefix.text().strip()
        prefix_display = f"<b>[{prefix}]</b> + " if prefix else ""
        cols_line = prefix_display + f" <b>{sep}</b> ".join(selected)

        if self._df is not None and len(self._df) > 0:
            try:
                from processing.description_helpers import combine_columns, normalize_units, clean_description
                sample = combine_columns(self._df.head(1), selected, sep=sep).iloc[0]
                sample = clean_description(pd.Series([sample])).iloc[0]
                if self._chk_title_case.isChecked():
                    sample = sample.title()
                if self._chk_norm_units.isChecked():
                    sample = normalize_units(pd.Series([sample])).iloc[0]
                if prefix:
                    sample = prefix + " " + sample
                if len(sample) > 55:
                    sample = sample[:52] + "…"
                char_count = len(sample)
                color = "#842029" if char_count > 40 else "#155724"
                self._lbl_desc_preview.setText(
                    f"{cols_line}<br>"
                    f"<span style='color:{color};'>Sample: {sample} ({char_count} chars)</span>"
                )
            except Exception:
                self._lbl_desc_preview.setText(cols_line)
        else:
            self._lbl_desc_preview.setText(cols_line)

    def _move_desc_col_up(self):
        selected = [i for i in range(self._lst_desc_cols.count())
                    if self._lst_desc_cols.item(i).isSelected()]
        if not selected or selected[0] == 0:
            return
        for i in selected:
            item = self._lst_desc_cols.takeItem(i)
            self._lst_desc_cols.insertItem(i - 1, item)
            item.setSelected(True)
        self._update_desc_preview()

    def _move_desc_col_down(self):
        selected = [i for i in range(self._lst_desc_cols.count())
                    if self._lst_desc_cols.item(i).isSelected()]
        if not selected or selected[-1] == self._lst_desc_cols.count() - 1:
            return
        for i in reversed(selected):
            item = self._lst_desc_cols.takeItem(i)
            self._lst_desc_cols.insertItem(i + 1, item)
            item.setSelected(True)
        self._update_desc_preview()

    # ─── Rules ───────────────────────────────────────────────────────────────
    def _open_rules(self):
        path, _ = QFileDialog.getOpenFileName(self, "Open Rules CSV", "", "CSV (*.csv)")
        if path:
            self._rules_path = path
            self._lbl_rules.setText(Path(path).name)

    def _apply_rules(self):
        if self._preview_model is None:
            QMessageBox.warning(self, "Rules", "Run Preview Changes first.")
            return
        if not self._rules_path:
            QMessageBox.warning(self, "Rules", "Load a rules CSV file first.")
            return
        try:
            rules_df = pd.read_csv(self._rules_path)
            rules_df.columns = [c.strip().lower() for c in rules_df.columns]
            if "from" not in rules_df.columns or "to" not in rules_df.columns:
                raise ValueError("Rules file must have 'from' and 'to' columns.")
            rules_df = rules_df[["from", "to"]].dropna()
            rules_df["from"] = rules_df["from"].astype("string").str.strip()
            rules_df["to"] = rules_df["to"].astype("string").str.strip()
            rules_df = rules_df[rules_df["from"] != ""]

            df = self._preview_model.dataFrame()
            df["New_Description"], count = apply_replacements(
                df["New_Description"],
                rules_df,
                whole_word=self._chk_whole_word.isChecked(),
                case_insensitive=self._chk_case.isChecked(),
            )
            df["Chars"] = df["New_Description"].str.len()
            df["Valid"] = df["Chars"].apply(lambda x: "✓" if x <= 40 else "⚠")
            self._show_preview(df)
            self.status_message.emit(f"Applied rules — {count} replacements made.")
        except Exception as e:
            QMessageBox.critical(self, "Rules Error", str(e))

    # ─── Build output (stores _export_df, no table shown) ────────────────────
    def _build_output(self):
        if self._df is None:
            QMessageBox.warning(self, "Build Output", "Open a price file first.")
            return
        if self._get_supplier_id() is None:
            QMessageBox.warning(self, "Build Output", "Select a supplier first.")
            return
        retail_col = self._col_combos["retail"].currentText()
        cost_col = self._col_combos["cost"].currentText()
        try:
            df = self._df
            decimals = int(self._cmb_decimals.currentText())

            def col(key: str) -> str:
                return self._col_combos[key].currentText()

            out = pd.DataFrame(index=df.index)
            out["PLU"] = safe_str(df[col("plu")]) if col("plu") != "(None)" else ""
            out["Supplier_Code"] = safe_str(df[col("sc")]) if col("sc") != "(None)" else ""
            out["Pharmacode"] = safe_str(df[col("pharmacode")]) if col("pharmacode") != "(None)" else ""
            out["Retail"] = parse_money(df[retail_col]).round(decimals) if retail_col != "(None)" else 0.0
            out["Cost"] = parse_money(df[cost_col]).round(decimals) if cost_col != "(None)" else 0.0

            desc_cols = [
                self._lst_desc_cols.item(i).text()
                for i in range(self._lst_desc_cols.count())
                if self._lst_desc_cols.item(i).isSelected()
            ]
            trade = (
                combine_columns(df, desc_cols, sep=self._cmb_sep.currentText())
                if desc_cols
                else pd.Series([""] * len(df), index=df.index)
            )
            trade = clean_description(trade)
            if self._chk_title_case.isChecked():
                trade = trade.str.title()
            if self._chk_norm_units.isChecked():
                trade = normalize_units(trade)
            prefix = self._txt_prefix.text().strip()
            if prefix:
                trade = prefix + " " + trade
            out["TradeName"] = trade.fillna("").astype(str)

            # Auto-apply per-supplier shortening rules if any exist
            rules = self._load_supplier_rules(self._get_supplier_id())
            if rules is not None and not rules.empty:
                out["TradeName"], n_replacements = apply_replacements(
                    out["TradeName"], rules, whole_word=True, case_insensitive=True
                )
                if n_replacements:
                    self.status_message.emit(
                        f"Supplier rules: {n_replacements} replacement(s) applied."
                    )

            barcode_key = col("barcode")
            out["Barcode"] = safe_str(df[barcode_key]) if barcode_key != "(None)" else ""

            outer_key = col("outer")
            if outer_key != "(None)":
                outer_vals = pd.to_numeric(df[outer_key], errors="coerce").fillna(0).round().astype(int)
                out["Outers"] = outer_vals.where(outer_vals > 0, 1)
            else:
                out["Outers"] = 1

            out = out.dropna(how="all")

            # De-dupe non-blank PLUs
            if col("plu") != "(None)":
                non_blank = out["PLU"].astype("string").str.strip() != ""
                out = pd.concat([
                    out[~non_blank],
                    out[non_blank].drop_duplicates(subset=["PLU"], keep="last"),
                ]).sort_index()

            # Exclude rows where barcode, supplier code AND pharmacode are all blank
            if self._chk_excl_no_id.isChecked():
                has_id = (
                    out["Barcode"].astype(str).str.strip().ne("") |
                    out["Supplier_Code"].astype(str).str.strip().ne("") |
                    out["Pharmacode"].astype(str).str.strip().ne("")
                )
                excluded = (~has_id).sum()
                out = out[has_id]
                if excluded:
                    self.status_message.emit(
                        f"Excluded {excluded} row(s) with no Barcode, Supplier Code or Pharmacode."
                    )

            self._export_df = out.reset_index(drop=True)
            self._preview_model = None
            self._preview_table.setModel(None)
            self._lbl_preview_info.setText(
                f"Built {len(self._export_df)} rows — click <b>Preview Changes</b> to compare with Access."
            )
            self.status_message.emit(f"Built {len(self._export_df)} rows.")
            self._save_supplier_settings()
        except Exception as e:
            QMessageBox.critical(self, "Build Error", str(e))

    # ─── Supplier rules (supplier_rules.json) ────────────────────────────────
    def _load_supplier_rules(self, supplier_id: Optional[int]) -> Optional[pd.DataFrame]:
        """Return a rules DataFrame for the given supplier, or None if none defined."""
        if supplier_id is None:
            return None
        try:
            p = Path(SUPPLIER_RULES_FILE)
            if not p.exists():
                return None
            data = json.loads(p.read_text(encoding="utf-8"))
            raw = data.get(str(supplier_id), [])
            if not raw:
                return None
            rules = pd.DataFrame(raw)
            rules.columns = [c.strip().lower() for c in rules.columns]
            if "from" not in rules.columns or "to" not in rules.columns:
                return None
            rules = rules[["from", "to"]].dropna()
            rules["from"] = rules["from"].astype("string").str.strip()
            rules["to"] = rules["to"].astype("string").str.strip()
            return rules[rules["from"] != ""].reset_index(drop=True)
        except Exception:
            return None

    def _update_supplier_rules_label(self):
        supplier_id = self._get_supplier_id()
        rules = self._load_supplier_rules(supplier_id)
        if rules is not None and not rules.empty:
            self._lbl_supplier_rules.setText(f"⚡ {len(rules)} shortening rule(s) will auto-apply")
            self._lbl_supplier_rules.setStyleSheet("font-size: 11px; color: #155724; padding: 1px 2px;")
        else:
            self._lbl_supplier_rules.setText("No shortening rules for this supplier")
            self._lbl_supplier_rules.setStyleSheet("font-size: 11px; color: #6c757d; padding: 1px 2px;")

    def _open_rules_dialog(self):
        supplier_id = self._get_supplier_id()
        if supplier_id is None:
            QMessageBox.warning(self, "Rules", "Select a supplier first.")
            return
        supplier_name = self._cmb_supplier.currentText()

        # Load existing rules as raw list of dicts
        try:
            p = Path(SUPPLIER_RULES_FILE)
            all_rules = json.loads(p.read_text(encoding="utf-8")) if p.exists() else {}
        except Exception:
            all_rules = {}
        existing = [
            r for r in all_rules.get(str(supplier_id), [])
            if isinstance(r, dict) and "from" in r
        ]

        dlg = SupplierRulesDialog(supplier_name, existing, parent=self)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return

        new_rules = dlg.get_rules()

        # Save back to JSON (preserve other suppliers, preserve _readme key)
        all_rules[str(supplier_id)] = new_rules
        try:
            Path(SUPPLIER_RULES_FILE).write_text(
                json.dumps(all_rules, indent=2, ensure_ascii=False), encoding="utf-8"
            )
        except Exception as e:
            QMessageBox.warning(self, "Rules", f"Could not save rules:\n{e}")
            return

        self._update_supplier_rules_label()

        # Apply the updated rules to the current preview if one exists
        if self._preview_model is not None and new_rules:
            rules_df = pd.DataFrame(new_rules)
            rules_df["from"] = rules_df["from"].astype("string").str.strip()
            rules_df["to"] = rules_df["to"].astype("string").str.strip()
            df = self._preview_model.dataFrame()
            df["New_Description"], count = apply_replacements(
                df["New_Description"], rules_df, whole_word=True, case_insensitive=True
            )
            df["Chars"] = df["New_Description"].str.len()
            df["Valid"] = df["Chars"].apply(lambda x: "✓" if x <= 40 else "⚠")
            self._show_preview(df)
            self.status_message.emit(
                f"Supplier rules saved — {count} replacement(s) applied to preview."
            )
        else:
            self.status_message.emit(f"Rules saved for {supplier_name}.")

    # ─── Supplier settings persistence ───────────────────────────────────────
    def _read_all_supplier_settings(self) -> dict:
        try:
            p = Path(SUPPLIER_SETTINGS_FILE)
            if p.exists():
                return json.loads(p.read_text(encoding="utf-8"))
        except Exception:
            pass
        return {}

    def _save_supplier_settings(self):
        supplier_id = self._get_supplier_id()
        if supplier_id is None:
            return
        sheet_key = self._get_sheet_key()
        sheet_settings = {
            "template": self._cmb_template.currentText(),
            "skip_rows": self._spin_skip.value(),
            "columns": {key: cb.currentText() for key, cb in self._col_combos.items()},
            "description": {
                "columns": [
                    self._lst_desc_cols.item(i).text()
                    for i in range(self._lst_desc_cols.count())
                    if self._lst_desc_cols.item(i).isSelected()
                ],
                "prefix": self._txt_prefix.text().strip(),
                "separator": self._cmb_sep.currentText(),
                "title_case": self._chk_title_case.isChecked(),
                "normalize_units": self._chk_norm_units.isChecked(),
            },
        }
        all_s = self._read_all_supplier_settings()
        sup_key = str(supplier_id)
        entry = all_s.get(sup_key, {})
        # Migrate old flat format to sheet-keyed format on first save
        if "sheets" not in entry and ("template" in entry or "columns" in entry):
            entry = {"sheets": {"_default": entry}}
        if "sheets" not in entry:
            entry = {"sheets": {}}
        sheets = entry["sheets"]
        sheets[sheet_key] = sheet_settings
        # Enforce max 6 sheet layouts (drop oldest if exceeded)
        if len(sheets) > 6:
            del sheets[next(iter(sheets))]
        entry["sheets"] = sheets
        all_s[sup_key] = entry
        try:
            Path(SUPPLIER_SETTINGS_FILE).write_text(
                json.dumps(all_s, indent=2, ensure_ascii=False), encoding="utf-8"
            )
            self.status_message.emit(
                f"Settings saved for {self._cmb_supplier.currentText()} [{sheet_key}]"
            )
            self._update_sheet_combo(sheets)
        except Exception as e:
            QMessageBox.warning(self, "Settings", f"Could not save supplier settings:\n{e}")

    def _apply_supplier_settings(self, supplier_id: int):
        """Apply saved supplier settings for the current sheet key to the UI."""
        all_s = self._read_all_supplier_settings()
        entry = all_s.get(str(supplier_id))
        if not entry:
            self._update_sheet_combo({})
            self._update_desc_preview()
            return

        # Migrate old flat format to sheet-keyed format on read
        if "sheets" not in entry and ("template" in entry or "columns" in entry):
            entry = {"sheets": {"_default": entry}}
        sheets = entry.get("sheets", {})
        self._update_sheet_combo(sheets)

        # Pick settings for current sheet key, fall back to first available
        sheet_key = self._get_sheet_key()
        settings = sheets.get(sheet_key) or next(iter(sheets.values()), None)
        if not settings:
            self._update_desc_preview()
            return

        # Template
        tpl = settings.get("template", "")
        if tpl:
            idx = self._cmb_template.findText(tpl)
            if idx >= 0:
                self._cmb_template.blockSignals(True)
                self._cmb_template.setCurrentIndex(idx)
                self._cmb_template.blockSignals(False)

        # Skip rows
        self._spin_skip.blockSignals(True)
        self._spin_skip.setValue(settings.get("skip_rows", 0))
        self._spin_skip.blockSignals(False)

        # Column mappings — only apply if the saved column name exists in the combo
        for key, saved_col in settings.get("columns", {}).items():
            cb = self._col_combos.get(key)
            if cb is None:
                continue
            idx = cb.findText(saved_col)
            if idx >= 0:
                cb.setCurrentIndex(idx)

        # Description columns — reorder list widget and set selection
        desc = settings.get("description", {})
        saved_cols = desc.get("columns", [])
        if saved_cols and self._df is not None:
            all_items = [
                self._lst_desc_cols.item(i).text()
                for i in range(self._lst_desc_cols.count())
            ]
            saved_existing = [c for c in saved_cols if c in all_items]
            remaining = [c for c in all_items if c not in saved_existing]
            self._lst_desc_cols.blockSignals(True)
            self._lst_desc_cols.clear()
            for col in saved_existing + remaining:
                self._lst_desc_cols.addItem(col)
            for i in range(self._lst_desc_cols.count()):
                item = self._lst_desc_cols.item(i)
                item.setSelected(item.text() in saved_existing)
            self._lst_desc_cols.blockSignals(False)

        # Prefix / separator / title case / normalize units
        self._txt_prefix.blockSignals(True)
        self._txt_prefix.setText(desc.get("prefix", ""))
        self._txt_prefix.blockSignals(False)
        sep = desc.get("separator", " ")
        sep_idx = self._cmb_sep.findText(sep)
        if sep_idx >= 0:
            self._cmb_sep.setCurrentIndex(sep_idx)
        self._chk_title_case.setChecked(bool(desc.get("title_case", False)))
        self._chk_norm_units.setChecked(bool(desc.get("normalize_units", True)))

        self._update_desc_preview()
        self._update_supplier_rules_label()

    def _on_supplier_changed(self):
        """When the user picks a different supplier, load their saved settings and prefetch Access data."""
        self._update_supplier_rules_label()
        supplier_id = self._get_supplier_id()
        if supplier_id is not None:
            self._start_prefetch(supplier_id)
        if self._df is None:
            return
        if supplier_id is not None:
            self._apply_supplier_settings(supplier_id)

    def _start_prefetch(self, supplier_id: int):
        """Start a background fetch of existing Access Details for the given supplier."""
        # Cancel any in-progress prefetch for a different supplier
        if self._prefetch_worker is not None and self._prefetch_worker.isRunning():
            self._prefetch_worker.done.disconnect()
            self._prefetch_worker.error.disconnect()
            self._prefetch_worker = None
        self._prefetched_details = None
        self._prefetch_supplier_id = None
        self.status_message.emit("Fetching supplier data from Access…")
        self._prefetch_worker = PrefetchWorker(supplier_id)
        self._prefetch_worker.done.connect(self._on_prefetch_done)
        self._prefetch_worker.error.connect(self._on_prefetch_error)
        self._prefetch_worker.start()

    def _on_prefetch_done(self, supplier_id: int, df: pd.DataFrame):
        self._prefetch_supplier_id = supplier_id
        self._prefetched_details = df
        self.status_message.emit(
            f"Supplier data ready — {len(df)} existing records cached."
        )

    def _on_prefetch_error(self, msg: str):
        self._prefetched_details = None
        self._prefetch_supplier_id = None
        self.status_message.emit(f"Prefetch failed (will fetch on demand): {msg}")

    def _get_sheet_key(self) -> str:
        """Return the settings key for the currently active sheet (or '_default')."""
        return self._current_sheet or "_default"

    def _update_sheet_combo(self, sheets: dict):
        """Populate and show/hide the sheet layout combo based on saved sheet layouts."""
        self._cmb_sheet_config.blockSignals(True)
        self._cmb_sheet_config.clear()
        for sheet_key in sheets:
            label = sheet_key if sheet_key != "_default" else "(default)"
            self._cmb_sheet_config.addItem(label, userData=sheet_key)
        # Select current sheet in the combo
        current_key = self._get_sheet_key()
        for i in range(self._cmb_sheet_config.count()):
            if self._cmb_sheet_config.itemData(i) == current_key:
                self._cmb_sheet_config.setCurrentIndex(i)
                break
        visible = self._cmb_sheet_config.count() > 1
        self._lbl_sheet_config.setVisible(visible)
        self._cmb_sheet_config.setVisible(visible)
        self._cmb_sheet_config.blockSignals(False)

    def _on_sheet_config_changed(self, index: int):
        """User selected a different saved sheet layout — apply its stored settings."""
        if index < 0:
            return
        sheet_key = self._cmb_sheet_config.itemData(index)
        if sheet_key:
            self._current_sheet = None if sheet_key == "_default" else sheet_key
        supplier_id = self._get_supplier_id()
        if supplier_id is not None:
            self._apply_supplier_settings(supplier_id)


    def _show_preview(self, df: pd.DataFrame):
        self._preview_model = EditablePandasModel(
            df.copy(),
            editable_cols=["New_Description"],
            currency_cols=["Old_Cost", "New_Cost", "Old_Retail", "New_Retail"],
        )
        self._preview_table.setModel(self._preview_model)
        self._preview_table.setItemDelegate(CursorPlacingDelegate(self._preview_table))
        self._preview_table.setSortingEnabled(True)
        self._preview_table.resizeColumnsToContents()
        if "New_Description" in df.columns:
            idx = list(df.columns).index("New_Description")
            self._preview_table.setColumnWidth(idx, 280)

    # ─── Bulk truncate ────────────────────────────────────────────────────────
    def _bulk_truncate(self):
        if self._preview_model is None:
            QMessageBox.warning(self, "Truncate", "Run Preview Changes first.")
            return
        df = self._preview_model.dataFrame()
        df["New_Description"] = df["New_Description"].str.slice(0, 40)
        df["Chars"] = df["New_Description"].str.len()
        df["Valid"] = df["Chars"].apply(lambda x: "✓" if x <= 40 else "⚠")
        self._show_preview(df)
        self.status_message.emit("Bulk truncated all New Descriptions to 40 chars.")

    # ─── Export CSV ───────────────────────────────────────────────────────────
    def _export_csv(self):
        export = self._get_export_df()
        if export is None:
            QMessageBox.warning(self, "Export", "Build output first.")
            return
        tpl_name = self._cmb_template.currentText()
        path, _ = QFileDialog.getSaveFileName(
            self, "Save Price File CSV",
            f"price_file_{tpl_name}.csv",
            "CSV (*.csv)",
        )
        if not path:
            return
        delim_map = {",": ",", ";": ";", "Tab": "\t"}
        delim = delim_map.get(self._cmb_delim.currentText(), ",")
        try:
            export.to_csv(path, index=False, sep=delim)
            self.status_message.emit(f"Saved to {Path(path).name}")
        except Exception as e:
            QMessageBox.critical(self, "Export Error", str(e))

    @staticmethod
    def _preview_to_export_df(df: pd.DataFrame) -> pd.DataFrame:
        """Convert a preview-table DataFrame to the export format expected by upsert_details."""
        result = pd.DataFrame({
            "PLU": df.get("PLU", pd.Series([""] * len(df))),
            "Supplier_Code": df["Supplier_Code"],
            "Pharmacode": df.get("Pharmacode", pd.Series([""] * len(df))),
            "Retail": df["New_Retail"],
            "Cost": df["New_Cost"],
            "TradeName": df["New_Description"],
            "Barcode": df["Barcode"],
        })
        if "Outer" in df.columns:
            result["Outer"] = df["Outer"].values
        return result

    def _get_export_df(self) -> Optional[pd.DataFrame]:
        """Export-format DataFrame from the preview table (current edits) or raw build."""
        if self._preview_model is not None:
            return self._preview_to_export_df(self._preview_model.dataFrame())
        return self._export_df

    # ─── Upsert helpers ───────────────────────────────────────────────────────
    def _get_supplier_id(self) -> Optional[int]:
        if self._suppliers_df is None:
            return None
        label = self._cmb_supplier.currentText()
        row = self._suppliers_df[self._suppliers_df["Label"] == label]
        return int(row["SupplierID"].iloc[0]) if not row.empty else None

    def _get_supplier_name_code(self) -> tuple[str, str]:
        """Return (SupplierName, SupplierCode) for the currently selected supplier."""
        if self._suppliers_df is None:
            return ("", "")
        label = self._cmb_supplier.currentText()
        row = self._suppliers_df[self._suppliers_df["Label"] == label]
        if row.empty:
            return ("", "")
        return (str(row["SupplierName"].iloc[0] or ""), str(row["SupplierCode"].iloc[0] or ""))

    # ─── Preview upsert ───────────────────────────────────────────────────────
    def _preview_upsert(self):
        if self._export_df is None:
            QMessageBox.warning(self, "Preview", "Build output first.")
            return
        supplier_id = self._get_supplier_id()
        if supplier_id is None:
            QMessageBox.warning(self, "Preview", "No supplier selected.")
            return
        self._btn_preview_upsert.setEnabled(False)

        # If the prefetch is still running for this supplier, wait for it to
        # finish rather than opening a second concurrent Access connection.
        if (
            self._prefetch_worker is not None
            and self._prefetch_worker.isRunning()
        ):
            self._btn_preview_upsert.setText("Waiting for supplier data…")
            self._prefetch_worker.done.connect(lambda sid, df: self._preview_upsert())
            return

        # Use pre-fetched Access data if available for the current supplier
        cached = (
            self._prefetched_details
            if self._prefetch_supplier_id == supplier_id and self._prefetched_details is not None
            else None
        )
        self._btn_preview_upsert.setText("Comparing…" if cached is not None else "Loading…")
        self._preview_worker = PreviewWorker(self._export_df, supplier_id, existing_df=cached)
        self._preview_worker.done.connect(self._on_preview_done)
        self._preview_worker.error.connect(self._on_preview_error)
        self._preview_worker.start()

    def _on_preview_done(self, preview_df: pd.DataFrame, stats: dict):
        self._btn_preview_upsert.setEnabled(True)
        self._btn_preview_upsert.setText("Preview Changes")
        self._last_preview_stats = stats

        # Add Chars + Valid columns
        preview_df["Chars"] = preview_df["New_Description"].str.len()
        preview_df["Valid"] = preview_df["Chars"].apply(lambda x: "✓" if x <= 40 else "⚠")
        # Reorder so Chars/Valid appear right after New_Description
        cols = list(preview_df.columns)
        for extra in ["Valid", "Chars"]:
            if extra in cols:
                cols.remove(extra)
                cols.insert(cols.index("New_Description") + 1, extra)
        preview_df = preview_df[cols]

        new_count = (preview_df["Status"] == "NEW").sum()
        upd_count = (preview_df["Status"] == "UPDATE").sum()
        unchanged_count = (preview_df["Status"] == "UNCHANGED").sum()
        long_count = (preview_df["Chars"] > 40).sum()
        info = (
            f"{stats['existing_count']} in Access  |  "
            f"matched by code: {stats['matched_by_code']}  |  "
            f"matched by barcode: {stats['matched_by_barcode']}  |  "
            f"{new_count} new  |  {upd_count} updates  |  "
            f"<span style='color:#6c757d'>{unchanged_count} unchanged</span>"
        )
        if long_count:
            info += f"  |  <span style='color:#842029'>⚠ {long_count} exceed 40 chars</span>"
        self._lbl_preview_info.setText(info)

        self._show_preview(preview_df)
        self.status_message.emit(
            f"Preview: {new_count} new, {upd_count} updates, {unchanged_count} unchanged."
        )

    def _on_preview_error(self, msg: str):
        self._btn_preview_upsert.setEnabled(True)
        self._btn_preview_upsert.setText("Preview Changes")
        QMessageBox.critical(self, "Preview Error", msg)

    # ─── Execute upsert ───────────────────────────────────────────────────────
    def _exec_upsert(self):
        if self._preview_model is None:
            QMessageBox.warning(self, "Upsert", "Run Preview Changes first.")
            return
        supplier_id = self._get_supplier_id()
        if supplier_id is None:
            QMessageBox.warning(self, "Upsert", "No supplier selected.")
            return

        preview_df = self._preview_model.dataFrame().copy()
        invalid_mask = preview_df.get("Chars", pd.Series([0] * len(preview_df))) > 40
        fallback_mask = invalid_mask & preview_df.get("Status", "").eq("UPDATE")
        skip_mask = invalid_mask & ~fallback_mask

        if fallback_mask.any():
            preview_df.loc[fallback_mask, "New_Description"] = (
                preview_df.loc[fallback_mask, "Old_Description"].fillna("").astype(str)
            )
            preview_df.loc[fallback_mask, "Chars"] = (
                preview_df.loc[fallback_mask, "New_Description"].str.len()
            )
            preview_df.loc[fallback_mask, "Valid"] = preview_df.loc[fallback_mask, "Chars"].apply(
                lambda x: "✓" if x <= 40 else "⚠"
            )

        invalid_rows = preview_df[skip_mask].copy()
        valid_rows = preview_df[~skip_mask].copy()

        if valid_rows.empty:
            QMessageBox.warning(
                self, "Upsert",
                "All rows have descriptions > 40 chars.\nFix them before upserting.",
            )
            return

        supplier_label = self._cmb_supplier.currentText()
        skip_note = ""
        fallback_note = ""
        if fallback_mask.any():
            fallback_note = (
                f"\n\n{int(fallback_mask.sum())} update row(s) have descriptions > 40 chars "
                f"and will keep their existing Access description."
            )
        if not invalid_rows.empty:
            skip_note = (
                f"\n\n{len(invalid_rows)} row(s) with descriptions > 40 chars "
                f"will be skipped and remain in the table."
            )
        if QMessageBox.question(
            self, "Confirm Upsert",
            f"Upsert {len(valid_rows)} row(s) for:\n{supplier_label}"
            f"{fallback_note}{skip_note}\n\nContinue?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        ) != QMessageBox.StandardButton.Yes:
            return

        self._upsert_skipped_rows = invalid_rows if not invalid_rows.empty else None
        export = self._preview_to_export_df(valid_rows)
        self._btn_exec_upsert.setEnabled(False)
        self._btn_exec_upsert.setText("Starting…")
        cached = (
            self._prefetched_details
            if self._prefetch_supplier_id == supplier_id and self._prefetched_details is not None
            else None
        )
        self._upsert_worker = UpsertWorker(
            export, supplier_id, self._chk_mark_updated.isChecked(),
            existing_df=cached,
        )
        self._upsert_worker.progress.connect(self._on_upsert_progress)
        self._upsert_worker.done.connect(self._on_upsert_done)
        self._upsert_worker.error.connect(self._on_upsert_error)
        self._upsert_worker.start()

    def _on_upsert_progress(self, done: int, total: int):
        remaining = total - done
        self._btn_exec_upsert.setText(f"{remaining} rows remaining…")

    def _on_upsert_done(self, inserted: int, updated: int):
        self._btn_exec_upsert.setEnabled(True)
        self._btn_exec_upsert.setText("Execute Upsert")
        supplier_label = self._cmb_supplier.currentText()

        skipped = self._upsert_skipped_rows
        self._upsert_skipped_rows = None

        skip_note = f"\n\n{len(skipped)} row(s) skipped (description > 40 chars) — fix and upsert again." if skipped is not None else ""
        QMessageBox.information(
            self, "Upsert Complete",
            f"Done — {inserted} inserted, {updated} updated\nfor {supplier_label}{skip_note}",
        )

        if skipped is not None and not skipped.empty:
            # Rebuild the preview with only the invalid rows remaining
            self._show_preview(skipped)
            self._lbl_preview_info.setText(
                f"<span style='color:#842029;'>⚠ {len(skipped)} row(s) still have descriptions "
                f"> 40 chars — fix and upsert again.</span>"
            )
        else:
            self._preview_model = None
            self._preview_table.setModel(None)
            self._lbl_preview_info.setText("")

        self.status_message.emit(f"Upsert done: {inserted} inserted, {updated} updated.")
        # Refresh Access cache so the next operation sees the new data
        supplier_id = self._get_supplier_id()
        if supplier_id is not None:
            self._start_prefetch(supplier_id)

    def _on_upsert_error(self, msg: str):
        self._btn_exec_upsert.setEnabled(True)
        self._btn_exec_upsert.setText("Execute Upsert")
        QMessageBox.critical(self, "Upsert Error", msg)

    # ─── Clear / Start Fresh ──────────────────────────────────────────────────
    def _clear_state(self):
        if self._preview_model is not None or self._export_df is not None:
            if QMessageBox.question(
                self, "Clear / Start Fresh",
                "This will discard the loaded file and all unsaved work.\n\nContinue?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            ) != QMessageBox.StandardButton.Yes:
                return

        self._file_path = None
        self._current_sheet = None
        self._excel_sheets = []
        self._df = None
        self._export_df = None
        self._preview_model = None
        self._upsert_skipped_rows = None
        self._last_preview_stats = None
        self._prefetched_details = None
        self._prefetch_supplier_id = None

        self._lbl_file.setText("No file loaded")
        self._cmb_sheet_config.blockSignals(True)
        self._cmb_sheet_config.clear()
        self._cmb_sheet_config.blockSignals(False)
        self._lbl_sheet_config.setVisible(False)
        self._cmb_sheet_config.setVisible(False)
        self._spin_skip.blockSignals(True)
        self._spin_skip.setValue(0)
        self._spin_skip.blockSignals(False)
        self._source_table.setModel(None)
        self._preview_table.setModel(None)
        self._lbl_preview_info.setText("")
        # Clear column combos back to (None)-only
        for cb in self._col_combos.values():
            cb.blockSignals(True)
            cb.clear()
            cb.addItem("(None)")
            cb.blockSignals(False)
        self._lst_desc_cols.clear()
        self._txt_prefix.blockSignals(True)
        self._txt_prefix.clear()
        self._txt_prefix.blockSignals(False)
        self._lbl_desc_preview.setText("(load a file and select columns)")
        self.status_message.emit("Cleared — ready for next file.")

    # ─── Replace All in Access ────────────────────────────────────────────────
    def _exec_replace(self):
        supplier_id = self._get_supplier_id()
        if supplier_id is None:
            QMessageBox.warning(self, "Replace All", "Select a supplier first.")
            return
        export = self._get_export_df()
        if export is None:
            QMessageBox.warning(self, "Replace All", "Build output first.")
            return

        # If a preview is open, only use valid rows
        if self._preview_model is not None:
            preview_df = self._preview_model.dataFrame()
            invalid_mask = preview_df.get("Chars", 0) > 40
            if invalid_mask.any():
                QMessageBox.warning(
                    self, "Replace All",
                    f"{invalid_mask.sum()} row(s) have descriptions > 40 chars.\n"
                    "Fix or truncate them before replacing.",
                )
                return
            export = self._preview_to_export_df(preview_df)

        supplier_label = self._cmb_supplier.currentText()
        if QMessageBox.warning(
            self, "⚠ Replace All in Access — ARE YOU SURE?",
            f"This will PERMANENTLY DELETE every existing record for:\n\n"
            f"    {supplier_label}\n\n"
            f"and replace them with {len(export)} row(s) from the current file.\n\n"
            f"This CANNOT be undone. Continue?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        ) != QMessageBox.StandardButton.Yes:
            return

        self._btn_replace_all.setEnabled(False)
        self._btn_replace_all.setText("Replacing…")
        self._replace_worker = ReplaceWorker(
            export, supplier_id, self._chk_mark_updated.isChecked()
        )
        self._replace_worker.done.connect(self._on_replace_done)
        self._replace_worker.error.connect(self._on_replace_error)
        self._replace_worker.progress.connect(self._on_replace_progress)
        self._replace_worker.start()

    def _on_replace_progress(self, done: int, total: int):
        self._btn_replace_all.setText(f"Replacing… {done} / {total}")

    def _on_replace_done(self, inserted: int):
        self._btn_replace_all.setEnabled(True)
        self._btn_replace_all.setText("Replace All in Access")
        supplier_label = self._cmb_supplier.currentText()
        QMessageBox.information(
            self, "Replace Complete",
            f"Done — all previous records deleted and {inserted} row(s) inserted\nfor {supplier_label}.",
        )
        self._preview_model = None
        self._preview_table.setModel(None)
        self._lbl_preview_info.setText("")
        self.status_message.emit(f"Replace complete: {inserted} rows inserted for {supplier_label}.")
        # Refresh Access cache so the next operation sees the new data
        supplier_id = self._get_supplier_id()
        if supplier_id is not None:
            self._start_prefetch(supplier_id)

    def _on_replace_error(self, msg: str):
        self._btn_replace_all.setEnabled(True)
        self._btn_replace_all.setText("Replace All in Access")
        QMessageBox.critical(self, "Replace Error", msg)

    # ─── SQLite upsert ────────────────────────────────────────────────────────
    def _exec_sqlite_upsert(self):
        supplier_id = self._get_supplier_id()
        if supplier_id is None:
            QMessageBox.warning(self, "SQLite Upsert", "Select a supplier first.")
            return
        export = self._get_export_df()
        if export is None:
            QMessageBox.warning(self, "SQLite Upsert", "Build output first.")
            return
        supplier_label = self._cmb_supplier.currentText()
        if QMessageBox.question(
            self, "Confirm SQLite Upsert",
            f"Upsert {len(export)} row(s) into suppliers.sqlite for:\n{supplier_label}\n\nContinue?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        ) != QMessageBox.StandardButton.Yes:
            return
        self._btn_sqlite_upsert.setEnabled(False)
        self._btn_sqlite_upsert.setText("Running…")
        supplier_name, supplier_code = self._get_supplier_name_code()
        self._sqlite_upsert_worker = SqliteUpsertWorker(
            export, supplier_id, self._chk_mark_updated.isChecked(),
            supplier_name=supplier_name, supplier_code=supplier_code,
        )
        self._sqlite_upsert_worker.done.connect(self._on_sqlite_upsert_done)
        self._sqlite_upsert_worker.error.connect(self._on_sqlite_upsert_error)
        self._sqlite_upsert_worker.start()

    def _on_sqlite_upsert_done(self, updated: int, inserted: int):
        self._btn_sqlite_upsert.setEnabled(True)
        self._btn_sqlite_upsert.setText("Upsert → SQLite")
        supplier_label = self._cmb_supplier.currentText()
        QMessageBox.information(
            self, "SQLite Upsert Complete",
            f"Done — {inserted} inserted, {updated} updated\nfor {supplier_label} in suppliers.sqlite",
        )
        self.status_message.emit(f"SQLite upsert done: {inserted} inserted, {updated} updated.")

    def _on_sqlite_upsert_error(self, msg: str):
        self._btn_sqlite_upsert.setEnabled(True)
        self._btn_sqlite_upsert.setText("Upsert → SQLite")
        QMessageBox.critical(self, "SQLite Upsert Error", msg)

    # ─── SQLite replace all ───────────────────────────────────────────────────
    def _exec_sqlite_replace(self):
        supplier_id = self._get_supplier_id()
        if supplier_id is None:
            QMessageBox.warning(self, "SQLite Replace", "Select a supplier first.")
            return
        export = self._get_export_df()
        if export is None:
            QMessageBox.warning(self, "SQLite Replace", "Build output first.")
            return
        supplier_label = self._cmb_supplier.currentText()
        if QMessageBox.warning(
            self, "⚠ Replace All in SQLite — ARE YOU SURE?",
            f"This will PERMANENTLY DELETE every existing SQLite record for:\n\n"
            f"    {supplier_label}\n\n"
            f"and replace them with {len(export)} row(s) from the current file.\n\n"
            f"This CANNOT be undone. Continue?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        ) != QMessageBox.StandardButton.Yes:
            return
        self._btn_sqlite_replace.setEnabled(False)
        self._btn_sqlite_replace.setText("Replacing…")
        supplier_name, supplier_code = self._get_supplier_name_code()
        self._sqlite_replace_worker = SqliteReplaceWorker(
            export, supplier_id, self._chk_mark_updated.isChecked(),
            supplier_name=supplier_name, supplier_code=supplier_code,
        )
        self._sqlite_replace_worker.done.connect(self._on_sqlite_replace_done)
        self._sqlite_replace_worker.error.connect(self._on_sqlite_replace_error)
        self._sqlite_replace_worker.start()

    def _on_sqlite_replace_done(self, inserted: int):
        self._btn_sqlite_replace.setEnabled(True)
        self._btn_sqlite_replace.setText("Replace All → SQLite")
        supplier_label = self._cmb_supplier.currentText()
        QMessageBox.information(
            self, "SQLite Replace Complete",
            f"Done — all previous records deleted and {inserted} row(s) inserted\n"
            f"for {supplier_label} in suppliers.sqlite",
        )
        self.status_message.emit(f"SQLite replace complete: {inserted} rows for {supplier_label}.")

    def _on_sqlite_replace_error(self, msg: str):
        self._btn_sqlite_replace.setEnabled(True)
        self._btn_sqlite_replace.setText("Replace All → SQLite")
        QMessageBox.critical(self, "SQLite Replace Error", msg)

    # ─── Safe-close helper ────────────────────────────────────────────────────
    def active_write_worker(self):
        """Return the currently running write worker, or None."""
        for w in (
            self._upsert_worker,
            self._replace_worker,
            self._sqlite_upsert_worker,
            self._sqlite_replace_worker,
        ):
            if w is not None and w.isRunning():
                return w
        return None

    # ─── SQLite browser ───────────────────────────────────────────────────────
    def _open_sqlite_browser(self):
        """Open the SQLite database browser / editor dialog."""
        from db.sqlite_browser import SqliteBrowserDialog
        dlg = SqliteBrowserDialog(self)
        dlg.exec()

    # ─── Diagnostics ──────────────────────────────────────────────────────────
    def _show_diagnostics(self):
        stats = self._last_preview_stats
        if stats is None:
            QMessageBox.information(self, "Match Diagnostics", "Run Preview Changes first.")
            return

        def fmt(lst):
            return "\n    ".join(lst) if lst else "(none)"

        msg = (
            f"Access records for this supplier: {stats['existing_count']}\n"
            f"Matched by supplier code: {stats['matched_by_code']}\n"
            f"Matched by barcode:       {stats['matched_by_barcode']}\n"
            f"New (no match found):     {stats['new_count']}\n"
            f"\n"
            f"── Access codes (sample) ──\n    {fmt(stats['access_codes'])}\n"
            f"\n"
            f"── Export codes (sample) ──\n    {fmt(stats['export_codes'])}\n"
            f"\n"
            f"── Access barcodes (sample) ──\n    {fmt(stats['access_barcodes'])}\n"
            f"\n"
            f"── Export barcodes (sample) ──\n    {fmt(stats['export_barcodes'])}\n"
            f"\n"
            f"If all rows show as NEW, check that the codes/barcodes above match "
            f"in format (leading zeros, spaces, etc.)."
        )
        dlg = QMessageBox(self)
        dlg.setWindowTitle("Match Diagnostics")
        dlg.setText(msg)
        dlg.setTextFormat(Qt.TextFormat.PlainText)
        dlg.exec()

    # ─── Sort helpers ─────────────────────────────────────────────────────────
    def _sort_new_top(self):
        if self._preview_model is None:
            return
        df = self._preview_model.dataFrame()
        if "Status" not in df.columns:
            return
        df["_sort"] = df["Status"].map({"NEW": 0, "UPDATE": 1}).fillna(2)
        df = df.sort_values("_sort").drop(columns=["_sort"]).reset_index(drop=True)
        self._show_preview(df)

    def _sort_invalid_top(self):
        if self._preview_model is None:
            return
        df = self._preview_model.dataFrame()
        if "Valid" not in df.columns:
            return
        df["_sort"] = df["Valid"].apply(lambda v: 0 if str(v) == "⚠" else 1)
        df = df.sort_values("_sort").drop(columns=["_sort"]).reset_index(drop=True)
        self._show_preview(df)


# ── Main window ───────────────────────────────────────────────────────────────
class SettingsDialog(QDialog):
    """Dialog for editing settings.ini paths and Access password."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Settings")
        self.setMinimumWidth(520)

        layout = QGridLayout(self)
        layout.setSpacing(8)
        layout.setContentsMargins(12, 12, 12, 12)

        # Access MDB path
        layout.addWidget(QLabel("Access database (.mdb):"), 0, 0)
        self._mdb_edit = QLineEdit(config.get_mdb_path())
        layout.addWidget(self._mdb_edit, 0, 1)
        btn_mdb = QPushButton("Browse…")
        btn_mdb.clicked.connect(self._browse_mdb)
        layout.addWidget(btn_mdb, 0, 2)

        # Access password
        layout.addWidget(QLabel("Access password:"), 1, 0)
        self._pwd_edit = QLineEdit(config.get_access_password())
        self._pwd_edit.setEchoMode(QLineEdit.EchoMode.Password)
        layout.addWidget(self._pwd_edit, 1, 1, 1, 2)

        # SQLite path
        layout.addWidget(QLabel("SQLite database (.sqlite):"), 2, 0)
        self._sqlite_edit = QLineEdit(config.get_sqlite_path())
        layout.addWidget(self._sqlite_edit, 2, 1)
        btn_sqlite = QPushButton("Browse…")
        btn_sqlite.clicked.connect(self._browse_sqlite)
        layout.addWidget(btn_sqlite, 2, 2)

        # OK / Cancel
        btn_row = QHBoxLayout()
        btn_row.addStretch()
        btn_ok = QPushButton("OK")
        btn_ok.setDefault(True)
        btn_ok.clicked.connect(self._save_and_accept)
        btn_cancel = QPushButton("Cancel")
        btn_cancel.clicked.connect(self.reject)
        btn_row.addWidget(btn_ok)
        btn_row.addWidget(btn_cancel)
        layout.addLayout(btn_row, 3, 0, 1, 3)

    def _browse_mdb(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Select Access Database", "", "Access Database (*.mdb *.accdb)"
        )
        if path:
            self._mdb_edit.setText(path)

    def _browse_sqlite(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Select SQLite Database", "", "SQLite Database (*.sqlite *.db)",
            options=QFileDialog.Option.DontConfirmOverwrite,
        )
        if path:
            self._sqlite_edit.setText(path)

    def _save_and_accept(self):
        config.set_mdb_path(self._mdb_edit.text().strip())
        config.set_access_password(self._pwd_edit.text())
        config.set_sqlite_path(self._sqlite_edit.text().strip())
        self.accept()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Price File Builder")

        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(4, 4, 4, 4)
        root.setSpacing(4)

        # Mode radio buttons
        mode_row = QHBoxLayout()
        mode_row.addWidget(QLabel("Mode:"))
        self._mode_group = QButtonGroup(self)
        rb_price = QRadioButton("Price File")
        rb_specials = QRadioButton("Specials File")
        rb_price.setChecked(True)
        self._mode_group.addButton(rb_price, 0)
        self._mode_group.addButton(rb_specials, 1)
        self._mode_group.idClicked.connect(self._on_mode_changed)
        mode_row.addWidget(rb_price)
        mode_row.addWidget(rb_specials)
        mode_row.addStretch()
        btn_settings = QPushButton("Settings…")
        btn_settings.clicked.connect(self._open_settings)
        mode_row.addWidget(btn_settings)
        root.addLayout(mode_row)

        self._stack = QStackedWidget()

        self._price_widget = PriceFileWidget()
        self._price_widget.status_message.connect(self.statusBar().showMessage)
        self._stack.addWidget(self._price_widget)

        self._specials_widget = SpecialsWidget()
        self._specials_widget.status_message.connect(self.statusBar().showMessage)
        self._stack.addWidget(self._specials_widget)

        root.addWidget(self._stack)
        self.statusBar().showMessage("Ready")

    def _on_mode_changed(self, idx: int):
        self._stack.setCurrentIndex(idx)

    def _open_settings(self):
        dlg = SettingsDialog(self)
        dlg.exec()

    def closeEvent(self, event):
        """Block close if a database write is in progress to prevent partial commits."""
        writers = []
        w = self._price_widget.active_write_worker()
        if w:
            writers.append(w)
        if self._specials_widget.is_write_in_progress():
            writers.append(self._specials_widget._worker)

        if writers:
            reply = QMessageBox.question(
                self,
                "Write in progress",
                "A database write is still in progress.\n\n"
                "Wait for it to finish before closing?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.Yes,
            )
            if reply == QMessageBox.StandardButton.Yes:
                for w in writers:
                    w.wait(30_000)   # wait up to 30 seconds
        event.accept()


# ── Entry point ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    win = MainWindow()
    win.showMaximized()
    sys.exit(app.exec())
