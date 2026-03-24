"""sqlite_browser.py — SQLite database browser / editor dialog."""
from __future__ import annotations

from typing import Optional

import pandas as pd
from PyQt6.QtCore import Qt, QModelIndex, QThread, pyqtSignal
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QPushButton, QComboBox, QGroupBox,
    QPlainTextEdit, QMessageBox, QSizePolicy, QFrame, QSplitter,
)

from . import sqlite_helper


# ─── Thin worker for slow saves ───────────────────────────────────────────────
class _SaveWorker(QThread):
    done = pyqtSignal(int)
    error = pyqtSignal(str)

    def __init__(self, supplier_id: int, df: pd.DataFrame):
        super().__init__()
        self.supplier_id = supplier_id
        self.df = df

    def run(self):
        try:
            n = sqlite_helper.save_details_from_browser(self.supplier_id, self.df)
            self.done.emit(n)
        except Exception as e:
            self.error.emit(str(e))


# ─── Dialog ───────────────────────────────────────────────────────────────────
class SqliteBrowserDialog(QDialog):
    """Browse, edit, and delete rows in the SQLite suppliers database."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("SQLite Database Browser")
        self.resize(1150, 720)
        self.setMinimumSize(800, 500)

        self._df: Optional[pd.DataFrame] = None          # currently displayed data
        self._model = None                                # PandasModel or EditablePandasModel
        self._supplier_id: Optional[int] = None
        self._save_worker: Optional[_SaveWorker] = None

        # Import Qt table helpers from qt_app at runtime to avoid circular imports
        from qt_app import EditablePandasModel, DeletableTableView, CursorPlacingDelegate
        self._EditablePandasModel = EditablePandasModel
        self._DeletableTableView = DeletableTableView
        self._CursorPlacingDelegate = CursorPlacingDelegate

        self._build_ui()
        self._load_supplier_list()

    # ── UI construction ───────────────────────────────────────────────────────
    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(8, 8, 8, 8)
        root.setSpacing(6)

        # ── Toolbar ──────────────────────────────────────────────────────────
        toolbar = QHBoxLayout()

        toolbar.addWidget(QLabel("Table:"))
        self._cmb_table = QComboBox()
        self._cmb_table.setMinimumWidth(120)
        self._cmb_table.currentIndexChanged.connect(self._on_table_changed)
        toolbar.addWidget(self._cmb_table)

        self._lbl_sup = QLabel("Supplier:")
        toolbar.addWidget(self._lbl_sup)
        self._cmb_supplier = QComboBox()
        self._cmb_supplier.setMinimumWidth(220)
        toolbar.addWidget(self._cmb_supplier)

        btn_load = QPushButton("Load")
        btn_load.setStyleSheet("font-weight: bold; padding: 4px 14px;")
        btn_load.clicked.connect(self._load_data)
        toolbar.addWidget(btn_load)

        toolbar.addStretch()
        self._lbl_count = QLabel("")
        self._lbl_count.setStyleSheet("color: #6c757d; font-size: 11px;")
        toolbar.addWidget(self._lbl_count)

        root.addLayout(toolbar)

        # ── Splitter: table (top) | code lookup (bottom) ─────────────────────
        splitter = QSplitter(Qt.Orientation.Vertical)
        splitter.setChildrenCollapsible(False)

        # Main table
        self._table = self._DeletableTableView()
        splitter.addWidget(self._table)

        # Bottom pane
        bottom_w = QFrame()
        bottom_w.setFrameShape(QFrame.Shape.StyledPanel)
        bottom_layout = QHBoxLayout(bottom_w)
        bottom_layout.setContentsMargins(4, 4, 4, 4)

        # Code/barcode lookup group
        lookup_grp = QGroupBox("Lookup / Delete by Code or Barcode")
        lookup_v = QVBoxLayout(lookup_grp)
        lookup_v.addWidget(QLabel(
            "Paste codes or barcodes (one per line or comma-separated):"
        ))
        self._txt_codes = QPlainTextEdit()
        self._txt_codes.setMaximumHeight(70)
        self._txt_codes.setPlaceholderText("12345\n67890\nor:  12345, 67890")
        lookup_v.addWidget(self._txt_codes)
        lbtn_row = QHBoxLayout()
        btn_filter = QPushButton("Filter to matching rows")
        btn_filter.clicked.connect(self._filter_by_codes)
        btn_del_codes = QPushButton("Delete matching rows")
        btn_del_codes.setStyleSheet("color: #842029;")
        btn_del_codes.clicked.connect(self._delete_by_codes)
        btn_reset_filter = QPushButton("Show all rows")
        btn_reset_filter.clicked.connect(self._load_data)
        lbtn_row.addWidget(btn_filter)
        lbtn_row.addWidget(btn_del_codes)
        lbtn_row.addWidget(btn_reset_filter)
        lbtn_row.addStretch()
        lookup_v.addLayout(lbtn_row)
        bottom_layout.addWidget(lookup_grp, 3)

        # Action buttons
        act_grp = QGroupBox("Actions")
        act_v = QVBoxLayout(act_grp)
        btn_del_sel = QPushButton("Delete Selected Rows")
        btn_del_sel.setStyleSheet("color: #842029; font-weight: bold; padding: 4px 10px;")
        btn_del_sel.clicked.connect(self._delete_selected)
        self._btn_save = QPushButton("Save Changes to SQLite")
        self._btn_save.setStyleSheet(
            "background-color: #0d6efd; color: white; font-weight: bold; padding: 6px 16px;"
        )
        self._btn_save.clicked.connect(self._save_changes)
        btn_discard = QPushButton("Discard Changes")
        btn_discard.clicked.connect(self._load_data)
        act_v.addWidget(btn_del_sel)
        act_v.addWidget(self._btn_save)
        act_v.addWidget(btn_discard)
        act_v.addStretch()
        bottom_layout.addWidget(act_grp, 1)

        splitter.addWidget(bottom_w)
        splitter.setSizes([480, 180])
        root.addWidget(splitter, 1)

        # Status bar
        self._lbl_status = QLabel("Load a table to begin.")
        self._lbl_status.setStyleSheet("color: #6c757d; font-size: 11px; padding: 2px 0;")
        root.addWidget(self._lbl_status)

    # ── Data loading ──────────────────────────────────────────────────────────
    def _load_supplier_list(self):
        """Populate table and supplier combos from the SQLite DB."""
        # Tables
        try:
            tables = sqlite_helper.get_table_names()
        except Exception:
            tables = ["Details", "Suppliers"]
        self._cmb_table.blockSignals(True)
        self._cmb_table.clear()
        self._cmb_table.addItems(tables if tables else ["Details", "Suppliers"])
        self._cmb_table.blockSignals(False)

        # Suppliers
        try:
            sup_df = sqlite_helper.load_suppliers()
            self._sup_df = sup_df
            self._cmb_supplier.clear()
            for _, row in sup_df.iterrows():
                label = (
                    f'{row["SupplierName"]} ({row["SupplierCode"]})'
                    if str(row.get("SupplierCode") or "").strip()
                    else str(row["SupplierName"])
                )
                self._cmb_supplier.addItem(label, userData=int(row["SupplierID"]))
        except Exception as e:
            self._set_status(f"Could not load suppliers: {e}")

        self._on_table_changed()

    def _on_table_changed(self):
        """Show/hide the supplier selector depending on whether Details is selected."""
        is_details = self._cmb_table.currentText() == "Details"
        self._lbl_sup.setVisible(is_details)
        self._cmb_supplier.setVisible(is_details)

    def _load_data(self):
        """Load the selected table (and supplier if Details) into the grid."""
        table = self._cmb_table.currentText()
        try:
            if table == "Details":
                sid = self._cmb_supplier.currentData()
                if sid is None:
                    self._set_status("No supplier selected.")
                    return
                self._supplier_id = int(sid)
                df = sqlite_helper.load_details_for_browser(self._supplier_id)
                editable = sqlite_helper.BROWSER_EDITABLE_COLS
            else:
                self._supplier_id = None
                df = sqlite_helper.load_full_table(table)
                editable = []

            self._df = df.copy()
            self._show_grid(df, editable)
            self._lbl_count.setText(f"{len(df)} rows")
            self._set_status(f"Loaded {len(df)} rows from {table}.")
        except Exception as e:
            QMessageBox.critical(self, "Load Error", str(e))

    def _show_grid(self, df: pd.DataFrame, editable_cols: list[str]):
        currency = ["Cost", "Retail"] if "Cost" in df.columns else []
        if editable_cols:
            model = self._EditablePandasModel(
                df, editable_cols=editable_cols, currency_cols=currency
            )
        else:
            from qt_app import PandasModel
            model = PandasModel(df, currency_cols=currency)
        self._model = model
        self._table.setModel(model)
        self._table.setItemDelegate(self._CursorPlacingDelegate(self._table))
        self._table.resizeColumnsToContents()
        # Give Description more room
        if "Description" in df.columns:
            idx = list(df.columns).index("Description")
            self._table.setColumnWidth(idx, 260)

    # ── Code/barcode lookup ───────────────────────────────────────────────────
    def _parse_codes(self) -> list[str]:
        """Parse the lookup text box into a list of non-empty trimmed strings."""
        raw = self._txt_codes.toPlainText()
        codes = []
        for token in raw.replace(",", "\n").splitlines():
            t = token.strip()
            if t:
                codes.append(t)
        return codes

    def _filter_by_codes(self):
        """Filter the current grid to rows whose Code or Barcode is in the lookup list."""
        if self._df is None:
            return
        codes = self._parse_codes()
        if not codes:
            self._set_status("Enter at least one code to filter.")
            return
        code_set = set(codes)
        mask = (
            self._df.get("Code", pd.Series(dtype=str)).astype(str).str.strip().isin(code_set) |
            self._df.get("Barcode", pd.Series(dtype=str)).astype(str).str.strip().isin(code_set)
        )
        filtered = self._df[mask].reset_index(drop=True)
        editable = sqlite_helper.BROWSER_EDITABLE_COLS if self._supplier_id else []
        self._show_grid(filtered, editable)
        self._lbl_count.setText(f"{len(filtered)} / {len(self._df)} rows")
        self._set_status(f"Filtered to {len(filtered)} matching rows.")

    def _delete_by_codes(self):
        """Delete rows from SQLite directly by code/barcode list without going through Save."""
        if self._supplier_id is None:
            QMessageBox.warning(self, "Delete", "Select a supplier and load Details first.")
            return
        codes = self._parse_codes()
        if not codes:
            self._set_status("Enter at least one code to delete.")
            return
        if QMessageBox.question(
            self, "Confirm Delete",
            f"Permanently delete all Details rows for this supplier\n"
            f"where Code or Barcode matches any of the {len(codes)} value(s)?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        ) != QMessageBox.StandardButton.Yes:
            return
        try:
            n = sqlite_helper.delete_details_by_keys(self._supplier_id, codes)
            self._set_status(f"Deleted {n} row(s).")
            self._load_data()   # refresh grid
        except Exception as e:
            QMessageBox.critical(self, "Delete Error", str(e))

    # ── Grid row deletion ─────────────────────────────────────────────────────
    def _delete_selected(self):
        """Remove selected rows from the in-memory grid (not saved until Save is clicked)."""
        if self._model is None:
            return
        rows = sorted(
            {idx.row() for idx in self._table.selectedIndexes()},
            reverse=True,
        )
        if not rows:
            self._set_status("No rows selected.")
            return
        self._model.remove_rows(rows)
        self._df = self._model.dataFrame().copy()
        self._lbl_count.setText(f"{len(self._df)} rows")
        self._set_status(f"Removed {len(rows)} row(s) from view — click Save to commit.")

    # ── Save ──────────────────────────────────────────────────────────────────
    def _save_changes(self):
        """Write the current grid contents back to SQLite (replace all for supplier)."""
        if self._supplier_id is None:
            QMessageBox.warning(self, "Save", "Only Details (by supplier) can be saved here.")
            return
        if self._model is None:
            return
        df = self._model.dataFrame()
        sup_label = self._cmb_supplier.currentText()
        if QMessageBox.question(
            self, "Confirm Save",
            f"Replace all {len(df)} row(s) in SQLite Details\nfor supplier: {sup_label}\n\nContinue?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        ) != QMessageBox.StandardButton.Yes:
            return
        self._btn_save.setEnabled(False)
        self._btn_save.setText("Saving…")
        self._save_worker = _SaveWorker(self._supplier_id, df)
        self._save_worker.done.connect(self._on_save_done)
        self._save_worker.error.connect(self._on_save_error)
        self._save_worker.start()

    def _on_save_done(self, n: int):
        self._btn_save.setEnabled(True)
        self._btn_save.setText("Save Changes to SQLite")
        self._set_status(f"Saved {n} rows to SQLite.")
        self._load_data()   # reload to confirm round-trip

    def _on_save_error(self, msg: str):
        self._btn_save.setEnabled(True)
        self._btn_save.setText("Save Changes to SQLite")
        QMessageBox.critical(self, "Save Error", msg)

    # ── Helpers ───────────────────────────────────────────────────────────────
    def _set_status(self, msg: str):
        self._lbl_status.setText(msg)
