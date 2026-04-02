# stockcard_dialog.py — Search MSSQL Healthsoft stock table for a barcode
"""
Opens from the "Open Stockcards" button in the specials grid.
Pre-fills the search with the selected row's Description, queries
SUPPORT0630\\HEALTHSOFTLOTS (configured in settings.ini [MSSQL]),
and writes the chosen Barcode back to the grid row.
"""
from __future__ import annotations

import pyodbc
import pandas as pd
from PyQt6.QtCore import Qt, QAbstractTableModel, QModelIndex, QSortFilterProxyModel, QThread, pyqtSignal
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton,
    QTableView, QAbstractItemView, QMessageBox,
)

import config


# ── MSSQL connection + query ───────────────────────────────────────────────────

def _get_conn():
    server   = config.get_mssql_server()
    database = config.get_mssql_database()
    if not server:
        raise RuntimeError(
            "MSSQL server not configured.\n"
            "Set [MSSQL] server= in settings.ini."
        )
    if not database:
        raise RuntimeError(
            "MSSQL database not configured.\n"
            "Set [MSSQL] database= in settings.ini."
        )
    conn_str = (
        f"DRIVER={{SQL Server}};"
        f"SERVER={server};"
        f"DATABASE={database};"
        f"Trusted_Connection=yes;"
    )
    return pyodbc.connect(conn_str, timeout=10)


def search_stock(search_text: str) -> pd.DataFrame:
    """
    Search the stock table by tradename (LIKE %term%), joining to the Barcode
    table on StockID to retrieve barcodes.
    Returns DataFrame with columns: Description, Barcode, Pharmacode.
    Results capped at 200 rows.
    """
    sql = """
        SELECT TOP 200
            s.tradename  AS Description,
            b.Barcode    AS Barcode,
            s.PLU AS Pharmacode
        FROM stock s
        LEFT JOIN Barcode b ON b.StockID = s.StockID
        WHERE s.tradename LIKE ?
        ORDER BY s.tradename,
                 CASE WHEN b.Barcode IS NULL OR b.Barcode = '' THEN 0 ELSE 1 END,
                 b.Barcode
    """
    term = f"%{search_text.strip()}%"
    with _get_conn() as conn:
        cur = conn.cursor()
        cur.execute(sql, (term,))
        cols = [d[0] for d in cur.description]
        rows = cur.fetchall()
    return pd.DataFrame([list(r) for r in rows], columns=cols)


# ── Read-only result table model ──────────────────────────────────────────────

class _ResultModel(QAbstractTableModel):
    def __init__(self, df: pd.DataFrame, parent=None):
        super().__init__(parent)
        self._df = df

    def rowCount(self, parent=QModelIndex()):
        return len(self._df)

    def columnCount(self, parent=QModelIndex()):
        return len(self._df.columns)

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if role == Qt.ItemDataRole.DisplayRole and index.isValid():
            val = self._df.iloc[index.row(), index.column()]
            return "" if pd.isna(val) else str(val)
        return None

    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
        if role != Qt.ItemDataRole.DisplayRole:
            return None
        if orientation == Qt.Orientation.Horizontal:
            return str(self._df.columns[section])
        return str(section + 1)

    def row_data(self, row: int) -> dict:
        return self._df.iloc[row].to_dict() if 0 <= row < len(self._df) else {}


# ── Background search worker ──────────────────────────────────────────────────

class _SearchWorker(QThread):
    finished = pyqtSignal(object)   # pd.DataFrame
    error    = pyqtSignal(str)

    def __init__(self, search_text: str):
        super().__init__()
        self._text = search_text

    def run(self):
        try:
            self.finished.emit(search_stock(self._text))
        except Exception as exc:
            self.error.emit(str(exc))


# ── Dialog ────────────────────────────────────────────────────────────────────

class StockcardSearchDialog(QDialog):
    """
    Search Healthsoft stock DB by description and pick a barcode.

    After exec() == Accepted:
        .selected_barcode    — barcode string from the chosen row
        .selected_pharmacode — pharmacode string from the chosen row (may be empty)
    """

    def __init__(self, initial_description: str = "", parent=None):
        super().__init__(parent)
        self.setWindowTitle("Find in Stock Database")
        self.resize(720, 460)
        self.selected_barcode    = ""
        self.selected_pharmacode = ""
        self._worker = None

        root = QVBoxLayout(self)
        root.setSpacing(6)

        # ── Search bar ────────────────────────────────────────────────────────
        search_row = QHBoxLayout()
        search_row.addWidget(QLabel("Search:"))
        self._txt_search = QLineEdit(initial_description)
        self._txt_search.returnPressed.connect(self._do_search)
        search_row.addWidget(self._txt_search)
        self._btn_search = QPushButton("&Search")
        self._btn_search.clicked.connect(self._do_search)
        search_row.addWidget(self._btn_search)
        root.addLayout(search_row)

        # ── Status label ──────────────────────────────────────────────────────
        self._lbl_status = QLabel("Enter a description and press Search (or press Enter).")
        self._lbl_status.setStyleSheet("font-size: 11px; color: #555; padding: 1px 0;")
        root.addWidget(self._lbl_status)

        # ── Results table ─────────────────────────────────────────────────────
        self._table = QTableView()
        self._table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self._table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self._table.horizontalHeader().setStretchLastSection(True)
        self._table.verticalHeader().setDefaultSectionSize(22)
        self._table.setAlternatingRowColors(True)
        self._table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self._table.setSortingEnabled(True)
        self._table.doubleClicked.connect(self._accept_selection)
        self._set_model(pd.DataFrame(columns=["Description", "Barcode", "Pharmacode"]))
        root.addWidget(self._table)

        # ── Bottom buttons ────────────────────────────────────────────────────
        btn_row = QHBoxLayout()
        self._btn_apply = QPushButton("&Apply Barcode to Row")
        self._btn_apply.setStyleSheet(
            "background-color: #0d6efd; color: white; font-weight: bold; padding: 4px 12px;"
        )
        self._btn_apply.setEnabled(False)
        self._btn_apply.clicked.connect(self._accept_selection)
        btn_cancel = QPushButton("Cancel")
        btn_cancel.clicked.connect(self.reject)
        btn_row.addWidget(self._btn_apply)
        btn_row.addStretch()
        btn_row.addWidget(btn_cancel)
        root.addLayout(btn_row)

        # Enable Apply when a row is selected
        self._table.selectionModel().selectionChanged.connect(
            lambda: self._btn_apply.setEnabled(bool(self._table.selectedIndexes()))
        )

        # Auto-search if we have an initial description
        if initial_description.strip():
            self._do_search()

    def _set_model(self, df: pd.DataFrame):
        self._result_model = _ResultModel(df)
        self._proxy = QSortFilterProxyModel(self)
        self._proxy.setSourceModel(self._result_model)
        self._table.setModel(self._proxy)
        # Re-wire selection signal after model replacement
        self._table.selectionModel().selectionChanged.connect(
            lambda: self._btn_apply.setEnabled(bool(self._table.selectedIndexes()))
        )

    def _do_search(self):
        term = self._txt_search.text().strip()
        if not term:
            return
        self._btn_search.setEnabled(False)
        self._btn_apply.setEnabled(False)
        self._lbl_status.setText("Searching…")
        self._worker = _SearchWorker(term)
        self._worker.finished.connect(self._on_results)
        self._worker.error.connect(self._on_error)
        self._worker.start()

    def _on_results(self, df: pd.DataFrame):
        self._btn_search.setEnabled(True)
        self._set_model(df)
        n = len(df)
        if n:
            self._lbl_status.setText(
                f"{n} result(s). Double-click a row or select and click Apply Barcode."
            )
        else:
            self._lbl_status.setText("No results found — try a shorter or different search term.")

    def _on_error(self, msg: str):
        self._btn_search.setEnabled(True)
        self._lbl_status.setText(f"Error: {msg}")
        QMessageBox.warning(self, "Stock Database Error", msg)

    def _accept_selection(self):
        idx = self._table.currentIndex()
        if not idx.isValid():
            return
        source_row = self._proxy.mapToSource(idx).row()
        row_data = self._result_model.row_data(source_row)
        barcode    = str(row_data.get("Barcode", "") or "").strip()
        pharmacode = str(row_data.get("Pharmacode", "") or "").strip()
        if not barcode and not pharmacode:
            QMessageBox.warning(
                self, "No Barcode",
                "The selected item has no barcode or pharmacode in the stock database."
            )
            return
        self.selected_barcode    = barcode
        self.selected_pharmacode = pharmacode
        self.accept()
