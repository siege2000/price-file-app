import json
from pathlib import Path
from typing import Any, Dict

import pandas as pd


def load_file(file) -> pd.DataFrame:
    name = file.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(file)
    if name.endswith(".xlsx") or name.endswith(".xls"):
        return pd.read_excel(file)
    raise ValueError("Unsupported file type (must be .csv or .xlsx)")


def load_templates(path: str = "templates.json") -> Dict[str, Any]:
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"Template file not found: {path}")
    data = json.loads(p.read_text(encoding="utf-8"))
    if "templates" not in data:
        raise ValueError("templates.json must contain a top-level key: 'templates'")
    return data["templates"]
