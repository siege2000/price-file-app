"""export_find_replace.py

One-off (and re-runnable) script to export the FindAndReplace table from
Settings1.mdb into data/supplier_rules.json.

Usage:
    python scripts/export_find_replace.py

Behaviour:
- Reads FindAndReplace (SupplierID, Find, Replace) from Settings1.mdb.
- Merges into data/supplier_rules.json using the format:
      {"<SupplierID>": [{"from": "...", "to": "..."}, ...], ...}
- Rules imported from Access REPLACE any previous Access-sourced rules for
  that supplier, but manually-added rules prefixed with "regex:" are
  preserved alongside them.
- Prints a per-supplier summary.

Run from the project root so that settings.ini and data/ are found correctly.
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

# Allow running from project root without installing the package
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from db.access_helpers import load_find_replace_rules  # noqa: E402

RULES_FILE = Path("data/supplier_rules.json")


def main() -> None:
    print("Connecting to Settings1.mdb …")
    try:
        df = load_find_replace_rules()
    except Exception as e:
        print(f"ERROR reading FindAndReplace table: {e}")
        sys.exit(1)

    if df.empty:
        print("FindAndReplace table is empty — nothing to export.")
        return

    # Normalise column names (Access drivers can vary case)
    df.columns = [c.strip() for c in df.columns]
    df = df.dropna(subset=["Find"])
    df["Find"] = df["Find"].astype(str).str.strip()
    df["Replace"] = df["Replace"].fillna("").astype(str).str.strip()
    df = df[df["Find"] != ""]

    # Load existing rules file (preserve _readme and any regex rules)
    if RULES_FILE.exists():
        try:
            existing: dict = json.loads(RULES_FILE.read_text(encoding="utf-8"))
        except Exception:
            existing = {}
    else:
        existing = {}

    readme = existing.get("_readme", [
        "Per-supplier description shortening rules.",
        "Key = SupplierID (integer, as a string).",
        "Each rule has 'from' and 'to'.",
        "Rules are applied as whole-word, case-insensitive replacements during Build Output.",
        "Prefix 'from' with 'regex:' to use a regular expression instead of a literal match.",
    ])

    totals: dict[str, int] = {}

    for supplier_id, group in df.groupby("SupplierID"):
        key = str(int(supplier_id))
        access_rules = [
            {"from": row["Find"], "to": row["Replace"]}
            for _, row in group.iterrows()
        ]

        # Keep any existing regex rules that were manually added
        old_rules = existing.get(key, [])
        manual_regex = [
            r for r in old_rules
            if isinstance(r, dict) and str(r.get("from", "")).lower().startswith("regex:")
        ]

        existing[key] = access_rules + manual_regex
        totals[key] = len(access_rules)

    # Write back — _readme first, then numeric keys in order
    ordered: dict = {"_readme": readme}
    for k in sorted((k for k in existing if k != "_readme"), key=lambda x: int(x) if x.isdigit() else 0):
        ordered[k] = existing[k]

    RULES_FILE.write_text(json.dumps(ordered, indent=2, ensure_ascii=False), encoding="utf-8")

    print(f"\nExported to {RULES_FILE}:")
    for key, count in sorted(totals.items(), key=lambda x: int(x) if x.isdigit() else 0):
        print(f"  Supplier {key}: {count} rule(s)")
    print(f"\nTotal suppliers updated: {len(totals)}")


if __name__ == "__main__":
    main()
