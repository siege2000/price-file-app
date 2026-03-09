# helpers.py — re-export shim for backwards compatibility
# Logic has been split into:
#   file_helpers.py        — file + template loading
#   column_helpers.py      — column mapping / template defaults
#   description_helpers.py — text building / normalization
#   editing_helpers.py     — money parsing + replacement rules
#   access_helpers.py      — Access DB connection, suppliers, save/upsert

from file_helpers import load_file, load_templates
from column_helpers import safe_str, idx_for, tpl_field_default, tpl_desc_defaults
from description_helpers import combine_columns, normalize_units, clean_description, desc_len
from editing_helpers import parse_money, load_replacements_csv, apply_replacements
from access_helpers import (
    MAX_DESC_LEN,
    ACCESS_PASSWORD,
    ACCESS_FILE,
    get_access_conn,
    load_suppliers,
    new_guid,
    save_to_details,
    upsert_details,
)

__all__ = [
    "MAX_DESC_LEN",
    "ACCESS_PASSWORD",
    "ACCESS_FILE",
    "load_file",
    "load_templates",
    "safe_str",
    "idx_for",
    "tpl_field_default",
    "tpl_desc_defaults",
    "combine_columns",
    "normalize_units",
    "clean_description",
    "desc_len",
    "parse_money",
    "load_replacements_csv",
    "apply_replacements",
    "get_access_conn",
    "load_suppliers",
    "new_guid",
    "save_to_details",
    "upsert_details",
]
