"""config.py — Read/write settings.ini for configurable paths and credentials."""
from __future__ import annotations

import os
import configparser

SETTINGS_FILE = "settings.ini"

DEFAULTS = {
    "Access": {
        "mdb_path": "suppliers.mdb",
        "password": "LOCKIE MONDAY",
    },
    "SQLite": {
        "db_path": "suppliers.sqlite",
    },
}


def _settings_path() -> str:
    return os.path.join(os.getcwd(), SETTINGS_FILE)


def load() -> configparser.ConfigParser:
    """Load settings.ini, creating it with defaults if it doesn't exist."""
    cfg = configparser.ConfigParser()
    path = _settings_path()
    if not os.path.exists(path):
        for section, values in DEFAULTS.items():
            cfg[section] = values
        with open(path, "w") as f:
            cfg.write(f)
    else:
        cfg.read(path)
        # Ensure all default keys exist (non-destructive merge)
        changed = False
        for section, values in DEFAULTS.items():
            if not cfg.has_section(section):
                cfg.add_section(section)
                changed = True
            for key, val in values.items():
                if not cfg.has_option(section, key):
                    cfg.set(section, key, val)
                    changed = True
        if changed:
            save(cfg)
    return cfg


def save(cfg: configparser.ConfigParser) -> None:
    with open(_settings_path(), "w") as f:
        cfg.write(f)


# ── Getters ───────────────────────────────────────────────────────────────────

def get_mdb_path() -> str:
    return load().get("Access", "mdb_path", fallback=DEFAULTS["Access"]["mdb_path"])


def get_access_password() -> str:
    return load().get("Access", "password", fallback=DEFAULTS["Access"]["password"])


def get_sqlite_path() -> str:
    return load().get("SQLite", "db_path", fallback=DEFAULTS["SQLite"]["db_path"])


# ── Setters ───────────────────────────────────────────────────────────────────

def set_mdb_path(path: str) -> None:
    cfg = load()
    cfg.set("Access", "mdb_path", path)
    save(cfg)


def set_access_password(password: str) -> None:
    cfg = load()
    cfg.set("Access", "password", password)
    save(cfg)


def set_sqlite_path(path: str) -> None:
    cfg = load()
    cfg.set("SQLite", "db_path", path)
    save(cfg)
