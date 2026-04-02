# app_settings.py — Settings sidebar for the Streamlit app
from __future__ import annotations

import streamlit as st
import config


def render_settings_sidebar():
    """Render the connection settings expander in the sidebar."""
    with st.sidebar:
        st.title("Price File Builder")
        with st.expander("Connection Settings", expanded=False):
            mdb = st.text_input(
                "Access database path (.mdb)",
                value=config.get_mdb_path(),
                key="settings_mdb_path",
            )
            pwd = st.text_input(
                "Access password",
                value=config.get_access_password(),
                type="password",
                key="settings_access_pwd",
            )
            sqlite = st.text_input(
                "SQLite database path (.sqlite)",
                value=config.get_sqlite_path(),
                key="settings_sqlite_path",
            )
            if st.button("Save Settings", key="settings_save"):
                try:
                    config.set_mdb_path(mdb.strip())
                    config.set_access_password(pwd)
                    config.set_sqlite_path(sqlite.strip())
                    st.success("Settings saved.")
                    # Clear any cached supplier/details data
                    for k in list(st.session_state.keys()):
                        if k.startswith("price_prefetch") or k.startswith("spec_brands"):
                            del st.session_state[k]
                    st.rerun()
                except Exception as e:
                    st.error(f"Could not save settings: {e}")
