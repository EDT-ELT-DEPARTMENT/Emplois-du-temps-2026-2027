# -*- coding: utf-8 -*-
"""Page indépendante : Gestion des EDTs & administration."""
from pathlib import Path
import runpy
import streamlit as st

st.session_state["legacy_module"] = "gestion"
st.session_state["legacy_page"] = "accueil_admin"
runtime = Path(__file__).resolve().parents[1] / "core" / "legacy_runtime.py"
runpy.run_path(str(runtime), run_name="__main__")
