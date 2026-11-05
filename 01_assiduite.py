# -*- coding: utf-8 -*-
"""Page indépendante : Suivi d’assiduité."""
from pathlib import Path
import runpy
import streamlit as st

st.session_state["legacy_module"] = "assiduite"
st.session_state["legacy_page"] = "accueil"
runtime = Path(__file__).resolve().parents[1] / "core" / "legacy_runtime.py"
runpy.run_path(str(runtime), run_name="__main__")
