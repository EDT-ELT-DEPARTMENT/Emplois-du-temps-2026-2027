# -*- coding: utf-8 -*-
"""Page indépendante du portail étudiant : Mes absences."""
from pathlib import Path
import runpy
import streamlit as st

st.session_state["legacy_module"] = "portail"
st.session_state["legacy_page"] = "mes_absences"
runtime = Path(__file__).resolve().parents[1] / "core" / "legacy_runtime.py"
runpy.run_path(str(runtime), run_name="__main__")
