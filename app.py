# -*- coding: utf-8 -*-
"""Point d'entrée multipage de la plateforme EDT UDL-SBA."""
from pathlib import Path
import streamlit as st

st.set_page_config(
    page_title="EDT UDL-SBA",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="expanded",
)

BASE_DIR = Path(__file__).parent
st.markdown("""
<style>
:root { --udl-navy:#102a43; --udl-blue:#1d4e89; --udl-accent:#2f80c9; }
[data-testid="stSidebar"] {
    background: linear-gradient(180deg,#102a43 0%,#173f67 58%,#1d4e89 100%);
    border-right: 1px solid rgba(255,255,255,.12);
}
[data-testid="stSidebar"] * { color:#f7fbff; }
[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p { color:#e7f0f8; }
[data-testid="stSidebar"] button {
    border:1px solid rgba(255,255,255,.16);
    border-radius:9px;
}
[data-testid="stSidebar"] button:hover {
    background:rgba(255,255,255,.12);
    border-color:rgba(255,255,255,.35);
}
.block-container { padding-top:1.6rem; padding-bottom:3rem; max-width:1600px; }
h1,h2,h3 { color:#102a43; letter-spacing:-.02em; }
[data-testid="stMetric"] {
    background:#f4f8fc; border:1px solid #dce7f1;
    border-radius:12px; padding:14px 16px;
}
[data-testid="stExpander"] {
    border:1px solid #dce7f1; border-radius:12px;
}
.stButton>button, .stDownloadButton>button, .stFormSubmitButton>button {
    border-radius:9px; font-weight:600;
}
hr { border-color:#dce7f1; }
@media(max-width:800px) {
    .block-container { padding-left:1rem; padding-right:1rem; }
}
</style>
""", unsafe_allow_html=True)

st.sidebar.markdown(
    "<div style='text-align:center;padding:18px 4px 12px'>"
    "<div style='font-size:34px'>🏛️</div>"
    "<div style='font-size:24px;font-weight:800;letter-spacing:.08em'>UDL-SBA</div>"
    "<div style='font-size:11px;opacity:.85'>Département d'Électrotechnique · FGE</div>"
    "</div>",
    unsafe_allow_html=True,
)
st.sidebar.caption("PLATEFORME UNIVERSITAIRE · 2026–2027")

pages = {
    "PILOTAGE": [
        st.Page("pages/01_assiduite.py", title="Suivi d’assiduité", icon="📊", default=True),
        st.Page("pages/02_gestion_edt.py", title="Gestion des EDT & Admin", icon="📅"),
        st.Page("pages/03_edt_intelligent.py", title="EDT intelligent", icon="🧠"),
    ],
    "ESPACE ÉTUDIANT": [
        st.Page("pages/04_portail_etudiant.py", title="Mon compte", icon="🎓"),
        st.Page("pages/06_portail_edt.py", title="Mon emploi du temps", icon="🗓️"),
        st.Page("pages/07_portail_pdf.py", title="Télécharger mon EDT", icon="📄"),
        st.Page("pages/08_portail_absences.py", title="Mes absences", icon="📋"),
        st.Page("pages/09_portail_messages.py", title="Messages", icon="✉️"),
    ],
    "VIE PÉDAGOGIQUE": [
        st.Page("pages/05_pv_pedagogiques.py", title="PV des comités pédagogiques", icon="📝"),
    ],
}
navigation = st.navigation(pages, position="sidebar")
navigation.run()
