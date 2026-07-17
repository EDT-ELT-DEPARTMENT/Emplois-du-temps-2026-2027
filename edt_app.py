import streamlit as st
import pandas as pd
import os
import hashlib
import io
import re
from datetime import datetime
from supabase import create_client

# =============================================================================
# NOUVEAUTÉ : FONCTIONS UTILITAIRES PRO POUR L'EXPORT (PDF / HTML / EXCEL)
# =============================================================================

def sanitize_for_pdf(text):
    """Nettoie le texte pour fpdf (latin-1) en remplaçant les caractères problématiques."""
    if text is None or pd.isna(text):
        return ""
    text = str(text)
    # Remplacements des caractères Unicode problématiques pour latin-1
    replacements = {
        "'": "'", "'": "'", """: """, """: """, "–": "-", "—": "-",
        "…": "...", "«": """, "»": """, "œ": "oe", "Œ": "OE",
        "à": "a", "â": "a", "ä": "a", "á": "a", "ã": "a", "å": "a",
        "è": "e", "é": "e", "ê": "e", "ë": "e", "ē": "e", "ė": "e", "ę": "e",
        "ì": "i", "í": "i", "î": "i", "ï": "i", "ī": "i", "į": "i",
        "ò": "o", "ó": "o", "ô": "o", "ö": "o", "õ": "o", "ō": "o",
        "ù": "u", "ú": "u", "û": "u", "ü": "u", "ū": "u",
        "ç": "c", "ć": "c", "č": "c", "ñ": "n", "ń": "n",
        "ÿ": "y", "ý": "y",
        "À": "A", "Â": "A", "Ä": "A", "Á": "A", "Ã": "A",
        "È": "E", "É": "E", "Ê": "E", "Ë": "E",
        "Ì": "I", "Í": "I", "Î": "I", "Ï": "I",
        "Ò": "O", "Ó": "O", "Ô": "O", "Ö": "O", "Õ": "O",
        "Ù": "U", "Ú": "U", "Û": "U", "Ü": "U",
        "Ç": "C", "Ñ": "N",
    }
    for old, new in replacements.items():
        text = text.replace(old, new)
    # Supprime tout ce qui reste hors latin-1
    return text.encode('latin-1', 'ignore').decode('latin-1')


def generate_pro_pdf(df_source, title, subtitle="", orientation="L"):
    """Génère un PDF professionnel avec fpdf."""
    try:
        from fpdf import FPDF
    except ImportError:
        return None, "fpdf non installé"

    class ProPDF(FPDF):
        def header(self):
            self.set_font('Arial', 'B', 9)
            self.set_text_color(30, 58, 138)  # #1E3A8A
            header_text = sanitize_for_pdf("Plateforme de gestion des EDTs-S2-2027 - Departement d'Electrotechnique - FGE/UDL-SBA")
            self.cell(0, 6, header_text, 0, 1, 'C')
            self.set_draw_color(212, 175, 55)  # #D4AF37
            self.line(10, self.get_y(), self.w - 10, self.get_y())
            self.ln(3)

        def footer(self):
            self.set_y(-15)
            self.set_font('Arial', 'I', 8)
            self.set_text_color(128, 128, 128)
            self.cell(0, 10, f'Page {self.page_no()}', 0, 0, 'C')

    pdf = ProPDF(orientation=orientation, unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    # Titre principal
    pdf.set_font("Arial", "B", 16)
    pdf.set_text_color(30, 58, 138)
    pdf.cell(0, 10, sanitize_for_pdf(title), 0, 1, "C")

    if subtitle:
        pdf.set_font("Arial", "I", 10)
        pdf.set_text_color(100, 100, 100)
        pdf.cell(0, 6, sanitize_for_pdf(subtitle), 0, 1, "C")

    pdf.ln(5)

    # Tableau de données
    if df_source is not None and not df_source.empty:
        df_clean = df_source.fillna("").astype(str)
        cols = list(df_clean.columns)
        n_cols = len(cols)

        # Largeurs dynamiques
        page_w = pdf.w - 20
        col_w = page_w / n_cols

        # En-tête
        pdf.set_font("Arial", "B", 8)
        pdf.set_fill_color(30, 58, 138)
        pdf.set_text_color(255, 255, 255)

        for col in cols:
            pdf.cell(col_w, 8, sanitize_for_pdf(str(col)), 1, 0, "C", True)
        pdf.ln()

        # Données avec alternance de couleurs
        pdf.set_font("Arial", "", 7)
        pdf.set_text_color(0, 0, 0)

        for idx, row in df_clean.iterrows():
            if idx % 2 == 0:
                pdf.set_fill_color(248, 250, 252)
            else:
                pdf.set_fill_color(255, 255, 255)

            for val in row:
                cell_text = sanitize_for_pdf(str(val))
                # Tronquer si trop long
                if len(cell_text) > 50:
                    cell_text = cell_text[:47] + "..."
                pdf.cell(col_w, 6, cell_text, 1, 0, "L", True)
            pdf.ln()

    # Pied de page info
    pdf.ln(5)
    pdf.set_font("Arial", "I", 8)
    pdf.set_text_color(150, 150, 150)
    pdf.cell(0, 5, sanitize_for_pdf(f"Document genere le {datetime.now().strftime('%d/%m/%Y a %H:%M')}"), 0, 0, "R")

    return pdf.output(), None


def generate_pro_html(df_source, title, subtitle=""):
    """Génère un HTML professionnel et responsive."""

    # Construction du tableau HTML
    if df_source is not None and not df_source.empty:
        df_clean = df_source.fillna("").astype(str)
        rows_html = ""
        for idx, row in df_clean.iterrows():
            bg = "#f8fafc" if idx % 2 == 0 else "#ffffff"
            cells = "".join([f'<td style="padding:10px;border:1px solid #e2e8f0;font-size:13px;">{val}</td>' for val in row])
            rows_html += f'<tr style="background-color:{bg};">{cells}</tr>'

        headers = "".join([f'<th style="padding:10px;border:1px solid #e2e8f0;background:#1E3A8A;color:white;font-size:13px;">{c}</th>' for c in df_clean.columns])
        table_html = f'<table style="width:100%;border-collapse:collapse;margin-top:15px;">{headers}{rows_html}</table>'
    else:
        table_html = '<p style="text-align:center;color:#999;">Aucune donnee disponible</p>'

    html_content = f"""<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{title}</title>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap');
        body {{ font-family: 'Inter', 'Segoe UI', Arial, sans-serif; background: linear-gradient(135deg, #f1f5f9 0%, #e2e8f0 100%); margin: 0; padding: 30px; color: #1e293b; }}
        .container {{ max-width: 1200px; margin: 0 auto; background: white; border-radius: 16px; box-shadow: 0 10px 40px rgba(0,0,0,0.08); overflow: hidden; }}
        .header {{ background: linear-gradient(135deg, #1E3A8A 0%, #3B82F6 100%); color: white; padding: 30px 40px; text-align: center; }}
        .header h1 {{ margin: 0; font-size: 24px; font-weight: 700; letter-spacing: -0.5px; }}
        .header p {{ margin: 8px 0 0 0; opacity: 0.9; font-size: 14px; }}
        .badge {{ display: inline-block; background: #D4AF37; color: #1E3A8A; padding: 4px 14px; border-radius: 20px; font-size: 11px; font-weight: 700; margin-top: 10px; }}
        .content {{ padding: 30px 40px; }}
        .meta {{ display: flex; justify-content: space-between; color: #64748b; font-size: 12px; margin-bottom: 20px; padding-bottom: 15px; border-bottom: 1px solid #f1f5f9; }}
        table {{ width: 100%; border-collapse: separate; border-spacing: 0; }}
        th {{ position: sticky; top: 0; z-index: 10; }}
        tr:hover {{ background-color: #eff6ff !important; transition: background 0.2s; }}
        .footer {{ text-align: center; padding: 20px; color: #94a3b8; font-size: 12px; border-top: 1px solid #f1f5f9; }}
        @media print {{ body {{ background: white; padding: 0; }} .container {{ box-shadow: none; border-radius: 0; }} }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>{title}</h1>
            <p>{subtitle}</p>
            <span class="badge">EDT S2-2027 - FGE/UDL-SBA</span>
        </div>
        <div class="content">
            <div class="meta">
                <span>📅 Genere le {datetime.now().strftime('%d/%m/%Y a %H:%M')}</span>
                <span>{len(df_source) if df_source is not None else 0} lignes</span>
            </div>
            {table_html}
        </div>
        <div class="footer">
            Plateforme de gestion des EDTs - Departement d'Electrotechnique - Faculte de Genie Electrique - UDL-SBA
        </div>
    </div>
</body>
</html>"""
    return html_content


def generate_pro_excel(df_source, title, sheet_name="Donnees"):
    """Génère un Excel professionnel avec xlsxwriter."""
    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df_clean = df_source.fillna("").astype(str) if df_source is not None else pd.DataFrame()
        df_clean.to_excel(writer, index=False, sheet_name=sheet_name)

        workbook = writer.book
        worksheet = writer.sheets[sheet_name]

        # Formats
        header_fmt = workbook.add_format({
            'bold': True, 'font_size': 11, 'font_color': 'white',
            'bg_color': '#1E3A8A', 'border': 1, 'align': 'center', 'valign': 'vcenter'
        })
        cell_fmt = workbook.add_format({
            'font_size': 10, 'border': 1, 'valign': 'vcenter', 'text_wrap': True
        })
        alt_fmt = workbook.add_format({
            'font_size': 10, 'border': 1, 'valign': 'vcenter', 'text_wrap': True, 'bg_color': '#F8FAFC'
        })
        title_fmt = workbook.add_format({
            'bold': True, 'font_size': 14, 'font_color': '#1E3A8A', 'bottom': 2, 'bottom_color': '#D4AF37'
        })

        # Titre
        worksheet.write(0, 0, title, title_fmt)
        worksheet.merge_range(0, 0, 0, len(df_clean.columns)-1, title, title_fmt)

        # Ré-écrire les données avec formatage
        for col_num, col_name in enumerate(df_clean.columns):
            worksheet.write(1, col_num, col_name, header_fmt)
            # Auto-ajustement de la largeur
            max_len = max(df_clean[col_name].astype(str).map(len).max(), len(str(col_name))) + 3
            worksheet.set_column(col_num, col_num, min(max_len, 50))

        for row_num, (_, row) in enumerate(df_clean.iterrows(), start=2):
            fmt = alt_fmt if row_num % 2 == 0 else cell_fmt
            for col_num, val in enumerate(row):
                worksheet.write(row_num, col_num, val, fmt)

        # Figer l'en-tête
        worksheet.freeze_panes(2, 0)

        # Ajouter un onglet récap si pertinent
        if len(df_clean) > 0:
            recap = pd.DataFrame({
                'Metrique': ['Total lignes', 'Date generation', 'Source'],
                'Valeur': [len(df_clean), datetime.now().strftime('%d/%m/%Y %H:%M'), 'Plateforme EDT UDL']
            })
            recap.to_excel(writer, index=False, sheet_name='Recap')
            ws_recap = writer.sheets['Recap']
            ws_recap.set_column(0, 0, 20)
            ws_recap.set_column(1, 1, 30)

    buffer.seek(0)
    return buffer.getvalue()


# =============================================================================
# HUB DE TELECHARGEMENT RAPIDE (A afficher en haut de page apres auth)
# =============================================================================

def render_download_hub(df_global, user_data, is_admin):
    """Affiche un hub de telechargement rapide en haut de page."""

    st.markdown("""
        <style>
        .dl-hub { background: linear-gradient(135deg, #1E3A8A 0%, #3B82F6 100%); 
                   border-radius: 12px; padding: 20px; color: white; margin-bottom: 20px; }
        .dl-hub h3 { margin: 0 0 10px 0; font-size: 18px; }
        .dl-hub p { margin: 0 0 15px 0; opacity: 0.9; font-size: 13px; }
        </style>
    """, unsafe_allow_html=True)

    st.markdown("""
        <div class="dl-hub">
            <h3>📥 Centre de Telechargement Rapide</h3>
            <p>Exportez vos emplois du temps dans le format de votre choix (PDF, HTML, Excel)</p>
        </div>
    """, unsafe_allow_html=True)

    if df_global is None or df_global.empty:
        st.warning("Aucune donnee chargee. Verifiez votre connexion Supabase ou votre fichier Excel.")
        return

    # Déterminer les filtres disponibles
    promos = sorted([p for p in df_global["Promotion"].unique() if p and p != "Non défini"])
    profs = sorted([p for p in df_global["Enseignants"].unique() if p and p != "Non défini"])
    salles = sorted([s for s in df_global["Lieu"].unique() if s and s != "Non défini"])

    col1, col2, col3 = st.columns(3)

    with col1:
        st.markdown("**🎓 Par Promotion**")
        sel_promo = st.selectbox("Choisir promotion", ["Toutes"] + promos, key="hub_promo")

        df_filtre = df_global.copy()
        if sel_promo != "Toutes":
            df_filtre = df_filtre[df_filtre["Promotion"] == sel_promo]

        c1, c2, c3 = st.columns(3)

        # PDF
        pdf_data, err = generate_pro_pdf(df_filtre, f"EDT - {sel_promo}", "Export promotion")
        if pdf_data:
            c1.download_button("📄 PDF", pdf_data, f"EDT_{sel_promo}_2027.pdf", "application/pdf", use_container_width=True)
        else:
            c1.button("📄 PDF", disabled=True, use_container_width=True)

        # HTML
        html_data = generate_pro_html(df_filtre, f"EDT {sel_promo}", "Faculte de Genie Electrique - UDL-SBA")
        c2.download_button("🌐 HTML", html_data, f"EDT_{sel_promo}_2027.html", "text/html", use_container_width=True)

        # Excel
        xlsx_data = generate_pro_excel(df_filtre, f"EDT {sel_promo}")
        c3.download_button("📊 Excel", xlsx_data, f"EDT_{sel_promo}_2027.xlsx", 
                          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

    with col2:
        st.markdown("**👤 Par Enseignant**")
        sel_prof = st.selectbox("Choisir enseignant", ["Tous"] + profs, key="hub_prof")

        df_filtre_p = df_global.copy()
        if sel_prof != "Tous":
            df_filtre_p = df_filtre_p[df_filtre_p["Enseignants"].str.contains(sel_prof, case=False, na=False)]

        c1, c2, c3 = st.columns(3)

        pdf_data_p, _ = generate_pro_pdf(df_filtre_p, f"EDT - {sel_prof}", "Export individuel")
        if pdf_data_p:
            c1.download_button("📄 PDF", pdf_data_p, f"EDT_{sel_prof}_2027.pdf", "application/pdf", use_container_width=True, key="dp")
        else:
            c1.button("📄 PDF", disabled=True, use_container_width=True, key="dp")

        html_data_p = generate_pro_html(df_filtre_p, f"EDT {sel_prof}", "Faculte de Genie Electrique - UDL-SBA")
        c2.download_button("🌐 HTML", html_data_p, f"EDT_{sel_prof}_2027.html", "text/html", use_container_width=True, key="dh")

        xlsx_data_p = generate_pro_excel(df_filtre_p, f"EDT {sel_prof}")
        c3.download_button("📊 Excel", xlsx_data_p, f"EDT_{sel_prof}_2027.xlsx", 
                          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, key="dx")

    with col3:
        st.markdown("**🏢 Par Salle**")
        sel_salle = st.selectbox("Choisir salle", ["Toutes"] + salles, key="hub_salle")

        df_filtre_s = df_global.copy()
        if sel_salle != "Toutes":
            df_filtre_s = df_filtre_s[df_filtre_s["Lieu"] == sel_salle]

        c1, c2, c3 = st.columns(3)

        pdf_data_s, _ = generate_pro_pdf(df_filtre_s, f"Planning - {sel_salle}", "Export salle")
        if pdf_data_s:
            c1.download_button("📄 PDF", pdf_data_s, f"Planning_{sel_salle}_2027.pdf", "application/pdf", use_container_width=True, key="sp")
        else:
            c1.button("📄 PDF", disabled=True, use_container_width=True, key="sp")

        html_data_s = generate_pro_html(df_filtre_s, f"Planning {sel_salle}", "Faculte de Genie Electrique - UDL-SBA")
        c2.download_button("🌐 HTML", html_data_s, f"Planning_{sel_salle}_2027.html", "text/html", use_container_width=True, key="sh")

        xlsx_data_s = generate_pro_excel(df_filtre_s, f"Planning {sel_salle}")
        c3.download_button("📊 Excel", xlsx_data_s, f"Planning_{sel_salle}_2027.xlsx", 
                          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, key="sx")

    # --- Export Global (Admin uniquement) ---
    if is_admin:
        st.divider()
        st.markdown("**🌍 Export Global (Admin)**")

        cg1, cg2, cg3, cg4 = st.columns(4)

        # Global PDF
        pdf_g, _ = generate_pro_pdf(df_global, "EDT GLOBAL S2-2027", "Departement d'Electrotechnique - Toutes promotions")
        if pdf_g:
            cg1.download_button("📄 PDF Global", pdf_g, "EDT_GLOBAL_S2_2027.pdf", "application/pdf", use_container_width=True)

        # Global HTML
        html_g = generate_pro_html(df_global, "EDT Global S2-2027", "Departement d'Electrotechnique - FGE/UDL-SBA")
        cg2.download_button("🌐 HTML Global", html_g, "EDT_GLOBAL_S2_2027.html", "text/html", use_container_width=True)

        # Global Excel
        xlsx_g = generate_pro_excel(df_global, "EDT Global S2-2027", "EDT_Global")
        cg3.download_button("📊 Excel Global", xlsx_g, "EDT_GLOBAL_S2_2027.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

        # Pack ZIP (tous les formats)
        import zipfile
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            if pdf_g:
                zf.writestr("EDT_GLOBAL.pdf", pdf_g)
            zf.writestr("EDT_GLOBAL.html", html_g)
            zf.writestr("EDT_GLOBAL.xlsx", xlsx_g)
        cg4.download_button("🗜️ Pack ZIP", zip_buffer.getvalue(), "Pack_EDT_GLOBAL_S2_2027.zip", "application/zip", use_container_width=True)

    st.divider()


# =============================================================================
# RESTE DU CODE (votre code existant, nettoye et integre)
# =============================================================================

import streamlit as st
import pandas as pd
import os
import hashlib
import io
from datetime import datetime
from supabase import create_client
import streamlit as st

# Masquer les elements du menu superieur
hide_st_style = """
            <style>
            #MainMenu {visibility: hidden;}
            header {visibility: hidden;}
            footer {visibility: hidden;}
            .stAppDeployButton {display:none;}
            #stDecoration {display:none;}
            </style>
            """
st.markdown(hide_st_style, unsafe_allow_html=True)

# --- CONFIGURATION DE LA PAGE ---
st.set_page_config(
    page_title="EDT UDL 2027",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- CONNEXION BASE DE DONNEES ---
URL = st.secrets["SUPABASE_URL"]
KEY = st.secrets["SUPABASE_KEY"]
supabase = create_client(URL, KEY)

def hash_pw(password):
    return hashlib.sha256(str.encode(password)).hexdigest()

# --- GESTION DU TEMPS ---
now = datetime.now()
date_str = now.strftime("%d/%m/%Y")
jours_semaine = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]
nom_jour_fr = jours_semaine[now.weekday()]

# --- STYLE CSS DETAILLE ---
st.markdown("""
    <style>
    .main-title { 
        color: #1E3A8A; 
        text-align: center; 
        font-family: 'serif'; 
        font-weight: bold; 
        border-bottom: 3px solid #D4AF37; 
        padding-bottom: 15px; 
        font-size: 18px; 
        margin-top: 5px;
    }
    .portal-badge { 
        background-color: #D4AF37; 
        color: #1E3A8A; 
        padding: 5px 15px; 
        border-radius: 5px; 
        font-weight: bold; 
        text-align: center; 
        margin-bottom: 20px; 
    }
    .date-badge { 
        background-color: #1E3A8A; 
        color: white; 
        padding: 5px 15px; 
        border-radius: 20px; 
        font-size: 12px; 
        float: right; 
    }
    .metric-card { 
        background-color: #f8f9fa; 
        border: 1px solid #1E3A8A; 
        padding: 10px; 
        border-radius: 10px; 
        text-align: center; 
        height: 100%; 
    }
    .stat-container { 
        display: flex; 
        justify-content: space-around; 
        margin: 20px 0; 
        gap: 10px; 
    }
    .stat-box { 
        flex: 1; 
        padding: 15px; 
        border-radius: 12px; 
        color: white; 
        font-weight: bold; 
        text-align: center; 
        font-size: 16px; 
        box-shadow: 2px 2px 5px rgba(0,0,0,0.1); 
    }
    .bg-cours { background: linear-gradient(135deg, #1E3A8A, #3B82F6); }
    .bg-td { background: linear-gradient(135deg, #15803d, #22c55e); }
    .bg-tp { background: linear-gradient(135deg, #b45309, #f59e0b); }

    table { 
        width: 100%; 
        border-collapse: collapse; 
        table-layout: fixed; 
        margin-top: 10px; 
        background-color: white; 
    }
    th { 
        background-color: #1E3A8A !important; 
        color: white !important; 
        border: 1px solid #000; 
        padding: 6px; 
        text-align: center; 
        font-size: 11px; 
    }
    td { 
        border: 1px solid #000; 
        padding: 4px !important; 
        vertical-align: top; 
        text-align: center; 
        background-color: white; 
        height: 95px; 
        font-size: 11px; 
    }
    .separator { 
        border-top: 1px dashed #bbb; 
        margin: 4px 0; 
    }
    </style>
""", unsafe_allow_html=True)

# --- CHARGEMENT DES DONNEES ---
NOM_FICHIER_FIXE = "dataEDT-ELT-S1-2027.xlsx"
NOM_FICHIER_CONTACTS = "Permanents-Vacataires-ELT2-2026-2027.xlsx"

df = None
repertoire_source = {}
repertoire_noms_complets = {}
repertoire_qualites = {}
repertoire_grades = {}

def normalize(s):
    if not s or s == "Non défini": 
        return "vide"
    s = str(s).strip().lower()
    s = s.replace(" ", "").replace("-", "").replace("–", "")
    s = s.replace(":00", "").replace("h00", "h")
    return s

# --- CHARGEMENT DEPUIS SUPABASE ---
def charger_donnees_supabase():
    try:
        reponse = supabase.table("edt_data").select("*").execute()
        if reponse.data:
            df_temp = pd.DataFrame(reponse.data)
            df_temp['h_norm'] = df_temp['Horaire'].apply(normalize)
            df_temp['j_norm'] = df_temp['Jours'].apply(normalize)
            return df_temp
        else:
            colonnes = ['Enseignements', 'Code', 'Enseignants', 'Horaire', 'Jours', 'Lieu', 'Promotion']
            return pd.DataFrame(columns=colonnes)
    except Exception as e:
        st.error(f"Erreur lors du chargement Supabase : {e}")
        return None

df = charger_donnees_supabase()

# --- CHARGEMENT DU REPERTOIRE ---
if os.path.exists(NOM_FICHIER_CONTACTS):
    try:
        df_contacts = pd.read_excel(NOM_FICHIER_CONTACTS)
        df_contacts.columns = [str(c).strip() for c in df_contacts.columns]
        for _, row in df_contacts.iterrows():
            nom_brut = str(row.get('NOM', '')).strip().upper()
            prenom_brut = str(row.get('PRENOM', '')).strip().capitalize()
            email_brut = str(row.get('Email', '')).strip()
            qualite_brute = str(row.get('Qualite', 'Non defini')).strip()
            grade_brut = str(row.get('Grade', 'N/A')).strip()
            if nom_brut:
                if email_brut and email_brut.lower() != 'nan':
                    repertoire_source[nom_brut] = email_brut
                repertoire_noms_complets[nom_brut] = f"{nom_brut} {prenom_brut}"
                repertoire_qualites[nom_brut] = qualite_brute
                repertoire_grades[nom_brut] = grade_brut
    except Exception as e:
        st.error(f"Erreur lors de la lecture du fichier contacts: {e}")

# --- CHARGEMENT EXCEL LOCAL (FALLBACK) ---
if df is None or df.empty:
    if os.path.exists(NOM_FICHIER_FIXE):
        df = pd.read_excel(NOM_FICHIER_FIXE)
        df.columns = [str(c).strip() for c in df.columns]
        colonnes_cles = ['Enseignements', 'Code', 'Enseignants', 'Horaire', 'Jours', 'Lieu', 'Promotion']
        for col in colonnes_cles:
            if col in df.columns:
                df[col] = df[col].fillna("Non defini").astype(str).str.strip()
            else:
                df[col] = "Non defini"
        df['h_norm'] = df['Horaire'].apply(normalize)
        df['j_norm'] = df['Jours'].apply(normalize)

# --- SYSTEME D'AUTH ---
if "user_data" not in st.session_state:
    st.session_state["user_data"] = None

if not st.session_state["user_data"]:
    st.markdown("<h1 class='main-title'>🏛️ DEPARTEMENT D'ELECTROTECHNIQUE - UDL SBA</h1>", unsafe_allow_html=True)
    t_conn, t_ins, t_adm = st.tabs(["🔑 Connexion", "📝 Inscription", "🛡️ Admin"])

    with t_conn:
        email_input = st.text_input("Adresse Email", key="login_email")
        pass_input = st.text_input("Mot de passe", type="password", key="login_pass")
        if st.button("Se connecter au portail", use_container_width=True):
            result = supabase.table("enseignants_auth").select("*").eq("email", email_input).eq("password_hash", hash_pw(pass_input)).execute()
            if result.data:
                st.session_state["user_data"] = result.data[0]
                st.rerun()
            else:
                st.error("Email ou mot de passe incorrect.")

    with t_ins:
        st.subheader("📝 Creer un nouveau compte Enseignant")
        noms_possibles = sorted(df["Enseignants"].unique()) if df is not None else []
        col1, col2 = st.columns(2)
        with col1:
            new_nom = st.selectbox("Selectionnez votre nom (dans l'EDT)", noms_possibles)
            new_email = st.text_input("Votre adresse Email")
        with col2:
            statut_user = st.radio("Statut de l'enseignant", ["Permanent", "Vacataire"], horizontal=True)
            new_phone = ""
            if statut_user == "Vacataire":
                new_phone = st.text_input("📱 Numero de telephone (Obligatoire)", placeholder="06XXXXXXXX")
        st.divider()
        c_p1, c_p2 = st.columns(2)
        with c_p1:
            new_pass = st.text_input("Choisissez un mot de passe", type="password")
        with c_p2:
            confirm_pass = st.text_input("Confirmez le mot de passe", type="password")

        if st.button("Creer mon compte", use_container_width=True, type="primary"):
            if not new_email or not new_pass:
                st.warning("Veuillez remplir les champs obligatoires.")
            elif statut_user == "Vacataire" and not new_phone:
                st.error("Le numero de telephone est requis pour les vacataires.")
            elif new_pass != confirm_pass:
                st.error("Les mots de passe ne correspondent pas.")
            else:
                check = supabase.table("enseignants_auth").select("email").eq("email", new_email).execute()
                if check.data:
                    st.error("Cet email est deja utilise.")
                else:
                    data_ins = {
                        "nom_officiel": new_nom,
                        "email": new_email,
                        "password_hash": hash_pw(new_pass),
                        "role": "enseignant",
                        "statut": statut_user,
                        "telephone": new_phone if statut_user == "Vacataire" else None
                    }
                    try:
                        supabase.table("enseignants_auth").insert(data_ins).execute()
                        st.success("✅ Compte cree avec succes ! Connectez-vous maintenant.")
                        st.balloons()
                    except Exception as e:
                        st.error(f"Erreur Supabase : {e}")

    with t_adm:
        code_admin = st.text_input("Code de securite Administration", type="password", key="admin_code")
        if st.button("Acces Administration", use_container_width=True):
            if code_admin == "doctorat2026":
                st.session_state["user_data"] = {
                    "nom_officiel": "ADMINISTRATEUR", 
                    "role": "admin",
                    "email": "milouafarid@gmail.com"
                }
                st.rerun()
            else:
                st.error("Code admin incorrect.")

# --- GARDIEN DE SESSION ---
user = st.session_state.get("user_data")
if user is None:
    st.stop() 

is_admin = user.get("role") == "admin"

# =============================================================================
# >>>>> HUB DE TELECHARGEMENT RAPIDE (NOUVEAU - EN HAUT DE PAGE) <<<<<
# =============================================================================
st.markdown("---")
render_download_hub(df, user, is_admin)

# =============================================================================
# SUITE DE VOTRE CODE (Barre laterale, Logique principale, etc.)
# =============================================================================

horaires_list = [
    "8h - 9h", "8h - 9h30", "8h - 10h", "9h - 10h", "9h30 - 11h", 
    "10h - 11h", "11h - 12h", "11h - 12h30", 
    "12h - 13h", "12h30 - 14h", "13h - 14h30", "14h - 15h30", "14h - 16h", "15h30 - 17h"
]
jours_list = ["Dimanche", "Lundi", "Mardi", "Mercredi", "Jeudi"]
map_h = {normalize(h): h for h in horaires_list}
map_j = {normalize(j): j for j in jours_list}

# --- BARRE LATERALE ---
with st.sidebar:
    st.header(f"👤 {user.get('nom_officiel', 'Utilisateur')}")
    portail = st.selectbox("🚀 Selectionner Espace", [
        "📖 Emploi du Temps", "📅 Surveillances Examens", 
        "🤖 Generateur Automatique", "👥 Portail Enseignants", 
        "🎓 Portail mise a jour EDT", "📢 Gestion Administrative - Bordereaux & PVs"
    ])
    st.divider()

    mode_view = "Personnel"
    poste_sup = False

    if portail == "📖 Emploi du Temps":
        if is_admin:
            mode_view = st.radio("Vue Administration :", ["Promotion", "Enseignant", "🏢 Planning Salles", "🚩 Verificateur de conflits", "✍️ Editeur de donnees"])
        else:
            mode_view = "Personnel"
        poste_sup = st.checkbox("Poste Superieur (Decharge 3h)")

    if st.button("🚪 Deconnexion du compte"):
        st.session_state["user_data"] = None
        st.rerun()

# [NOTE : Le reste de votre code original continue ici...]
# Vous pouvez conserver toute votre logique existante (EDT, Surveillances, etc.)
# Les fonctions generate_pro_pdf / generate_pro_html / generate_pro_excel 
# peuvent etre reutilisees partout ou vous generez des fichiers.

# Pour remplacer vos anciens boutons de telechargement dans le code, utilisez :
#   pdf_bytes, _ = generate_pro_pdf(df_export, "Titre", "Sous-titre")
#   st.download_button("📄 PDF", pdf_bytes, "nom.pdf", "application/pdf")
#
#   html_str = generate_pro_html(df_export, "Titre", "Sous-titre")  
#   st.download_button("🌐 HTML", html_str, "nom.html", "text/html")
#
#   xlsx_bytes = generate_pro_excel(df_export, "Titre")
#   st.download_button("📊 Excel", xlsx_bytes, "nom.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
