import streamlit as st
import pandas as pd
import os
import hashlib
import io
from datetime import datetime
from supabase import create_client
import streamlit as st

# =============================================================================
# FONCTIONS UTILITAIRES PRO POUR L'EXPORT (PDF / HTML / EXCEL)
# =============================================================================

def sanitize_for_pdf(text):
    """Nettoie le texte pour fpdf (latin-1) en remplacant les caracteres problematiques."""
    if text is None or pd.isna(text):
        return ""
    text = str(text)
    replacements = {
        "'": "'", "'": "'", """: "\"", """: "\"", "–": "-", "—": "-",
        "…": "...", "«": "\"", "»": "\"", "œ": "oe", "Œ": "OE",
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
    return text.encode('latin-1', 'ignore').decode('latin-1')


def generate_pro_pdf(df_source, title, subtitle="", orientation="L"):
    """Genere un PDF professionnel avec fpdf. Retourne des bytes utilisables par st.download_button."""
    try:
        from fpdf import FPDF
    except ImportError:
        return None, "fpdf non installe"

    class ProPDF(FPDF):
        def header(self):
            self.set_font('Arial', 'B', 9)
            self.set_text_color(30, 58, 138)
            header_text = sanitize_for_pdf("Plateforme de gestion des EDTs-S2-2027 - Departement d'Electrotechnique - FGE/UDL-SBA")
            self.cell(0, 6, header_text, 0, 1, 'C')
            self.set_draw_color(212, 175, 55)
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

    pdf.set_font("Arial", "B", 16)
    pdf.set_text_color(30, 58, 138)
    pdf.cell(0, 10, sanitize_for_pdf(title), 0, 1, "C")

    if subtitle:
        pdf.set_font("Arial", "I", 10)
        pdf.set_text_color(100, 100, 100)
        pdf.cell(0, 6, sanitize_for_pdf(subtitle), 0, 1, "C")

    pdf.ln(5)

    if df_source is not None and not df_source.empty:
        df_clean = df_source.fillna("").astype(str)
        cols = list(df_clean.columns)
        n_cols = len(cols)
        page_w = pdf.w - 20
        col_w = page_w / n_cols if n_cols > 0 else page_w

        pdf.set_font("Arial", "B", 8)
        pdf.set_fill_color(30, 58, 138)
        pdf.set_text_color(255, 255, 255)

        for col in cols:
            pdf.cell(col_w, 8, sanitize_for_pdf(str(col)), 1, 0, "C", True)
        pdf.ln()

        pdf.set_font("Arial", "", 7)
        pdf.set_text_color(0, 0, 0)

        for idx, row in df_clean.iterrows():
            if idx % 2 == 0:
                pdf.set_fill_color(248, 250, 252)
            else:
                pdf.set_fill_color(255, 255, 255)

            for val in row:
                cell_text = sanitize_for_pdf(str(val))
                if len(cell_text) > 50:
                    cell_text = cell_text[:47] + "..."
                pdf.cell(col_w, 6, cell_text, 1, 0, "L", True)
            pdf.ln()

    pdf.ln(5)
    pdf.set_font("Arial", "I", 8)
    pdf.set_text_color(150, 150, 150)
    pdf.cell(0, 5, sanitize_for_pdf(f"Document genere le {datetime.now().strftime('%d/%m/%Y a %H:%M')}"), 0, 0, "R")

    # CORRECTION CRITIQUE : convertir explicitement en bytes pour Streamlit
    return bytes(pdf.output()), None


def generate_pro_html(df_source, title, subtitle=""):
    """Genere un HTML professionnel et responsive. Retourne une str."""
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
                <span>Genere le {datetime.now().strftime('%d/%m/%Y a %H:%M')}</span>
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
    """Genere un Excel professionnel avec xlsxwriter. Retourne des bytes."""
    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df_clean = df_source.fillna("").astype(str) if df_source is not None else pd.DataFrame()
        df_clean.to_excel(writer, index=False, sheet_name=sheet_name)

        workbook = writer.book
        worksheet = writer.sheets[sheet_name]

        header_fmt = workbook.add_format({{
            'bold': True, 'font_size': 11, 'font_color': 'white',
            'bg_color': '#1E3A8A', 'border': 1, 'align': 'center', 'valign': 'vcenter'
        }})
        cell_fmt = workbook.add_format({{
            'font_size': 10, 'border': 1, 'valign': 'vcenter', 'text_wrap': True
        }})
        alt_fmt = workbook.add_format({{
            'font_size': 10, 'border': 1, 'valign': 'vcenter', 'text_wrap': True, 'bg_color': '#F8FAFC'
        }})
        title_fmt = workbook.add_format({{
            'bold': True, 'font_size': 14, 'font_color': '#1E3A8A', 'bottom': 2, 'bottom_color': '#D4AF37'
        }})

        worksheet.write(0, 0, title, title_fmt)
        worksheet.merge_range(0, 0, 0, len(df_clean.columns)-1, title, title_fmt)

        for col_num, col_name in enumerate(df_clean.columns):
            worksheet.write(1, col_num, col_name, header_fmt)
            max_len = max(df_clean[col_name].astype(str).map(len).max(), len(str(col_name))) + 3
            worksheet.set_column(col_num, col_num, min(max_len, 50))

        for row_num, (_, row) in enumerate(df_clean.iterrows(), start=2):
            fmt = alt_fmt if row_num % 2 == 0 else cell_fmt
            for col_num, val in enumerate(row):
                worksheet.write(row_num, col_num, val, fmt)

        worksheet.freeze_panes(2, 0)

        if len(df_clean) > 0:
            recap = pd.DataFrame({{
                'Metrique': ['Total lignes', 'Date generation', 'Source'],
                'Valeur': [len(df_clean), datetime.now().strftime('%d/%m/%Y %H:%M'), 'Plateforme EDT UDL']
            }})
            recap.to_excel(writer, index=False, sheet_name='Recap')
            ws_recap = writer.sheets['Recap']
            ws_recap.set_column(0, 0, 20)
            ws_recap.set_column(1, 1, 30)

    buffer.seek(0)
    return buffer.getvalue()


def render_download_hub(df_global, user_data, is_admin):
    """Affiche un hub de telechargement rapide en haut de page."""
    st.markdown("""
        <style>
        .dl-hub {{ background: linear-gradient(135deg, #1E3A8A 0%, #3B82F6 100%); 
                   border-radius: 12px; padding: 20px; color: white; margin-bottom: 20px; }}
        .dl-hub h3 {{ margin: 0 0 10px 0; font-size: 18px; }}
        .dl-hub p {{ margin: 0 0 15px 0; opacity: 0.9; font-size: 13px; }}
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

    promos = sorted([p for p in df_global["Promotion"].unique() if p and p != "Non defini"])
    profs = sorted([p for p in df_global["Enseignants"].unique() if p and p != "Non defini"])
    salles = sorted([s for s in df_global["Lieu"].unique() if s and s != "Non defini"])

    col1, col2, col3 = st.columns(3)

    with col1:
        st.markdown("**🎓 Par Promotion**")
        sel_promo = st.selectbox("Choisir promotion", ["Toutes"] + promos, key="hub_promo")
        df_filtre = df_global.copy()
        if sel_promo != "Toutes":
            df_filtre = df_filtre[df_filtre["Promotion"] == sel_promo]
        c1, c2, c3 = st.columns(3)
        pdf_data, err = generate_pro_pdf(df_filtre, f"EDT - {sel_promo}", "Export promotion")
        if pdf_data is not None:
            c1.download_button("📄 PDF", pdf_data, f"EDT_{sel_promo}_2027.pdf", "application/pdf", use_container_width=True)
        else:
            c1.button("📄 PDF", disabled=True, use_container_width=True)
        html_data = generate_pro_html(df_filtre, f"EDT {sel_promo}", "Faculte de Genie Electrique - UDL-SBA")
        c2.download_button("🌐 HTML", html_data, f"EDT_{sel_promo}_2027.html", "text/html", use_container_width=True)
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
        if pdf_data_p is not None:
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
        if pdf_data_s is not None:
            c1.download_button("📄 PDF", pdf_data_s, f"Planning_{sel_salle}_2027.pdf", "application/pdf", use_container_width=True, key="sp")
        else:
            c1.button("📄 PDF", disabled=True, use_container_width=True, key="sp")
        html_data_s = generate_pro_html(df_filtre_s, f"Planning {sel_salle}", "Faculte de Genie Electrique - UDL-SBA")
        c2.download_button("🌐 HTML", html_data_s, f"Planning_{sel_salle}_2027.html", "text/html", use_container_width=True, key="sh")
        xlsx_data_s = generate_pro_excel(df_filtre_s, f"Planning {sel_salle}")
        c3.download_button("📊 Excel", xlsx_data_s, f"Planning_{sel_salle}_2027.xlsx", 
                          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, key="sx")

    if is_admin:
        st.divider()
        st.markdown("**🌍 Export Global (Admin)**")
        cg1, cg2, cg3, cg4 = st.columns(4)
        pdf_g, _ = generate_pro_pdf(df_global, "EDT GLOBAL S2-2027", "Departement d'Electrotechnique - Toutes promotions")
        if pdf_g is not None:
            cg1.download_button("📄 PDF Global", pdf_g, "EDT_GLOBAL_S2_2027.pdf", "application/pdf", use_container_width=True)
        html_g = generate_pro_html(df_global, "EDT Global S2-2027", "Departement d'Electrotechnique - FGE/UDL-SBA")
        cg2.download_button("🌐 HTML Global", html_g, "EDT_GLOBAL_S2_2027.html", "text/html", use_container_width=True)
        xlsx_g = generate_pro_excel(df_global, "EDT Global S2-2027", "EDT_Global")
        cg3.download_button("📊 Excel Global", xlsx_g, "EDT_GLOBAL_S2_2027.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            if pdf_g is not None:
                zf.writestr("EDT_GLOBAL.pdf", pdf_g)
            zf.writestr("EDT_GLOBAL.html", html_g)
            zf.writestr("EDT_GLOBAL.xlsx", xlsx_g)
        cg4.download_button("🗜️ Pack ZIP", zip_buffer.getvalue(), "Pack_EDT_GLOBAL_S2_2027.zip", "application/zip", use_container_width=True)

    st.divider()


# =============================================================================
# Masquer les éléments du menu supérieur (Share, Star, Edit, etc.)
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

# --- CONNEXION BASE DE DONNÉES ---
URL = st.secrets["SUPABASE_URL"]
KEY = st.secrets["SUPABASE_KEY"]
supabase = create_client(URL, KEY)

def hash_pw(password):
    return hashlib.sha256(str.encode(password)).hexdigest()

# --- GESTION DU TEMPS ---
now = datetime.now()
date_str = now.strftime("%d/%m/%Y")
jours_semaine = [
    "Lundi", "Mardi", "Mercredi", 
    "Jeudi", "Vendredi", "Samedi", "Dimanche"
]
nom_jour_fr = jours_semaine[now.weekday()]

# --- STYLE CSS DÉTAILLÉ ---
st.markdown(f"""
    <style>
    .main-title {{ 
        color: #1E3A8A; 
        text-align: center; 
        font-family: 'serif'; 
        font-weight: bold; 
        border-bottom: 3px solid #D4AF37; 
        padding-bottom: 12px; 
        font-size: 12px; 
        margin-top: 5px;
    }}
    .portal-badge {{ 
        background-color: #D4AF37; 
        color: #1E3A8A; 
        padding: 5px 12px; 
        border-radius: 5px; 
        font-weight: bold; 
        text-align: center; 
        margin-bottom: 12px; 
    }}
    .date-badge {{ 
        background-color: #1E3A8A; 
        color: white; 
        padding: 5px 12px; 
        border-radius: 15px; 
        font-size: 12px; 
        float: right; 
    }}
    .metric-card {{ 
        background-color: #f8f9fa; 
        border: 1px solid #1E3A8A; 
        padding: 10px; 
        border-radius: 10px; 
        text-align: center; 
        height: 100%; 
    }}
    .stat-container {{ 
        display: flex; 
        justify-content: space-around; 
        margin: 15px 0; 
        gap: 10px; 
    }}
    .stat-box {{ 
        flex: 1; 
        padding: 15px; 
        border-radius: 12px; 
        color: white; 
        font-weight: bold; 
        text-align: center; 
        font-size: 16px; 
        box-shadow: 2px 2px 5px rgba(0,0,0,0.1); 
    }}
    .bg-cours {{ background: linear-gradient(135deg, #1E3A8A, #3B82F6); }}
    .bg-td {{ background: linear-gradient(135deg, #15803d, #22c55e); }}
    .bg-tp {{ background: linear-gradient(135deg, #b45309, #f59e0b); }}
    
    table {{ 
        width: 100%; 
        border-collapse: collapse; 
        table-layout: fixed; 
        margin-top: 10px; 
        background-color: white; 
    }}
    th {{ 
        background-color: #1E3A8A !important; 
        color: white !important; 
        border: 1px solid #000; 
        padding: 6px; 
        text-align: center; 
        font-size: 11px; 
    }}
    td {{ 
        border: 1px solid #000; 
        padding: 4px !important; 
        vertical-align: top; 
        text-align: center; 
        background-color: white; 
        height: 95px; 
        font-size: 11px; 
    }}
    .separator {{ 
        border-top: 1px dashed #bbb; 
        margin: 4px 0; 
    }}
    </style>
""", unsafe_allow_html=True)

# --- CHARGEMENT DES DONNÉES ---
NOM_FICHIER_FIXE = "dataEDT-ELT-S1-2027.xlsx"
NOM_FICHIER_CONTACTS = "Permanents-Vacataires-ELT2-2026-2027.xlsx"

df = None
repertoire_source = {}        # Pour stocker les Emails : { "NOM": "email" }
repertoire_noms_complets = {} # Pour stocker l'affichage : { "NOM": "NOM Prénom" }

def normalize(s):
    if not s or s == "Non défini": 
        return "vide"
    s = str(s).strip().lower()
    s = s.replace(" ", "").replace("-", "").replace("–", "")
    s = s.replace(":00", "").replace("h00", "h")
    return s

# --- ÉTAPE 2 : CHARGEMENT DEPUIS LA BASE DE DONNÉES (SUPABASE) ---
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

# Initialisation du DataFrame principal
df = charger_donnees_supabase()

# --- ÉTAPE 3 : CHARGEMENT DU RÉPERTOIRE (CONTACTS, PRÉNOMS, STATUTS & GRADES) ---
repertoire_qualites = {} 
repertoire_grades = {} # Nouveau dictionnaire pour stocker le Grade (MCA, Prof, etc.)

if os.path.exists(NOM_FICHIER_CONTACTS):
    try:
        df_contacts = pd.read_excel(NOM_FICHIER_CONTACTS)
        # Nettoyage strict des noms de colonnes
        df_contacts.columns = [str(c).strip() for c in df_contacts.columns]
        
        for _, row in df_contacts.iterrows():
            nom_brut = str(row.get('NOM', '')).strip().upper()
            prenom_brut = str(row.get('PRÉNOM', '')).strip().capitalize()
            email_brut = str(row.get('Email', '')).strip()
            
            # Récupération de la colonne 'Qualité' (Statut)
            qualite_brute = str(row.get('Qualité', 'Non défini')).strip()
            
            # Récupération de la colonne 'Grade'
            grade_brut = str(row.get('Grade', 'N/A')).strip()
            
            if nom_brut:
                # 1. Dictionnaire pour les Emails (Inscription)
                if email_brut and email_brut.lower() != 'nan':
                    repertoire_source[nom_brut] = email_brut
                
                # 2. Dictionnaire pour l'affichage complet (Ex: ABID Mohamed)
                repertoire_noms_complets[nom_brut] = f"{nom_brut} {prenom_brut}"
                
                # 3. Dictionnaire pour le statut (Ex: Permanent)
                repertoire_qualites[nom_brut] = qualite_brute
                
                # 4. Dictionnaire pour le Grade (Ex: MCA)
                repertoire_grades[nom_brut] = grade_brut
                
    except Exception as e:
        st.error(f"Erreur lors de la lecture du fichier contacts: {e}")
# --- SYSTÈME D'AUTH ---
if "user_data" not in st.session_state:
    st.session_state["user_data"] = None

if not st.session_state["user_data"]:
    st.markdown("<h1 class='main-title'>🏛️ DÉPARTEMENT D'ÉLECTROTECHNIQUE - UDL SBA</h1>", unsafe_allow_html=True)
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
        st.subheader("📝 Créer un nouveau compte Enseignant")
        # Récupération des noms depuis l'Excel
        noms_possibles = sorted(df["Enseignants"].unique()) if df is not None else []
        
        col1, col2 = st.columns(2)
        with col1:
            new_nom = st.selectbox("Sélectionnez votre nom (dans l'EDT)", noms_possibles)
            new_email = st.text_input("Votre adresse Email")
            
        with col2:
            # Nouveau : Choix du Statut
            statut_user = st.radio("Statut de l'enseignant", ["Permanent", "Vacataire"], horizontal=True)
            
            # Nouveau : Champ téléphone conditionnel
            new_phone = ""
            if statut_user == "Vacataire":
                new_phone = st.text_input("📱 Numéro de téléphone (Obligatoire)", placeholder="06XXXXXXXX")

        st.divider()
        c_p1, c_p2 = st.columns(2)
        with c_p1:
            new_pass = st.text_input("Choisissez un mot de passe", type="password")
        with c_p2:
            confirm_pass = st.text_input("Confirmez le mot de passe", type="password")
        
        if st.button("Créer mon compte", use_container_width=True, type="primary"):
            if not new_email or not new_pass:
                st.warning("Veuillez remplir les champs obligatoires.")
            elif statut_user == "Vacataire" and not new_phone:
                st.error("Le numéro de téléphone est requis pour les vacataires.")
            elif new_pass != confirm_pass:
                st.error("Les mots de passe ne correspondent pas.")
            else:
                # Vérifier si l'email existe déjà
                check = supabase.table("enseignants_auth").select("email").eq("email", new_email).execute()
                if check.data:
                    st.error("Cet email est déjà utilisé.")
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
                        st.success("✅ Compte créé avec succès ! Connectez-vous maintenant.")
                        st.balloons()
                    except Exception as e:
                        st.error(f"Erreur Supabase : {e}")

    with t_adm:
        code_admin = st.text_input("Code de sécurité Administration", type="password", key="admin_code")
        if st.button("Accès Administration", use_container_width=True):
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
# --- CONFIGURATION DE LA PAGE ---
st.set_page_config(
    page_title="EDT UDL 2027",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- CONNEXION BASE DE DONNÉES ---
URL = st.secrets["SUPABASE_URL"]
KEY = st.secrets["SUPABASE_KEY"]
supabase = create_client(URL, KEY)

def hash_pw(password):
    return hashlib.sha256(str.encode(password)).hexdigest()

# --- GESTION DU TEMPS ---
now = datetime.now()
date_str = now.strftime("%d/%m/%Y")
jours_semaine = [
    "Lundi", "Mardi", "Mercredi", 
    "Jeudi", "Vendredi", "Samedi", "Dimanche"
]
nom_jour_fr = jours_semaine[now.weekday()]

# --- STYLE CSS DÉTAILLÉ ---
st.markdown(f"""
    <style>
    .main-title {{ 
        color: #1E3A8A; 
        text-align: center; 
        font-family: 'serif'; 
        font-weight: bold; 
        border-bottom: 3px solid #D4AF37; 
        padding-bottom: 15px; 
        font-size: 18px; 
        margin-top: 5px;
    }}
    .portal-badge {{ 
        background-color: #D4AF37; 
        color: #1E3A8A; 
        padding: 5px 15px; 
        border-radius: 5px; 
        font-weight: bold; 
        text-align: center; 
        margin-bottom: 20px; 
    }}
    .date-badge {{ 
        background-color: #1E3A8A; 
        color: white; 
        padding: 5px 15px; 
        border-radius: 20px; 
        font-size: 12px; 
        float: right; 
    }}
    .metric-card {{ 
        background-color: #f8f9fa; 
        border: 1px solid #1E3A8A; 
        padding: 10px; 
        border-radius: 10px; 
        text-align: center; 
        height: 100%; 
    }}
    .stat-container {{ 
        display: flex; 
        justify-content: space-around; 
        margin: 20px 0; 
        gap: 10px; 
    }}
    .stat-box {{ 
        flex: 1; 
        padding: 15px; 
        border-radius: 12px; 
        color: white; 
        font-weight: bold; 
        text-align: center; 
        font-size: 16px; 
        box-shadow: 2px 2px 5px rgba(0,0,0,0.1); 
    }}
    .bg-cours {{ background: linear-gradient(135deg, #1E3A8A, #3B82F6); }}
    .bg-td {{ background: linear-gradient(135deg, #15803d, #22c55e); }}
    .bg-tp {{ background: linear-gradient(135deg, #b45309, #f59e0b); }}
    
    table {{ 
        width: 100%; 
        border-collapse: collapse; 
        table-layout: fixed; 
        margin-top: 10px; 
        background-color: white; 
    }}
    th {{ 
        background-color: #1E3A8A !important; 
        color: white !important; 
        border: 1px solid #000; 
        padding: 6px; 
        text-align: center; 
        font-size: 11px; 
    }}
    td {{ 
        border: 1px solid #000; 
        padding: 4px !important; 
        vertical-align: top; 
        text-align: center; 
        background-color: white; 
        height: 95px; 
        font-size: 11px; 
    }}
    .separator {{ 
        border-top: 1px dashed #bbb; 
        margin: 4px 0; 
    }}
    </style>
""", unsafe_allow_html=True)

# --- CHARGEMENT DES DONNÉES ---
NOM_FICHIER_FIXE = "dataEDT-ELT-S1-2027.xlsx"
df = None

def normalize(s):
    if not s or s == "Non défini": 
        return "vide"
    s = str(s).strip().lower()
    s = s.replace(" ", "").replace("-", "").replace("–", "")
    s = s.replace(":00", "").replace("h00", "h")
    return s

if os.path.exists(NOM_FICHIER_FIXE):
    df = pd.read_excel(NOM_FICHIER_FIXE)
    df.columns = [str(c).strip() for c in df.columns]
    
    colonnes_cles = [
        'Enseignements', 
        'Code', 
        'Enseignants', 
        'Horaire', 
        'Jours', 
        'Lieu', 
        'Promotion'
    ]
    
    for col in colonnes_cles:
        if col in df.columns: 
            df[col] = df[col].fillna("Non défini").astype(str).str.strip()
        else:
            df[col] = "Non défini"
            
    df['h_norm'] = df['Horaire'].apply(normalize)
    df['j_norm'] = df['Jours'].apply(normalize)

# --- SYSTÈME D'AUTH ---
if "user_data" not in st.session_state:
    st.session_state["user_data"] = None

if not st.session_state["user_data"]:
    st.markdown("<h1 class='main-title'>🏛️ DÉPARTEMENT D'ÉLECTROTECHNIQUE - UDL SBA</h1>", unsafe_allow_html=True)
    t_conn, t_ins, t_adm = st.tabs(["🔑 Connexion", "📝 Inscription", "🛡️ Admin"])
    
    with t_conn:
        email_input = st.text_input("Adresse Email", key="login_email")
        pass_input = st.text_input("Mot de passe", type="password", key="login_pass")
        if st.button("Se connecter au portail"):
            result = supabase.table("enseignants_auth").select("*").eq("email", email_input).eq("password_hash", hash_pw(pass_input)).execute()
            if result.data:
                st.session_state["user_data"] = result.data[0]
                st.rerun()
            else:
                st.error("Email ou mot de passe incorrect.")
                
    with t_ins:
        st.subheader("Créer un nouveau compte Enseignant")
        # On récupère la liste des noms depuis l'Excel pour éviter les erreurs de saisie
        noms_possibles = sorted(df["Enseignants"].unique()) if df is not None else []
        
        new_nom = st.selectbox("Sélectionnez votre nom (tel qu'il apparaît dans l'EDT)", noms_possibles)
        new_email = st.text_input("Votre adresse Email")
        new_pass = st.text_input("Choisissez un mot de passe", type="password")
        confirm_pass = st.text_input("Confirmez le mot de passe", type="password")
        
        if st.button("Créer mon compte"):
            if not new_email or not new_pass:
                st.warning("Veuillez remplir tous les champs.")
            elif new_pass != confirm_pass:
                st.error("Les mots de passe ne correspondent pas.")
            else:
                # Vérifier si l'email existe déjà
                check = supabase.table("enseignants_auth").select("email").eq("email", new_email).execute()
                if check.data:
                    st.error("Cet email est déjà utilisé.")
                else:
                    data_ins = {
                        "nom_officiel": new_nom,
                        "email": new_email,
                        "password_hash": hash_pw(new_pass),
                        "role": "enseignant"
                    }
                    supabase.table("enseignants_auth").insert(data_ins).execute()
                    st.success("✅ Compte créé avec succès ! Vous pouvez maintenant vous connecter.")
                    st.balloons()

    with t_adm:
        code_admin = st.text_input("Code de sécurité Administration", type="password", key="admin_code")
        if st.button("Accès Administration"):
            if code_admin == "doctorat2026":
                # On force l'email ici pour activer vos droits maître
                st.session_state["user_data"] = {
                    "nom_officiel": "ADMINISTRATEUR", 
                    "role": "admin",
                    "email": "milouafarid@gmail.com"  # <--- AJOUTER CETTE LIGNE
                }
                st.rerun()
            else:
                st.error("Code admin incorrect.")
# --- SOLUTIONS AUX ERREURS (Remplace le bloc supprimé) ---
user = st.session_state.get("user_data")

# Le st.stop() est le gardien : si pas de login, on n'affiche pas la suite
if user is None:
    st.stop() 

is_admin = user.get("role") == "admin"

# =============================================================================
# >>>>> HUB DE TELECHARGEMENT RAPIDE (CENTRE DE TELECHARGEMENT) <<<<<
# =============================================================================
st.markdown("---")
render_download_hub(df, user, is_admin)


# 1. Définition précise de votre nouvelle liste d'horaires (14 créneaux)
horaires_list = [
    "8h - 9h", "8h - 9h30", "8h - 10h", "9h - 10h", "9h30 - 11h", 
    "10h - 11h", "11h - 12h", "11h - 12h30", 
    "12h - 13h", "12h30 - 14h", "13h - 14h30", "14h - 15h30", "14h - 16h", "15h30 - 17h"
]

# 2. Définition des jours de la semaine
jours_list = ["Dimanche", "Lundi", "Mardi", "Mercredi", "Jeudi"]

# 3. Mapping pour la normalisation (crucial pour faire le lien avec l'Excel)
# Cela permet de faire correspondre "8h-9h30" (Excel) avec "8h - 9h30" (Affichage)
map_h = {normalize(h): h for h in horaires_list}
map_j = {normalize(j): j for j in jours_list}

# --- BARRE LATÉRALE ---
with st.sidebar:
    # On utilise .get() pour éviter le ch si la donnée est corrompue
    st.header(f"👤 {user.get('nom_officiel', 'Utilisateur')}")
    portail = st.selectbox("🚀 Sélectionner Espace", [
        "📖 Emploi du Temps", "📅 Surveillances Examens", 
        "🤖 Générateur Automatique", "👥 Portail Enseignants", "🎓 Portail mise à jour EDT", "📢 Gestion Administrative - Bordereaux & PVs"
    ])
    st.divider()
    
    mode_view = "Personnel"
    poste_sup = False
    
    if portail == "📖 Emploi du Temps":
        if is_admin:
            mode_view = st.radio("Vue Administration :", ["Promotion", "Enseignant", "🏢 Planning Salles", "🚩 Vérificateur de conflits","✍️ Éditeur de données"])
        else:
            mode_view = "Personnel"
        poste_sup = st.checkbox("Poste Supérieur (Décharge 3h)")
        
    if st.button("🚪 Déconnexion du compte"):
        st.session_state["user_data"] = None
        st.rerun()
# --- ESPACE ÉDITEUR AVANCÉ (ADMIN UNIQUEMENT) ---
# --- ESPACE ÉDITEUR AVANCÉ (ADMIN UNIQUEMENT) ---
if is_admin and mode_view == "✍️ Éditeur de données":
    st.divider()
    st.subheader("✍️ Plateforme de gestion des EDTs-S2-2027-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA")

    # 1. VÉRIFICATION DE L'EXISTENCE DE df
    if 'df' not in locals() and 'df' not in globals():
        st.error("Erreur : Les données (df) n'ont pas été chargées. Veuillez vérifier votre source de données.")
        st.stop() # Arrête l'exécution pour éviter le plantage

    # 2. STRUCTURE ET NETTOYAGE
    cols_format = ['Enseignements', 'Code', 'Enseignants', 'Horaire', 'Jours', 'Lieu', 'Promotion', 'Chevauchement']

    if 'df_admin' not in st.session_state:
        # On s'assure que df est bien un DataFrame valide avant de le copier
        if df is not None:
            temp_df = df.copy()
            for col in cols_format:
                if col not in temp_df.columns:
                    temp_df[col] = ""
                temp_df[col] = temp_df[col].astype(str).replace(['nan', 'None', '<NA>'], '')
            st.session_state.df_admin = temp_df
        else:
            st.warning("Le DataFrame est vide ou non initialisé.")
    # 2. PRÉPARATION DES OPTIONS
    horaires_ref = ["8h - 9h30", "9h30 - 11h", "11h - 12h30", "12h30 - 14h00", "14h00 - 15h30", "15h30 - 17h00"]
    h_existants = [h for h in st.session_state.df_admin["Horaire"].unique() if h and h.strip() != ""]
    liste_horaires = sorted(list(set(h_existants + horaires_ref)))
    jours_std = ["Dimanche", "Lundi", "Mardi", "Mercredi", "Jeudi"]
    promos_existantes = [p for p in st.session_state.df_admin["Promotion"].unique() if p and p.strip() != ""]

    # --- NOUVEAUTÉ : FILTRE DE RECHERCHE ---
    st.markdown("### 🔍 Filtrer par Enseignant")
    search_prof = st.text_input("Tapez le nom de l'enseignant pour filtrer le tableau :", "")

    # Application du filtre
    if search_prof:
        # On filtre les données pour l'affichage
        df_to_edit = st.session_state.df_admin[
            st.session_state.df_admin["Enseignants"].str.contains(search_prof, case=False, na=False)
        ]
        st.info(f"💡 Affichage des cours de : **{search_prof}**. Les modifications ou ajouts ne concernernt que cette sélection.")
    else:
        df_to_edit = st.session_state.df_admin

    # 3. TABLEAU GLOBAL (ÉDITION, AJOUT & DÉTECTION DE CONFLITS)
    st.markdown("### 🌍 Tableau d'édition")
    
    # --- FORMULAIRE D'AJOUT AVEC VÉRIFICATION ---
    with st.expander("➕ Ajouter une nouvelle ligne (Vérification automatique)"):
        with st.form("form_nouvelle_ligne"):
            c1, c2, c3 = st.columns(3)
            with c1:
                n_ensg = st.text_input("📚 Enseignements")
                n_code = st.text_input("🔑 Code")
                n_promo = st.selectbox("🎓 Promotion", options=promos_existantes if promos_existantes else ["M2RE"])
            with c2:
                n_prof = st.text_input("👤 Enseignants")
                n_horaire = st.selectbox("🕒 Horaire", options=liste_horaires)
            with c3:
                n_jour = st.selectbox("📅 Jours", options=jours_std)
                n_lieu = st.text_input("🏢 Lieu (Salle)")
                n_chev = "Non"

            submit_add = st.form_submit_button("🔍 Vérifier et Insérer", use_container_width=True)

            if submit_add:
                # 1. Vérification des conflits avec extraction de la promotion concernée
                conflit_salle = st.session_state.df_admin[
                    (st.session_state.df_admin['Jours'] == n_jour) & 
                    (st.session_state.df_admin['Horaire'] == n_horaire) & 
                    (st.session_state.df_admin['Lieu'] == n_lieu)
                ]
                
                conflit_prof = st.session_state.df_admin[
                    (st.session_state.df_admin['Jours'] == n_jour) & 
                    (st.session_state.df_admin['Horaire'] == n_horaire) & 
                    (st.session_state.df_admin['Enseignants'] == n_prof)
                ]

                if not conflit_salle.empty:
                    # On affiche quelle promotion occupe déjà la salle
                    promo_conflit = conflit_salle.iloc[0]['Promotion']
                    prof_conflit = conflit_salle.iloc[0]['Enseignants']
                    st.error(f"❌ CONFLIT SALLE : La salle {n_lieu} est déjà prise par **{prof_conflit}** pour la promotion **{promo_conflit}**.")
                
                elif not conflit_prof.empty:
                    # On affiche quelle promotion l'enseignant a déjà
                    promo_conflit = conflit_prof.iloc[0]['Promotion']
                    lieu_conflit = conflit_prof.iloc[0]['Lieu']
                    st.error(f"❌ CONFLIT ENSEIGNANT : M. {n_prof} a déjà un cours avec la promotion **{promo_conflit}** en salle {lieu_conflit}.")
                
                else:
                    # --- ÉTAPE 3 : INSERTION RÉELLE DANS LA TABLE SUPABASE ---
                    nouvelle_ligne_db = {
                        "Enseignements": n_ensg,
                        "Code": n_code,
                        "Enseignants": n_prof,
                        "Horaire": n_horaire,
                        "Jours": n_jour,
                        "Lieu": n_lieu,
                        "Promotion": n_promo
                    }
                    
                    try:
                        # 1. Envoi à la base de données Cloud
                        supabase.table("edt_data").insert(nouvelle_ligne_db).execute()
                        
                        st.success(f"✅ Félicitations ! Le cours de {n_ensg} pour la promotion {n_promo} est désormais enregistré dans la base de données Cloud.")
                        
                        # 2. On efface la version temporaire pour forcer le rechargement depuis le Cloud
                        if 'df_admin' in st.session_state:
                            del st.session_state.df_admin
                        
                        # 3. Relance de l'application pour tout mettre à jour
                        st.rerun()
                        
                    except Exception as e:
                        st.error(f"❌ Erreur technique lors de l'enregistrement : {e}")

   

   # --- ÉDITEUR DE TABLEAU (VERSION CORRIGÉE 2027) ---
    st.markdown("### 📝 Modification des données")
    
    # Changement de la clé pour éviter le conflit StreamlitDuplicateElementKey
    edited_df = st.data_editor(
        df_to_edit[cols_format],
        use_container_width=True,
        num_rows="dynamic",
        key="editor_final_unique_v3", 
        column_config={
            "Enseignements": st.column_config.TextColumn("📚 Matière"),
            "Horaire": st.column_config.SelectboxColumn("🕒 Horaire", options=liste_horaires),
            "Jours": st.column_config.SelectboxColumn("📅 Jours", options=jours_std),
            "Promotion": st.column_config.SelectboxColumn("🎓 Promotion", options=promos_existantes if promos_existantes else ["M2RE"]),
            "Chevauchement": st.column_config.TextColumn("⚠️ État Conflit"),
        }
    )

    # Synchronisation intelligente
    if edited_df is not None and not edited_df.equals(df_to_edit[cols_format]):
        if search_prof:
            indices_modifies = df_to_edit.index
            df_others = st.session_state.df_admin.drop(indices_modifies)
            st.session_state.df_admin = pd.concat([df_others, edited_df], ignore_index=True)
        else:
            st.session_state.df_admin = edited_df

    # --- BLOC D'ANALYSE VISUELLE (STYLE PERSONNALISÉ : SALLE/PROF/PROMO) ---
    st.divider()
    st.markdown("### 🔍 Analyse Visuelle des Chevauchements")

    def afficher_grille_anomalie(df_source, type_tri):
        jours_ordre = ["Dimanche", "Lundi", "Mardi", "Mercredi", "Jeudi"]
        horaires_ordre = [
            "8h - 9h", "8h - 9h30", "8h - 10h", "9h - 10h", "9h30 - 11h", 
            "10h - 11h", "11h - 12h", "11h - 12h30", "12h - 13h", 
            "12h30 - 14h", "13h - 14h", "14h - 15h30", "14h - 16h", "15h30 - 17h"
        ]
        
        grid = pd.DataFrame("", index=horaires_ordre, columns=jours_ordre)
        df_temp = df_source.copy()
        
        def format_horaire(h):
            h_str = str(h).replace(" ", "").lower()
            for target in horaires_ordre:
                if h_str == target.replace(" ", "").lower(): return target
            return h

        df_temp['Horaire_Normalise'] = df_temp['Horaire'].apply(format_horaire)
        df_temp['Jours'] = df_temp['Jours'].astype(str).str.strip().str.capitalize()

        # Détection des doublons
        doublons = df_temp.duplicated(subset=['Jours', 'Horaire_Normalise', type_tri], keep=False)
        mask_valid = (df_temp[type_tri].astype(str).str.len() > 1) & (df_temp[type_tri].astype(str).str.lower() != "nan")
        df_conflits = df_temp[doublons & mask_valid].copy()
        
        if not df_conflits.empty:
            for _, row in df_conflits.iterrows():
                idx_h = row['Horaire_Normalise']
                col_j = row['Jours']
                
                if idx_h in horaires_ordre and col_j in jours_ordre:
                    # Formatage selon votre exemple
                    salle_label = f"🏢 {row['Lieu']}"
                    prof_label = f"(Prof: {row['Enseignants']})"
                    promo_label = f"🎓 {row['Promotion']}"
                    matiere_label = f"📚 {row['Enseignements']}"
                    heure_label = f"🕒 {row['Horaire']}"

                    cell_html = (
                        f"<div style='color: #b91c1c; font-size: 0.75rem; border-left: 4px solid #b91c1c; "
                        f"padding: 6px; margin-bottom: 8px; background-color: #fff5f5; line-height: 1.3;'>"
                        f"<b>{salle_label}</b><br>"
                        f"{prof_label}<br>"
                        f"<b>{promo_label}</b><br>"
                        f"{matiere_label}<br>"
                        f"{heure_label}"
                        f"</div>"
                    )
                    
                    prev = grid.at[idx_h, col_j]
                    grid.at[idx_h, col_j] = (prev + cell_html) if prev else cell_html
            
            st.write(grid.to_html(escape=False, justify='center'), unsafe_allow_html=True)
        else:
            st.success(f"✅ Aucun conflit de type **{type_tri}** détecté.")

    # Onglets de navigation
    t_salle, t_prof, t_promo = st.tabs(["🏢 Conflits Salles", "👤 Conflits Enseignants", "🎓 Conflits Promotions"])
    
    with t_salle:
        afficher_grille_anomalie(st.session_state.df_admin, "Lieu")
    with t_prof:
        afficher_grille_anomalie(st.session_state.df_admin, "Enseignants")
    with t_promo:
        afficher_grille_anomalie(st.session_state.df_admin, "Promotion")

    # 4. SAUVEGARDE ET EXPORT AVEC RAPPORT DE CONFLITS DYNAMIQUE
    st.write("---")
    c1, c2, c3 = st.columns(3)
    
    with c1:
        if st.button("💾 Enregistrer sur Serveur", type="primary", use_container_width=True):
            try:
                st.session_state.df_admin[cols_format].to_excel(NOM_FICHIER_FIXE, index=False)
                st.success("✅ Modifications enregistrées sur le serveur !")
                st.balloons()
            except Exception as e:
                st.error(f"Erreur d'écriture : {e}")

    with c2:
        if st.button("🔄 Réinitialiser l'éditeur", use_container_width=True):
            if 'df_admin' in st.session_state:
                del st.session_state.df_admin
            st.rerun()

    with c3:
        import io
        import re
        df_complet = st.session_state.df_admin.copy()
        conflits_list = []

        # 1. Détection des doublons sur les colonnes clés
        doublons_salle = df_complet.duplicated(subset=['Jours', 'Horaire', 'Lieu'], keep=False) & (df_complet['Lieu'].astype(str).str.len() > 1)
        doublons_prof = df_complet.duplicated(subset=['Jours', 'Horaire', 'Enseignants'], keep=False) & (df_complet['Enseignants'] != "ND") & (df_complet['Enseignants'] != "")
        doublons_promo = df_complet.duplicated(subset=['Jours', 'Horaire', 'Promotion'], keep=False) & (df_complet['Promotion'] != "")

        # 2. Construction du rapport ligne par ligne pour garantir l'affichage de la Promotion
        for i, row in df_complet.iterrows():
            #--- CONFLIT SALLE ---
            if doublons_salle[i]:
                conflits_list.append({
                    "Type de Conflit": "❌ SALLE OCCUPÉE",
                    "Promotion": row['Promotion'],
                    "Intervenant/Salle": row['Lieu'],
                    "Jour": row['Jours'],
                    "Horaire": row['Horaire'],
                    "Détails": f"La salle {row['Lieu']} est réservée par plusieurs groupes."
                })
            
            #--- CONFLIT ENSEIGNANT ---
            if doublons_prof[i]:
                conflits_list.append({
                    "Type de Conflit": "👤 CONFLIT ENSEIGNANT",
                    "Promotion": row['Promotion'],
                    "Intervenant/Salle": row['Enseignants'],
                    "Jour": row['Jours'],
                    "Horaire": row['Horaire'],
                    "Détails": f"L'enseignant {row['Enseignants']} a deux cours en même temps."
                })

            #--- CONFLIT PROMOTION (Chevauchement de cours) ---
            if doublons_promo[i]:
                conflits_list.append({
                    "Type de Conflit": "⚠️ CONFLIT PROMOTION",
                    "Promotion": row['Promotion'],
                    "Intervenant/Salle": row['Promotion'],
                    "Jour": row['Jours'],
                    "Horaire": row['Horaire'],
                    "Détails": "Cette promotion a plusieurs enseignements affectés au même créneau."
                })

        # Création du DataFrame final
        df_rapport = pd.DataFrame(conflits_list).drop_duplicates()

        # 3. Génération du fichier Excel
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            # Onglet 1 : Emploi du Temps
            df_complet[cols_format].to_excel(writer, sheet_name='Emploi du Temps', index=False)
            
            # Onglet 2 : Rapport des Conflits
            if not df_rapport.empty:
                # Disposition demandée : Promotion bien isolée en 2ème colonne
                colonnes_rapport = ["Type de Conflit", "Promotion", "Intervenant/Salle", "Jour", "Horaire", "Détails"]
                df_rapport[colonnes_rapport].to_excel(writer, sheet_name='Rapport Conflits', index=False)
                
                # Mise en forme (largeur colonnes)
                worksheet = writer.sheets['Rapport Conflits']
                for idx, col in enumerate(colonnes_rapport):
                    worksheet.set_column(idx, idx, 22)
            else:
                pd.DataFrame({"Résultat": ["Aucun conflit détecté"]}).to_excel(writer, sheet_name='Rapport Conflits', index=False)

        st.download_button(
            label="📥 Télécharger le Rapport d'Erreurs Excel",
            data=buffer.getvalue(),
            file_name=f"Rapport_Conflits_EDT_2027.xlsx",
            mime="application/vnd.ms-excel",
            use_container_width=True
        ) 


# --- EN-TÊTE HARMONISÉ (LOGO + TITRE + DATE) ---
col_logo, col_titre, col_date = st.columns([1, 5, 1.2])

with col_logo:
    try:
        st.image("logo.PNG", width=90)
    except:
        st.markdown("🏛️") # Secours si le fichier est manquant

with col_titre:
    st.markdown("<h1 class='main-title' style='border-bottom: none; margin-top: 0;'>Plateforme de gestion des emplois du temps 2026-2027-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA</h1>", unsafe_allow_html=True)

with col_date:
    st.markdown(f"<div class='date-badge' style='float: right;'>📅 {nom_jour_fr}<br>{date_str}</div>", unsafe_allow_html=True)

# Ligne dorée décorative et badge du mode
st.markdown("<div style='border-bottom: 3px solid #D4AF37; margin-bottom: 10px;'></div>", unsafe_allow_html=True)
st.markdown(f"<div class='portal-badge'>MODE ACTIF : {portail.upper()}</div>", unsafe_allow_html=True)

# --- LOGIQUE PRINCIPALE ---
if df is not None:
    if portail == "📖 Emploi du Temps":
        # Sélection de la cible (Enseignant ou Personnel)
        if mode_view == "Personnel" or (is_admin and mode_view == "Enseignant"):
            
            if mode_view == "Personnel":
                # Pour l'enseignant connecté, on utilise ses informations de session
                cible = user['nom_officiel']
                # On essaie de récupérer le nom complet pour l'affichage, sinon on garde le nom officiel
                nom_affichage_complet = repertoire_noms_complets.get(cible.strip().upper(), cible)
            else:
                # 1. On récupère les noms uniques (NOM seulement) présents dans le fichier EDT
                noms_bruts = sorted(df["Enseignants"].unique())
                
                # 2. On prépare la liste d'affichage : "NOM Prénom" si trouvé dans le fichier contacts, sinon NOM seul
                options_affichage = [repertoire_noms_complets.get(n.strip().upper(), n) for n in noms_bruts]
                
                # 3. On crée un mapping inverse pour retrouver le NOM brut à partir du choix "NOM Prénom"
                # Exemple : {"ABID Mohamed": "ABID"}
                inverse_map = {repertoire_noms_complets.get(n.strip().upper(), n): n for n in noms_bruts}
                
                # 4. Affichage de la liste déroulante avec les noms complets
                choix_utilisateur = st.selectbox(
                    "Sélectionner l'Enseignant :", 
                    options=options_affichage,
                    index=0
                )
                
                # 5. 'cible' récupère la valeur NOM brute (nécessaire pour filtrer le DataFrame Excel)
                cible = inverse_map[choix_utilisateur]
                # 6. 'nom_affichage_complet' servira pour le titre du bilan
                nom_affichage_complet = choix_utilisateur
            
            # --- FILTRAGE ET CALCULS ---
            # Filtrage des données basé sur le NOM (variable 'cible')
            df_f = df[df["Enseignants"].str.contains(cible, case=False, na=False)].copy()
            
            # Définition des types pour le calcul (COURS, TD ou TP)
            df_f['Type'] = df_f['Code'].apply(lambda x: "COURS" if "COURS" in str(x).upper() else ("TD" if "TD" in str(x).upper() else "TP"))
            
            # Suppression des doublons basés sur le jour et l'heure pour éviter les erreurs de calcul
            df_u = df_f.drop_duplicates(subset=['j_norm', 'h_norm'])
            
            # --- 1. CALCUL DES COMPTEURS (LOGIQUE BILAN DIRECT) ---
            nb_cours = len(df_u[df_u['Type'] == 'COURS'])
            nb_td    = len(df_u[df_u['Type'] == 'TD'])
            nb_tp    = len(df_u[df_u['Type'] == 'TP'])

            # Le seuil réglementaire (3.0 si poste sup, sinon 6.0)
            seuil_obligatoire = 3.0 if poste_sup else 6.0
            
            # Calcul de la charge totale en Équivalent (Cours = 1.5, TD/TP = 1.0)
            charge_totale_eq = (nb_cours * 1.5) + (nb_td + nb_tp)
            
            # Calcul du bilan (Déficit ou Heures Sup) par rapport au seuil
            delta_eq = charge_totale_eq - seuil_obligatoire
            
            # Calcul de la valeur brute en heures pour l'affichage
            h_sup = delta_eq * 1.5
            
            # --- LOGIQUE DE CONVERSION EN HEURES ET MINUTES ---
            abs_h_sup = abs(h_sup)
            heures_entieres = int(abs_h_sup)
            minutes_restantes = int((abs_h_sup - heures_entieres) * 60)
            
            # Formatage du signe (+ ou -) et de la chaîne de caractères
            signe_str = "+" if h_sup >= 0 else "-"
            h_sup_formattee = f"{signe_str}{heures_entieres}h{minutes_restantes:02d}"
            
            # Charge effective enseignée (Nombre total de séances réelles * 1.5h)
            charge_effective = (nb_cours + nb_td + nb_tp) * 1.5

            # --- 2. RÉCUPÉRATION DES INFOS D'AFFICHAGE (GRADE & STATUT) ---
            # On récupère les infos depuis les dictionnaires mis à jour à l'Etape 3
            statut_enseignant = repertoire_qualites.get(cible.strip().upper(), "Statut inconnu")
            grade_enseignant = repertoire_grades.get(cible.strip().upper(), "Grade inconnu")
            
            # Définition des couleurs des badges
            # Vert si Permanent, Orange si Vacataire/Autre
            color_statut = "#2ecc71" if "PERMANENT" in statut_enseignant.upper() else "#e67e22"
            color_grade = "#3498db" # Bleu professionnel pour le Grade
            
            # --- 3. AFFICHAGE DU TITRE AVEC LES DEUX BADGES ---
            st.markdown(f"""
                <div style="display: flex; align-items: center; flex-wrap: wrap; gap: 10px; margin-bottom: 20px;">
                    <h3 style="margin: 0;">📊 Charge Horaire hebdomadaire : {nom_affichage_complet}</h3>
                    <span style="background-color: {color_grade}; color: white; padding: 3px 12px; 
                                 border-radius: 15px; font-size: 0.8em; font-weight: bold; border: 1px solid rgba(255,255,255,0.1);">
                        {grade_enseignant}
                    </span>
                    <span style="background-color: {color_statut}; color: white; padding: 3px 12px; 
                                 border-radius: 15px; font-size: 0.8em; font-weight: bold; border: 1px solid rgba(255,255,255,0.1);">
                        {statut_enseignant}
                    </span>
                </div>
            """, unsafe_allow_html=True)
            
            # --- 4. AFFICHAGE DES COMPTEURS (COURS, TD, TP) ---
            st.markdown(f"""
                <div class="stat-container">
                    <div class="stat-box bg-cours">📘 {nb_cours} Cours</div>
                    <div class="stat-box bg-td">📗 {nb_td} TD</div>
                    <div class="stat-box bg-tp">🔴 {nb_tp} TP</div>
                </div>
            """, unsafe_allow_html=True)
            
            # Affichage des détails de charge (Optionnel selon votre interface)          
            if h_sup < 0:
                st.warning(f"⚠️ Attention : Sous-charge détectée de {abs(delta_eq)} eq/h par rapport au seuil.")

            c1, c2, c3 = st.columns(3)
            
            with c1:
                st.markdown(f"<div class='metric-card'>Charge Effective<br><h2>{round(charge_effective, 2)} h</h2></div>", unsafe_allow_html=True)
            
            with c2:
                st.markdown(f"<div class='metric-card'>Seuil Réglementaire<br><h2>{seuil_obligatoire} eq/h</h2></div>", unsafe_allow_html=True)
            
            with c3:
                # Utilisation de h_sup pour la logique de couleur (résout la NameError)
                color_res = "#2ecc71" if h_sup >= 0 else "#e74c3c"
                label_res = "Heures Sup. Réelles" if h_sup >= 0 else "Déficit Horaire"
                
                st.markdown(f"""
                    <div class='metric-card' style='border-bottom: 5px solid {color_res};'>
                        {label_res}<br>
                        <h2 style='color: {color_res};'>{h_sup_formattee}</h2>
                    </div>
                """, unsafe_allow_html=True)
            # --- NOTES DE SYNTHÈSE ---
            if h_sup > 0:
                st.caption(f"✅ L'enseignant a complété sa charge et totalise {round(h_sup, 2)}h en supplément.")
            elif h_sup < 0:
                st.caption(f"⚠️ Attention : Sous-charge détectée de {round(abs(h_sup), 2)}h par rapport au seuil.")
            else:
                st.caption("⚖️ Service réglementaire exactement rempli (Pile 6.0 eq/h).")

            # --- 3. SECTION ADMINISTRATIVE : EXPORT EXCEL GLOBAL ---
            if is_admin:
                st.markdown("---")
                import io
                
                # Bouton de génération globale
                if st.button("📑 Préparer le Bilan Global (Tous les enseignants)", use_container_width=True):
                    liste_profs = sorted(df["Enseignants"].unique())
                    recap_data = []

                    for p in liste_profs:
                        # Calculs miroirs de la logique individuelle pour chaque enseignant
                        df_p = df[df["Enseignants"].str.contains(p, case=False, na=False)].copy()
                        
                        # Détermination du Type (Cours, TD, TP)
                        df_p['Type'] = df_p['Code'].apply(lambda x: "COURS" if "COURS" in str(x).upper() else ("TD" if "TD" in str(x).upper() else "TP"))
                        
                        # Suppression des doublons basés sur l'horaire normalisé
                        df_up = df_p.drop_duplicates(subset=['j_norm', 'h_norm'])
                        
                        # --- ÉTAPE CRUCIALE : DÉFINITION DES VARIABLES (CORRECTION NameError) ---
                        # On cherche le nom complet (NOM Prénom) dans le dictionnaire créé au chargement
                        nom_complet = repertoire_noms_complets.get(p.strip().upper(), p)
            
                        # Récupération du Grade et de la Qualité
                        grade_enseignant = repertoire_grades.get(p.strip().upper(), "N/A")
                        qualite_enseignant = repertoire_qualites.get(p.strip().upper(), "Non spécifié")
                        
                        # Comptage des séances
                        n_co = len(df_up[df_up['Type'] == 'COURS'])
                        n_td = len(df_up[df_up['Type'] == 'TD'])
                        n_tp = len(df_up[df_up['Type'] == 'TP'])
                        
                        # --- VOTRE LOGIQUE DE CALCUL (NON MODIFIÉE) ---
                        s_oblig = 6.0 
                        c_eq = (n_co * 1.5) + (n_td + n_tp)
                        b_h = (c_eq - s_oblig) * 1.5
                        c_eff = (n_co + n_td + n_tp) * 1.5

                        # Ajout à la liste avec la nouvelle colonne Qualité
                        recap_data.append({
                            "Enseignant": nom_complet,
                            "Grade": grade_enseignant,
                            "Qualité": qualite_enseignant,
                            "Cours": n_co,
                            "TD": n_td,
                            "TP": n_tp,
                            "Charge Effective (h)": c_eff,
                            "Total (Eq)": c_eq,
                            "Heures Sup. Réelles/Déficit Horaire": round(b_h, 2)
                        })

                    df_global = pd.DataFrame(recap_data)
                    
                    # Génération Excel avec formatage strict
                    buf = io.BytesIO()
                    with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
                        df_global.to_excel(writer, index=False, sheet_name='Bilan_Global_Charges')
                        
                        workbook  = writer.book
                        worksheet = writer.sheets['Bilan_Global_Charges']

                        # Définition du format pour les valeurs négatives
                        format_rouge = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
                        
                        # 1. Ajustement automatique de la largeur des colonnes au texte
                        for i, col in enumerate(df_global.columns):
                            # Calcul de la longueur maximale entre l'entête et le contenu
                            max_len = max(df_global[col].astype(str).map(len).max(), len(col)) + 2
                            worksheet.set_column(i, i, max_len)
                        
                        # 2. Coloration des charges inférieures à 0 (Colonne "Heures Sup. Réelles/Déficit Horaire" - index 7)
                        last_row = len(df_global)
                        worksheet.conditional_format(1, 7, last_row, 7, {
                            'type':     'cell',
                            'criteria': '<',
                            'value':    0,
                            'format':   format_rouge
                        })

                        # Figer l'entête
                        worksheet.freeze_panes(1, 0)

                    st.download_button(
                        label="📥 Télécharger le fichier Excel Global",
                        data=buf.getvalue(),
                        file_name="Bilan_Global_Charges_2027.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                    
            st.divider()
            st.markdown("### 📅 Emploi du Temps Individuel")
            
            # Récupération des informations d'identification pour les entêtes
            nom_complet_entete = repertoire_noms_complets.get(cible.strip().upper(), cible)
            grade_entete = repertoire_grades.get(cible.strip().upper(), "Grade non spécifié")
            statut_entete = repertoire_qualites.get(cible.strip().upper(), "Statut non spécifié")
            
            def format_case(rows):
                items = []
                for _, r in rows.iterrows():
                    # Disposition : Enseignements, Code, Lieu, Promotion
                    nat = '📘' if 'COURS' in str(r['Code']).upper() else '📗' if 'TD' in str(r['Code']).upper() else '🔴'
                    txt = f"<div style='margin-bottom:8px;'>{nat} <b>{r['Enseignements']}</b><br><small>({r['Code']})</small><br><i>{r['Lieu']}</i><br><b>{r['Promotion']}</b></div>"
                    items.append(txt)
                return "<div class='separator'></div>".join(items)

            if not df_f.empty:
                # --- AFFICHAGE À L'ÉCRAN ---
                grid = df_f.groupby(['h_norm', 'j_norm']).apply(format_case, include_groups=False).unstack('j_norm')
                grid = grid.reindex(index=[normalize(h) for h in horaires_list], columns=[normalize(j) for j in jours_list]).fillna("")
                grid.index = [map_h.get(i, i) for i in grid.index]
                grid.columns = [map_j.get(c, c) for c in grid.columns]
                st.write(grid.to_html(escape=False), unsafe_allow_html=True)

                # --- SECTION : BOUTONS DE TÉLÉCHARGEMENT ---
                st.markdown("---")
                col_dl1, col_dl2, col_dl3 = st.columns(3)
                
                # --- 1. EXPORT EXCEL (AVEC INFOS ENSEIGNANT) ---
                import io
                buf_ex = io.BytesIO()
    
                # Création d'un mini-tableau d'entête pour l'Excel
                df_infos = pd.DataFrame([
                    ["Enseignant :", nom_complet_entete],
                    ["Grade :", grade_entete],
                    ["Qualité :", statut_entete],
                    ["", ""] # Ligne vide de séparation
                ])
                
                df_to_export = df_f.drop(columns=['h_norm', 'j_norm'], errors='ignore')
                
                with pd.ExcelWriter(buf_ex, engine='xlsxwriter') as writer:
                    # On écrit d'abord les infos en haut à gauche
                    df_infos.to_excel(writer, index=False, header=False, sheet_name='Mon_EDT')
                    # On écrit la liste des cours juste en dessous (ligne 5)
                    df_to_export.to_excel(writer, index=False, startrow=5, sheet_name='Mon_EDT')
                        
                col_dl1.download_button(
                    label="📥 Liste (Excel)",
                    data=buf_ex.getvalue(),
                    file_name="EDT_Individuel_2027.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="btn_indiv_xl_final_v12"
                )

                # --- 2. EXPORT HTML ---
                if not df_f.empty:
                    style_css_safe = style_css if 'style_css' in locals() else ""
                    content_html_safe = content_html if 'content_html' in locals() else df_f.to_html()

                    col_dl2.download_button(
                        label="🌐 Tableau (HTML)",
                        data=f"<!DOCTYPE html><html><head><meta charset='UTF-8'>{style_css_safe}</head><body>{content_html_safe}</body></html>",
                        file_name="EDT_Individuel_2027.html",
                        mime="text/html",
                        use_container_width=True,
                        key="btn_indiv_html_final_v12"
                    )

                    # --- 3. EXPORT PDF (CENTRAGE & MARGES DE SÉCURITÉ) ---
                    try:
                        from fpdf import FPDF
                        import re

                        class INDIV_PDF(FPDF):
                            def header(self):
                                self.set_font('Arial', 'B', 10)
                                t = "Plateforme de gestion des EDTs-S2-2027-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA"
                                self.cell(0, 8, t.encode('latin-1', 'replace').decode('latin-1'), 0, 1, 'C')
                                self.ln(2)
                            
                            def get_nb_lines(self, w, txt, margin_h=4):
                                """Calcule le nombre de lignes en tenant compte d'une marge horizontale interne"""
                                if not txt: return 1
                                # Largeur utile réelle pour le texte (Largeur cellule - marges gauche/droite)
                                effective_w = w - margin_h
                                lines = 0
                                for paragraph in txt.split('\n'):
                                    width = self.get_string_width(paragraph)
                                    # On ajoute 1 pour forcer le passage à la ligne si ça dépasse
                                    lines += max(1, int(width / (effective_w - 1)) + 1)
                                return lines
                        def clean_indiv(text_val):
                            if not text_val: return ""
                            # 1. On convertit en chaîne de caractères
                            t = str(text_val)
                            # 2. Nettoyage des balises HTML
                            t = t.replace('<b>','').replace('</b>','')
                            # 3. CORRECTION DE L'APOSTROPHE (Remplace l'apostrophe courbe par la droite)
                            t = t.replace("’", "'").replace("‘", "'")
                            # 4. Encodage final pour FPDF
                            return t.encode('latin-1', 'replace').decode('latin-1')
                        pdf = INDIV_PDF(orientation="L", unit="mm", format="A4")
                        pdf.set_margins(7, 10, 7)
                        pdf.add_page()

                        # Titres
                        pdf.set_font("Arial", "B", 11)
                        pdf.cell(0, 8, "EMPLOI DU TEMPS INDIVIDUEL".encode('latin-1', 'replace').decode('latin-1'), 0, 1, "C")
                        
                        # Récapitulatif
                        nb_cours = len(df_f[df_f['Enseignements'].str.contains('Cours', case=False, na=False)])
                        nb_td = len(df_f[df_f['Enseignements'].str.contains('Td', case=False, na=False)])
                        nb_tp = len(df_f[df_f['Enseignements'].str.contains('Tp', case=False, na=False)])
                        pdf.set_font("Arial", "I", 9)
                        pdf.cell(0, 6, f"Recapitulatif : {nb_cours} Cours | {nb_td} TD | {nb_tp} TP", 0, 1, "C")
                        pdf.ln(4)

                        # --- LOGIQUE DE TRI & FUSION ---
                        ordre_horaires = ["8h-9h30", "9h30-11h", "11h-12h30", "12h30-14h", "13h-14h30", "14h-15h30", "15h30-17h"]
                        df_pdf = df_f.copy()
                        
                        def merge_info(row):
                            txt = f"{row['Enseignements']}"
                            if 'Enseignants' in row and row['Enseignants']: txt += f"\n- {row['Enseignants']}"
                            if 'Lieu' in row and row['Lieu']: txt += f"\n{row['Lieu']}"
                            if 'Lieu' in row and row['Promotion']: txt += f"\n{row['Promotion']}"         
                            return txt

                        df_pdf['Info_Cell'] = df_pdf.apply(merge_info, axis=1)
                        grid = df_pdf.pivot_table(index='Horaire', columns='Jours', values='Info_Cell', aggfunc=lambda x: "\n".join(x)).fillna("")
                        
                        index_present = [h for h in ordre_horaires if h in grid.index]
                        grid = grid.reindex(index=index_present)
                        jours_ordre = ["Dimanche", "Lundi", "Mardi", "Mercredi", "Jeudi"]
                        grid = grid.reindex(columns=[j for j in jours_ordre if j in grid.columns])

                        # --- CONFIGURATION DU TRACÉ ---
                        col_h_w = 32
                        # Marge de sécurité horizontale par cellule (mm)
                        m_h = 6 
                        col_j_w = (pdf.w - col_h_w - 20) / len(grid.columns) if len(grid.columns) > 0 else 100
                        interline = 3.8
                        padding_v = 3 # Marge de sécurité verticale

                        # En-tête
                        pdf.set_font("Arial", "B", 8)
                        pdf.set_fill_color(220, 220, 220)
                        pdf.cell(col_h_w, 10, "HORAIRE", 1, 0, "C", True)
                        for jour in grid.columns:
                            pdf.cell(col_j_w, 10, str(jour).encode('latin-1', 'replace').decode('latin-1'), 1, 0, "C", True)
                        pdf.ln()

                        # Données
                        for horaire, row in grid.iterrows():
                            texts = [clean_indiv(val) for val in row]
                            
                            # 1. Calcul de la hauteur de ligne (Balayage avec marges)
                            max_h = 12 
                            pdf.set_font("Arial", "", 6.5)
                            for t in texts:
                                n_lines = pdf.get_nb_lines(col_j_w, t, margin_h=m_h)
                                h_total = (n_lines * interline) + (padding_v * 2)
                                if h_total > max_h: max_h = h_total

                            # 2. Rendu Horaire
                            pdf.set_font("Arial", "B", 7.5)
                            pdf.set_fill_color(248, 248, 248)
                            pdf.cell(col_h_w, max_h, str(horaire), 1, 0, "C", True)

                            # 3. Rendu Contenu (Centrage vertical et horizontal sans toucher les traits)
                            pdf.set_font("Arial", "", 6.8)
                            for idx, content in enumerate(texts):
                                # Couleur fond
                                raw_c = str(row.iloc[idx]).upper()
                                if "COURS" in raw_c: pdf.set_fill_color(225, 238, 255)
                                elif "TD" in raw_c: pdf.set_fill_color(232, 252, 235)
                                elif "TP" in raw_c: pdf.set_fill_color(255, 235, 235)
                                else: pdf.set_fill_color(255, 255, 255)

                                x, y = pdf.get_x(), pdf.get_y()
                                # Dessine le rectangle (la cellule)
                                pdf.rect(x, y, col_j_w, max_h, 'FD')
                                
                                # Calcul du bloc de texte pour le centrage vertical
                                n_l = pdf.get_nb_lines(col_j_w, content, margin_h=m_h)
                                text_h = n_l * interline
                                
                                # Positionnement du curseur avec marge interne
                                pdf.set_xy(x + (m_h/2), y + (max_h - text_h) / 2)
                                
                                # Rendu multi-cellule centré horizontalement dans la zone utile
                                pdf.multi_cell(col_j_w - m_h, interline, content, 0, "C")
                                
                                # Retour au curseur pour la cellule suivante
                                pdf.set_xy(x + col_j_w, y)
                            
                            pdf.ln(max_h)

                        col_dl3.download_button(
                            label="📄 Emploi du temps individuel (PDF)",
                            data=bytes(pdf.output()),
                            file_name="EDT_Individuel_2027.pdf",
                            mime="application/pdf",
                            use_container_width=True,
                            key="btn_indiv_pdf_v12_final"
                        )
                    except Exception as e:
                        col_dl3.error(f"Erreur rendu PDF : {e}")
                    # --- LOGIQUE DE TRI CHRONOLOGIQUE (Ajoutée pour l'ordre) ---
                    ordre_horaires = [
                        "8h-9h30", "8h-10h", "8h-11h", "9h30-11h", "10h-11h", 
                        "11h-12h30", "11h-12h", "12h30-14h", "13h-14h", 
                        "14h-15h30", "14h-16h", "15h30-17h"
                    ]
                    # Normalisation pour éviter les erreurs d'espaces
                    df_f['Horaire'] = df_f['Horaire'].astype(str).str.replace(' ', '').str.strip()
                    df_f['Horaire'] = pd.Categorical(df_f['Horaire'], categories=ordre_horaires, ordered=True)
                    df_f = df_f.sort_values(['Horaire'])

                    # Préparation de la grille : on ne garde que les horaires de l'enseignant
                    charge_group = df_f.groupby(['Horaire', 'Jours'], observed=True).apply(format_case, include_groups=False)
                    grid_charge = charge_group.unstack('Jours').fillna("")
                    
                    # Réordonner les jours présents
                    jours_present = [j for j in jours_list if j in grid_charge.columns]
                    grid_charge = grid_charge.reindex(columns=jours_present)
                    
                    # Conversion en HTML avec styles optimisés
                    html_table_content = grid_charge.to_html(escape=False, classes='table-charge')
                    
                    # Utilisation du nom de l'enseignant (cible) défini plus haut dans votre code
                    nom_affiche = cible if 'cible' in locals() else "Enseignant"

                    html_final_doc = f"""
                    <!DOCTYPE html>
                    <html lang="fr">
                    <head>
                        <meta charset="UTF-8">
                        <style>
                            @media screen {{
                                body {{ 
                                    font-family: 'Segoe UI', Arial, sans-serif; 
                                    padding: 30px; 
                                    background-color: #f1f5f9; 
                                    color: #334e68;
                                }}
                                .btn-print {{
                                    display: inline-block;
                                    padding: 12px 25px;
                                    background-color: #2563eb;
                                    color: white;
                                    text-decoration: none;
                                    border-radius: 8px;
                                    font-weight: bold;
                                    margin-bottom: 20px;
                                    box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1);
                                }}
                            }}

                            /* --- CONFIGURATION IMPRESSION FORCEE SUR UNE PAGE --- */
                            @media print {{
                                .btn-print {{ display: none !important; }}
                                body {{ 
                                    background-color: white !important; 
                                    padding: 0 !important; 
                                    margin: 0 !important;
                                }}
                                @page {{
                                    size: A4 landscape;
                                    margin: 0.5cm;
                                }}
                                .header-box {{ margin-bottom: 10px !important; padding: 10px !important; }}
                                .prof-name {{ font-size: 22px !important; }}
                                .recap-container {{ margin-bottom: 10px !important; }}
                                .table-charge {{ font-size: 10px !important; }}
                            }}

                            .header-box {{
                                text-align: center;
                                border: 1px solid #e2e8f0;
                                padding: 15px;
                                border-radius: 8px;
                                background-color: #f8fafc;
                                margin-bottom: 20px;
                            }}
                            .main-title {{ 
                                color: #64748b; font-size: 12px; font-weight: bold; 
                                text-transform: uppercase; letter-spacing: 1px;
                            }}
                            .prof-name {{
                                font-size: 28px; color: #1e293b; font-weight: 800; margin: 5px 0;
                            }}
                            .recap-container {{
                                display: flex; justify-content: center; gap: 15px; margin-bottom: 25px;
                            }}
                            .recap-box {{
                                border: 1px solid #e2e8f0; padding: 10px 20px; border-radius: 8px;
                                background: #ffffff; text-align: center; min-width: 100px;
                            }}
                            .recap-box b {{ display: block; font-size: 18px; color: #1e293b; }}
                            .recap-box span {{ font-size: 11px; color: #64748b; text-transform: uppercase; }}

                            .table-charge {{ width: 100%; border-collapse: collapse; background: white; }}
                            .table-charge th, .table-charge td {{
                                border: 1px solid #e2e8f0; padding: 10px; text-align: center;
                            }}
                            .table-charge th {{ background-color: #f1f5f9; color: #1e293b; font-size: 12px; }}

                            div:has(b:contains("📘")) {{ background-color: #f0f9ff; border-left: 5px solid #3b82f6; padding: 8px; text-align: left; }}
                            div:has(b:contains("📗")) {{ background-color: #f0fdf4; border-left: 5px solid #22c55e; padding: 8px; text-align: left; }}
                            div:has(b:contains("🔴")) {{ background-color: #fef2f2; border-left: 5px solid #ef4444; padding: 8px; text-align: left; }}
                            
                            i {{ font-weight: bold; background: #f1f5f9; padding: 2px 4px; border-radius: 3px; }}
                            .footer {{ margin-top: 20px; text-align: center; font-size: 10px; color: #94a3b8; }}
                        </style>
                    </head>
                    <body>
                        <a href="#" class="btn-print" onclick="window.print();return false;">🖨️ Imprimer la page (A4 Paysage)</a>

                        <div class="header-box">
                            <div class="main-title">Plateforme de gestion des EDTs-S2-2027-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA</div>
                            <div class="prof-name">Enseignant (e) : {nom_affiche}</div>
                        </div>

                        <div class="recap-container">
                            <div class="recap-box"><span>📘 Cours</span><b>{nb_cours}</b></div>
                            <div class="recap-box"><span>📗 TD</span><b>{nb_td}</b></div>
                            <div class="recap-box"><span>🔴 TP</span><b>{nb_tp}</b></div>
                        </div>
                        
                        {html_table_content}
                        
                        <div class="footer">
                            Document généré numériquement le {pd.Timestamp.now().strftime('%d/%m/%Y à %H:%M')}
                        </div>
                    </body>
                    </html>
                    """

                    col_dl2.download_button(
                        label="🌐 Télécharger la Charge HTML",
                        data=html_final_doc,
                        file_name=f"EDT_{nom_affiche.replace(' ', '_')}.html",
                        mime="text/html",
                        use_container_width=True
                    )
        elif is_admin and mode_view == "Promotion":
            # 1. Sélection de la promotion via le menu déroulant
            p_sel = st.selectbox("Choisir Promotion :", sorted(df["Promotion"].unique()))
            df_p = df[df["Promotion"] == p_sel].copy()
            
            # --- 2. CALCUL DES STATISTIQUES POUR L'EN-TÊTE ---
            # Identification des types pour le décompte (Cours, TD, TP)
            df_p['Type_Tmp'] = df_p['Code'].apply(lambda x: "COURS" if "COURS" in str(x).upper() else ("TD" if "TD" in str(x).upper() else "TP"))
            
            # Décompte basé sur les séances uniques (évite les doublons par groupe sur un même créneau)
            df_stats = df_p.drop_duplicates(subset=['j_norm', 'h_norm'])
            n_p_co = len(df_stats[df_stats['Type_Tmp'] == 'COURS'])
            n_p_td = len(df_stats[df_stats['Type_Tmp'] == 'TD'])
            n_p_tp = len(df_stats[df_stats['Type_Tmp'] == 'TP'])

            # --- 3. FONCTION DE FORMATAGE DES CELLULES (HTML) ---
            def fmt_p(rows):
                items = []
                for _, r in rows.iterrows():
                    # Choix de la couleur selon la nature de l'enseignement
                    code_up = str(r['Code']).upper()
                    if 'COURS' in code_up:
                        nat, color = '📘', '#1e40af' # Bleu
                    elif 'TD' in code_up:
                        nat, color = '📗', '#166534' # Vert
                    else:
                        nat, color = '🔴', '#991b1b' # Rouge
                        
                    # Structure HTML de la séance (identique au fichier exporté)
                    txt = f"""
                    <div style='margin-bottom:8px; padding:5px; border-left:3px solid {color}; background-color:#f8fafc; border-radius:4px;'>
                        <b style='color:{color};'>{nat} {r['Enseignements']}</b><br>
                        <span style='font-size:11px; font-weight:bold;'>👤 {r['Enseignants']}</span><br>
                        <span style='font-size:11px;'>📍 {r['Lieu']}</span>
                    </div>
                    """
                    items.append(txt)
                return "".join(items)

            # --- 4. CONSTRUCTION ET FILTRAGE DE LA GRILLE ---
            # Groupement des données par horaire et par jour
            grid_p = df_p.groupby(['h_norm', 'j_norm']).apply(fmt_p, include_groups=False).unstack('j_norm')
            
            # Réindexation sur tous les créneaux et jours définis globalement
            idx_h = [normalize(h) for h in horaires_list]
            cols_j = [normalize(j) for j in jours_list]
            grid_p = grid_p.reindex(index=idx_h, columns=cols_j).fillna("")

            # FILTRAGE : Suppression des lignes horaires totalement vides pour cette promotion
            grid_p = grid_p[grid_p.any(axis=1)]

            # Application des noms réels (labels) sur les index et colonnes
            grid_p.index = [map_h.get(i, i) for i in grid_p.index]
            grid_p.columns = [map_j.get(c, c) for c in grid_p.columns]

            # --- 5. DÉFINITION DU STYLE CSS (PARTAGÉ STREAMLIT / EXPORT) ---
            style_css = """
            <style>
                .p-container { font-family: 'Segoe UI', Tahoma, sans-serif; padding: 20px; color: #1e293b; background-color: #ffffff; }
                .p-header { text-align: center; margin-bottom: 20px; border-bottom: 2px solid #3b82f6; padding-bottom: 10px; }
                .p-promo-name { font-size: 24px; font-weight: bold; color: #1e40af; text-transform: uppercase; }
                .p-stats { margin-top: 10px; font-size: 16px; font-weight: 600; display: flex; justify-content: center; gap: 20px; }
                .p-table { width: 100%; border-collapse: collapse; margin-top: 20px; table-layout: fixed; border: 1px solid #cbd5e1; }
                .p-table th { background-color: #f1f5f9; border: 1px solid #cbd5e1; padding: 12px; font-size: 14px; text-align: center; }
                .p-table td { border: 1px solid #cbd5e1; padding: 10px; vertical-align: top; min-height: 80px; word-wrap: break-word; }
                .p-time-col { width: 120px; background-color: #f8fafc !important; font-weight: bold; text-align: center; vertical-align: middle !important; }
            </style>
            """

            # --- 6. GÉNÉRATION DU CONTENU HTML DU TABLEAU ---
            content_html = f"""
            <div class="p-container">
                <div class="p-header">
                    <div class="p-promo-name">PROMOTION : {p_sel}</div>
                    <div class="p-stats">
                        <span>📘 COURS : {n_p_co}</span>
                        <span>📗 TD : {n_p_td}</span>
                        <span>🔴 TP : {n_p_tp}</span>
                    </div>
                </div>
                <table class="p-table">
                    <thead>
                        <tr>
                            <th style="width:120px;">HORAIRE</th>
                            {" ".join([f"<th>{day}</th>" for day in grid_p.columns])}
                        </tr>
                    </thead>
                    <tbody>
            """

            # Construction des lignes du tableau
            for time_label, row in grid_p.iterrows():
                content_html += f"<tr><td class='p-time-col'>{time_label}</td>"
                for day_label in grid_p.columns:
                    content_html += f"<td>{row[day_label]}</td>"
                content_html += "</tr>"

            content_html += "</tbody></table></div>"

            # --- AFFICHAGE SYNTHÉTIQUE DES ENSEIGNEMENTS PAR ENSEIGNANT ---
            
            st.subheader(f"📚 Récapitulatif des enseignements : {p_sel}")
            
            # --- NOUVEAU : BILAN GLOBAL DE LA PROMOTION (SANS DOUBLONS) ---
            # On retire les doublons basés sur le nom de l'enseignement et le code (type)
            # pour ne compter chaque matière qu'une seule fois pour toute la promo
            df_unique_matieres = df_p.drop_duplicates(subset=['Enseignements', 'Code'])
            
            total_p_cours = len(df_unique_matieres[df_unique_matieres['Code'].str.contains('COURS', case=False, na=False)])
            total_p_td = len(df_unique_matieres[df_unique_matieres['Code'].str.contains('TD', case=False, na=False)])
            total_p_tp = len(df_unique_matieres[~df_unique_matieres['Code'].str.contains('COURS|TD', case=False, na=False)])
                    # --- 8. BOUTONS DE TÉLÉCHARGEMENT ---
            st.markdown("---")
            cp1, cp2, cp3 = st.columns(3)

            # --- 8.1. Export Excel Promotion ---
            import io
            buf_p = io.BytesIO()
            df_export = df_p.drop(columns=['h_norm', 'j_norm', 'Type_Tmp'], errors='ignore')
            df_export.to_excel(buf_p, index=False)
            
            cp1.download_button(
                label=f"📥 Liste {p_sel} (Excel)",
                data=buf_p.getvalue(),
                file_name=f"EDT_{p_sel}_2027.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key=f"btn_xl_v8_{p_sel}" 
            )

            # --- 8.2. Export HTML Promotion ---
            content_html_fixed = content_html.replace('\\n', '<br>').replace('\n', '<br>')
            full_html_doc = f"<!DOCTYPE html><html lang='fr'><head><meta charset='UTF-8'><title>EDT {p_sel}</title>{style_css}</head><body>{content_html_fixed}</body></html>"
            
            cp2.download_button(
                label=f"🌐 Tableau {p_sel} (HTML)",
                data=full_html_doc,
                file_name=f"EDT_{p_sel}_2027.html",
                mime="text/html",
                use_container_width=True,
                key=f"btn_html_v8_{p_sel}"
            )

            # --- 8.3. Export PDF Promotion (Correction CoreFont) ---
            try:
                from fpdf import FPDF
                import re

                class EDT_PDF(FPDF):
                    def header(self):
                        self.set_font('Arial', 'B', 10)
                        title = "Plateforme de gestion des EDTs-S2-2027-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA"
                        self.cell(0, 8, title.encode('latin-1', 'replace').decode('latin-1'), 0, 1, 'C')
                        self.ln(2)

                    def get_nb_lines(self, w, txt):
                        """Calcule le nombre de lignes réelles après retour à la ligne automatique"""
                        if not txt: return 1
                        lines = 0
                        for paragraph in txt.split('\n'):
                            # get_string_width donne la largeur totale du texte sans retours
                            width = self.get_string_width(paragraph)
                            # On ajoute le nombre de lignes créées par le wrap automatique
                            lines += max(1, int((width + (2 * self.c_margin)) / (w - (2 * self.c_margin))) + 1)
                        return lines

                def clean_text_for_pdf(html_str):
                    if not html_str: return ""
                    t = str(html_str).replace('</div>', '\n').replace('<br>', '\n').replace('<br/>', '\n')
                    t = t.replace('<b>','').replace('</b>','')
                    t = t.replace('👤', '- ').replace('📘', '').replace('📗', '').replace('🔴', '').replace('📍', '').replace('Lieu:', '').replace('Lieu', '')
                    t = re.sub(r'<[^>]+>', '', t)
                    lines = [l.strip() for l in t.split('\n') if l.strip()]
                    return "\n".join([line.encode('latin-1', 'replace').decode('latin-1') for line in lines])

                pdf = EDT_PDF(orientation="L", unit="mm", format="A4")
                pdf.set_margins(left=5, top=10, right=5)
                pdf.add_page()

                pdf.set_font("Arial", "B", 11)
                pdf.cell(0, 8, f"PROMOTION : {p_sel}".encode('latin-1', 'replace').decode('latin-1'), 0, 1, "C")
                pdf.ln(2)

                grid_pdf = df_p.groupby(['h_norm', 'j_norm']).apply(fmt_p, include_groups=False).unstack('j_norm')
                grid_pdf = grid_pdf.reindex(index=idx_h, columns=cols_j).fillna("")
                grid_pdf = grid_pdf[grid_pdf.any(axis=1)] 
                grid_pdf.index = [map_h.get(i, i) for i in grid_pdf.index]
                grid_pdf.columns = [map_j.get(c, c) for c in grid_pdf.columns]

                # --- Dimensions ---
                col_h_w = 25
                col_j_w = (pdf.w - 35) / len(grid_pdf.columns)
                interline = 3.5
                padding_v = 2

                # En-tête
                pdf.set_font("Arial", "B", 8)
                pdf.set_fill_color(220, 220, 220)
                pdf.cell(col_h_w, 9, "HORAIRE", 1, 0, "C", True)
                for jour in grid_pdf.columns:
                    pdf.cell(col_j_w, 9, jour.encode('latin-1', 'replace').decode('latin-1'), 1, 0, "C", True)
                pdf.ln()

                # Corps
                for heure, row in grid_pdf.iterrows():
                    row_texts = []
                    max_h = 10
                    
                    # Balayage pour hauteur
                    pdf.set_font("Arial", "", 6)
                    for jour in grid_pdf.columns:
                        txt = clean_text_for_pdf(row[jour])
                        row_texts.append(txt)
                        nb_l = pdf.get_nb_lines(col_j_w, txt)
                        h_calc = (nb_l * interline) + (padding_v * 2)
                        if h_calc > max_h: max_h = h_calc

                    # Rendu Heure
                    pdf.set_font("Arial", "B", 7)
                    pdf.set_fill_color(248, 248, 248)
                    pdf.cell(col_h_w, max_h, str(heure), 1, 0, "C", True)

                    # Rendu Cours
                    pdf.set_font("Arial", "", 6)
                    for idx, jour in enumerate(grid_pdf.columns):
                        content = row_texts[idx]
                        raw_c = str(row[jour]).upper()
                        
                        if "COURS" in raw_c: pdf.set_fill_color(225, 238, 255)
                        elif "TD" in raw_c: pdf.set_fill_color(232, 252, 235)
                        elif "TP" in raw_c: pdf.set_fill_color(255, 235, 235)
                        else: pdf.set_fill_color(255, 255, 255)

                        x, y = pdf.get_x(), pdf.get_y()
                        pdf.rect(x, y, col_j_w, max_h, 'FD')
                        
                        nb_l = pdf.get_nb_lines(col_j_w, content)
                        pdf.set_xy(x, y + (max_h - (nb_l * interline)) / 2)
                        pdf.multi_cell(col_j_w, interline, content, 0, "C")
                        pdf.set_xy(x + col_j_w, y)
                    pdf.ln(max_h)

                pdf_bytes = pdf.output()
                cp3.download_button(
                    label=f"📄 Emploi du temps {p_sel} (PDF)",
                    data=bytes(pdf_bytes),
                    file_name=f"EDT_{p_sel}_2027.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                    key=f"btn_pdf_v8_{p_sel}" 
                )
            except Exception as e:
                cp3.error(f"Erreur technique PDF : {e}")
            # Affichage du bandeau récapitulatif global
            st.markdown(f"""
            <div style='display: flex; justify-content: space-around; background-color: #f8fafc; padding: 15px; border-radius: 10px; border: 1px solid #e2e8f0; margin-bottom: 20px; text-align: center;'>
                <div>
                    <div style='font-size: 12px; color: #64748b; text-transform: uppercase; font-weight: bold;'>Total Cours</div>
                    <div style='font-size: 22px; font-weight: bold; color: #1e40af;'>📘 {total_p_cours}</div>
                </div>
                <div style='border-left: 1px solid #e2e8f0; height: 40px;'></div>
                <div>
                    <div style='font-size: 12px; color: #64748b; text-transform: uppercase; font-weight: bold;'>Total TD</div>
                    <div style='font-size: 22px; font-weight: bold; color: #166534;'>📗 {total_p_td}</div>
                </div>
                <div style='border-left: 1px solid #e2e8f0; height: 40px;'></div>
                <div>
                    <div style='font-size: 12px; color: #64748b; text-transform: uppercase; font-weight: bold;'>Total TP</div>
                    <div style='font-size: 22px; font-weight: bold; color: #991b1b;'>🔴 {total_p_tp}</div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            # --- 1. Extraction et tri des enseignants uniques de la promotion ---
            enseignants_promo = sorted(df_p["Enseignants"].unique())
            
            for ens in enseignants_promo:
                # Filtrer les matières pour cet enseignant précis
                df_ens = df_p[df_p["Enseignants"] == ens].copy()
                
                # 2. Organisation par type (Ordre : COURS > TD > TP) sans doublons pour l'affichage des badges
                matieres_brutes = df_ens.drop_duplicates(subset=['Enseignements', 'Code'])
                
                cours_list = matieres_brutes[matieres_brutes['Code'].str.contains('COURS', case=False, na=False)]['Enseignements'].unique()
                td_list = matieres_brutes[matieres_brutes['Code'].str.contains('TD', case=False, na=False)]['Enseignements'].unique()
                tp_list = matieres_brutes[~matieres_brutes['Code'].str.contains('COURS|TD', case=False, na=False)]['Enseignements'].unique()
                
                # 3. Calcul du nombre de séances par enseignant (avec groupes)
                n_cours = len(df_ens[df_ens['Code'].str.contains('COURS', case=False, na=False)])
                n_td = len(df_ens[df_ens['Code'].str.contains('TD', case=False, na=False)])
                n_tp = len(df_ens[~df_ens['Code'].str.contains('COURS|TD', case=False, na=False)])
                
                # 4. Construction de l'affichage HTML pour l'enseignant (Couleurs claires)
                items_html = ""
                for c in cours_list:
                    items_html += f"<span style='background-color:#dbeafe; color:#1e40af; padding:4px 10px; border-radius:15px; margin:3px; display:inline-block; font-size:12px; border:1px solid #bfdbfe;'>📘 {c}</span>"
                for t in td_list:
                    items_html += f"<span style='background-color:#dcfce7; color:#166534; padding:4px 10px; border-radius:15px; margin:3px; display:inline-block; font-size:12px; border:1px solid #bbf7d0;'>📗 {t}</span>"
                for p in tp_list:
                    items_html += f"<span style='background-color:#fee2e2; color:#991b1b; padding:4px 10px; border-radius:15px; margin:3px; display:inline-block; font-size:12px; border:1px solid #fecaca;'>🔴 {p}</span>"
                
                # 5. Affichage final dans un conteneur stylisé
                if items_html: 
                    st.markdown(f"""
                    <div style='padding:15px; border:1px solid #e2e8f0; border-radius:10px; margin-bottom:15px; background-color:white; box-shadow: 0 2px 4px rgba(0,0,0,0.05);'>
                        <div style='display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px; border-bottom: 1px solid #f1f5f9; padding-bottom: 8px;'>
                            <div style='font-weight:bold; color:#1e293b; font-size:16px;'>👤 M. {ens}</div>
                            <div style='font-size:11px; color:#64748b; font-weight: 600;'>
                                <span style='margin-left:8px;'>📘 {n_cours} Séc.</span>
                                <span style='margin-left:8px;'>📗 {n_td} Groupes</span>
                                <span style='margin-left:8px;'>🔴 {n_tp} Groupes</span>
                            </div>
                        </div>
                        <div>{items_html}</div>
                    </div>
                    """, unsafe_allow_html=True)

            # --- 8. BOUTONS DE TÉLÉCHARGEMENT ---
            st.markdown("---")
            cp1, cp2, cp3 = st.columns(3)

            # --- 8.1. Export Excel Promotion ---
            import io
            buf_p = io.BytesIO()
            df_export = df_p.drop(columns=['h_norm', 'j_norm', 'Type_Tmp'], errors='ignore')
            df_export.to_excel(buf_p, index=False)
            
            cp1.download_button(
                label=f"📥 Liste {p_sel} (Excel)",
                data=buf_p.getvalue(),
                file_name=f"EDT_{p_sel}_2027.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key=f"btn_xl_{p_sel}"
            )

            # --- 8.2. Export HTML Promotion ---
            content_html_fixed = content_html.replace('\\n', '<br>').replace('\n', '<br>')
            full_html_doc = f"<!DOCTYPE html><html lang='fr'><head><meta charset='UTF-8'><title>EDT {p_sel}</title>{style_css}</head><body>{content_html_fixed}</body></html>"
            
            cp2.download_button(
                label=f"🌐 Tableau {p_sel} (HTML)",
                data=full_html_doc,
                file_name=f"EDT_{p_sel}_2027.html",
                mime="text/html",
                use_container_width=True,
                key=f"btn_html_{p_sel}"
            )

            # --- 8.3. Export PDF Promotion (Haute Lisibilité) ---
            try:
                from fpdf import FPDF
                import re

                class EDT_PDF(FPDF):
                    def header(self):
                        self.set_font('Arial', 'B', 10)
                        # Rappel du titre officiel
                        self.cell(0, 8, "Plateforme de gestion des EDTs-S2-2027-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA", 0, 1, 'C')
                        self.ln(2)
                    def footer(self): pass

                def clean_text_for_pdf(html_str):
                    if not html_str: return ""
                    t = str(html_str).replace('</div>', '\n').replace('<br>', '\n').replace('<br/>', '\n')
                    t = t.replace('<b>','').replace('</b>','')
                    t = re.sub(r'<[^>]+>', '', t)
                    t = re.sub(r'[^\x00-\x7F]+', '', t) 
                    lines = [l.strip() for l in t.split('\n') if l.strip()]
                    return "\n".join(lines)

                # Initialisation PDF Paysage
                pdf = EDT_PDF(orientation="L", unit="mm", format="A4")
                pdf.set_margins(left=7, top=7, right=7)
                pdf.set_auto_page_break(auto=True, margin=10)
                pdf.add_page()

                # Titre de la promotion
                pdf.set_font("Arial", "B", 12)
                pdf.cell(0, 10, f"PROMOTION : {p_sel}", 0, 1, "C")
                pdf.ln(2)

                # Reconstruction de la grille pour le PDF
                grid_pdf = df_p.groupby(['h_norm', 'j_norm']).apply(fmt_p, include_groups=False).unstack('j_norm')
                grid_pdf = grid_pdf.reindex(index=idx_h, columns=cols_j).fillna("")
                grid_pdf = grid_pdf[grid_pdf.any(axis=1)] 
                grid_pdf.index = [map_h.get(i, i) for i in grid_pdf.index]
                grid_pdf.columns = [map_j.get(c, c) for c in grid_pdf.columns]

                # Paramètres de style
                col_h_w = 26
                col_j_w = (pdf.w - col_h_w - 20) / len(grid_pdf.columns)
                interline = 3.5
                padding_v = 3

                # En-tête du tableau
                pdf.set_font("Arial", "B", 9)
                pdf.set_fill_color(230, 230, 230)
                pdf.cell(col_h_w, 10, "HORAIRE", 1, 0, "C", True)
                for jour in grid_pdf.columns:
                    pdf.cell(col_j_w, 10, jour, 1, 0, "C", True)
                pdf.ln()

                # Remplissage du tableau
                for heure, row in grid_pdf.iterrows():
                    # Calcul hauteur de ligne
                    max_h_needed = 14
                    row_texts = []
                    for jour in grid_pdf.columns:
                        txt_propre = clean_text_for_pdf(row[jour])
                        row_texts.append(txt_propre)
                        nb_l = txt_propre.count('\n') + 1
                        h_calc = (nb_l * interline) + (padding_v * 2)
                        if h_calc > max_h_needed: max_h_needed = h_calc

                    # Cellule Heure
                    pdf.set_font("Arial", "B", 7.5)
                    pdf.set_fill_color(245, 245, 245)
                    pdf.cell(col_h_w, max_h_needed, heure, 1, 0, "C", True)

                    # Cellules Cours
                    pdf.set_font("Arial", "", 6.5)
                    for idx, jour in enumerate(grid_pdf.columns):
                        content = row_texts[idx]
                        raw_val = str(row[jour]).upper()
                        
                        if "COURS" in raw_val: pdf.set_fill_color(225, 238, 255)
                        elif "TD" in raw_val: pdf.set_fill_color(232, 252, 235)
                        elif "TP" in raw_val: pdf.set_fill_color(255, 235, 235)
                        else: pdf.set_fill_color(255, 255, 255)

                        cur_x, cur_y = pdf.get_x(), pdf.get_y()
                        pdf.rect(cur_x, cur_y, col_j_w, max_h_needed, 'FD')
                        
                        text_block_h = (content.count('\n') + 1) * interline
                        margin_top = (max_h_needed - text_block_h) / 2
                        
                        pdf.set_xy(cur_x, cur_y + margin_top)
                        pdf.multi_cell(col_j_w, interline, content, 0, "C")
                        pdf.set_xy(cur_x + col_j_w, cur_y)
                    pdf.ln(max_h_needed)

                pdf_bytes = pdf.output()

                cp3.download_button(
                    label=f"📄 Emploi du temps {p_sel} (PDF)",
                    data=bytes(pdf_bytes),
                    file_name=f"EDT_{p_sel}_2027.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                    key=f"btn_pdf_{p_sel}"
                )
            except Exception as e:
                cp3.error(f"Erreur lors de la génération PDF : {e}")
            # --- 8.2. EXPORT PACK HTML (CORRIGÉ & NETTOYÉ) ---
            st.markdown("---") # <--- Vérifiez que cette ligne est alignée avec le "if" précédent
            if st.button("📦 Générer le Pack de fichiers HTML (Version Finale)", use_container_width=True):
                try:
                    import io
                    import zipfile
                    import re

                    zip_buffer = io.BytesIO()
                    promos_disponibles = sorted(df["Promotion"].unique())

                    # Style CSS amélioré
                    style_edt_html = """
                    <style>
                        body { font-family: 'Segoe UI', Arial, sans-serif; background-color: #f8fafc; padding: 15px; }
                        .container { max-width: 1000px; margin: auto; background: white; padding: 15px; border-radius: 10px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); }
                        h2 { color: #1e293b; text-align: center; border-bottom: 2px solid #3b82f6; padding-bottom: 10px; }
                        table { width: 100%; border-collapse: collapse; margin-top: 15px; table-layout: fixed; }
                        th { background-color: #1e293b; color: white; padding: 10px; border: 1px solid #cbd5e1; }
                        td { border: 1px solid #cbd5e1; padding: 10px; vertical-align: top; text-align: center; font-size: 10px; line-height: 1.5; }
                    </style>
                    """

                    def clean_html_content(text):
                        """Remplace les \n par des sauts de ligne réels et nettoie le texte"""
                        if not text or text == "None": return ""
                        # Transforme les sauts de ligne en balises HTML <br>
                        return str(text).replace('\\n', '<br>').replace('\n', '<br>')

                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                        for p_name in promos_disponibles:
                            df_p = df[df["Promotion"] == p_name].copy()
                            
                            # Reconstruction de la grille
                            grid_p = df_p.groupby(['h_norm', 'j_norm']).apply(fmt_p, include_groups=False).unstack('j_norm')
                            grid_p = grid_p.reindex(index=idx_h, columns=cols_j).fillna("")
                            grid_p = grid_p[grid_p.any(axis=1)]
                            
                            # Nettoyage des cellules avant conversion
                            grid_p = grid_p.map(clean_html_content)
                            
                            grid_p.index = [map_h.get(i, i) for i in grid_p.index]
                            grid_p.columns = [map_j.get(c, c) for c in grid_p.columns]

                            # Conversion HTML (escape=False pour garder les <br>)
                            table_html = grid_p.to_html(escape=False, classes='edt-table')

                            full_html = f"""
                            <!DOCTYPE html>
                            <html lang='fr'>
                            <head>
                                <meta charset='UTF-8'>
                                <title>EDT {p_name}</title>
                                {style_edt_html}
                            </head>
                            <body>
                                <div class='container'>
                                    <h2>PROMOTION : {p_name}</h2>
                                    <p style='text-align:center; color:#64748b;'>Faculté de Génie Électrique - UDL-SBA</p>
                                    {table_html}
                                </div>
                            </body>
                            </html>
                            """
                            zip_file.writestr(f"EDT_{p_name}_2027.html", full_html)

                    st.success(f"✅ Pack généré ({len(promos_disponibles)} fichiers)")
                    st.download_button(
                        label="⬇️ Télécharger le Pack ZIP",
                        data=zip_buffer.getvalue(),
                        file_name="Pack_EDT_HTML_S2.zip",
                        mime="application/zip",
                        use_container_width=True
                    )
                except Exception as e:
                    st.error(f"Erreur technique : {e}")
            # --- 8.3. GÉNÉRATION DU PACK PDF (ANTI-COLLISION & HAUTEUR ADAPTATIVE) ---
            st.markdown("---")
            if st.button("📁 Générer le Pack PDF (Haute Lisibilité - Pleine Page)", use_container_width=True):
                try:
                    from fpdf import FPDF
                    import re

                    class EDT_PDF(FPDF):
                        def header(self): pass
                        def footer(self): pass

                    def clean_text_for_pdf(html_str):
                        """Nettoie le HTML en préservant la structure verticale pour la lisibilité"""
                        if not html_str: return ""
                        # Conversion des balises structurelles en sauts de ligne
                        t = html_str.replace('</div>', '\n').replace('<br>', '\n').replace('<br/>', '\n')
                        t = t.replace('<b>','').replace('</b>','')
                        # Nettoyage HTML et suppression des caractères spéciaux/émojis
                        t = re.sub(r'<[^>]+>', '', t)
                        t = re.sub(r'[^\x00-\x7F]+', '', t) 
                        lines = [l.strip() for l in t.split('\n') if l.strip()]
                        return "\n".join(lines)

                    # Initialisation en Paysage A4
                    pdf = EDT_PDF(orientation="L", unit="mm", format="A4")
                    # Marges minimales pour occuper toute la surface
                    pdf.set_margins(left=7, top=7, right=7)
                    pdf.set_auto_page_break(auto=True, margin=10)
                    
                    promotions_liste = sorted(df["Promotion"].unique())
                    
                    for p_name in promotions_liste:
                        pdf.add_page()
                        
                        # Titre de la promotion centré et aéré
                        pdf.set_font("Arial", "B", 12)
                        pdf.set_text_color(0, 0, 0)
                        pdf.cell(0, 10, f"PROMOTION : {p_name}", 0, 1, "C")
                        pdf.ln(2)
                        
                        # Préparation des données de la promotion
                        df_p = df[df["Promotion"] == p_name].copy()
                        grid_p = df_p.groupby(['h_norm', 'j_norm']).apply(fmt_p, include_groups=False).unstack('j_norm')
                        grid_p = grid_p.reindex(index=idx_h, columns=cols_j).fillna("")
                        grid_p = grid_p[grid_p.any(axis=1)] # Ne garde que les créneaux avec des cours
                        
                        grid_p.index = [map_h.get(i, i) for i in grid_p.index]
                        grid_p.columns = [map_j.get(c, c) for c in grid_p.columns]
                        
                        # Dimensions des colonnes
                        col_h_w = 26
                        col_j_w = (pdf.w - col_h_w - 20) / len(grid_p.columns)
                        
                        # --- EN-TÊTE DU TABLEAU ---
                        pdf.set_font("Arial", "B", 9)
                        pdf.set_fill_color(230, 230, 230)
                        pdf.cell(col_h_w, 10, "HORAIRE", 1, 0, "C", True)
                        for jour in grid_p.columns:
                            pdf.cell(col_j_w, 10, jour, 1, 0, "C", True)
                        pdf.ln()
                        
                        # --- CORPS DU TABLEAU ---
                        # Police réduite pour la lisibilité des zones denses
                        pdf.set_font("Arial", "", 6.5)
                        interline = 3.5 # Hauteur d'une ligne de texte
                        padding_v = 3   # Marge interne (padding) pour ne pas toucher les lignes
                        
                        for heure, row in grid_p.iterrows():
                            # 1. Calculer la hauteur nécessaire pour la ligne (basé sur le jour le plus rempli)
                            max_h_needed = 14 # Hauteur minimum
                            row_texts = []
                            
                            for jour in grid_p.columns:
                                txt_propre = clean_text_for_pdf(str(row[jour]))
                                row_texts.append(txt_propre)
                                
                                # Calcul : (Nombre de lignes * hauteur de ligne) + padding haut et bas
                                nb_l = txt_propre.count('\n') + 1
                                h_calc = (nb_l * interline) + (padding_v * 2)
                                if h_calc > max_h_needed:
                                    max_h_needed = h_calc
                            
                            # 2. Rendu de la cellule Heure (Grise)
                            pdf.set_font("Arial", "B", 7.5)
                            pdf.set_fill_color(245, 245, 245)
                            pdf.cell(col_h_w, max_h_needed, heure, 1, 0, "C", True)
                            
                            # 3. Rendu des cellules Jours (Colorées selon type)
                            pdf.set_font("Arial", "", 6.5)
                            for idx, jour in enumerate(grid_p.columns):
                                content = row_texts[idx]
                                raw_val = str(row[jour]).upper()
                                
                                # Attribution des couleurs
                                if "COURS" in raw_val: pdf.set_fill_color(225, 238, 255)
                                elif "TD" in raw_val: pdf.set_fill_color(232, 252, 235)
                                elif "TP" in raw_val: pdf.set_fill_color(255, 235, 235)
                                else: pdf.set_fill_color(255, 255, 255)
                                
                                cur_x, cur_y = pdf.get_x(), pdf.get_y()
                                # Dessin de la bordure et du fond
                                pdf.rect(cur_x, cur_y, col_j_w, max_h_needed, 'FD')
                                
                                # Calcul du centrage vertical pour le padding
                                nb_l_c = content.count('\n') + 1
                                text_block_h = nb_l_c * interline
                                margin_top = (max_h_needed - text_block_h) / 2
                                
                                pdf.set_xy(cur_x, cur_y + margin_top)
                                
                                # Écriture du texte multi-ligne
                                pdf.multi_cell(col_j_w, interline, content, 0, "C")
                                pdf.set_xy(cur_x + col_j_w, cur_y)
                            
                            pdf.ln(max_h_needed)

                    # --- EXPORT ---
                    # Export PDF (existant)
                    pdf_final = pdf.output()
                    st.success(f"✅ Pack PDF généré avec succès ({len(promotions_liste)} pages).")
                    
                    st.download_button(
                        label="⬇️ Télécharger le Pack PDF (Version Corrigée)",
                        data=bytes(pdf_final),
                        file_name="Pack_EDT_S2_2027_Lisible.pdf",
                        mime="application/pdf",
                        use_container_width=True
                    )

                    # --- AJOUT EXPORT EXCEL ---
                    import pandas as pd
                    import io

                    # Conversion des données en Excel (Utilise le DataFrame source de votre tableau)
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                        # On suppose que 'df' est votre DataFrame contenant les colonnes :
                        # Enseignements, Code, Enseignants, Horaire, Jours, Lieu, Promotion
                        df.to_excel(writer, index=False, sheet_name='EDT_S2_2027')
                    
                    excel_data = buffer.getvalue()

                    st.download_button(
                        label="⬇️ Télécharger le Pack Excel (Format .xlsx)",
                        data=excel_data,
                        file_name="Pack_EDT_S2_2027.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

                except Exception as e:
                    st.error(f"Une erreur est survenue lors de la génération : {e}")

        elif is_admin and mode_view == "🏢 Planning Salles":
            s_sel = st.selectbox("Choisir Salle :", sorted(df["Lieu"].unique()))
            df_s = df[df["Lieu"] == s_sel]
            
            def fmt_s(rows):
                items = [f"<b>{r['Promotion']}</b><br>{r['Enseignements']}<br><i>{r['Enseignants']}</i>" for _, r in rows.iterrows()]
                return "<div class='separator'></div>".join(items)
                
            grid_s = df_s.groupby(['h_norm', 'j_norm']).apply(fmt_s, include_groups=False).unstack('j_norm')
            grid_s = grid_s.reindex(index=[normalize(h) for h in horaires_list], columns=[normalize(j) for j in jours_list]).fillna("")
            grid_s.index = horaires_list
            grid_s.columns = jours_list
            
            # Affichage Écran
            st.write(grid_s.to_html(escape=False), unsafe_allow_html=True)

            # --- SECTION TÉLÉCHARGEMENT ---
            st.markdown("---")
            cs1, cs2 = st.columns(2)

            # 1. EXCEL
            import io
            buf_s = io.BytesIO()
            # On exporte la liste brute pour l'Excel (plus exploitable)
            df_s.drop(columns=['h_norm', 'j_norm'], errors='ignore').to_excel(buf_s, index=False)
            cs1.download_button(
                label=f"📥 Liste {s_sel} (Excel)",
                data=buf_ex.getvalue() if 'buf_ex' in locals() else buf_s.getvalue(), # Sécurité buffer
                file_name=f"Planning_{s_sel}_2027.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key=f"xl_salle_{s_sel}"
            )

            # 2. PDF (Centrage et Marges de sécurité)
            try:
                from fpdf import FPDF
                import re

                class SALLE_PDF(FPDF):
                    def header(self):
                        self.set_font('Arial', 'B', 10)
                        t = "Plateforme de gestion des EDTs-S2-2027-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA"
                        self.cell(0, 8, t.encode('latin-1', 'replace').decode('latin-1'), 0, 1, 'C')
                        self.ln(2)
                    
                    def get_nb_lines(self, w, txt, m_h=6):
                        if not txt: return 1
                        lines = 0
                        for p in txt.split('\n'):
                            lines += max(1, int(self.get_string_width(p) / (w - m_h)) + 1)
                        return lines

                def clean_salle_pdf(val):
                    if not val: return ""
                    # Nettoie les balises HTML de fmt_s pour le PDF
                    t = str(val).replace('<b>','').replace('</b>','').replace('<i>','').replace('</i>','')
                    t = t.replace('<br>', '\n').replace('<div class=\'separator\'></div>', '\n---\n')
                    return t.encode('latin-1', 'replace').decode('latin-1')

                pdf = SALLE_PDF(orientation="L", unit="mm", format="A4")
                pdf.set_margins(7, 10, 7)
                pdf.add_page()

                # Titre
                pdf.set_font("Arial", "B", 12)
                pdf.cell(0, 10, f"PLANNING SALLE : {s_sel}".encode('latin-1', 'replace').decode('latin-1'), 0, 1, "C")
                pdf.ln(3)

                # Config PDF
                col_h_w = 32
                m_h = 6 # Marge pour ne pas toucher les traits verticaux
                col_j_w = (pdf.w - col_h_w - 20) / len(grid_s.columns)
                interline = 3.6
                padding_v = 4 # Marge pour ne pas toucher les traits horizontaux

                # Header Tableau
                pdf.set_font("Arial", "B", 8)
                pdf.set_fill_color(230, 230, 230)
                pdf.cell(col_h_w, 10, "HORAIRE", 1, 0, "C", True)
                for j in grid_s.columns:
                    pdf.cell(col_j_w, 10, str(j).upper().encode('latin-1', 'replace').decode('latin-1'), 1, 0, "C", True)
                pdf.ln()

                # Remplissage
                for horaire, row in grid_s.iterrows():
                    texts = [clean_salle_pdf(v) for v in row]
                    
                    # Balayage hauteur
                    pdf.set_font("Arial", "", 6)
                    max_h = 12
                    for t in texts:
                        n = pdf.get_nb_lines(col_j_w, t, m_h=m_h)
                        h_c = (n * interline) + (padding_v * 2)
                        if h_c > max_h: max_h = h_c

                    # Cellule Horaire
                    pdf.set_font("Arial", "B", 7.5)
                    pdf.set_fill_color(248, 248, 248)
                    pdf.cell(col_h_w, max_h, str(horaire), 1, 0, "C", True)

                    # Cellules Contenu
                    pdf.set_font("Arial", "", 6.5)
                    for idx, content in enumerate(texts):
                        # Détection couleur (Cours/TD/TP)
                        raw = str(row.iloc[idx]).upper()
                        if "COURS" in raw: pdf.set_fill_color(225, 238, 255)
                        elif "TD" in raw: pdf.set_fill_color(232, 252, 235)
                        elif "TP" in raw: pdf.set_fill_color(255, 235, 235)
                        else: pdf.set_fill_color(255, 255, 255)

                        x, y = pdf.get_x(), pdf.get_y()
                        pdf.rect(x, y, col_j_w, max_h, 'FD')
                        
                        # Calcul centrage
                        n_l = pdf.get_nb_lines(col_j_w, content, m_h=m_h)
                        th = n_l * interline
                        
                        # Dessin texte avec marges de sécurité
                        pdf.set_xy(x + (m_h/2), y + (max_h - th) / 2)
                        pdf.multi_cell(col_j_w - m_h, interline, content, 0, "C")
                        pdf.set_xy(x + col_j_w, y)
                    pdf.ln(max_h)

                cs2.download_button(
                    label=f"📄 Planning {s_sel} (PDF)",
                    data=bytes(pdf.output()),
                    file_name=f"Planning_{s_sel}_2027.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                    key=f"pdf_salle_{s_sel}"
                )
            except Exception as e:
                cs2.error(f"Erreur PDF : {e}")

        elif is_admin and mode_view == "🚩 Vérificateur de conflits":
            st.subheader("🚩 Analyse des Conflits Individuels")
            st.markdown("---")
            
            errs_text = []      
            errs_for_df = []    
            
            # --- 1. DÉTECTION DES CONFLITS (ENSEIGNANTS, SALLES ET PROMOS) ---
            errs_text = []      
            errs_for_df = []    

            # A. CONFLITS D'ENSEIGNANTS (Un prof ne peut pas être à 2 lieux/matières)
            p_groups = df[df["Enseignants"] != "Non défini"].groupby(['Jours', 'Horaire', 'Enseignants'])
            for (jour, horaire, prof), group in p_groups:
                lieux_uniques = group['Lieu'].unique()
                matieres_uniques = group['Enseignements'].unique()
                if len(lieux_uniques) > 1 or len(matieres_uniques) > 1:
                    type_err = "❌ CONFLIT ENSEIGNANT"
                    style = "error"
                    detail = f"L'enseignant est affecté à plusieurs lieux ({', '.join(lieux_uniques)}) ou matières."
                    
                    msg = f"**{type_err}** : {prof} | {jour} {horaire}"
                    errs_text.append((style, msg))
                    errs_for_df.append({
                        "Type": type_err, "Enseignant": prof, "Jour": jour, "Horaire": horaire, 
                        "Détail": detail, "Lieu": ", ".join(lieux_uniques), 
                        "Matières": ", ".join(matieres_uniques), "Promotions": ", ".join(group['Promotion'].unique())
                    })

            # B. CONFLITS DE SALLES (Deux profs différents dans la même salle) -> RÉSOUT VOTRE PROBLÈME
            s_groups = df[(df["Lieu"] != "Non défini") & (df["Lieu"] != "A distance")].groupby(['Jours', 'Horaire', 'Lieu'])
            for (jour, horaire, lieu), group in s_groups:
                if len(group['Enseignants'].unique()) > 1:
                    type_err = "❌ CONFLIT SALLE OCCUPÉE"
                    style = "error"
                    profs_concernees = group['Enseignants'].unique()
                    detail = f"La salle '{lieu}' est utilisée par : {', '.join(profs_concernees)}"
                    
                    msg = f"**{type_err}** : {lieu} | {jour} {horaire} ({', '.join(profs_concernees)})"
                    errs_text.append((style, msg))
                    
                    # On ajoute l'erreur pour chaque enseignant impliqué pour qu'ils la voient dans leur filtre
                    for p in profs_concernees:
                        errs_for_df.append({
                            "Type": type_err, "Enseignant": p, "Jour": jour, "Horaire": horaire, 
                            "Détail": detail, "Lieu": lieu, 
                            "Matières": ", ".join(group['Enseignements'].unique()), 
                            "Promotions": ", ".join(group['Promotion'].unique())
                        })

            # C. CONFLITS DE PROMOTION (Une classe ne peut pas avoir deux cours en même temps)
            pr_groups = df[df["Promotion"] != "Non défini"].groupby(['Jours', 'Horaire', 'Promotion'])
            for (jour, horaire, promo), group in pr_groups:
                if len(group['Enseignements'].unique()) > 1:
                    type_err = "⚠️ CONFLIT PROMOTION"
                    style = "warning"
                    matieres = group['Enseignements'].unique()
                    detail = f"La promotion {promo} a plusieurs cours simultanés : {', '.join(matieres)}"
                    
                    msg = f"**{type_err}** : {promo} | {jour} {horaire}"
                    errs_text.append((style, msg))
                    errs_for_df.append({
                        "Type": type_err, "Enseignant": "Multi-enseignants", "Jour": jour, "Horaire": horaire, 
                        "Détail": detail, "Lieu": ", ".join(group['Lieu'].unique()), 
                        "Matières": ", ".join(matieres), "Promotions": promo
                    })

            # --- 2. INTERFACE DE FILTRAGE ET BOUTON RESET ---
            if errs_for_df:
                st.markdown("### 🔍 Résolution ciblée")
                
                # Récupération de la liste des enseignants ayant au moins un conflit
                profs_en_conflit = sorted(list(set([e["Enseignant"] for e in errs_for_df])))
                options_menu = ["Tous"] + profs_en_conflit

                # Initialisation de la clé dans le session_state si elle n'existe pas
                if "filtre_prof_conflit" not in st.session_state:
                    st.session_state.filtre_prof_conflit = "Tous"

                # Sélecteur d'enseignant
                selected_prof = st.selectbox(
                    "🎯 Filtrer par enseignant :", 
                    options=options_menu,
                    key="filtre_prof_conflit"
                )

                # --- LE BOUTON RESET ---
                if selected_prof != "Tous":
                    st.write("") # Espacement visuel
                    if st.button("🔄 Réinitialiser la vue (Afficher tout)", use_container_width=True):
                        # Suppression sécurisée pour éviter l'erreur StreamlitAPIException
                        if "filtre_prof_conflit" in st.session_state:
                            del st.session_state.filtre_prof_conflit
                        st.rerun()

                st.divider()

                # --- 3. AFFICHAGE DES DÉTAILS (SI FILTRÉ) ---
                if selected_prof != "Tous":
                    st.info(f"Analyse précise pour : **{selected_prof}**")
                    
                    # Filtrage des erreurs pour l'enseignant sélectionné
                    conflits_specifiques = [e for e in errs_for_df if e["Enseignant"] == selected_prof]
                    
                    for i, cp in enumerate(conflits_specifiques):
                        with st.expander(f"📌 {cp['Type']} - {cp['Jour']} {cp['Horaire']}", expanded=True):
                            st.error(f"**Problème :** {cp['Détail']}")
                            
                            st.markdown("💡 **Solutions suggérées :**")
                            st.write("- Vérifiez que le nom de la matière est identique pour les deux groupes.")
                            st.write("- Modifiez l'horaire ou la salle dans l'éditeur de données.")
                            
                            # Bouton pour naviguer vers l'éditeur
                            btn_key = f"btn_solve_{cp['Enseignant']}_{i}"
                            if st.button(f"🔗 Aller à l'éditeur pour {selected_prof}", key=btn_key):
                                st.session_state.mode_view = "✍️ Éditeur de données"
                                st.rerun()

                # --- 4. RAPPORT GLOBAL ---
                st.markdown("### 🌍 Rapport Global des Anomalies")
                for style, m in errs_text:
                    # On affiche le message si on est en mode "Tous" ou si le nom du prof est dans le message
                    if selected_prof == "Tous" or selected_prof in m:
                        if style == "error":
                            st.error(m)
                        else:
                            st.warning(m)

                # --- 5. ASSISTANT DE RÉSOLUTION ET EXPORT DES SOLUTIONS ---
            if errs_for_df:
                st.divider()
                st.subheader("💡 Assistant de Résolution Intelligent")
                st.info("L'assistant propose des créneaux libres (Horaire + Salle) en respectant le type de lieu initial.")

                # On récupère la liste de tous les lieux possibles à partir du fichier
                tous_les_lieux = sorted([l for l in df['Lieu'].unique() if str(l) != "nan" and l != "Non défini"])
                
                solutions_finales = []

                # Affichage interactif pour chaque conflit
                for i, cp in enumerate(errs_for_df):
                    with st.expander(f"📍 Conflit n°{i+1} : {cp['Enseignant']} ({cp['Jour']} - {cp['Horaire']})", expanded=True):
                        c1, c2 = st.columns([2, 1])
                        
                        with c1:
                            st.error(f"**Anomalie :** {cp['Détail']}")
                            st.caption(f"Matières impliquées : {cp.get('Matières', 'N/A')}")
                        
                        with c2:
                            # 1. ANALYSE DU TYPE DE LIEU INITIAL
                            lieu_initial = str(cp['Lieu']).upper()
                            
                            # Détermination intelligente du type (Labo, Amphi, ou Salle)
                            est_tp = any(keyword in lieu_initial for keyword in ["LABO", "TP", "ATELIER", "CC", "MICRO"])
                            est_amphi = "AMPHI" in lieu_initial or "A0" in lieu_initial
                            
                            # 2. FILTRAGE DES LIEUX DU MÊME GENRE UNIQUEMENT
                            lieux_compatibles = []
                            for l in tous_les_lieux:
                                l_str = str(l).upper()
                                if est_tp and any(k in l_str for k in ["LABO", "TP", "CC", "MICRO"]):
                                    lieux_compatibles.append(l)
                                elif est_amphi and ("AMPHI" in l_str or "A0" in l_str):
                                    lieux_compatibles.append(l)
                                elif not est_tp and not est_amphi and ("S" in l_str or "SALLE" in l_str):
                                    lieux_compatibles.append(l)

                            # 3. RECHERCHE DE CRÉNEAUX ET LIEUX DISPONIBLES (Même Jour)
                            tous_horaires = ["8h - 9h30", "9h30 - 11h", "11h - 12h30", "12h30 - 14h", "14h - 15h30", "15h30 - 17h"]
                            suggestions_valides = []
                            
                            for hor in tous_horaires:
                                # A. Vérifier si l'ENSEIGNANT est libre à cette heure 'hor' ce jour-là
                                # (On ignore la vérification pour "ND" car c'est un placeholder multi-profs)
                                prof_occupe = False
                                if cp['Enseignant'] not in ["ND", "Multi-enseignants"]:
                                    prof_occupe = not df[(df['Jours'] == cp['Jour']) & 
                                                         (df['Horaire'] == hor) & 
                                                         (df['Enseignants'] == cp['Enseignant'])].empty
                                
                                if not prof_occupe:
                                    # B. Vérifier quels LIEUX COMPATIBLES sont libres à cette heure 'hor'
                                    lieux_occupes = df[(df['Jours'] == cp['Jour']) & 
                                                       (df['Horaire'] == hor)]['Lieu'].unique()
                                    
                                    libres = [l for l in lieux_compatibles if l not in lieux_occupes]
                                    
                                    for salle_libre in libres:
                                        # Éviter de proposer l'option qui est déjà en conflit
                                        if not (hor == cp['Horaire'] and salle_libre in cp['Lieu']):
                                            suggestions_valides.append(f"{hor} en {salle_libre}")

                            # 4. SÉLECTEUR DE SOLUTION
                            choix_sol = st.selectbox(
                                "🚀 Solution (Heure + Lieu compatible) :",
                                options=["-- Garder actuel --"] + suggestions_valides[:30], # Top 30 suggestions
                                key=f"assistant_sol_{i}",
                                help="Propose uniquement des créneaux où l'enseignant et la salle sont libres."
                            )
                        
                        # Construction de la ligne pour le rapport Excel final
                        solutions_finales.append({
                            "Type de Conflit": cp['Type'],
                            "Personne/Salle concernée": cp['Enseignant'] if cp['Enseignant'] != "Multi-enseignants" else cp['Détail'],
                            "Jour": cp['Jour'],
                            "Horaire Initial": cp['Horaire'],
                            "Lieu Initial": cp['Lieu'],
                            "SOLUTION PROPOSÉE": choix_sol if choix_sol != "-- Garder actuel --" else "À CORRIGER MANUELLEMENT"
                        })

                # --- 6. ACTIONS : GÉNÉRATION DU RAPPORT ET RÉINITIALISATION ---
                st.divider()
                st.markdown("### 📥 Actions sur le plan de correction")
                
                col_down, col_reset = st.columns(2)

                with col_down:
                    df_sol = pd.DataFrame(solutions_finales)
                    buf_sol = io.BytesIO()
                    with pd.ExcelWriter(buf_sol, engine='xlsxwriter') as writer:
                        df_sol.to_excel(writer, index=False, sheet_name='Solutions_Proposees')
                        
                        workbook = writer.book
                        worksheet = writer.sheets['Solutions_Proposees']
                        header_fmt = workbook.add_format({
                            'bold': True, 'bg_color': '#10B981', 'font_color': 'white', 'border': 1
                        })
                        
                        for col_num, value in enumerate(df_sol.columns.values):
                            worksheet.write(0, col_num, value, header_fmt)
                        worksheet.set_column('A:F', 25)

                    st.download_button(
                        label="💾 Télécharger le Tableau des Solutions (Excel)",
                        data=buf_sol.getvalue(),
                        file_name=f"Solutions_Conflits_EDT_S2_2027.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        type="primary"
                    )

                with col_reset:
                    if st.button("🔄 Réinitialiser tous les choix", use_container_width=True):
                        for key in list(st.session_state.keys()):
                            if key.startswith("assistant_sol_"):
                                del st.session_state[key]
                        st.rerun()

                st.caption("ℹ️ Utilisez ce fichier Excel pour appliquer les corrections dans l'Éditeur de données.")

            else:
                st.success("✅ Félicitations ! Aucun conflit détecté dans l'emploi du temps actuel.")
                st.balloons()
    elif portail == "📅 Surveillances Examens":
        FILE_S = "surveillances_2027.xlsx"
        if os.path.exists(FILE_S):
            df_surv = pd.read_excel(FILE_S)
            df_surv.columns = [str(c).strip() for c in df_surv.columns]
            df_surv['Date_Tri'] = pd.to_datetime(df_surv['Date'], dayfirst=True, errors='coerce')
            
            for c in df_surv.columns: 
                df_surv[c] = df_surv[c].fillna("").astype(str).str.strip()
                
            c_prof = 'Surveillant(s)' if 'Surveillant(s)' in df_surv.columns else 'Enseignants'
            u_nom = user['nom_officiel']
            u_email = user.get('email', '').lower().strip()

            is_master_admin = (u_email == "milouafarid@gmail.com")

            if is_master_admin:
                tous_les_profs = []
                for entry in df_surv[c_prof].unique():
                    for p in entry.split('&'):
                        clean_p = p.strip()
                        if clean_p and clean_p not in ["nan", "Non défini", ""]:
                            tous_les_profs.append(clean_p)
                liste_profs = sorted(list(set(tous_les_profs)))
                st.success("🔓 Accès Maître : milouafarid@gmail.com")
                prof_sel = st.selectbox("🔍 Choisir un enseignant :", liste_profs)
            else:
                prof_sel = u_nom
                st.info(f"👤 Espace Personnel : **{u_nom}**")

            df_u_surv = df_surv[df_surv[c_prof].str.contains(prof_sel, case=False, na=False)].sort_values(by='Date_Tri')
            st.markdown(f"### 📋 Planning de : {prof_sel}")
            
            c1, c2, c3 = st.columns(3)
            nb_mat = len(df_u_surv[df_u_surv['Heure'].str.contains("08h|09h|10h", case=False)])
            c1.metric("Séances Total", len(df_u_surv))
            c2.metric("Matin", nb_mat)
            c3.metric("Après-midi", len(df_u_surv) - nb_mat)
            
            st.divider()

            if not df_u_surv.empty:
                for _, r in df_u_surv.iterrows():
                    st.markdown(f"""
                    <div style="background:#f9f9f9;padding:12px;border-radius:8px;border-left:5px solid #1E3A8A;margin-bottom:8px;">
                        <span style="font-weight:bold;color:#1E3A8A;">📅 {r['Jour']} {r['Date']}</span> | 🕒 {r['Heure']}<br>
                        <b>📖 {r['Matière']}</b><br>
                        <small>📍 {r['Salle']} | 🎓 {r['Promotion']} | 👥 {r[c_prof]}</small>
                    </div>""", unsafe_allow_html=True)
                
                buf = io.BytesIO()
                df_u_surv.drop(columns=['Date_Tri']).to_excel(buf, index=False)
                st.download_button(f"📥 Télécharger l'EDT de {prof_sel}", buf.getvalue(), f"Surv_{prof_sel}.xlsx")
            else:
                st.warning(f"⚠️ Aucune surveillance trouvée pour : {prof_sel}")
        else:
            st.error("Le fichier 'surveillances_2027.xlsx' est absent.")

    elif portail == "🤖 Générateur Automatique":
        if not is_admin:
            st.error("Accès réservé au Bureau des Examens.")
        else:
            st.header("⚙️ Moteur de Génération de Surveillances")
            if "effectifs_db" not in st.session_state:
                st.session_state.effectifs_db = {"ING1": [50, 4], "MCIL1": [40, 3], "L1MCIL": [288, 4], "L2ELT": [90, 2], "M1RE": [15, 1], "ING2": [16, 1]}

            with st.expander("📦 Gestion des Effectifs", expanded=False):
                data_eff = [{"Promotion": k, "Effectif Total": v[0], "Nb de Salles": v[1]} for k, v in st.session_state.effectifs_db.items()]
                edited_eff = st.data_editor(pd.DataFrame(data_eff), use_container_width=True, num_rows="dynamic", hide_index=True)
                if st.button("💾 Sauvegarder la configuration"):
                    st.session_state.effectifs_db = {row["Promotion"]: [int(row["Effectif Total"]), int(row["Nb de Salles"])] for _, row in edited_eff.iterrows()}
                    st.success("Mis à jour !")

            SRC = "surveillances_2027.xlsx"
            if os.path.exists(SRC):
                df_src = pd.read_excel(SRC)
                df_src.columns = [str(c).strip() for c in df_src.columns]
                for c in df_src.columns: df_src[c] = df_src[c].fillna("").astype(str).str.strip()
                
                C_MAT, C_RESP, C_SURV, C_DATE, C_HEURE, C_SALLE, C_PROMO = "Matière", "Chargé de matière", "Surveillant(s)", "Date", "Heure", "Salle", "Promotion"
                df_src = df_src[~df_src[C_MAT].str.contains(r'\bTP\b|\bTD\b', case=False, na=False)]
                liste_profs = sorted([p for p in df_src[C_SURV].unique() if p not in ["", "nan", "Non défini"]])

                with st.expander("⚖️ Plafonnement", expanded=True):
                    col1, col2 = st.columns(2)
                    m_base = col1.number_input("Max séances", min_value=1, value=10)
                    ratio = col2.number_input("Ratio Étud/Surv", min_value=1, value=25)
                
                p_cible = st.multiselect("🎓 Promotions :", sorted(df_src[C_PROMO].unique()))
                if st.button("🚀 GÉNÉRER LE PLANNING") and p_cible:
                    stats = {p: 0 for p in liste_profs}
                    tracker, res_list = [], []
                    for p_name in p_cible:
                        df_p = df_src[df_src[C_PROMO] == p_name].drop_duplicates(subset=[C_MAT, C_DATE, C_HEURE])
                        conf = st.session_state.effectifs_db.get(p_name, [30, 1])
                        eff_total, nb_salles = conf[0], int(conf[1])
                        for _, row in df_p.iterrows():
                            for s_idx in range(1, nb_salles + 1):
                                eff_salle = eff_total // nb_salles
                                nb_req = max(2, (eff_salle // ratio) + (1 if eff_salle % ratio > 0 else 0))
                                equipe = []
                                tri_prio = sorted(liste_profs, key=lambda x: stats[x])
                                for p in tri_prio:
                                    if len(equipe) < nb_req and stats[p] < m_base:
                                        if not any(t for t in tracker if t['D']==row[C_DATE] and t['H']==row[C_HEURE] and t['N']==p):
                                            equipe.append(p); stats[p] += 1
                                            tracker.append({'D': row[C_DATE], 'H': row[C_HEURE], 'N': p})
                                res_list.append({"Enseignements": row[C_MAT], "Code": "S2-2027", "Enseignants": " & ".join(equipe) if len(equipe) >= 2 else "⚠️ BESOIN RENFORT", "Horaire": row[C_HEURE], "Jours": row[C_DATE], "Lieu": f"Salle {s_idx}" if nb_salles > 1 else row[C_SALLE], "Promotion": f"{p_name} (S{s_idx})" if nb_salles > 1 else p_name})
                    st.session_state.df_genere = pd.DataFrame(res_list)
                    st.session_state.stats_charge = stats
                    st.rerun()

                if st.session_state.get("df_genere") is not None:
                    st.dataframe(st.session_state.df_genere, use_container_width=True, hide_index=True)
                    xlsx_buf = io.BytesIO()
                    with pd.ExcelWriter(xlsx_buf, engine='xlsxwriter') as writer: st.session_state.df_genere.to_excel(writer, index=False)
                    st.download_button("📥 TÉLÉCHARGER LE PLANNING", xlsx_buf.getvalue(), "EDT_Surveillances_2027.xlsx")

    elif portail == "👥 Portail Enseignants":
        if not is_admin:
            st.error("🚫 ACCÈS RESTREINT.")
            st.stop()

        # --- CHARGEMENT DE SECOURS DU FICHIER EXCEL ---
        if df is None or df.empty:
            if os.path.exists(NOM_FICHIER_FIXE):
                df = pd.read_excel(NOM_FICHIER_FIXE)
                df.columns = [str(c).strip() for c in df.columns]
                colonnes_cles = ['Enseignements', 'Code', 'Enseignants', 'Horaire', 'Jours', 'Lieu', 'Promotion']
                for col in colonnes_cles:
                    if col in df.columns: 
                        df[col] = df[col].fillna("Non défini").astype(str).str.strip()
                    else:
                        df[col] = "Non défini"
                df['h_norm'] = df['Horaire'].apply(normalize)
                df['j_norm'] = df['Jours'].apply(normalize)
            else:
                st.error(f"❌ Le fichier {NOM_FICHIER_FIXE} est introuvable. Veuillez vérifier le chemin.")
                st.stop()

        # --- EN-TÊTE ---
        col_l, col_t = st.columns([1, 5])
        with col_l:
            st.image("logo.PNG", width=80)
        with col_t:
            st.header("🏢 Répertoire et Envoi Automatisé")
            st.write("Plateforme de gestion des emplois du temps 2026-2027-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA")

        # 1. RÉCUPÉRATION DES DONNÉES (Supabase + Répertoire Source Excel)
        res_auth = supabase.table("enseignants_auth").select("nom_officiel, email, last_sent").execute()
        dict_auth = {str(row['nom_officiel']).strip().upper(): {
            "email": row['email'], 
            "statut": "✅ Envoyé" if row['last_sent'] else "⏳ En attente"
        } for row in res_auth.data} if res_auth.data else {}

        noms_excel = sorted([e for e in df['Enseignants'].unique() if str(e) not in ["Non défini", "nan", ""]])
        
        donnees_finales = []
        for nom in noms_excel:
            nom_key = str(nom).strip().upper()
            
            # Logique de récupération de l'email
            if nom_key in dict_auth:
                email = dict_auth[nom_key]["email"]
                etat = dict_auth[nom_key]["statut"]
            elif nom_key in repertoire_source:
                email = repertoire_source[nom_key]
                etat = "🟡 Dispo (Source Excel)"
            else:
                email = "⚠️ Mail introuvable"
                etat = "❌ Adresse non communiquée"
                
            donnees_finales.append({
                "Enseignant": nom,
                "Email": email,
                "État d'envoi": etat
            })

        # 2. BOUTONS D'ACTION
        c1, c2 = st.columns(2)
        with c1:
            if st.button("🔄 Réinitialiser les statuts (Comptes)", use_container_width=True):
                supabase.table("enseignants_auth").update({"last_sent": None}).neq("email", "").execute()
                st.success("✅ Statuts réinitialisés !")
                st.rerun()
        
        with c2:
            if st.button("🚀 Lancer l'envoi groupé", type="primary", use_container_width=True):
                import smtplib
                import io
                import os
                import pandas as pd
                from email.mime.text import MIMEText
                from email.mime.multipart import MIMEMultipart
                from email.mime.base import MIMEBase
                from email import encoders
                from datetime import datetime

                try:
                    server = smtplib.SMTP('smtp.gmail.com', 587)
                    server.starttls()
                    
                    # --- CONFIGURATION EXPÉDITEUR ---
                    expediteur_email = "chef.department.elt.fge@gmail.com"
                    mot_de_passe = "gkzs pdza yodb icvd"
                    nom_affichage = "Département d'Électrotechnique UDL-SBA"
                    
                    server.login(expediteur_email, mot_de_passe)
                    
                    for row in donnees_finales:
                        if (row["État d'envoi"] in ["⏳ En attente", "🟡 Dispo (Source Excel)"]) and "@" in str(row["Email"]):
                            nom_cible = str(row['Enseignant']).strip().upper()
                            df_perso = df[df["Enseignants"].astype(str).str.upper().str.contains(nom_cible, na=False)]
                            df_mail = df_perso[['Enseignements', 'Code', 'Enseignants', 'Horaire', 'Jours', 'Lieu', 'Promotion']]
                            
                            nb_cours = df_mail['Enseignements'].str.contains('Cours', case=False).sum()
                            nb_td = df_mail['Enseignements'].str.contains('TD', case=False).sum()
                            nb_tp = df_mail['Enseignements'].str.contains('TP', case=False).sum()

                            msg = MIMEMultipart()
                            msg['Subject'] = f"Votre Emploi du Temps S2-2027 - {row['Enseignant']}"
                            
                            # --- CORRECTION DES EN-TÊTES ---
                            msg['From'] = f"{nom_affichage} <{expediteur_email}>"
                            msg['To'] = row["Email"]
                            
                            # --- CORPS DU MESSAGE (TEXTE NON CONDENSÉ) ---
                            corps_html = f"""
                            <html>
                            <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
                                <h2 style="color: #1E3A8A; border-bottom: 2px solid #1E3A8A; padding-bottom: 10px;">
                                    Plateforme de gestion des EDTs-S2-2027-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA
                                </h2>
                                
                                <p>Sallem M./Mme <b>{row['Enseignant']}</b>,</p>
                                
                                <div style="background-color: #f8f9fa; padding: 10px; border: 1px solid #dee2e6; border-radius: 5px; margin-bottom: 15px;">
                                    <b>📊 Récapitulatif de votre charge (S2-2027) :</b><br>
                                    <ul>
                                        <li>Nombre de Cours : <b>{nb_cours}</b></li>
                                        <li>Nombre de TD : <b>{nb_td}</b></li>
                                        <li>Nombre d'unité de TP : <b>{nb_tp}</b></li>
                                    </ul>
                                </div>

                                <div style="background-color: #fff4e5; border-left: 5px solid #ffa500; padding: 15px; margin: 20px 0;">
                                    <p style="font-weight: bold; color: #d97706; margin-top: 0;">
                                        Objet : Urgent : Vérification de l’emploi du temps – Semestre 2
                                    </p>
                                    
                                    <p>Cher collègue, Sallem,</p>
                                    
                                    <p>Vous trouverez ci-joint votre emploi du temps individuel pour le second semestre.<br>
                                    Afin de permettre au service des enseignements d'accomplir sa mission dans les meilleures conditions, il est impératif que vous procédiez à sa vérification immédiate. Cette étape est cruciale pour :</p>
                                    
                                    <ul style="margin-top: 5px;">
                                        <li>1- Valider la charge horaire exacte de chaque enseignant.</li>
                                        <li>2- Planifier précisément le démarrage effectif des différents enseignements.</li>
                                    </ul>

                                    <p><b>🚀 Action requise :</b><br>
                                    - <b>En cas d'anomalie :</b> nous retourner le fichier Excel dûment corrigé à l'adresse d'envoi : <b>chef.department.elt.fge@gmail.com</b><br>
                                    - <b>Si tout est conforme :</b> nous répondre simplement par « <b>RAS</b> ».</p>
                                    
                                    <p>Votre retour est indispensable pour la stabilisation des emplois du temps. Sans réponse de votre part, nous ne pourrons garantir la mise à jour de vos charges pédagogiques.<br>
                                    <span style="color: #b91c1c; font-weight: bold;">""</span></p>
                                    
                                    <p><b>Saha Ftourkoum</b></p>
                                </div>

                                <div style="margin: 20px 0;">
                                    {df_mail.to_html(index=False, border=1, justify='center')}
                                </div>
                                
                                <p>Cordialement.</p>
                                <hr>
                                <p style="color: #555;">
                                    <b>Service d'enseignement</b><br>
                                    Département d'Électrotechnique<br>
                                    Faculté de Génie Électrique (FGE)
                                </p>
                            </body>
                            </html>
                            """
                            msg.attach(MIMEText(corps_html, 'html'))

                            # --- GÉNÉRATION EXCEL ET ATTACHEMENT ---
                            buffer = io.BytesIO()
                            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                                df_mail.to_excel(writer, index=False, sheet_name='Mon EDT')
                                workbook = writer.book
                                worksheet = writer.sheets['Mon EDT']
                                fmt_header = workbook.add_format({'bold': True, 'bg_color': '#4472C4', 'font_color': 'white', 'border': 1})
                                for col_num, value in enumerate(df_mail.columns.values):
                                    worksheet.write(0, col_num, value, fmt_header)
                                worksheet.set_column('A:G', 20)
                            
                            buffer.seek(0)
                            part = MIMEBase('application', 'vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                            part.set_payload(buffer.read())
                            encoders.encode_base64(part)
                            part.add_header('Content-Disposition', f'attachment; filename="EDT_S2_2027_{row["Enseignant"]}.xlsx"')
                            msg.attach(part)
                            
                            server.send_message(msg)
                            
                            # Mise à jour optionnelle si vous utilisez Supabase
                            try:
                                supabase.table("enseignants_auth").update({"last_sent": datetime.now().isoformat()}).eq("email", row["Email"]).execute()
                            except:
                                pass
                    
                    server.quit()
                    st.success("✅ Envoi groupé terminé !")
                    st.rerun()
                except Exception as e:
                    st.error(f"Erreur : {e}")
        # 3. AFFICHAGE DU TABLEAU RECAPITULATIF
        st.divider()
        st.dataframe(pd.DataFrame(donnees_finales), use_container_width=True, hide_index=True)

        # --- 3. GESTION DES ENVOIS PERSONNALISÉS ---
st.divider()
st.subheader("📬 Gestion des envois personnalisés")

# --- CONFIGURATION SMTP ---
EMAIL_EXPEDITEUR = "chef.department.elt.fge@gmail.com"
SECRET_APP = "gkzs pdza yodb icvd"

# --- PRÉPARATION DES DONNÉES ---
donnees_finales = []
if df is not None:
    for ens in sorted(df["Enseignants"].unique()):
        email = repertoire_source.get(str(ens).strip().upper(), "Non communiquée")
        donnees_finales.append({
            "Enseignant": ens,
            "Email": email,
            "État d'envoi": "✅ Prêt" if email != "Non communiquée" else "❌ Adresse non communiquée"
        })

mode_envoi = st.radio("Choisir le mode d'envoi :", 
                      ["Un par un (Individuel)", "Sélection groupée (Multi-choix)", "Par Promotion (Automatique)"], 
                      horizontal=True)

# --- LOGIQUE DE FILTRAGE COMMUNE ---
col_f1, col_f2 = st.columns(2)
with col_f1:
    liste_noms = ["TOUS"] + [d["Enseignant"] for d in donnees_finales]
    choix_enseignant = st.selectbox("🔍 Chercher un nom :", liste_noms)
with col_f2:
    choix_statut = st.selectbox("📊 Filtrer par statut :", ["TOUS", "✅ Prêt", "❌ Adresse non communiquée"])

enseignants_filtres = [
    e for e in donnees_finales 
    if (choix_enseignant == "TOUS" or e["Enseignant"] == choix_enseignant) and
       (choix_statut == "TOUS" or e["État d'envoi"] == choix_statut)
]

# --- FONCTION D'ENVOI (Pour éviter la répétition du code) ---
def envoyer_emails(liste_destinataires, promotion_label="Individuel"):
    import smtplib, io
    from email.mime.text import MIMEText
    from email.mime.multipart import MIMEMultipart
    from email.mime.base import MIMEBase
    from email import encoders

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(EMAIL_EXPEDITEUR, SECRET_APP)
        
        barre_prog = st.progress(0)
        status_txt = st.empty()
        
        for i, info in enumerate(liste_destinataires):
            nom_ens = info["Enseignant"]
            email_ens = info["Email"]
            status_txt.text(f"Envoi en cours : {nom_ens} ({i+1}/{len(liste_destinataires)})")
            
            # Extraction et mise en forme des données
            df_perso = df[df["Enseignants"].astype(str).str.contains(str(nom_ens).strip(), na=False)]
            df_mail = df_perso[['Enseignements', 'Code', 'Enseignants', 'Horaire', 'Jours', 'Lieu', 'Promotion']]
            
            nb_cours = df_mail['Enseignements'].str.contains('Cours', case=False).sum()
            nb_td = df_mail['Enseignements'].str.contains('TD', case=False).sum()
            nb_tp = df_mail['Enseignements'].str.contains('TP', case=False).sum()

            msg = MIMEMultipart()
            msg['Subject'] = f"Votre Emploi du Temps S2-2027 - {nom_ens}"
            msg['From'] = f"Département d'Électrotechnique <{EMAIL_EXPEDITEUR}>"
            msg['To'] = email_ens
            
            table_html = df_mail.to_html(index=False, border=1, justify='center')
            
            corps_html = f"""
            <html>
            <body style="font-family: Arial, sans-serif;">
                <div style="background-color: #f4f4f4; padding: 15px; border-radius: 5px; border: 1px solid #1E3A8A;">
                    <h2 style="color: #1E3A8A; text-align: center;">Plateforme de gestion des EDTs-S2-2027-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA</h2>
                    <p>Sallem M./Mme <b>{nom_ens}</b>,</p>
                    <p><b>Récapitulatif de votre charge :</b> {nb_cours} Cours, {nb_td} TD, {nb_tp} TP.</p>
                    <p style="font-weight: bold; color: #b91c1c;">Objet : Urgent : Vérification de l’emploi du temps – Semestre 2</p>
                    <p>Merci de bien renseigner le fichier Excel joint. Envoie RAS si c'est bon.</p>
                    <p style="font-size: 1.2em; color: #b91c1c; font-weight: bold; text-align: center;"></p>
                    <div style="background-color: white;">{table_html}</div>
                    <p>Cordialement.<br><b>Service d'enseignement</b></p>
                </div>
            </body>
            </html>
            """
            msg.attach(MIMEText(corps_html, 'html'))

            # Fichier Excel formaté
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
                df_mail.to_excel(writer, index=False, sheet_name='Mon EDT')
                wb = writer.book
                ws = writer.sheets['Mon EDT']
                header_fmt = wb.add_format({'bold': True, 'bg_color': '#4472C4', 'font_color': 'white', 'border': 1})
                for col_num, value in enumerate(df_mail.columns.values):
                    ws.write(0, col_num, value, header_fmt)
                ws.set_column('A:G', 18)
            
            buf.seek(0)
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(buf.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename="EDT_S2_2027_{nom_ens}.xlsx"')
            msg.attach(part)
            
            server.send_message(msg)
            barre_prog.progress((i + 1) / len(liste_destinataires))

        server.quit()
        status_txt.success(f"✅ {len(liste_destinataires)} emails envoyés avec succès !")
        st.balloons()
    except Exception as e:
        st.error(f"Erreur lors de l'envoi : {e}")

# --- AFFICHAGE SELON LE MODE ---
if mode_envoi == "Un par un (Individuel)":
    st.dataframe(pd.DataFrame(enseignants_filtres), use_container_width=True, hide_index=True)
    if st.button("🚀 ENVOYER AUX ENSEIGNANTS FILTRÉS", type="primary", use_container_width=True):
        destinataires = [e for e in enseignants_filtres if "@" in str(e["Email"])]
        if destinataires:
            envoyer_emails(destinataires)
        else:
            st.warning("Aucun email valide trouvé dans le filtre.")

elif mode_envoi == "Sélection groupée (Multi-choix)":
    noms_dispo = [e["Enseignant"] for e in enseignants_filtres if "@" in str(e["Email"])]
    selection = st.multiselect("Sélectionner les enseignants :", noms_dispo)
    if st.button("🚀 ENVOYER À LA SÉLECTION", type="primary", use_container_width=True):
        destinataires = [e for e in enseignants_filtres if e["Enseignant"] in selection]
        if destinataires:
            envoyer_emails(destinataires)
        else:
            st.warning("Veuillez sélectionner au moins un enseignant.")

elif mode_envoi == "Par Promotion (Automatique)":
    promos = sorted(df['Promotion'].unique().tolist())
    choix_promo = st.selectbox("🎯 Sélectionner la Promotion :", promos)
    ens_promo = df[df['Promotion'] == choix_promo]['Enseignants'].unique().tolist()
    liste_promo = [e for e in donnees_finales if e["Enseignant"] in ens_promo]
    
    st.write(f"### 📋 Contrôle : {choix_promo}")
    if liste_promo:
        df_export = pd.DataFrame(liste_promo)
        
        # Sélection des colonnes demandées : Nom/Prénom (Enseignant) et Email
        colonnes_export = ["Enseignant", "Email"]
        df_download = df_export[colonnes_export].drop_duplicates()

        nb_ok = sum(1 for e in liste_promo if "@" in str(e["Email"]))
        st.metric("Emails opérationnels", f"{nb_ok} / {len(liste_promo)}")
        
        st.dataframe(df_export, use_container_width=True, hide_index=True)

        # --- GÉNÉRATION DU FICHIER EXCEL ---
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_download.to_excel(writer, index=False, sheet_name='Liste_Emails')
            # Optionnel : Ajustement automatique de la largeur des colonnes
            worksheet = writer.sheets['Liste_Emails']
            for i, col in enumerate(df_download.columns):
                column_len = max(df_download[col].astype(str).map(len).max(), len(col)) + 2
                worksheet.set_column(i, i, column_len)
        
        buffer.seek(0)

        st.download_button(
            label=f"🟢 Télécharger la liste Excel ({choix_promo})",
            data=buffer,
            file_name=f"Emails_{choix_promo}_S2_2027.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        # -----------------------------------

        st.divider()

        if st.button(f"🚀 LANCER L'ENVOI POUR {choix_promo}", type="primary", use_container_width=True):
            destinataires = [e for e in liste_promo if "@" in str(e["Email"])]
            if destinataires:
                envoyer_emails(destinataires, choix_promo)
            else:
                st.error("Aucun email valide pour cette promotion.")
            # --- SECTION PRÉVISUALISATION ---
            if selection:
                st.write(f"🔍 **Prévisualisation de la sélection ({len(selection)}) :**")
                # Filtrage pour afficher uniquement les enseignants sélectionnés dans le tableau de contrôle
                donnees_previsu = [e for e in enseignants_filtres if e["Enseignant"] in selection]
                df_previsu = pd.DataFrame(donnees_previsu)[["Enseignant", "Email", "État d'envoi"]]
                st.dataframe(df_previsu, use_container_width=True, hide_index=True)
            # --------------------------------

            if st.button(f"🚀 Envoyer à la sélection ({len(selection)})", type="primary", use_container_width=True):
                if not selection:
                    st.warning("Veuillez sélectionner au moins un enseignant.")
                else:
                    import smtplib, io, pandas as pd
                    from email.mime.text import MIMEText
                    from email.mime.multipart import MIMEMultipart
                    from email.mime.base import MIMEBase
                    from email import encoders

                    try:
                        server = smtplib.SMTP('smtp.gmail.com', 587)
                        server.starttls()
                        server.login(st.secrets["EMAIL_USER"], st.secrets["EMAIL_PASS"])
                        
                        progress_bar = st.progress(0)
                        for i, nom in enumerate(selection):
                            info_ens = next(e for e in enseignants_filtres if e["Enseignant"] == nom)
                            nom_cible = str(nom).strip().upper()
                            
                            # Extraction des données spécifiques à l'enseignant pour le tableau
                            df_perso = df[df["Enseignants"].astype(str).str.upper().str.contains(nom_cible, na=False)]
                            df_mail = df_perso[['Enseignements', 'Code', 'Enseignants', 'Horaire', 'Jours', 'Lieu', 'Promotion']]
                            
                            # Calcul du récapitulatif de charge
                            nb_cours = df_mail['Enseignements'].str.contains('Cours', case=False).sum()
                            nb_td = df_mail['Enseignements'].str.contains('TD', case=False).sum()
                            nb_tp = df_mail['Enseignements'].str.contains('TP', case=False).sum()

                            msg = MIMEMultipart()
                            msg['Subject'] = f"Votre Emploi du Temps S2-2027 - {nom}"
                            msg['From'] = st.secrets["EMAIL_USER"]
                            msg['To'] = info_ens["Email"]

                            # --- CORPS DU MESSAGE (IDENTIQUE À L'INDIVIDUEL) ---
                            corps_html = f"""
                            <html>
                            <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
                                <h2 style="color: #1E3A8A; border-bottom: 2px solid #1E3A8A; padding-bottom: 10px;">
                                    Plateforme de gestion des EDTs-S2-2027-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA
                                </h2>
                                
                                <p>Sallem M./Mme <b>{row['Enseignant']}</b>,</p>
                                
                                <div style="background-color: #f8f9fa; padding: 10px; border: 1px solid #dee2e6; border-radius: 5px; margin-bottom: 15px;">
                                    <b>📊 Récapitulatif de votre charge (S2-2027) :</b><br>
                                    <ul>
                                        <li>Nombre de Cours : <b>{nb_cours}</b></li>
                                        <li>Nombre de TD : <b>{nb_td}</b></li>
                                        <li>Nombre d'unité de TP : <b>{nb_tp}</b></li>
                                    </ul>
                                </div>

                                <div style="background-color: #fff4e5; border-left: 5px solid #ffa500; padding: 15px; margin: 20px 0;">
                                    <p style="font-weight: bold; color: #d97706; margin-top: 0;">
                                        Objet : Urgent : Vérification de l’emploi du temps – Semestre 2
                                    </p>
                                    
                                    <p>Cher collègue, Sallem,</p>
                                    
                                    <p>Vous trouverez ci-joint votre emploi du temps individuel pour le second semestre.<br>
                                    Afin de permettre au service des enseignements d'accomplir sa mission dans les meilleures conditions, il est impératif que vous procédiez à sa vérification immédiate. Cette étape est cruciale pour :</p>
                                    
                                    <ul style="margin-top: 5px;">
                                        <li>1- Valider la charge horaire exacte de chaque enseignant.</li>
                                        <li>2- Planifier précisément le démarrage effectif des différents enseignements.</li>
                                    </ul>

                                    <p><b>🚀 Action requise :</b><br>
                                    - <b>En cas d'anomalie :</b> nous retourner le fichier Excel dûment corrigé à l'adresse d'envoi : <b>chef.department.elt.fge@gmail.com</b><br>
                                    - <b>Si tout est conforme :</b> nous répondre simplement par « <b>RAS</b> ».</p>
                                    
                                    <p>Votre retour est indispensable pour la stabilisation des emplois du temps. Sans réponse de votre part, nous ne pourrons garantir la mise à jour de vos charges pédagogiques.<br>
                                    <span style="color: #b91c1c; font-weight: bold;">""</span></p>
                                    
                                    <p><b>Saha Ftourkoum</b></p>
                                </div>

                                <div style="margin: 20px 0;">
                                    {df_mail.to_html(index=False, border=1, justify='center')}
                                </div>
                                
                                <p>Cordialement.</p>
                                <hr>
                                <p style="color: #555;">
                                    <b>Service d'enseignement</b><br>
                                    Département d'Électrotechnique<br>
                                    Faculté de Génie Électrique (FGE)
                                </p>
                            </body>
                            </html>
                            """
                            msg.attach(MIMEText(corps_html, 'html'))

                            # Génération de la pièce jointe Excel formatée
                            buffer = io.BytesIO()
                            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                                df_mail.to_excel(writer, index=False, sheet_name='Mon EDT')
                                workbook, worksheet = writer.book, writer.sheets['Mon EDT']
                                f_h = workbook.add_format({'bold': True, 'bg_color': '#4472C4', 'font_color': 'white', 'border': 1})
                                f_c = workbook.add_format({'bg_color': '#D9EAD3', 'border': 1})
                                f_d = workbook.add_format({'bg_color': '#FFF2CC', 'border': 1})
                                f_p = workbook.add_format({'bg_color': '#F4CCCC', 'border': 1})
                                
                                # Entête
                                for c_n, v_l in enumerate(df_mail.columns.values): 
                                    worksheet.write(0, c_n, v_l, f_h)
                                
                                # Coloration des lignes par type d'enseignement
                                for i_x, e_n in enumerate(df_mail['Enseignements']):
                                    f_r = None
                                    if 'Cours' in str(e_n): f_r = f_c
                                    elif 'TD' in str(e_n): f_r = f_d
                                    elif 'TP' in str(e_n): f_r = f_p
                                    if f_r: worksheet.set_row(i_x + 1, None, f_r)
                                worksheet.set_column('A:G', 18)
                            
                            buffer.seek(0)
                            part = MIMEBase('application', 'octet-stream')
                            part.set_payload(buffer.read())
                            encoders.encode_base64(part)
                            part.add_header('Content-Disposition', f'attachment; filename="EDT_S2_2027_{nom}.xlsx"')
                            msg.attach(part)
                            
                            # Envoi effectif
                            server.send_message(msg)
                            progress_bar.progress((i + 1) / len(selection))
                        
                        server.quit()
                        st.success(f"✅ Envoi terminé avec succès pour la sélection !")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Erreur lors de l'envoi : {e}")
        else:
            # --- MODE INDIVIDUEL (Bouton par ligne) ---
            for idx, row in enumerate(enseignants_filtres):
                col_ens, col_mail, col_stat, col_act = st.columns([2, 2, 1, 1])
                col_ens.write(f"**{row['Enseignant']}**")
                col_mail.write(row['Email'])
                col_stat.write(row["État d'envoi"])
                
                if "@" in str(row["Email"]):
                    if col_act.button("📧 Envoyer", key=f"btn_unit_{row['Enseignant']}_{idx}"):
                        import smtplib, io, pandas as pd
                        from email.mime.text import MIMEText
                        from email.mime.multipart import MIMEMultipart
                        from email.mime.base import MIMEBase
                        from email import encoders

                        try:
                            server = smtplib.SMTP('smtp.gmail.com', 587)
                            server.starttls()
                            
                            # --- CONFIGURATION EXPÉDITEUR ---
                            exp_mail = "chef.department.elt.fge@gmail.com"
                            exp_pass = "gkzs pdza yodb icvd"
                            nom_aff = "Département d'Électrotechnique UDL-SBA"
                            
                            server.login(exp_mail, exp_pass)
                            
                            nom_c = str(row['Enseignant']).strip().upper()
                            df_p = df[df["Enseignants"].astype(str).str.upper().str.contains(nom_c, na=False)]
                            df_m = df_p[['Enseignements', 'Code', 'Enseignants', 'Horaire', 'Jours', 'Lieu', 'Promotion']]
                            
                            msg = MIMEMultipart()
                            msg['Subject'] = f"Votre Emploi du Temps S2-2027 - {row['Enseignant']}"
                            
                            # --- MODIFICATION DES EN-TÊTES ---
                            msg['From'] = f"{nom_aff} <{exp_mail}>"
                            msg['To'] = row["Email"]
                            
                            # --- CORPS DU MESSAGE MIS À JOUR ---
                            corps = f"""
                            <html>
                            <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
                                <h2 style="color: #1E3A8A; border-bottom: 2px solid #1E3A8A; padding-bottom: 10px;">
                                    Plateforme de gestion des EDTs-S2-2027-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA
                                </h2>
                                
                                <p>Sallem M./Mme <b>{row['Enseignant']}</b>,</p>
                                
                                <div style="background-color: #fff4e5; border-left: 5px solid #ffa500; padding: 15px; margin: 20px 0;">
                                    <p style="font-weight: bold; color: #d97706; margin-top: 0;">
                                        Objet : Urgent : Vérification de l’emploi du temps – Semestre 2
                                    </p>
                                    
                                    <p>Cher collègue, Sallem,</p>
                                    
                                    <p>Vous trouverez ci-joint votre emploi du temps individuel pour le second semestre.<br>
                                    Afin de permettre au service des enseignements d'accomplir sa mission dans les meilleures conditions, il est impératif que vous procédiez à sa vérification immédiate. Cette étape est cruciale pour :</p>
                                    
                                    <ul style="margin-top: 5px;">
                                        <li>1- Valider la charge horaire exacte de chaque enseignant.</li>
                                        <li>2- Planifier précisément le démarrage effectif des différents enseignements.</li>
                                    </ul>

                                    <p><b>🚀 Action requise :</b><br>
                                    - <b>En cas d'anomalie :</b> nous retourner le fichier Excel dûment corrigé à l'adresse d'envoi : <b>chef.department.elt.fge@gmail.com</b><br>
                                    - <b>Si tout est conforme :</b> nous répondre simplement par « <b>RAS</b> ».</p>
                                    
                                    <p>Votre retour est indispensable pour la stabilisation des emplois du temps. Sans réponse de votre part, nous ne pourrons garantir la mise à jour de vos charges pédagogiques.<br>
                                    <span style="color: #b91c1c; font-weight: bold;">""</span></p>
                                    
                                    <p><b>Saha Ftourkoum</b></p>
                                </div>

                                <p>Voici le récapitulatif de votre emploi du temps pour le semestre 02 (S2-2027) :</p>

                                <div style="margin: 20px 0;">
                                    {df_m.to_html(index=False, border=1, justify='center')}
                                </div>
                                
                                <p>Cordialement.</p>
                                <hr>
                                <p style="color: #555;">
                                    <b>Service d'enseignement</b><br>
                                    Département d'Électrotechnique<br>
                                    Faculté de Génie Électrique (FGE)
                                </p>
                            </body>
                            </html>
                            """
                            msg.attach(MIMEText(corps, 'html'))
                            
                            buf = io.BytesIO()
                            with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
                                df_m.to_excel(writer, index=False, sheet_name='Mon EDT')
                            buf.seek(0)
                            
                            part = MIMEBase('application', 'octet-stream')
                            part.set_payload(buf.read())
                            encoders.encode_base64(part)
                            part.add_header('Content-Disposition', f'attachment; filename="EDT_2027_{row["Enseignant"]}.xlsx"')
                            msg.attach(part)
                            
                            server.send_message(msg)
                            server.quit()
                            
                            st.success(f"✅ Envoyé à {row['Enseignant']}")
                            st.rerun()
                        except Exception as e: 
                            st.error(f"Erreur : {e}")
        # =================================================================
        # =================================================================
        # SECTION COURRIER OFFICIEL : MULTI-EXPÉDITEURS (CHEF / ADJOINT / SEC)
        # =================================================================
        st.divider()
        with st.expander("✉️ ENVOYER UN COURRIER OFFICIEL (Direction / Secrétariat)", expanded=False):
            st.info("""
            **Mode Multi-Profils :** Sélectionnez votre fonction. L'email officiel correspondant sera utilisé. 
            Chaque utilisateur doit entrer son propre 'Mot de passe d'application' Google.
            """)
            
            # --- 1. CONFIGURATION DE L'EXPÉDITEUR ---
            st.subheader("🔑 1. Identification de l'expéditeur")
            
            # Configuration des profils
            options_exp = {
                "Chef de Département": "chef.department.elt.fge@gmail.com",
                "Chef de Département Adjoint": st.secrets.get("EMAIL_ADJOINT", "Non configuré"),
                "Secrétariat ELT": st.secrets.get("EMAIL_SEC", "Non configuré"),
                "Chef de départemet ELT": st.secrets.get("EMAIL_USER", "Non configuré")
            }
            
            col_auth1, col_auth2 = st.columns(2)
            
            with col_auth1:
                role_choisi = st.selectbox("Expéditeur officiel :", list(options_exp.keys()))
                expediteur_mail = options_exp[role_choisi]
                st.success(f"📧 Compte : {expediteur_mail}")
            
            with col_auth2:
                # Dictionnaire contenant vos codes de 16 lettres
                # Remplacez les textes par vos codes réels
                codes_secrets = {
                    "Chef de Département": "gkzs pdza yodb icvd", 
                    "Chef de départemet ELT": "kmtk zmkd kwpd cqzz",
                    "Chef de Département Adjoint": "", # Vide pour le moment
                    "Secrétariat ELT": ""              # Vide pour le moment
                }

                # On récupère le code selon le rôle choisi
                code_auto = codes_secrets.get(role_choisi, "")

                # Affichage du champ (rempli automatiquement si le code existe)
                expediteur_pass = st.text_input(
                    f"Mot de passe d'application ({role_choisi}) :", 
                    value=code_auto,
                    type="password", 
                    help="Le code est rempli automatiquement pour les comptes autorisés.",
                    key=f"pass_{role_choisi}" # Clé dynamique pour que Streamlit rafraîchisse bien le champ
                )

            st.divider()

            # --- 2. RÉDACTION DU MESSAGE ---
            st.subheader("📝 2. Rédaction du message")
            dict_emails = {row["Enseignant"]: row["Email"] for row in donnees_finales if "@" in str(row["Email"])}
            
            col_msg1, col_msg2 = st.columns([1, 2])
            
            with col_msg1:
                cible_courrier = st.radio("Destinataires :", ["Tous les enseignants", "Sélection spécifique"])
                destinataires_mails = []
                
                if cible_courrier == "Tous les enseignants":
                    destinataires_mails = list(dict_emails.values())
                    st.warning(f"⚠️ Envoi groupé à {len(destinataires_mails)} enseignants.")
                else:
                    selection_profs = st.multiselect("Choisir les enseignants :", sorted(dict_emails.keys()))
                    destinataires_mails = [dict_emails[p] for p in selection_profs]

            with col_msg2:
                sujet_libre = st.text_input("Objet du message :", placeholder="Ex: Convocation réunion...")
                corps_libre = st.text_area("Corps du message (Texte libre) :", height=150)
                fichier_joint = st.file_uploader("📎 Pièce jointe (PDF, Excel, Image...)", type=["pdf", "png", "jpg", "docx", "xlsx"])

            # --- 3. LOGIQUE D'ENVOI AVEC SIGNATURE AUTOMATIQUE ---
            if st.button("🚀 LANCER L'ENVOI OFFICIEL", type="primary", use_container_width=True):
                if not expediteur_pass:
                    st.error(f"❌ Veuillez saisir le mot de passe d'application pour {expediteur_mail}")
                elif not destinataires_mails:
                    st.error("❌ Aucun destinataire sélectionné.")
                elif not sujet_libre or not corps_libre:
                    st.error("❌ L'objet et le corps du message sont obligatoires.")
                else:
                    try:
                        import smtplib, mimetypes
                        from email.mime.multipart import MIMEMultipart
                        from email.mime.text import MIMEText
                        from email.mime.base import MIMEBase
                        from email import encoders

                        # Génération de la signature selon le rôle
                        if role_choisi == "Chef de Département":
                            signature = (
                                "\n\n---\n"
                                "Cordialement,\n\n"
                                "Pr. MILOUA Farid\n"
                                "Chef de Département d'Électrotechnique\n"
                                "Faculté de Génie Électrique (FGE)\n"
                                "Université Djillali Liabes (UDL-SBA)"
                            )
                        elif role_choisi == "Chef de Département Adjoint":
                            signature = "\n\n---\nCordialement,\nChef de Département Adjoint\nDépartement d'Électrotechnique - FGE - UDL-SBA"
                        elif role_choisi == "Secrétariat ELT":
                            signature = "\n\n---\nSecrétariat du Département d'Électrotechnique\nFGE - UDL-SBA"
                        else:
                            signature = "\n\n---\nEnvoyé via la Plateforme de Gestion des EDTs (FGE-UDL-SBA)"

                        corps_final = corps_libre + signature

                        with st.spinner(f"Envoi en cours par le {role_choisi}..."):
                            server = smtplib.SMTP('smtp.gmail.com', 587)
                            server.starttls()
                            server.login(expediteur_mail, expediteur_pass)

                            msg = MIMEMultipart()
                            msg['From'] = f"{role_choisi} <{expediteur_mail}>"
                            msg['To'] = ", ".join(destinataires_mails)
                            msg['Subject'] = sujet_libre
                            
                            # Support complet des accents et caractères spéciaux
                            msg.attach(MIMEText(corps_final, 'plain', 'utf-8'))

                            if fichier_joint:
                                content_type, _ = mimetypes.guess_type(fichier_joint.name)
                                main_type, sub_type = (content_type or 'application/octet-stream').split('/', 1)
                                part = MIMEBase(main_type, sub_type)
                                part.set_payload(fichier_joint.read())
                                encoders.encode_base64(part)
                                part.add_header('Content-Disposition', f'attachment; filename="{fichier_joint.name}"')
                                msg.attach(part)

                            server.send_message(msg)
                            server.quit()
                        
                        st.success(f"✅ Courrier de la part de {role_choisi} envoyé avec succès !")
                        st.balloons()
                        
                    except Exception as e:
                        st.error(f"❌ Erreur technique : {e}")
                        st.info("💡 Rappel : Vérifiez votre connexion et votre code de 16 lettres.")
        # =================================================================
        # =================================================================
# --- LOGIQUE D'AFFICHAGE DU PORTAIL MISE À JOUR ---
if portail == "🎓 Portail mise à jour EDT":
    st.write(f"**MODE ACTIF :** {portail}")
    st.subheader("📚 Espace mise à jour EDT")
    
    # Rappel du titre obligatoire
    st.info("Plateforme de gestion des EDTs-S2-2027-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA")

    # --- 1. AFFICHAGE DE L'EMPLOI DU TEMPS (VUE ÉTUDIANT) ---
    st.markdown("### 📋 Consultation par Promotion")
    
    # Récupération sécurisée des promotions
    if df is not None and not df.empty:
        liste_promotions = sorted(df["Promotion"].unique().tolist())
    else:
        liste_promotions = ["ING1", "L3-ELT", "M1-RE", "M2-RE"] # Valeurs par défaut si fichier vide

    choix_promo = st.selectbox("Choisir votre Promotion :", liste_promotions)
    
    # Disposition stricte : Enseignements, Code, Enseignants, Horaire, Jours, Lieu, Promotion
    colonnes_ordonnees = ['Enseignements', 'Code', 'Enseignants', 'Horaire', 'Jours', 'Lieu', 'Promotion']
    
    # Filtrage du tableau de vue
    df_vue = df[df["Promotion"] == choix_promo][colonnes_ordonnees].sort_values(by=["Jours", "Horaire"])
    
    st.write(f"**Emploi du temps actuel : {choix_promo}**")
    st.table(df_vue)

    # Boutons de téléchargement pour la vue actuelle
    c1, c2 = st.columns(2)
    with c1:
        # Export Excel
        output_vue = io.BytesIO()
        with pd.ExcelWriter(output_vue, engine='xlsxwriter') as writer:
            df_vue.to_excel(writer, index=False)
        st.download_button("📊 Télécharger Excel (Vue actuelle)", output_vue.getvalue(), f"EDT_{choix_promo}.xlsx")
    with c2:
        # Export HTML
        st.download_button("📄 Télécharger HTML (Vue actuelle)", df_vue.to_html(index=False), f"EDT_{choix_promo}.html", "text/html")

    # --- 2. ESPACE ADMINISTRATEUR (ÉDITION & AJOUT DE LIGNE) ---
    if is_admin:
        st.write("---")
        st.subheader("✍️ Espace Éditeur de Données (Admin)")
        st.info("💡 Pour ajouter une charge : Filtrez pour isoler l'EDT concerné, puis cliquez sur le (+) en bas du tableau.")

        # Barre de recherche pour filtrer l'éditeur
        recherche = st.text_input("🔍 Rechercher une ligne (Enseignant, Salle ou Code) :", key="admin_search_bar")

        # Préparation du DataFrame maître
        df_master = df[colonnes_ordonnees].copy()
        
        # Application du filtre si recherche active
        if recherche:
            masque = df_master.apply(lambda r: r.astype(str).str.contains(recherche, case=False).any(), axis=1)
            df_edition = df_master[masque].copy()
        else:
            df_edition = df_master.copy()

        # Affichage du compteur de lignes pour le suivi de l'index (Ex: 532)
        total_lignes = len(df)
        st.caption(f"Lignes totales dans le fichier source : {total_lignes} | Prochain index : {total_lignes}")

        # L'ÉDITEUR DYNAMIQUE
        # num_rows="dynamic" permet d'ajouter la ligne 533 etc.
        df_edite = st.data_editor(
            df_edition,
            use_container_width=True,
            num_rows="dynamic",
            key="admin_data_editor_main"
        )

        # Boutons de téléchargement pour l'éditeur (données filtrées)
        ca1, ca2 = st.columns(2)
        with ca1:
            out_ed = io.BytesIO()
            df_edite.to_excel(out_ed, index=False)
            st.download_button("📊 Télécharger l'EDT filtré (Excel)", out_ed.getvalue(), "EDT_Edition.xlsx")
        with ca2:
            st.download_button("📄 Télécharger l'EDT filtré (HTML)", df_edite.to_html(index=False), "EDT_Edition.html", "text/html")

        # --- 3. LOGIQUE DE SAUVEGARDE ET AUTO-INDEXATION ---
        if st.button("💾 Sauvegarder les modifications et la nouvelle charge"):
            try:
                if recherche:
                    # On fusionne : (Tout ce qui n'était pas affiché) + (Ce qui est dans l'éditeur + ajouts)
                    df_final = pd.concat([df_master[~masque], df_edite], ignore_index=True)
                else:
                    df_final = df_edite

                # Nettoyage : suppression des lignes vides (si on a cliqué sur + sans écrire)
                df_final = df_final.dropna(subset=['Enseignements'])
                
                # Tri pour l'organisation
                df_final = df_final.sort_values(by=["Promotion", "Jours", "Horaire"])
                
                # Sauvegarde sur le fichier Excel source
                df_final.to_excel(NOM_FICHIER_FIXE, index=False)
                
                st.success(f"✅ Enregistrement réussi ! Le fichier contient maintenant {len(df_final)} lignes.")
                st.rerun()
                
            except Exception as e:
                st.error(f"❌ Erreur technique lors de la sauvegarde : {e}")
import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import io
import os
import pandas as pd
from datetime import datetime

# ==========================================
# CONFIGURATION ET CONSTANTES
# ==========================================
TITRE_PLATEFORME = "Plateforme de gestion des EDTs-S2-2027-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA"

DEPARTEMENTS = [
    "Département d'Électrotechnique",
    "Département d'Électronique",
    "Département d'Automatique",
    "Département de Télécommunications"
]

TYPES_DOCUMENTS = [
    "Bordereau d'envoi",
    "Procès-verbal (PV) de réunion",
    "PV de surveillance",
    "PV du Comité Pédagogique",
    "Réunion du Conseil de Discipline"
]

OPTIONS_DESTINATAIRES = [
    "Le Doyen de la faculté",
    "Le vice Doyen de la Post graduation",
    "Le vice Doyen de la graduation",
    "Autres"
]

# ==========================================
# FONCTIONS TECHNIQUES DE STRUCTURE
# ==========================================
def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    """Définit l'espacement interne (padding) des cellules d'un tableau."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('w:top', top), ('w:bottom', bottom), ('w:left', left), ('w:right', right)]:
        node = OxmlElement(m)
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def ajouter_champ_page(run, type_champ):
    """Injecte un champ de numérotation dynamique (PAGE ou NUMPAGES) dans un paragraphe Word."""
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = type_champ
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'separate')
    fldChar3 = OxmlElement('w:fldChar')
    fldChar3.set(qn('w:fldCharType'), 'end')
    
    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)
    run._r.append(fldChar3)

# ==========================================
# GÉNÉRATEUR DE BORDEREAU ISO STRICT
# ==========================================
def generer_bordereau_iso(departement, donnees):
    doc = Document()
    
    # Configuration des marges globales de la page (0.8 pouce partout)
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)
        
        # Propagation du pied de page sur toutes les pages
        section.different_first_page_header_footer = False
        
        # Structure du pied de page rectifié
        footer = section.footer
        footer_p = footer.paragraphs[0]
        
        footer_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        footer_pPr = footer_p._p.get_or_add_pPr()
        tabs = OxmlElement('w:tabs')
        
        # 1. Taquet au centre pour la référence (Centre à ~ 3.45 pouces = 4968 dxa)
        tab_centre = OxmlElement('w:tab')
        tab_centre.set(qn('w:val'), 'center')
        tab_centre.set(qn('w:pos'), '4968')
        tabs.append(tab_centre)
        
        # 2. Taquet à l'extrême droite pour les numéros de page (Extrémité à 6.9 pouces = 9936 dxa)
        tab_droite = OxmlElement('w:tab')
        tab_droite.set(qn('w:val'), 'right')
        tab_droite.set(qn('w:pos'), '9936')
        tabs.append(tab_droite)
        
        footer_pPr.append(tabs)
        
        # Premier saut vers le centre pour y écrire le code de référence
        footer_p.add_run("\t")
        r_ref_fixe = footer_p.add_run("Réf : UDL-GEL-ER-004-2027")
        r_ref_fixe.font.name = 'Calibri'
        r_ref_fixe.font.size = Pt(11)
        
        # Deuxième saut vers l'extrême droite pour y loger la pagination automatique
        footer_p.add_run("\t")
        
        r_page_actuelle = footer_p.add_run()
        r_page_actuelle.font.name = 'Calibri'
        r_page_actuelle.font.size = Pt(11)
        ajouter_champ_page(r_page_actuelle, "PAGE")
        
        r_separateur = footer_p.add_run("/")
        r_separateur.font.name = 'Calibri'
        r_separateur.font.size = Pt(11)
        
        r_total_pages = footer_p.add_run()
        r_total_pages.font.name = 'Calibri'
        r_total_pages.font.size = Pt(11)
        ajouter_champ_page(r_total_pages, "NUMPAGES")

    # 1. STRUCTURE DE L'EN-TÊTE VIA UN TABLEAU INVISIBLE
    header_table = doc.add_table(rows=1, cols=2)
    header_table.alignment = WD_ALIGN_PARAGRAPH.CENTER
    header_table.autofit = False
    
    header_table.columns[0].width = Inches(1.2)
    header_table.columns[1].width = Inches(5.7)
    
    cell_logo = header_table.rows[0].cells[0]
    cell_texte = header_table.rows[0].cells[1]
    
    tblPr = header_table._tbl.tblPr
    tblBorders = OxmlElement('w:tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'none')
        tblBorders.append(border)
    tblPr.append(tblBorders)

    # Insertion du Logo (Largeur 80 pixels = 0.833 pouces)
    p_logo = cell_logo.paragraphs[0]
    p_logo.alignment = WD_ALIGN_PARAGRAPH.LEFT
    
    nom_fichier_logo = "logo.PNG"
    if os.path.exists(nom_fichier_logo):
        p_logo.add_run().add_picture(nom_fichier_logo, width=Inches(0.833))
    else:
        r_alt = p_logo.add_run("[LOGO UNIVERSITÉ]")
        r_alt.font.name = 'Calibri'
        r_alt.font.size = Pt(8)
        r_alt.font.italic = True

    # Insertion des textes officiels de l'en-tête (Calibri)
    p_en_tete = cell_texte.paragraphs[0]
    p_en_tete.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    r1 = p_en_tete.add_run("RÉPUBLIQUE ALGÉRIENNE DÉMOCRATIQUE ET POPULAIRE\n")
    r1.bold = True
    r1.font.size = Pt(11)
    r1.font.name = 'Calibri'
    
    r2 = p_en_tete.add_run(
        "Ministère de l'Enseignement Supérieur et de la Recherche Scientifique\n"
        "Université Djillali Liabes - Sidi Bel Abbès\n"
        "Faculté de Génie Électrique\n"
    )
    r2.font.size = Pt(10)
    r2.font.name = 'Calibri'
    
    r_dept = p_en_tete.add_run(f"{departement.upper()}\n")
    r_dept.bold = True
    r_dept.font.size = Pt(11)
    r_dept.font.name = 'Calibri'

    doc.add_paragraph("\n")

    # 2. RÉFÉRENCE CHRONOLOGIQUE
    p_ref = doc.add_paragraph()
    p_ref.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r_ref = p_ref.add_run(f"N° : {donnees['num_reference']}/ F.G.E/ V.D.E.Q.L.E/2027")
    r_ref.font.size = Pt(10)
    r_ref.font.name = 'Calibri'
    r_ref.bold = True

    doc.add_paragraph("\n")

    # 3. TITRE DU BORDEREAU (Taille 36, Calibri, Italique, Souligné)
    p_titre = doc.add_paragraph()
    p_titre.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_titre = p_titre.add_run("BORDEREAU D’ENVOI")
    r_titre.font.name = 'Calibri'
    r_titre.font.size = Pt(36)
    r_titre.italic = True
    r_titre.underline = True
    r_titre.bold = True
    
    doc.add_paragraph("\n")

    # 4. DESTINATAIRE CONSTRUIT DYNAMIQUEMENT (Calibri)
    p_dest = doc.add_paragraph()
    p_dest.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r_dest = p_dest.add_run(f"A monsieur : {donnees['destinataire']}")
    r_dest.bold = True
    r_dest.font.size = Pt(12)
    r_dest.font.name = 'Calibri'

    doc.add_paragraph("\n")

    # 5. TABLEAU DE TRANSMISSION MULTI-LIGNES
    liste_pieces = donnees['liste_pieces']
    nb_lignes_totatles = 2 + len(liste_pieces)
    
    table = doc.add_table(rows=nb_lignes_totatles, cols=3)
    table.style = 'Table Grid'
    
    table.columns[0].width = Inches(4.5)
    table.columns[1].width = Inches(0.8)
    table.columns[2].width = Inches(1.7)

    # Ligne 1 : En-têtes fixes
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = "Désignation des pièces"
    hdr_cells[1].text = "Nbre"
    hdr_cells[2].text = "Observations"
    
    for cell in hdr_cells:
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        cell.paragraphs[0].runs[0].font.bold = True
        cell.paragraphs[0].runs[0].font.name = 'Calibri'
        cell.paragraphs[0].runs[0].font.size = Pt(10)
        set_cell_margins(cell, top=120, bottom=120)

    # Ligne 2 : Formule d'accompagnement
    row_joint = table.rows[1].cells
    row_joint[0].text = "Veuillez trouver ci-joint :"
    row_joint[0].paragraphs[0].runs[0].font.italic = True
    row_joint[0].paragraphs[0].runs[0].font.name = 'Calibri'
    row_joint[0].paragraphs[0].runs[0].font.size = Pt(10)
    set_cell_margins(row_joint[0], top=80, bottom=80)

    # Lignes Dynamiques
    for index, piece in enumerate(liste_pieces):
        row_idx = 2 + index
        current_row = table.rows[row_idx].cells
        
        current_row[0].text = str(piece["Désignation des pièces"])
        current_row[1].text = str(piece["Nbre"])
        current_row[2].text = str(piece["Observations"])
        
        for i, cell in enumerate(current_row):
            set_cell_margins(cell, top=150, bottom=300)
            if len(cell.paragraphs[0].runs) > 0:
                cell.paragraphs[0].runs[0].font.name = 'Calibri'
                cell.paragraphs[0].runs[0].font.size = Pt(10)
            if i == 1:
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph("\n\n")

    # 6. SIGNATURES ET ACCUSÉ DE RÉCEPTION
    p_signatures = doc.add_paragraph()
    p_signatures.alignment = WD_ALIGN_PARAGRAPH.LEFT
    date_texte = donnees['date_creation'].strftime('%d/%m/%Y')
    run_sig = p_signatures.add_run(f"Sidi bel Abbès le : {date_texte}\t\t\t\tChef de département")
    run_sig.font.name = 'Calibri'
    run_sig.font.size = Pt(11)
    run_sig.bold = True

    doc.add_paragraph("\n\n\n\n")

    p_accuse = doc.add_paragraph()
    p_accuse.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run_accuse = p_accuse.add_run("Accusé de réception    ")
    run_accuse.font.name = 'Calibri'
    run_accuse.font.size = Pt(10)
    run_accuse.font.underline = True
    run_accuse.bold = True

    return doc

def generer_pv_generique(departement, type_pv, donnees):
    """Générateur secondaire de secours (Calibri)."""
    doc = Document()
    p = doc.add_paragraph()
    run = p.add_run(f"{type_pv} - {departement}\nDocument en cours.")
    run.font.name = 'Calibri'
    return doc

# ==========================================
# INTERFACE UTILISATEUR STREAMLIT
# ==========================================
st.set_page_config(page_title="Générateur ISO Destinataire Dynamique", layout="wide")

st.caption(TITRE_PLATEFORME)
st.title("Gestion Administrative - Bordereaux & PVs")

col_dept, col_doc = st.columns(2)
with col_dept:
    dept_choisi = st.selectbox("Département émetteur :", DEPARTEMENTS)
with col_doc:
    doc_choisi = st.selectbox("Nature du document à générer :", TYPES_DOCUMENTS)

st.divider()
st.subheader(f"Formulaire d'édition - {doc_choisi}")

donnees_doc = {}

if doc_choisi == "Bordereau d'envoi":
    col_ref, col_date = st.columns(2)
    with col_ref:
        donnees_doc['num_reference'] = st.text_input("Référence séquentielle (Ex: 27)", value="27")
    with col_date:
        donnees_doc['date_creation'] = st.date_input("Date d'édition", datetime.now())
        
    # ----------------------------------------------------
    # ZONE DESTINATAIRE : SÉLECTEUR ET CHAMP LIBRE DYNAMIQUE
    # ----------------------------------------------------
    st.markdown("##### Destinataire officiel")
    choix_dest = st.selectbox(
        "Sélectionnez le destinataire dans la liste :", 
        OPTIONS_DESTINATAIRES,
        index=0
    )
    
    if choix_dest == "Autres":
        donnees_doc['destinataire'] = st.text_input("Veuillez saisir la destination personnalisée :", value="")
    else:
        donnees_doc['destinataire'] = choix_dest
        
    st.markdown("---")
    st.write("**Configuration du Tableau de Transmission**")
    
    df_initial = pd.DataFrame([
        {"Désignation des pièces": "Fiches de vœux du second semestre", "Nbre": 12, "Observations": "Pour examen"},
        {"Désignation des pièces": "Procès-verbal de délibération", "Nbre": 2, "Observations": "Pour affichage"}
    ])
    
    df_edite = st.data_editor(
        df_initial, 
        num_rows="dynamic", 
        use_container_width=True,
        column_config={
            "Désignation des pièces": st.column_config.TextColumn(width="medium", required=True),
            "Nbre": st.column_config.NumberColumn(width="small", min_value=1, required=True),
            "Observations": st.column_config.TextColumn(width="medium")
        }
    )
    donnees_doc['liste_pieces'] = df_edite.to_dict(orient="records")

else:
    with st.form("form_autres"):
        donnees_doc['date_creation'] = st.date_input("Date", datetime.now())
        donnees_doc['contenu'] = st.text_area("Contenu textuel")
        st.form_submit_button("Valider")

# Action finale de compilation
if doc_choisi == "Bordereau d'envoi":
    if st.button("Compiler et Générer le Bordereau Officiel"):
        # Blocage de sécurité si le choix "Autres" est laissé vide
        if not donnees_doc['destinataire'].strip():
            st.error("Erreur : Le champ de destination personnalisée ne peut pas être vide.")
        else:
            try:
                document_final = generer_bordereau_iso(dept_choisi, donnees_doc)
                
                output_stream = io.BytesIO()
                document_final.save(output_stream)
                output_stream.seek(0)
                
                st.success("✓ Bordereau généré avec succès avec le destinataire sélectionné.")
                
                nom_fichier_export = f"Bordereau_{dept_choisi.replace(' ', '_')}.docx"
                st.download_button(
                    label="⬇️ Télécharger le document (.docx)",
                    data=output_stream,
                    file_name=nom_fichier_export,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except Exception as error:
                st.error(f"Échec de l'opération de génération : {str(error)}")













































