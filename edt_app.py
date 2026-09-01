# =============================================================================
# IMPORTS UNIFIES
# =============================================================================
import streamlit as st
import pandas as pd
import base64
import io
import time
import re
import os
import hashlib
import zipfile
import math
import smtplib
import secrets
import mimetypes
from datetime import datetime, timedelta, timezone
from collections import defaultdict, Counter
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from pathlib import Path

try:
    from supabase import create_client, Client
except ImportError:
    create_client = None
    Client = None

try:
    from fpdf import FPDF
except ImportError:
    FPDF = None

try:
    from docx import Document
    from docx.shared import Inches, Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
except ImportError:
    Document = None

# ═══════════════════════════════════════════════════════════════════════════
# FONCTIONS UTILITAIRES : Demande EDT + Email Admin
# ═══════════════════════════════════════════════════════════════════════════

def envoyer_email_notification_admin(demande_info, fichier_bytes=None):
    """Notifie l'admin par email qu'une nouvelle demande EDT est arrivée."""
    try:
        SMTP_SERVER = 'smtp.gmail.com'
        SMTP_PORT = 587
        SMTP_USER = "chef.department.elt.fge@gmail.com"
        SMTP_PASS = "gkzs pdza yodb icvd"
        EMAIL_ADMIN = "chef.department.elt.fge@gmail.com"

        nom_ens = demande_info.get('enseignant_nom', 'Inconnu')
        email_ens = demande_info.get('enseignant_email', 'N/A')
        date_dem = demande_info.get('date_demande', datetime.now().strftime("%d/%m/%Y %H:%M"))
        nb_lignes = len(demande_info.get('lignes', []))

        msg = MIMEMultipart("alternative")
        msg["Subject"] = f"📬 Plateforme de gestion des EDTs-S2-2026 — Nouvelle demande de {nom_ens}"
        msg["From"] = SMTP_USER
        msg["To"] = EMAIL_ADMIN

        html_body = f"""<!DOCTYPE html>
<html><head><meta charset="UTF-8"></head>
<body style="font-family:'Segoe UI',Arial,sans-serif;background:#f1f5f9;margin:0;padding:20px;">
<div style="max-width:600px;margin:auto;background:white;border-radius:12px;overflow:hidden;box-shadow:0 4px 12px rgba(0,0,0,0.08);">
<div style="background:linear-gradient(135deg,#DC2626 0%,#EF4444 100%);color:white;padding:25px;text-align:center;">
<h2 style="margin:0;font-size:18px;">Plateforme de gestion des EDTs-S2-2026-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA</h2>
<p style="margin:8px 0 0 0;opacity:0.9;font-size:13px;">Nouvelle demande de mise à jour</p>
</div>
<div style="padding:30px;">
<p style="color:#334155;font-size:15px;">Salem Admin,</p>
<p style="color:#64748b;font-size:14px;">
    L'enseignant <strong style="color:#1E3A8A;">{nom_ens}</strong> vient de soumettre 
    une demande de mise à jour de son emploi du temps.
</p>
<div style="background:#fef2f2;border-left:5px solid #DC2626;padding:15px;margin:20px 0;border-radius:0 8px 8px 0;">
    <p style="margin:0 0 8px 0;color:#991b1b;font-weight:600;">📋 Récapitulatif</p>
    <table style="width:100%;font-size:13px;color:#334155;">
        <tr><td style="padding:4px 0;"><b>Enseignant :</b></td><td>{nom_ens}</td></tr>
        <tr><td style="padding:4px 0;"><b>Email :</b></td><td>{email_ens}</td></tr>
        <tr><td style="padding:4px 0;"><b>Date :</b></td><td>{date_dem}</td></tr>
        <tr><td style="padding:4px 0;"><b>Créneaux proposés :</b></td><td><strong style="color:#DC2626;">{nb_lignes}</strong> ligne(s)</td></tr>
    </table>
</div>
<div style="background:#eff6ff;border:1px solid #3b82f6;border-radius:8px;padding:15px;margin:20px 0;color:#1e40af;font-size:13px;">
    <strong>⚡ Action requise :</strong><br>
    Connectez-vous à la plateforme, rubrique <strong>"📝 Demandes de Mise à Jour EDT"</strong>.
</div>
<p style="color:#64748b;font-size:13px;"><i>Cet email est généré automatiquement. Le fichier Excel est joint.</i></p>
</div>
<div style="text-align:center;padding:20px;background:#f8fafc;font-size:12px;color:#94a3b8;">
    Faculté de Génie Electrique - Université Djillali Liabes - Sidi Bel Abbes
</div>
</div>
</body></html>"""

        msg.attach(MIMEText(html_body, "html"))

        if fichier_bytes:
            part = MIMEBase("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            part.set_payload(fichier_bytes)
            encoders.encode_base64(part)
            part.add_header(
                "Content-Disposition",
                f"attachment; filename=Demande_EDT_{nom_ens.replace(' ', '_')}.xlsx"
            )
            msg.attach(part)

        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(SMTP_USER, SMTP_PASS)
        server.send_message(msg)
        server.quit()
        return True

    except Exception as e:
        print(f"[EMAIL ADMIN] Erreur : {e}")
        return False


def generer_excel_demande_edt(donnees_lignes, nom_enseignant=""):
    """Génère un fichier Excel à partir des données de demande (Ordre garanti)."""
    try:
        df = pd.DataFrame(donnees_lignes)
        ordre_cols = ['Enseignements', 'Code', 'Enseignants', 'Horaire', 'Jours', 'Lieu', 'Promotion']
        cols_presentes = [c for c in ordre_cols if c in df.columns]
        df = df[cols_presentes]

        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Demande EDT', index=False)
            ws = writer.sheets['Demande EDT']
            
            from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
            header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
            header_font = Font(color="FFFFFF", bold=True, size=11)
            thin_border = Border(
                left=Side(style='thin'), right=Side(style='thin'),
                top=Side(style='thin'), bottom=Side(style='thin')
            )
            
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = thin_border
            
            for column in ws.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if cell.value:
                            max_length = max(max_length, len(str(cell.value)))
                    except:
                        pass
                ws.column_dimensions[column_letter].width = min(max_length + 4, 40)
        
        excel_buffer.seek(0)
        return excel_buffer

    except Exception as e:
        st.error(f"Erreur génération Excel : {e}")
        return None


def sauvegarder_demande_edt(email_prof, nom_prof, donnees_lignes, supabase_client):
    """Sauvegarde une demande EDT (Supabase ou Local)."""
    try:
        excel_buffer = generer_excel_demande_edt(donnees_lignes, nom_prof)
        if excel_buffer is None:
            return False, "❌ Erreur lors de la génération Excel", None
        
        excel_buffer.seek(0)
        fichier_bytes = excel_buffer.getvalue()

        date_now = datetime.now()
        fichier_data = {
            "enseignant_email": email_prof,
            "enseignant_nom": nom_prof,
            "lignes": donnees_lignes,
            "date_generation": date_now.strftime("%d/%m/%Y %H:%M")
        }

        supabase_ok = False
        if supabase_client:
            try:
                supabase_client.table("edt_update_requests").insert({
                    "enseignant_id": email_prof,
                    "enseignant_email": email_prof,
                    "enseignant_nom": nom_prof,
                    "fichier_data": fichier_data,
                    "statut": "En attente",
                    "date_demande": date_now.isoformat()
                }).execute()
                supabase_ok = True
            except Exception as e:
                print(f"[SUPABASE] Erreur : {e}")

        if not supabase_ok:
            if "demandes_edt_local" not in st.session_state:
                st.session_state.demandes_edt_local = []
            st.session_state.demandes_edt_local.append({
                "id": len(st.session_state.demandes_edt_local) + 1,
                "enseignant_email": email_prof,
                "enseignant_nom": nom_prof,
                "fichier_data": fichier_data,
                "statut": "En attente",
                "date_demande": date_now.strftime("%d/%m/%Y %H:%M"),
                "excel_bytes": fichier_bytes
            })

        demande_info = {
            "enseignant_nom": nom_prof,
            "enseignant_email": email_prof,
            "date_demande": date_now.strftime("%d/%m/%Y %H:%M"),
            "lignes": donnees_lignes
        }
        email_envoye = envoyer_email_notification_admin(demande_info, fichier_bytes)

        msg = "✅ Demande enregistrée en ligne." if supabase_ok else "✅ Demande enregistrée (mode local)."
        msg += " 📧 Admin notifié." if email_envoye else " (Email non transmis.)"

        return True, msg, fichier_bytes

    except Exception as e:
        return False, f"❌ Erreur : {str(e)[:150]}", None


# =============================================================================
# CONFIGURATION STREAMLIT
# =============================================================================
st.set_page_config(
    page_title="Plateforme de gestion des EDTs-S2-2026-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA",
    layout="wide",
    page_icon="🏛️",
    initial_sidebar_state="expanded"
)

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

# =============================================================================
# CONNEXION SUPABASE GLOBALE
# =============================================================================
MODE_SUPABASE = False
supabase = None

if create_client:
    try:
        SUPABASE_URL = st.secrets.get("SUPABASE_URL", "")
        SUPABASE_KEY = st.secrets.get("SUPABASE_KEY", "")
        if SUPABASE_URL and SUPABASE_KEY:
            supabase = create_client(SUPABASE_URL, SUPABASE_KEY)
            MODE_SUPABASE = True
    except Exception:
        pass

# =============================================================================
# CONSTANTES COMMUNES
# =============================================================================
_BASE_DIR = Path(__file__).parent.resolve()

FILE_ETUDIANTS = str(_BASE_DIR / "Liste des étudiants_2026-2027.xlsx")
FILE_EDT       = str(_BASE_DIR / "dataEDT-ELT-S1-2027.xlsx")
FILE_ENS       = str(_BASE_DIR / "Permanents-Vacataires-ELT2-2026-2027.xlsx")
FILE_CARTES    = str(_BASE_DIR / "Fichier_cartes.pdf")
NOM_FICHIER_FIXE = FILE_EDT
NOM_FICHIER_CONTACTS = FILE_ENS

HORAIRES_LIST = [
    "8h - 9h30", "9h30 - 11h", "11h - 12h30", "12h30 - 14h", "14h - 15h","14h - 15h30","15h - 16h", "15h30 - 17h"
]
JOURS_SEMAINE = ["Dimanche", "Lundi", "Mardi", "Mercredi", "Jeudi"]

CAUSES_ABSENCES = [
    "Non justifiee",
    "Décès dans l'ascendance, la déscendance ou la parenté",
    "Mariage de l'interessé(e)",
    "Congé de paternité ou de maternité de l'interessé(e)",
    "Mission ou convocation officielle",
    "Maladie de l'interessé(e)",
    "Autres"
]

CODE_ADMIN = "1234"
CODE_ADMIN_EDT = "doctorat2026"

# =============================================================================
# CSS PERSONNALISÉ & STYLES
# =============================================================================
st.markdown("""
<style>
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #1e3a8a 0%, #3b82f6 100%);
        min-height: 100vh;
    }
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p {
        color: white !important;
    }
    [data-testid="stSidebar"] button {
        width: 100%;
        text-align: left;
    }
    .sidebar-active {
        background: linear-gradient(90deg, #4f46e5, #6366f1) !important;
        color: white !important;
        border-radius: 8px !important;
    }
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 15px;
        border-radius: 10px;
        color: white;
    }
    #MainMenu {visibility: hidden;}
    header {visibility: hidden;}
    footer {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

# =============================================================================
# SESSION STATE NAVIGATION
# =============================================================================
if "page_active" not in st.session_state:
    st.session_state.page_active = "accueil_suivi"
if "module_sel" not in st.session_state:
    st.session_state.module_sel = "📊 Suivi d'Assiduite"

MODULES_NAVIGATION = {
    "📊 Suivi d'Assiduite": {
        "icon": "📊",
        "pages": [
            ("🏠 Accueil", "accueil_suivi"),
            ("📋 Saisir Absences", "saisir_abs"),
            ("📊 Statistiques", "stats"),
            ("📈 Rapports", "rapports"),
            ("👥 Répertoire", "repertoire"),
            ("🔔 Alertes", "alertes")
        ]
    },
    "📅 Gestion des EDTs & Admin": {
        "icon": "📅",
        "pages": [
            ("🏠 Accueil Admin", "accueil_admin"),
            ("📅 Créer EDTs", "creer_edt"),
            ("👨‍🏫 Enseignants", "gerer_ens"),
            ("🎓 Promotions", "gerer_promo"),
            ("⚙️ Paramètres", "parametres")
        ]
    },
    "🧠 EDT Intelligent": {
        "icon": "🧠",
        "pages": [
            ("🏠 Accueil IA", "accueil_ia"),
            ("📊 Affichage EDT", "affichage"),
            ("🔍 Recherche", "recherche"),
            ("📤 Import/Export", "import_export"),
            ("📊 Analytics", "analytics")
        ]
    }
}

# =============================================================================
# SIDEBAR PRINCIPALE
# =============================================================================
with st.sidebar:
    st.markdown("""
    <div style='text-align: center; padding: 20px 0;'>
        <h1 style='color: white; margin: 0; font-size: 36px;'>🏛️</h1>
        <h2 style='color: white; margin: 10px 0 0 0; font-size: 16px; font-weight: bold;'>Plateforme de gestion des EDTs-S2-2026-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA</h2>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("---")
    
    st.markdown("<p style='color: white; font-weight: bold; margin-bottom: 10px;'>📂 MODULES</p>", unsafe_allow_html=True)
    
    for module_name, module_info in MODULES_NAVIGATION.items():
        if st.button(
            f"{module_info['icon']} {module_name}",
            use_container_width=True,
            key=f"module_btn_{module_name.replace(' ', '_')}"
        ):
            st.session_state.module_sel = module_name
            st.session_state.page_active = module_info['pages'][0][1]
            st.rerun()
    
    st.markdown("---")
    st.markdown("<p style='color: white; font-weight: bold; margin-bottom: 10px;'>📄 PAGES</p>", unsafe_allow_html=True)
    
    for module_name, module_info in MODULES_NAVIGATION.items():
        if st.session_state.module_sel == module_name:
            for page_name, page_key in module_info['pages']:
                is_active_page = st.session_state.page_active == page_key
                
                if is_active_page:
                    st.markdown(f"""
                    <div style='background: linear-gradient(90deg, #4f46e5, #6366f1); color: white; padding: 10px 12px; border-radius: 8px; font-weight: 600; margin: 5px 0;'>{page_name}</div>
                    """, unsafe_allow_html=True)
                else:
                    if st.button(page_name, use_container_width=True, key=f"page_btn_{page_key}"):
                        st.session_state.page_active = page_key
                        st.rerun()
            break
    
    st.markdown("---")
    col1, col2 = st.columns(2)
    with col1:
        st.metric("Année", "2026-2027", label_visibility="collapsed")
    with col2:
        st.metric("Semestre", "S2", label_visibility="collapsed")

module_sel = st.session_state.module_sel

# =============================================================================
# FONCTIONS UTILITAIRES COMMUNES
# =============================================================================
def nettoyer_nom_enseignant(nom):
    n = str(nom).strip()
    for prefix in ["Pr ", "Dr ", "Mme ", "Mr ", "Dr. ", "Pr. ", "M. "]:
        if n.startswith(prefix):
            n = n[len(prefix):]
    return n.strip()

def extraire_nom_famille(nom_complet):
    n = nettoyer_nom_enseignant(nom_complet)
    parts = n.split()
    return parts[0].upper() if parts else ""

def mapper_promotion(promo_edt):
    p = str(promo_edt).strip().upper()
    mapping_direct = {
        "ING1": "ING1", "ING2RSE": "ING2", "ING3EI": "ING3EI", "ING3RSE": "ING3RSE",
        "ING4EI": "ING4", "ING4RSE": "ING4RSE", "ING5RSE": "ING5RSE", "L1MCIL": "L1MCIL",
        "L2ELT": "L2ELT", "L2MCIL": "MCIL2", "L3ELT": "L3ELT", "MCIL2": "MCIL2",
        "MCIL3": "MCIL3", "M1CE": "M1CE", "M1ER": "M1ER", "M1MCIL": "M1MCIL",
        "M1ME": "M1ME", "M1RE": "M1RE", "M2CE": "M2CE", "M2ER": "M2ER",
        "M2MCIL": "M2MCIL", "M2ME": "M2ME", "M2RE": "M2RE",
    }
    if p in mapping_direct:
        return mapping_direct[p]
    for key, val in mapping_direct.items():
        if key in p or p in key:
            return val
    return p

def trouver_matiere_promo(nom_ens_complet, df_edt):
    nom_fam = extraire_nom_famille(nom_ens_complet)
    if not nom_fam or df_edt.empty:
        return pd.DataFrame()
    mask = df_edt["Enseignants"].astype(str).str.upper().str.contains(re.escape(nom_fam), na=False, regex=True)
    df_filtre = df_edt[mask].copy()
    if df_filtre.empty:
        return pd.DataFrame()
    df_filtre["Promotion_Mappee"] = df_filtre["Promotion"].apply(mapper_promotion)
    return df_filtre[df_filtre["Enseignants"].astype(str).str.strip().str.lower() != "non defini"]

def trouver_toutes_promos_matiere(nom_matiere, df_edt, df_etu):
    if df_edt.empty or not nom_matiere.strip():
        return [], pd.DataFrame(), ""
    mask_matiere = df_edt["Enseignements"].astype(str).str.strip().str.lower() == nom_matiere.strip().lower()
    df_mat = df_edt[mask_matiere].copy()
    if df_mat.empty:
        return [], pd.DataFrame(), ""
    
    promos_edt = list(set([mapper_promotion(str(p).strip()) for p in df_mat["Promotion"].dropna().unique()]))
    if not promos_edt:
        return [], pd.DataFrame(), ""
    
    promos_etu_uniques = df_etu["Promotion"].dropna().astype(str).str.strip().unique()
    promos_finales = []
    
    for p_edt in promos_edt:
        p_edt_upper = p_edt.upper()
        if p_edt_upper in promos_etu_uniques:
            promos_finales.append(p_edt)
        else:
            for p_etu in promos_etu_uniques:
                if p_edt_upper in p_etu.upper() or p_etu.upper() in p_edt_upper:
                    promos_finales.append(p_etu)
                    break
    
    promos_finales = list(set(promos_finales))
    df_etudiants_tous = pd.DataFrame()
    for promo_final in promos_finales:
        df_promo = df_etu[df_etu["Promotion"].astype(str).str.strip().str.upper() == promo_final.upper()]
        df_etudiants_tous = pd.concat([df_etudiants_tous, df_promo], ignore_index=True)
    
    df_etudiants_tous = df_etudiants_tous.drop_duplicates(subset=["Nom_Complet"])
    return promos_finales, df_etudiants_tous, promos_finales[0] if promos_finales else ""

def lire_excel_robuste(chemin_ou_fichier, sheet_name=0):
    if chemin_ou_fichier is None:
        return None
    if hasattr(chemin_ou_fichier, 'seek'):
        chemin_ou_fichier.seek(0)
    for engine in ['openpyxl', 'xlrd', 'pyxlsb']:
        try:
            if hasattr(chemin_ou_fichier, 'seek'):
                chemin_ou_fichier.seek(0)
            return pd.read_excel(chemin_ou_fichier, sheet_name=sheet_name, engine=engine)
        except Exception:
            continue
    raise ValueError("❌ Lecture impossible du fichier Excel.")

# =============================================================================
# MODULE PRINCIPAL
# =============================================================================
def run_Assiduité():
    st.title("Plateforme de gestion des EDTs-S2-2026-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA")
    st.caption("Département d'Électrotechnique - Faculté de génie électrique - UDL-SBA - Année 2026-2027")

    fichiers_locaux_ok = all(os.path.exists(c) for c in [FILE_ETUDIANTS, FILE_EDT, FILE_ENS])
    df_etu, df_edt, df_ens = pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    if not fichiers_locaux_ok:
        st.warning("⚠️ Fichiers locaux manquants. Veuillez uploader les fichiers requis :")
        c1, c2, c3 = st.columns(3)
        with c1: up_etu = st.file_uploader("Liste des étudiants", type=["xlsx", "xls"], key="up_etu")
        with c2: up_edt = st.file_uploader("Données EDT", type=["xlsx", "xls"], key="up_edt")
        with c3: up_ens = st.file_uploader("Liste enseignants", type=["xlsx", "xls"], key="up_ens")
        if not all([up_etu, up_edt, up_ens]):
            return
        df_etu = lire_excel_robuste(up_etu)
        df_edt = lire_excel_robuste(up_edt)
        df_ens = lire_excel_robuste(up_ens)
    else:
        df_etu = lire_excel_robuste(FILE_ETUDIANTS)
        df_edt = lire_excel_robuste(FILE_EDT)
        df_ens = lire_excel_robuste(FILE_ENS)

    df_etu.columns = df_etu.columns.str.strip()
    df_edt.columns = df_edt.columns.str.strip()
    df_ens.columns = df_ens.columns.str.strip()

    col_nom = next((c for c in df_etu.columns if c.upper() == "NOM"), None)
    col_prenom = next((c for c in df_etu.columns if c.upper() in ["PRÉNOM", "PRENOM"]), None)
    if col_nom and col_prenom:
        df_etu["Nom_Complet"] = df_etu[col_nom].astype(str).str.strip().str.upper() + " " + df_etu[col_prenom].astype(str).str.strip().str.title()
    else:
        st.error("❌ Colonnes Nom et Prénom introuvables.")
        return

    LISTE_PROFS = []
    if "NOM" in df_ens.columns and "PRÉNOM" in df_ens.columns:
        df_ens["Nom_Complet"] = df_ens["NOM"].astype(str).str.strip().str.upper() + " " + df_ens["PRÉNOM"].astype(str).str.strip().str.title()
        LISTE_PROFS = sorted(df_ens["Nom_Complet"].dropna().unique().tolist())

    if "absences" not in st.session_state: st.session_state.absences = []

    tab1, tab2, tab3, tab4 = st.tabs(["📝 Suivi d'Assiduité", "📩 Justificatifs", "📊 Bilans & Exports", "👤 Infos Étudiant"])
    
    with tab1:
        st.header("📝 Suivi de l'Assiduité et Compteur d'Absences")
        sel_prof = st.selectbox("👤 Sélectionnez l'Enseignant :", [""] + LISTE_PROFS, key="ens_T1_main")
        
        if sel_prof:
            df_matiere = trouver_matiere_promo(sel_prof, df_edt)
            if not df_matiere.empty:
                liste_mats = sorted(df_matiere["Enseignements"].dropna().unique().tolist())
                sel_mat = st.selectbox("📚 Sélectionnez la matière :", [""] + liste_mats, key="mat_T1_main")
                
                if sel_mat:
                    promos_toutes, df_etudiants_tous, promo_c = trouver_toutes_promos_matiere(sel_mat, df_edt, df_etu)
                    if not df_etudiants_tous.empty:
                        noms_e = sorted(df_etudiants_tous["Nom_Complet"].tolist())
                        st.info(f"📍 **{len(noms_e)}** étudiants concernés — Promotion : **{promo_c}**")

                        cn1, cn2, cn3 = st.columns(3)
                        with cn1: etud_non = st.selectbox("👤 Étudiant :", [""] + noms_e, key="ne_et_t1")
                        with cn2: status_assid = st.selectbox("📊 Statut :", ["", "Absent"], key="status_t1")
                        with cn3: cause_s = st.selectbox("❓ Motif :", CAUSES_ABSENCES, key="cause_t1")

                        if etud_non and status_assid == "Absent":
                            absences_etu = [a for a in st.session_state.absences if a.get("etud_non_eligible") == etud_non and a.get("matiere") == sel_mat]
                            nb_abs_matiere = len(absences_etu)
                            color = "#ef4444" if nb_abs_matiere >= 5 else "#f59e0b" if nb_abs_matiere >= 3 else "#22c55e"
                            
                            st.markdown(f"""
                            <div style='background:{color}15; border:2px solid {color}; border-radius:12px; padding:14px; text-align:center;'>
                                <div style='font-size:28px; font-weight:800; color:{color};'>{nb_abs_matiere}</div>
                                <div style='font-size:12px; color:#475569;'>Absence(s) enregistrée(s) dans cette matière</div>
                            </div>
                            """, unsafe_allow_html=True)
                            
                            if st.button("💾 Enregistrer l'absence", use_container_width=True):
                                st.session_state.absences.append({
                                    "etud_non_eligible": etud_non,
                                    "matiere": sel_mat,
                                    "promotion": promo_c,
                                    "cause": cause_s,
                                    "date": datetime.now().strftime("%d/%m/%Y")
                                })
                                st.success(f"Absence enregistrée pour {etud_non}")
                                time.sleep(0.5)
                                st.rerun()

if __name__ == "__main__":
    run_Assiduité()
