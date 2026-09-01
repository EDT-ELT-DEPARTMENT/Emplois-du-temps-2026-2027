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
    """
    Notifie l'admin par email qu'une nouvelle demande EDT est arrivée.
    """
    try:
        import smtplib
        from email.mime.text import MIMEText
        from email.mime.multipart import MIMEMultipart
        from email.mime.base import MIMEBase
        from email import encoders

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
        msg["Subject"] = f"📬 Nouvelle demande EDT — {nom_ens}"
        msg["From"] = SMTP_USER
        msg["To"] = EMAIL_ADMIN

        html_body = f"""<!DOCTYPE html>
<html><head><meta charset="UTF-8"></head>
<body style="font-family:'Segoe UI',Arial,sans-serif;background:#f1f5f9;margin:0;padding:20px;">
<div style="max-width:600px;margin:auto;background:white;border-radius:12px;overflow:hidden;box-shadow:0 4px 12px rgba(0,0,0,0.08);">
<div style="background:linear-gradient(135deg,#DC2626 0%,#EF4444 100%);color:white;padding:25px;text-align:center;">
<h2 style="margin:0;font-size:20px;">🔔 Plateforme EDT — Nouvelle Demande</h2>
<p style="margin:8px 0 0 0;opacity:0.9;font-size:13px;">département d'Électrotechnique - FGE/UDL-SBA</p>
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
    """
    Génère un fichier Excel à partir des données de demande.
    Retourne un BytesIO.
    """
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
    """
    Sauvegarde une demande EDT. Si Supabase échoue, bascule en local.
    """
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

        if supabase_ok:
            msg = "✅ Demande enregistrée en ligne."
        else:
            msg = "✅ Demande enregistrée (mode local)."

        if email_envoye:
            msg += " 📧 Admin notifié."
        else:
            msg += " (Email non transmis.)"

        return True, msg, fichier_bytes

    except Exception as e:
        return False, f"❌ Erreur : {str(e)[:150]}", None

# =============================================================================
# CONFIGURATION STREAMLIT (UNIQUE)
# =============================================================================
st.set_page_config(
    page_title="Plateforme ELT - UDL-SBA",
    layout="wide",
    page_icon="🏛️",
    initial_sidebar_state="expanded"
)

# Masquer les éléments du menu supérieur
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
# CONNEXION SUPABASE GLOBALE (partagée)
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
CARTES_PREFIX = "Fich"
CARTES_EXTENSION = ".pdf"
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
# SESSION STATE POUR NAVIGATION INTELLIGENTE
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
        <h2 style='color: white; margin: 10px 0 0 0; font-size: 24px;'>UDL-SBA</h2>
        <p style='color: rgba(255,255,255,0.85); margin: 5px 0 0 0; font-size: 11px;'>
            Département d'Électrotechnique - FGE
        </p>
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
                    <div style='
                        background: linear-gradient(90deg, #4f46e5, #6366f1);
                        color: white;
                        padding: 10px 12px;
                        border-radius: 8px;
                        font-weight: 600;
                        margin: 5px 0;
                        text-align: left;
                    '>{page_name}</div>
                    """, unsafe_allow_html=True)
                else:
                    if st.button(
                        page_name,
                        use_container_width=True,
                        key=f"page_btn_{page_key}"
                    ):
                        st.session_state.page_active = page_key
                        st.rerun()
            break
    
    st.markdown("---")
    col1, col2 = st.columns(2)
    with col1:
        st.metric("Année", "2026-2027", label_visibility="collapsed")
    with col2:
        st.metric("Semestre", "S1", label_visibility="collapsed")
    
    st.markdown("---")
    col1, col2, col3 = st.columns(3)
    with col1:
        if st.button("🔄", use_container_width=True, help="Rafraîchir"):
            st.rerun()
    with col2:
        if st.button("🌓", use_container_width=True, help="Thème"):
            st.info("🌓 Thème switcher")
    with col3:
        if st.button("⚙️", use_container_width=True, help="Paramètres"):
            st.session_state.module_sel = "📅 Gestion des EDTs & Admin"
            st.session_state.page_active = "parametres"
            st.rerun()

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
    if not parts:
        return ""
    return parts[0].upper()

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

def detecter_colonnes_etudiant(df):
    """Détecte automatiquement les colonnes étudiantes avec tolérance maximale."""
    import unicodedata
    
    def normalize_col(name):
        if pd.isna(name):
            return ""
        s = str(name).strip().lower()
        s = unicodedata.normalize('NFKD', s).encode('ASCII', 'ignore').decode('ASCII')
        s = s.replace(' ', '').replace('-', '').replace('_', '').replace('.', '').replace('/', '')
        return s
    
    cols_norm = {normalize_col(c): c for c in df.columns}
    
    def find_col(variants):
        best_match = None
        best_score = 0
        MIN_LEN = 4
        
        for v in variants:
            v_norm = normalize_col(v)
            if not v_norm:
                continue
            
            for key, orig in cols_norm.items():
                if not key:
                    continue
                score = 0
                if v_norm == key:
                    score = 1000 + len(key)
                elif len(v_norm) >= MIN_LEN and len(key) >= MIN_LEN:
                    if v_norm in key:
                        score = 500 + len(v_norm)
                    elif key in v_norm:
                        score = 200 + len(key)
                
                if score > best_score:
                    best_score = score
                    best_match = orig
        
        return best_match
    
    mapping = {}
    mapping['nom']           = find_col(['nom', 'name', 'familyname'])
    mapping['prenom']        = find_col(['prenom', 'firstname', 'givenname'])
    mapping['email']         = find_col(['email', 'e-mail', 'mail', 'courriel', 'adressemail'])
    mapping['promotion']     = find_col(['promotion', 'promo', 'niveau', 'annee'])
    mapping['mat_bac']       = find_col(['matbac', 'matriculebac', 'nombac', 'numbac', 'matriculedebac', 'mat.bac'])
    mapping['mat_etud']      = find_col(['matetudiant', 'matriculeetudiant', 'numetudiant', 'netudiant', 'matetud', 'nometudiant', 'codeetudiant'])
    mapping['groupe']        = find_col(['groupe', 'grp', 'group', 'section'])
    mapping['sous_groupe']   = find_col(['sousgroupe', 'sousgrp', 'sg', 'subgroup'])
    mapping['date_naiss']    = find_col(['datedenaissance', 'datenaiss', 'datenaissance', 'naissance'])
    mapping['lieu_naiss']    = find_col(['lieudenaissance', 'lieunaiss', 'lieunaissance', 'birthplace'])
    mapping['admis_dette']   = find_col(['admisdette', 'admis_dette', 'endette', 'dette'])
    mapping['conge_acad']    = find_col(['congeacademique', 'conge_academique', 'congeacad'])
    
    return mapping

# =============================================================================
# MOTEUR DE RECHERCHE ET EXTRACTION DES CARTES PDF PAR MATRICULE STRICTE
# =============================================================================

def lister_fichiers_cartes():
    """Liste automatiquement Fich01.pdf, Fich02.pdf, ... FichN.pdf dans le répertoire."""
    fichiers = []
    for p in _BASE_DIR.glob(f"{CARTES_PREFIX}*{CARTES_EXTENSION}"):
        if p.is_file():
            m = re.search(r"(\d+)", p.stem)
            numero = int(m.group(1)) if m else 10**12
            fichiers.append((numero, p.name.lower(), str(p)))
    fichiers.sort(key=lambda x: (x[0], x[1]))
    return [x[2] for x in fichiers]

def _extraire_page_par_matricule_exacte(chemin_pdf, matricule_cible, df_index=None, mode_debug=False):
    """
    Parcourt un PDF page par page et NE SELECTIONNE la page QUE SI
    la matricule exacte y est présente dans le texte (avec ou sans OCR).
    """
    import unicodedata
    from io import BytesIO

    def norm_num(t):
        if not t:
            return ""
        return re.sub(r"[^0-9A-Z]", "", str(t).upper().strip())

    mat_clean = norm_num(matricule_cible)
    if not mat_clean or len(mat_clean) < 3:
        return None, None, "matricule_invalide"

    def extraire_octets_page(page_index):
        try:
            from pypdf import PdfReader, PdfWriter
            reader = PdfReader(chemin_pdf)
            if not (0 <= page_index < len(reader.pages)):
                return None
            writer = PdfWriter()
            writer.add_page(reader.pages[page_index])
            out = BytesIO()
            writer.write(out)
            return out.getvalue()
        except Exception as ex:
            if mode_debug:
                print(f"[DEBUG] extraction page: {ex}")
            return None

    # 1. Vérification par Indexation Matricule | Fichier | Page
    if df_index is not None and not df_index.empty:
        try:
            courant = os.path.basename(chemin_pdf).lower()
            for _, row in df_index.iterrows():
                idx_mat = norm_num(row.get("Matricule", row.get("Mat_Etud", row.get("Matricule_Etudiant", ""))))
                idx_page = row.get("Page", row.get("Num_Page", None))
                idx_file = str(row.get("Fichier", row.get("PDF", ""))).strip()

                fichier_ok = True
                if idx_file and idx_file.lower() not in (courant, os.path.splitext(courant)[0]):
                    fichier_ok = (os.path.splitext(idx_file.lower())[0] == os.path.splitext(courant)[0])

                if fichier_ok and idx_mat == mat_clean:
                    if str(idx_page).strip().lower() not in ("", "nan", "none"):
                        page = int(float(str(idx_page).replace(",", "."))) - 1
                        data = extraire_octets_page(page)
                        if data:
                            return data, page + 1, "matricule_index"
        except Exception as ex:
            if mode_debug:
                print(f"[DEBUG] index matricule: {ex}")

    # 2. Recherche Textuelle via PyMuPDF (fitz)
    try:
        import fitz
        doc = fitz.open(chemin_pdf)
        for i in range(len(doc)):
            raw_text = doc.load_page(i).get_text() or ""
            text_clean = norm_num(raw_text)
            
            # Recherche stricte sur la matricule exacte
            if mat_clean in text_clean:
                doc.close()
                data = extraire_octets_page(i)
                if data:
                    return data, i + 1, "matricule_fitz"
        doc.close()
    except ImportError:
        pass
    except Exception as ex:
        if mode_debug:
            print(f"[DEBUG] fitz error: {ex}")

    # 3. Fallback pdfplumber
    try:
        import pdfplumber
        with pdfplumber.open(chemin_pdf) as pdf:
            for i, page in enumerate(pdf.pages):
                raw_text = page.extract_text() or ""
                text_clean = norm_num(raw_text)
                if mat_clean in text_clean:
                    data = extraire_octets_page(i)
                    if data:
                        return data, i + 1, "matricule_plumber"
    except Exception as ex:
        if mode_debug:
            print(f"[DEBUG] pdfplumber error: {ex}")

    return None, None, "matricule_differente"


def extraire_page_etudiant_pdf(chemin_pdf, matricule_etudiant, df_index=None, mode_debug=False):
    """
    Recherche la carte exclusivement par MATRICULE dans tous les Fich01.pdf, Fich02.pdf, etc.
    """
    if isinstance(chemin_pdf, (list, tuple, set)):
        fichiers = [str(x) for x in chemin_pdf if os.path.isfile(str(x))]
    elif chemin_pdf and os.path.isfile(str(chemin_pdf)):
        fichiers = [str(chemin_pdf)]
    else:
        fichiers = lister_fichiers_cartes()

    if not fichiers:
        fichiers = lister_fichiers_cartes()

    extraire_page_etudiant_pdf.dernier_fichier_trouve = None

    for fichier in fichiers:
        data, page, methode = _extraire_page_par_matricule_exacte(
            fichier, matricule_etudiant, df_index, mode_debug
        )
        if data:
            extraire_page_etudiant_pdf.dernier_fichier_trouve = fichier
            return data, page, "matricule correspondante"

    return None, None, "matricule différente"

extraire_page_etudiant_pdf.dernier_fichier_trouve = None

# =============================================================================
# LECTURE EXCEL ROBUSTE
# =============================================================================
def lire_excel_robuste(chemin_ou_fichier, sheet_name=0):
    if chemin_ou_fichier is None:
        return None
    
    if hasattr(chemin_ou_fichier, 'seek'):
        chemin_ou_fichier.seek(0)
    
    nom = ""
    if hasattr(chemin_ou_fichier, 'name'):
        nom = chemin_ou_fichier.name.lower()
    elif isinstance(chemin_ou_fichier, str):
        nom = os.path.basename(chemin_ou_fichier).lower()
    
    engines = ['openpyxl', 'xlrd', 'pyxlsb']
    if nom.endswith('.xls') and not nom.endswith('.xlsx'):
        engines = ['xlrd', 'openpyxl', 'pyxlsb']
    
    last_err = None
    for engine in engines:
        try:
            if hasattr(chemin_ou_fichier, 'seek'):
                chemin_ou_fichier.seek(0)
            return pd.read_excel(chemin_ou_fichier, sheet_name=sheet_name, engine=engine)
        except Exception as e:
            last_err = e
            continue
            
    raise ValueError(f"❌ Format non reconnu : {last_err}")

# =============================================================================
# APPLICATION PRINCIPALE & SUIVI ASSIDUITE
# =============================================================================
def run_Assiduité():
    st.title("📊 Plateforme de gestion des EDTs & Suivi d'Assiduité des Étudiants")
    st.caption("Plateforme de gestion des EDTs-S2-2026-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA")
    
    fichiers_locaux_ok = all(os.path.exists(c) for c in [FILE_ETUDIANTS, FILE_EDT, FILE_ENS])
    df_etu, df_edt, df_ens = pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    
    if not fichiers_locaux_ok:
        st.warning("⚠️ Fichiers locaux manquants. Veuillez uploader les 3 fichiers Excel :")
        c1, c2, c3 = st.columns(3)
        with c1:
            up_etu = st.file_uploader("Liste des étudiants", type=["xlsx", "xls", "xlsb"], key="up_etu")
        with c2:
            up_edt = st.file_uploader("Données EDT", type=["xlsx", "xls", "xlsb"], key="up_edt")
        with c3:
            up_ens = st.file_uploader("Liste enseignants", type=["xlsx", "xls", "xlsb"], key="up_ens")
        
        if not all([up_etu, up_edt, up_ens]):
            st.info("📤 En attente des fichiers...")
            return
        
        try:
            df_etu = lire_excel_robuste(up_etu)
            df_etu.columns = df_etu.columns.str.strip()
            df_edt = lire_excel_robuste(up_edt)
            df_edt.columns = df_edt.columns.str.strip()
            df_ens = lire_excel_robuste(up_ens, sheet_name=0)
            df_ens.columns = df_ens.columns.str.strip()
        except Exception as e:
            st.error(f"❌ Erreur lors du chargement : {e}")
            return
    else:
        try:
            df_etu = lire_excel_robuste(FILE_ETUDIANTS)
            df_etu.columns = df_etu.columns.str.strip()
            df_edt = lire_excel_robuste(FILE_EDT)
            df_edt.columns = df_edt.columns.str.strip()
            df_ens = lire_excel_robuste(FILE_ENS, sheet_name=0)
            df_ens.columns = df_ens.columns.str.strip()
        except Exception as e:
            st.error(f"❌ Erreur de chargement local : {e}")
            return

    # Mappage des colonnes de la liste étudiant
    cols_etu_map = detecter_colonnes_etudiant(df_etu)
    col_nom = cols_etu_map.get('nom')
    col_prenom = cols_etu_map.get('prenom')

    if col_nom and col_prenom:
        df_etu["Nom_Complet"] = df_etu[col_nom].astype(str).str.strip().str.upper() + " " + df_etu[col_prenom].astype(str).str.strip().str.title()
    else:
        st.error("❌ Impossible d'identifier les colonnes Nom et Prénom dans la liste des étudiants.")
        return

    # Section Sélection & Consultation Étudiant avec Carte Stricte par Matricule
    st.subheader("🎓 Consultation Étudiant & Carte Étudiante (Vérification Stricte)")
    
    promotions_dispo = df_etu[cols_etu_map['promotion']].dropna().unique().tolist() if cols_etu_map.get('promotion') else []
    promo_sel = st.selectbox("Sélectionner la Promotion :", sorted(promotions_dispo))
    
    if promo_sel and cols_etu_map.get('promotion'):
        df_etu_promo = df_etu[df_etu[cols_etu_map['promotion']].astype(str).str.strip() == str(promo_sel)].copy()
        etudiants_liste = sorted(df_etu_promo["Nom_Complet"].unique().tolist())
        etud_sel = st.selectbox("Sélectionner l'étudiant :", etudiants_liste)

        if etud_sel:
            row_etud = df_etu_promo[df_etu_promo["Nom_Complet"] == etud_sel].iloc[0]
            
            # Récupération exacte de la matricule à partir de df_etu_promo & cols_etu_map
            col_mat = cols_etu_map.get('mat_etud') or cols_etu_map.get('mat_bac')
            matricule_exacte = str(row_etud[col_mat]).strip() if col_mat and col_mat in row_etud else ""

            col_info, col_carte = st.columns([1, 1])
            
            with col_info:
                st.markdown("### 📌 Fiche Étudiant")
                st.write(f"**Nom & Prénom :** {etud_sel}")
                st.write(f"**Matricule Source :** {matricule_exacte if matricule_exacte else 'Non renseignée'}")
                if cols_etu_map.get('email'):
                    st.write(f"**Email :** {row_etud.get(cols_etu_map['email'], 'N/A')}")
                if cols_etu_map.get('groupe'):
                    st.write(f"**Groupe :** {row_etud.get(cols_etu_map['groupe'], 'N/A')}")

            with col_carte:
                st.markdown("### 🪪 Carte Étudiante")
                if not matricule_exacte or matricule_exacte.lower() in ['nan', 'none', '']:
                    st.error("❌ Recherche impossible : aucune matricule disponible pour cet étudiant dans le fichier source.")
                else:
                    pdf_bytes, page_num, statut_verification = extraire_page_etudiant_pdf(
                        lister_fichiers_cartes(),
                        matricule_exacte
                    )

                    if statut_verification == "matricule correspondante" and pdf_bytes:
                        st.success(f"✅ {statut_verification} (Page {page_num} dans {os.path.basename(extraire_page_etudiant_pdf.dernier_fichier_trouve)})")
                        base64_pdf = base64.b64encode(pdf_bytes).decode('utf-8')
                        pdf_display = f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="100%" height="400" type="application/pdf"></iframe>'
                        st.markdown(pdf_display, unsafe_allow_html=True)
                    else:
                        st.error("❌ matricule différente")
                        st.warning("Carte refusée : Aucune page dans Fich01.pdf, Fich02.pdf, etc. ne contient cette matricule exacte. La correspondance sur le Nom/Prénom est strictement ignorée.")

# Point d'entrée Streamlit
if __name__ == "__main__":
    run_Assiduité()
