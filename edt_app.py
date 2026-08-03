"""
================================================================================
Application Unifiée : Suivi d'Assiduité + Gestion des EDTs
Département d'Electrotechnique - Faculté de Genie Electrique - UDL-SBA
Année universitaire 2026-2027
================================================================================
"""

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
from datetime import datetime, timedelta
from collections import defaultdict
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
NOM_FICHIER_FIXE = FILE_EDT
NOM_FICHIER_CONTACTS = FILE_ENS

HORAIRES_LIST = [
    "08h00 - 09h30", "09h30 - 11h00", "11h00 - 12h30",
    "12h30 - 14h00", "14h00 - 15h30", "15h30 - 17h00"
]
JOURS_SEMAINE = ["Dimanche", "Lundi", "Mardi", "Mercredi", "Jeudi"]

CAUSES_ABSENCES = [
    "Non justifie",
    "Deces dans l'ascendance, la descendance ou la parente",
    "Mariage de l'interesse(e)",
    "Conge de paternite ou de maternite de l'interesse(e)",
    "Mission ou convocation officielle",
    "Maladie de l'interesse(e)",
    "Autres"
]

CODE_ADMIN = "1234"
CODE_ADMIN_EDT = "doctorat2026"

# =============================================================================
# SIDEBAR PRINCIPALE (TOUJOURS VISIBLE)
# =============================================================================
with st.sidebar:
    st.markdown("<h2 style='text-align:center;color:#1E3A8A;'>🏛️ UDL-SBA</h2>", unsafe_allow_html=True)
    st.caption("Département d'Électrotechnique - FGE")
    st.markdown("---")
    
    module_sel = st.radio(
        "📂 Choix du module :",
        ["📊 Suivi d'Assiduité", "📅 Gestion des EDTs & Admin"],
        index=0,
        key="module_selector"
    )
    
    st.markdown("---")
    st.caption("Année universitaire 2026-2027")

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
    # Mapping direct sans doublons ni erreurs
    mapping_direct = {
        "ING1": "ING1",
        "ING2RSE": "ING2",
        "ING3EI": "ING3EI",
        "ING3RSE": "ING3RSE",
        "ING4EI": "ING4",
        "ING4RSE": "ING4RSE",
        "ING5RSE": "ING5RSE",
        "L1MCIL": "L1MCIL",
        "L2ELT": "L2ELT",
        "L2MCIL": "MCIL2",
        "L3ELT": "L3ELT",
        "MCIL2": "MCIL2",
        "MCIL3": "MCIL3",
        "M1CE": "M1CE",
        "M1ER": "M1ER",
        "M1MCIL": "M1MCIL",
        "M1ME": "M1ME",
        "M1RE": "M1RE",
        "M2CE": "M2CE",
        "M2ER": "M2ER",
        "M2MCIL": "M2MCIL",
        "M2ME": "M2ME",
        "M2RE": "M2RE",
    }
    if p in mapping_direct:
        return mapping_direct[p]
    for key, val in mapping_direct.items():
        if key in p or p in key:
            return val
    if "ING1" in p: return "ING1"
    elif "ING2" in p: return "ING2"
    elif "ING3" in p: return "ING3RSE" if "RSE" in p else "ING3EI"
    elif "ING4" in p: return "ING4"
    elif "L1" in p and "MCIL" in p: return "L1MCIL"
    elif "L2" in p and "ELT" in p: return "L2ELT"
    elif "L2" in p and "MCIL" in p: return "MCIL2"
    elif "L3" in p and "ELT" in p: return "L3ELT"
    elif "MCIL3" in p: return "MCIL3"
    elif "M1" in p:
        for code in ["CE", "ER", "MCIL", "ME", "RE"]:
            if code in p: return f"M1{code}"
    elif "M2" in p:
        for code in ["CE", "ER", "MCIL", "ME", "RE"]:
            if code in p: return f"M2{code}"
    return p

def trouver_matiere_promo(nom_ens_complet, df_edt):
    nom_fam = extraire_nom_famille(nom_ens_complet)
    if not nom_fam or df_edt.empty:
        return pd.DataFrame()
    mask = df_edt["Enseignants"].astype(str).str.upper().str.contains(
        re.escape(nom_fam), na=False, regex=True
    )
    df_filtre = df_edt[mask].copy()
    if df_filtre.empty:
        return pd.DataFrame()
    df_filtre["Promotion_Mappee"] = df_filtre["Promotion"].apply(mapper_promotion)
    df_filtre = df_filtre[df_filtre["Enseignants"].astype(str).str.strip().str.lower() != "non defini"]
    return df_filtre

def generer_page_html(df_data, titre_bilan, colonnes, entetes):
    html_doc = f"""<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <title>{titre_bilan}</title>
    <style>
        body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
               background-color: #f1f5f9; color: #1e293b; padding: 30px; margin: 0; }}
        .header {{ background-color: #1e3a8a; color: white; padding: 20px;
                   border-radius: 8px 8px 0 0; }}
        .header h1 {{ margin: 0; font-size: 1.5rem; }}
        .content {{ background: white; padding: 20px; border-radius: 0 0 8px 8px;
                    box-shadow: 0 4px 6px rgba(0,0,0,0.05); }}
        table {{ width: 100%; border-collapse: collapse; margin-top: 15px; }}
        th {{ background-color: #0f172a; color: white; padding: 12px 15px; text-align: left; }}
        td {{ padding: 12px 15px; border-bottom: 1px solid #cbd5e1; font-size: 0.9rem; }}
        tr:nth-child(even) {{ background-color: #f8fafc; }}
        .abs-count {{ color: #b91c1c; font-weight: bold; }}
        .badge-fav {{ background-color: #dcfce7; color: #166534; padding: 4px 8px;
                      border-radius: 4px; font-weight: 600; font-size: 0.85em; }}
        .badge-def {{ background-color: #fee2e2; color: #991b1b; padding: 4px 8px;
                      border-radius: 4px; font-weight: 600; font-size: 0.85em; }}
        .badge-att {{ background-color: #fef3c7; color: #92400e; padding: 4px 8px;
                      border-radius: 4px; font-weight: 600; font-size: 0.85em; }}
        .footer {{ text-align: center; margin-top: 25px; font-size: 0.8rem; color: #64748b; }}
    </style>
</head>
<body>
    <div class="header">
        <h1>📊 {titre_bilan}</h1>
        <p>Suivi d'Assiduite - Departement d'Electrotechnique - UDL-SBA</p>
    </div>
    <div class="content">
        <p>Genere le : {datetime.now().strftime('%d/%m/%Y a %H:%M')}</p>
        <table>
            <thead><tr>"""
    for h in entetes:
        html_doc += f"<th>{h}</th>"
    html_doc += "</tr></thead><tbody>"
    for _, row in df_data.iterrows():
        html_doc += "<tr>"
        for col in colonnes:
            val = row.get(col, "")
            if "Absences" in str(col) or "Total" in str(col):
                html_doc += f"<td class='abs-count'>{val}</td>"
            elif str(col).lower() == "statut":
                if "favor" in str(val).lower() and "de" not in str(val).lower():
                    html_doc += f"<td><span class='badge-fav'>{val}</span></td>"
                elif "defavor" in str(val).lower():
                    html_doc += f"<td><span class='badge-def'>{val}</span></td>"
                else:
                    html_doc += f"<td><span class='badge-att'>{val}</span></td>"
            else:
                html_doc += f"<td>{val}</td>"
        html_doc += "</tr>"
    html_doc += """</tbody></table></div>
    <div class="footer">&copy; 2026 Departement d'Electrotechnique - UDL-SBA</div>
</body>
</html>"""
    return html_doc

# =============================================================================
# MODULE 1 : SUIVI ASSIDUITE DES ETUDIANTS
# =============================================================================
# FONCTION DE LECTURE EXCEL ROBUSTE
    # =============================================================================
def lire_excel_robuste(chemin_ou_fichier, sheet_name=0):
    """Lit un fichier Excel en essayant plusieurs engines (.xlsx, .xls, .xlsb)."""
    if chemin_ou_fichier is None:
        return None
    
    if hasattr(chemin_ou_fichier, 'seek'):
        chemin_ou_fichier.seek(0)
    
    # Détection prioritaire selon l'extension
    nom = ""
    if hasattr(chemin_ou_fichier, 'name'):
        nom = chemin_ou_fichier.name.lower()
    elif isinstance(chemin_ou_fichier, str):
        nom = os.path.basename(chemin_ou_fichier).lower()
    
    # Ordre des engines : xlrd prioritaire pour les .xls anciens
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
            
    raise ValueError(f"❌ Format non reconnu. Utilisez un fichier Excel valide (.xlsx, .xls, .xlsb). Erreur : {last_err}")


def run_assiduite():
    st.title("📊 Plateforme de Suivi d'Assiduite des Etudiants")
    st.caption("Departement d'Electrotechnique - Faculte de Genie Electrique - UDL-SBA - Annee 2026-2027")
    
    # =============================================================================
    # CHARGEMENT DES DONNÉES (UNIFIÉ)
    # =============================================================================
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
        except Exception as e:
            st.error(f"❌ Erreur lecture étudiants : {e}")
            return
        try:
            df_edt = lire_excel_robuste(up_edt)
            df_edt.columns = df_edt.columns.str.strip()
        except Exception as e:
            st.error(f"❌ Erreur lecture EDT : {e}")
            return
        try:
            df_ens = lire_excel_robuste(up_ens, sheet_name=0)
            df_ens.columns = df_ens.columns.str.strip()
        except Exception as e:
            st.error(f"❌ Erreur lecture enseignants : {e}")
            return
    else:
        try:
            df_etu = lire_excel_robuste(FILE_ETUDIANTS)
            df_etu.columns = df_etu.columns.str.strip()
        except Exception as e:
            st.error(f"❌ Erreur chargement étudiants : {e}")
            return
        try:
            df_edt = lire_excel_robuste(FILE_EDT)
            df_edt.columns = df_edt.columns.str.strip()
        except Exception as e:
            st.error(f"❌ Erreur chargement EDT : {e}")
            return
        try:
            df_ens = lire_excel_robuste(FILE_ENS, sheet_name=0)
            df_ens.columns = df_ens.columns.str.strip()
        except Exception as e:
            st.error(f"❌ Erreur chargement enseignants : {e}")
            return

    # Vérification finale
    if df_etu.empty or df_edt.empty or df_ens.empty:
        st.error("❌ Données incomplètes après chargement. Vérifiez vos fichiers source.")
        return
    
    # ... suite de votre code ...    
    
    # --- Preparation des listes ---
    if "NOM" in df_ens.columns and "PRÉNOM" in df_ens.columns:
        df_ens["Nom_Complet"] = df_ens["NOM"].astype(str).str.strip().str.upper() + " " + df_ens["PRÉNOM"].astype(str).str.strip().str.title()
        LISTE_PROFS = sorted(df_ens["Nom_Complet"].dropna().unique().tolist())
    elif "Nom" in df_ens.columns and "Prénom" in df_ens.columns:
        df_ens["Nom_Complet"] = df_ens["Nom"].astype(str).str.strip().str.upper() + " " + df_ens["Prénom"].astype(str).str.strip().str.title()
        LISTE_PROFS = sorted(df_ens["Nom_Complet"].dropna().unique().tolist())
    else:
        LISTE_PROFS = []

    # Détection automatique des colonnes Nom/Prénom (insensible à la casse)
    col_nom = None
    col_prenom = None
    for c in df_etu.columns:
        if c.strip().upper() == "NOM":
            col_nom = c
        if c.strip().upper() in ["PRÉNOM", "PRENOM"]:
            col_prenom = c

    if col_nom and col_prenom:
        df_etu["Nom_Complet"] = df_etu[col_nom].astype(str).str.strip().str.upper() + " " + df_etu[col_prenom].astype(str).str.strip().str.title()
    else:
        st.error(f"❌ Colonnes 'Nom' et 'Prénom' introuvables dans le fichier étudiants. Colonnes trouvées : {list(df_etu.columns)}")
        return

    # =============================================================================
    # INITIALISATION SESSION STATE
    # =============================================================================
    if "absences" not in st.session_state:
        st.session_state.absences = []
    if "requetes" not in st.session_state:
        st.session_state.requetes = []
    if "confirm_reset" not in st.session_state:
        st.session_state.confirm_reset = False
    if "confirm_reset_abs" not in st.session_state:
        st.session_state.confirm_reset_abs = False

    # =============================================================================
    if 'etudiant_auth' not in st.session_state:
        st.session_state.etudiant_auth = None
    if 'etudiant_otp' not in st.session_state:
        st.session_state.etudiant_otp = None
    if 'etudiant_otp_email' not in st.session_state:
        st.session_state.etudiant_otp_email = None

    # FONCTIONS SUPABASE
    # =============================================================================
    def charger_absences_supabase(matiere=None, promotion=None):
        if not MODE_SUPABASE:
            return []
        try:
            query = supabase.table("suivi_assiduite_2026").select("*")
            if matiere:
                query = query.eq("matiere", matiere)
            if promotion:
                query = query.eq("promotion", promotion)
            res = query.execute()
            return res.data if res.data else []
        except Exception as e:
            st.error(f"Erreur Supabase (charger absences) : {e}")
            return []

    def enregistrer_absence_supabase(payload):
        if not MODE_SUPABASE:
            return False
        try:
            supabase.table("suivi_assiduite_2026").insert(payload).execute()
            return True
        except Exception as e:
            st.error(f"Erreur Supabase (enregistrer) : {e}")
            return False

    def supprimer_absences_supabase(matiere, promotion):
        if not MODE_SUPABASE:
            return False
        try:
            supabase.table("suivi_assiduite_2026").delete().eq("matiere", matiere).eq("promotion", promotion).execute()
            return True
        except Exception as e:
            st.error(f"Erreur Supabase (supprimer) : {e}")
            return False

    def rehabiliter_absences_etudiant_supabase(etudiant, matiere):
        if not MODE_SUPABASE:
            return False
        try:
            supabase.table("suivi_assiduite_2026").update({"justifie": True}).eq("etud_non_eligible", etudiant).eq("matiere", matiere).execute()
            return True
        except Exception as e:
            st.error(f"Erreur Supabase (rehabilitation) : {e}")
            return False

    def charger_requetes_supabase(statut=None, promotion=None):
        if not MODE_SUPABASE:
            return []
        try:
            query = supabase.table("requetes_absences").select("*")
            if statut:
                query = query.eq("statut", statut)
            if promotion:
                query = query.eq("promotion", promotion)
            res = query.execute()
            return res.data if res.data else []
        except Exception as e:
            st.error(f"Erreur Supabase (charger requetes) : {e}")
            return []

    def enregistrer_requete_supabase(payload):
        if not MODE_SUPABASE:
            return False
        try:
            supabase.table("requetes_absences").insert(payload).execute()
            return True
        except Exception as e:
            st.error(f"Erreur Supabase (requete) : {e}")
            return False

    def mettre_a_jour_statut_requete_supabase(req_id, statut):
        if not MODE_SUPABASE:
            return False
        try:
            supabase.table("requetes_absences").update({"statut": statut}).eq("id", req_id).execute()
            return True
        except Exception as e:
            st.error(f"Erreur Supabase (maj statut) : {e}")
            return False

    def reinitialiser_requetes_supabase():
        if not MODE_SUPABASE:
            return False
        try:
            supabase.table("requetes_absences").delete().neq("id", -1).execute()
            return True
        except Exception as e:
            st.error(f"Erreur Supabase (reset) : {e}")
            return False

    def get_absences_etudiant(nom_etudiant):
        if MODE_SUPABASE:
            try:
                res = supabase.table("suivi_assiduite_2026").select("*").eq("etud_non_eligible", nom_etudiant).execute()
                return res.data if res.data else []
            except Exception as e:
                st.error(f"Erreur chargement absences : {e}")
                return []
        else:
            return [a for a in st.session_state.absences if a.get("etud_non_eligible") == nom_etudiant]

    def trouver_requete_existante(nom_etudiant, matiere):
        if MODE_SUPABASE:
            try:
                res = supabase.table("requetes_absences").select("*")\
                    .eq("nom_etudiant", nom_etudiant)\
                    .eq("matiere", matiere)\
                    .eq("statut", "En attente").execute()
                return res.data[0] if res.data else None
            except Exception as e:
                st.error(f"Erreur recherche requête : {e}")
                return None
        for r in st.session_state.requetes:
            if (r.get("nom_etudiant") == nom_etudiant and 
                r.get("matiere") == matiere and 
                r.get("statut") == "En attente"):
                return r
        return None

    def supprimer_derniere_absence_supabase(etudiant, matiere, promotion):
        if not MODE_SUPABASE:
            return False
        try:
            res = supabase.table("suivi_assiduite_2026").select("*")\
                .eq("etud_non_eligible", etudiant)\
                .eq("matiere", matiere)\
                .eq("promotion", promotion)\
                .order("id", desc=True).limit(1).execute()
            if res.data:
                last_id = res.data[0]["id"]
                supabase.table("suivi_assiduite_2026").delete().eq("id", last_id).execute()
                return True
            return False
        except Exception as e:
            st.error(f"Erreur Supabase (annulation) : {e}")
            return False

    def supprimer_derniere_absence_locale(etudiant, matiere, promotion):
        candidates = [
            (idx, a) for idx, a in enumerate(st.session_state.absences)
            if a.get("etud_non_eligible") == etudiant
            and a.get("matiere") == matiere
            and a.get("promotion") == promotion
        ]
        if candidates:
            last_idx = candidates[-1][0]
            st.session_state.absences.pop(last_idx)
            return True
        return False


    # =============================================================================

    def envoyer_otp_etudiant(email_dest, nom_etud, code_otp):
        try:
            import smtplib
            from email.mime.text import MIMEText
            body = f"Bonjour {nom_etud},\n\nVotre code d'accès à la Plateforme de Suivi d'Assiduité est : {code_otp}\n\nCe code est valable 10 minutes.\n\nDépartement d'Électrotechnique - FGE/UDL-SBA"
            msg = MIMEText(body)
            msg["Subject"] = "Code d'accès - Plateforme Assiduité"
            msg["From"] = "chef.department.elt.fge@gmail.com"
            msg["To"] = str(email_dest).strip()
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login("chef.department.elt.fge@gmail.com", "gkzs pdza yodb icvd")
            server.send_message(msg)
            server.quit()
            return True
        except Exception as e:
            st.error(f"Erreur envoi email : {e}")
            return False

    # =============================================================================
    # AUTHENTIFICATION ETUDIANT (Mat. BAC + OTP)
    # =============================================================================
    # Récupération connexion enseignant (depuis Module 2 EDT)
    user = st.session_state.get("user_data")
    is_enseignant_connecte = user is not None and user.get("role") != "admin"

    etudiant_connecte = st.session_state.get("etudiant_auth") is not None

    if not is_enseignant_connecte and not etudiant_connecte:
        st.markdown("<h3 style='text-align:center;color:#1E3A8A;'>🔐 Portail Étudiant</h3>", unsafe_allow_html=True)
        st.info("Accédez à votre espace pour consulter vos absences et déposer des justificatifs.")

        mat_bac_input = st.text_input("🎓 Numéro de Matricule BAC :", key="mat_bac_auth", placeholder="Ex: 12345678")

        if mat_bac_input:
            mat_bac_clean = str(mat_bac_input).strip().upper().replace(" ", "").replace("-", "")
            df_match = pd.DataFrame()
            col_mat_bac = None
            for c in df_etu.columns:
                c_up = str(c).strip().upper().replace('.', '').replace(' ', '').replace('_', '').replace('-', '')
                if "MAT" in c_up and "BAC" in c_up:
                    col_mat_bac = c
                    break

            if col_mat_bac:
                mask_mat = df_etu[col_mat_bac].astype(str).str.strip().str.upper().str.replace(' ', '').str.replace('-', '') == mat_bac_clean
                df_match = df_etu[mask_mat]
            else:
                for c in df_etu.columns:
                    vals = df_etu[c].astype(str).str.strip().str.upper().str.replace(' ', '').str.replace('-', '')
                    if vals.eq(mat_bac_clean).any():
                        df_match = df_etu[vals == mat_bac_clean]
                        break

            if not df_match.empty:
                etud_nom = str(df_match.iloc[0]['Nom_Complet']).strip()
                etud_promo = str(df_match.iloc[0]['Promotion']).strip() if 'Promotion' in df_match.columns else ''
                st.success(f"✅ Étudiant trouvé : **{etud_nom}** ({etud_promo})")

                email_input = st.text_input("📧 Votre adresse email :", key="email_etud_auth", placeholder="ex: nom@email.com")

                if email_input and "@" in str(email_input):
                    if st.button("📧 Recevoir mon code d'accès", use_container_width=True, key="btn_otp"):
                        import random
                        otp_code = str(random.randint(100000, 999999))
                        st.session_state.etudiant_otp = otp_code
                        st.session_state.etudiant_otp_email = str(email_input).strip()

                        sent = envoyer_otp_etudiant(email_input, etud_nom, otp_code)
                        if sent:
                            st.success(f"✅ Code envoyé à : `{email_input}` — Vérifiez votre boîte mail (et les spams).")
                        else:
                            st.warning(f"⚠️ Impossible d'envoyer l'email. Votre code (mode démo) : `{otp_code}`")

                if st.session_state.get("etudiant_otp"):
                    otp_input = st.text_input("🔑 Saisissez le code reçu par email :", type="password", key="otp_input_auth")
                    if st.button("✅ Valider mon accès", use_container_width=True, key="btn_valider_otp"):
                        if otp_input == st.session_state.get("etudiant_otp"):
                            st.session_state.etudiant_auth = {
                                "mat_bac": mat_bac_clean,
                                "nom": etud_nom,
                                "email": st.session_state.etudiant_otp_email,
                                "promotion": etud_promo
                            }
                            st.session_state.etudiant_otp = None
                            st.session_state.etudiant_otp_email = None
                            st.success(f"🎓 Bienvenue {etud_nom} ! Accès autorisé...")
                            time.sleep(0.8)
                            st.rerun()
                        else:
                            st.error("❌ Code incorrect. Veuillez réessayer.")
            else:
                st.error("❌ Matricule BAC non reconnu dans la base étudiants.")

        return

    # =============================================================================
    # ONGLETS
    # =============================================================================
    # Création des onglets (toujours 3 pour éviter UnboundLocalError)
    tab1, tab2, tab3 = st.tabs(["📝 Suivi d'Assiduite", "📩 Justificatifs", "📊 Bilans & Exports"])



    # =============================================================================
    # ONGLET 1 : SUIVI D'ASSIDUITE
    # =============================================================================
    with tab1:
        st.header("📝 Suivi de l'Assiduite et Compteur d'Absences")

        sel_prof = ""
        sel_mat = ""
        promo_c = ""
        df_matiere = pd.DataFrame()

        if is_enseignant_connecte:
            sel_prof = user['nom_officiel']
            st.success(f"👤 Bienvenue **{sel_prof}** — Espace Suivi d'Assiduité")
            c1, c2 = st.columns(2)
            with c1:
                st.markdown(f"**Enseignant :** `{sel_prof}`")
            with c2:
                st.markdown("*Accès direct — Aucun code requis*")
        else:
            pwd = st.text_input("🔑 Code d'acces :", type="password", key="pwd_tab1")
            if pwd == CODE_ADMIN:
                c1, c2 = st.columns(2)
                with c1:
                    sel_prof = st.selectbox("👤 Selectionnez l'Enseignant :", [""] + LISTE_PROFS, key="ens_T1")
            elif pwd != "":
                st.error("❌ Code incorrect.")

        # SUITE COMMUNE
        if sel_prof:
            df_matiere = trouver_matiere_promo(sel_prof, df_edt)
            if not df_matiere.empty:
                liste_mats = sorted(df_matiere["Enseignements"].dropna().unique().tolist())
                with c2:
                    sel_mat = st.selectbox("📚 Selectionnez la Matiere :", [""] + liste_mats, key="mat_T1")
                
                if sel_mat:
                    info_rows = df_matiere[df_matiere["Enseignements"] == sel_mat]
                    if not info_rows.empty:
                        # 1. Promotion brute dans l'EDT
                        promo_edt_brut = str(info_rows.iloc[0]["Promotion"]).strip()
                        # 2. Mapping standard (ING2RSE → ING2, etc.)
                        promo_mapped = mapper_promotion(promo_edt_brut)
                        
                        # 3. Recherche intelligente dans le fichier ÉTUDIANTS
                        promos_etu_uniques = df_etu["Promotion"].dropna().astype(str).str.strip().unique()
                        promo_c = promo_mapped  # fallback
                        
                        # A. Correspondance exacte
                        if promo_mapped in promos_etu_uniques:
                            promo_c = promo_mapped
                        else:
                            # B. Correspondance partielle (ex: ING2 dans ING2RSE ou inverse)
                            pm_upper = promo_mapped.upper()
                            for p in promos_etu_uniques:
                                p_upper = p.upper()
                                if pm_upper == p_upper or pm_upper in p_upper or p_upper in pm_upper:
                                    promo_c = p  # On prend la valeur EXACTE du fichier étudiants
                                    break
                                 
        if sel_mat and promo_c:
            df_p = df_etu[df_etu["Promotion"].astype(str).str.strip().str.upper() == promo_c.upper()].copy()

            if not df_p.empty:
                noms_e = sorted(df_p["Nom_Complet"].tolist())
                st.info(f"📍 Promotion detectee : **{promo_c}** | **{len(noms_e)}** étudiants")

                if MODE_SUPABASE:
                    absences_filtrees = charger_absences_supabase(sel_mat, promo_c)
                else:
                    absences_filtrees = [
                        a for a in st.session_state.absences
                        if a.get("matiere") == sel_mat and a.get("promotion") == promo_c
                    ]
                df_db_full = pd.DataFrame(absences_filtrees)

                # Détection automatique de la colonne Mat. BAC
                col_mat_bac = None
                for c in df_p.columns:
                    c_up = str(c).strip().upper().replace(".", "").replace(" ", "").replace("_", "")
                    if "MAT" in c_up and "BAC" in c_up:
                        col_mat_bac = c
                        break

                st.markdown("#### 📥 Enregistrement d'une Absence")
                cn1, cn2, cn3 = st.columns(3)
                with cn1:
                    etud_non = st.selectbox("👤 Etudiant :", [""] + noms_e, key="ne_et_t1")
                    if etud_non and col_mat_bac:
                        mat_bac_val = df_p[df_p["Nom_Complet"] == etud_non][col_mat_bac]
                        if not mat_bac_val.empty:
                            mat_bac_str = str(mat_bac_val.iloc[0])
                            st.markdown(f"<div style='background:linear-gradient(90deg,#1E3A8A,#3B82F6);color:white;padding:6px 12px;border-radius:6px;font-size:13px;font-weight:600;text-align:center;margin-top:4px;'>🎓 Mat. BAC : {mat_bac_str}</div>", unsafe_allow_html=True)
                with cn2:
                    status_assid = st.selectbox("📊 Statut :", ["", "Absent"], key="status_t1")
                with cn3:
                    cause_s = st.selectbox("❓ Motif :", CAUSES_ABSENCES, key="cause_t1")

                c_d1, c_d2, c_d3 = st.columns(3)
                with c_d1:
                    date_abs = st.date_input("📅 Date :", key="date_t1")
                with c_d2:
                    jour_abs = st.selectbox("🗓️ Jour :", JOURS_SEMAINE, key="jour_t1")
                with c_d3:
                    horaire_abs = st.selectbox("🕒 Horaire :", HORAIRES_LIST, key="horaire_t1")

                # ─── COMPTEUR NUMÉRIQUE D'ABSENCES ───
                if etud_non and status_assid == "Absent":
                    if not df_db_full.empty and "etud_non_eligible" in df_db_full.columns:
                        absences_etu_matiere = df_db_full[df_db_full["etud_non_eligible"] == etud_non]
                        nb_abs_matiere = len(absences_etu_matiere)
                        nb_abs_justif = len(absences_etu_matiere[absences_etu_matiere.get("justifie") == True]) if "justifie" in absences_etu_matiere.columns else 0
                    else:
                        nb_abs_matiere = 0
                        nb_abs_justif = 0

                    if MODE_SUPABASE:
                        abs_global = [a for a in charger_absences_supabase() if a.get("etud_non_eligible") == etud_non]
                    else:
                        abs_global = [a for a in st.session_state.absences if a.get("etud_non_eligible") == etud_non]
                    nb_abs_global = len(abs_global)

                    # Afficheur numérique stylisé
                    st.markdown("<div style='margin:10px 0;'>", unsafe_allow_html=True)
                    c1, c2, c3, c4 = st.columns(4)
                    with c1:
                        color = "#ef4444" if nb_abs_matiere >= 5 else "#f59e0b" if nb_abs_matiere >= 3 else "#22c55e"
                        st.markdown(f"<div style='background:{color}15;border:2px solid {color};border-radius:12px;padding:14px;text-align:center;'><div style='font-size:28px;font-weight:800;color:{color};'>{nb_abs_matiere}</div><div style='font-size:11px;color:#64748b;font-weight:600;'>🔢 Absences matière</div></div>", unsafe_allow_html=True)
                    with c2:
                        st.markdown(f"<div style='background:#22c55e15;border:2px solid #22c55e;border-radius:12px;padding:14px;text-align:center;'><div style='font-size:28px;font-weight:800;color:#22c55e;'>{nb_abs_justif}</div><div style='font-size:11px;color:#64748b;font-weight:600;'>✅ Justifiées</div></div>", unsafe_allow_html=True)
                    with c3:
                        st.markdown(f"<div style='background:#3b82f615;border:2px solid #3b82f6;border-radius:12px;padding:14px;text-align:center;'><div style='font-size:28px;font-weight:800;color:#3b82f6;'>{nb_abs_global}</div><div style='font-size:11px;color:#64748b;font-weight:600;'>🌍 Total global</div></div>", unsafe_allow_html=True)
                    with c4:
                        reste = max(0, 5 - nb_abs_matiere)
                        color_r = "#ef4444" if reste == 0 else "#f59e0b" if reste <= 2 else "#22c55e"
                        st.markdown(f"<div style='background:{color_r}15;border:2px solid {color_r};border-radius:12px;padding:14px;text-align:center;'><div style='font-size:28px;font-weight:800;color:{color_r};'>{reste}</div><div style='font-size:11px;color:#64748b;font-weight:600;'>⏳ Avant exclusion</div></div>", unsafe_allow_html=True)
                    st.markdown("</div>", unsafe_allow_html=True)

                    if nb_abs_matiere >= 5:
                        st.error(f"🚫 **EXCLU de la matière {sel_mat}** — Seuil de 5 absences atteint.")
                    elif nb_abs_matiere == 4:
                        st.warning(f"⚠️ Attention : 4 absences dans {sel_mat}. Une prochaine absence = exclusion.")
                    else:
                        st.info(f"ℹ️ {nb_abs_matiere} absence(s) dans {sel_mat}. Seuil d'exclusion : 5.")

                # ─── BOUTONS D'ACTION ───
                col_btn1, col_btn2 = st.columns(2)
                with col_btn1:
                    if st.button("💾 ENREGISTRER L'ABSENCE", use_container_width=True, type="primary"):
                        if not etud_non:
                            st.error("❌ Veuillez selectionner un etudiant.")
                        elif status_assid != "Absent":
                            st.warning("⚠️ L'enregistrement necessite le statut 'Absent'.")
                        else:
                            payload = {
                                "enseignant": sel_prof,
                                "matiere": sel_mat,
                                "promotion": promo_c,
                                "etud_non_eligible": etud_non,
                                "cause_non_eligibilite": cause_s if cause_s else "Non justifie",
                                "date_absence": str(date_abs),
                                "jour_absence": jour_abs,
                                "horaire_absence": horaire_abs,
                                "date_saisie": datetime.now().strftime("%d/%m/%Y %H:%M"),
                                "justifie": False
                            }
                            if MODE_SUPABASE:
                                if enregistrer_absence_supabase(payload):
                                    st.success(f"✅ Absence enregistree pour {etud_non} !")
                                    time.sleep(0.5)
                                    st.rerun()
                            else:
                                payload["id"] = len(st.session_state.absences) + 1
                                st.session_state.absences.append(payload)
                                st.success(f"✅ Absence enregistree (mode local) pour {etud_non} !")
                                time.sleep(0.5)
                                st.rerun()

                with col_btn2:
                    if etud_non and st.button("🔄 ANNULER LA DERNIERE ABSENCE", use_container_width=True):
                        if MODE_SUPABASE:
                            if supprimer_derniere_absence_supabase(etud_non, sel_mat, promo_c):
                                st.success(f"✅ Dernière absence annulée pour {etud_non} dans {sel_mat} ! L'exclusion est levée si applicable.")
                                time.sleep(0.5)
                                st.rerun()
                            else:
                                st.warning("⚠️ Aucune absence à annuler pour cet étudiant dans cette matière.")
                        else:
                            if supprimer_derniere_absence_locale(etud_non, sel_mat, promo_c):
                                st.success(f"✅ Dernière absence annulée (mode local) pour {etud_non} ! L'exclusion est levée si applicable.")
                                time.sleep(0.5)
                                st.rerun()
                            else:
                                st.warning("⚠️ Aucune absence à annuler pour cet étudiant dans cette matière.")
                # LISTE GLOBALE DES ABSENCES
                st.divider()
                st.subheader("📋 Liste Globale des Absences")

                if not df_db_full.empty and "etud_non_eligible" in df_db_full.columns:
                    if "justifie" not in df_db_full.columns:
                        df_db_full["justifie"] = False
                    df_db_full["justifie"] = df_db_full["justifie"].fillna(False)
                    df_liste = df_db_full.copy()
                    df_liste["Statut Justif"] = df_liste["justifie"].apply(lambda x: "✅ Justifiée" if x else "❌ Non justifiée")
                    df_count_mat = df_liste.groupby(["etud_non_eligible", "matiere"]).size().reset_index(name="Abs Matiere")
                    df_liste = df_liste.merge(df_count_mat, on=["etud_non_eligible", "matiere"], how="left")
                    df_liste["Statut Exclusion"] = df_liste["Abs Matiere"].apply(lambda x: "🚫 EXCLU" if x >= 5 else "Eligible")

                    affichage_cols = {
                        "enseignant": "Charge de Cours",
                        "matiere": "Matiere",
                        "promotion": "Promotion",
                        "etud_non_eligible": "Etudiant",
                        "jour_absence": "Jour",
                        "date_absence": "Date",
                        "horaire_absence": "Horaire",
                        "cause_non_eligibilite": "Motif",
                        "Statut Justif": "Justification",
                        "Abs Matiere": "🔢 Nb (cette matière)",
                        "Statut Exclusion": "Statut"
                    }
                    df_aff = df_liste[list(affichage_cols.keys())].rename(columns=affichage_cols)
                    df_aff = df_aff.sort_values(by=["Etudiant", "Date"], ascending=[True, False])
                    st.dataframe(df_aff, use_container_width=True, hide_index=True)

                    if st.button("🗑️ Effacer TOUT l'historique des absences", type="primary"):
                        if MODE_SUPABASE:
                            try:
                                supabase.table("suivi_assiduite_2026").delete().neq("id", -1).execute()
                                st.success("✅ Historique Supabase effacé !")
                                time.sleep(0.5)
                                st.rerun()
                            except Exception as e:
                                st.error(f"Erreur : {e}")
                        else:
                            st.session_state.absences = []
                            st.success("✅ Historique local effacé !")
                            time.sleep(0.5)
                            st.rerun()
                else:
                    st.info("ℹ️ Aucune absence enregistrée dans l'historique global.")

                # RAPPORT OFFICIEL EXCEL
                st.divider()
                st.subheader("📥 Rapport Officiel Excel — Liste d'Éligibilité")

                toutes_promos = sorted(df_etu["Promotion"].dropna().unique().tolist())
                promo_rapport = st.selectbox(
                    "🎓 Sélectionner la promotion pour le rapport :",
                    options=toutes_promos,
                    index=toutes_promos.index(promo_c) if promo_c in toutes_promos else 0,
                    key="promo_rapport_select"
                )

                df_p_rapport = df_etu[df_etu["Promotion"].astype(str).str.strip().str.upper() == promo_rapport.upper()].copy()

                if not df_p_rapport.empty:
                    try:
                        if MODE_SUPABASE:
                            absences_cours = charger_absences_supabase(sel_mat, promo_rapport)
                        else:
                            absences_cours = [
                                a for a in st.session_state.absences
                                if a.get("matiere") == sel_mat and a.get("promotion") == promo_rapport
                            ]
                        df_abs_cours = pd.DataFrame(absences_cours)

                        absents_noms = []
                        absents_details = {}
                        if not df_abs_cours.empty and "etud_non_eligible" in df_abs_cours.columns:
                            for _, row in df_abs_cours.iterrows():
                                nom = row["etud_non_eligible"]
                                absents_noms.append(nom)
                                absents_details[nom] = {
                                    "motif": row.get("cause_non_eligibilite", "Non justifie"),
                                    "date": row.get("date_absence", ""),
                                    "jour": row.get("jour_absence", ""),
                                    "horaire": row.get("horaire_absence", "")
                                }

                        df_liste_finale = df_p_rapport.copy()
                        df_liste_finale["Statut"] = df_liste_finale["Nom_Complet"].apply(
                            lambda x: "❌ Non Eligible (Absent)" if x in absents_noms else "✅ Eligible"
                        )
                        df_liste_finale["Motif du Retrait"] = df_liste_finale["Nom_Complet"].map(
                            lambda x: absents_details.get(x, {}).get("motif", "")
                        )
                        df_liste_finale["Date Absence"] = df_liste_finale["Nom_Complet"].map(
                            lambda x: absents_details.get(x, {}).get("date", "")
                        )
                        df_liste_finale["Jour"] = df_liste_finale["Nom_Complet"].map(
                            lambda x: absents_details.get(x, {}).get("jour", "")
                        )
                        df_liste_finale["Horaire"] = df_liste_finale["Nom_Complet"].map(
                            lambda x: absents_details.get(x, {}).get("horaire", "")
                        )

                        df_export = df_liste_finale[[
                            "Nom_Complet", "Statut", "Motif du Retrait",
                            "Date Absence", "Jour", "Horaire"
                        ]].rename(columns={
                            "Nom_Complet": "Nom et Prénom",
                            "Motif du Retrait": "Motif Absence",
                            "Date Absence": "Date",
                            "Horaire": "Horaire"
                        })
                        df_export["Matiere"] = sel_mat
                        df_export["Charge"] = sel_prof
                        df_export["Promotion"] = promo_rapport

                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                            workbook = writer.book
                            fmt_title = workbook.add_format({
                                'bold': True, 'font_size': 14, 'align': 'center', 'valign': 'vcenter',
                                'bg_color': '#1E3A8A', 'font_color': 'white'
                            })
                            fmt_sub = workbook.add_format({
                                'italic': True, 'font_size': 11, 'align': 'center'
                            })
                            fmt_bold = workbook.add_format({'bold': True})
                            fmt_eligible = workbook.add_format({
                                'bg_color': '#dcfce7', 'font_color': '#166534', 'border': 1
                            })
                            fmt_non_eligible = workbook.add_format({
                                'bg_color': '#fee2e2', 'font_color': '#991b1b', 'border': 1
                            })

                            df_export.to_excel(writer, sheet_name='Liste_Eligibilite', startrow=8, index=False)
                            ws = writer.sheets['Liste_Eligibilite']

                            ws.merge_range('A1:I1', "UNIVERSITE DJILLALI LIABES - SIDI BEL ABBES", fmt_title)
                            ws.merge_range('A2:I2', "Faculte de Genie Electrique - Departement d'Electrotechnique", fmt_sub)
                            ws.merge_range('A3:I3', "LISTE D'ELIGIBILITE A L'EXAMEN", fmt_title)
                            ws.write('A5', "Matiere :", fmt_bold); ws.write('B5', sel_mat)
                            ws.write('A6', "Enseignant :", fmt_bold); ws.write('B6', sel_prof)
                            ws.write('D5', "Promotion :", fmt_bold); ws.write('E5', promo_rapport)
                            ws.write('D6', "Date export :", fmt_bold); ws.write('E6', datetime.now().strftime('%d/%m/%Y'))

                            ws.set_column('A:A', 28)
                            ws.set_column('B:B', 22)
                            ws.set_column('C:C', 28)
                            ws.set_column('D:F', 14)
                            ws.set_column('G:I', 20)

                            for row_num in range(9, 9 + len(df_export)):
                                statut_val = df_export.iloc[row_num - 9]["Statut"]
                                if "Eligible" in str(statut_val) and "Non" not in str(statut_val):
                                    ws.set_row(row_num, None, fmt_eligible)
                                else:
                                    ws.set_row(row_num, None, fmt_non_eligible)

                            ws.freeze_panes(9, 0)

                        nb_eligibles = len(df_export[df_export["Statut"].str.contains("✅")])
                        nb_non_eligibles = len(df_export[df_export["Statut"].str.contains("❌")])

                        st.success(
                            f"✅ Rapport généré pour **{promo_rapport}** : **{nb_eligibles}** éligible(s) | **{nb_non_eligibles}** signalé(s) comme absent(s)."
                        )
                        st.download_button(
                            label="📥 TELECHARGER LE RAPPORT (XLSX)",
                            data=output.getvalue(),
                            file_name=f"Rapport_{sel_mat.replace(' ', '_')}_{promo_rapport}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                    except Exception as e:
                        st.error(f"❌ Erreur Excel : {e}")
                else:
                    st.warning(f"⚠️ Aucun étudiant trouvé pour la promotion {promo_rapport}.")

    # =============================================================================
    # ONGLET 2 : GESTION DES JUSTIFICATIFS
    # =============================================================================
    if not is_enseignant_connecte:
        with tab2:
            st.header("📩 Système de Gestion des Justificatifs")
            st.caption("Dépôt étudiant et validation administration")

            if etudiant_connecte:
                # Mode étudiant connecté : accès direct au dépôt
                choix_vue = "Etudiant (Depot)"
                st.success(f"👤 Connecté en tant qu'étudiant : **{etudiant_connecte['nom']}** — Mat. BAC: {etudiant_connecte['mat_bac']}")
                if st.button("🚪 Se déconnecter", use_container_width=True):
                    st.session_state.etudiant_auth = None
                    st.rerun()
                st.divider()
            else:
                choix_vue = st.radio("Profil :", ["Etudiant (Depot)", "Administration (Decision)"], horizontal=True)
                st.divider()

            if choix_vue == "Etudiant (Depot)":
                st.subheader("📤 Soumettre une demande de réhabilitation")

                col1, col2 = st.columns(2)
                with col1:
                    if etudiant_connecte:
                        promo_sel = etudiant_connecte["promotion"]
                        st.markdown(f"**🎓 Promotion :** `{promo_sel}`")
                        df_etu_promo = df_etu[df_etu['Promotion'] == promo_sel]
                        etudiant_sel = etudiant_connecte["nom"]
                        st.markdown(f"**👤 Nom :** `{etudiant_sel}`")
                    else:
                        promo_dispo = sorted(df_etu["Promotion"].dropna().unique().tolist())
                        promo_sel = st.selectbox("Promotion :", promo_dispo, key="promo_depot")
                        df_etu_promo = df_etu[df_etu['Promotion'] == promo_sel]
                        noms_dispo = sorted(df_etu_promo["Nom_Complet"].tolist())
                        etudiant_sel = st.selectbox("Votre Nom :", noms_dispo, key="etud_depot")
                with col2:
                    st.markdown("**ℹ️ Informations**")
                    st.caption("Sélectionnez votre promotion et votre nom pour voir automatiquement vos absences signalées.")

                st.subheader("📤 Soumettre une demande de rehabilitation")

                col1, col2 = st.columns(2)
                with col1:
                    promo_dispo = sorted(df_etu["Promotion"].dropna().unique().tolist())
                    promo_sel = st.selectbox("Promotion :", promo_dispo, key="promo_depot")
                    df_etu_promo = df_etu[df_etu["Promotion"] == promo_sel]
                    noms_dispo = sorted(df_etu_promo["Nom_Complet"].tolist())
                    etudiant_sel = st.selectbox("Votre Nom :", noms_dispo, key="etud_depot")
                with col2:
                    st.markdown("**ℹ️ Informations**")
                    st.caption("Sélectionnez votre promotion et votre nom pour voir automatiquement vos absences signalées.")

                st.divider()
                st.markdown("### 📋 Mes absences signalées")

                absences_etu = get_absences_etudiant(etudiant_sel)

                if absences_etu:
                    data_display = []
                    for abs_item in absences_etu:
                        mat = abs_item.get("matiere", "")
                        req = trouver_requete_existante(etudiant_sel, mat)
                        is_justif = abs_item.get("justifie", False)
                        if is_justif:
                            statut_j = "🟢 Justifiée (acceptée)"
                        elif req:
                            statut_j = "🟡 " + req.get("statut", "En attente")
                        else:
                            statut_j = "🔴 Non déposé"
                        date_dep = req.get("date_demande", "-") if req else "-"

                        data_display.append({
                            "Matière": mat,
                            "Date d'absence": abs_item.get("date_absence", ""),
                            "Jour": abs_item.get("jour_absence", ""),
                            "Horaire": abs_item.get("horaire_absence", ""),
                            "Motif initial": abs_item.get("cause_non_eligibilite", ""),
                            "Statut justificatif": statut_j,
                            "Date dépôt": date_dep
                        })

                    df_disp = pd.DataFrame(data_display)
                    st.dataframe(df_disp, use_container_width=True, hide_index=True)

                    st.markdown("### 📎 Envoyer un justificatif pour une absence")
                    absences_sans_justif = [
                        a for a in absences_etu 
                        if not trouver_requete_existante(etudiant_sel, a.get("matiere", ""))
                    ]

                    if absences_sans_justif:
                        with st.form("form_depot_cible", clear_on_submit=True):
                            options_abs = {
                                f"{a['matiere']} — {a['date_absence']} ({a['jour_absence']} {a['horaire_absence']})": a 
                                for a in absences_sans_justif
                            }
                            sel_abs = st.selectbox("Sélectionnez l'absence concernée :", list(options_abs.keys()))
                            motif_dep = st.selectbox("Motif du justificatif :", CAUSES_ABSENCES, key="motif_dep_cible")
                            fichier_pdf_cible = st.file_uploader("Joindre le justificatif (PDF)", type=["pdf"], key="pdf_cible")
                            submit_cible = st.form_submit_button("🚀 ENVOYER LE JUSTIFICATIF")

                        if submit_cible:
                            if not fichier_pdf_cible:
                                st.error("❌ Vous devez joindre un fichier PDF.")
                            else:
                                try:
                                    pdf_bytes = fichier_pdf_cible.read()
                                    pdf_encoded = base64.b64encode(pdf_bytes).decode('utf-8')
                                    abs_conc = options_abs[sel_abs]

                                    req_ex = trouver_requete_existante(etudiant_sel, abs_conc["matiere"])

                                    if MODE_SUPABASE:
                                        if req_ex:
                                            supabase.table("requetes_absences").update({
                                                "justificatif_pdf": pdf_encoded,
                                                "motif": motif_dep,
                                                "date_demande": datetime.now().strftime("%d/%m/%Y")
                                            }).eq("id", req_ex["id"]).execute()
                                            st.success(f"✅ Justificatif ajouté à la demande existante pour **{abs_conc['matiere']}** !")
                                        else:
                                            data_insert = {
                                                "date_demande": datetime.now().strftime("%d/%m/%Y"),
                                                "nom_etudiant": etudiant_sel,
                                                "matiere": abs_conc["matiere"],
                                                "promotion": abs_conc.get("promotion", promo_sel),
                                                "motif": motif_dep,
                                                "justificatif_pdf": pdf_encoded,
                                                "statut": "En attente"
                                            }
                                            enregistrer_requete_supabase(data_insert)
                                            st.success(f"✅ Demande enregistrée pour **{etudiant_sel}** !")
                                        st.balloons()
                                        time.sleep(0.5)
                                        st.rerun()
                                    else:
                                        if req_ex:
                                            req_ex["justificatif_pdf"] = pdf_encoded
                                            req_ex["motif"] = motif_dep
                                            req_ex["date_demande"] = datetime.now().strftime("%d/%m/%Y")
                                            st.success(f"✅ Justificatif mis à jour (mode local) pour **{abs_conc['matiere']}** !")
                                        else:
                                            data_insert = {
                                                "date_demande": datetime.now().strftime("%d/%m/%Y"),
                                                "nom_etudiant": etudiant_sel,
                                                "matiere": abs_conc["matiere"],
                                                "promotion": abs_conc.get("promotion", promo_sel),
                                                "motif": motif_dep,
                                                "justificatif_pdf": pdf_encoded,
                                                "statut": "En attente",
                                                "id": len(st.session_state.requetes) + 1
                                            }
                                            st.session_state.requetes.append(data_insert)
                                            st.success(f"✅ Demande enregistrée (mode local) pour **{etudiant_sel}** !")
                                        st.balloons()
                                        time.sleep(0.5)
                                        st.rerun()
                                except Exception as e:
                                    st.error(f"❌ Erreur : {e}")
                    else:
                        st.info("✅ Toutes vos absences ont déjà un justificatif déposé ou une demande en cours de traitement.")
                else:
                    st.info("ℹ️ Aucune absence signalée pour vous actuellement.")
                    st.caption("Si vous pensez qu'il s'agit d'une erreur, contactez l'enseignant de la matière concernée.")

            else:
                pwd_admin = st.text_input("🔑 Code Admin :", type="password", key="pwd_admin")

                if pwd_admin == CODE_ADMIN:
                    st.subheader("⚖️ Dossiers en attente")

                    if MODE_SUPABASE:
                        resultats = charger_requetes_supabase(statut="En attente")
                    else:
                        resultats = [r for r in st.session_state.requetes if r.get("statut") == "En attente"]

                    if not resultats:
                        st.info("📭 Aucun dossier en attente.")
                    else:
                        for req in resultats:
                            with st.expander(f"📄 {req['nom_etudiant']} — {req['matiere']}"):
                                st.write(f"**Promotion :** {req['promotion']}")
                                st.write(f"**Motif :** {req['motif']}")
                                st.write(f"**Date :** {req['date_demande']}")

                                pdf_decoded = base64.b64decode(req['justificatif_pdf'])
                                st.download_button(
                                    label="👁️ Telecharger le PDF",
                                    data=pdf_decoded,
                                    file_name=f"Justif_{req['nom_etudiant']}_{req['matiere']}.pdf",
                                    mime="application/pdf",
                                    key=f"dl_{req['id']}"
                                )

                                col_acc, col_rej = st.columns(2)
                                if col_acc.button("✅ ACCORDER", key=f"acc_{req['id']}", use_container_width=True):
                                    if MODE_SUPABASE:
                                        mettre_a_jour_statut_requete_supabase(req["id"], "Favorable")
                                        rehabiliter_absences_etudiant_supabase(req['nom_etudiant'], req['matiere'])
                                    else:
                                        for r in st.session_state.requetes:
                                            if r["id"] == req["id"]:
                                                r["statut"] = "Favorable"
                                        for a in st.session_state.absences:
                                            if (a.get("etud_non_eligible") == req['nom_etudiant']
                                                    and a.get("matiere") == req['matiere']):
                                                a["justifie"] = True
                                                a["cause_non_eligibilite"] = "Justifiee - " + str(a.get("cause_non_eligibilite", ""))
                                    st.success(f"✔️ Justificatif de {req['nom_etudiant']} pour {req['matiere']} accepté.")
                                    time.sleep(0.5)
                                    st.rerun()

                                if col_rej.button("❌ REJETER", key=f"rej_{req['id']}", use_container_width=True):
                                    if MODE_SUPABASE:
                                        mettre_a_jour_statut_requete_supabase(req["id"], "Defavorable")
                                    else:
                                        for r in st.session_state.requetes:
                                            if r["id"] == req["id"]:
                                                r["statut"] = "Defavorable"
                                    st.warning(f"❌ Dossier de {req['nom_etudiant']} rejete.")
                                    time.sleep(0.5)
                                    st.rerun()

                    with st.expander("🛠️ Zone Maintenance"):
                        st.write("Effacer toutes les requetes de justificatifs.")
                        if st.button("🔄 REINITIALISER", type="primary"):
                            st.session_state.confirm_reset = True

                        if st.session_state.get("confirm_reset", False):
                            st.error("⚠️ Cette action est IRREVERSIBLE !")
                            c_ok, c_cancel = st.columns(2)
                            if c_ok.button("🔥 CONFIRMER", type="primary"):
                                if MODE_SUPABASE:
                                    reinitialiser_requetes_supabase()
                                else:
                                    st.session_state.requetes = []
                                st.session_state.confirm_reset = False
                                st.success("✅ Base reinitialisee.")
                                time.sleep(0.5)
                                st.rerun()
                            if c_cancel.button("❌ ANNULER"):
                                st.session_state.confirm_reset = False
                                st.info("Action annulee.")
                                time.sleep(0.5)
                                st.rerun()

                elif pwd_admin != "":
                    st.error("❌ Code incorrect.")

    # =============================================================================
    # ONGLET 3 : BILANS ET EXPORTS
    # =============================================================================
    if not is_enseignant_connecte:
        with tab3:
            st.header("📊 Registres et Bilans Agreges")

            promo_filtre = st.selectbox(
                "Filtrer par Promotion :",
                sorted(df_etu["Promotion"].dropna().unique().tolist()),
                key="promo_bilan"
            )

            if MODE_SUPABASE:
                data_hist = charger_requetes_supabase(promotion=promo_filtre)
            else:
                data_hist = [r for r in st.session_state.requetes if r.get("promotion") == promo_filtre]

            if data_hist:
                df_tab = pd.DataFrame(data_hist)

                def trouver_enseignant_par_matiere(matiere):
                    rows = df_edt[df_edt["Enseignements"] == matiere]
                    if not rows.empty:
                        return str(rows.iloc[0]["Enseignants"])
                    return "Non assigne"

                df_tab["Charge"] = df_tab["matiere"].apply(trouver_enseignant_par_matiere)

                df_tab = df_tab[["date_demande", "promotion", "Charge",
                                 "nom_etudiant", "matiere", "motif", "statut"]]
                df_tab.columns = ["Date", "Promotion", "Charge", "Etudiant", "Matiere", "Motif", "Statut"]

                st.subheader("📋 Registre General")
                st.dataframe(df_tab, use_container_width=True, hide_index=True)

                buf_xl = io.BytesIO()
                with pd.ExcelWriter(buf_xl, engine='xlsxwriter') as w:
                    df_tab.to_excel(w, index=False, sheet_name='Registre')

                c1, c2 = st.columns(2)
                with c1:
                    st.download_button(
                        "📥 EXCEL",
                        buf_xl.getvalue(),
                        f"Registre_{promo_filtre}.xlsx",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                with c2:
                    html_reg = generer_page_html(df_tab, "Registre General", df_tab.columns, df_tab.columns)
                    st.download_button(
                        "🌐 HTML",
                        html_reg,
                        f"Registre_{promo_filtre}.html",
                        "text/html",
                        use_container_width=True
                    )

                st.subheader("📚 Bilan par Etudiant et Matiere")
                df_bilan_mat = df_tab.groupby(["Etudiant", "Matiere", "Charge", "Promotion"]).size().reset_index(name="Nombre d'Absences")
                st.dataframe(df_bilan_mat, use_container_width=True, hide_index=True)

                buf_mat = io.BytesIO()
                with pd.ExcelWriter(buf_mat, engine='xlsxwriter') as w:
                    df_bilan_mat.to_excel(w, index=False, sheet_name='Absences_Matiere')

                c3, c4 = st.columns(2)
                with c3:
                    st.download_button(
                        "📥 EXCEL",
                        buf_mat.getvalue(),
                        f"Bilan_Matiere_{promo_filtre}.xlsx",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                with c4:
                    html_mat = generer_page_html(df_bilan_mat, "Bilan par Matiere", df_bilan_mat.columns, df_bilan_mat.columns)
                    st.download_button(
                        "🌐 HTML",
                        html_mat,
                        f"Bilan_Matiere_{promo_filtre}.html",
                        "text/html",
                        use_container_width=True
                    )

                st.subheader("👥 Total des Absences par Etudiant")
                df_bilan_etud = df_tab.groupby(["Etudiant", "Promotion"]).size().reset_index(name="Total Absences")
                df_bilan_etud = df_bilan_etud.sort_values(by="Total Absences", ascending=False)
                st.dataframe(df_bilan_etud, use_container_width=True, hide_index=True)

                buf_etud = io.BytesIO()
                with pd.ExcelWriter(buf_etud, engine='xlsxwriter') as w:
                    df_bilan_etud.to_excel(w, index=False, sheet_name='Total_Etudiant')

                c5, c6 = st.columns(2)
                with c5:
                    st.download_button(
                        "📥 EXCEL",
                        buf_etud.getvalue(),
                        f"Total_Etudiants_{promo_filtre}.xlsx",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                with c6:
                    html_etud = generer_page_html(df_bilan_etud, "Total par Etudiant", df_bilan_etud.columns, df_bilan_etud.columns)
                    st.download_button(
                        "🌐 HTML",
                        html_etud,
                        f"Total_Etudiants_{promo_filtre}.html",
                        "text/html",
                        use_container_width=True
                    )
            else:
                st.info(f"ℹ️ Aucun historique pour {promo_filtre}.")

            st.divider()
            st.subheader("📊 Bilan des Absences Directes (Onglet Suivi)")

            promo_abs = st.selectbox(
                "Filtrer les absences par Promotion :",
                sorted(df_etu["Promotion"].dropna().unique().tolist()),
                key="promo_abs_bilan"
            )

            if MODE_SUPABASE:
                abs_promo = charger_absences_supabase(promotion=promo_abs)
            else:
                abs_promo = [a for a in st.session_state.absences if a.get("promotion") == promo_abs]

            if abs_promo:
                df_abs = pd.DataFrame(abs_promo)
                if "justifie" not in df_abs.columns:
                    df_abs["justifie"] = False

                df_abs_count = df_abs.groupby(["etud_non_eligible", "matiere"]).agg(
                    Nombre_Absences=("etud_non_eligible", "size"),
                    Dont_Justifiees=("justifie", lambda x: (x == True).sum())
                ).reset_index()
                df_abs_count["Dont_Non_Justifiees"] = df_abs_count["Nombre_Absences"] - df_abs_count["Dont_Justifiees"]
                df_abs_count["Statut"] = df_abs_count["Nombre_Absences"].apply(lambda x: "🚫 EXCLU" if x >= 5 else "Sous seuil")
                df_abs_count = df_abs_count.sort_values(by="Nombre_Absences", ascending=False)

                st.dataframe(df_abs_count, use_container_width=True, hide_index=True)

                buf_abs = io.BytesIO()
                with pd.ExcelWriter(buf_abs, engine='xlsxwriter') as w:
                    df_abs_count.to_excel(w, index=False, sheet_name='Absences_Directes')
                st.download_button(
                    "📥 EXPORTER ABSENCES DIRECTES (XLSX)",
                    buf_abs.getvalue(),
                    f"Absences_Directes_{promo_abs}.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            else:
                st.info(f"ℹ️ Aucune absence directe enregistree pour {promo_abs}.")

# =============================================================================
# MODULE 2 : GESTION DES EDTs & ADMINISTRATION
# =============================================================================
def run_edt():
    st.markdown("<h1 class='main-title'>🏛️ DÉPARTEMENT D'ÉLECTROTECHNIQUE-FGE- UDL-SBA</h1>", unsafe_allow_html=True)
    
    # --- CONNEXION BASE DE DONNÉES ---
    try:
        URL = st.secrets["SUPABASE_URL"]
        KEY = st.secrets["SUPABASE_KEY"]
        supabase_edt = create_client(URL, KEY)
    except Exception as e:
        st.error(f"Erreur connexion Supabase : {e}")
        supabase_edt = None

    def hash_pw(password):
        return hashlib.sha256(str.encode(password)).hexdigest()

    # --- GESTION DU TEMPS ---
    now = datetime.now()
    date_str = now.strftime("%d/%m/%Y")
    jours_semaine = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]
    nom_jour_fr = jours_semaine[now.weekday()]

    # --- STYLE CSS ---
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
        </style>
    """, unsafe_allow_html=True)

    # --- CHARGEMENT DES DONNÉES ---
    df = None
    repertoire_qualites = {} 
    repertoire_grades = {} 
    repertoire_source = {}
    repertoire_noms_complets = {}
    repertoire_telephones = {}
    df_contacts = None

    def normalize(s):
        if not s or s == "Non défini": 
            return "vide"
        s = str(s).strip().lower()
        s = s.replace(" ", "").replace("-", "").replace("–", "")
        s = s.replace(":00", "").replace("h00", "h")
        return s

    if os.path.exists(NOM_FICHIER_FIXE):
        try:
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
        except Exception as e:
            st.error(f"Erreur chargement EDT : {e}")

    if os.path.exists(NOM_FICHIER_CONTACTS):
        try:
            df_contacts = pd.read_excel(NOM_FICHIER_CONTACTS)
            df_contacts.columns = [str(c).strip() for c in df_contacts.columns]
            for _, row in df_contacts.iterrows():
                nom_brut = str(row.get('NOM', '')).strip().upper()
                prénom_brut = str(row.get('PRÉNOM', '')).strip().capitalize()
                email_brut = str(row.get('Email', '')).strip()
                qualite_brute = str(row.get('Qualité', 'Non défini')).strip()
                grade_brut = str(row.get('Grade', 'N/A')).strip()
                tel_brut = str(row.get('N°/TEL', '')).strip()
                tel_nettoye = ''.join([c for c in tel_brut if c.isdigit()])
                if tel_nettoye and tel_nettoye.lower() != 'nan':
                    repertoire_telephones[nom_brut] = tel_nettoye
                if nom_brut:
                    nom_complet = f"{nom_brut} {prénom_brut}".strip()
                    if email_brut and email_brut.lower() != 'nan':
                        repertoire_source[nom_brut] = email_brut
                    repertoire_noms_complets[nom_brut] = nom_complet
                    repertoire_qualites[nom_brut] = qualite_brute
                    repertoire_grades[nom_brut] = grade_brut
        except Exception as e:
            st.warning(f"Fichier contacts : {e}")

    # --- SYSTÈME D'AUTH ---
    if "user_data" not in st.session_state:
        st.session_state["user_data"] = None

    if not st.session_state["user_data"]:
        st.info("🔒 **Module 2 (Gestion des EDTs).** Connexion requise.")
        t_conn, t_ins, t_adm = st.tabs(["🔑 Connexion", "📝 Inscription", "🛡️ Admin"])

        with t_conn:
            email_input = st.text_input("Adresse Email", key="login_email")
            pass_input = st.text_input("Mot de passe", type="password", key="login_pass")
            if st.button("Se connecter au portail", use_container_width=True):
                if supabase_edt:
                    result = supabase_edt.table("enseignants_auth").select("*").eq("email", email_input).eq("password_hash", hash_pw(pass_input)).execute()
                    if result.data:
                        st.session_state["user_data"] = result.data[0]
                        st.rerun()
                    else:
                        st.error("Email ou mot de passe incorrect.")
                else:
                    st.error("Base de données non disponible.")

        with t_ins:
            st.subheader("Créer un nouveau compte Enseignant")
            if df is not None:
                noms_possibles = sorted(df["Enseignants"].unique())
            else:
                noms_possibles = []
            new_nom = st.selectbox("Sélectionnez votre nom", noms_possibles)
            new_email = st.text_input("Votre adresse Email")
            new_pass = st.text_input("Choisissez un mot de passe", type="password")
            confirm_pass = st.text_input("Confirmez le mot de passe", type="password")
            if st.button("Créer mon compte", use_container_width=True):
                if not new_email or not new_pass:
                    st.warning("Veuillez remplir tous les champs.")
                elif new_pass != confirm_pass:
                    st.error("Les mots de passe ne correspondent pas.")
                elif supabase_edt:
                    check = supabase_edt.table("enseignants_auth").select("email").eq("email", new_email).execute()
                    if check.data:
                        st.error("Cet email est déjà utilisé.")
                    else:
                        data_ins = {
                            "nom_officiel": new_nom,
                            "email": new_email,
                            "password_hash": hash_pw(new_pass),
                            "role": "enseignant"
                        }
                        supabase_edt.table("enseignants_auth").insert(data_ins).execute()
                        st.success("✅ Compte créé avec succès !")
                        st.balloons()

        with t_adm:
            code_admin = st.text_input("Code de sécurité Administration", type="password", key="admin_code")
            if st.button("Accès Administration", use_container_width=True):
                if code_admin == CODE_ADMIN_EDT:
                    st.session_state["user_data"] = {
                        "nom_officiel": "ADMINISTRATEUR", 
                        "role": "admin",
                        "email": "admin@udl-sba.dz"
                    }
                    st.rerun()
                else:
                    st.error("Code admin incorrect.")
        
        st.warning("⚠️ Veuillez vous connecter ci-dessus pour accéder au Module 2.")
        return

    user = st.session_state.get("user_data")
    is_admin = user.get("role") == "admin"

    # --- INTERFACE PRINCIPALE APRÈS CONNEXION ---
    st.markdown(f"<div class='portal-badge'>MODE ACTIF : {'ADMINISTRATEUR' if is_admin else 'ENSEIGNANT'}</div>", unsafe_allow_html=True)


    # --- MENU DE NAVIGATION INTERNE (Main Panel) ---
    col_user, col_deco = st.columns([4, 1])
    with col_user:
        st.markdown(f"**👤 Connecté :** `{user.get('nom_officiel', 'Utilisateur')}` | Rôle : `{user.get('role', 'enseignant').upper()}`")
    with col_deco:
        if st.button("🚪 Déconnexion", use_container_width=True):
            st.session_state["user_data"] = None
            st.rerun()
    
    st.divider()
    
    if is_admin:
        options_portail = [
            "📖 Emploi du Temps", 
            "📅 Surveillances Examens", 
            "🤖 Générateur Automatique", 
            "👥 Portail Enseignants", 
            "🎓 Portail mise à jour EDT", 
            "📢 Gestion Administrative"
        ]
    else:
        options_portail = [
            "👤 Mon Espace Enseignant",
            "📅 Surveillances Examens"
        ]

    portail = st.selectbox("🚀 Sélectionner Espace", options_portail)
    
    mode_view = "Personnel"
    poste_sup = False
    
    if portail == "📖 Emploi du Temps" and is_admin:
        mode_view = st.radio("Vue Administration :", [
            "Promotion", "Enseignant", "🏢 Planning Salles", 
            "🚩 Vérificateur de conflits", "✍️ Éditeur de données"
        ], horizontal=True)
        poste_sup = st.checkbox("Poste Supérieur (Décharge 3h)")
    elif portail == "👤 Mon Espace Enseignant":
        poste_sup = st.checkbox("Poste Supérieur (Décharge 3h)", key="poste_sup_ens")

    # --- LOGIQUE PRINCIPALE SELON LE PORTAIL SÉLECTIONNÉ ---
    
    # Constantes locales pour EDT
    horaires_list = [
        "8h - 9h30", "9h30 - 11h", "11h - 12h30", 
        "12h30 - 14h", "14h - 15h30", "15h30 - 17h"
    ]
    jours_list = ["Dimanche", "Lundi", "Mardi", "Mercredi", "Jeudi"]
    map_h = {normalize(h): h for h in horaires_list}
    map_j = {normalize(j): j for j in jours_list}

    # En-tête harmonisé
    col_logo, col_titre, col_date = st.columns([1, 5, 1.2])
    with col_logo:
        try:
            st.image(str(_BASE_DIR / "logo.PNG"), width=90)
        except:
            st.markdown("🏛️")
    with col_titre:
        st.markdown("<h3 style='color:#1E3A8A;margin:0;'>Plateforme de gestion des EDTs 2026-2027</h3>", unsafe_allow_html=True)
        st.caption("Département d'Électrotechnique - Faculté de Génie Électrique - UDL-SBA")
    with col_date:
        st.markdown(f"<div style='background:#1E3A8A;color:white;padding:8px 12px;border-radius:8px;text-align:center;font-size:12px;'>📅 {nom_jour_fr}<br>{date_str}</div>", unsafe_allow_html=True)

    st.markdown("<hr style='border:2px solid #D4AF37;margin:15px 0;'>", unsafe_allow_html=True)

    if df is None or df.empty:
        st.error("❌ Les données EDT ne sont pas disponibles. Vérifiez le fichier source.")
        return

    # ============================================================
    # PORTAIL : EMPLOI DU TEMPS (ADMIN)
    # ============================================================
    if portail == "📖 Emploi du Temps" and is_admin:
        if mode_view == "Enseignant":
            cible = st.selectbox("Sélectionner l'Enseignant :", 
                                sorted([e for e in df["Enseignants"].unique() if e and e != "Non défini"]))
            
            df_f = df[df["Enseignants"].str.contains(cible, case=False, na=False)].copy()
            df_f['Type'] = df_f['Code'].apply(lambda x: "COURS" if "COURS" in str(x).upper() else ("TD" if "TD" in str(x).upper() else "TP"))
            df_u = df_f.drop_duplicates(subset=['j_norm', 'h_norm'])

            nb_cours = len(df_u[df_u['Type'] == 'COURS'])
            nb_td = len(df_u[df_u['Type'] == 'TD'])
            nb_tp = len(df_u[df_u['Type'] == 'TP'])
            seuil = 3.0 if poste_sup else 6.0
            charge_eq = (nb_cours * 1.5) + (nb_td + nb_tp)
            delta = charge_eq - seuil
            h_sup = delta * 1.5

            st.markdown(f"### 📊 Charge Horaire : {cible}")
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("📘 Cours", nb_cours)
            c2.metric("📗 TD", nb_td)
            c3.metric("🔴 TP", nb_tp)
            c4.metric("Charge Eq/h", f"{charge_eq:.1f}")

            if h_sup > 0:
                st.success(f"✅ Heures supplémentaires : +{h_sup:.1f}h")
            elif h_sup < 0:
                st.warning(f"⚠️ Déficit : {h_sup:.1f}h")
            else:
                st.info("⚖️ Seuil exact")

            # Grille EDT
            def format_case(rows):
                items = []
                for _, r in rows.iterrows():
                    code_up = str(r['Code']).upper()
                    if 'COURS' in code_up:
                        nat = '📘'
                    elif 'TD' in code_up:
                        nat = '📗'
                    else:
                        nat = '🔴'
                    items.append(f"<b>{nat} {r['Enseignements']}</b><br><small>{r['Lieu']} | {r['Promotion']}</small>")
                return "<hr style='margin:4px 0;'>".join(items)

            if not df_f.empty:
                grid = df_f.groupby(['h_norm', 'j_norm']).apply(format_case, include_groups=False).unstack('j_norm')
                grid = grid.reindex(index=[normalize(h) for h in horaires_list], columns=[normalize(j) for j in jours_list]).fillna("")
                grid.index = [map_h.get(i, i) for i in grid.index]
                grid.columns = [map_j.get(c, c) for c in grid.columns]
                st.write(grid.to_html(escape=False), unsafe_allow_html=True)

        elif mode_view == "Promotion":
            p_sel = st.selectbox("Choisir Promotion :", sorted(df["Promotion"].unique()))
            df_p = df[df["Promotion"] == p_sel].copy()
            
            st.markdown(f"### 📚 EDT Promotion : {p_sel}")
            
            def fmt_p(rows):
                items = []
                for _, r in rows.iterrows():
                    code_up = str(r['Code']).upper()
                    color = '#1e40af' if 'COURS' in code_up else ('#166534' if 'TD' in code_up else '#991b1b')
                    nat = '📘' if 'COURS' in code_up else ('📗' if 'TD' in code_up else '🔴')
                    items.append(f"<div style='border-left:3px solid {color};padding:4px;margin:2px 0;background:#f8fafc;'><b>{nat} {r['Enseignements']}</b><br><small>👤 {r['Enseignants']} | 📍 {r['Lieu']}</small></div>")
                return "".join(items)

            grid_p = df_p.groupby(['h_norm', 'j_norm']).apply(fmt_p, include_groups=False).unstack('j_norm')
            grid_p = grid_p.reindex(index=[normalize(h) for h in horaires_list], columns=[normalize(j) for j in jours_list]).fillna("")
            grid_p = grid_p[grid_p.any(axis=1)]
            grid_p.index = [map_h.get(i, i) for i in grid_p.index]
            grid_p.columns = [map_j.get(c, c) for c in grid_p.columns]
            st.write(grid_p.to_html(escape=False), unsafe_allow_html=True)

            # Export
            c1, c2 = st.columns(2)
            buf_p = io.BytesIO()
            df_p[['Enseignements', 'Code', 'Enseignants', 'Horaire', 'Jours', 'Lieu', 'Promotion']].to_excel(buf_p, index=False)
            c1.download_button("📥 Excel", buf_p.getvalue(), f"EDT_{p_sel}.xlsx", 
                              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            c2.download_button("🌐 HTML", grid_p.to_html(escape=False), f"EDT_{p_sel}.html", "text/html")

        elif mode_view == "🏢 Planning Salles":
            s_sel = st.selectbox("Choisir Salle :", sorted([s for s in df["Lieu"].unique() if s and s != "Non défini"]))
            df_s = df[df["Lieu"] == s_sel]
            st.markdown(f"### 🏢 Planning : {s_sel}")
            st.dataframe(df_s[['Jours', 'Horaire', 'Enseignements', 'Enseignants', 'Promotion']], use_container_width=True, hide_index=True)

        elif mode_view == "🚩 Vérificateur de conflits":
            st.subheader("🚩 Détection des Conflits")
            
            conflits = []
            # Conflits salle
            grp_salle = df[(df["Lieu"] != "Non défini")].groupby(['Jours', 'Horaire', 'Lieu'])
            for (j, h, l), g in grp_salle:
                if len(g) > 1:
                    conflits.append({"Type": "Salle", "Jour": j, "Horaire": h, "Lieu": l, "Détail": f"{len(g)} cours simultanés"})
            # Conflits prof
            grp_prof = df[(df["Enseignants"] != "Non défini")].groupby(['Jours', 'Horaire', 'Enseignants'])
            for (j, h, p), g in grp_prof:
                if len(g) > 1:
                    conflits.append({"Type": "Enseignant", "Jour": j, "Horaire": h, "Enseignant": p, "Détail": f"{len(g)} affectations"})

            if conflits:
                st.warning(f"⚠️ {len(conflits)} conflit(s) détecté(s)")
                st.dataframe(pd.DataFrame(conflits), use_container_width=True, hide_index=True)
            else:
                st.success("✅ Aucun conflit détecté")
                st.balloons()

        elif mode_view == "✍️ Éditeur de données":
            st.subheader("✍️ Éditeur de données EDT")
            
            cols_ed = ['Enseignements', 'Code', 'Enseignants', 'Horaire', 'Jours', 'Lieu', 'Promotion']
            for c in cols_ed:
                if c not in df.columns:
                    df[c] = ""

            if 'df_admin' not in st.session_state:
                st.session_state.df_admin = df[cols_ed].copy()

            search = st.text_input("🔍 Rechercher (Enseignant, Salle, Matière) :")
            df_edit = st.session_state.df_admin.copy()
            if search:
                mask = df_edit.apply(lambda r: r.astype(str).str.contains(search, case=False).any(), axis=1)
                df_edit = df_edit[mask]

            edited = st.data_editor(df_edit, use_container_width=True, num_rows="dynamic", key="edt_editor")

            c1, c2 = st.columns(2)
            if c1.button("💾 Sauvegarder", use_container_width=True):
                st.session_state.df_admin = edited
                try:
                    edited.to_excel(NOM_FICHIER_FIXE, index=False)
                    st.success("✅ Sauvegardé !")
                except Exception as e:
                    st.error(f"Erreur sauvegarde : {e}")
            if c2.button("🔄 Réinitialiser", use_container_width=True):
                if 'df_admin' in st.session_state:
                    del st.session_state.df_admin
                st.rerun()

    # ============================================================
    # PORTAIL : MON ESPACE ENSEIGNANT
    # ============================================================
    elif portail == "👤 Mon Espace Enseignant":
        cible = user['nom_officiel']
        nom_aff = repertoire_noms_complets.get(cible.strip().upper(), cible)
        
        # ═══════════════════════════════════════════════════════
        # EN-TÊTE IDENTITÉ + DÉCONNEXION
        # ═══════════════════════════════════════════════════════
        col_id, col_deco = st.columns([4, 1])
        with col_id:
            st.markdown(f"""
                <div style="background: linear-gradient(135deg, #1E3A8A, #3B82F6); padding: 20px; border-radius: 12px; color: white; margin-bottom: 20px;">
                    <h2 style="margin:0;">👤 {nom_aff}</h2>
                    <p style="margin:5px 0 0 0; opacity:0.9;">Espace Personnel Enseignant - S1 2026-2027</p>
                </div>
            """, unsafe_allow_html=True)
        with col_deco:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("🚪 Déconnexion", use_container_width=True, type="primary", key="deco_ens_indiv_main"):
                st.session_state["user_data"] = None
                st.rerun()

        df_f = df[df["Enseignants"].str.contains(cible, case=False, na=False)].copy()
        if df_f.empty:
            st.warning("⚠️ Aucun cours programmé pour vous.")
            return

        df_f['Type'] = df_f['Code'].apply(lambda x: "COURS" if "COURS" in str(x).upper() else ("TD" if "TD" in str(x).upper() else "TP"))
        df_u = df_f.drop_duplicates(subset=['j_norm', 'h_norm'])
        
        nb_cours = len(df_u[df_u['Type'] == 'COURS'])
        nb_td = len(df_u[df_u['Type'] == 'TD'])
        nb_tp = len(df_u[df_u['Type'] == 'TP'])
        seuil = 3.0 if poste_sup else 6.0
        charge_eq = (nb_cours * 1.5) + (nb_td + nb_tp)
        delta = charge_eq - seuil

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("📘 Cours", nb_cours)
        c2.metric("📗 TD", nb_td)
        c3.metric("🔴 TP", nb_tp)
        c4.metric("Équivalent", f"{charge_eq:.1f} eq/h")

        if delta > 0:
            st.success(f"✅ Heures supplémentaires : +{delta * 1.5:.1f}h")
        elif delta < 0:
            st.warning(f"⚠️ Déficit horaire : {delta * 1.5:.1f}h")
        else:
            st.info("⚖️ Seuil réglementaire atteint")

        st.divider()
        st.markdown("### 📅 Mon Emploi du Temps")
        st.dataframe(df_f[['Jours', 'Horaire', 'Enseignements', 'Code', 'Lieu', 'Promotion']].sort_values(['Jours', 'Horaire']), 
                    use_container_width=True, hide_index=True)

        # Export perso
        col_ex1, col_ex2 = st.columns(2)
        buf_ex = io.BytesIO()
        df_f[['Enseignements', 'Code', 'Horaire', 'Jours', 'Lieu', 'Promotion']].to_excel(buf_ex, index=False)
        col_ex1.download_button("📊 Excel", buf_ex.getvalue(), f"Mon_EDT_{cible}.xlsx", 
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        col_ex2.download_button("🌐 HTML", df_f.to_html(index=False), f"Mon_EDT_{cible}.html", "text/html")

    # ============================================================
    # PORTAIL : SURVEILLANCES EXAMENS
    # ============================================================
    elif portail == "📅 Surveillances Examens":
        FILE_S = str(_BASE_DIR / "surveillances_2027.xlsx")
        if not os.path.exists(FILE_S):
            st.error("❌ Fichier 'surveillances_2027.xlsx' introuvable.")
            return

        df_surv = pd.read_excel(FILE_S)
        df_surv.columns = [str(c).strip() for c in df_surv.columns]
        
        c_prof = 'Surveillant(s)' if 'Surveillant(s)' in df_surv.columns else 'Enseignants'
        u_nom = user['nom_officiel']
        
        if is_admin:
            profs_surv = sorted([p for p in df_surv[c_prof].unique() if p and p != "Non défini"])
            prof_sel = st.selectbox("🔍 Filtrer par enseignant :", profs_surv)
        else:
            prof_sel = u_nom
            st.info(f"👤 Vos surveillances : **{u_nom}**")

        df_u = df_surv[df_surv[c_prof].str.contains(prof_sel, case=False, na=False)]
        st.markdown(f"### 📋 Planning de surveillance : {prof_sel}")
        st.dataframe(df_u, use_container_width=True, hide_index=True)

        if not df_u.empty:
            buf_s = io.BytesIO()
            df_u.to_excel(buf_s, index=False)
            st.download_button("📥 Télécharger", buf_s.getvalue(), f"Surv_{prof_sel}.xlsx",
                              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # ============================================================
    # PORTAIL : GÉNÉRATEUR AUTOMATIQUE (ADMIN)
    # ============================================================
    elif portail == "🤖 Générateur Automatique":
        if not is_admin:
            st.error("🚫 Accès réservé à l'administration.")
            return
        
        st.header("⚙️ Générateur de Surveillances")
        st.info("Cet outil génère automatiquement les plannings de surveillance d'examens.")
        
        # Interface simplifiée
        st.markdown("### 🎓 Promotions à traiter")
        promos_dispo = sorted(df['Promotion'].unique().tolist())
        promos_sel = st.multiselect("Sélectionner :", promos_dispo)
        
        if promos_sel and st.button("🚀 Générer le planning", use_container_width=True):
            st.success(f"✅ Planning généré pour : {', '.join(promos_sel)}")
            st.info("(Simulation - Intégrez votre algorithme de répartition ici)")
            st.balloons()

    # ============================================================
    # PORTAIL : PORTAIL ENSEIGNANTS (ADMIN - ENVOI MAIL)
    # ============================================================
    elif portail == "👥 Portail Enseignants":
        if not is_admin:
            st.error("🚫 Accès réservé à l'administration.")
            return

        st.header("📧 Portail Enseignants - Envoi des EDT")
        
        # Liste des enseignants avec emails
        donnees_envoi = []
        for ens in sorted(df["Enseignants"].unique()):
            if ens and ens != "Non défini":
                email = repertoire_source.get(str(ens).strip().upper(), "Non communiqué")
                donnees_envoi.append({"Enseignant": ens, "Email": email, "Statut": "✅ Prêt" if "@" in str(email) else "❌ Sans email"})

        df_envoi = pd.DataFrame(donnees_envoi)
        st.dataframe(df_envoi, use_container_width=True, hide_index=True)

        # Filtre
        filtre_statut = st.selectbox("Filtrer :", ["Tous", "✅ Prêt", "❌ Sans email"])
        if filtre_statut != "Tous":
            df_envoi = df_envoi[df_envoi["Statut"] == filtre_statut]

        st.download_button("📥 Télécharger la liste", df_envoi.to_csv(index=False), "liste_enseignants.csv", "text/csv")

        st.markdown("---")
        st.info("💡 Pour l'envoi automatisé par email, configurez les identifiants SMTP dans les secrets Streamlit.")

    # ============================================================
    # PORTAIL : MISE À JOUR EDT (ADMIN)
    # ============================================================
    elif portail == "🎓 Portail mise à jour EDT":
        st.subheader("📚 Mise à jour des Emplois du Temps")
        
        promo_vue = st.selectbox("Promotion à consulter :", sorted(df["Promotion"].unique()))
        df_vue = df[df["Promotion"] == promo_vue][['Enseignements', 'Code', 'Enseignants', 'Horaire', 'Jours', 'Lieu', 'Promotion']]
        
        st.dataframe(df_vue.sort_values(['Jours', 'Horaire']), use_container_width=True, hide_index=True)
        
        c1, c2 = st.columns(2)
        buf_v = io.BytesIO()
        df_vue.to_excel(buf_v, index=False)
        c1.download_button("📊 Excel", buf_v.getvalue(), f"EDT_{promo_vue}.xlsx", 
                          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        c2.download_button("🌐 HTML", df_vue.to_html(index=False), f"EDT_{promo_vue}.html", "text/html")

        if is_admin:
            st.divider()
            st.markdown("### ✍️ Import / Mise à jour par fichier")
            fichier_import = st.file_uploader("Importer un fichier Excel EDT", type=["xlsx"])
            if fichier_import:
                try:
                    df_imp = pd.read_excel(fichier_import)
                    st.success(f"✅ {len(df_imp)} lignes importées. Cliquez sur Sauvegarder pour fusionner.")
                    if st.button("💾 Fusionner avec l'EDT actuel", use_container_width=True):
                        df_new = pd.concat([df, df_imp], ignore_index=True)
                        df_new.to_excel(NOM_FICHIER_FIXE, index=False)
                        st.success("✅ Fichier maître mis à jour !")
                        st.rerun()
                except Exception as e:
                    st.error(f"Erreur import : {e}")

    # ============================================================
    # PORTAIL : GESTION ADMINISTRATIVE (BORDEREAUX)
    # ============================================================
    elif portail == "📢 Gestion Administrative":
        if not is_admin:
            st.error("🚫 Accès réservé à l'administration.")
            return

        st.header("📋 Gestion Administrative - Bordereaux & Documents")
        
        tab_bord, tab_pv = st.tabs(["📨 Bordereau d'envoi", "📄 PV / Procès-verbal"])

        with tab_bord:
            st.subheader("Génération de Bordereau d'Envoi")
            
            col1, col2 = st.columns(2)
            with col1:
                destinataire = st.selectbox("Destinataire :", [
                    "Le Doyen de la Faculté", 
                    "Le Vice-Doyen", 
                    "Le Chef de Département", 
                    "La Scolarité", 
                    "Autre"
                ])
                if destinataire == "Autre":
                    destinataire = st.text_input("Préciser :")
            with col2:
                ref_num = st.text_input("N° Référence :", f"001/FGE/ELT/{datetime.now().year}")
                date_bord = st.date_input("Date :", datetime.now())

            st.markdown("### 📎 Pièces jointes")
            pieces_df = st.data_editor(
                pd.DataFrame([
                    {"Désignation": "Emploi du temps S1", "Nombre": 1, "Observation": "Pour diffusion"},
                    {"Désignation": "Liste des étudiants", "Nombre": 1, "Observation": "Pour contrôle"}
                ]),
                column_config={
                    "Désignation": st.column_config.TextColumn("Désignation", required=True),
                    "Nombre": st.column_config.NumberColumn("Nombre", min_value=1),
                    "Observation": st.column_config.TextColumn("Observation")
                },
                num_rows="dynamic",
                use_container_width=True,
                key="pieces_bord"
            )

            if st.button("📄 Générer le Bordereau", use_container_width=True, type="primary"):
                # Génération simple HTML (fallback si python-docx non dispo)
                html_bord = f"""
                <div style="border:2px solid #1E3A8A; padding:30px; max-width:800px; margin:auto; font-family:Arial;">
                    <div style="text-align:center; border-bottom:2px solid #D4AF37; padding-bottom:15px; margin-bottom:20px;">
                        <h2 style="color:#1E3A8A; margin:0;">UNIVERSITÉ DJILLALI LIABES</h2>
                        <p style="margin:5px 0;">Faculté de Génie Électrique - Sidi Bel Abbès</p>
                        <h3 style="color:#D4AF37; margin:10px 0 0 0;">BORDEREAU D'ENVOI</h3>
                    </div>
                    <p><b>N° Référence :</b> {ref_num}</p>
                    <p><b>Date :</b> {date_bord.strftime('%d/%m/%Y')}</p>
                    <p><b>Destinataire :</b> {destinataire}</p>
                    <hr>
                    <table style="width:100%; border-collapse:collapse; margin-top:20px;">
                        <tr style="background:#1E3A8A; color:white;">
                            <th style="padding:10px; border:1px solid #333;">Désignation des pièces</th>
                            <th style="padding:10px; border:1px solid #333;">Nombre</th>
                            <th style="padding:10px; border:1px solid #333;">Observations</th>
                        </tr>
                """
                for _, row in pieces_df.iterrows():
                    html_bord += f"""
                        <tr>
                            <td style="padding:8px; border:1px solid #333;">{row['Désignation']}</td>
                            <td style="padding:8px; border:1px solid #333; text-align:center;">{row['Nombre']}</td>
                            <td style="padding:8px; border:1px solid #333;">{row['Observation']}</td>
                        </tr>
                    """
                html_bord += """
                    </table>
                    <div style="margin-top:40px; display:flex; justify-content:space-between;">
                        <div><b>Signature du responsable</b><br><br>_________________</div>
                        <div><b>Accusé de réception</b><br><br>_________________</div>
                    </div>
                </div>
                """
                st.success("✅ Bordereau généré")
                st.download_button("📥 Télécharger (HTML)", html_bord, f"Bordereau_{ref_num.replace('/', '_')}.html", "text/html")

        with tab_pv:
            st.subheader("📄 Génération de PV")
            st.info("Module de génération de Procès-verbaux de délibération")
            st.text_area("Contenu du PV :", height=200, placeholder="Saisir le contenu du PV ici...")
            if st.button("Générer le PV", use_container_width=True):
                st.success("✅ PV généré (simulation)")
                st.download_button("📥 Télécharger", "<html><body><h1>PV</h1></body></html>", "PV.html", "text/html")


# =============================================================================
# POINT D'ENTRÉE PRINCIPAL
# =============================================================================
if module_sel == "📊 Suivi d'Assiduité":
    run_assiduite()
else:
    run_edt() 



                     

import streamlit as st
import pandas as pd
import os
import hashlib
import io
from datetime import datetime, timedelta
from supabase import create_client
import streamlit as st

# =============================================================================
# FONCTIONS UTILITAIRES PRO POUR L'EXPORT (PDF / HTML / EXCEL)
# =============================================================================

import zipfile

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
    """Généré un PDF professionnel avec fpdf. Retourne des bytes utilisables par st.download_button."""
    try:
        from fpdf import FPDF
    except ImportError:
        return None, "fpdf non installe"

    class ProPDF(FPDF):
        def header(self):
            self.set_font('Arial', 'B', 9)
            self.set_text_color(30, 58, 138)
            header_text = sanitize_for_pdf("Plateforme de gestion des EDTs-Semestre 01_2026-2027 - Département d'Electrotechnique - FGE/UDL-SBA")
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
    pdf.alias_nb_pages()
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
    pdf.cell(0, 5, sanitize_for_pdf(f"Document géneré le {datetime.now().strftime('%d/%m/%Y a %H:%M')}"), 0, 0, "R")

    return bytes(pdf.output()), None


def generate_pro_html(df_source, title, subtitle=""):
    """Généré un HTML professionnel et responsive. Retourne une str."""
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
            <span class="badge">EDT Semestre 01_2026-2027 - Département d'électrotechnuqe-FGE/UDL-SBA</span>
        </div>
        <div class="content">
            <div class="meta">
                <span>Généré le {datetime.now().strftime('%d/%m/%Y a %H:%M')}</span>
                <span>{len(df_source) if df_source is not None else 0} lignes</span>
            </div>
            {table_html}
        </div>
        <div class="footer">
            Plateforme de gestion des EDTs - Département d'Electrotechnique - Faculte de Genie Electrique - UDL-SBA
        </div>
    </div>
</body>
</html>"""
    return html_content


def generate_pro_excel(df_source, title, sheet_name="Donnees"):
    """Généré un Excel professionnel avec xlsxwriter. Retourne des bytes."""
    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df_clean = df_source.fillna("").astype(str) if df_source is not None else pd.DataFrame()
        df_clean.to_excel(writer, index=False, sheet_name=sheet_name)

        workbook = writer.book
        worksheet = writer.sheets[sheet_name]

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

def generate_edt_individuel_pdf_classique(df_source, nom_enseignant):
    """Généré un PDF individuel avec en-tete PPER.03 centree et texte infos aligne a gauche."""
    try:
        from fpdf import FPDF
        import math
    except ImportError:
        return None, "fpdf non installe"
    
    if df_source is None or df_source.empty:
        return None, "Aucune donnee"
    
    # Dimensions exactes PPER.03
    W_LOGO = 1.19 * 25.4
    W_MILIEU = 3.70 * 25.4
    W_INFO = 1.40 * 25.4
    H_ENTETE = 1.04 * 25.4
    H_HAUT_MILIEU = 0.60 * 25.4
    H_BAS_MILIEU = H_ENTETE - H_HAUT_MILIEU
    W_TOT = W_LOGO + W_MILIEU + W_INFO
    
    jours_ordre = ["Dimanche", "Lundi", "Mardi", "Mercredi", "Jeudi"]
    horaires_ordre = [
        "8h - 9h", "8h - 9h30", "8h - 10h", "9h - 10h", "9h30 - 11h", 
        "10h - 11h", "11h - 12h", "11h - 12h30", "12h - 13h", 
        "12h30 - 14h", "13h - 14h30", "14h - 15h30", "14h - 16h", "15h30 - 17h"
    ]
    
    def norm(x):
        if not x or str(x).strip().lower() in ["non defini", "nan", "none", ""]:
            return "vide"
        s = str(x).strip().lower().replace(" ", "").replace("-", "").replace("–", "")
        s = s.replace(":00", "").replace("h00", "h")
        return s
    
    map_j = {norm(j): j for j in jours_ordre}
    map_h = {norm(h): h for h in horaires_ordre}
    
    # Remplacer les emojis par des codes courts pour un calcul de largeur fiable
    def format_cell(rows):
        items = []
        for _, r in rows.iterrows():
            code_up = str(r.get('Code', '')).upper()
            if 'COURS' in code_up:
                nat = '[C]'
            elif 'TD' in code_up:
                nat = '[T]'
            else:
                nat = '[P]'
            txt = f"{nat} {r.get('Enseignements', '')}\n{r.get('Lieu', '')}\n{r.get('Promotion', '')}"
            items.append(txt)
        return "\n".join(items)
    
    df = df_source.copy()
    df['Jours_Norm'] = df['Jours'].apply(norm)
    df['Horaire_Norm'] = df['Horaire'].apply(norm)
    
    # Pivot : Jours en lignes, Horaires en colonnes
    grouped = df.groupby(['Jours_Norm', 'Horaire_Norm']).apply(format_cell, include_groups=False)
    if grouped.empty:
        grid = pd.DataFrame(index=[norm(j) for j in jours_ordre], columns=[norm(h) for h in horaires_ordre]).fillna("")
    else:
        grid = grouped.unstack(fill_value="")
    
    jours_present = [j for j in [norm(j) for j in jours_ordre] if j in grid.index]
    horaires_present = [h for h in [norm(h) for h in horaires_ordre] if h in grid.columns]
    
    if not jours_present or not horaires_present:
        grid = pd.DataFrame(index=["Aucun"], columns=["Aucun"]).fillna("Aucun cours")
    else:
        grid = grid.reindex(index=jours_present, columns=horaires_present)
        grid.index = [map_j.get(i, i) for i in grid.index]
        grid.columns = [map_h.get(c, c) for c in grid.columns]
    
    class EDTIndivPDF(FPDF):
        def header(self):
            X0 = 10 + ((self.w - 20) - W_TOT) / 2
            Y0 = 10
            X_MILIEU = X0 + W_LOGO
            X_INFO = X_MILIEU + W_MILIEU
            Y_SEP = Y0 + H_HAUT_MILIEU
            
            self.set_draw_color(0, 0, 0)
            self.set_line_width(0.3)
            self.rect(X0, Y0, W_TOT, H_ENTETE, 'D')
            self.line(X_MILIEU, Y0, X_MILIEU, Y0 + H_ENTETE)
            self.line(X_INFO, Y0, X_INFO, Y0 + H_ENTETE)
            self.line(X_MILIEU, Y_SEP, X_INFO, Y_SEP)
            
            if os.path.exists("logo.PNG"):
                logo_w = W_LOGO - 4
                logo_h = H_ENTETE - 4
                self.image("logo.PNG", x=X0 + 2, y=Y0 + 2, w=logo_w, h=logo_h)
            
            self.set_xy(X_MILIEU, Y0 + 1.5)
            self.set_font('Arial', 'B', 11)
            self.cell(W_MILIEU, 5.5, sanitize_for_pdf("Universite Djillali Liabes"), 0, 2, "C")
            self.set_font('Arial', '', 10)
            self.cell(W_MILIEU, 5, sanitize_for_pdf("Sidi Bel Abbes"), 0, 2, "C")
            
            self.set_xy(X_MILIEU, Y_SEP + 0.5)
            self.set_font('Arial', 'B', 12)
            self.cell(W_MILIEU, H_BAS_MILIEU - 1, sanitize_for_pdf("EMPLOI DU TEMPS"), 0, 0, "C")
            
            self.set_font('Arial', '', 9)
            line_h = H_ENTETE / 4
            infos = [
                "Code : PPER.03",
                "Révision : 00",
                "Date : 16/05/2026",
                f"Pages : {self.page_no()}/{{nb}}"
            ]
            for i, info in enumerate(infos):
                self.set_xy(X_INFO + 1.5, Y0 + 0.5 + i * line_h)
                self.cell(W_INFO - 3, line_h, sanitize_for_pdf(info), 0, 2, "L")
            
            self.set_y(Y0 + H_ENTETE + 5)
        
        def footer(self):
            self.set_y(-12)
            self.set_font('Arial', 'I', 7)
            self.set_text_color(128, 128, 128)
            self.cell(0, 10, sanitize_for_pdf(f"{self.page_no()}/{{nb}}"), 0, 0, "R")
    
    pdf = EDTIndivPDF(orientation="L", unit="mm", format="A4")
    pdf.alias_nb_pages()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    
    # Sous-titre
    pdf.set_font("Arial", "B", 11)
    pdf.set_text_color(30, 58, 138)
    pdf.cell(0, 8, sanitize_for_pdf(f"EMPLOI DU TEMPS INDIVIDUEL - {nom_enseignant.upper()}"), 0, 1, "C")
    pdf.set_font("Arial", "I", 8)
    pdf.set_text_color(100, 100, 100)
    pdf.cell(0, 5, sanitize_for_pdf("Semestre 01 - Département d'Electrotechnique - FGE/UDL-SBA"), 0, 1, "C")
    pdf.ln(3)
        
    if grid.empty:
        pdf.set_font("Arial", "", 10)
        pdf.cell(0, 10, "Aucun cours programme pour cet enseignant.", 0, 1, "C")
        return bytes(pdf.output()), None
    
    n_cols = len(grid.columns)
    page_w = pdf.w - 20
    col_jour_w = 22
    col_h_w = (page_w - col_jour_w) / n_cols if n_cols > 0 else page_w
    
    # Parametres de rendu (marges internes de securite)
    interline = 3.2      # hauteur d'une ligne de texte
    margin_h = 4.0       # marge horizontale totale (2mm de chaque cote)
    padding_v = 3.0      # marge verticale de securite
    
    # === CALCUL PRECIS DES HAUTEURS DE LIGNE ===
    pdf.set_font("Arial", "", 5.5)
    row_heights = []
    for _, row in grid.iterrows():
        max_lines = 1
        for val in row:
            if val and str(val).strip():
                txt_propre = sanitize_for_pdf(str(val))
                lines = 0
                for para in txt_propre.split('\n'):
                    w_txt = pdf.get_string_width(para)
                    if w_txt == 0:
                        lines += 1
                    else:
                        # Largeur utile reelle avec marge de securite
                        usable_w = max(col_h_w - margin_h - 1.0, 10)
                        lines += max(1, math.ceil(w_txt / usable_w))
                if lines > max_lines:
                    max_lines = lines
        # Hauteur exacte necessaire (PAS de plafond arbitraire)
        h_needed = max_lines * interline + padding_v * 2 + 2.0
        row_heights.append(max(10, h_needed))
    
    # En-tetes du tableau EDT
    pdf.set_font("Arial", "B", 7)
    pdf.set_fill_color(30, 58, 138)
    pdf.set_text_color(255, 255, 255)
    pdf.cell(col_jour_w, 8, "JOUR", 1, 0, "C", True)
    for h in grid.columns:
        h_txt = sanitize_for_pdf(str(h))
        if len(h_txt) > 12:
            h_txt = h_txt.replace(" - ", "-").replace(" ", "")
        pdf.cell(col_h_w, 8, h_txt, 1, 0, "C", True)
    pdf.ln()
    
    # Donnees du tableau EDT
    pdf.set_text_color(0, 0, 0)
    
    for idx, ((jour, row), row_h) in enumerate(zip(grid.iterrows(), row_heights)):
        # Saut de page si la ligne entiere ne tient pas dans l'espace restant
        if pdf.get_y() + row_h > pdf.h - 15:
            pdf.add_page()
            # Re-imprimer l'en-tete du tableau sur la nouvelle page
            pdf.set_font("Arial", "B", 7)
            pdf.set_fill_color(30, 58, 138)
            pdf.set_text_color(255, 255, 255)
            pdf.cell(col_jour_w, 8, "JOUR", 1, 0, "C", True)
            for h in grid.columns:
                h_txt = sanitize_for_pdf(str(h))
                if len(h_txt) > 12:
                    h_txt = h_txt.replace(" - ", "-").replace(" ", "")
                pdf.cell(col_h_w, 8, h_txt, 1, 0, "C", True)
            pdf.ln()
            pdf.set_text_color(0, 0, 0)
        
        bg_color = (248, 250, 252) if idx % 2 == 0 else (255, 255, 255)
        
        pdf.set_font("Arial", "B", 7)
        pdf.set_fill_color(*bg_color)
        pdf.set_draw_color(180, 180, 180)
        pdf.cell(col_jour_w, row_h, sanitize_for_pdf(str(jour)), 1, 0, "C", True)
        
        pdf.set_font("Arial", "", 5.5)
        for val in row:
            cell_text = sanitize_for_pdf(str(val)) if val else ""
            x, y = pdf.get_x(), pdf.get_y()
            
            if cell_text.strip():
                raw_up = str(val).upper()
                if "COURS" in raw_up:
                    bg = (225, 238, 255)
                elif "TD" in raw_up:
                    bg = (232, 252, 235)
                elif "TP" in raw_up:
                    bg = (255, 235, 235)
                else:
                    bg = bg_color
            else:
                bg = bg_color
            
            pdf.set_fill_color(*bg)
            pdf.set_draw_color(180, 180, 180)
            pdf.rect(x, y, col_h_w, row_h, 'FD')
            
            if cell_text.strip():
                # Recalcul precis du nombre de lignes pour le centrage vertical
                n_lines = 0
                for para in cell_text.split('\n'):
                    w_txt = pdf.get_string_width(para)
                    if w_txt == 0:
                        n_lines += 1
                    else:
                        usable_w = max(col_h_w - margin_h - 1.0, 10)
                        n_lines += max(1, math.ceil(w_txt / usable_w))
                
                text_block_h = n_lines * interline
                offset_y = max((row_h - text_block_h) / 2, padding_v / 2)
                
                pdf.set_xy(x + margin_h / 2, y + offset_y)
                pdf.multi_cell(col_h_w - margin_h, interline, cell_text, 0, "L")
                pdf.set_xy(x + col_h_w, y)
            else:
                pdf.set_xy(x + col_h_w, y)
        pdf.ln(row_h)
    
    return bytes(pdf.output()), None    
    # ═══════════════════════════════════════════════════════════════
    # PASSE 1 : COMPTAGE SILENCIEUX (sans barre de progression)
    # ═══════════════════════════════════════════════════════════════
    pdf_count = _build_pdf(total_pages=0, show_progress=False)
    total_pages = pdf_count.page_no()
def generate_edt_tous_enseignants_pdf(df_source, progress_bar=None):
    """Genere un PDF unique avec l'EDT de TOUS les enseignants (format grille jour/horaire)."""
    try:
        from fpdf import FPDF
        import math
    except ImportError:
        return None, "fpdf non installe"

    if df_source is None or df_source.empty:
        return None, "Aucune donnee"

    W_LOGO = 1.19 * 25.4
    W_MILIEU = 3.70 * 25.4
    W_INFO = 1.40 * 25.4
    H_ENTETE = 1.04 * 25.4
    H_HAUT_MILIEU = 0.60 * 25.4
    H_BAS_MILIEU = H_ENTETE - H_HAUT_MILIEU
    W_TOT = W_LOGO + W_MILIEU + W_INFO
    MARGE_BAS = 15

    jours_ordre = ["Dimanche", "Lundi", "Mardi", "Mercredi", "Jeudi"]
    horaires_ordre = [
        "8h - 9h", "8h - 9h30", "8h - 10h", "9h - 10h", "9h30 - 11h", 
        "10h - 11h", "11h - 12h", "11h - 12h30", "12h - 13h", 
        "12h30 - 14h", "13h - 14h30", "14h - 15h30", "14h - 16h", "15h30 - 17h"
    ]

    def norm(x):
        if not x or str(x).strip().lower() in ["non defini", "nan", "none", ""]:
            return "vide"
        s = str(x).strip().lower().replace(" ", "").replace("-", "").replace("–", "")
        s = s.replace(":00", "").replace("h00", "h")
        return s

    map_j = {norm(j): j for j in jours_ordre}
    map_h = {norm(h): h for h in horaires_ordre}

    def format_cell(rows):
        items = []
        for _, r in rows.iterrows():
            code_up = str(r.get('Code', '')).upper()
            if 'COURS' in code_up:
                nat = '[C]'
            elif 'TD' in code_up:
                nat = '[TD]'
            else:
                nat = '[TP]'
            txt = f"{nat} {r.get('Enseignements', '')}\nPromo: {r.get('Promotion', '')}\nSalle: {r.get('Lieu', '')}"
            items.append(txt)
        return "\n".join(items)

    df = df_source.copy()
    df['Jours_Norm'] = df['Jours'].apply(norm)
    df['Horaire_Norm'] = df['Horaire'].apply(norm)

    enseignants = sorted([e for e in df['Enseignants'].unique() 
                         if e and str(e).strip() not in ["", "nan", "None", "Non defini", "Non défini"]])
    if not enseignants:
        return None, "Aucun enseignant trouve"

    def _build_pdf(total_pages, show_progress=False):
        class EnsGlobalPDF(FPDF):
            def header(self):
                X0 = 10 + ((self.w - 20) - W_TOT) / 2
                Y0 = 10
                X_MILIEU = X0 + W_LOGO
                X_INFO = X_MILIEU + W_MILIEU
                Y_SEP = Y0 + H_HAUT_MILIEU

                self.set_draw_color(0, 0, 0)
                self.set_line_width(0.3)
                self.rect(X0, Y0, W_TOT, H_ENTETE, 'D')
                self.line(X_MILIEU, Y0, X_MILIEU, Y0 + H_ENTETE)
                self.line(X_INFO, Y0, X_INFO, Y0 + H_ENTETE)
                self.line(X_MILIEU, Y_SEP, X_INFO, Y_SEP)

                if os.path.exists("logo.PNG"):
                    self.image("logo.PNG", x=X0 + 2, y=Y0 + 2, w=W_LOGO - 4, h=H_ENTETE - 4)

                self.set_xy(X_MILIEU, Y0 + 1.5)
                self.set_font('Arial', 'B', 11)
                self.cell(W_MILIEU, 5.5, sanitize_for_pdf("Universite Djillali Liabes"), 0, 2, "C")
                self.set_font('Arial', '', 10)
                self.cell(W_MILIEU, 5, sanitize_for_pdf("Sidi Bel Abbes"), 0, 2, "C")

                self.set_xy(X_MILIEU, Y_SEP + 0.5)
                self.set_font('Arial', 'B', 12)
                self.cell(W_MILIEU, H_BAS_MILIEU - 1, sanitize_for_pdf("EMPLOI DU TEMPS"), 0, 0, "C")

                self.set_font('Arial', '', 9)
                line_h = H_ENTETE / 4
                infos = [
                    "Code : PPER.03",
                    "Revision : 00",
                    "Date : 16/05/2026",
                    f"Pages : {self.page_no()}/{total_pages}"
                ]
                for i, info in enumerate(infos):
                    self.set_xy(X_INFO + 1.5, Y0 + 0.5 + i * line_h)
                    self.cell(W_INFO - 3, line_h, sanitize_for_pdf(info), 0, 2, "L")

                self.set_y(Y0 + H_ENTETE + 5)

            def footer(self):
                self.set_y(-15)
                self.set_font('Arial', '', 8)
                self.set_text_color(80, 80, 80)
                self.cell(0, 10, f"{self.page_no()}/{total_pages}", 0, 0, "R")

        pdf = EnsGlobalPDF(orientation="L", unit="mm", format="A4")
        pdf.set_auto_page_break(auto=False, margin=MARGE_BAS)

        def draw_title(pdf, ens):
            pdf.set_font("Arial", "B", 12)
            pdf.set_text_color(30, 58, 138)
            pdf.cell(0, 8, sanitize_for_pdf(f"EMPLOI DU TEMPS - {str(ens).upper()}"), 0, 1, "C")
            pdf.set_font("Arial", "I", 8)
            pdf.set_text_color(100, 100, 100)
            pdf.cell(0, 5, sanitize_for_pdf("Semestre 01 - Departement d'Electrotechnique - FGE/UDL-SBA"), 0, 1, "C")
            pdf.ln(2)

        def draw_table_header(pdf, grid, col_jour_w, col_h_w):
            pdf.set_font("Arial", "B", 7)
            pdf.set_fill_color(30, 58, 138)
            pdf.set_text_color(255, 255, 255)
            pdf.cell(col_jour_w, 8, "JOUR", 1, 0, "C", True)
            for h in grid.columns:
                h_txt = sanitize_for_pdf(str(h))
                if len(h_txt) > 12:
                    h_txt = h_txt.replace(" - ", "-").replace(" ", "")
                pdf.cell(col_h_w, 8, h_txt, 1, 0, "C", True)
            pdf.ln()
            pdf.set_text_color(0, 0, 0)

        n_ens = len(enseignants)

        for idx_ens, ens in enumerate(enseignants):
            if show_progress and progress_bar is not None:
                progress_bar.progress((idx_ens) / n_ens, text=f"Generation : {ens} ({idx_ens+1}/{n_ens})")

            df_ens = df[df['Enseignants'] == ens].copy()
            if df_ens.empty:
                continue

            pdf.add_page()

            grouped = df_ens.groupby(['Jours_Norm', 'Horaire_Norm']).apply(format_cell, include_groups=False)
            grid = grouped.unstack(fill_value="") if not grouped.empty else pd.DataFrame()

            jours_present = [j for j in [norm(j) for j in jours_ordre] if j in grid.index]
            horaires_present = [h for h in [norm(h) for h in horaires_ordre] if h in grid.columns]

            if not jours_present or not horaires_present:
                draw_title(pdf, ens)
                pdf.set_font("Arial", "", 10)
                pdf.cell(0, 10, "Aucun cours programme pour cet enseignant.", 0, 1, "C")
                continue

            grid = grid.reindex(index=jours_present, columns=horaires_present)
            grid.index = [map_j.get(i, i) for i in grid.index]
            grid.columns = [map_h.get(c, c) for c in grid.columns]

            n_cols = len(grid.columns)
            page_w = pdf.w - 20
            col_jour_w = 22
            col_h_w = (page_w - col_jour_w) / n_cols if n_cols > 0 else page_w

            interline = 3.2
            margin_h  = 4.0
            padding_v = 3.0

            pdf.set_font("Arial", "", 5.5)
            row_heights = []
            for _, row in grid.iterrows():
                max_lines = 1
                for val in row:
                    if val and str(val).strip():
                        txt_propre = sanitize_for_pdf(str(val))
                        lines = 0
                        for para in txt_propre.split('\n'):
                            w_txt = pdf.get_string_width(para)
                            if w_txt == 0:
                                lines += 1
                            else:
                                usable_w = max(col_h_w - margin_h - 1.0, 10)
                                lines += max(1, math.ceil(w_txt / usable_w))
                        if lines > max_lines:
                            max_lines = lines
                h_needed = max_lines * interline + padding_v * 2 + 2.0
                row_heights.append(max(10, h_needed))

            title_h = 8 + 5 + 2
            header_h = 8
            if pdf.get_y() + title_h + header_h + row_heights[0] > pdf.h - MARGE_BAS:
                pdf.add_page()

            draw_title(pdf, ens)
            draw_table_header(pdf, grid, col_jour_w, col_h_w)

            for idx, ((jour, row), row_h) in enumerate(zip(grid.iterrows(), row_heights)):
                if pdf.get_y() + row_h > pdf.h - MARGE_BAS:
                    pdf.add_page()
                    draw_title(pdf, ens)
                    draw_table_header(pdf, grid, col_jour_w, col_h_w)

                bg_color = (248, 250, 252) if idx % 2 == 0 else (255, 255, 255)

                pdf.set_font("Arial", "B", 7)
                pdf.set_fill_color(*bg_color)
                pdf.set_draw_color(180, 180, 180)
                pdf.cell(col_jour_w, row_h, sanitize_for_pdf(str(jour)), 1, 0, "C", True)

                pdf.set_font("Arial", "", 5.5)
                for val in row:
                    cell_text = sanitize_for_pdf(str(val)) if val else ""
                    x, y = pdf.get_x(), pdf.get_y()

                    if cell_text.strip():
                        raw_up = str(val).upper()
                        if "COURS" in raw_up:
                            bg = (225, 238, 255)
                        elif "TD" in raw_up:
                            bg = (232, 252, 235)
                        elif "TP" in raw_up:
                            bg = (255, 235, 235)
                        else:
                            bg = bg_color
                    else:
                        bg = bg_color

                    pdf.set_fill_color(*bg)
                    pdf.set_draw_color(180, 180, 180)
                    pdf.rect(x, y, col_h_w, row_h, 'FD')

                    if cell_text.strip():
                        n_lines = 0
                        for para in cell_text.split('\n'):
                            w_txt = pdf.get_string_width(para)
                            if w_txt == 0:
                                n_lines += 1
                            else:
                                usable_w = max(col_h_w - margin_h - 1.0, 10)
                                n_lines += max(1, math.ceil(w_txt / usable_w))

                        text_block_h = n_lines * interline
                        offset_y = max((row_h - text_block_h) / 2, padding_v / 2)

                        pdf.set_xy(x + margin_h / 2, y + offset_y)
                        pdf.multi_cell(col_h_w - margin_h, interline, cell_text, 0, "L")
                        pdf.set_xy(x + col_h_w, y)
                    else:
                        pdf.set_xy(x + col_h_w, y)

                pdf.ln(row_h)

        return pdf

    pdf_count = _build_pdf(total_pages=0, show_progress=False)
    total_pages = pdf_count.page_no()

    if progress_bar is not None:
        progress_bar.progress(0, text=f"Generation finale sur {total_pages} pages...")

    pdf_final = _build_pdf(total_pages=total_pages, show_progress=True)

    if progress_bar is not None:
        progress_bar.empty()

    return bytes(pdf_final.output()), None



def generate_edt_toutes_promotions_pdf(df_source, progress_bar=None):
    """Généré un PDF unique avec l'EDT de TOUTES les promotions (format grille jour/horaire)."""
    try:
        from fpdf import FPDF
        import math
    except ImportError:
        return None, "fpdf non installe"
    
    if df_source is None or df_source.empty:
        return None, "Aucune donnee"
    
    # Dimensions en-tete PPER.03 (inch -> mm)
    W_LOGO = 1.19 * 25.4
    W_MILIEU = 3.70 * 25.4
    W_INFO = 1.40 * 25.4
    H_ENTETE = 1.04 * 25.4
    H_HAUT_MILIEU = 0.60 * 25.4
    H_BAS_MILIEU = H_ENTETE - H_HAUT_MILIEU
    W_TOT = W_LOGO + W_MILIEU + W_INFO
    MARGE_BAS = 15
    
    jours_ordre = ["Dimanche", "Lundi", "Mardi", "Mercredi", "Jeudi"]
    horaires_ordre = [
        "8h - 9h", "8h - 9h30", "8h - 10h", "9h - 10h", "9h30 - 11h", 
        "10h - 11h", "11h - 12h", "11h - 12h30", "12h - 13h", 
        "12h30 - 14h", "13h - 14h30", "14h - 15h30", "14h - 16h", "15h30 - 17h"
    ]
    
    def norm(x):
        if not x or str(x).strip().lower() in ["non defini", "nan", "none", ""]:
            return "vide"
        s = str(x).strip().lower().replace(" ", "").replace("-", "").replace("–", "")
        s = s.replace(":00", "").replace("h00", "h")
        return s
    
    map_j = {norm(j): j for j in jours_ordre}
    map_h = {norm(h): h for h in horaires_ordre}
    
    def format_cell(rows):
        items = []
        for _, r in rows.iterrows():
            code_up = str(r.get('Code', '')).upper()
            if 'COURS' in code_up:
                nat = '[C]'
            elif 'TD' in code_up:
                nat = '[TD]'
            else:
                nat = '[TP]'
            
            txt = f"{nat} {r.get('Enseignements', '')}\nProf: {r.get('Enseignants', '')}\nSalle: {r.get('Lieu', '')}"
            items.append(txt)
        return "\n".join(items)
    
    df = df_source.copy()
    df['Jours_Norm'] = df['Jours'].apply(norm)
    df['Horaire_Norm'] = df['Horaire'].apply(norm)
    
    promotions = sorted([p for p in df['Promotion'].unique() 
                        if p and str(p).strip() not in ["", "nan", "None", "Non defini"]])
    if not promotions:
        return None, "Aucune promotion trouvee"
    
    def _build_pdf(total_pages, show_progress=False):
        class PromoGlobalPDF(FPDF):
            def header(self):
                X0 = 10 + ((self.w - 20) - W_TOT) / 2
                Y0 = 10
                X_MILIEU = X0 + W_LOGO
                X_INFO = X_MILIEU + W_MILIEU
                Y_SEP = Y0 + H_HAUT_MILIEU
                
                self.set_draw_color(0, 0, 0)
                self.set_line_width(0.3)
                self.rect(X0, Y0, W_TOT, H_ENTETE, 'D')
                self.line(X_MILIEU, Y0, X_MILIEU, Y0 + H_ENTETE)
                self.line(X_INFO, Y0, X_INFO, Y0 + H_ENTETE)
                self.line(X_MILIEU, Y_SEP, X_INFO, Y_SEP)
                
                if os.path.exists("logo.PNG"):
                    self.image("logo.PNG", x=X0 + 2, y=Y0 + 2, w=W_LOGO - 4, h=H_ENTETE - 4)
                
                self.set_xy(X_MILIEU, Y0 + 1.5)
                self.set_font('Arial', 'B', 11)
                self.cell(W_MILIEU, 5.5, sanitize_for_pdf("Universite Djillali Liabes"), 0, 2, "C")
                self.set_font('Arial', '', 10)
                self.cell(W_MILIEU, 5, sanitize_for_pdf("Sidi Bel Abbes"), 0, 2, "C")
                
                self.set_xy(X_MILIEU, Y_SEP + 0.5)
                self.set_font('Arial', 'B', 12)
                self.cell(W_MILIEU, H_BAS_MILIEU - 1, sanitize_for_pdf("EMPLOI DU TEMPS"), 0, 0, "C")
                
                self.set_font('Arial', '', 9)
                line_h = H_ENTETE / 4
                infos = [
                    "Code : PPER.03",
                    "Révision : 00",
                    "Date : 16/05/2026",
                    f"Pages : {total_pages}/{total_pages}"
                ]
                for i, info in enumerate(infos):
                    self.set_xy(X_INFO + 1.5, Y0 + 0.5 + i * line_h)
                    self.cell(W_INFO - 3, line_h, sanitize_for_pdf(info), 0, 2, "L")
                
                self.set_y(Y0 + H_ENTETE + 5)
            
            def footer(self):
                self.set_y(-15)
                self.set_font('Arial', '', 8)
                self.set_text_color(80, 80, 80)
                self.cell(0, 10, f"{self.page_no()}/{total_pages}", 0, 0, "R")
        
        pdf = PromoGlobalPDF(orientation="L", unit="mm", format="A4")
        pdf.set_auto_page_break(auto=False, margin=MARGE_BAS)
        
        def draw_title(pdf, promo):
            pdf.set_font("Arial", "B", 12)
            pdf.set_text_color(30, 58, 138)
            pdf.cell(0, 8, sanitize_for_pdf(f"PROMOTION {str(promo).upper()}"), 0, 1, "C")
            pdf.set_font("Arial", "I", 8)
            pdf.set_text_color(100, 100, 100)
            pdf.cell(0, 5, sanitize_for_pdf("Semestre 01 - Département d'Electrotechnique - FGE/UDL-SBA"), 0, 1, "C")
            pdf.ln(2)
        
        def draw_table_header(pdf, grid, col_jour_w, col_h_w):
            pdf.set_font("Arial", "B", 7)
            pdf.set_fill_color(30, 58, 138)
            pdf.set_text_color(255, 255, 255)
            pdf.cell(col_jour_w, 8, "JOUR", 1, 0, "C", True)
            for h in grid.columns:
                h_txt = sanitize_for_pdf(str(h))
                if len(h_txt) > 12:
                    h_txt = h_txt.replace(" - ", "-").replace(" ", "")
                pdf.cell(col_h_w, 8, h_txt, 1, 0, "C", True)
            pdf.ln()
            pdf.set_text_color(0, 0, 0)
        
        n_promo = len(promotions)
        
        for idx_promo, promo in enumerate(promotions):
            if show_progress and progress_bar is not None:
                progress_bar.progress((idx_promo) / n_promo, text=f"Generation : {promo} ({idx_promo+1}/{n_promo})")
            
            df_promo = df[df['Promotion'] == promo].copy()
            if df_promo.empty:
                continue
            
            pdf.add_page()
            
            grouped = df_promo.groupby(['Jours_Norm', 'Horaire_Norm']).apply(format_cell, include_groups=False)
            grid = grouped.unstack(fill_value="") if not grouped.empty else pd.DataFrame()
            
            jours_present = [j for j in [norm(j) for j in jours_ordre] if j in grid.index]
            horaires_present = [h for h in [norm(h) for h in horaires_ordre] if h in grid.columns]
            
            if not jours_present or not horaires_present:
                draw_title(pdf, promo)
                pdf.set_font("Arial", "", 10)
                pdf.cell(0, 10, "Aucun cours programme pour cette promotion.", 0, 1, "C")
                continue
            
            grid = grid.reindex(index=jours_present, columns=horaires_present)
            grid.index = [map_j.get(i, i) for i in grid.index]
            grid.columns = [map_h.get(c, c) for c in grid.columns]
            
            n_cols = len(grid.columns)
            page_w = pdf.w - 20
            col_jour_w = 22
            col_h_w = (page_w - col_jour_w) / n_cols if n_cols > 0 else page_w
            
            interline = 3.2
            margin_h  = 4.0
            padding_v = 3.0
            
            pdf.set_font("Arial", "", 5.5)
            row_heights = []
            for _, row in grid.iterrows():
                max_lines = 1
                for val in row:
                    if val and str(val).strip():
                        txt_propre = sanitize_for_pdf(str(val))
                        lines = 0
                        for para in txt_propre.split('\n'):
                            w_txt = pdf.get_string_width(para)
                            if w_txt == 0:
                                lines += 1
                            else:
                                usable_w = max(col_h_w - margin_h - 1.0, 10)
                                lines += max(1, math.ceil(w_txt / usable_w))
                        if lines > max_lines:
                            max_lines = lines
                h_needed = max_lines * interline + padding_v * 2 + 2.0
                row_heights.append(max(10, h_needed))
            
            title_h = 8 + 5 + 2
            header_h = 8
            if pdf.get_y() + title_h + header_h + row_heights[0] > pdf.h - MARGE_BAS:
                pdf.add_page()
            
            draw_title(pdf, promo)
            draw_table_header(pdf, grid, col_jour_w, col_h_w)
            
            for idx, ((jour, row), row_h) in enumerate(zip(grid.iterrows(), row_heights)):
                if pdf.get_y() + row_h > pdf.h - MARGE_BAS:
                    pdf.add_page()
                    draw_title(pdf, promo)
                    draw_table_header(pdf, grid, col_jour_w, col_h_w)
                
                bg_color = (248, 250, 252) if idx % 2 == 0 else (255, 255, 255)
                
                pdf.set_font("Arial", "B", 7)
                pdf.set_fill_color(*bg_color)
                pdf.set_draw_color(180, 180, 180)
                pdf.cell(col_jour_w, row_h, sanitize_for_pdf(str(jour)), 1, 0, "C", True)
                
                pdf.set_font("Arial", "", 5.5)
                for val in row:
                    cell_text = sanitize_for_pdf(str(val)) if val else ""
                    x, y = pdf.get_x(), pdf.get_y()
                    
                    if cell_text.strip():
                        raw_up = str(val).upper()
                        if "COURS" in raw_up:
                            bg = (225, 238, 255)
                        elif "TD" in raw_up:
                            bg = (232, 252, 235)
                        elif "TP" in raw_up:
                            bg = (255, 235, 235)
                        else:
                            bg = bg_color
                    else:
                        bg = bg_color
                    
                    pdf.set_fill_color(*bg)
                    pdf.set_draw_color(180, 180, 180)
                    pdf.rect(x, y, col_h_w, row_h, 'FD')
                    
                    if cell_text.strip():
                        n_lines = 0
                        for para in cell_text.split('\n'):
                            w_txt = pdf.get_string_width(para)
                            if w_txt == 0:
                                n_lines += 1
                            else:
                                usable_w = max(col_h_w - margin_h - 1.0, 10)
                                n_lines += max(1, math.ceil(w_txt / usable_w))
                        
                        text_block_h = n_lines * interline
                        offset_y = max((row_h - text_block_h) / 2, padding_v / 2)
                        
                        pdf.set_xy(x + margin_h / 2, y + offset_y)
                        pdf.multi_cell(col_h_w - margin_h, interline, cell_text, 0, "L")
                        pdf.set_xy(x + col_h_w, y)
                    else:
                        pdf.set_xy(x + col_h_w, y)
                
                pdf.ln(row_h)
        
        return pdf
    
    # PASSE 1 : comptage silencieux
    pdf_count = _build_pdf(total_pages=0, show_progress=False)
    total_pages = pdf_count.page_no()
    
    # PASSE 2 : generation avec pagination correcte
    if progress_bar is not None:
        progress_bar.progress(0, text=f"Generation finale sur {total_pages} pages...")
    
    pdf_final = _build_pdf(total_pages=total_pages, show_progress=True)
    
    if progress_bar is not None:
        progress_bar.empty()
    
    return bytes(pdf_final.output()), None    
def generate_edt_tous_lieux_pdf(df_source, progress_bar=None):
    """Genere un PDF unique avec le planning de TOUS les lieux (format grille jour/horaire)."""
    try:
        from fpdf import FPDF
        import math
    except ImportError:
        return None, "fpdf non installe"
    
    if df_source is None or df_source.empty:
        return None, "Aucune donnee"
    
    # Dimensions en-tete PPER.03 (inch -> mm)
    W_LOGO = 1.19 * 25.4
    W_MILIEU = 3.70 * 25.4
    W_INFO = 1.40 * 25.4
    H_ENTETE = 1.04 * 25.4
    H_HAUT_MILIEU = 0.60 * 25.4
    H_BAS_MILIEU = H_ENTETE - H_HAUT_MILIEU
    W_TOT = W_LOGO + W_MILIEU + W_INFO
    MARGE_BAS = 15
    
    jours_ordre = ["Dimanche", "Lundi", "Mardi", "Mercredi", "Jeudi"]
    horaires_ordre = [
        "8h - 9h", "8h - 9h30", "8h - 10h", "9h - 10h", "9h30 - 11h", 
        "10h - 11h", "11h - 12h", "11h - 12h30", "12h - 13h", 
        "12h30 - 14h", "13h - 14h30", "14h - 15h30", "14h - 16h", "15h30 - 17h"
    ]
    
    def norm(x):
        if not x or str(x).strip().lower() in ["non defini", "nan", "none", ""]:
            return "vide"
        s = str(x).strip().lower().replace(" ", "").replace("-", "").replace("–", "")
        s = s.replace(":00", "").replace("h00", "h")
        return s
    
    map_j = {norm(j): j for j in jours_ordre}
    map_h = {norm(h): h for h in horaires_ordre}
    
    def format_cell(rows):
        items = []
        for _, r in rows.iterrows():
            code_up = str(r.get('Code', '')).upper()
            if 'COURS' in code_up:
                nat = '[C]'
            elif 'TD' in code_up:
                nat = '[TD]'
            else:
                nat = '[TP]'
            txt = f"{nat} {r.get('Enseignements', '')}\nProf: {r.get('Enseignants', '')}\nPromo: {r.get('Promotion', '')}"
            items.append(txt)
        return "\n".join(items)
    
    df = df_source.copy()
    df['Jours_Norm'] = df['Jours'].apply(norm)
    df['Horaire_Norm'] = df['Horaire'].apply(norm)
    
    lieux = sorted([l for l in df['Lieu'].unique() 
                    if l and str(l).strip() not in ["", "nan", "None", "Non defini"]])
    if not lieux:
        return None, "Aucun lieu trouve"
    
    def _build_pdf(total_pages, show_progress=False):
        class LieuGlobalPDF(FPDF):
            def header(self):
                X0 = 10 + ((self.w - 20) - W_TOT) / 2
                Y0 = 10
                X_MILIEU = X0 + W_LOGO
                X_INFO = X_MILIEU + W_MILIEU
                Y_SEP = Y0 + H_HAUT_MILIEU
                
                self.set_draw_color(0, 0, 0)
                self.set_line_width(0.3)
                self.rect(X0, Y0, W_TOT, H_ENTETE, 'D')
                self.line(X_MILIEU, Y0, X_MILIEU, Y0 + H_ENTETE)
                self.line(X_INFO, Y0, X_INFO, Y0 + H_ENTETE)
                self.line(X_MILIEU, Y_SEP, X_INFO, Y_SEP)
                
                if os.path.exists("logo.PNG"):
                    self.image("logo.PNG", x=X0 + 2, y=Y0 + 2, w=W_LOGO - 4, h=H_ENTETE - 4)
                
                self.set_xy(X_MILIEU, Y0 + 1.5)
                self.set_font('Arial', 'B', 11)
                self.cell(W_MILIEU, 5.5, sanitize_for_pdf("Universite Djillali Liabes"), 0, 2, "C")
                self.set_font('Arial', '', 10)
                self.cell(W_MILIEU, 5, sanitize_for_pdf("Sidi Bel Abbes"), 0, 2, "C")
                
                self.set_xy(X_MILIEU, Y_SEP + 0.5)
                self.set_font('Arial', 'B', 12)
                self.cell(W_MILIEU, H_BAS_MILIEU - 1, sanitize_for_pdf("EMPLOI DU TEMPS"), 0, 0, "C")
                
                self.set_font('Arial', '', 9)
                line_h = H_ENTETE / 4
                infos = [
                    "Code : PPER.03",
                    "Revision : 00",
                    "Date : 16/05/2026",
                    f"Pages : {self.page_no()}/{total_pages}"
                ]
                for i, info in enumerate(infos):
                    self.set_xy(X_INFO + 1.5, Y0 + 0.5 + i * line_h)
                    self.cell(W_INFO - 3, line_h, sanitize_for_pdf(info), 0, 2, "L")
                
                self.set_y(Y0 + H_ENTETE + 5)
            
            def footer(self):
                self.set_y(-15)
                self.set_font('Arial', '', 8)
                self.set_text_color(80, 80, 80)
                self.cell(0, 10, f"{self.page_no()}/{total_pages}", 0, 0, "R")
        
        pdf = LieuGlobalPDF(orientation="L", unit="mm", format="A4")
        pdf.set_auto_page_break(auto=False, margin=MARGE_BAS)
        
        def draw_title(pdf, lieu):
            pdf.set_font("Arial", "B", 12)
            pdf.set_text_color(30, 58, 138)
            pdf.cell(0, 8, sanitize_for_pdf(f"PLANNING - {str(lieu).upper()}"), 0, 1, "C")
            pdf.set_font("Arial", "I", 8)
            pdf.set_text_color(100, 100, 100)
            pdf.cell(0, 5, sanitize_for_pdf("Semestre 01 - Departement d'Electrotechnique - FGE/UDL-SBA"), 0, 1, "C")
            pdf.ln(2)
        
        def draw_table_header(pdf, grid, col_jour_w, col_h_w):
            pdf.set_font("Arial", "B", 7)
            pdf.set_fill_color(30, 58, 138)
            pdf.set_text_color(255, 255, 255)
            pdf.cell(col_jour_w, 8, "JOUR", 1, 0, "C", True)
            for h in grid.columns:
                h_txt = sanitize_for_pdf(str(h))
                if len(h_txt) > 12:
                    h_txt = h_txt.replace(" - ", "-").replace(" ", "")
                pdf.cell(col_h_w, 8, h_txt, 1, 0, "C", True)
            pdf.ln()
            pdf.set_text_color(0, 0, 0)
        
        n_lieu = len(lieux)
        
        for idx_lieu, lieu in enumerate(lieux):
            if show_progress and progress_bar is not None:
                progress_bar.progress((idx_lieu) / n_lieu, text=f"Generation : {lieu} ({idx_lieu+1}/{n_lieu})")
            
            df_lieu = df[df['Lieu'] == lieu].copy()
            if df_lieu.empty:
                continue
            
            pdf.add_page()
            
            grouped = df_lieu.groupby(['Jours_Norm', 'Horaire_Norm']).apply(format_cell, include_groups=False)
            grid = grouped.unstack(fill_value="") if not grouped.empty else pd.DataFrame()
            
            jours_present = [j for j in [norm(j) for j in jours_ordre] if j in grid.index]
            horaires_present = [h for h in [norm(h) for h in horaires_ordre] if h in grid.columns]
            
            if not jours_present or not horaires_present:
                draw_title(pdf, lieu)
                pdf.set_font("Arial", "", 10)
                pdf.cell(0, 10, "Aucun cours programme pour ce lieu.", 0, 1, "C")
                continue
            
            grid = grid.reindex(index=jours_present, columns=horaires_present)
            grid.index = [map_j.get(i, i) for i in grid.index]
            grid.columns = [map_h.get(c, c) for c in grid.columns]
            
            n_cols = len(grid.columns)
            page_w = pdf.w - 20
            col_jour_w = 22
            col_h_w = (page_w - col_jour_w) / n_cols if n_cols > 0 else page_w
            
            interline = 3.2
            margin_h  = 4.0
            padding_v = 3.0
            
            pdf.set_font("Arial", "", 5.5)
            row_heights = []
            for _, row in grid.iterrows():
                max_lines = 1
                for val in row:
                    if val and str(val).strip():
                        txt_propre = sanitize_for_pdf(str(val))
                        lines = 0
                        for para in txt_propre.split('\n'):
                            w_txt = pdf.get_string_width(para)
                            if w_txt == 0:
                                lines += 1
                            else:
                                usable_w = max(col_h_w - margin_h - 1.0, 10)
                                lines += max(1, math.ceil(w_txt / usable_w))
                        if lines > max_lines:
                            max_lines = lines
                h_needed = max_lines * interline + padding_v * 2 + 2.0
                row_heights.append(max(10, h_needed))
            
            title_h = 8 + 5 + 2
            header_h = 8
            if pdf.get_y() + title_h + header_h + row_heights[0] > pdf.h - MARGE_BAS:
                pdf.add_page()
            
            draw_title(pdf, lieu)
            draw_table_header(pdf, grid, col_jour_w, col_h_w)
            
            for idx, ((jour, row), row_h) in enumerate(zip(grid.iterrows(), row_heights)):
                if pdf.get_y() + row_h > pdf.h - MARGE_BAS:
                    pdf.add_page()
                    draw_title(pdf, lieu)
                    draw_table_header(pdf, grid, col_jour_w, col_h_w)
                
                bg_color = (248, 250, 252) if idx % 2 == 0 else (255, 255, 255)
                
                pdf.set_font("Arial", "B", 7)
                pdf.set_fill_color(*bg_color)
                pdf.set_draw_color(180, 180, 180)
                pdf.cell(col_jour_w, row_h, sanitize_for_pdf(str(jour)), 1, 0, "C", True)
                
                pdf.set_font("Arial", "", 5.5)
                for val in row:
                    cell_text = sanitize_for_pdf(str(val)) if val else ""
                    x, y = pdf.get_x(), pdf.get_y()
                    
                    if cell_text.strip():
                        raw_up = str(val).upper()
                        if "COURS" in raw_up:
                            bg = (225, 238, 255)
                        elif "TD" in raw_up:
                            bg = (232, 252, 235)
                        elif "TP" in raw_up:
                            bg = (255, 235, 235)
                        else:
                            bg = bg_color
                    else:
                        bg = bg_color
                    
                    pdf.set_fill_color(*bg)
                    pdf.set_draw_color(180, 180, 180)
                    pdf.rect(x, y, col_h_w, row_h, 'FD')
                    
                    if cell_text.strip():
                        n_lines = 0
                        for para in cell_text.split('\n'):
                            w_txt = pdf.get_string_width(para)
                            if w_txt == 0:
                                n_lines += 1
                            else:
                                usable_w = max(col_h_w - margin_h - 1.0, 10)
                                n_lines += max(1, math.ceil(w_txt / usable_w))
                        
                        text_block_h = n_lines * interline
                        offset_y = max((row_h - text_block_h) / 2, padding_v / 2)
                        
                        pdf.set_xy(x + margin_h / 2, y + offset_y)
                        pdf.multi_cell(col_h_w - margin_h, interline, cell_text, 0, "L")
                        pdf.set_xy(x + col_h_w, y)
                    else:
                        pdf.set_xy(x + col_h_w, y)
                
                pdf.ln(row_h)
        
        return pdf
    
    # PASSE 1 : comptage silencieux
    pdf_count = _build_pdf(total_pages=0, show_progress=False)
    total_pages = pdf_count.page_no()
    
    # PASSE 2 : generation avec pagination correcte
    if progress_bar is not None:
        progress_bar.progress(0, text=f"Generation finale sur {total_pages} pages...")
    
    pdf_final = _build_pdf(total_pages=total_pages, show_progress=True)
    
    if progress_bar is not None:
        progress_bar.empty()
    
    return bytes(pdf_final.output()), None    
def generate_edt_individuel_lieu_pdf(df_source, nom_lieu):
    """Génère un PDF individuel pour UN lieu (Amphi/Salle) avec grille Jours×Horaires et en-tête PPER.03."""
    try:
        from fpdf import FPDF
        import math
    except ImportError:
        return None, "fpdf non installe"
    
    if df_source is None or df_source.empty:
        return None, "Aucune donnee"
    
    # Dimensions exactes PPER.03
    W_LOGO = 1.19 * 25.4
    W_MILIEU = 3.70 * 25.4
    W_INFO = 1.40 * 25.4
    H_ENTETE = 1.04 * 25.4
    H_HAUT_MILIEU = 0.60 * 25.4
    H_BAS_MILIEU = H_ENTETE - H_HAUT_MILIEU
    W_TOT = W_LOGO + W_MILIEU + W_INFO
    MARGE_BAS = 15
    
    jours_ordre = ["Dimanche", "Lundi", "Mardi", "Mercredi", "Jeudi"]
    horaires_ordre = [
        "8h - 9h", "8h - 9h30", "8h - 10h", "9h - 10h", "9h30 - 11h", 
        "10h - 11h", "11h - 12h", "11h - 12h30", "12h - 13h", 
        "12h30 - 14h", "13h - 14h30", "14h - 15h30", "14h - 16h", "15h30 - 17h"
    ]
    
    def norm(x):
        if not x or str(x).strip().lower() in ["non defini", "nan", "none", ""]:
            return "vide"
        s = str(x).strip().lower().replace(" ", "").replace("-", "").replace("–", "")
        s = s.replace(":00", "").replace("h00", "h")
        return s
    
    map_j = {norm(j): j for j in jours_ordre}
    map_h = {norm(h): h for h in horaires_ordre}
    
    def format_cell(rows):
        items = []
        for _, r in rows.iterrows():
            code_up = str(r.get('Code', '')).upper()
            if 'COURS' in code_up:
                nat = '[C]'
            elif 'TD' in code_up:
                nat = '[TD]'
            else:
                nat = '[TP]'
            txt = f"{nat} {r.get('Enseignements', '')}\nProf: {r.get('Enseignants', '')}\nPromo: {r.get('Promotion', '')}"
            items.append(txt)
        return "\n".join(items)
    
    df = df_source.copy()
    df['Jours_Norm'] = df['Jours'].apply(norm)
    df['Horaire_Norm'] = df['Horaire'].apply(norm)
    
    grouped = df.groupby(['Jours_Norm', 'Horaire_Norm']).apply(format_cell, include_groups=False)
    grid = grouped.unstack(fill_value="") if not grouped.empty else pd.DataFrame()
    
    jours_present = [j for j in [norm(j) for j in jours_ordre] if j in grid.index]
    horaires_present = [h for h in [norm(h) for h in horaires_ordre] if h in grid.columns]
    
    if not jours_present or not horaires_present:
        grid = pd.DataFrame(index=["Aucun"], columns=["Aucun"]).fillna("Aucun cours")
    else:
        grid = grid.reindex(index=jours_present, columns=horaires_present)
        grid.index = [map_j.get(i, i) for i in grid.index]
        grid.columns = [map_h.get(c, c) for c in grid.columns]
    
    class LieuIndivPDF(FPDF):
        def header(self):
            X0 = 10 + ((self.w - 20) - W_TOT) / 2
            Y0 = 10
            X_MILIEU = X0 + W_LOGO
            X_INFO = X_MILIEU + W_MILIEU
            Y_SEP = Y0 + H_HAUT_MILIEU
            
            self.set_draw_color(0, 0, 0)
            self.set_line_width(0.3)
            self.rect(X0, Y0, W_TOT, H_ENTETE, 'D')
            self.line(X_MILIEU, Y0, X_MILIEU, Y0 + H_ENTETE)
            self.line(X_INFO, Y0, X_INFO, Y0 + H_ENTETE)
            self.line(X_MILIEU, Y_SEP, X_INFO, Y_SEP)
            
            if os.path.exists("logo.PNG"):
                self.image("logo.PNG", x=X0 + 2, y=Y0 + 2, w=W_LOGO - 4, h=H_ENTETE - 4)
            
            self.set_xy(X_MILIEU, Y0 + 1.5)
            self.set_font('Arial', 'B', 11)
            self.cell(W_MILIEU, 5.5, sanitize_for_pdf("Universite Djillali Liabes"), 0, 2, "C")
            self.set_font('Arial', '', 10)
            self.cell(W_MILIEU, 5, sanitize_for_pdf("Sidi Bel Abbes"), 0, 2, "C")
            
            self.set_xy(X_MILIEU, Y_SEP + 0.5)
            self.set_font('Arial', 'B', 12)
            self.cell(W_MILIEU, H_BAS_MILIEU - 1, sanitize_for_pdf("EMPLOI DU TEMPS"), 0, 0, "C")
            
            self.set_font('Arial', '', 9)
            line_h = H_ENTETE / 4
            infos = [
                "Code : PPER.03",
                "Revision : 00",
                "Date : 16/05/2026",
                f"Page : {self.page_no()}/{{nb}}"
            ]
            for i, info in enumerate(infos):
                self.set_xy(X_INFO + 1.5, Y0 + 0.5 + i * line_h)
                self.cell(W_INFO - 3, line_h, sanitize_for_pdf(info), 0, 2, "L")
            
            self.set_y(Y0 + H_ENTETE + 5)
        
        def footer(self):
            self.set_y(-15)
            self.set_font('Arial', '', 8)
            self.set_text_color(80, 80, 80)
            self.cell(0, 10, f"{self.page_no()}/{{nb}}", 0, 0, "R")
    
    pdf = LieuIndivPDF(orientation="L", unit="mm", format="A4")
    pdf.alias_nb_pages()
    pdf.set_auto_page_break(auto=True, margin=MARGE_BAS)
    pdf.add_page()
    
    # Sous-titre
    pdf.set_font("Arial", "B", 11)
    pdf.set_text_color(30, 58, 138)
    pdf.cell(0, 8, sanitize_for_pdf(f"PLANNING - {nom_lieu.upper()}"), 0, 1, "C")
    pdf.set_font("Arial", "I", 8)
    pdf.set_text_color(100, 100, 100)
    pdf.cell(0, 5, sanitize_for_pdf("Semestre 01 - Departement d'Electrotechnique - FGE/UDL-SBA"), 0, 1, "C")
    pdf.ln(3)
    
    if grid.empty or (grid.shape == (1,1) and grid.iloc[0,0] == "Aucun cours"):
        pdf.set_font("Arial", "", 10)
        pdf.cell(0, 10, "Aucun cours programme pour ce lieu.", 0, 1, "C")
        return bytes(pdf.output()), None
    
    n_cols = len(grid.columns)
    page_w = pdf.w - 20
    col_jour_w = 22
    col_h_w = (page_w - col_jour_w) / n_cols if n_cols > 0 else page_w
    
    interline = 3.2
    margin_h = 4.0
    padding_v = 3.0
    
    # Calcul des hauteurs de ligne
    pdf.set_font("Arial", "", 5.5)
    row_heights = []
    for _, row in grid.iterrows():
        max_lines = 1
        for val in row:
            if val and str(val).strip():
                txt_propre = sanitize_for_pdf(str(val))
                lines = 0
                for para in txt_propre.split('\n'):
                    w_txt = pdf.get_string_width(para)
                    if w_txt == 0:
                        lines += 1
                    else:
                        usable_w = max(col_h_w - margin_h - 1.0, 10)
                        lines += max(1, math.ceil(w_txt / usable_w))
                if lines > max_lines:
                    max_lines = lines
        h_needed = max_lines * interline + padding_v * 2 + 2.0
        row_heights.append(max(10, h_needed))
    
    # En-têtes
    pdf.set_font("Arial", "B", 7)
    pdf.set_fill_color(30, 58, 138)
    pdf.set_text_color(255, 255, 255)
    pdf.cell(col_jour_w, 8, "JOUR", 1, 0, "C", True)
    for h in grid.columns:
        h_txt = sanitize_for_pdf(str(h))
        if len(h_txt) > 12:
            h_txt = h_txt.replace(" - ", "-").replace(" ", "")
        pdf.cell(col_h_w, 8, h_txt, 1, 0, "C", True)
    pdf.ln()
    
    # Données
    pdf.set_text_color(0, 0, 0)
    
    for idx, ((jour, row), row_h) in enumerate(zip(grid.iterrows(), row_heights)):
        if pdf.get_y() + row_h > pdf.h - MARGE_BAS:
            pdf.add_page()
            # Ré-imprimer l'en-tête du tableau
            pdf.set_font("Arial", "B", 7)
            pdf.set_fill_color(30, 58, 138)
            pdf.set_text_color(255, 255, 255)
            pdf.cell(col_jour_w, 8, "JOUR", 1, 0, "C", True)
            for h in grid.columns:
                h_txt = sanitize_for_pdf(str(h))
                if len(h_txt) > 12:
                    h_txt = h_txt.replace(" - ", "-").replace(" ", "")
                pdf.cell(col_h_w, 8, h_txt, 1, 0, "C", True)
            pdf.ln()
            pdf.set_text_color(0, 0, 0)
        
        bg_color = (248, 250, 252) if idx % 2 == 0 else (255, 255, 255)
        
        pdf.set_font("Arial", "B", 7)
        pdf.set_fill_color(*bg_color)
        pdf.set_draw_color(180, 180, 180)
        pdf.cell(col_jour_w, row_h, sanitize_for_pdf(str(jour)), 1, 0, "C", True)
        
        pdf.set_font("Arial", "", 5.5)
        for val in row:
            cell_text = sanitize_for_pdf(str(val)) if val else ""
            x, y = pdf.get_x(), pdf.get_y()
            
            if cell_text.strip():
                raw_up = str(val).upper()
                if "COURS" in raw_up:
                    bg = (225, 238, 255)
                elif "TD" in raw_up:
                    bg = (232, 252, 235)
                elif "TP" in raw_up:
                    bg = (255, 235, 235)
                else:
                    bg = bg_color
            else:
                bg = bg_color
            
            pdf.set_fill_color(*bg)
            pdf.set_draw_color(180, 180, 180)
            pdf.rect(x, y, col_h_w, row_h, 'FD')
            
            if cell_text.strip():
                n_lines = 0
                for para in cell_text.split('\n'):
                    w_txt = pdf.get_string_width(para)
                    if w_txt == 0:
                        n_lines += 1
                    else:
                        usable_w = max(col_h_w - margin_h - 1.0, 10)
                        n_lines += max(1, math.ceil(w_txt / usable_w))
                
                text_block_h = n_lines * interline
                offset_y = max((row_h - text_block_h) / 2, padding_v / 2)
                
                pdf.set_xy(x + margin_h / 2, y + offset_y)
                pdf.multi_cell(col_h_w - margin_h, interline, cell_text, 0, "L")
                pdf.set_xy(x + col_h_w, y)
            else:
                pdf.set_xy(x + col_h_w, y)
        pdf.ln(row_h)
    
    return bytes(pdf.output()), None
    # ═══════════════════════════════════════════════════════════════
    # PASSE 2 : GENERATION FINALE AVEC BONNE PAGINATION
    # ═══════════════════════════════════════════════════════════════
    if progress_bar is not None:
        progress_bar.progress(0, text=f"Generation finale sur {total_pages} pages...")
    
    pdf_final = _build_pdf(total_pages=total_pages, show_progress=True)
    
    if progress_bar is not None:
        progress_bar.empty()
    
    return bytes(pdf_final.output()), None
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

    # ═══════════════════════════════════════════════════════
    # BARRE DE SESSION : UTILISATEUR + DÉCONNEXION
    # ═══════════════════════════════════════════════════════
    if user_data is not None:
        nom_session = user_data.get('nom_officiel', 'Utilisateur')
        role_session = user_data.get('role', 'enseignant').upper()
        col_user, col_deco = st.columns([4, 1])
        with col_user:
            st.markdown(f"""
                <div style="background: linear-gradient(90deg, #f0f9ff 0%, #e0f2fe 100%); 
                            border-left: 4px solid #1E3A8A; padding: 10px 16px; border-radius: 8px; 
                            margin-bottom: 12px; display: flex; align-items: center; gap: 10px;">
                    <span style="font-size: 20px;">👤</span>
                    <div>
                        <div style="font-weight: bold; color: #1E3A8A; font-size: 14px;">{nom_session}</div>
                        <div style="font-size: 11px; color: #64748b; text-transform: uppercase; letter-spacing: 0.5px;">
                            Connecté · {role_session}
                        </div>
                    </div>
                </div>
            """, unsafe_allow_html=True)
        with col_deco:
            if st.button("🚪 Déconnexion", use_container_width=True, type="primary"):
                st.session_state["user_data"] = None
                st.session_state.pop("pdf_all_ready", None)
                st.session_state.pop("pdf_all_data", None)
                st.session_state.pop("pdf_all_promo_ready", None)
                st.session_state.pop("pdf_all_promo_data", None)
                st.session_state.pop("pdf_all_lieu_ready", None)
                st.session_state.pop("pdf_all_lieu_data", None)
                st.rerun()

    if df_global is None or df_global.empty:
        st.warning("Aucune donnee chargee. Verifiez votre connexion Supabase ou votre fichier Excel.")
        return

    # Nettoyage : suppression des colonnes techniques internes pour tous les exports
    COLONNES_CACHEES = ['h_norm', 'j_norm']
    df_propre = df_global.drop(columns=[c for c in COLONNES_CACHEES if c in df_global.columns], errors='ignore')

    promos = sorted([p for p in df_propre["Promotion"].unique() if p and p != "Non defini"])
    profs = sorted([p for p in df_propre["Enseignants"].unique() if p and p != "Non defini"])
    salles = sorted([s for s in df_propre["Lieu"].unique() if s and s != "Non defini"])

    col1, col2, col3 = st.columns(3)

    with col1:
    
        st.markdown("**🎓 Par Promotion**")
        sel_promo = st.selectbox("Choisir promotion", ["Toutes"] + promos, key="hub_promo")
        df_filtre = df_propre.copy()
        if sel_promo != "Toutes":
            df_filtre = df_filtre[df_filtre["Promotion"] == sel_promo]
        c1, c2, c3 = st.columns(3)
        
        # ═══════════════════════════════════════════════════════
        # PDF : individuel ou global (meme logique que Enseignants)
        # ═══════════════════════════════════════════════════════
        if sel_promo != "Toutes":
            # Une seule promotion → generation immediate
            pdf_data, err = generate_pro_pdf(df_filtre, f"EDT - {sel_promo}", "Export promotion")
            if pdf_data is not None:
                c1.download_button("📄 PDF", pdf_data, f"EDT_{sel_promo}_2027.pdf", "application/pdf", use_container_width=True, key="dp_promo_single")
            else:
                c1.button("📄 PDF", disabled=True, use_container_width=True, key="dp_promo_single")
        else:
            # Toutes les promotions → generation au clic avec progression
            if c1.button("📄 Générer PDF Global", use_container_width=True, key="btn_gen_all_pdf_promo"):
                with st.spinner("Preparation du fichier global..."):
                    prog = st.progress(0, text="Demarrage...")
                    pdf_data_all, err_all = generate_edt_toutes_promotions_pdf(df_propre, progress_bar=prog)
                    if pdf_data_all:
                        st.session_state['pdf_all_promo_data'] = pdf_data_all
                        st.session_state['pdf_all_promo_ready'] = True
                        st.success(f"✅ PDF Généré : {len(promos)} promotions")
                        st.rerun()
                    else:
                        st.error(f"❌ Erreur : {err_all}")
            
            if st.session_state.get('pdf_all_promo_ready') and 'pdf_all_promo_data' in st.session_state:
                c1.download_button("⬇️ Telecharger PDF Global", st.session_state['pdf_all_promo_data'],
                                  "EDT_Toutes_Promotions_2027.pdf", "application/pdf",
                                  use_container_width=True, key="dp_down_promo")
        
        # HTML et Excel (toujours disponibles)
        html_data = generate_pro_html(df_filtre, f"EDT {sel_promo}", "Faculte de Genie Electrique - UDL-SBA")
        c2.download_button("🌐 HTML", html_data, f"EDT_{sel_promo}_2027.html", "text/html", use_container_width=True, key="dh_promo")
        xlsx_data = generate_pro_excel(df_filtre, f"EDT {sel_promo}")
        c3.download_button("📊 Excel", xlsx_data, f"EDT_{sel_promo}_2027.xlsx", 
                          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, key="dx_promo")
    
    
    with col2:
        st.markdown("**👤 Par Enseignant**")
        sel_prof = st.selectbox("Choisir enseignant", ["Tous"] + profs, key="hub_prof")
        df_filtre_p = df_propre.copy()
        if sel_prof != "Tous":
            df_filtre_p = df_filtre_p[df_filtre_p["Enseignants"].str.contains(sel_prof, case=False, na=False)]
            
            # ═══════════════════════════════════════════════════════
            # AFFICHEUR NUMÉRIQUE : CHARGE HORAIRE & HEURES SUP
            # ═══════════════════════════════════════════════════════
            df_filtre_p['Type'] = df_filtre_p['Code'].apply(
                lambda x: "COURS" if "COURS" in str(x).upper() else ("TD" if "TD" in str(x).upper() else "TP")
            )
            # Déduplication sur Horaire + Jours pour compter les séances uniques
            df_u = df_filtre_p.drop_duplicates(subset=['Horaire', 'Jours'])
            nb_cours = len(df_u[df_u['Type'] == 'COURS'])
            nb_td    = len(df_u[df_u['Type'] == 'TD'])
            nb_tp    = len(df_u[df_u['Type'] == 'TP'])

            # Calcul de la charge équivalente
            # 1h Cours = 1.5 eq/h  |  1h TD/TP = 1.0 eq/h
            charge_eq = round((nb_cours * 1.5) + (nb_td * 1.0) + (nb_tp * 1.0), 2)
            SEUIL_REGLEMENTAIRE = 6.0  # 6 heures équivalent/semaine

            delta = round(charge_eq - SEUIL_REGLEMENTAIRE, 2)

            # Style conditionnel
            if delta > 0:
                statut_label = "Heures Supplémentaires"
                statut_color = "#22c55e"  # vert
                statut_bg = "#f0fdf4"
                delta_str = f"+{delta} eq/h"
                emoji = "✅"
            elif delta < 0:
                statut_label = "Déficit Horaire"
                statut_color = "#ef4444"  # rouge
                statut_bg = "#fef2f2"
                delta_str = f"{delta} eq/h"
                emoji = "⚠️"
            else:
                statut_label = "Charge Exacte"
                statut_color = "#3b82f6"  # bleu
                statut_bg = "#eff6ff"
                delta_str = "0 eq/h"
                emoji = "⚖️"

            st.markdown(f"""
            <div style="background: linear-gradient(135deg, #1E3A8A 0%, #3B82F6 100%); 
                        padding: 16px; border-radius: 12px; color: white; margin: 10px 0;
                        box-shadow: 0 4px 12px rgba(0,0,0,0.1);">
                <div style="text-align: center; font-size: 13px; opacity: 0.9; text-transform: uppercase; letter-spacing: 1px;">
                    📊 Bilan Hebdomadaire — {sel_prof}
                </div>
                <div style="display: flex; justify-content: center; gap: 20px; margin-top: 12px; flex-wrap: wrap;">
                    <div style="text-align: center; min-width: 80px;">
                        <div style="font-size: 24px; font-weight: bold;">{nb_cours}</div>
                        <div style="font-size: 11px; opacity: 0.8;">📘 Cours</div>
                    </div>
                    <div style="text-align: center; min-width: 80px;">
                        <div style="font-size: 24px; font-weight: bold;">{nb_td}</div>
                        <div style="font-size: 11px; opacity: 0.8;">📗 TD</div>
                    </div>
                    <div style="text-align: center; min-width: 80px;">
                        <div style="font-size: 24px; font-weight: bold;">{nb_tp}</div>
                        <div style="font-size: 11px; opacity: 0.8;">🔴 TP</div>
                    </div>
                    <div style="border-left: 1px solid rgba(255,255,255,0.3); height: 40px; align-self: center;"></div>
                    <div style="text-align: center; min-width: 100px;">
                        <div style="font-size: 28px; font-weight: bold;">{charge_eq}</div>
                        <div style="font-size: 11px; opacity: 0.8;">Charge Eq/h</div>
                    </div>
                    <div style="text-align: center; min-width: 100px;">
                        <div style="font-size: 28px; font-weight: bold;">{SEUIL_REGLEMENTAIRE}</div>
                        <div style="font-size: 11px; opacity: 0.8;">Seuil Réglem.</div>
                    </div>
                </div>
            </div>
            <div style="background-color: {statut_bg}; border-left: 5px solid {statut_color}; 
                        padding: 12px 16px; border-radius: 8px; margin-top: 8px;">
                <div style="font-size: 18px; font-weight: bold; color: {statut_color};">
                    {emoji} {statut_label} : {delta_str}
                </div>
                <div style="font-size: 12px; color: #64748b; margin-top: 4px;">
                    Règle : 1h Cours = 1.5 eq/h | 1h TD/TP = 1.0 eq/h | Seuil = 6.0 eq/h/semaine
                </div>
            </div>
            """, unsafe_allow_html=True)
            # ═══════════════════════════════════════════════════════

              
        # ═══════════════════════════════════════════════════════
        # PDF : individuel ou global
        # ═══════════════════════════════════════════════════════
        if sel_prof != "Tous":
            # Un seul enseignant → generation immediate
            pdf_data_p, _ = generate_edt_individuel_pdf_classique(df_filtre_p, sel_prof)
            if pdf_data_p is not None:
                c1.download_button("📄 PDF", pdf_data_p, f"EDT_{sel_prof}_2027.pdf", 
                                  "application/pdf", use_container_width=True, key="dp")
            else:
                c1.button("📄 PDF", disabled=True, use_container_width=True, key="dp")
        else:
            # Tous les enseignants → generation au clic avec progression
            if c1.button("📄 Générer PDF Global", use_container_width=True, key="btn_gen_all_pdf"):
                with st.spinner("Preparation du fichier global..."):
                    prog = st.progress(0, text="Demarrage...")
                    pdf_data_all, err_all = generate_edt_tous_enseignants_pdf(df_propre, progress_bar=prog)
                    if pdf_data_all:
                        st.session_state['pdf_all_data'] = pdf_data_all
                        st.session_state['pdf_all_ready'] = True
                        st.success(f"✅ PDF Généré : {len(profs)} enseignants")
                        st.rerun()
                    else:
                        st.error(f"❌ Erreur : {err_all}")
            
            if st.session_state.get('pdf_all_ready') and 'pdf_all_data' in st.session_state:
                c1.download_button("⬇️ Telecharger PDF Global", st.session_state['pdf_all_data'],
                                  "EDT_Tous_Enseignants_2027.pdf", "application/pdf",
                                  use_container_width=True, key="dp_down")
        
        # HTML et Excel (toujours disponibles)
        html_data_p = generate_pro_html(df_filtre_p, f"EDT {sel_prof}", "Faculte de Genie Electrique - UDL-SBA")
        if html_data_p:
            c2.download_button("🌐 HTML", html_data_p, f"EDT_{sel_prof}_2027.html", 
                              "text/html", use_container_width=True, key="dh")
        else:
            c2.button("🌐 HTML", disabled=True, use_container_width=True, key="dh")
            
        xlsx_data_p = generate_pro_excel(df_filtre_p, f"EDT {sel_prof}")
        if xlsx_data_p:
            c3.download_button("📊 Excel", xlsx_data_p, f"EDT_{sel_prof}_2027.xlsx", 
                              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
                              use_container_width=True, key="dx")
        else:
            c3.button("📊 Excel", disabled=True, use_container_width=True, key="dx")                
            
    
    with col3:
        st.markdown("**🏢 Par Lieu**")
        sel_salle = st.selectbox("Choisir lieu (Salle, Amphi, Labo, Autres)", ["Toutes"] + salles, key="hub_salle")
        df_filtre_s = df_propre.copy()
        
        if sel_salle != "Toutes":
            df_filtre_s = df_filtre_s[df_filtre_s["Lieu"].astype(str).str.startswith(sel_salle)]
        
        c1, c2, c3 = st.columns(3)
        
        # ═══════════════════════════════════════════════════════
        # PDF : individuel ou global (meme logique que Promotions)
        # ═══════════════════════════════════════════════════════
        if sel_salle != "Toutes":
            # Un seul lieu → generation immediate avec grille standard PPER.03
            pdf_data_s, _ = generate_edt_individuel_lieu_pdf(df_filtre_s, sel_salle)
            if pdf_data_s is not None:
                c1.download_button("📄 PDF", pdf_data_s, f"Planning_{sel_salle}_2027.pdf", "application/pdf", use_container_width=True, key="sp")
            else:
                c1.button("📄 PDF", disabled=True, use_container_width=True, key="sp")
        else:
            # Tous les lieux → generation au clic avec progression
            if c1.button("📄 Generer PDF Global", use_container_width=True, key="btn_gen_all_pdf_lieu"):
                with st.spinner("Preparation du fichier global..."):
                    prog = st.progress(0, text="Demarrage...")
                    pdf_data_all, err_all = generate_edt_tous_lieux_pdf(df_propre, progress_bar=prog)
                    if pdf_data_all:
                        st.session_state['pdf_all_lieu_data'] = pdf_data_all
                        st.session_state['pdf_all_lieu_ready'] = True
                        st.success(f"✅ PDF Genere : {len(salles)} lieux")
                        st.rerun()
                    else:
                        st.error(f"❌ Erreur : {err_all}")
            
            if st.session_state.get('pdf_all_lieu_ready') and 'pdf_all_lieu_data' in st.session_state:
                c1.download_button("⬇️ Telecharger PDF Global", st.session_state['pdf_all_lieu_data'],
                                  "Planning_Tous_Lieux_2027.pdf", "application/pdf",
                                  use_container_width=True, key="dp_down_lieu")
        
        # HTML et Excel (toujours disponibles)
        html_data_s = generate_pro_html(df_filtre_s, f"Planning {sel_salle}", "Faculte de Genie Electrique - UDL-SBA")
        c2.download_button("🌐 HTML", html_data_s, f"Planning_{sel_salle}_2027.html", "text/html", use_container_width=True, key="sh")
        xlsx_data_s = generate_pro_excel(df_filtre_s, f"Planning {sel_salle}")
        c3.download_button("📊 Excel", xlsx_data_s, f"Planning_{sel_salle}_2027.xlsx", 
                          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, key="sx")
    
    if is_admin:
        st.divider()
        st.markdown("**🌍 Export Global (Admin)**")
        cg1, cg2, cg3, cg4 = st.columns(4)
        pdf_g, _ = generate_pro_pdf(df_propre, "EDT GLOBAL S1-2027", "Département d'Electrotechnique - Toutes promotions")
        if pdf_g is not None:
            cg1.download_button("📄 PDF Global", pdf_g, "EDT_GLOBAL_S1_2027.pdf", "application/pdf", use_container_width=True)
        html_g = generate_pro_html(df_propre, "EDT Global S1-2027", "Département d'Electrotechnique - FGE/UDL-SBA")
        cg2.download_button("🌐 HTML Global", html_g, "EDT_GLOBAL_S1_2027.html", "text/html", use_container_width=True)
        xlsx_g = generate_pro_excel(df_propre, "EDT Global S1-2027", "EDT_Global")
        cg3.download_button("📊 Excel Global", xlsx_g, "EDT_GLOBAL_S1_2027.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            if pdf_g is not None:
                zf.writestr("EDT_GLOBAL.pdf", pdf_g)
            zf.writestr("EDT_GLOBAL.html", html_g)
            zf.writestr("EDT_GLOBAL.xlsx", xlsx_g)
        cg4.download_button("🗜️ Pack ZIP", zip_buffer.getvalue(), "Pack_EDT_GLOBAL_S1_2027.zip", "application/zip", use_container_width=True)

    st.divider()

# =============================================================================
