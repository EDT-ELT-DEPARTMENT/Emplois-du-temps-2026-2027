"""
================================================================================
Application Unifiée : Suivi d'Assiduité + Gestion des EDTs
Département d'Electrotechnique - Faculté de Genie Electrique - UDL-SBA
Année universitaire 2026-2027
================================================================================
Module 1 : Suivi d'Assiduite des Etudiants
  - Enregistrement des absences
  - Gestion des justificatifs
  - Bilans et exports (Excel/HTML)

Module 2 : Plateforme de Gestion des EDTs
  - Emplois du temps par promotion/enseignant/salle
  - Surveillances d'examens
  - Générateur automatique
  - Portail enseignants (envoi mail)
  - Gestion administrative (Bordereaux & PVs)
  - Éditeur de données avec détection de conflits

Fichiers sources requis (même dossier que le script) :
  1. Liste des étudiants_2026-2027.xlsx
  2. dataEDT-ELT-S1-2027.xlsx
  3. Permanents-Vacataires-ELT2-2026-2027.xlsx
  4. surveillances_2027.xlsx (optionnel)
  5. logo.PNG (optionnel)

Configuration Supabase (optionnel) :
  Créer .streamlit/secrets.toml avec SUPABASE_URL et SUPABASE_KEY
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

# Masquer les éléments du menu supérieur (style du module EDT)
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
# Chemin absolu vers le dossier du script (compatible Streamlit Cloud)
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
# MODULE 1 : SUIVI ASSIDUITE DES ETUDIANTS
# =============================================================================
def run_assiduite():
  

    import streamlit as st
    import pandas as pd
    import base64
    import io
    import time
    import re
    import os
    from datetime import datetime
    from collections import defaultdict
     
    # =============================================================================
    # CONFIGURATION SUPABASE
    # =============================================================================
    MODE_SUPABASE = False
    supabase = None

    try:
        from supabase import create_client, Client
        SUPABASE_URL = st.secrets.get("SUPABASE_URL", "")
        SUPABASE_KEY = st.secrets.get("SUPABASE_KEY", "")
        if SUPABASE_URL and SUPABASE_KEY:
            supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)
            MODE_SUPABASE = True
            st.sidebar.success("☁️ Supabase connecte")
        else:
            st.sidebar.info("💾 Mode local (Supabase non configure)")
    except Exception as e:
        st.sidebar.warning(f"⚠️ Mode local : {e}")

    # =============================================================================
    # CONFIGURATION FICHIERS
    # =============================================================================
    FILE_ETUDIANTS = "Liste des étudiants_2026-2027.xlsx"
    FILE_EDT       = "dataEDT-ELT-S1-2027.xlsx"
    FILE_ENS       = "Permanents-Vacataires-ELT2-2026-2027.xlsx"
        # =============================================================================
    # CONFIGURATION FICHIERS (CORRIGÉE — utilise _BASE_DIR comme le module EDT)
    # =============================================================================
    FILE_ETUDIANTS = str(_BASE_DIR / "Liste des étudiants_2026-2027.xlsx")
    FILE_EDT       = str(_BASE_DIR / "dataEDT-ELT-S1-2027.xlsx")
    FILE_ENS       = str(_BASE_DIR / "Permanents-Vacataires-ELT2-2026-2027.xlsx")
    
    # =============================================================================
    # FONCTIONS UTILITAIRES (CORRIGÉE — avec upload fallback)
    # =============================================================================
    
    @st.cache_data(show_spinner=False)
    def charger_donnees():
        """Charge les 3 fichiers Excel sources. Permet l'upload manuel si fichiers locaux manquants."""
        
        fichiers = {
            "étudiants": FILE_ETUDIANTS,
            "EDT": FILE_EDT,
            "enseignants": FILE_ENS
        }
        
        df_etu, df_edt, df_ens = pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
        manquants = []
        
        # Vérification des fichiers locaux
        for nom, chemin in fichiers.items():
            if not os.path.exists(chemin):
                manquants.append(nom)
        
        # Si des fichiers sont manquants, proposer l'upload
        if manquants:
            st.warning(f"⚠️ Fichiers locaux manquants : {', '.join(manquants)}")
            st.info("📤 Veuillez uploader les fichiers Excel manquants ci-dessous pour continuer :")
            
            uploaded = {}
            if "étudiants" in manquants:
                uploaded["étudiants"] = st.file_uploader("Liste des étudiants (.xlsx)", type=["xlsx"], key="up_etu")
            if "EDT" in manquants:
                uploaded["EDT"] = st.file_uploader("Données EDT (.xlsx)", type=["xlsx"], key="up_edt")
            if "enseignants" in manquants:
                uploaded["enseignants"] = st.file_uploader("Liste enseignants (.xlsx)", type=["xlsx"], key="up_ens")
            
            # Vérifier que tous les uploads sont présents
            if not all(v is not None for v in uploaded.values()):
                st.error("❌ Tous les fichiers sont requis pour continuer.")
                st.stop()
            
            # Lecture depuis les uploads
            try:
                df_etu = pd.read_excel(uploaded["étudiants"])
                df_etu.columns = df_etu.columns.str.strip()
            except Exception as e:
                st.error(f"❌ Erreur lecture étudiants (upload) : {e}")
                
            try:
                df_edt = pd.read_excel(uploaded["EDT"])
                df_edt.columns = df_edt.columns.str.strip()
            except Exception as e:
                st.error(f"❌ Erreur lecture EDT (upload) : {e}")
                
            try:
                df_ens = pd.read_excel(uploaded["enseignants"], sheet_name=0)
                df_ens.columns = df_ens.columns.str.strip()
            except Exception as e:
                st.error(f"❌ Erreur lecture enseignants (upload) : {e}")
                
        else:
            # Chargement normal depuis les fichiers locaux
            try:
                df_etu = pd.read_excel(FILE_ETUDIANTS)
                df_etu.columns = df_etu.columns.str.strip()
            except Exception as e:
                st.error(f"❌ Erreur chargement étudiants : {e}")
    
            try:
                df_edt = pd.read_excel(FILE_EDT)
                df_edt.columns = df_edt.columns.str.strip()
            except Exception as e:
                st.error(f"❌ Erreur chargement EDT : {e}")
    
            try:
                df_ens = pd.read_excel(FILE_ENS, sheet_name=0)
                df_ens.columns = df_ens.columns.str.strip()
            except Exception as e:
                st.error(f"❌ Erreur chargement enseignants : {e}")
    
        return df_etu, df_edt, df_ens

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
    # FONCTIONS SUPABASE
    # =============================================================================

    def charger_absences_supabase(matiere=None, promotion=None):
        """Charge les absences depuis Supabase."""
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
        """Enregistre une absence dans Supabase."""
        if not MODE_SUPABASE:
            return False
        try:
            supabase.table("suivi_assiduite_2026").insert(payload).execute()
            return True
        except Exception as e:
            st.error(f"Erreur Supabase (enregistrer) : {e}")
            return False


    def supprimer_absences_supabase(matiere, promotion):
        """Supprime les absences d'une matiere/promotion."""
        if not MODE_SUPABASE:
            return False
        try:
            supabase.table("suivi_assiduite_2026").delete().eq("matiere", matiere).eq("promotion", promotion).execute()
            return True
        except Exception as e:
            st.error(f"Erreur Supabase (supprimer) : {e}")
            return False
   
    def rehabiliter_absences_etudiant_supabase(etudiant, matiere):
        """Marque les absences d'un etudiant comme justifiees (rehabilitation conserve l'historique)."""
        if not MODE_SUPABASE:
            return False
        try:
            supabase.table("suivi_assiduite_2026").update({"justifie": True}).eq("etud_non_eligible", etudiant).eq("matiere", matiere).execute()
            return True
        except Exception as e:
            st.error(f"Erreur Supabase (rehabilitation) : {e}")
            return False    


    def charger_requetes_supabase(statut=None, promotion=None):
        """Charge les requetes de justificatifs."""
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
        """Enregistre une requete de justificatif."""
        if not MODE_SUPABASE:
            return False
        try:
            supabase.table("requetes_absences").insert(payload).execute()
            return True
        except Exception as e:
            st.error(f"Erreur Supabase (requete) : {e}")
            return False


    def mettre_a_jour_statut_requete_supabase(req_id, statut):
        """Met a jour le statut d'une requete."""
        if not MODE_SUPABASE:
            return False
        try:
            supabase.table("requetes_absences").update({"statut": statut}).eq("id", req_id).execute()
            return True
        except Exception as e:
            st.error(f"Erreur Supabase (maj statut) : {e}")
            return False


    def reinitialiser_requetes_supabase():
        """Reinitialise toutes les requetes."""
        if not MODE_SUPABASE:
            return False
        try:
            supabase.table("requetes_absences").delete().neq("id", -1).execute()
            return True
        except Exception as e:
            st.error(f"Erreur Supabase (reset) : {e}")
            return False


    # =============================================================================
    # FONCTIONS UTILITAIRES
    # ============================================================================

    def nettoyer_nom_enseignant(nom):
        """Retire les titres academiques."""
        n = str(nom).strip()
        for prefix in ["Pr ", "Dr ", "Mme ", "Mr ", "Dr. ", "Pr. ", "M. "]:
            if n.startswith(prefix):
                n = n[len(prefix):]
        return n.strip()


    def extraire_nom_famille(nom_complet):
        """Extrait le premier mot (nom de famille)."""
        n = nettoyer_nom_enseignant(nom_complet)
        parts = n.split()
        if not parts:
            return ""
        return parts[0].upper()


    def mapper_promotion(promo_edt):
        """Mappe une promotion EDT vers celle du fichier etudiants."""
        p = str(promo_edt).strip().upper()

        mapping_direct = {
            "ING1": "ING1", "ING2ST": "ING2", "ING2TM": "ING2",
            "ING3EI": "ING3EI", "ING3RSE": "ING3RSE", "ING3TM": "ING3EI",
            "ING4EI": "ING4", "ING4": "ING4", "ING5RSE": "ING3RSE",
            "L1MCIL": "L1MCIL", "L2ELT": "L2ELT", "L2MCIL": "L2MCIL",
            "L3ELT": "L3ELT", "MCIL2": "L2MCIL", "MCIL3": "MCIL3",
            "M1CE": "M1CE", "M1ER": "M1ER", "M1MCIL": "M1MCIL",
            "M1ME": "M1ME", "M1RE": "M1RE", "M2CE": "M2CE",
            "M2ER": "M2ER", "M2MCIL": "M2MCIL", "M2ME": "M2ME", "M2RE": "M1RE",
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
        elif "L2" in p and "MCIL" in p: return "L2MCIL"
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
        """Retourne un DataFrame filtre sur l'enseignant."""
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
        """Genere une page HTML stylisee."""
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
    # INTERFACE STREAMLIT
    # =============================================================================
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
        """Trouve une requête en attente pour cet étudiant et cette matière."""
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
    
        for r in st.session_state.requetes:
            if (r.get("nom_etudiant") == nom_etudiant and 
                r.get("matiere") == matiere and 
                r.get("statut") == "En attente"):
                return r
        return None

    st.title("📊 Plateforme de Suivi d'Assiduite des Etudiants")
    st.caption("Departement d'Electrotechnique - Faculte de Genie Electrique - UDL-SBA - Annee 2026-2027")

    # --- Chargement des donnees ---
    df_etu, df_edt, df_ens = charger_donnees()

    if df_etu.empty or df_edt.empty or df_ens.empty:
        st.error("❌ Un ou plusieurs fichiers sources sont manquants. Verifiez que les 3 fichiers .xlsx sont dans le meme dossier que ce script.")
        st.stop()

    # --- Preparation des listes ---
    if "NOM" in df_ens.columns and "PRÉNOM" in df_ens.columns:
        df_ens["Nom_Complet"] = df_ens["NOM"].astype(str).str.strip().str.upper() + " " + df_ens["PRÉNOM"].astype(str).str.strip().str.title()
        LISTE_PROFS = sorted(df_ens["Nom_Complet"].dropna().unique().tolist())
    elif "Nom" in df_ens.columns and "Prénom" in df_ens.columns:
        df_ens["Nom_Complet"] = df_ens["Nom"].astype(str).str.strip().str.upper() + " " + df_ens["Prénom"].astype(str).str.strip().str.title()
        LISTE_PROFS = sorted(df_ens["Nom_Complet"].dropna().unique().tolist())
    else:
        LISTE_PROFS = []

    df_etu["Nom_Complet"] = df_etu["Nom"].astype(str).str.strip().str.upper() + " " + df_etu["Prénom"].astype(str).str.strip().str.title()

    # =============================================================================
    # ONGLETS
    # =============================================================================
    if is_enseignant_connecte:
        tab1, = st.tabs(["📝 Suivi d'Assiduite"])
    else:
        tab1, tab2, tab3 = st.tabs(["📝 Suivi d'Assiduite", "📩 Justificatifs", "📊 Bilans & Exports"])


    # =============================================================================
    # ONGLET 1 : SUIVI D'ASSIDUITE
    # =============================================================================
        # =============================================================================
    # GESTION DE LA CONNEXION (depuis le portail EDT)
    # =============================================================================
    user = st.session_state.get("user_data")
    is_enseignant_connecte = user is not None and user.get("role") != "admin"

    # =============================================================================
    # ONGLETS
    # =============================================================================
    if is_enseignant_connecte:
        tab1 = st.tabs(["📝 Suivi d'Assiduite"])[0]
    else:
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
            # >>> MODE ENSEIGNANT : accès direct, pas de code
            sel_prof = user['nom_officiel']
            st.success(f"👤 Bienvenue **{sel_prof}** — Espace Suivi d'Assiduité")
            c1, c2 = st.columns(2)
            with c1:
                st.markdown(f"**Enseignant :** `{sel_prof}`")
            with c2:
                st.markdown("*Accès direct — Aucun code requis*")
        else:
            # >>> MODE ADMIN : code requis
            pwd = st.text_input("🔑 Code d'acces :", type="password", key="pwd_tab1")

            if pwd == CODE_ADMIN:
                c1, c2 = st.columns(2)
                with c1:
                    sel_prof = st.selectbox("👤 Selectionnez l'Enseignant :", [""] + LISTE_PROFS, key="ens_T1")
            elif pwd != "":
                st.error("❌ Code incorrect.")

        # >>> SUITE COMMUNE (enseignant connecté OU admin validé)
        if sel_prof:
            df_matiere = trouver_matiere_promo(sel_prof, df_edt)
            if not df_matiere.empty:
                liste_mats = sorted(df_matiere["Enseignements"].dropna().unique().tolist())
                with c2:
                    sel_mat = st.selectbox("📚 Selectionnez la Matiere :", [""] + liste_mats, key="mat_T1")

                if sel_mat:
                    info_rows = df_matiere[df_matiere["Enseignements"] == sel_mat]
                    if not info_rows.empty:
                        promo_c = str(info_rows.iloc[0]["Promotion_Mappee"]).strip()

            if sel_mat and promo_c:
                      
                df_p = df_etu[df_etu["Promotion"].astype(str).str.strip().str.upper() == promo_c.upper()].copy()

                if not df_p.empty:
                    noms_e = sorted(df_p["Nom_Complet"].tolist())
                    st.info(f"📍 Promotion detectee : **{promo_c}** | **{len(noms_e)}** etudiants")

                    # Recuperation des absences (Supabase ou local)
                    if MODE_SUPABASE:
                        absences_filtrees = charger_absences_supabase(sel_mat, promo_c)
                    else:
                        absences_filtrees = [
                            a for a in st.session_state.absences
                            if a.get("matiere") == sel_mat and a.get("promotion") == promo_c
                        ]
                    df_db_full = pd.DataFrame(absences_filtrees)

                    st.markdown("#### 📥 Enregistrement d'une Absence")
                    cn1, cn2, cn3 = st.columns(3)

                    with cn1:
                        etud_non = st.selectbox("👤 Etudiant :", [""] + noms_e, key="ne_et_t1")
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

                    # Compteur
                    # Compteur par matière (seuil d'exclusion = 5) + global
                    if etud_non and status_assid == "Absent":
                        if not df_db_full.empty and "etud_non_eligible" in df_db_full.columns:
                            absences_etu_matiere = df_db_full[df_db_full["etud_non_eligible"] == etud_non]
                            nb_abs_matiere = len(absences_etu_matiere)
                            nb_abs_justif = len(absences_etu_matiere[absences_etu_matiere.get("justifie") == True]) if "justifie" in absences_etu_matiere.columns else 0
                        else:
                            nb_abs_matiere = 0
                            nb_abs_justif = 0

                        # Comptage global (toutes matières)
                        if MODE_SUPABASE:
                            abs_global = [a for a in charger_absences_supabase() if a.get("etud_non_eligible") == etud_non]
                        else:
                            abs_global = [a for a in st.session_state.absences if a.get("etud_non_eligible") == etud_non]
                        nb_abs_global = len(abs_global)

                        c1, c2, c3 = st.columns(3)
                        with c1:
                            st.metric("🔢 Absences dans cette matière", f"{nb_abs_matiere}/5")
                        with c2:
                            st.metric("✅ Dont justifiées", f"{nb_abs_justif}")
                        with c3:
                            st.metric("🌍 Total global (toutes matières)", f"{nb_abs_global}")

                        # Alerte seuil d'exclusion
                        if nb_abs_matiere >= 5:
                            st.error(f"🚫 **EXCLU de la matière {sel_mat}** — Seuil de 5 absences atteint (justifiées ou non).")
                        elif nb_abs_matiere == 4:
                            st.warning(f"⚠️ Attention : 4 absences dans {sel_mat}. Une prochaine absence = exclusion.")
                        else:
                            st.info(f"ℹ️ {nb_abs_matiere} absence(s) dans {sel_mat}. Seuil d'exclusion : 5.")
                    if st.button("💾 ENREGISTRER L'ABSENCE", use_container_width=True):
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

                    # Liste globale des absences
                    # LISTE GLOBALE DES ABSENCES (HISTORIQUE COMPLET)
                    # ─────────────────────────────────────────────────────────────
                    # Liste globale des absences
                    st.divider()
                    st.subheader("📋 Liste Globale des Absences")

                    if not df_db_full.empty and "etud_non_eligible" in df_db_full.columns:
                        # Ajout colonne Justifiée si elle n'existe pas encore (compatibilité anciennes données)
                        if "justifie" not in df_db_full.columns:
                            df_db_full["justifie"] = False

                        # Compteur par étudiant × matière (pour détecter l'exclusion)
                        df_db_full["justifie"] = df_db_full["justifie"].fillna(False)
                        df_liste = df_db_full.copy()

                        # Statut justification
                        df_liste["Statut Justif"] = df_liste["justifie"].apply(lambda x: "✅ Justifiée" if x else "❌ Non justifiée")

                        # Comptage par matière (pour l'exclusion)
                        df_count_mat = df_liste.groupby(["etud_non_eligible", "matiere"]).size().reset_index(name="Abs Matiere")
                        df_liste = df_liste.merge(df_count_mat, on=["etud_non_eligible", "matiere"], how="left")

                        # Statut exclusion
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
                        # Bouton pour effacer l'historique global (Admin)
                        if st.button("🗑️ Effacer TOUT l'historique des absences", type="primary"):
                            if MODE_SUPABASE:
                                # Suppression totale (attention : irréversible)
                                try:
                                    # Récupération des absences pour cette MATIERE et cette PROMOTION
                                    if MODE_SUPABASE:
                                        absences_cours = charger_absences_supabase(sel_mat, promo_c)
                                    else:
                                        absences_cours = [
                                            a for a in st.session_state.absences
                                            if a.get("matiere") == sel_mat and a.get("promotion") == promo_c
                                        ]
                                    df_abs_cours = pd.DataFrame(absences_cours)
                                    if "justifie" not in df_abs_cours.columns:
                                        df_abs_cours["justifie"] = False
            
                                    # Indexation des absences par étudiant
                                    absents_details = {}
                                    abs_count_par_etu = {}
                                    if not df_abs_cours.empty:
                                        for _, row in df_abs_cours.iterrows():
                                            nom = row["etud_non_eligible"]
                                            abs_count_par_etu[nom] = abs_count_par_etu.get(nom, 0) + 1
                                            if nom not in absents_details:
                                                absents_details[nom] = {
                                                    "motif": row.get("cause_non_eligibilite", "Non justifie"),
                                                    "date": row.get("date_absence", ""),
                                                    "jour": row.get("jour_absence", ""),
                                                    "horaire": row.get("horaire_absence", ""),
                                                    "justifie": row.get("justifie", False)
                                                }
            
                                    # Construction de la liste complète (Tous les étudiants + Statut)
                                    df_liste_finale = df_p.copy()
                                    df_liste_finale["Statut"] = df_liste_finale["Nom_Complet"].apply(
                                        lambda x: "🚫 EXCLU (5 absences)" if abs_count_par_etu.get(x, 0) >= 5 else ("❌ Absent" if x in abs_count_par_etu else "✅ Eligible")
                                    )
                                    df_liste_finale["Justifiee"] = df_liste_finale["Nom_Complet"].map(
                                        lambda x: "Oui" if absents_details.get(x, {}).get("justifie") == True else ("Non" if x in abs_count_par_etu else "")
                                    )
                                    df_liste_finale["Motif Absence"] = df_liste_finale["Nom_Complet"].map(
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
                                    df_liste_finale["Nb Absences (matiere)"] = df_liste_finale["Nom_Complet"].map(
                                        lambda x: abs_count_par_etu.get(x, 0)
                                    )
            
                                    # Réorganisation
                                    df_export = df_liste_finale[[
                                        "Nom_Complet", "Statut", "Justifiee", "Nb Absences (matiere)",
                                        "Motif Absence", "Date Absence", "Jour", "Horaire"
                                    ]].rename(columns={
                                        "Nom_Complet": "Nom et Prenom",
                                        "Nb Absences (matiere)": "Nb Absences"
                                    })
                                    df_export["Matiere"] = sel_mat
                                    df_export["Charge"] = sel_prof
                                    df_export["Promotion"] = promo_c
            
                                    # Génération Excel
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
                                        fmt_header = workbook.add_format({
                                            'bold': True, 'bg_color': '#0f172a', 'font_color': 'white',
                                            'border': 1, 'align': 'center', 'valign': 'vcenter'
                                        })
                                        fmt_exclu = workbook.add_format({
                                            'bg_color': '#fee2e2', 'font_color': '#991b1b', 'border': 1, 'bold': True
                                        })
                                        fmt_absent = workbook.add_format({
                                            'bg_color': '#fef3c7', 'font_color': '#92400e', 'border': 1
                                        })
                                        fmt_eligible = workbook.add_format({
                                            'bg_color': '#dcfce7', 'font_color': '#166534', 'border': 1
                                        })
                                        fmt_cell = workbook.add_format({'border': 1, 'valign': 'vcenter'})
            
                                        df_export.to_excel(writer, sheet_name='Liste_Eligibilite', startrow=8, index=False)
                                        ws = writer.sheets['Liste_Eligibilite']
            
                                        # En-tête institutionnel
                                        ws.merge_range('A1:J1', "UNIVERSITE DJILLALI LIABES - SIDI BEL ABBES", fmt_title)
                                        ws.merge_range('A2:J2', "Faculte de Genie Electrique - Departement d'Electrotechnique", fmt_sub)
                                        ws.merge_range('A3:J3', "LISTE D'ELIGIBILITE A L'EXAMEN", fmt_title)
                                        ws.write('A5', "Matiere :", fmt_bold); ws.write('B5', sel_mat)
                                        ws.write('A6', "Enseignant :", fmt_bold); ws.write('B6', sel_prof)
                                        ws.write('D5', "Promotion :", fmt_bold); ws.write('E5', promo_c)
                                        ws.write('D6', "Date export :", fmt_bold); ws.write('E6', datetime.now().strftime('%d/%m/%Y'))
            
                                        # Largeurs
                                        ws.set_column('A:A', 28)
                                        ws.set_column('B:B', 22)
                                        ws.set_column('C:C', 12)
                                        ws.set_column('D:D', 12)
                                        ws.set_column('E:H', 18)
                                        ws.set_column('I:J', 20)
            
                                        # Mise en forme conditionnelle par statut
                                        for row_num in range(9, 9 + len(df_export)):
                                            statut_val = df_export.iloc[row_num - 9]["Statut"]
                                            if "EXCLU" in str(statut_val):
                                                ws.set_row(row_num, None, fmt_exclu)
                                            elif "Absent" in str(statut_val):
                                                ws.set_row(row_num, None, fmt_absent)
                                            else:
                                                ws.set_row(row_num, None, fmt_eligible)
            
                                        ws.freeze_panes(9, 0)
            
                                    nb_eligibles = len(df_export[df_export["Statut"].str.contains("Eligible")])
                                    nb_absents = len(df_export[df_export["Statut"].str.contains("Absent")])
                                    nb_exclus = len(df_export[df_export["Statut"].str.contains("EXCLU")])
            
                                    st.success(
                                        f"✅ Rapport généré : **{nb_eligibles}** éligible(s) | **{nb_absents}** absent(s) | **{nb_exclus}** exclu(s) (≥5 absences)."
                                    )
                                    st.download_button(
                                        label="📥 TELECHARGER LE RAPPORT (XLSX)",
                                        data=output.getvalue(),
                                        file_name=f"Rapport_{sel_mat.replace(' ', '_')}_{promo_c}.xlsx",
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                        use_container_width=True
                                    )
                                except Exception as e:
                                    st.error(f"❌ Erreur Excel : {e}")
                                                
                            else:
                                st.session_state.absences = []
                                st.success("✅ Historique local effacé !")
                                time.sleep(0.5)
                                st.rerun()
                    else:
                        st.info("ℹ️ Aucune absence enregistrée dans l'historique global.")

                    # Export Rapport Officiel
                    # ─────────────────────────────────────────────────────────────
                    # RAPPORT OFFICIEL EXCEL (LISTE UNIQUE AVEC STATUT)
                    # ─────────────────────────────────────────────────────────────
                    st.divider()
                    st.subheader("📥 Rapport Officiel Excel — Liste d'Éligibilité")

                    # >>> NOUVEAUTÉ : Sélection de la promotion via liste déroulante <<<
                    toutes_promos = sorted(df_etu["Promotion"].dropna().unique().tolist())
                    promo_rapport = st.selectbox(
                        "🎓 Sélectionner la promotion pour le rapport :",
                        options=toutes_promos,
                        index=toutes_promos.index(promo_c) if promo_c in toutes_promos else 0,
                        key="promo_rapport_select"
                    )

                    # Rechargement des étudiants pour la promotion choisie
                    df_p_rapport = df_etu[df_etu["Promotion"].astype(str).str.strip().str.upper() == promo_rapport.upper()].copy()

                    if not df_p_rapport.empty:
                        try:
                            # 1. Récupération des absences pour cette MATIERE et cette PROMOTION
                            if MODE_SUPABASE:
                                absences_cours = charger_absences_supabase(sel_mat, promo_rapport)
                            else:
                                absences_cours = [
                                    a for a in st.session_state.absences
                                    if a.get("matiere") == sel_mat and a.get("promotion") == promo_rapport
                                ]
                            df_abs_cours = pd.DataFrame(absences_cours)

                            # 2. Liste des absents pour CE cours
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

                            # 3. Construction de la liste complète (Tous les étudiants + Statut)
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

                            # Réorganisation des colonnes
                            df_export = df_liste_finale[[
                                "Nom_Complet", "Statut", "Motif du Retrait",
                                "Date Absence", "Jour", "Horaire"
                            ]].rename(columns={
                                "Nom_Complet": "Nom et Prénom",
                                "Motif du Retrait": "Motif Absence",
                                "Date Absence": "Date",
                                "Horaire": "Horaire"
                            })

                            # Ajout des infos du cours en colonnes fixes
                            df_export["Matiere"] = sel_mat
                            df_export["Charge"] = sel_prof
                            df_export["Promotion"] = promo_rapport

                            # 4. Génération Excel avec mise en forme conditionnelle
                            output = io.BytesIO()
                            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                                workbook = writer.book

                                # Formats
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

                                # Écriture
                                df_export.to_excel(writer, sheet_name='Liste_Eligibilite', startrow=8, index=False)
                                ws = writer.sheets['Liste_Eligibilite']

                                # En-tête institutionnel
                                ws.merge_range('A1:I1', "UNIVERSITE DJILLALI LIABES - SIDI BEL ABBES", fmt_title)
                                ws.merge_range('A2:I2', "Faculte de Genie Electrique - Departement d'Electrotechnique", fmt_sub)
                                ws.merge_range('A3:I3', "LISTE D'ELIGIBILITE A L'EXAMEN", fmt_title)
                                ws.write('A5', "Matiere :", fmt_bold); ws.write('B5', sel_mat)
                                ws.write('A6', "Enseignant :", fmt_bold); ws.write('B6', sel_prof)
                                ws.write('D5', "Promotion :", fmt_bold); ws.write('E5', promo_rapport)
                                ws.write('D6', "Date export :", fmt_bold); ws.write('E6', datetime.now().strftime('%d/%m/%Y'))

                                # Largeurs
                                ws.set_column('A:A', 28)
                                ws.set_column('B:B', 22)
                                ws.set_column('C:C', 28)
                                ws.set_column('D:F', 14)
                                ws.set_column('G:I', 20)

                                # Mise en forme conditionnelle des lignes
                                for row_num in range(9, 9 + len(df_export)):
                                    statut_val = df_export.iloc[row_num - 9]["Statut"]
                                    if "Eligible" in str(statut_val) and "Non" not in str(statut_val):
                                        ws.set_row(row_num, None, fmt_eligible)
                                    else:
                                        ws.set_row(row_num, None, fmt_non_eligible)

                                # Figer l'en-tête
                                ws.freeze_panes(9, 0)

                            # Compteurs
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
            st.header("📩 Systeme de Gestion des Justificatifs")
            st.caption("Depot etudiant et validation administration")

        choix_vue = st.radio("Profil :", ["Etudiant (Depot)", "Administration (Decision)"], horizontal=True)
        st.divider()

        
        if choix_vue == "Etudiant (Depot)":
            st.subheader("📤 Soumettre une demande de rehabilitation")
    
            # --- Sélection identité ---
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
    
            # --- MES ABSENCES SIGNALÉES ---
            st.divider()
            st.markdown("### 📋 Mes absences signalées")
    
            absences_etu = get_absences_etudiant(etudiant_sel)
    
            
            if absences_etu:
            # Construction d'un tableau récapitulatif avec statut du justificatif
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
                                                                        
              
              
              # Construction d'un tableau récapitulatif avec statut du justificatif
                data_display = []
                for abs_item in absences_etu:
                    mat = abs_item.get("matiere", "")
                    req = trouver_requete_existante(etudiant_sel, mat)
                    
                    if req:
                        statut_j = "🟡 " + req.get("statut", "En attente")
                        date_dep = req.get("date_demande", "-")
                    else:
                        statut_j = "🔴 Non déposé"
                        date_dep = "-"
    
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
    
                # --- FORMULAIRE D'ENVOI CIBLÉ ---
                st.markdown("### 📎 Envoyer un justificatif pour une absence")
    
                # On ne propose que les absences sans justificatif déposé
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
    
                                # Vérifier si une requête auto existe déjà (créée par l'enseignant)
                                req_ex = trouver_requete_existante(etudiant_sel, abs_conc["matiere"])
    
                                if MODE_SUPABASE:
                                    if req_ex:
                                        # Mise à jour de la demande déjà présente (créée auto par l'enseignant)
                                        supabase.table("requetes_absences").update({
                                            "justificatif_pdf": pdf_encoded,
                                            "motif": motif_dep,
                                            "date_demande": datetime.now().strftime("%d/%m/%Y")
                                        }).eq("id", req_ex["id"]).execute()
                                        st.success(f"✅ Justificatif ajouté à la demande existante pour **{abs_conc['matiere']}** !")
                                    else:
                                        # Création nouvelle demande
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
                                    # Marquer les absences comme justifiees (sans suppression)
                                    for a in st.session_state.absences:
                                        if (a.get("etud_non_eligible") == req['nom_etudiant']
                                                and a.get("matiere") == req['matiere']):
                                            a["justifie"] = True
                                            a["cause_non_eligibilite"] = "Justifiee - " + str(a.get("cause_non_eligibilite", ""))
                                st.success(f"✔️ Justificatif de {req['nom_etudiant']} pour {req['matiere']} accepté. Absence conservée mais marquée comme justifiée.")
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

                # Maintenance
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
    
                # Export Registre
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
    
                # Bilan par matiere
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
    
                # Total par etudiant
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
    
            # --- Bilan des absences directes ---
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
    
                # Comptage total (justifiées ou non) + dont justifiées
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

                if os.path.exists(str(_BASE_DIR / "logo.PNG")):
                    logo_w = W_LOGO - 4
                    logo_h = H_ENTETE - 4
                    self.image(str(_BASE_DIR / "logo.PNG"), x=X0 + 2, y=Y0 + 2, w=logo_w, h=logo_h)

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

                    if os.path.exists(str(_BASE_DIR / "logo.PNG")):
                        self.image(str(_BASE_DIR / "logo.PNG"), x=X0 + 2, y=Y0 + 2, w=W_LOGO - 4, h=H_ENTETE - 4)

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

                    if os.path.exists(str(_BASE_DIR / "logo.PNG")):
                        self.image(str(_BASE_DIR / "logo.PNG"), x=X0 + 2, y=Y0 + 2, w=W_LOGO - 4, h=H_ENTETE - 4)

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

                if os.path.exists(str(_BASE_DIR / "logo.PNG")):
                    self.image(str(_BASE_DIR / "logo.PNG"), x=X0 + 2, y=Y0 + 2, w=W_LOGO - 4, h=H_ENTETE - 4)

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
            c1, c2, c3 = st.columns(3)

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
    # NOM_FICHIER_FIXE déjà défini plus haut
    # NOM_FICHIER_CONTACTS déjà défini plus haut

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
    # =============================================================================
    # CHARGEMENT DU RÉPERTOIRE (CONTACTS, PRÉNOMS, STATUTS & GRADES)
    # =============================================================================
    repertoire_qualites = {} 
    repertoire_grades = {} 
    repertoire_source = {}        # Email par nom de famille
    repertoire_noms_complets = {} # Affichage "NOM Prénom"
    repertoire_telephones = {}
    df_contacts = None

    # NOM_FICHIER_CONTACTS déjà défini plus haut

    def extraire_nom_famille(nom_complet):
        """Extrait le nom de famille (premier mot) pour la correspondance."""
        if not nom_complet or pd.isna(nom_complet):
            return ""
        return str(nom_complet).strip().upper().split()[0]

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
                    repertoire_telephones[nom_complet.upper()] = tel_nettoye

                if nom_brut:
                    nom_complet = f"{nom_brut} {prénom_brut}".strip()

                    # Stockage par NOM SEUL (clé principale)
                    if email_brut and email_brut.lower() != 'nan':
                        repertoire_source[nom_brut] = email_brut
                    repertoire_noms_complets[nom_brut] = nom_complet
                    repertoire_qualites[nom_brut] = qualite_brute
                    repertoire_grades[nom_brut] = grade_brut

                    # Stockage aussi par NOM COMPLET (fallback)
                    repertoire_noms_complets[nom_complet.upper()] = nom_complet
                    repertoire_qualites[nom_complet.upper()] = qualite_brute
                    repertoire_grades[nom_complet.upper()] = grade_brut
                    if email_brut and email_brut.lower() != 'nan':
                        repertoire_source[nom_complet.upper()] = email_brut

        except Exception as e:
            st.error(f"Erreur lors de la lecture du fichier contacts: {e}")
    else:
        st.warning(f"⚠️ Fichier {NOM_FICHIER_CONTACTS} introuvable. Les noms complets et emails ne seront pas disponibles.")
    # =============================================================================
    # ACTIVATION DE COMPTE PAR TOKEN (depuis lien email)
    # =============================================================================
    # =============================================================================
    # ACTIVATION DE COMPTE PAR TOKEN (depuis lien email)
    # =============================================================================
    query_params = st.query_params
    if "activation_token" in query_params and query_params["activation_token"]:
        st.markdown("<h1 class='main-title'>🏛️ ACTIVATION DU COMPTE</h1>", unsafe_allow_html=True)

        token = str(query_params["activation_token"]).strip()

        # Vérification du token
        res = supabase.table("enseignants_auth").select("*").eq("activation_token", token).execute()

        if not res.data:
            st.error("❌ Lien d'activation invalide ou déjà utilisé.")
            st.stop()

        user_row = res.data[0]
        expires_str = user_row.get('activation_expires')

        # Vérification expiration
        if expires_str:
            from datetime import timezone
            expires = datetime.fromisoformat(str(expires_str).replace("Z", "+00:00"))
            if datetime.now(timezone.utc) > expires:
                st.error("⏰ Ce lien d'activation a expiré. Veuillez refaire une demande d'inscription.")
                st.stop()

        st.success(f"👋 Bienvenue **{user_row['nom_officiel']}**, définissez votre mot de passe pour activer le compte.")

        # ═══════════════════════════════════════════════════════════════
        # FORMULAIRE D'ACTIVATION (sans bouton de navigation à l'intérieur)
        # ═══════════════════════════════════════════════════════════════
        with st.form("form_activation_token"):
            new_pass = st.text_input("Choisissez un mot de passe", type="password")
            confirm_pass = st.text_input("Confirmez le mot de passe", type="password")
            submitted = st.form_submit_button("🚀 Activer mon compte", use_container_width=True, type="primary")

            if submitted:
                if not new_pass or not confirm_pass:
                    st.error("Veuillez remplir les deux champs.")
                elif len(new_pass) < 6:
                    st.error("Le mot de passe doit contenir au moins 6 caractères.")
                elif new_pass != confirm_pass:
                    st.error("Les mots de passe ne correspondent pas.")
                else:
                    try:
                        supabase.table("enseignants_auth").update({
                            "password_hash": hash_pw(new_pass),
                            "activation_token": None,
                            "activation_expires": None
                        }).eq("id", user_row['id']).execute()

                        # On stocke le succès dans le session_state pour l'afficher hors du form
                        st.session_state['activation_success'] = True
                        st.rerun()

                    except Exception as e:
                        st.error(f"Erreur technique : {e}")

        # ═══════════════════════════════════════════════════════════════
        # BOUTON DE NAVIGATION (HORS DU FORMULAIRE)
        # ═══════════════════════════════════════════════════════════════
        if st.session_state.get('activation_success'):
            st.success("✅ Compte activé avec succès ! Vous pouvez maintenant vous connecter.")
            if st.button("🔑 Aller à la connexion", use_container_width=True, type="primary"):
                st.session_state['activation_success'] = False  # Reset
                st.query_params.clear()
                st.rerun()

        st.stop()
    # --- SYSTÈME D'AUTH ---
    if "user_data" not in st.session_state:
        st.session_state["user_data"] = None

    if not st.session_state["user_data"]:
        st.markdown("<h1 class='main-title'>🏛️ DÉPARTEMENT D'ÉLECTROTECHNIQUE-FGE- UDL-SBA</h1>", unsafe_allow_html=True)
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

            # ═══════════════════════════════════════════════════════════════
            # 🔒 RÉCUPÉRATION DE MOT DE PASSE
            # ═══════════════════════════════════════════════════════════════
            st.divider()
            with st.expander("🔒 Mot de passe oublié ?"):
                email_reset = st.text_input("Votre email enregistré", key="reset_email")

                c1, c2 = st.columns(2)
                with c1:
                    if st.button("📧 Générer un token de réinitialisation", use_container_width=True, key="btn_gen_token"):
                        import secrets
                        check = supabase.table("enseignants_auth").select("email").eq("email", email_reset).execute()
                        if check.data:
                            token = secrets.token_urlsafe(32)
                            # Expiration dans 1 heure (format ISO pour Supabase)
                            expiration = (datetime.now().replace(microsecond=0)).isoformat() + "+01:00"
                            supabase.table("enseignants_auth").update({
                                "reset_token": token,
                                "reset_expires": expiration
                            }).eq("email", email_reset).execute()
                            st.success("✅ Token généré (valable 1h) :")
                            st.code(token, language="text")
                            st.caption("💡 Dans un système en production, ce token serait envoyé par email automatiquement.")
                        else:
                            st.error("❌ Cet email n'est pas enregistré dans la base.")

                with c2:
                    token_input = st.text_input("Token reçu", type="password", key="token_input")
                    new_pass_reset = st.text_input("Nouveau mot de passe", type="password", key="new_pass_reset")
                    confirm_pass_reset = st.text_input("Confirmer le mot de passe", type="password", key="confirm_pass_reset")

                    if st.button("🔄 Valider la réinitialisation", use_container_width=True, key="btn_reset_pass"):
                        if not token_input or not new_pass_reset:
                            st.error("Veuillez remplir tous les champs.")
                        elif new_pass_reset != confirm_pass_reset:
                            st.error("Les mots de passe ne correspondent pas.")
                        elif len(new_pass_reset) < 6:
                            st.error("Le mot de passe doit contenir au moins 6 caractères.")
                        else:
                            res = supabase.table("enseignants_auth").select("*")\
                                .eq("email", email_reset)\
                                .eq("reset_token", token_input).execute()
                            if res.data:
                                # Vérification de l'expiration
                                try:
                                    from datetime import timezone
                                    expires_str = res.data[0]['reset_expires'].replace("Z", "+00:00")
                                    expires = datetime.fromisoformat(expires_str)
                                    if datetime.now(timezone.utc) < expires:
                                        supabase.table("enseignants_auth").update({
                                            "password_hash": hash_pw(new_pass_reset),
                                            "reset_token": None,
                                            "reset_expires": None
                                        }).eq("email", email_reset).execute()
                                        st.success("✅ Mot de passe mis à jour avec succès ! Vous pouvez maintenant vous connecter.")
                                    else:
                                        st.error("⏰ Token expiré. Veuillez en générer un nouveau.")
                                except Exception as e:
                                    st.error(f"Erreur de validation du token : {e}")
                            else:
                                st.error("Token invalide ou email incorrect.")



        with t_ins:
            st.subheader("📝 Demande d'activation de compte")
            st.info("NB : Saisissez votre email professionnel tel qu'envoyé au service d'enseignement. Vos informations personnelles se rempliront automatiquement. Une fois inscrit, veuillez accéder à votre compte via la page de connexion.")

            # ═══════════════════════════════════════════════════════════════
            # ÉTAPE 1 : AUTHENTIFICATION PAR EMAIL (Filtrage de l'identité)
            # ═══════════════════════════════════════════════════════════════
            email_verif = st.text_input(
                "📧 Saisissez votre email (celui envoyé au service d'enseignement du département)",
                key="verif_email_insc",
                placeholder="ex: nom.prénom@univ-sba.dz"
            )

            # Initialisation des variables de session pour l'inscription
            if "contact_match" not in st.session_state:
                st.session_state.contact_match = None

            # Bouton de vérification
            col_verif, _ = st.columns([1, 3])
            with col_verif:
                verifier = st.button("🔍 Vérifier mon identité", use_container_width=True, key="btn_verif_id")

            if verifier and email_verif:
                if df_contacts is not None and not df_contacts.empty and 'Email' in df_contacts.columns:
                    match = df_contacts[df_contacts["Email"].astype(str).str.strip().str.lower() == email_verif.strip().lower()]
                    if not match.empty:
                        st.session_state.contact_match = match.iloc[0]
                        st.success("✅ Identité confirmée. Vos coordonnées ont été récupérées.")
                    else:
                        st.session_state.contact_match = None
                        st.error("❌ Cet email n'est pas reconnu dans le répertoire officiel du département. Veuillez contacter l'administrateur.")
                else:
                    st.error("⚠️ Le fichier répertoire des contacts est introuvable ou corrompu.")

            # ═══════════════════════════════════════════════════════════════
            # ÉTAPE 2 : AFFICHAGE DU FORMULAIRE PRÉ-REMPLI (SI IDENTITÉ VÉRIFIÉE)
            # ═══════════════════════════════════════════════════════════════

            if st.session_state.contact_match is not None:
                row = st.session_state.contact_match

                # Extraction sécurisée des données
                nom_brut = str(row.get('NOM', '')).strip().upper()
                prénom_brut = str(row.get('PRÉNOM', '')).strip().capitalize()
                email_brut = str(row.get('Email', '')).strip()
                qualite_brute = str(row.get('Qualité', 'Non défini')).strip()
                tel_brut = str(row.get('N°/TEL', '')).strip()

                # Nettoyage strict : uniquement les chiffres
                tel_nettoye = ''.join([c for c in tel_brut if c.isdigit()])

                nom_complet = f"{nom_brut} {prénom_brut}"

                # ═══════════════════════════════════════════════════════════════
                # CORRECTION CRITIQUE : Forcer la réinitialisation du widget 
                # téléphone quand on change d'enseignant
                # ═══════════════════════════════════════════════════════════════
                contact_id = f"{nom_brut}_{email_brut}"
                if st.session_state.get("last_verified_contact") != contact_id:
                    if "tel_insc_modifiable" in st.session_state:
                        del st.session_state["tel_insc_modifiable"]
                    st.session_state["last_verified_contact"] = contact_id

                st.divider()
                st.markdown("### 👤 Votre fiche enseignant")

                # Affichage du nom (LECTURE SEULE)
                st.markdown(f"""
                <div style="background: linear-gradient(135deg, #1E3A8A 0%, #3B82F6 100%); 
                            padding: 15px; border-radius: 10px; color: white; margin-bottom: 15px;
                            box-shadow: 0 4px 6px rgba(0,0,0,0.1);">
                    <div style="font-size: 12px; opacity: 0.9; text-transform: uppercase; letter-spacing: 1px;">
                        Enseignant identifié
                    </div>
                    <div style="font-size: 22px; font-weight: bold; margin-top: 5px;">
                        {nom_complet}
                    </div>
                </div>
                """, unsafe_allow_html=True)

                c1, c2 = st.columns(2)

                with c1:
                    st.markdown(f"""
                    <div style="background-color:#f0f2f6;padding:12px;border-radius:8px;border-left:4px solid #22c55e;margin-bottom:10px;">
                        <span style="font-size:11px;color:#64748b;">📧 Adresse Email</span><br>
                        <span style="font-weight:bold;color:#1E3A8A;font-size:14px;">{email_brut}</span>
                    </div>
                    """, unsafe_allow_html=True)

                with c2:
                    st.markdown(f"""
                    <div style="background-color:#f0f2f6;padding:12px;border-radius:8px;border-left:4px solid #D4AF37;margin-bottom:10px;">
                        <span style="font-size:11px;color:#64748b;">🏷️ Qualité (auto-détectée)</span><br>
                        <span style="font-weight:bold;color:#1E3A8A;font-size:14px;">{qualite_brute}</span>
                    </div>
                    """, unsafe_allow_html=True)

                # Téléphone : pré-rempli avec le numéro nettoyé depuis le fichier source
                default_phone = tel_nettoye if len(tel_nettoye) == 10 else ""
                new_phone = st.text_input(
                    "📱 Numéro de téléphone (Obligatoire — 10 chiffres)",
                    value=default_phone,
                    key="tel_insc_modifiable",
                    help="Format strict : exactement 10 chiffres, sans espaces ni tirets (ex: 0555123456)",
                    max_chars=10
                )

                st.divider()

                # ═══════════════════════════════════════════════════════════════
                # ÉTAPE 3 : ENVOI DU LIEN D'ACTIVATION
                # ═══════════════════════════════════════════════════════════════
                BASE_URL = "https://emplois-du-temps-2026-2027-xadotqqqjnevp7zk2w2gbm.streamlit.app/"

                if st.button("📧 Envoyer le lien d'activation à votre adresse Email", use_container_width=True, type="primary"):
                    # Validation...
                    phone_clean = new_phone.strip()

                    if not new_phone:
                        st.error("📱 Le numéro de téléphone est obligatoire.")
                    elif not phone_clean.isdigit():
                        st.error("📱 Le numéro ne doit contenir que des chiffres.")
                    elif len(phone_clean) != 10:
                        st.error(f"📱 Le numéro doit contenir exactement 10 chiffres (actuellement {len(phone_clean)}).")
                    elif not email_brut or "@" not in email_brut:
                        st.error("❌ L'adresse email récupérée est invalide. Contactez l'administrateur.")
                    elif not qualite_brute or qualite_brute.lower() == "non défini":
                        st.error("❌ La qualité n'est pas reconnue dans le fichier source. Contactez l'administrateur.")
                    else:
                        # Vérifier si un compte actif existe déjà
                        check = supabase.table("enseignants_auth").select("email,activation_token").eq("email", email_brut).execute()
                        if check.data and not check.data[0].get('activation_token'):
                            st.error("❌ Cet email est déjà associé à un compte actif. Utilisez l'onglet **Connexion**.")
                        else:
                            import secrets
                            token = secrets.token_urlsafe(32)
                            expiration = (datetime.now().replace(microsecond=0) + timedelta(hours=24)).isoformat() + "+01:00"

                            data_upsert = {
                                "nom_officiel": nom_brut,
                                "email": email_brut,
                                "password_hash": hash_pw(secrets.token_urlsafe(16)),
                                "role": "enseignant",
                                "statut": qualite_brute,
                                "telephone": phone_clean,
                                "activation_token": token,
                                "activation_expires": expiration
                            }

                            try:
                                if check.data:
                                    supabase.table("enseignants_auth").update(data_upsert).eq("email", email_brut).execute()
                                else:
                                    supabase.table("enseignants_auth").insert(data_upsert).execute()

                                # ─── ENVOI EMAIL SMTP ───
                                import smtplib
                                from email.mime.text import MIMEText
                                from email.mime.multipart import MIMEMultipart

                                lien_activation = f"{BASE_URL}/?activation_token={token}"

                                msg = MIMEMultipart()
                                msg['Subject'] = "Activation de votre compte EDT - Département ELT"
                                msg['From'] = "chef.department.elt.fge@gmail.com"
                                msg['To'] = email_brut

                                body_html = f"""
                                <html>
                                <body style="font-family:Arial,sans-serif;color:#333;">
                                    <div style="max-width:600px;margin:auto;border:1px solid #e2e8f0;border-radius:10px;overflow:hidden;">
                                        <div style="background:linear-gradient(135deg,#1E3A8A 0%,#3B82F6 100%);padding:20px;color:white;text-align:center;">
                                            <h2 style="margin:0;">Département d'Électrotechnique - UDL SBA</h2>
                                        </div>
                                        <div style="padding:25px;background:#fff;">
                                            <p>Sallem Aleykoum <b>{nom_complet}</b>,</p>
                                            <p>Votre demande d'inscription a été enregistrée dans la <b>Plateforme de gestion des EDTs du département d'électrotechnique</b>.</p>
                                            <p><b>Qualité détectée :</b> {qualite_brute}<br>
                                            <b>Téléphone :</b> {phone_clean}</p>
                                            <p>Cliquez sur le bouton ci-dessous pour définir votre mot de passe :</p>
                                            <div style="text-align:center;margin:25px 0;">
                                                <a href="{lien_activation}" style="background:#1E3A8A;color:white;padding:12px 24px;text-decoration:none;border-radius:6px;font-weight:bold;display:inline-block;">
                                                    Activer mon compte
                                                </a>
                                            </div>
                                            <p>Ou copiez ce lien dans votre navigateur :<br>
                                            <code style="background:#f1f5f9;padding:8px;border-radius:4px;display:block;word-break:break-all;">{lien_activation}</code></p>
                                            <p style="color:#64748b;font-size:13px;"><i>Ce lien est valable 24 heures.</i></p>
                                            <hr style="border:none;border-top:1px solid #e2e8f0;">
                                            <p style="font-size:12px;color:#94a3b8;">Faculté de Génie Électrique - UDL SBA</p>
                                        </div>
                                    </div>
                                </body>
                                </html>
                                """
                                msg.attach(MIMEText(body_html, 'html'))

                                server = smtplib.SMTP('smtp.gmail.com', 587)
                                server.starttls()
                                server.login("chef.department.elt.fge@gmail.com", "gkzs pdza yodb icvd")
                                server.send_message(msg)
                                server.quit()

                                st.success("✅ Lien d'activation envoyé ! Consultez votre boîte mail (et vos spams).")
                                st.info(f"📧 Email envoyé à : `{email_brut}`")
                                st.balloons()

                            except Exception as e:
                                st.error(f"❌ Erreur lors de l'envoi : {e}")
            else:
                if not email_verif:
                    st.info("👆 Saisissez votre email professionnel et cliquez sur **Vérifier mon identité** pour commencer.")
                # Si email saisi mais pas de match, le message d'erreur est déjà affiché ci-dessus                    
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
    # NOM_FICHIER_FIXE déjà défini plus haut
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
        st.markdown("<h1 class='main-title'>🏛️ DÉPARTEMENT D'ÉLECTROTECHNIQUE-FGE- UDL-SBA</h1>", unsafe_allow_html=True)
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
    if is_admin:
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
        st.header(f"👤 {user.get('nom_officiel', 'Utilisateur')}")

        # ─── RESTRICTION DES PORTAILS SELON LE RÔLE ───
        if is_admin:
            options_portail = [
                "📖 Emploi du Temps", 
                "📅 Surveillances Examens", 
                "🤖 Générateur Automatique", 
                "👥 Portail Enseignants", 
                "🎓 Portail mise à jour EDT", 
                "📢 Gestion Administrative - Bordereaux & PVs"
            ]
        else:
            # ENSEIGNANT : accès strictement limité
            options_portail = [
                "👤 Mon Espace Enseignant",
                "📅 Surveillances Examens"
            ]

        portail = st.selectbox("🚀 Sélectionner Espace", options_portail)
        st.divider()

        mode_view = "Personnel"
        poste_sup = False

        if portail == "📖 Emploi du Temps" and is_admin:
            mode_view = st.radio("Vue Administration :", [
                "Promotion", "Enseignant", "🏢 Planning Salles", 
                "🚩 Vérificateur de conflits", "✍️ Éditeur de données"
            ])
            poste_sup = st.checkbox("Poste Supérieur (Décharge 3h)")
        elif portail == "👤 Mon Espace Enseignant":
            poste_sup = st.checkbox("Poste Supérieur (Décharge 3h)", key="poste_sup_ens")

        if st.button("🚪 Déconnexion du compte"):
            st.session_state["user_data"] = None
            st.rerun()
    # --- ESPACE ÉDITEUR AVANCÉ (ADMIN UNIQUEMENT) ---
    # --- ESPACE ÉDITEUR AVANCÉ (ADMIN UNIQUEMENT) ---
    if is_admin and mode_view == "✍️ Éditeur de données":
        st.divider()
        st.subheader("✍️ Plateforme de gestion des EDTs-Semestre 01__2026-2027-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA")

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
            st.image(str(_BASE_DIR / "logo.PNG"), width=90)
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
                        code_up = str(r['Code']).upper()
                        if 'COURS' in code_up:
                            nat, color, bg = '📘', '#1e40af', '#dbeafe'
                        elif 'TD' in code_up:
                            nat, color, bg = '📗', '#166534', '#dcfce7'
                        else:
                            nat, color, bg = '🔴', '#b91c1c', '#fee2e2'

                        txt = (f"<div style='margin-bottom:6px;padding:6px;border-left:3px solid {color};"
                               f"background-color:{bg};border-radius:4px;'>"
                               f"<b style='color:{color};'>{nat} {r['Enseignements']}</b><br>"
                               f"<span style='font-size:11px;'>({r['Code']})</span><br>"
                               f"<span style='font-size:11px;'>📍 {r['Lieu']}</span><br>"
                               f"<b style='font-size:11px;'>🎓 {r['Promotion']}</b></div>")
                        items.append(txt)
                    return "".join(items)

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
                                    t = "Plateforme de gestion des EDTs-Semeste 01__2026-2027-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA"
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
                            pdf.cell(0, 6, f"Récapitulatif : {nb_cours} Cours | {nb_td} TD | {nb_tp} TP", 0, 1, "C")
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
                                <div class="main-title">Plateforme de gestion des EDTs-Semestre 01__2026-2027-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA</div>
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
                            title = "Plateforme de gestion des EDTs-Semestre 01__2026-2027-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA"
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
                            self.cell(0, 8, "Plateforme de gestion des EDTs-Semestre 01__2026-2027-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA", 0, 1, 'C')
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
                            file_name="Pack_EDT_HTML_S1.zip",
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
                            file_name="Pack_EDT_S1_2027_Lisible.pdf",
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
                            df.to_excel(writer, index=False, sheet_name='EDT_S1_2027')

                        excel_data = buffer.getvalue()

                        st.download_button(
                            label="⬇️ Télécharger le Pack Excel (Format .xlsx)",
                            data=excel_data,
                            file_name="Pack_EDT_S1_2027.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )

                    except Exception as e:
                        st.error(f"Une erreur est survenue lors de la génération : {e}")

            elif is_admin and mode_view == "🏢 Planning Salles":
                s_sel = st.selectbox("Choisir Salle / Amphi :", sorted(df["Lieu"].unique()))
                df_s = df[df["Lieu"].astype(str).str.startswith(s_sel)]

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
                    pdf_data_s, err_s = generate_edt_individuel_lieu_pdf(df_s, s_sel)
                    if pdf_data_s:
                        cs2.download_button(
                            label=f"📄 Planning {s_sel} (PDF)",
                            data=pdf_data_s,
                            file_name=f"Planning_{s_sel}_2027.pdf",
                            mime="application/pdf",
                            use_container_width=True,
                            key=f"pdf_salle_{s_sel}"
                        )
                    else:
                        cs2.error(f"Erreur PDF : {err_s}")
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
                            file_name=f"Solutions_Conflits_EDT_S1_2027.xlsx",
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
        elif portail == "👤 Mon Espace Enseignant":
                # ─────────────────────────────────────────────────────────────
                # 1. IDENTITÉ & INFOS PERSONNELLES
                # ─────────────────────────────────────────────────────────────
                cible = user['nom_officiel']
                nom_affichage_complet = repertoire_noms_complets.get(cible.strip().upper(), cible)
                grade_enseignant = repertoire_grades.get(cible.strip().upper(), "Grade non spécifié")
                statut_enseignant = repertoire_qualites.get(cible.strip().upper(), "Statut non spécifié")
                email_ens = user.get('email', 'Non renseigné')
                tel_ens = repertoire_telephones.get(cible.strip().upper(), user.get('telephone', 'Non renseigné'))

                # Carte d'identité stylisée
                st.markdown(f"""
                    <div style="background: linear-gradient(135deg, #1E3A8A 0%, #3B82F6 100%); 
                                padding: 20px; border-radius: 14px; color: white; margin-bottom: 20px;
                                box-shadow: 0 6px 12px rgba(0,0,0,0.12);">
                        <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 15px;">
                            <div>
                                <div style="font-size: 12px; opacity: 0.9; text-transform: uppercase; letter-spacing: 1px;">
                                    Espace Personnel Sécurisé
                                </div>
                                <div style="font-size: 24px; font-weight: bold; margin-top: 6px;">
                                    {nom_affichage_complet}
                                </div>
                                <div style="margin-top: 10px; display: flex; gap: 8px;">
                                    <span style="background: rgba(255,255,255,0.2); padding: 4px 12px; border-radius: 20px; font-size: 12px; font-weight: 600;">
                                        {grade_enseignant}
                                    </span>
                                    <span style="background: rgba(255,255,255,0.2); padding: 4px 12px; border-radius: 20px; font-size: 12px; font-weight: 600;">
                                        {statut_enseignant}
                                    </span>
                                </div>
                            </div>
                            <div style="text-align: right; font-size: 13px; line-height: 1.8;">
                                <div>📧 {email_ens}</div>
                                <div>📱 {tel_ens}</div>
                                <div style="opacity: 0.8; font-size: 11px; margin-top: 4px;">S1 — 2026-2027</div>
                            </div>
                        </div>
                    </div>
                """, unsafe_allow_html=True)

                # ─────────────────────────────────────────────────────────────
                # 2. CHARGEMENT DES DONNÉES PERSONNELLES
                # ─────────────────────────────────────────────────────────────
                if df is None or df.empty:
                    st.error("❌ Les données EDT ne sont pas disponibles. Contactez l'administrateur.")
                    st.stop()

                df_f = df[df["Enseignants"].str.contains(cible, case=False, na=False)].copy()

                if df_f.empty:
                    st.warning("⚠️ Aucun cours n'est programmé pour vous actuellement.")
                    st.stop()

                # Détermination des types
                df_f['Type'] = df_f['Code'].apply(
                    lambda x: "COURS" if "COURS" in str(x).upper() else ("TD" if "TD" in str(x).upper() else "TP")
                )
                df_u = df_f.drop_duplicates(subset=['j_norm', 'h_norm'])

                # Calculs de charge
                nb_cours = len(df_u[df_u['Type'] == 'COURS'])
                nb_td = len(df_u[df_u['Type'] == 'TD'])
                nb_tp = len(df_u[df_u['Type'] == 'TP'])

                seuil_obligatoire = 3.0 if poste_sup else 6.0
                charge_totale_eq = (nb_cours * 1.5) + (nb_td + nb_tp)
                delta_eq = charge_totale_eq - seuil_obligatoire
                h_sup = delta_eq * 1.5

                abs_h_sup = abs(h_sup)
                heures_entieres = int(abs_h_sup)
                minutes_restantes = int((abs_h_sup - heures_entieres) * 60)
                signe_str = "+" if h_sup >= 0 else "-"
                h_sup_formattee = f"{signe_str}{heures_entieres}h{minutes_restantes:02d}"
                charge_effective = (nb_cours + nb_td + nb_tp) * 1.5

                # ─────────────────────────────────────────────────────────────
                # 3. MÉTRIQUES DE CHARGE (Style badges)
                # ─────────────────────────────────────────────────────────────
                st.markdown(f"""
                    <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(140px, 1fr)); gap: 12px; margin-bottom: 20px;">
                        <div style="background: white; border-radius: 12px; padding: 16px; text-align: center; 
                                    box-shadow: 0 2px 8px rgba(0,0,0,0.04); border-top: 4px solid #3b82f6;">
                            <div style="font-size: 26px; font-weight: 800; color: #1e40af;">{nb_cours}</div>
                            <div style="font-size: 12px; color: #64748b; margin-top: 4px;">📘 Cours</div>
                        </div>
                        <div style="background: white; border-radius: 12px; padding: 16px; text-align: center; 
                                    box-shadow: 0 2px 8px rgba(0,0,0,0.04); border-top: 4px solid #22c55e;">
                            <div style="font-size: 26px; font-weight: 800; color: #166534;">{nb_td}</div>
                            <div style="font-size: 12px; color: #64748b; margin-top: 4px;">📗 TD</div>
                        </div>
                        <div style="background: white; border-radius: 12px; padding: 16px; text-align: center; 
                                    box-shadow: 0 2px 8px rgba(0,0,0,0.04); border-top: 4px solid #f59e0b;">
                            <div style="font-size: 26px; font-weight: 800; color: #b45309;">{nb_tp}</div>
                            <div style="font-size: 12px; color: #64748b; margin-top: 4px;">🔴 TP</div>
                        </div>
                        <div style="background: white; border-radius: 12px; padding: 16px; text-align: center; 
                                    box-shadow: 0 2px 8px rgba(0,0,0,0.04); border-top: 4px solid #1E3A8A;">
                            <div style="font-size: 26px; font-weight: 800; color: #1E3A8A;">{round(charge_effective, 1)}h</div>
                            <div style="font-size: 12px; color: #64748b; margin-top: 4px;">Charge Effective</div>
                        </div>
                        <div style="background: white; border-radius: 12px; padding: 16px; text-align: center; 
                                    box-shadow: 0 2px 8px rgba(0,0,0,0.04); border-top: 4px solid {'#22c55e' if h_sup >= 0 else '#ef4444'};">
                            <div style="font-size: 26px; font-weight: 800; color: {'#166534' if h_sup >= 0 else '#dc2626'};">{h_sup_formattee}</div>
                            <div style="font-size: 12px; color: #64748b; margin-top: 4px;">{'Heures Sup.' if h_sup >= 0 else 'Déficit'}</div>
                        </div>
                    </div>
                """, unsafe_allow_html=True)

                # Note de synthèse
                if h_sup > 0:
                    st.caption(f"✅ Vous avez complété votre charge et totalisez **{h_sup_formattee}** en supplément.")
                elif h_sup < 0:
                    st.caption(f"⚠️ Sous-charge détectée de **{h_sup_formattee}** par rapport au seuil de {seuil_obligatoire} eq/h.")
                else:
                    st.caption("⚖️ Service réglementaire exactement rempli.")

                # ─────────────────────────────────────────────────────────────
                # 4. EMPLOI DU TEMPS INDIVIDUEL (Grille Jours × Horaires)
                # ─────────────────────────────────────────────────────────────
                st.divider()
                st.markdown("### 📅 Mon Emploi du Temps Individuel")
                def format_case(rows):
                    items = []
                    for _, r in rows.iterrows():
                        code_up = str(r['Code']).upper()
                        if 'COURS' in code_up:
                            nat, color, bg = '📘', '#1e40af', '#dbeafe'
                        elif 'TD' in code_up:
                            nat, color, bg = '📗', '#166534', '#dcfce7'
                        else:
                            nat, color, bg = '🔴', '#b91c1c', '#fee2e2'

                        txt = (f"<div style='margin-bottom:6px;padding:6px;border-left:3px solid {color};"
                               f"background-color:{bg};border-radius:4px;'>"
                               f"<b style='color:{color};'>{nat} {r['Enseignements']}</b><br>"
                               f"<span style='font-size:11px;'>({r['Code']})</span><br>"
                               f"<span style='font-size:11px;'>📍 {r['Lieu']}</span><br>"
                               f"<b style='font-size:11px;'>🎓 {r['Promotion']}</b></div>")
                        items.append(txt)
                    return "".join(items)   
                grid = df_f.groupby(['h_norm', 'j_norm']).apply(format_case, include_groups=False).unstack('j_norm')
                grid = grid.reindex(
                    index=[normalize(h) for h in horaires_list], 
                    columns=[normalize(j) for j in jours_list]
                ).fillna("")
                grid.index = [map_h.get(i, i) for i in grid.index]
                grid.columns = [map_j.get(c, c) for c in grid.columns]

                # Suppression des lignes totalement vides
                grid = grid[grid.any(axis=1)]

                st.write(grid.to_html(escape=False), unsafe_allow_html=True)

                # ─────────────────────────────────────────────────────────────
                # 5. MES COLLÈGUES PAR PROMOTION (Accès restreint)
                # ─────────────────────────────────────────────────────────────
                st.divider()
                st.markdown("### 👥 Mes collègues par promotion")
                st.caption("🔒 Vous ne voyez que les enseignants des promotions où vous intervenez.")

                mes_promotions = [p for p in df_f['Promotion'].unique() 
                                  if p and str(p).strip() not in ["", "nan", "None", "Non defini", "Non défini"]]

                if len(mes_promotions) > 0:
                    for promo in sorted(mes_promotions):
                        df_promo = df[df['Promotion'] == promo].copy()

                        # Extraction du nom de famille de l'utilisateur connecté
                        cible_nom_famille = extraire_nom_famille(cible)

                        # Enseignants uniques de cette promo (exclusion exacte par nom de famille)
                        autres_ens = [e for e in df_promo['Enseignants'].unique() 
                                      if e and str(e).strip() not in ["", "nan", "None", "Non defini", "Non défini"]
                                      and extraire_nom_famille(e) != cible_nom_famille]

                        if len(autres_ens) == 0:
                            continue

                        with st.expander(f"🎓 {promo} — {len(autres_ens)} enseignant(s)", expanded=True):
                            for nom_col in sorted(autres_ens):
                                # Recherche d'abord par nom de famille, puis par nom complet
                                nom_key_famille = extraire_nom_famille(nom_col)
                                nom_key_complet = str(nom_col).strip().upper()

                                nom_complet_col = repertoire_noms_complets.get(
                                    nom_key_famille, 
                                    repertoire_noms_complets.get(nom_key_complet, nom_col)
                                )
                                grade_col = repertoire_grades.get(
                                    nom_key_famille,
                                    repertoire_grades.get(nom_key_complet, "N/A")
                                )
                                qualite_col = repertoire_qualites.get(
                                    nom_key_famille,
                                    repertoire_qualites.get(nom_key_complet, "N/A")
                                )
                                email_col = repertoire_source.get(
                                    nom_key_famille,
                                    repertoire_source.get(nom_key_complet, None)
                                )

                                # Badge couleur selon qualité
                                badge_color = "#22c55e" if "PERMANENT" in str(qualite_col).upper() else "#f59e0b"

                                col1, col2, col3 = st.columns([3, 2, 2])
                                with col1:
                                    st.markdown(f"**{nom_complet_col}**")
                                    st.caption(f"🏷️ {grade_col}")
                                with col2:
                                    st.markdown(f"<span style='color:{badge_color}; font-weight:600; font-size:12px;'>{qualite_col}</span>", 
                                               unsafe_allow_html=True)
                                with col3:
                                    if email_col and "@" in str(email_col):
                                        st.markdown(f"📧 `{email_col}`")
                                    else:
                                        st.markdown("<span style='color:#94a3b8; font-size:12px;'>📧 Non communiqué</span>", 
                                                   unsafe_allow_html=True)
                                st.markdown("<hr style='margin: 8px 0; border: none; border-top: 1px solid #f1f5f9;'>", 
                                           unsafe_allow_html=True)
                else:
                    st.info("Vous n'êtes assigné à aucune promotion actuellement.")

                # ─────────────────────────────────────────────────────────────
                # 6. TÉLÉCHARGEMENTS (UNIQUEMENT SES PROPRES DONNÉES)
                # ─────────────────────────────────────────────────────────────
                st.divider()
                st.markdown("### 📥 Exporter mes données")

                col_dl1, col_dl2, col_dl3 = st.columns(3)

                # Excel
                with col_dl1:
                    buf_ex = io.BytesIO()
                    df_export = df_f.drop(columns=['h_norm', 'j_norm', 'Type'], errors='ignore')
                    with pd.ExcelWriter(buf_ex, engine='xlsxwriter') as writer:
                        df_export.to_excel(writer, index=False, sheet_name='Mon_EDT')
                        wb = writer.book
                        ws = writer.sheets['Mon_EDT']
                        header_fmt = wb.add_format({'bold': True, 'bg_color': '#1E3A8A', 'font_color': 'white', 'border': 1})
                        for col_num, value in enumerate(df_export.columns.values):
                            ws.write(0, col_num, value, header_fmt)
                        ws.set_column('A:G', 20)
                    st.download_button("📊 Excel", buf_ex.getvalue(), f"EDT_{cible}_2027.xlsx", 
                                      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                      use_container_width=True)

                # HTML
                with col_dl2:
                    html_content = generate_pro_html(df_export, f"EDT - {nom_affichage_complet}", 
                                                     "Département d'Électrotechnique - FGE/UDL-SBA")
                    st.download_button("🌐 HTML", html_content, f"EDT_{cible}_2027.html", "text/html",
                                      use_container_width=True)

                # PDF (grille PPER.03)
                with col_dl3:
                    pdf_data, err_pdf = generate_edt_individuel_pdf_classique(df_f, nom_affichage_complet)
                    if pdf_data:
                        st.download_button("📄 PDF", pdf_data, f"EDT_{cible}_2027.pdf", "application/pdf",
                                          use_container_width=True)
                    else:
                        st.button("📄 PDF", disabled=True, use_container_width=True, help=err_pdf)

                st.stop()  # Empêche l'enseignant d'accéder au reste du code
        elif portail == "📅 Surveillances Examens":
            FILE_S = str(_BASE_DIR / "surveillances_2027.xlsx")
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
                st.error("Le fichier 'surveillances_Semestre 01-2026-2027.xlsx' non établi pour le moment.")

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

                SRC = str(_BASE_DIR / "surveillances_2027.xlsx")
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
                                    res_list.append({"Enseignements": row[C_MAT], "Code": "S1-2027", "Enseignants": " & ".join(equipe) if len(equipe) >= 2 else "⚠️ BESOIN RENFORT", "Horaire": row[C_HEURE], "Jours": row[C_DATE], "Lieu": f"Salle {s_idx}" if nb_salles > 1 else row[C_SALLE], "Promotion": f"{p_name} (S{s_idx})" if nb_salles > 1 else p_name})
                        st.session_state.df_généré = pd.DataFrame(res_list)
                        st.session_state.stats_charge = stats
                        st.rerun()

                    if st.session_state.get("df_généré") is not None:
                        st.dataframe(st.session_state.df_généré, use_container_width=True, hide_index=True)
                        xlsx_buf = io.BytesIO()
                        with pd.ExcelWriter(xlsx_buf, engine='xlsxwriter') as writer: st.session_state.df_généré.to_excel(writer, index=False)
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
                st.image(str(_BASE_DIR / "logo.PNG"), width=80)
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
                                msg['Subject'] = f"Votre Emploi du Temps S1-2027 - {row['Enseignant']}"

                                # --- CORRECTION DES EN-TÊTES ---
                                msg['From'] = f"{nom_affichage} <{expediteur_email}>"
                                msg['To'] = row["Email"]

                                # --- CORPS DU MESSAGE (TEXTE NON CONDENSÉ) ---
                                corps_html = f"""
                                <html>
                                <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
                                    <h2 style="color: #1E3A8A; border-bottom: 2px solid #1E3A8A; padding-bottom: 10px;">
                                        Plateforme de gestion des EDTs-Semestre 01__2026-2027-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA
                                    </h2>

                                    <p>Sallem M./Mme <b>{row['Enseignant']}</b>,</p>

                                    <div style="background-color: #f8f9fa; padding: 10px; border: 1px solid #dee2e6; border-radius: 5px; margin-bottom: 15px;">
                                        <b>📊 Récapitulatif de votre charge (S1-2027) :</b><br>
                                        <ul>
                                            <li>Nombre de Cours : <b>{nb_cours}</b></li>
                                            <li>Nombre de TD : <b>{nb_td}</b></li>
                                            <li>Nombre d'unité de TP : <b>{nb_tp}</b></li>
                                        </ul>
                                    </div>

                                    <div style="background-color: #fff4e5; border-left: 5px solid #ffa500; padding: 15px; margin: 20px 0;">
                                        <p style="font-weight: bold; color: #d97706; margin-top: 0;">
                                            Objet : Urgent : Vérification de l’emploi du temps – Semestre 1
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
                                part.add_header('Content-Disposition', f'attachment; filename="EDT_S1_2027_{row["Enseignant"]}.xlsx"')
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
    if is_admin:
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
                    msg['Subject'] = f"Votre Emploi du Temps S1-2027 - {nom_ens}"
                    msg['From'] = f"Département d'Électrotechnique <{EMAIL_EXPEDITEUR}>"
                    msg['To'] = email_ens

                    table_html = df_mail.to_html(index=False, border=1, justify='center')

                    corps_html = f"""
                    <html>
                    <body style="font-family: Arial, sans-serif;">
                        <div style="background-color: #f4f4f4; padding: 15px; border-radius: 5px; border: 1px solid #1E3A8A;">
                            <h2 style="color: #1E3A8A; text-align: center;">Plateforme de gestion des EDTs-Semestre 01__2026-2027-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA</h2>
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
                    part.add_header('Content-Disposition', f'attachment; filename="EDT_S1_2027_{nom_ens}.xlsx"')
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
                    file_name=f"Emails_{choix_promo}_S1_2027.xlsx",
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
                                    msg['Subject'] = f"Votre Emploi du Temps S1-2027 - {nom}"
                                    msg['From'] = st.secrets["EMAIL_USER"]
                                    msg['To'] = info_ens["Email"]

                                    # --- CORPS DU MESSAGE (IDENTIQUE À L'INDIVIDUEL) ---
                                    corps_html = f"""
                                    <html>
                                    <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
                                        <h2 style="color: #1E3A8A; border-bottom: 2px solid #1E3A8A; padding-bottom: 10px;">
                                            Plateforme de gestion des EDTs-Semestre 01__2026-2027-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA
                                        </h2>

                                        <p>Sallem M./Mme <b>{row['Enseignant']}</b>,</p>

                                        <div style="background-color: #f8f9fa; padding: 10px; border: 1px solid #dee2e6; border-radius: 5px; margin-bottom: 15px;">
                                            <b>📊 Récapitulatif de votre charge (S1-2027) :</b><br>
                                            <ul>
                                                <li>Nombre de Cours : <b>{nb_cours}</b></li>
                                                <li>Nombre de TD : <b>{nb_td}</b></li>
                                                <li>Nombre d'unité de TP : <b>{nb_tp}</b></li>
                                            </ul>
                                        </div>

                                        <div style="background-color: #fff4e5; border-left: 5px solid #ffa500; padding: 15px; margin: 20px 0;">
                                            <p style="font-weight: bold; color: #d97706; margin-top: 0;">
                                                Objet : Urgent : Vérification de l’emploi du temps – Semestre 1
                                            </p>

                                            <p>Cher collègue, Sallem,</p>

                                            <p>Vous trouverez ci-joint votre emploi du temps individuel pour le premier semestre.<br>
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
                                    part.add_header('Content-Disposition', f'attachment; filename="EDT_S1_2027_{nom}.xlsx"')
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
                                    msg['Subject'] = f"Votre Emploi du Temps S1-2027 - {row['Enseignant']}"

                                    # --- MODIFICATION DES EN-TÊTES ---
                                    msg['From'] = f"{nom_aff} <{exp_mail}>"
                                    msg['To'] = row["Email"]

                                    # --- CORPS DU MESSAGE MIS À JOUR ---
                                    corps = f"""
                                    <html>
                                    <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
                                        <h2 style="color: #1E3A8A; border-bottom: 2px solid #1E3A8A; padding-bottom: 10px;">
                                            Plateforme de gestion des EDTs-Semestre 01__2026-2027-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA
                                        </h2>

                                        <p>Sallem M./Mme <b>{row['Enseignant']}</b>,</p>

                                        <div style="background-color: #fff4e5; border-left: 5px solid #ffa500; padding: 15px; margin: 20px 0;">
                                            <p style="font-weight: bold; color: #d97706; margin-top: 0;">
                                                Objet : Urgent : Vérification de l’emploi du temps – Semestre 1
                                            </p>

                                            <p>Cher collègue, Sallem,</p>

                                            <p>Vous trouverez ci-joint votre emploi du temps individuel pour le prmier semestre.<br>
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

                                        <p>Voici le récapitulatif de votre emploi du temps pour le semestre 01 (S1-2027) :</p>

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
        elif portail == "🎓 Portail mise à jour EDT":
            st.write(f"**MODE ACTIF :** {portail}")
            st.subheader("📚 Espace mise à jour EDT")

            # Rappel du titre obligatoire
            st.info("Plateforme de gestion des EDTs-Semestre 01__2026-2027-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA")

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

            # ═══════════════════════════════════════════════════════════════
            # SECTION IMPORT EXCEL
            # ═══════════════════════════════════════════════════════════════
            # ═══════════════════════════════════════════════════════════════
            # SECTION IMPORT EXCEL - CORRIGÉE (avec "Remplacer tout" par enseignant)
            # ═══════════════════════════════════════════════════════════════
            with st.expander("📥 Importer des données depuis un fichier Excel", expanded=False):
                st.markdown("**Format attendu :** `Enseignements | Code | Enseignants | Horaire | Jours | Lieu | Promotion`")

                # Liste déroulante des enseignants pour filtrage pré-import
                liste_enseignants = sorted(df['Enseignants'].dropna().unique().tolist())
                enseignant_filtre = st.selectbox(
                    "👤 Filtrer par enseignant avant l'import (optionnel) :",
                    options=["Tous les enseignants"] + liste_enseignants,
                    key="import_enseignant_filter"
                )

                uploaded_file = st.file_uploader(
                    "Choisir un fichier Excel (.xlsx) :",
                    type=["xlsx"],
                    key="excel_uploader"
                )

                if uploaded_file is not None:
                    try:
                        df_import = pd.read_excel(uploaded_file)

                        # Vérification des colonnes requises
                        colonnes_requises = ['Enseignements', 'Code', 'Enseignants', 'Horaire', 'Jours', 'Lieu', 'Promotion']
                        colonnes_manquantes = [c for c in colonnes_requises if c not in df_import.columns]

                        if colonnes_manquantes:
                            st.error(f"❌ Colonnes manquantes dans le fichier : {', '.join(colonnes_manquantes)}")
                        else:
                            # Filtrage par enseignant si sélectionné
                            if enseignant_filtre != "Tous les enseignants":
                                df_import = df_import[df_import['Enseignants'] == enseignant_filtre]

                            st.success(f"✅ {len(df_import)} lignes prêtes à être importées")
                            st.dataframe(df_import, use_container_width=True)

                            # Choix du mode d'intégration
                            mode_import = st.radio(
                                "Mode d'intégration :",
                                options=["➕ Ajouter (fusionner avec l'existant)", "🔄 Remplacer (supprimer l'ancien pour cette promotion)"],
                                key="mode_import"
                            )

                            # Sélection de la promotion cible pour le remplacement
                            promo_cible = None

                            if "Remplacer" in mode_import:
                                promos_import = sorted([str(p) for p in df_import['Promotion'].unique()])

                                # ═══════════════════════════════════════════════════════
                                # NOUVEAUTÉ : Option "Remplacer tout" par enseignant
                                # ═══════════════════════════════════════════════════════
                                options_promo = promos_import.copy()
                                if enseignant_filtre != "Tous les enseignants":
                                    options_promo.insert(0, f"🗑️ TOUTES LES LIGNES DE {enseignant_filtre}")

                                promo_cible = st.selectbox(
                                    "Promotion à remplacer :",
                                    options=options_promo,
                                    key="promo_remplacement"
                                )

                                # Détection du mode "Tout remplacer"
                                if enseignant_filtre != "Tous les enseignants" and isinstance(promo_cible, str) and promo_cible.startswith("🗑️ TOUTES"):
                                    st.session_state['promo_cible_import'] = "TOUT_ENSEIGNANT"
                                    st.warning(f"🗑️ **Mode Remplacer TOUT actif** : toutes les lignes de **{enseignant_filtre}** seront supprimées avant l'ajout.")
                                else:
                                    st.session_state['promo_cible_import'] = str(promo_cible).strip()
                                    st.warning(f"🗑️ **Mode Remplacer actif** : les anciennes lignes de la promotion **{promo_cible}** seront supprimées avant l'ajout.")
                            else:
                                st.session_state['promo_cible_import'] = None
                                st.info("➕ **Mode Ajouter actif** : les nouvelles lignes seront fusionnées avec les existantes.")

                            # Bouton d'intégration
                            if st.button("💾 Intégrer les données importées", key="btn_integrer"):
                                try:
                                    # Récupération fiable depuis le session state
                                    current_mode = st.session_state.get('mode_import', mode_import)
                                    current_promo = st.session_state.get('promo_cible_import')

                                    # Chargement du fichier maître (création si inexistant)
                                    if os.path.exists(NOM_FICHIER_FIXE):
                                        if NOM_FICHIER_FIXE.endswith('.xlsx'):
                                            df_master = pd.read_excel(NOM_FICHIER_FIXE)
                                        else:
                                            df_master = pd.read_csv(NOM_FICHIER_FIXE)
                                    else:
                                        df_master = pd.DataFrame(columns=colonnes_requises)

                                    # ═══════════════════════════════════════════════════════
                                    # CORRECTION CRITIQUE : uniformiser les types en string
                                    # ═══════════════════════════════════════════════════════
                                    if 'Promotion' in df_master.columns:
                                        df_master['Promotion'] = df_master['Promotion'].astype(str).str.strip()
                                    if 'Enseignants' in df_master.columns:
                                        df_master['Enseignants'] = df_master['Enseignants'].astype(str).str.strip()
                                    df_import['Promotion'] = df_import['Promotion'].astype(str).str.strip()
                                    df_import['Enseignants'] = df_import['Enseignants'].astype(str).str.strip()

                                    lignes_avant = len(df_master)
                                    lignes_supprimees = 0

                                    if "Remplacer" in current_mode and current_promo is not None:
                                        if current_promo == "TOUT_ENSEIGNANT":
                                            # MODE REMPLACER TOUT : suppression par enseignant
                                            masque_suppr = df_master['Enseignants'] == str(enseignant_filtre).strip()
                                            lignes_supprimees = int(masque_suppr.sum())

                                            if lignes_supprimees > 0:
                                                df_master = df_master[~masque_suppr].copy()
                                                st.info(f"🗑️ {lignes_supprimees} ligne(s) de l'enseignant '{enseignant_filtre}' supprimée(s).")
                                            else:
                                                st.warning(f"⚠️ Aucune ligne trouvée pour l'enseignant '{enseignant_filtre}' dans le fichier actuel.")
                                        else:
                                            # MODE REMPLACER PROMOTION : suppression par promotion
                                            masque_suppr = df_master['Promotion'] == str(current_promo).strip()
                                            lignes_supprimees = int(masque_suppr.sum())

                                            if lignes_supprimees > 0:
                                                df_master = df_master[~masque_suppr].copy()
                                                st.info(f"🗑️ {lignes_supprimees} ligne(s) de la promotion '{current_promo}' supprimée(s).")
                                            else:
                                                st.warning(f"⚠️ Aucune ligne trouvée pour la promotion '{current_promo}' dans le fichier actuel.")

                                    # Ajout des nouvelles lignes (dans les deux modes)
                                    df_final = pd.concat([df_master, df_import], ignore_index=True)

                                    # Nettoyage et tri
                                    df_final = df_final.dropna(subset=['Enseignements'])
                                    df_final = df_final.sort_values(by=["Promotion", "Jours", "Horaire"])

                                    # Sauvegarde
                                    if NOM_FICHIER_FIXE.endswith('.xlsx'):
                                        df_final.to_excel(NOM_FICHIER_FIXE, index=False)
                                    else:
                                        df_final.to_csv(NOM_FICHIER_FIXE, index=False)

                                    st.success(
                                        f"✅ Importation terminée !\n\n"
                                        f"• Lignes avant import : **{lignes_avant}**\n"
                                        f"• Lignes supprimées : **{lignes_supprimees}**\n"
                                        f"• Lignes importées : **{len(df_import)}**\n"
                                        f"• **Total après import : {len(df_final)}**"
                                    )
                                    st.rerun()

                                except Exception as e:
                                    st.error(f"❌ Erreur lors de l'intégration : {e}")
                                    import traceback
                                    st.code(traceback.format_exc())

                    except Exception as e:
                        st.error(f"❌ Erreur de lecture du fichier : {e}")

            # ═══════════════════════════════════════════════════════════════
            # BARRE DE RECHERCHE
            # ═══════════════════════════════════════════════════════════════
            # ═══════════════════════════════════════════════════════════════
            # BARRE DE RECHERCHE & ÉDITEUR
            # ═══════════════════════════════════════════════════════════════

            # Définition locale (sécurité contre NameError)
            colonnes_ordonnees = ['Enseignements', 'Code', 'Enseignants', 'Horaire', 'Jours', 'Lieu', 'Promotion']

            # Vérification que df est bien chargé
            if df is None or df.empty:
                st.error("❌ Les données (df) ne sont pas chargées. Impossible d'afficher l'éditeur.")
                st.stop()

            # S'assurer que toutes les colonnes existent dans df
            for col in colonnes_ordonnees:
                if col not in df.columns:
                    df[col] = ""

            recherche = st.text_input("🔍 Rechercher une ligne (Enseignant, Salle ou Code) :", key="admin_search_bar")

            # Préparation du DataFrame maître
            df_master = df[colonnes_ordonnees].copy()

            # Application du filtre si recherche active
            if recherche:
                masque = df_master.apply(lambda r: r.astype(str).str.contains(recherche, case=False).any(), axis=1)
                df_edition = df_master[masque].copy()
            else:
                df_edition = df_master.copy()

            # Affichage du compteur de lignes pour le suivi de l'index
            total_lignes = len(df)
            st.caption(f"Lignes totales dans le fichier source : {total_lignes} | Prochain index : {total_lignes}")

            # L'ÉDITEUR DYNAMIQUE
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
                    # Suppression des lignes où Enseignements est vide ou "nan"
                    df_final = df_final[df_final['Enseignements'].astype(str).str.strip() != '']
                    df_final = df_final[df_final['Enseignements'].astype(str).str.lower() != 'nan']

                    # Tri pour l'organisation
                    df_final = df_final.sort_values(by=["Promotion", "Jours", "Horaire"])

                    # Sauvegarde dans le fichier maître
                    if NOM_FICHIER_FIXE.endswith('.xlsx'):
                        df_final.to_excel(NOM_FICHIER_FIXE, index=False)
                    else:
                        df_final.to_csv(NOM_FICHIER_FIXE, index=False)

                    st.success(f"✅ Modifications sauvegardées avec succès ! {len(df_final)} lignes enregistrées.")
                    st.balloons()
                    st.rerun()

                except Exception as e:
                    st.error(f"❌ Erreur lors de la sauvegarde : {e}")

            # --- 4. GESTION DES DOUBLONS & CONFLITS ---
            st.write("---")
            st.subheader("🔍 Vérificateur de Conflits")

            col1, col2, col3 = st.columns(3)

            with col1:
                if st.button("🔎 Détecter les doublons de séances"):
                    doublons = df_master[df_master.duplicated(subset=['Enseignements', 'Code', 'Promotion'], keep=False)]
                    if not doublons.empty:
                        st.warning(f"⚠️ {len(doublons)} doublons trouvés :")
                        st.dataframe(doublons.sort_values(by=["Promotion", "Enseignements"]), use_container_width=True)
                    else:
                        st.success("✅ Aucun doublon détecté.")

            with col2:
                if st.button("🏫 Détecter les conflits de salles"):
                    conflits_salle = df_master.groupby(['Jours', 'Horaire', 'Lieu']).size().reset_index(name='count')
                    conflits_salle = conflits_salle[conflits_salle['count'] > 1]
                    if not conflits_salle.empty:
                        st.warning(f"⚠️ {len(conflits_salle)} conflits de salle trouvés :")
                        st.dataframe(conflits_salle, use_container_width=True)
                    else:
                        st.success("✅ Aucun conflit de salle détecté.")

            with col3:
                if st.button("👨‍🏫 Détecter les surcharges enseignants"):
                    # Comptage des heures par enseignant
                    df_h = df_master.copy()
                    df_h['Duree'] = df_h['Horaire'].str.extract(r'(\d+)').astype(float)
                    surcharge = df_h.groupby('Enseignants')['Duree'].sum().reset_index()
                    surcharge = surcharge[surcharge['Duree'] > 20]  # Seuil arbitraire
                    if not surcharge.empty:
                        st.warning(f"⚠️ {len(surcharge)} enseignants potentiellement surchargés :")
                        st.dataframe(surcharge.sort_values('Duree', ascending=False), use_container_width=True)
                    else:
                        st.success("✅ Aucune surcharge détectée.")
        # --- 5. ESPACE PUBLIC (VISUALISATION LECTURE SEULE) ---
        # --- 5. ESPACE PUBLIC (VISUALISATION LECTURE SEULE) ---
        else:
            st.subheader("📅 Emploi du Temps - Vue Publique")

            # ═══════════════════════════════════════════════════════════════
            # DÉFINITION LOCALE (sécurité contre NameError)
            # ═══════════════════════════════════════════════════════════════
            colonnes_ordonnees = ['Enseignements', 'Code', 'Enseignants', 'Horaire', 'Jours', 'Lieu', 'Promotion']

            # Vérification que df existe et n'est pas vide
            if df is None or df.empty:
                st.warning("⚠️ Aucune donnée EDT disponible pour l'affichage public.")
                st.stop()

            # Filtres publics
            col_f1, col_f2 = st.columns(2)
            with col_f1:
                promo_pub = st.selectbox(
                    "🎓 Filtrer par promotion :",
                    options=["Toutes"] + sorted(df['Promotion'].dropna().unique().tolist()),
                    key="pub_promo"
                )
            with col_f2:
                jour_pub = st.selectbox(
                    "📆 Filtrer par jour :",
                    options=["Tous"] + sorted(df['Jours'].dropna().unique().tolist()),
                    key="pub_jour"
                )

            df_pub = df.copy()
            if promo_pub != "Toutes":
                df_pub = df_pub[df_pub['Promotion'] == promo_pub]
            if jour_pub != "Tous":
                df_pub = df_pub[df_pub['Jours'] == jour_pub]

            # Affichage stylisé
            st.dataframe(
                df_pub[colonnes_ordonnees],
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Enseignements": st.column_config.TextColumn("Matière", width="large"),
                    "Code": st.column_config.TextColumn("Code", width="small"),
                    "Enseignants": st.column_config.TextColumn("Intervenant", width="medium"),
                    "Horaire": st.column_config.TextColumn("Horaire", width="small"),
                    "Jours": st.column_config.TextColumn("Jour", width="small"),
                    "Lieu": st.column_config.TextColumn("Salle", width="small"),
                    "Promotion": st.column_config.TextColumn("Promo", width="small"),
                }
            )

            # Export public
            st.write("---")
            c1, c2 = st.columns(2)
            with c1:
                out_pub = io.BytesIO()
                df_pub.to_excel(out_pub, index=False)
                st.download_button("📊 Télécharger la vue (Excel)", out_pub.getvalue(), "EDT_Vue_Publique.xlsx")
            with c2:
                st.download_button("📄 Télécharger la vue (HTML)", df_pub.to_html(index=False), "EDT_Vue_Publique.html", "text/html")

        if not is_admin:
            st.error("🚫 ACCÈS RESTREINT.")
            st.stop()
        # ==========================================
        # IMPORTS SPÉCIFIQUES (idéalement à déplacer en haut du fichier)
        # ==========================================
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
        TITRE_PLATEFORME = "Plateforme de gestion des EDTs-Semestre 01__2026-2027-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA"

        DEPARTEMENTS = [
            "Département d'Électrotechnique",
            "Département d'Électronique",
            "Département d'Automatique",
            "Département de Télécommunications"
        ]

        TYPES_DOCUMENTS = [
            "Bordereau d'envoi"
        ]

        OPTIONS_DESTINATAIRES = [
            "Le Doyen de la faculté",
            "Le vice Doyen de la Post graduation",
            "Le vice Doyen de la graduation",
            "Le chef de département",
            "Autres"
        ]
        OPTIONS_EXPEDITEURS = [
            "Chef de département",
            "Chef de département adjoint",
            "Chef service de scolarité",
            "Chef service d'enseignements",
            "Signataire"
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
        def construire_reference(numero, annee=None):
            """Construit la référence complète du bordereau."""
            if annee is None:
                annee = datetime.now().year
            num_str = str(numero).split('/')[0] if '/' in str(numero) else str(numero)
            return f"{num_str}/F.G.E/Département-ELT/{annee}"

        def générer_bordereau_iso(département, donnees):
            doc = Document()

            for section in doc.sections:
                section.top_margin = Inches(0.8)
                section.bottom_margin = Inches(0.8)
                section.left_margin = Inches(0.8)
                section.right_margin = Inches(0.8)
                section.different_first_page_header_footer = False

                footer = section.footer
                footer_p = footer.paragraphs[0]
                footer_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                footer_pPr = footer_p._p.get_or_add_pPr()
                tabs = OxmlElement('w:tabs')
                tab_centre = OxmlElement('w:tab')
                tab_centre.set(qn('w:val'), 'center')
                tab_centre.set(qn('w:pos'), '4968')
                tabs.append(tab_centre)
                tab_droite = OxmlElement('w:tab')
                tab_droite.set(qn('w:val'), 'right')
                tab_droite.set(qn('w:pos'), '9936')
                tabs.append(tab_droite)
                footer_pPr.append(tabs)

                footer_p.add_run("\t")
                annee_doc = donnees.get('annee_reference', datetime.now().year)
                r_ref_fixe = footer_p.add_run(f"Réf : UDL-GEL-ER-004-{annee_doc}")
                r_ref_fixe.font.name = 'Calibri'
                r_ref_fixe.font.size = Pt(11)

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

            # 1. EN-TÊTE : Tableau invisible Logo | Texte officiel
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

            p_logo = cell_logo.paragraphs[0]
            p_logo.alignment = WD_ALIGN_PARAGRAPH.LEFT
            nom_fichier_logo = str(_BASE_DIR / "logo.PNG")
            if os.path.exists(nom_fichier_logo):
                p_logo.add_run().add_picture(nom_fichier_logo, width=Inches(0.833))
            else:
                r_alt = p_logo.add_run("[LOGO UNIVERSITÉ]")
                r_alt.font.name = 'Calibri'
                r_alt.font.size = Pt(8)
                r_alt.font.italic = True

            p_en_tete = cell_texte.paragraphs[0]
            p_en_tete.alignment = WD_ALIGN_PARAGRAPH.CENTER

            r1 = p_en_tete.add_run("République Algérienne Démocratique et populaire\n")
            r1.bold = True
            r1.font.size = Pt(12)
            r1.font.name = 'Calibri'

            r2 = p_en_tete.add_run(
                "Ministère de l'enseignement supérieur et de la recherche scientifiques\n"
                "Université Djillali Liabes de Sidi Bel Abbés\n"
                "Faculté de Génie Electrique\n"
            )
            r2.bold = True
            r2.font.size = Pt(12)
            r2.font.name = 'Calibri'

            r_dept = p_en_tete.add_run(f"{département.upper()}\n")
            r_dept.bold = True
            r_dept.font.size = Pt(11)
            r_dept.font.name = 'Calibri'

            doc.add_paragraph("\n")

            # 2. RÉFÉRENCE
            p_ref = doc.add_paragraph()
            p_ref.alignment = WD_ALIGN_PARAGRAPH.LEFT
            r_ref = p_ref.add_run(f"N° : {donnees['num_reference']}")
            r_ref.font.size = Pt(10)
            r_ref.font.name = 'Calibri'
            r_ref.bold = True

            doc.add_paragraph("\n")

            # 3. TITRE
            p_titre = doc.add_paragraph()
            p_titre.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_titre = p_titre.add_run("BORDEREAU D'ENVOI")
            r_titre.font.name = 'Calibri'
            r_titre.font.size = Pt(36)
            r_titre.italic = True
            r_titre.underline = True
            r_titre.bold = True

            doc.add_paragraph("\n")

            # 4. DESTINATAIRE
            p_dest = doc.add_paragraph()
            p_dest.alignment = WD_ALIGN_PARAGRAPH.LEFT
            r_dest = p_dest.add_run(f"A monsieur : {donnees['destinataire']}")
            r_dest.bold = True
            r_dest.font.size = Pt(12)
            r_dest.font.name = 'Calibri'

            doc.add_paragraph("\n")

            # 5. TABLEAU DE TRANSMISSION
            liste_pieces = donnees['liste_pieces']
            nb_lignes_totales = 2 + len(liste_pieces)

            table = doc.add_table(rows=nb_lignes_totales, cols=3)
            table.style = 'Table Grid'
            table.columns[0].width = Inches(4.5)
            table.columns[1].width = Inches(0.8)
            table.columns[2].width = Inches(1.7)

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

            row_joint = table.rows[1].cells
            row_joint[0].text = "Veuillez trouver ci-joint :"
            row_joint[0].paragraphs[0].runs[0].font.italic = True
            row_joint[0].paragraphs[0].runs[0].font.name = 'Calibri'
            row_joint[0].paragraphs[0].runs[0].font.size = Pt(10)
            set_cell_margins(row_joint[0], top=80, bottom=80)

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
            qualite_expediteur = donnees.get('expediteur_qualite', 'Chef de département')

            run_sig = p_signatures.add_run(f"Sidi bel Abbès le : {date_texte}\t\t\t\t{qualite_expediteur}")
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

        def générer_pv_generique(département, type_pv, donnees):
            """Générateur secondaire de secours (Calibri)."""
            doc = Document()
            p = doc.add_paragraph()
            run = p.add_run(f"{type_pv} - {département}\nDocument en cours.")
            run.font.name = 'Calibri'
            return doc

        # ==========================================
        # HISTORIQUE DES BORDEREAUX (SUPABASE)
        # ==========================================
        def enregistrer_historique_bordereau(donnees, departement, user_email):
            """Enregistre un bordereau généré dans l'historique Supabase."""
            try:
                ref_pur = donnees.get('num_reference_pur', 1)
                annee_ref = donnees.get('annee_reference', datetime.now().year)

                data_histo = {
                    "generated_by": user_email,
                    "departement": departement,
                    "destinataire": donnees.get('destinataire', ''),
                    "num_reference": str(ref_pur),
                    "annee_reference": annee_ref,
                    "expediteur_qualite": donnees.get('expediteur_qualite', 'Chef de département'),
                    "date_creation": donnees.get('date_creation', datetime.now()).isoformat(),
                    "nombre_pieces": len(donnees.get('liste_pieces', [])),
                    "pieces_details": donnees.get('liste_pieces', []),
                    "fichier_nom": f"Bordereau_{departement.replace(' ', '_')}.docx"
                }
                supabase.table("bordereaux_historique").insert(data_histo).execute()
            except Exception as e:
                st.warning(f"⚠️ Sauvegarde historique échouée : {e}")

        def get_prochaine_reference():
            """Récupère le prochain numéro de référence depuis l'historique."""
            try:
                res = supabase.table("bordereaux_historique")\
                              .select("num_reference")\
                              .order("num_reference", desc=True)\
                              .limit(1)\
                              .execute()
                if res.data and len(res.data) > 0:
                    dernier = res.data[0].get('num_reference', '0')
                    try:
                        if isinstance(dernier, str) and '/' in dernier:
                            dernier = dernier.split('/')[0]
                        return int(dernier) + 1
                    except ValueError:
                        return 1
                return 1
            except Exception:
                return 1

        def afficher_historique_bordereaux():
            """Affiche l'historique des bordereaux avec export Excel et effacement sécurisé."""
            try:
                res = supabase.table("bordereaux_historique")\
                              .select("*")\
                              .order("created_at", desc=True)\
                              .limit(50)\
                              .execute()
                if res.data:
                    from collections import Counter

                    compteur_dest = Counter([row.get('destinataire', 'Non spécifié') for row in res.data])
                    total_bordereaux = len(res.data)

                    st.markdown("### 📊 Tableau de bord — Bordereaux par destination")

                    c_total, c_unique = st.columns(2)
                    c_total.metric("📨 Total bordereaux générés", total_bordereaux)
                    c_unique.metric("🏛️ Destinations distinctes", len(compteur_dest))

                    st.divider()

                    st.markdown("**Répartition par destinataire :**")
                    destinations = sorted(compteur_dest.items(), key=lambda x: x[1], reverse=True)

                    cols = st.columns(min(3, len(destinations)))
                    for idx, (dest, count) in enumerate(destinations):
                        with cols[idx % 3]:
                            st.markdown(f"""
                                <div style="
                                    background: linear-gradient(135deg, #1E3A8A 0%, #3B82F6 100%);
                                    border-radius: 12px;
                                    padding: 16px;
                                    color: white;
                                    text-align: center;
                                    margin-bottom: 12px;
                                    box-shadow: 0 4px 6px rgba(0,0,0,0.1);
                                ">
                                    <div style="font-size: 11px; opacity: 0.9; text-transform: uppercase; letter-spacing: 1px;">
                                        {dest}
                                    </div>
                                    <div style="font-size: 32px; font-weight: bold; margin: 8px 0;">
                                        {count}
                                    </div>
                                    <div style="font-size: 12px; opacity: 0.8;">
                                        bordereau{'x' if count > 1 else ''}
                                    </div>
                                </div>
                            """, unsafe_allow_html=True)

                    st.divider()

                    rows_recap = []
                    rows_detail = []

                    for row in res.data:
                        date_str = pd.to_datetime(row['created_at']).strftime('%d/%m/%Y %H:%M')

                        rows_recap.append({
                            'Date': date_str,
                            'Généré par': row.get('generated_by', '—'),
                            'Expéditeur': row.get('expediteur_qualite', '—'),
                            'Département': row.get('departement', '—'),
                            'Destinataire': row.get('destinataire', '—'),
                            'N° Référence': row.get('num_reference', '—'),
                            'Nb pièces': row.get('nombre_pieces', 0),
                            'Fichier': row.get('fichier_nom', '—')
                        })

                        pieces = row.get('pieces_details', [])
                        ref_num = row.get('num_reference', '—')
                        ref_annee = row.get('annee_reference', datetime.now().year)
                        ref_full = construire_reference(ref_num, ref_annee)

                        if isinstance(pieces, list) and len(pieces) > 0:
                            for p in pieces:
                                rows_detail.append({
                                    'Date génération': date_str,
                                    'Généré par': row.get('generated_by', '—'),
                                    'Expéditeur': row.get('expediteur_qualite', '—'),
                                    'Département': row.get('departement', '—'),
                                    'Destinataire': row.get('destinataire', '—'),
                                    'N° Référence': ref_full,
                                    'Désignation des pièces': p.get('Désignation des pièces', ''),
                                    'Nbre': p.get('Nbre', ''),
                                    'Observations': p.get('Observations', ''),
                                    'Fichier': row.get('fichier_nom', '—')
                                })
                        else:
                            rows_detail.append({
                                'Date génération': date_str,
                                'Généré par': row.get('generated_by', '—'),
                                'Expéditeur': row.get('expediteur_qualite', '—'),
                                'Département': row.get('departement', '—'),
                                'Destinataire': row.get('destinataire', '—'),
                                'N° Référence': ref_full,
                                'Désignation des pièces': '—',
                                'Nbre': '—',
                                'Observations': '—',
                                'Fichier': row.get('fichier_nom', '—')
                            })

                    df_recap = pd.DataFrame(rows_recap)
                    df_detail = pd.DataFrame(rows_detail)

                    col_titre, col_dl, col_del = st.columns([3, 1, 1])

                    with col_titre:
                        st.markdown("**📊 Vue d'ensemble des bordereaux**")

                    with col_dl:
                        buffer_excel = io.BytesIO()
                        with pd.ExcelWriter(buffer_excel, engine='xlsxwriter') as writer:
                            df_recap.to_excel(writer, index=False, sheet_name='Récapitulatif')
                            ws1 = writer.sheets['Récapitulatif']
                            header_fmt = writer.book.add_format({
                                'bold': True, 'bg_color': '#1E3A8A', 'font_color': 'white', 'border': 1
                            })
                            for col_num, value in enumerate(df_recap.columns.values):
                                ws1.write(0, col_num, value, header_fmt)
                                ws1.set_column(col_num, col_num, 18)

                            df_detail.to_excel(writer, index=False, sheet_name='Détail des pièces')
                            ws2 = writer.sheets['Détail des pièces']
                            for col_num, value in enumerate(df_detail.columns.values):
                                ws2.write(0, col_num, value, header_fmt)
                                ws2.set_column(col_num, col_num, 22)

                        st.download_button(
                            label="📥 Excel",
                            data=buffer_excel.getvalue(),
                            file_name=f"Historique_Bordereaux_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,
                            key="dl_histo_bordereaux"
                        )

                    with col_del:
                        if st.button("🗑️ Effacer", use_container_width=True, key="btn_del_histo"):
                            st.session_state['confirmer_suppression_historique'] = True

                    if st.session_state.get('confirmer_suppression_historique'):
                        st.warning("⚠️ **Action irréversible** — Tous les bordereaux enregistrés seront supprimés définitivement.")

                        c1, c2 = st.columns([1, 1])
                        with c1:
                            if st.button("✅ Oui, supprimer définitivement", type="primary", key="confirm_del_yes"):
                                try:
                                    supabase.table("bordereaux_historique").delete().neq('id', -1).execute()
                                    st.success("✅ Historique effacé avec succès.")
                                    del st.session_state['confirmer_suppression_historique']
                                    st.rerun()
                                except Exception as e:
                                    st.error(f"❌ Erreur lors de la suppression : {e}")
                        with c2:
                            if st.button("❌ Non, annuler", key="confirm_del_no"):
                                del st.session_state['confirmer_suppression_historique']
                                st.rerun()

                    st.dataframe(df_recap, use_container_width=True, hide_index=True)

                    st.divider()
                    st.markdown("**📋 Détail par bordereau**")

                    for i, row in enumerate(res.data):
                        date_str = pd.to_datetime(row['created_at']).strftime('%d/%m/%Y %H:%M')
                        ref_num = row.get('num_reference', '—')
                        ref_annee = row.get('annee_reference', datetime.now().year)
                        ref_full = construire_reference(ref_num, ref_annee)
                        dest = row.get('destinataire', '—')

                        with st.expander(
                            f"📝 Bordereau N° {ref_full} — {dest} — {date_str}", 
                            expanded=(i == 0)
                        ):
                            c1, c2, c3, c4, c5 = st.columns(5)
                            c1.markdown(f"**👤 Généré par**\n{row.get('generated_by', '—')}")
                            c2.markdown(f"**🏛️ Département**\n{row.get('departement', '—')}")
                            c3.markdown(f"**📤 Expéditeur**\n{row.get('expediteur_qualite', '—')}")
                            c4.markdown(f"**📅 Date**\n{date_str}")
                            c5.markdown(f"**📎 Fichier**\n{row.get('fichier_nom', '—')}")

                            st.markdown("---")
                            st.markdown("**Tableau de transmission :**")

                            pieces = row.get('pieces_details', [])
                            if isinstance(pieces, list) and len(pieces) > 0:
                                df_pieces = pd.DataFrame(pieces)
                                cols_ordre = []
                                for col in ['Désignation des pièces', 'Nbre', 'Observations']:
                                    if col in df_pieces.columns:
                                        cols_ordre.append(col)
                                if cols_ordre:
                                    df_pieces = df_pieces[cols_ordre]
                                    st.dataframe(df_pieces, use_container_width=True, hide_index=True)
                                else:
                                    st.json(pieces)
                            else:
                                st.info("Aucune pièce enregistrée pour ce bordereau.")

                else:
                    st.info("📭 Aucun bordereau enregistré dans l'historique.")

            except Exception as e:
                st.error(f"Erreur chargement historique : {e}")

        # ==========================================
        # INTERFACE UTILISATEUR STREAMLIT
        # ==========================================
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
            prochaine_ref = get_prochaine_reference()

            col_ref, col_date, col_exp = st.columns(3)

            with col_ref:
                annee_courante = datetime.now().year
                prochaine_ref_num = get_prochaine_reference()
                ref_auto = construire_reference(prochaine_ref_num, annee_courante)

                donnees_doc['num_reference'] = st.text_input(
                    "Référence séquentielle", 
                    value=ref_auto,
                    help="Auto-incrémentée selon l'historique d'envoi"
                )    
                try:
                    donnees_doc['num_reference_pur'] = int(str(donnees_doc['num_reference']).split('/')[0])
                except ValueError:
                    donnees_doc['num_reference_pur'] = prochaine_ref_num
                donnees_doc['annee_reference'] = annee_courante
            with col_date:
                donnees_doc['date_creation'] = st.date_input("Date d'édition", datetime.now())
            with col_exp:
                donnees_doc['expediteur_qualite'] = st.selectbox(
                    "Qualité de l'expéditeur :", 
                    OPTIONS_EXPEDITEURS,
                    index=0
                )

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

        if doc_choisi == "Bordereau d'envoi":
            if st.button("Compiler et Générer le Bordereau Officiel"):
                if not donnees_doc['destinataire'].strip():
                    st.error("Erreur : Le champ de destination personnalisée ne peut pas être vide.")
                else:
                    try:
                        document_final = générer_bordereau_iso(dept_choisi, donnees_doc)

                        output_stream = io.BytesIO()
                        document_final.save(output_stream)
                        output_stream.seek(0)

                        user_email = user.get('email', 'inconnu') if user else 'inconnu'
                        enregistrer_historique_bordereau(donnees_doc, dept_choisi, user_email)

                        st.success(f"✓ Bordereau {donnees_doc['num_reference']} généré et enregistré.")

                        nom_fichier_export = f"Bordereau_{dept_choisi.replace(' ', '_')}.docx"
                        st.download_button(
                            label="⬇️ Télécharger le document (.docx)",
                            data=output_stream,
                            file_name=nom_fichier_export,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                    except Exception as error:
                        st.error(f"Échec de l'opération de génération : {str(error)}")

        st.divider()
        with st.expander("📜 Historique détaillé des bordereaux générés", expanded=False):
            afficher_historique_bordereaux()

# =============================================================================
# POINT D'ENTREE PRINCIPAL
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

# Lancement du module sélectionné
if module_sel == "📊 Suivi d'Assiduité":
    run_assiduite()
else:
    run_edt()
