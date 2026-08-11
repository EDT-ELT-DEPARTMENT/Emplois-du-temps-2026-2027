# =============================================================================
# IMPORTS OPTIMISÉS
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
from collections import defaultdict, Counter
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from pathlib import Path

# Imports optionnels avec gestion d'erreur
try:
    from supabase import create_client, Client
    SUPABASE_AVAILABLE = True
except ImportError:
    SUPABASE_AVAILABLE = False

try:
    from fpdf import FPDF
    FPDF_AVAILABLE = True
except ImportError:
    FPDF_AVAILABLE = False

try:
    from docx import Document
    from docx.shared import Inches, Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

# =============================================================================
# CONFIGURATION STREAMLIT
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
# GESTION DES CHEMINS DE FICHIERS
# =============================================================================
def get_base_dir():
    """Retourne le répertoire de base de manière robuste"""
    try:
        # En production Streamlit
        base_dir = Path(__file__).parent.resolve()
    except:
        # Fallback - utiliser le répertoire courant
        base_dir = Path.cwd()
    return base_dir

BASE_DIR = get_base_dir()

# Dictionnaire des fichiers avec vérification
FICHIERS_CONFIG = {
    'etudiants': 'Liste des étudiants_2026-2027.xlsx',
    'edt': 'dataEDT-ELT-S1-2027.xlsx',
    'enseignants': 'Permanents-Vacataires-ELT2-2026-2027.xlsx'
}

def verifier_fichiers():
    """Vérifie l'existence des fichiers nécessaires"""
    fichiers_manquants = []
    for key, nom_fichier in FICHIERS_CONFIG.items():
        chemin = BASE_DIR / nom_fichier
        if not chemin.exists():
            fichiers_manquants.append(nom_fichier)
    return fichiers_manquants

# Vérifier les fichiers au démarrage
fichiers_manquants = verifier_fichiers()
if fichiers_manquants:
    st.warning(f"⚠️ Fichiers manquants : {', '.join(fichiers_manquants)}")
    st.info("Assurez-vous que les fichiers sont dans le même répertoire que l'application.")

FILE_ETUDIANTS = str(BASE_DIR / FICHIERS_CONFIG['etudiants'])
FILE_EDT = str(BASE_DIR / FICHIERS_CONFIG['edt'])
FILE_ENS = str(BASE_DIR / FICHIERS_CONFIG['enseignants'])

# =============================================================================
# CONNEXION SUPABASE OPTIMISÉE
# =============================================================================
@st.cache_resource
def init_supabase():
    """Initialise la connexion Supabase une seule fois"""
    if not SUPABASE_AVAILABLE:
        return None, False
    
    try:
        SUPABASE_URL = st.secrets.get("SUPABASE_URL", "")
        SUPABASE_KEY = st.secrets.get("SUPABASE_KEY", "")
        
        if SUPABASE_URL and SUPABASE_KEY:
            supabase = create_client(SUPABASE_URL, SUPABASE_KEY)
            return supabase, True
    except Exception as e:
        st.warning(f"Erreur Supabase : {str(e)}")
    
    return None, False

supabase, MODE_SUPABASE = init_supabase()

# =============================================================================
# CONSTANTES
# =============================================================================
HORAIRES_LIST = [
    "8h - 9h30", "9h30 - 11h", "11h - 12h30", "12h30 - 14h", "14h - 15h30", "15h30 - 17h"
]
JOURS_SEMAINE = ["Dimanche", "Lundi", "Mardi", "Mercredi", "Jeudi"]

CAUSES_ABSENCES = [
    "Non justifiée",
    "Décès dans l'ascendance, la descendance ou la parenté",
    "Mariage de l'intéressé(e)",
    "Congé de paternité ou de maternité de l'intéressé(e)",
    "Mission ou convocation officielle",
    "Maladie de l'intéressé(e)",
    "Autres"
]

CODE_ADMIN = "1234"
CODE_ADMIN_EDT = "doctorat2026"

# =============================================================================
# FONCTIONS UTILITAIRES PRINCIPALES
# =============================================================================

@st.cache_data
def charger_fichier_excel(chemin_fichier, sheet_name=0):
    """Charge un fichier Excel avec gestion d'erreur"""
    try:
        if not os.path.exists(chemin_fichier):
            st.error(f"❌ Fichier non trouvé : {chemin_fichier}")
            return None
        
        df = pd.read_excel(chemin_fichier, sheet_name=sheet_name)
        return df
    except Exception as e:
        st.error(f"❌ Erreur lors du chargement de {chemin_fichier}: {str(e)}")
        return None

def nettoyer_nom_enseignant(nom):
    """Nettoie les noms d'enseignants"""
    n = str(nom).strip()
    prefixes = ["Pr ", "Dr ", "Mme ", "Mr ", "Dr. ", "Pr. ", "M. "]
    for prefix in prefixes:
        if n.startswith(prefix):
            n = n[len(prefix):]
    return n.strip()

def extraire_nom_famille(nom_complet):
    """Extrait le nom de famille"""
    n = nettoyer_nom_enseignant(nom_complet)
    parts = n.split()
    return parts[0].upper() if parts else ""

def mapper_promotion(promo_edt):
    """Mappe les promotions avec logique robuste"""
    p = str(promo_edt).strip().upper()
    
    mapping_direct = {
        "ING1": "ING1", "ING2RSE": "ING2", "ING3EI": "ING3EI", "ING3RSE": "ING3RSE",
        "ING4EI": "ING4", "ING4RSE": "ING4RSE", "ING5RSE": "ING5RSE",
        "L1MCIL": "L1MCIL", "L2ELT": "L2ELT", "L2MCIL": "MCIL2",
        "L3ELT": "L3ELT", "MCIL2": "MCIL2", "MCIL3": "MCIL3",
        "M1CE": "M1CE", "M1ER": "M1ER", "M1MCIL": "M1MCIL",
        "M1ME": "M1ME", "M1RE": "M1RE",
        "M2CE": "M2CE", "M2ER": "M2ER", "M2MCIL": "M2MCIL",
        "M2ME": "M2ME", "M2RE": "M2RE",
    }
    
    if p in mapping_direct:
        return mapping_direct[p]
    
    # Logique de fallback
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

def format_date_naissance(date_val):
    """Formate les dates de naissance"""
    if pd.isna(date_val):
        return ""
    try:
        if isinstance(date_val, str):
            return date_val
        return pd.Timestamp(date_val).strftime("%d/%m/%Y")
    except:
        return str(date_val)

def detecter_colonnes_etudiant(df):
    """Détecte automatiquement les colonnes importantes"""
    colonnes_mapping = {}
    cols_lower = {col.lower(): col for col in df.columns}
    
    # Mappage des colonnes
    mappings = {
        'nom': ['nom', 'nom_family', 'lastname', 'name'],
        'prenom': ['prenom', 'firstName', 'firstname', 'prénoms'],
        'promotion': ['promotion', 'class', 'niveau', 'level'],
        'email': ['email', 'mail', 'adresse_email'],
        'mat_etud': ['matricule', 'student_id', 'mat_etud'],
        'mat_bac': ['bac', 'mat_bac', 'baccalaurea'],
        'groupe': ['groupe', 'group', 'class_group'],
        'sous_groupe': ['sous_groupe', 'subgroup', 'tp_group'],
        'date_naiss': ['date_naissance', 'date_birth', 'dob'],
        'lieu_naiss': ['lieu_naissance', 'birthplace', 'place']
    }
    
    for key, variations in mappings.items():
        for variation in variations:
            if variation in cols_lower:
                colonnes_mapping[key] = cols_lower[variation]
                break
    
    return colonnes_mapping

# =============================================================================
# SIDEBAR - NAVIGATION PRINCIPALE
# =============================================================================
with st.sidebar:
    st.markdown("""
        <div style='text-align: center; padding: 20px;'>
            <h2 style='color: #1E3A8A; margin: 0;'>🏛️ UDL-SBA</h2>
            <p style='color: #64748b; font-size: 12px; margin: 5px 0 0 0;'>
                Département d'Électrotechnique - FGE
            </p>
        </div>
    """, unsafe_allow_html=True)
    
    st.markdown("---")
    
    module_sel = st.radio(
        "📂 **Choix du module**",
        ["📊 Suivi d'Assiduité", "📅 Gestion des EDTs & Admin", "🧠 EDT Intelligent"],
        index=0,
        key="module_selector"
    )
    
    st.markdown("---")
    st.caption("**Année universitaire** : 2026-2027")
    
    # Infos en bas
    st.markdown("---")
    col1, col2 = st.columns(2)
    with col1:
        if st.button("🔄 Rafraîchir", use_container_width=True):
            st.rerun()
    with col2:
        if st.button("⚙️ Réglages", use_container_width=True):
            st.info("Paramètres : à implémenter")

# =============================================================================
# ROUTEUR PRINCIPAL DES PAGES
# =============================================================================
if module_sel == "📊 Suivi d'Assiduité":
    st.header("📊 Suivi d'Assiduité")
    st.info("Module Suivi d'Assiduité - À implémenter")
    
elif module_sel == "📅 Gestion des EDTs & Admin":
    st.header("📅 Gestion des EDTs & Administration")
    st.info("Module Gestion EDTs - À implémenter")
    
elif module_sel == "🧠 EDT Intelligent":
    st.header("🧠 EDT Intelligent")
    
    # Onglets principaux
    tab1, tab2, tab3 = st.tabs(["📊 Répertoire", "📋 Affichage", "🔍 Recherche"])
    
    with tab1:
        st.subheader("📊 Répertoire des Étudiants")
        
        # Charger les données
        df_etu = charger_fichier_excel(FILE_ETUDIANTS)
        
        if df_etu is not None:
            # Initialiser le filtre
            if "filtre_etudiants" not in st.session_state:
                st.session_state.filtre_etudiants = "TOUS"
            
            # Préparation des données
            df_etu_rep = df_etu.copy()
            cols_map = detecter_colonnes_etudiant(df_etu_rep)
            
            # Créer colonne Nom_Complet si n'existe pas
            if "Nom_Complet" not in df_etu_rep.columns:
                col_n = cols_map.get('nom')
                col_p = cols_map.get('prenom')
                if col_n and col_p and col_n in df_etu_rep.columns and col_p in df_etu_rep.columns:
                    df_etu_rep["Nom_Complet"] = (
                        df_etu_rep[col_n].astype(str).str.strip().str.upper() + " " +
                        df_etu_rep[col_p].astype(str).str.strip().str.title()
                    )
                else:
                    df_etu_rep["Nom_Complet"] = "N/A"
            
            # Normaliser la promotion
            promo_col = cols_map.get('promotion', 'Promotion')
            if promo_col and promo_col in df_etu_rep.columns:
                df_etu_rep["Promotion_Clean"] = df_etu_rep[promo_col].astype(str).str.strip().str.upper()
            else:
                df_etu_rep["Promotion_Clean"] = "NON DEFINI"
            
            # Compteurs
            promo_counts = df_etu_rep["Promotion_Clean"].value_counts().to_dict()
            total_etu = len(df_etu_rep)
            top_promos = sorted(promo_counts.items(), key=lambda x: x[1], reverse=True)
            
            # Afficheurs numériques
            st.markdown("**🎓 Promotions principales :**")
            cols_display = st.columns(min(6, len(top_promos) + 1))
            
            for idx, (promo, count) in enumerate(top_promos[:5]):
                with cols_display[idx]:
                    st.metric(promo, count)
            
            with cols_display[min(5, len(top_promos))]:
                st.metric("TOTAL", total_etu)
            
            # Filtre
            filtre = st.selectbox(
                "🔍 Filtrer par promotion",
                ["TOUS"] + [p[0] for p in top_promos],
                key="filtre_promo"
            )
            st.session_state.filtre_etudiants = filtre
            
            # Appliquer le filtre
            df_filtre = df_etu_rep.copy()
            if filtre != "TOUS":
                df_filtre = df_etu_rep[df_etu_rep["Promotion_Clean"] == filtre]
            
            # Graphiques
            col_g1, col_g2 = st.columns(2)
            with col_g1:
                st.subheader("📈 Répartition")
                chart_data = df_filtre["Promotion_Clean"].value_counts() if filtre == "TOUS" else df_filtre["Promotion_Clean"].value_counts()
                st.bar_chart(chart_data, height=300)
            
            with col_g2:
                st.subheader("📊 Total par promotion")
                effectifs = df_etu_rep["Promotion_Clean"].value_counts().sort_index()
                st.bar_chart(effectifs, height=300, color="#059669")
            
            # Tableau détaillé
            with st.expander(f"🔍 Consulter ({len(df_filtre)} résultat(s))", expanded=True):
                display_cols = []
                for key in ['nom', 'prenom', 'promotion', 'email']:
                    col_name = cols_map.get(key)
                    if col_name and col_name in df_filtre.columns:
                        display_cols.append(col_name)
                
                if not display_cols:
                    display_cols = df_filtre.columns.tolist()[:5]
                
                # Recherche
                search_term = st.text_input("🔎 Rechercher...", key="search_etu")
                df_display = df_filtre[display_cols].copy()
                
                if search_term:
                    mask = df_display.astype(str).apply(
                        lambda row: row.str.contains(search_term, case=False, na=False).any(), axis=1
                    )
                    df_display = df_display[mask]
                
                st.dataframe(df_display, use_container_width=True, height=400)
                
                # Export Excel
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    df_display.to_excel(writer, index=False, sheet_name="Étudiants")
                
                st.download_button(
                    label=f"📥 Télécharger Excel ({len(df_display)} lignes)",
                    data=buffer.getvalue(),
                    file_name=f"Repertoire_Etudiants_{filtre}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
        else:
            st.error("❌ Impossible de charger le fichier des étudiants")
    
    with tab2:
        st.info("Affichage EDT - À implémenter")
    
    with tab3:
        st.info("Recherche avancée - À implémenter")

# =============================================================================
# FOOTER
# =============================================================================
st.markdown("---")
st.markdown("""
    <div style='text-align: center; color: #64748b; font-size: 11px; padding: 10px;'>
        <p>© 2024 Plateforme ELT | Version 2.0 | UDL-SBA</p>
    </div>
""", unsafe_allow_html=True)
