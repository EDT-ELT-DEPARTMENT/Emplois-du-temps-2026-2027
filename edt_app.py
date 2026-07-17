import streamlit as st
import pandas as pd
import os
import hashlib
import io
from datetime import datetime
from supabase import create_client
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import mimetypes
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

# ==========================================
# CONFIGURATION UNIQUE DE LA PAGE
# ==========================================
st.set_page_config(
    page_title="EDT UDL 2027",
    layout="wide",
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

# ==========================================
# CONNEXION BASE DE DONNÉES
# ==========================================
URL = st.secrets["SUPABASE_URL"]
KEY = st.secrets["SUPABASE_KEY"]
supabase = create_client(URL, KEY)

# ==========================================
# FONCTIONS UTILITAIRES
# ==========================================
def hash_pw(password):
    return hashlib.sha256(str.encode(password)).hexdigest()

def normalize(s):
    if not s or s == "Non défini": 
        return "vide"
    s = str(s).strip().lower()
    s = s.replace(" ", "").replace("-", "").replace("–", "")
    s = s.replace(":00", "").replace("h00", "h")
    return s

# ==========================================
# GESTION DU TEMPS
# ==========================================
now = datetime.now()
date_str = now.strftime("%d/%m/%Y")
jours_semaine = [
    "Lundi", "Mardi", "Mercredi", 
    "Jeudi", "Vendredi", "Samedi", "Dimanche"
]
nom_jour_fr = jours_semaine[now.weekday()]

# ==========================================
# STYLE CSS DÉTAILLÉ
# ==========================================
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

# ==========================================
# CONSTANTES ET CONFIGURATION
# ==========================================
NOM_FICHIER_FIXE = "dataEDT-ELT-S1-2027.xlsx"
NOM_FICHIER_CONTACTS = "Permanents-Vacataires-ELT2-2026-2027.xlsx"

horaires_list = [
    "8h - 9h", "8h - 9h30", "8h - 10h", "9h - 10h", "9h30 - 11h", 
    "10h - 11h", "11h - 12h", "11h - 12h30", 
    "12h - 13h", "12h30 - 14h", "13h - 14h30", "14h - 15h30", "14h - 16h", "15h30 - 17h"
]

jours_list = ["Dimanche", "Lundi", "Mardi", "Mercredi", "Jeudi"]

map_h = {normalize(h): h for h in horaires_list}
map_j = {normalize(j): j for j in jours_list}

# ==========================================
# CHARGEMENT DES DONNÉES
# ==========================================
df = None
repertoire_source = {}
repertoire_noms_complets = {}
repertoire_qualites = {}
repertoire_grades = {}

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

# Chargement depuis Supabase (prioritaire) ou fichier local
df = charger_donnees_supabase()

if df is None or df.empty:
    if os.path.exists(NOM_FICHIER_FIXE):
        df = pd.read_excel(NOM_FICHIER_FIXE)
        df.columns = [str(c).strip() for c in df.columns]
        
        colonnes_cles = [
            'Enseignements', 'Code', 'Enseignants', 'Horaire', 
            'Jours', 'Lieu', 'Promotion'
        ]
        
        for col in colonnes_cles:
            if col in df.columns: 
                df[col] = df[col].fillna("Non défini").astype(str).str.strip()
            else:
                df[col] = "Non défini"
                
        df['h_norm'] = df['Horaire'].apply(normalize)
        df['j_norm'] = df['Jours'].apply(normalize)

# Chargement du répertoire contacts
if os.path.exists(NOM_FICHIER_CONTACTS):
    try:
        df_contacts = pd.read_excel(NOM_FICHIER_CONTACTS)
        df_contacts.columns = [str(c).strip() for c in df_contacts.columns]
        
        for _, row in df_contacts.iterrows():
            nom_brut = str(row.get('NOM', '')).strip().upper()
            prenom_brut = str(row.get('PRÉNOM', '')).strip().capitalize()
            email_brut = str(row.get('Email', '')).strip()
            qualite_brute = str(row.get('Qualité', 'Non défini')).strip()
            grade_brut = str(row.get('Grade', 'N/A')).strip()
            
            if nom_brut:
                if email_brut and email_brut.lower() != 'nan':
                    repertoire_source[nom_brut] = email_brut
                repertoire_noms_complets[nom_brut] = f"{nom_brut} {prenom_brut}"
                repertoire_qualites[nom_brut] = qualite_brute
                repertoire_grades[nom_brut] = grade_brut
                
    except Exception as e:
        st.error(f"Erreur lors de la lecture du fichier contacts: {e}")

# ==========================================
# SYSTÈME D'AUTHENTIFICATION
# ==========================================
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
        noms_possibles = sorted(df["Enseignants"].unique()) if df is not None and not df.empty else []
        
        col1, col2 = st.columns(2)
        with col1:
            new_nom = st.selectbox("Sélectionnez votre nom (dans l'EDT)", noms_possibles)
            new_email = st.text_input("Votre adresse Email")
            
        with col2:
            statut_user = st.radio("Statut de l'enseignant", ["Permanent", "Vacataire"], horizontal=True)
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

# ==========================================
# EN-TÊTE HARMONISÉ
# ==========================================
col_logo, col_titre, col_date = st.columns([1, 5, 1.2])

with col_logo:
    try:
        st.image("logo.PNG", width=90)
    except:
        st.markdown("🏛️")

with col_titre:
    st.markdown("<h1 class='main-title' style='border-bottom: none; margin-top: 0;'>Plateforme de gestion des emplois du temps 2026-2027-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA</h1>", unsafe_allow_html=True)

with col_date:
    st.markdown(f"<div class='date-badge' style='float: right;'>📅 {nom_jour_fr}<br>{date_str}</div>", unsafe_allow_html=True)

st.markdown("<div style='border-bottom: 3px solid #D4AF37; margin-bottom: 10px;'></div>", unsafe_allow_html=True)

# ==========================================
# BARRE LATÉRALE
# ==========================================
with st.sidebar:
    st.header(f"👤 {user.get('nom_officiel', 'Utilisateur')}")
    portail = st.selectbox("🚀 Sélectionner Espace", [
        "📖 Emploi du Temps", 
        "📅 Surveillances Examens", 
        "🤖 Générateur Automatique", 
        "👥 Portail Enseignants", 
        "🎓 Portail mise à jour EDT", 
        "📢 Gestion Administrative - Bordereaux & PVs"
    ])
    st.divider()
    
    mode_view = "Personnel"
    poste_sup = False
    
    if portail == "📖 Emploi du Temps":
        if is_admin:
            mode_view = st.radio("Vue Administration :", [
                "Promotion", "Enseignant", "🏢 Planning Salles", 
                "🚩 Vérificateur de conflits", "✍️ Éditeur de données"
            ])
        else:
            mode_view = "Personnel"
        poste_sup = st.checkbox("Poste Supérieur (Décharge 3h)")
        
    if st.button("🚪 Déconnexion du compte"):
        st.session_state["user_data"] = None
        st.rerun()

st.markdown(f"<div class='portal-badge'>MODE ACTIF : {portail.upper()}</div>", unsafe_allow_html=True)

# ==========================================
# FONCTIONS DE FORMATAGE COMMUNES
# ==========================================
def format_case(rows):
    items = []
    for _, r in rows.iterrows():
        nat = '📘' if 'COURS' in str(r['Code']).upper() else '📗' if 'TD' in str(r['Code']).upper() else '🔴'
        txt = f"<div style='margin-bottom:8px;'>{nat} <b>{r['Enseignements']}</b><br><small>({r['Code']})</small><br><i>{r['Lieu']}</i><br><b>{r['Promotion']}</b></div>"
        items.append(txt)
    return "<div class='separator'></div>".join(items)

# ==========================================
# LOGIQUE PRINCIPALE
# ==========================================
if df is not None and not df.empty:
    
    # ------------------------------------------------------------------
    # PORTAIL : EMPLOI DU TEMPS
    # ------------------------------------------------------------------
    if portail == "📖 Emploi du Temps":
        
        # --- VUE PERSONNEL / ENSEIGNANT ---
        if mode_view == "Personnel" or (is_admin and mode_view == "Enseignant"):
            
            if mode_view == "Personnel":
                cible = user['nom_officiel']
                nom_affichage_complet = repertoire_noms_complets.get(cible.strip().upper(), cible)
            else:
                noms_bruts = sorted(df["Enseignants"].unique())
                options_affichage = [repertoire_noms_complets.get(n.strip().upper(), n) for n in noms_bruts]
                inverse_map = {repertoire_noms_complets.get(n.strip().upper(), n): n for n in noms_bruts}
                
                choix_utilisateur = st.selectbox("Sélectionner l'Enseignant :", options=options_affichage, index=0)
                cible = inverse_map[choix_utilisateur]
                nom_affichage_complet = choix_utilisateur
            
            # Filtrage et calculs
            df_f = df[df["Enseignants"].str.contains(cible, case=False, na=False)].copy()
            df_f['Type'] = df_f['Code'].apply(
                lambda x: "COURS" if "COURS" in str(x).upper() else ("TD" if "TD" in str(x).upper() else "TP")
            )
            df_u = df_f.drop_duplicates(subset=['j_norm', 'h_norm'])
            
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
            
            statut_enseignant = repertoire_qualites.get(cible.strip().upper(), "Statut inconnu")
            grade_enseignant = repertoire_grades.get(cible.strip().upper(), "Grade inconnu")
            color_statut = "#2ecc71" if "PERMANENT" in statut_enseignant.upper() else "#e67e22"
            color_grade = "#3498db"
            
            # Affichage titre
            st.markdown(f"""
                <div style="display: flex; align-items: center; flex-wrap: wrap; gap: 10px; margin-bottom: 20px;">
                    <h3 style="margin: 0;">📊 Charge Horaire hebdomadaire : {nom_affichage_complet}</h3>
                    <span style="background-color: {color_grade}; color: white; padding: 3px 12px; 
                                 border-radius: 15px; font-size: 0.8em; font-weight: bold;">
                        {grade_enseignant}
                    </span>
                    <span style="background-color: {color_statut}; color: white; padding: 3px 12px; 
                                 border-radius: 15px; font-size: 0.8em; font-weight: bold;">
                        {statut_enseignant}
                    </span>
                </div>
            """, unsafe_allow_html=True)
            
            # Compteurs
            st.markdown(f"""
                <div class="stat-container">
                    <div class="stat-box bg-cours">📘 {nb_cours} Cours</div>
                    <div class="stat-box bg-td">📗 {nb_td} TD</div>
                    <div class="stat-box bg-tp">🔴 {nb_tp} TP</div>
                </div>
            """, unsafe_allow_html=True)
            
            if h_sup < 0:
                st.warning(f"⚠️ Attention : Sous-charge détectée de {abs(delta_eq)} eq/h par rapport au seuil.")

            c1, c2, c3 = st.columns(3)
            with c1:
                st.markdown(f"<div class='metric-card'>Charge Effective<br><h2>{round(charge_effective, 2)} h</h2></div>", unsafe_allow_html=True)
            with c2:
                st.markdown(f"<div class='metric-card'>Seuil Réglementaire<br><h2>{seuil_obligatoire} eq/h</h2></div>", unsafe_allow_html=True)
            with c3:
                color_res = "#2ecc71" if h_sup >= 0 else "#e74c3c"
                label_res = "Heures Sup. Réelles" if h_sup >= 0 else "Déficit Horaire"
                st.markdown(f"""
                    <div class='metric-card' style='border-bottom: 5px solid {color_res};'>
                        {label_res}<br>
                        <h2 style='color: {color_res};'>{h_sup_formattee}</h2>
                    </div>
                """, unsafe_allow_html=True)
            
            if h_sup > 0:
                st.caption(f"✅ L'enseignant a complété sa charge et totalise {round(h_sup, 2)}h en supplément.")
            elif h_sup < 0:
                st.caption(f"⚠️ Attention : Sous-charge détectée de {round(abs(h_sup), 2)}h par rapport au seuil.")
            else:
                st.caption("⚖️ Service réglementaire exactement rempli (Pile 6.0 eq/h).")

            # --- EXPORT GLOBAL ADMIN ---
            if is_admin:
                st.markdown("---")
                if st.button("📑 Préparer le Bilan Global (Tous les enseignants)", use_container_width=True):
                    liste_profs = sorted(df["Enseignants"].unique())
                    recap_data = []

                    for p in liste_profs:
                        df_p = df[df["Enseignants"].str.contains(p, case=False, na=False)].copy()
                        df_p['Type'] = df_p['Code'].apply(
                            lambda x: "COURS" if "COURS" in str(x).upper() else ("TD" if "TD" in str(x).upper() else "TP")
                        )
                        df_up = df_p.drop_duplicates(subset=['j_norm', 'h_norm'])
                        
                        nom_complet = repertoire_noms_complets.get(p.strip().upper(), p)
                        grade_ens = repertoire_grades.get(p.strip().upper(), "N/A")
                        qualite_ens = repertoire_qualites.get(p.strip().upper(), "Non spécifié")
                        
                        n_co = len(df_up[df_up['Type'] == 'COURS'])
                        n_td = len(df_up[df_up['Type'] == 'TD'])
                        n_tp = len(df_up[df_up['Type'] == 'TP'])
                        
                        s_oblig = 6.0 
                        c_eq = (n_co * 1.5) + (n_td + n_tp)
                        b_h = (c_eq - s_oblig) * 1.5
                        c_eff = (n_co + n_td + n_tp) * 1.5

                        recap_data.append({
                            "Enseignant": nom_complet,
                            "Grade": grade_ens,
                            "Qualité": qualite_ens,
                            "Cours": n_co,
                            "TD": n_td,
                            "TP": n_tp,
                            "Charge Effective (h)": c_eff,
                            "Total (Eq)": c_eq,
                            "Heures Sup. Réelles/Déficit Horaire": round(b_h, 2)
                        })

                    df_global = pd.DataFrame(recap_data)
                    buf = io.BytesIO()
                    with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
                        df_global.to_excel(writer, index=False, sheet_name='Bilan_Global_Charges')
                        workbook = writer.book
                        worksheet = writer.sheets['Bilan_Global_Charges']
                        format_rouge = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
                        
                        for i, col in enumerate(df_global.columns):
                            max_len = max(df_global[col].astype(str).map(len).max(), len(col)) + 2
                            worksheet.set_column(i, i, max_len)
                        
                        last_row = len(df_global)
                        worksheet.conditional_format(1, 7, last_row, 7, {
                            'type': 'cell',
                            'criteria': '<',
                            'value': 0,
                            'format': format_rouge
                        })
                        worksheet.freeze_panes(1, 0)

                    st.download_button(
                        label="📥 Télécharger le fichier Excel Global",
                        data=buf.getvalue(),
                        file_name="Bilan_Global_Charges_2027.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
            
            # --- EMPLOI DU TEMPS INDIVIDUEL ---
            st.divider()
            st.markdown("### 📅 Emploi du Temps Individuel")
            
            nom_complet_entete = repertoire_noms_complets.get(cible.strip().upper(), cible)
            grade_entete = repertoire_grades.get(cible.strip().upper(), "Grade non spécifié")
            statut_entete = repertoire_qualites.get(cible.strip().upper(), "Statut non spécifié")
            
            if not df_f.empty:
                grid = df_f.groupby(['h_norm', 'j_norm']).apply(format_case, include_groups=False).unstack('j_norm')
                grid = grid.reindex(index=[normalize(h) for h in horaires_list], columns=[normalize(j) for j in jours_list]).fillna("")
                grid.index = [map_h.get(i, i) for i in grid.index]
                grid.columns = [map_j.get(c, c) for c in grid.columns]
                st.write(grid.to_html(escape=False), unsafe_allow_html=True)

                # Boutons de téléchargement
                st.markdown("---")
                col_dl1, col_dl2, col_dl3 = st.columns(3)
                
                # Excel
                buf_ex = io.BytesIO()
                with pd.ExcelWriter(buf_ex, engine='xlsxwriter') as writer:
                    df_f.drop(columns=['h_norm', 'j_norm', 'Type'], errors='ignore').to_excel(writer, index=False, sheet_name='Mon EDT')
                col_dl1.download_button(
                    label="📥 Liste Excel",
                    data=buf_ex.getvalue(),
                    file_name=f"EDT_Individuel_{nom_complet_entete.replace(' ', '_')}_2027.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="btn_indiv_xl_pro_v1"
                )
                
                # HTML
                html_content = df_f.drop(columns=['h_norm', 'j_norm', 'Type'], errors='ignore').to_html(index=False)
                col_dl2.download_button(
                    label="🌐 Tableau (HTML)",
                    data=f"<!DOCTYPE html><html><head><meta charset='UTF-8'></head><body>{html_content}</body></html>",
                    file_name="EDT_Individuel_2027.html",
                    mime="text/html",
                    use_container_width=True,
                    key="btn_indiv_html_final_v12"
                )
                
                # PDF
                try:
                    from fpdf import FPDF
                    
                    class INDIV_PDF(FPDF):
                        def header(self):
                            self.set_font('Arial', 'B', 10)
                            t = "Plateforme de gestion des EDTs-S2-2027-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA"
                            self.cell(0, 8, t.encode('latin-1', 'replace').decode('latin-1'), 0, 1, 'C')
                            self.ln(2)
                    
                    def clean_indiv(text_val):
                        if not text_val: return ""
                        t = str(text_val)
                        t = t.replace('<b>','').replace('</b>','')
                        t = t.replace("’", "'").replace("‘", "'")
                        return t.encode('latin-1', 'replace').decode('latin-1')
                    
                    pdf = INDIV_PDF(orientation="L", unit="mm", format="A4")
                    pdf.set_margins(7, 10, 7)
                    pdf.add_page()
                    pdf.set_font("Arial", "B", 12)
                    pdf.cell(0, 10, f"EMPLOI DU TEMPS : {nom_complet_entete}".encode('latin-1', 'replace').decode('latin-1'), 0, 1, "C")
                    pdf.ln(3)
                    
                    # Tableau simple
                    col_w = (pdf.w - 20) / 7
                    pdf.set_font("Arial", "B", 8)
                    headers = ['Enseignements', 'Code', 'Enseignants', 'Horaire', 'Jours', 'Lieu', 'Promotion']
                    for h in headers:
                        pdf.cell(col_w, 8, h.encode('latin-1', 'replace').decode('latin-1'), 1, 0, 'C')
                    pdf.ln()
                    
                    pdf.set_font("Arial", "", 7)
                    for _, row in df_f.iterrows():
                        for h in headers:
                            val = clean_indiv(str(row.get(h, '')))
                            pdf.cell(col_w, 6, val, 1, 0, 'C')
                        pdf.ln()
                    
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
            else:
                st.info("Aucune donnée trouvée pour cet enseignant.")

        # --- VUE PROMOTION (ADMIN) ---
        elif is_admin and mode_view == "Promotion":
            p_sel = st.selectbox("Choisir Promotion :", sorted(df["Promotion"].unique()))
            df_p = df[df["Promotion"] == p_sel].copy()
            
            df_p['Type_Tmp'] = df_p['Code'].apply(
                lambda x: "COURS" if "COURS" in str(x).upper() else ("TD" if "TD" in str(x).upper() else "TP")
            )
            df_stats = df_p.drop_duplicates(subset=['j_norm', 'h_norm'])
            n_p_co = len(df_stats[df_stats['Type_Tmp'] == 'COURS'])
            n_p_td = len(df_stats[df_stats['Type_Tmp'] == 'TD'])
            n_p_tp = len(df_stats[df_stats['Type_Tmp'] == 'TP'])

            def fmt_p(rows):
                items = []
                for _, r in rows.iterrows():
                    code_up = str(r['Code']).upper()
                    if 'COURS' in code_up:
                        nat, color = '📘', '#1e40af'
                    elif 'TD' in code_up:
                        nat, color = '📗', '#166534'
                    else:
                        nat, color = '🔴', '#991b1b'
                    txt = f"""
                    <div style='margin-bottom:8px; padding:5px; border-left:3px solid {color}; background-color:#f8fafc; border-radius:4px;'>
                        <b style='color:{color};'>{nat} {r['Enseignements']}</b><br>
                        <span style='font-size:11px; font-weight:bold;'>👤 {r['Enseignants']}</span><br>
                        <span style='font-size:11px;'>📍 {r['Lieu']}</span>
                    </div>
                    """
                    items.append(txt)
                return "".join(items)

            grid_p = df_p.groupby(['h_norm', 'j_norm']).apply(fmt_p, include_groups=False).unstack('j_norm')
            idx_h = [normalize(h) for h in horaires_list]
            cols_j = [normalize(j) for j in jours_list]
            grid_p = grid_p.reindex(index=idx_h, columns=cols_j).fillna("")
            grid_p = grid_p[grid_p.any(axis=1)]
            grid_p.index = [map_h.get(i, i) for i in grid_p.index]
            grid_p.columns = [map_j.get(c, c) for c in grid_p.columns]

            st.write(grid_p.to_html(escape=False), unsafe_allow_html=True)

            # Stats
            df_unique_matieres = df_p.drop_duplicates(subset=['Enseignements', 'Code'])
            total_p_cours = len(df_unique_matieres[df_unique_matieres['Code'].str.contains('COURS', case=False, na=False)])
            total_p_td = len(df_unique_matieres[df_unique_matieres['Code'].str.contains('TD', case=False, na=False)])
            total_p_tp = len(df_unique_matieres[~df_unique_matieres['Code'].str.contains('COURS|TD', case=False, na=False)])
            total_p_seances = total_p_cours + total_p_td + total_p_tp

            st.markdown(f"""
            <div style='display: flex; justify-content: space-around; background: linear-gradient(135deg, #f8fafc, #e2e8f0); padding: 20px; border-radius: 15px; border: 1px solid #cbd5e1; margin-bottom: 25px; text-align: center; box-shadow: 0 4px 15px rgba(0,0,0,0.08);'>
                <div><div style='font-size: 11px; color: #64748b; text-transform: uppercase; font-weight: bold;'>Total Cours</div>
                <div style='font-size: 26px; font-weight: 800; color: #1e40af; margin-top: 5px;'>📘 {total_p_cours}</div></div>
                <div style='border-left: 2px solid #cbd5e1; height: 45px;'></div>
                <div><div style='font-size: 11px; color: #64748b; text-transform: uppercase; font-weight: bold;'>Total TD</div>
                <div style='font-size: 26px; font-weight: 800; color: #166534; margin-top: 5px;'>📗 {total_p_td}</div></div>
                <div style='border-left: 2px solid #cbd5e1; height: 45px;'></div>
                <div><div style='font-size: 11px; color: #64748b; text-transform: uppercase; font-weight: bold;'>Total TP</div>
                <div style='font-size: 26px; font-weight: 800; color: #991b1b; margin-top: 5px;'>🔴 {total_p_tp}</div></div>
                <div style='border-left: 2px solid #cbd5e1; height: 45px;'></div>
                <div><div style='font-size: 11px; color: #64748b; text-transform: uppercase; font-weight: bold;'>Total Séances</div>
                <div style='font-size: 26px; font-weight: 800; color: #0f3460; margin-top: 5px;'>📊 {total_p_seances}</div></div>
            </div>
            """, unsafe_allow_html=True)

            # Téléchargements
            st.markdown("---")
            st.markdown("### 📥 Téléchargements de l'Emploi du Temps")
            cp1, cp2, cp3 = st.columns(3)
            
            with cp1:
                buf_p = io.BytesIO()
                with pd.ExcelWriter(buf_p, engine='xlsxwriter') as writer:
                    df_p.drop(columns=['h_norm', 'j_norm', 'Type_Tmp'], errors='ignore').to_excel(writer, index=False, sheet_name=f'EDT {p_sel}')
                cp1.download_button(
                    label=f"📊 Excel",
                    data=buf_p.getvalue(),
                    file_name=f"EDT_{p_sel}_2027.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key=f"btn_xl_v10_{p_sel}" 
                )
            
            with cp2:
                html_p = grid_p.to_html(escape=False)
                cp2.download_button(
                    label="🌐 HTML",
                    data=f"<!DOCTYPE html><html><head><meta charset='UTF-8'></head><body>{html_p}</body></html>",
                    file_name=f"EDT_{p_sel}_2027.html",
                    mime="text/html",
                    use_container_width=True,
                    key=f"btn_html_v10_{p_sel}"
                )
            
            with cp3:
                try:
                    from fpdf import FPDF
                    pdf = FPDF(orientation="L", unit="mm", format="A4")
                    pdf.add_page()
                    pdf.set_font("Arial", "B", 12)
                    pdf.cell(0, 10, f"EDT {p_sel}".encode('latin-1', 'replace').decode('latin-1'), 0, 1, "C")
                    pdf.set_font("Arial", "", 8)
                    col_w = (pdf.w - 20) / (len(grid_p.columns) + 1)
                    
                    # Header
                    pdf.cell(col_w, 8, "Horaire", 1, 0, 'C')
                    for col in grid_p.columns:
                        pdf.cell(col_w, 8, str(col).encode('latin-1', 'replace').decode('latin-1'), 1, 0, 'C')
                    pdf.ln()
                    
                    for idx, row in grid_p.iterrows():
                        pdf.cell(col_w, 8, str(idx).encode('latin-1', 'replace').decode('latin-1'), 1, 0, 'C')
                        for val in row:
                            clean = str(val).replace('<br>', '\n').replace('<div', '').replace('</div>', '').replace('<b>', '').replace('</b>', '')[:50]
                            pdf.cell(col_w, 8, clean.encode('latin-1', 'replace').decode('latin-1'), 1, 0, 'C')
                        pdf.ln()
                    
                    cp3.download_button(
                        label="📄 PDF",
                        data=bytes(pdf.output()),
                        file_name=f"EDT_{p_sel}_2027.pdf",
                        mime="application/pdf",
                        use_container_width=True,
                        key=f"btn_pdf_v10_{p_sel}" 
                    )
                except Exception as e:
                    cp3.error(f"Erreur PDF : {e}")

        # --- VUE PLANNING SALLES (ADMIN) ---
        elif is_admin and mode_view == "🏢 Planning Salles":
            s_sel = st.selectbox("Choisir Salle :", sorted(df["Lieu"].unique()))
            df_s = df[df["Lieu"] == s_sel]
            
            def fmt_s(rows):
                items = [f"<b>{r['Promotion']}</b><br>{r['Enseignements']}<br><i>{r['Enseignants']}</i>" for _, r in rows.iterrows()]
                return "<div class='separator'></div>".join(items)
                
            grid_s = df_s.groupby(['h_norm', 'j_norm']).apply(fmt_s, include_groups=False).unstack('j_norm')
            grid_s = grid_s.reindex(index=[normalize(h) for h in horaires_list], columns=[normalize(j) for j in jours_list]).fillna("")
            grid_s.index = [map_h.get(i, i) for i in grid_s.index]
            grid_s.columns = [map_j.get(c, c) for c in grid_s.columns]
            
            st.write(grid_s.to_html(escape=False), unsafe_allow_html=True)

            st.markdown("---")
            cs1, cs2 = st.columns(2)
            
            buf_s = io.BytesIO()
            df_s.drop(columns=['h_norm', 'j_norm'], errors='ignore').to_excel(buf_s, index=False)
            cs1.download_button(
                label=f"📥 Liste {s_sel} (Excel)",
                data=buf_s.getvalue(),
                file_name=f"Planning_{s_sel}_2027.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key=f"xl_salle_{s_sel}"
            )
            
            try:
                from fpdf import FPDF
                pdf = FPDF(orientation="L", unit="mm", format="A4")
                pdf.add_page()
                pdf.set_font("Arial", "B", 12)
                pdf.cell(0, 10, f"PLANNING SALLE : {s_sel}".encode('latin-1', 'replace').decode('latin-1'), 0, 1, "C")
                pdf.set_font("Arial", "", 8)
                col_w = (pdf.w - 20) / (len(grid_s.columns) + 1)
                
                pdf.cell(col_w, 8, "Horaire", 1, 0, 'C')
                for col in grid_s.columns:
                    pdf.cell(col_w, 8, str(col).encode('latin-1', 'replace').decode('latin-1'), 1, 0, 'C')
                pdf.ln()
                
                for idx, row in grid_s.iterrows():
                    pdf.cell(col_w, 8, str(idx).encode('latin-1', 'replace').decode('latin-1'), 1, 0, 'C')
                    for val in row:
                        clean = str(val).replace('<br>', ' ').replace('<div', '').replace('</div>', '').replace('<b>', '').replace('</b>', '').replace('<i>', '').replace('</i>', '')[:50]
                        pdf.cell(col_w, 8, clean.encode('latin-1', 'replace').decode('latin-1'), 1, 0, 'C')
                    pdf.ln()
                
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

        # --- VUE VÉRIFICATEUR DE CONFLITS (ADMIN) ---
        elif is_admin and mode_view == "🚩 Vérificateur de conflits":
            st.subheader("🚩 Analyse des Conflits Individuels")
            st.markdown("---")
            
            errs_text = []      
            errs_for_df = []    

            # A. CONFLITS D'ENSEIGNANTS
            p_groups = df[df["Enseignants"] != "Non défini"].groupby(['Jours', 'Horaire', 'Enseignants'])
            for (jour, horaire, prof), group in p_groups:
                lieux_uniques = group['Lieu'].unique()
                matieres_uniques = group['Enseignements'].unique()
                if len(lieux_uniques) > 1 or len(matieres_uniques) > 1:
                    type_err = "❌ CONFLIT ENSEIGNANT"
                    detail = f"L'enseignant est affecté à plusieurs lieux ({', '.join(lieux_uniques)}) ou matières."
                    msg = f"**{type_err}** : {prof} | {jour} {horaire}"
                    errs_text.append(("error", msg))
                    errs_for_df.append({
                        "Type": type_err, "Enseignant": prof, "Jour": jour, "Horaire": horaire, 
                        "Détail": detail, "Lieu": ", ".join(lieux_uniques), 
                        "Matières": ", ".join(matieres_uniques), "Promotions": ", ".join(group['Promotion'].unique())
                    })

            # B. CONFLITS DE SALLES
            s_groups = df[(df["Lieu"] != "Non défini") & (df["Lieu"] != "A distance")].groupby(['Jours', 'Horaire', 'Lieu'])
            for (jour, horaire, lieu), group in s_groups:
                if len(group['Enseignants'].unique()) > 1:
                    type_err = "❌ CONFLIT SALLE OCCUPÉE"
                    profs_concernees = group['Enseignants'].unique()
                    detail = f"La salle '{lieu}' est utilisée par : {', '.join(profs_concernees)}"
                    msg = f"**{type_err}** : {lieu} | {jour} {horaire} ({', '.join(profs_concernees)})"
                    errs_text.append(("error", msg))
                    for p in profs_concernees:
                        errs_for_df.append({
                            "Type": type_err, "Enseignant": p, "Jour": jour, "Horaire": horaire, 
                            "Détail": detail, "Lieu": lieu, 
                            "Matières": ", ".join(group['Enseignements'].unique()), 
                            "Promotions": ", ".join(group['Promotion'].unique())
                        })

            # C. CONFLITS DE PROMOTION
            pr_groups = df[df["Promotion"] != "Non défini"].groupby(['Jours', 'Horaire', 'Promotion'])
            for (jour, horaire, promo), group in pr_groups:
                if len(group['Enseignements'].unique()) > 1:
                    type_err = "⚠️ CONFLIT PROMOTION"
                    matieres = group['Enseignements'].unique()
                    detail = f"La promotion {promo} a plusieurs cours simultanés : {', '.join(matieres)}"
                    msg = f"**{type_err}** : {promo} | {jour} {horaire}"
                    errs_text.append(("warning", msg))
                    errs_for_df.append({
                        "Type": type_err, "Enseignant": "Multi-enseignants", "Jour": jour, "Horaire": horaire, 
                        "Détail": detail, "Lieu": ", ".join(group['Lieu'].unique()), 
                        "Matières": ", ".join(matieres), "Promotions": promo
                    })

            if errs_for_df:
                st.markdown("### 🔍 Résolution ciblée")
                profs_en_conflit = sorted(list(set([e["Enseignant"] for e in errs_for_df])))
                options_menu = ["Tous"] + profs_en_conflit

                if "filtre_prof_conflit" not in st.session_state:
                    st.session_state.filtre_prof_conflit = "Tous"

                selected_prof = st.selectbox(
                    "🎯 Filtrer par enseignant :", 
                    options=options_menu,
                    key="filtre_prof_conflit"
                )

                if selected_prof != "Tous":
                    if st.button("🔄 Réinitialiser la vue (Afficher tout)", use_container_width=True):
                        if "filtre_prof_conflit" in st.session_state:
                            del st.session_state.filtre_prof_conflit
                        st.rerun()

                st.divider()

                if selected_prof != "Tous":
                    st.info(f"Analyse précise pour : **{selected_prof}**")
                    conflits_specifiques = [e for e in errs_for_df if e["Enseignant"] == selected_prof]
                    for i, cp in enumerate(conflits_specifiques):
                        with st.expander(f"📌 {cp['Type']} - {cp['Jour']} {cp['Horaire']}", expanded=True):
                            st.error(f"**Problème :** {cp['Détail']}")
                            st.markdown("💡 **Solutions suggérées :**")
                            st.write("- Vérifiez que le nom de la matière est identique pour les deux groupes.")
                            st.write("- Modifiez l'horaire ou la salle dans l'éditeur de données.")

                st.markdown("### 🌍 Rapport Global des Anomalies")
                for style, m in errs_text:
                    if selected_prof == "Tous" or selected_prof in m:
                        if style == "error":
                            st.error(m)
                        else:
                            st.warning(m)

                # Assistant de résolution
                st.divider()
                st.subheader("💡 Assistant de Résolution Intelligent")
                st.info("L'assistant propose des créneaux libres (Horaire + Salle) en respectant le type de lieu initial.")

                tous_les_lieux = sorted([l for l in df['Lieu'].unique() if str(l) != "nan" and l != "Non défini"])
                solutions_finales = []

                for i, cp in enumerate(errs_for_df):
                    with st.expander(f"📍 Conflit n°{i+1} : {cp['Enseignant']} ({cp['Jour']} - {cp['Horaire']})", expanded=True):
                        c1, c2 = st.columns([2, 1])
                        with c1:
                            st.error(f"**Anomalie :** {cp['Détail']}")
                            st.caption(f"Matières impliquées : {cp.get('Matières', 'N/A')}")
                        with c2:
                            lieu_initial = str(cp['Lieu']).upper()
                            est_tp = any(keyword in lieu_initial for keyword in ["LABO", "TP", "ATELIER", "CC", "MICRO"])
                            est_amphi = "AMPHI" in lieu_initial or "A0" in lieu_initial
                            
                            lieux_compatibles = []
                            for l in tous_les_lieux:
                                l_str = str(l).upper()
                                if est_tp and any(k in l_str for k in ["LABO", "TP", "CC", "MICRO"]):
                                    lieux_compatibles.append(l)
                                elif est_amphi and ("AMPHI" in l_str or "A0" in l_str):
                                    lieux_compatibles.append(l)
                                elif not est_tp and not est_amphi and ("S" in l_str or "SALLE" in l_str):
                                    lieux_compatibles.append(l)

                            tous_horaires = ["8h - 9h30", "9h30 - 11h", "11h - 12h30", "12h30 - 14h", "14h - 15h30", "15h30 - 17h"]
                            suggestions_valides = []
                            
                            for hor in tous_horaires:
                                prof_occupe = False
                                if cp['Enseignant'] not in ["ND", "Multi-enseignants"]:
                                    prof_occupe = not df[(df['Jours'] == cp['Jour']) & 
                                                         (df['Horaire'] == hor) & 
                                                         (df['Enseignants'] == cp['Enseignant'])].empty
                                
                                if not prof_occupe:
                                    lieux_occupes = df[(df['Jours'] == cp['Jour']) & 
                                                       (df['Horaire'] == hor)]['Lieu'].unique()
                                    libres = [l for l in lieux_compatibles if l not in lieux_occupes]
                                    for salle_libre in libres:
                                        if not (hor == cp['Horaire'] and salle_libre in cp['Lieu']):
                                            suggestions_valides.append(f"{hor} en {salle_libre}")

                            choix_sol = st.selectbox(
                                "🚀 Solution (Heure + Lieu compatible) :",
                                options=["-- Garder actuel --"] + suggestions_valides[:30],
                                key=f"assistant_sol_{i}",
                                help="Propose uniquement des créneaux où l'enseignant et la salle sont libres."
                            )
                        
                        solutions_finales.append({
                            "Type de Conflit": cp['Type'],
                            "Personne/Salle concernée": cp['Enseignant'] if cp['Enseignant'] != "Multi-enseignants" else cp['Détail'],
                            "Jour": cp['Jour'],
                            "Horaire Initial": cp['Horaire'],
                            "Lieu Initial": cp['Lieu'],
                            "SOLUTION PROPOSÉE": choix_sol if choix_sol != "-- Garder actuel --" else "À CORRIGER MANUELLEMENT"
                        })

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
                        file_name="Solutions_Conflits_EDT_S2_2027.xlsx",
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

        # --- VUE ÉDITEUR DE DONNÉES (ADMIN) ---
        elif is_admin and mode_view == "✍️ Éditeur de données":
            st.divider()
            st.subheader("✍️ Plateforme de gestion des EDTs-S2-2027-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA")

            cols_format = ['Enseignements', 'Code', 'Enseignants', 'Horaire', 'Jours', 'Lieu', 'Promotion', 'Chevauchement']

            if 'df_admin' not in st.session_state:
                if df is not None:
                    temp_df = df.copy()
                    for col in cols_format:
                        if col not in temp_df.columns:
                            temp_df[col] = ""
                        temp_df[col] = temp_df[col].astype(str).replace(['nan', 'None', '<NA>'], '')
                    st.session_state.df_admin = temp_df
                else:
                    st.warning("Le DataFrame est vide ou non initialisé.")
                    st.stop()

            horaires_ref = ["8h - 9h30", "9h30 - 11h", "11h - 12h30", "12h30 - 14h00", "14h00 - 15h30", "15h30 - 17h00"]
            h_existants = [h for h in st.session_state.df_admin["Horaire"].unique() if h and h.strip() != ""]
            liste_horaires = sorted(list(set(h_existants + horaires_ref)))
            jours_std = ["Dimanche", "Lundi", "Mardi", "Mercredi", "Jeudi"]
            promos_existantes = [p for p in st.session_state.df_admin["Promotion"].unique() if p and p.strip() != ""]

            # Filtre de recherche
            st.markdown("### 🔍 Filtrer par Enseignant")
            search_prof = st.text_input("Tapez le nom de l'enseignant pour filtrer le tableau :", "")

            if search_prof:
                df_to_edit = st.session_state.df_admin[
                    st.session_state.df_admin["Enseignants"].str.contains(search_prof, case=False, na=False)
                ]
                st.info(f"💡 Affichage des cours de : **{search_prof}**.")
            else:
                df_to_edit = st.session_state.df_admin

            # Formulaire d'ajout
            st.markdown("### 🌍 Tableau d'édition")
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
                            promo_conflit = conflit_salle.iloc[0]['Promotion']
                            prof_conflit = conflit_salle.iloc[0]['Enseignants']
                            st.error(f"❌ CONFLIT SALLE : La salle {n_lieu} est déjà prise par **{prof_conflit}** pour la promotion **{promo_conflit}**.")
                        
                        elif not conflit_prof.empty:
                            promo_conflit = conflit_prof.iloc[0]['Promotion']
                            lieu_conflit = conflit_prof.iloc[0]['Lieu']
                            st.error(f"❌ CONFLIT ENSEIGNANT : M. {n_prof} a déjà un cours avec la promotion **{promo_conflit}** en salle {lieu_conflit}.")
                        
                        else:
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
                                supabase.table("edt_data").insert(nouvelle_ligne_db).execute()
                                st.success(f"✅ Félicitations ! Le cours de {n_ensg} pour la promotion {n_promo} est désormais enregistré dans la base de données Cloud.")
                                
                                if 'df_admin' in st.session_state:
                                    del st.session_state.df_admin
                                st.rerun()
                                
                            except Exception as e:
                                st.error(f"❌ Erreur technique lors de l'enregistrement : {e}")

            # Éditeur de tableau
            st.markdown("### 📝 Modification des données")
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

            if edited_df is not None and not edited_df.equals(df_to_edit[cols_format]):
                if search_prof:
                    indices_modifies = df_to_edit.index
                    df_others = st.session_state.df_admin.drop(indices_modifies)
                    st.session_state.df_admin = pd.concat([df_others, edited_df], ignore_index=True)
                else:
                    st.session_state.df_admin = edited_df

            # Analyse visuelle des chevauchements
            st.divider()
            st.markdown("### 🔍 Analyse Visuelle des Chevauchements")

            def afficher_grille_anomalie(df_source, type_tri):
                jours_ordre = ["Dimanche", "Lundi", "Mardi", "Mercredi", "Jeudi"]
                horaires_ordre = [
                    "8h - 9h", "8h - 9h30", "8h - 10h", "9h - 10h", "9h30 - 11h", 
                    "10h - 11h", "11h - 12h", "11h - 12h30", "12h - 13h", 
                    "12h30 - 14h", "13h - 14h30", "14h - 15h30", "14h - 16h", "15h30 - 17h"
                ]
                
                grid = pd.DataFrame("", index=horaires_ordre, columns=jours_ordre)
                df_temp = df_source.copy()
                
                def format_horaire(h):
                    h_str = str(h).replace(" ", "").lower()
                    for target in horaires_ordre:
                        if h_str == target.replace(" ", "").lower(): 
                            return target
                    return h

                df_temp['Horaire_Normalise'] = df_temp['Horaire'].apply(format_horaire)
                df_temp['Jours'] = df_temp['Jours'].astype(str).str.strip().str.capitalize()

                doublons = df_temp.duplicated(subset=['Jours', 'Horaire_Normalise', type_tri], keep=False)
                mask_valid = (df_temp[type_tri].astype(str).str.len() > 1) & (df_temp[type_tri].astype(str).str.lower() != "nan")
                df_conflits = df_temp[doublons & mask_valid].copy()
                
                if not df_conflits.empty:
                    for _, row in df_conflits.iterrows():
                        idx_h = row['Horaire_Normalise']
                        col_j = row['Jours']
                        
                        if idx_h in horaires_ordre and col_j in jours_ordre:
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

            t_salle, t_prof, t_promo = st.tabs(["🏢 Conflits Salles", "👤 Conflits Enseignants", "🎓 Conflits Promotions"])
            
            with t_salle:
                afficher_grille_anomalie(st.session_state.df_admin, "Lieu")
            with t_prof:
                afficher_grille_anomalie(st.session_state.df_admin, "Enseignants")
            with t_promo:
                afficher_grille_anomalie(st.session_state.df_admin, "Promotion")

            # Sauvegarde et export
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
                df_complet = st.session_state.df_admin.copy()
                conflits_list = []

                doublons_salle = df_complet.duplicated(subset=['Jours', 'Horaire', 'Lieu'], keep=False) & (df_complet['Lieu'].astype(str).str.len() > 1)
                doublons_prof = df_complet.duplicated(subset=['Jours', 'Horaire', 'Enseignants'], keep=False) & (df_complet['Enseignants'] != "ND") & (df_complet['Enseignants'] != "")
                doublons_promo = df_complet.duplicated(subset=['Jours', 'Horaire', 'Promotion'], keep=False) & (df_complet['Promotion'] != "")

                for i, row in df_complet.iterrows():
                    if doublons_salle[i]:
                        conflits_list.append({
                            "Type de Conflit": "❌ SALLE OCCUPÉE",
                            "Promotion": row['Promotion'],
                            "Intervenant/Salle": row['Lieu'],
                            "Jour": row['Jours'],
                            "Horaire": row['Horaire'],
                            "Détails": f"La salle {row['Lieu']} est réservée par plusieurs groupes."
                        })
                    
                    if doublons_prof[i]:
                        conflits_list.append({
                            "Type de Conflit": "👤 CONFLIT ENSEIGNANT",
                            "Promotion": row['Promotion'],
                            "Intervenant/Salle": row['Enseignants'],
                            "Jour": row['Jours'],
                            "Horaire": row['Horaire'],
                            "Détails": f"L'enseignant {row['Enseignants']} a deux cours en même temps."
                        })

                    if doublons_promo[i]:
                        conflits_list.append({
                            "Type de Conflit": "⚠️ CONFLIT PROMOTION",
                            "Promotion": row['Promotion'],
                            "Intervenant/Salle": row['Promotion'],
                            "Jour": row['Jours'],
                            "Horaire": row['Horaire'],
                            "Détails": "Cette promotion a plusieurs enseignements affectés au même créneau."
                        })

                df_rapport = pd.DataFrame(conflits_list).drop_duplicates()

                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    df_complet[cols_format].to_excel(writer, sheet_name='Emploi du Temps', index=False)
                    
                    if not df_rapport.empty:
                        colonnes_rapport = ["Type de Conflit", "Promotion", "Intervenant/Salle", "Jour", "Horaire", "Détails"]
                        df_rapport[colonnes_rapport].to_excel(writer, sheet_name='Rapport Conflits', index=False)
                        worksheet = writer.sheets['Rapport Conflits']
                        for idx, col in enumerate(colonnes_rapport):
                            worksheet.set_column(idx, idx, 22)
                    else:
                        pd.DataFrame({"Résultat": ["Aucun conflit détecté"]}).to_excel(writer, sheet_name='Rapport Conflits', index=False)

                st.download_button(
                    label="📥 Télécharger le Rapport d'Erreurs Excel",
                    data=buffer.getvalue(),
                    file_name="Rapport_Conflits_EDT_2027.xlsx",
                    mime="application/vnd.ms-excel",
                    use_container_width=True
                ) 

    # ------------------------------------------------------------------
    # PORTAIL : SURVEILLANCES EXAMENS
    # ------------------------------------------------------------------
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

    # ------------------------------------------------------------------
    # PORTAIL : GÉNÉRATEUR AUTOMATIQUE
    # ------------------------------------------------------------------
    elif portail == "🤖 Générateur Automatique":
        if not is_admin:
            st.error("Accès réservé au Bureau des Examens.")
        else:
            st.header("⚙️ Moteur de Génération de Surveillances")
            if "effectifs_db" not in st.session_state:
                st.session_state.effectifs_db = {
                    "ING1": [50, 4], "MCIL1": [40, 3], "L1MCIL": [288, 4], 
                    "L2ELT": [90, 2], "M1RE": [15, 1], "ING2": [16, 1]
                }

            with st.expander("📦 Gestion des Effectifs", expanded=False):
                data_eff = [{"Promotion": k, "Effectif Total": v[0], "Nb de Salles": v[1]} for k, v in st.session_state.effectifs_db.items()]
                edited_eff = st.data_editor(pd.DataFrame(data_eff), use_container_width=True, num_rows="dynamic", hide_index=True)
                if st.button("💾 Sauvegarder la configuration"):
                    st.session_state.effectifs_db = {
                        row["Promotion"]: [int(row["Effectif Total"]), int(row["Nb de Salles"])] 
                        for _, row in edited_eff.iterrows()
                    }
                    st.success("Mis à jour !")

            SRC = "surveillances_2027.xlsx"
            if os.path.exists(SRC):
                df_src = pd.read_excel(SRC)
                df_src.columns = [str(c).strip() for c in df_src.columns]
                for c in df_src.columns: 
                    df_src[c] = df_src[c].fillna("").astype(str).str.strip()
                
                C_MAT = "Matière"
                C_SURV = "Surveillant(s)"
                C_DATE = "Date"
                C_HEURE = "Heure"
                C_SALLE = "Salle"
                C_PROMO = "Promotion"
                
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
                                            equipe.append(p)
                                            stats[p] += 1
                                            tracker.append({'D': row[C_DATE], 'H': row[C_HEURE], 'N': p})
                                
                                res_list.append({
                                    "Enseignements": row[C_MAT], 
                                    "Code": "S2-2027", 
                                    "Enseignants": " & ".join(equipe) if len(equipe) >= 2 else "⚠️ BESOIN RENFORT", 
                                    "Horaire": row[C_HEURE], 
                                    "Jours": row[C_DATE], 
                                    "Lieu": f"Salle {s_idx}" if nb_salles > 1 else row[C_SALLE], 
                                    "Promotion": f"{p_name} (S{s_idx})" if nb_salles > 1 else p_name
                                })
                    
                    st.session_state.df_genere = pd.DataFrame(res_list)
                    st.session_state.stats_charge = stats
                    st.rerun()

                if st.session_state.get("df_genere") is not None:
                    st.dataframe(st.session_state.df_genere, use_container_width=True, hide_index=True)
                    xlsx_buf = io.BytesIO()
                    with pd.ExcelWriter(xlsx_buf, engine='xlsxwriter') as writer: 
                        st.session_state.df_genere.to_excel(writer, index=False)
                    st.download_button("📥 TÉLÉCHARGER LE PLANNING", xlsx_buf.getvalue(), "EDT_Surveillances_2027.xlsx")

    # ------------------------------------------------------------------
    # PORTAIL : PORTAIL ENSEIGNANTS
    # ------------------------------------------------------------------
    elif portail == "👥 Portail Enseignants":
        if not is_admin:
            st.error("🚫 ACCÈS RESTREINT.")
            st.stop()

        # Chargement de secours
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
                st.error(f"❌ Le fichier {NOM_FICHIER_FIXE} est introuvable.")
                st.stop()

        col_l, col_t = st.columns([1, 5])
        with col_l:
            try:
                st.image("logo.PNG", width=80)
            except:
                st.markdown("🏛️")
        with col_t:
            st.header("🏢 Répertoire et Envoi Automatisé")
            st.write("Plateforme de gestion des emplois du temps 2026-2027-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA")

        # Récupération des données
        res_auth = supabase.table("enseignants_auth").select("nom_officiel, email, last_sent").execute()
        dict_auth = {
            str(row['nom_officiel']).strip().upper(): {
                "email": row['email'], 
                "statut": "✅ Envoyé" if row['last_sent'] else "⏳ En attente"
            } 
            for row in res_auth.data
        } if res_auth.data else {}

        noms_excel = sorted([e for e in df['Enseignants'].unique() if str(e) not in ["Non défini", "nan", ""]])
        
        donnees_finales = []
        for nom in noms_excel:
            nom_key = str(nom).strip().upper()
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

        # Boutons d'action
        c1, c2 = st.columns(2)
        with c1:
            if st.button("🔄 Réinitialiser les statuts (Comptes)", use_container_width=True):
                supabase.table("enseignants_auth").update({"last_sent": None}).neq("email", "").execute()
                st.success("✅ Statuts réinitialisés !")
                st.rerun()
        
        with c2:
            if st.button("🚀 Lancer l'envoi groupé", type="primary", use_container_width=True):
                try:
                    server = smtplib.SMTP('smtp.gmail.com', 587)
                    server.starttls()
                    
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
                            msg['From'] = f"{nom_affichage} <{expediteur_email}>"
                            msg['To'] = row["Email"]
                            
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
                                <p>Votre retour est indispensable pour la stabilisation des emplois du temps.</p>
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

                            buffer = io.BytesIO()
                            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                                df_mail.to_excel(writer, index=False, sheet_name='Mon EDT')
                                workbook = writer.book
                                worksheet = writer.sheets['Mon EDT']
                                fmt_header = workbook.add_format({
                                    'bold': True, 'bg_color': '#4472C4', 'font_color': 'white', 'border': 1
                                })
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
                            
                            try:
                                supabase.table("enseignants_auth").update(
                                    {"last_sent": datetime.now().isoformat()}
                                ).eq("email", row["Email"]).execute()
                            except:
                                pass
                    
                    server.quit()
                    st.success("✅ Envoi groupé terminé !")
                    st.rerun()
                except Exception as e:
                    st.error(f"Erreur : {e}")

        # Tableau récapitulatif
        st.divider()
        st.dataframe(pd.DataFrame(donnees_finales), use_container_width=True, hide_index=True)

        # Gestion des envois personnalisés
        st.divider()
        st.subheader("📬 Gestion des envois personnalisés")

        EMAIL_EXPEDITEUR = "chef.department.elt.fge@gmail.com"
        SECRET_APP = "gkzs pdza yodb icvd"

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

        def envoyer_emails(liste_destinataires, promotion_label="Individuel"):
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
                            <div style="background-color: white;">{table_html}</div>
                            <p>Cordialement.<br><b>Service d'enseignement</b></p>
                        </div>
                    </body>
                    </html>
                    """
                    msg.attach(MIMEText(corps_html, 'html'))

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
                colonnes_export = ["Enseignant", "Email"]
                df_download = df_export[colonnes_export].drop_duplicates()

                nb_ok = sum(1 for e in liste_promo if "@" in str(e["Email"]))
                st.metric("Emails opérationnels", f"{nb_ok} / {len(liste_promo)}")
                
                st.dataframe(df_export, use_container_width=True, hide_index=True)

                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    df_download.to_excel(writer, index=False, sheet_name='Liste_Emails')
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

                st.divider()

                if st.button(f"🚀 LANCER L'ENVOI POUR {choix_promo}", type="primary", use_container_width=True):
                    destinataires = [e for e in liste_promo if "@" in str(e["Email"])]
                    if destinataires:
                        envoyer_emails(destinataires, choix_promo)
                    else:
                        st.error("Aucun email valide pour cette promotion.")

        # Courrier officiel
        st.divider()
        with st.expander("✉️ ENVOYER UN COURRIER OFFICIEL (Direction / Secrétariat)", expanded=False):
            st.info("Mode Multi-Profils : Sélectionnez votre fonction. L'email officiel correspondant sera utilisé.")
            
            st.subheader("🔑 1. Identification de l'expéditeur")
            
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
                codes_secrets = {
                    "Chef de Département": "gkzs pdza yodb icvd", 
                    "Chef de départemet ELT": "kmtk zmkd kwpd cqzz",
                    "Chef de Département Adjoint": "",
                    "Secrétariat ELT": ""
                }
                code_auto = codes_secrets.get(role_choisi, "")
                expediteur_pass = st.text_input(
                    f"Mot de passe d'application ({role_choisi}) :", 
                    value=code_auto,
                    type="password", 
                    key=f"pass_{role_choisi}"
                )

            st.divider()
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

            if st.button("🚀 LANCER L'ENVOI OFFICIEL", type="primary", use_container_width=True):
                if not expediteur_pass:
                    st.error(f"❌ Veuillez saisir le mot de passe d'application pour {expediteur_mail}")
                elif not destinataires_mails:
                    st.error("❌ Aucun destinataire sélectionné.")
                elif not sujet_libre or not corps_libre:
                    st.error("❌ L'objet et le corps du message sont obligatoires.")
                else:
                    try:
                        if role_choisi == "Chef de Département":
                            signature = (
                                "\n\n---\nCordialement,\n\nPr. MILOUA Farid\n"
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

    # ------------------------------------------------------------------
    # PORTAIL : MISE À JOUR EDT
    # ------------------------------------------------------------------
    elif portail == "🎓 Portail mise à jour EDT":
        st.write(f"**MODE ACTIF :** {portail}")
        st.subheader("📚 Espace mise à jour EDT")
        st.info("Plateforme de gestion des EDTs-S2-2027-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA")

        st.markdown("### 📋 Consultation par Promotion")
        
        if df is not None and not df.empty:
            liste_promotions = sorted(df["Promotion"].unique().tolist())
        else:
            liste_promotions = ["ING1", "L3-ELT", "M1-RE", "M2-RE"]

        choix_promo = st.selectbox("Choisir votre Promotion :", liste_promotions)
        colonnes_ordonnees = ['Enseignements', 'Code', 'Enseignants', 'Horaire', 'Jours', 'Lieu', 'Promotion']
        df_vue = df[df["Promotion"] == choix_promo][colonnes_ordonnees].sort_values(by=["Jours", "Horaire"])
        
        st.write(f"**Emploi du temps actuel : {choix_promo}**")
        st.table(df_vue)

        c1, c2 = st.columns(2)
        with c1:
            output_vue = io.BytesIO()
            with pd.ExcelWriter(output_vue, engine='xlsxwriter') as writer:
                df_vue.to_excel(writer, index=False)
            st.download_button("📊 Télécharger Excel (Vue actuelle)", output_vue.getvalue(), f"EDT_{choix_promo}.xlsx")
        with c2:
            st.download_button("📄 Télécharger HTML (Vue actuelle)", df_vue.to_html(index=False), f"EDT_{choix_promo}.html", "text/html")

        if is_admin:
            st.write("---")
            st.subheader("✍️ Espace Éditeur de Données (Admin)")
            st.info("💡 Pour ajouter une charge : Filtrez pour isoler l'EDT concerné, puis cliquez sur le (+) en bas du tableau.")

            recherche = st.text_input("🔍 Rechercher une ligne (Enseignant, Salle ou Code) :", key="admin_search_bar")
            df_master = df[colonnes_ordonnees].copy()
            
            if recherche:
                masque = df_master.apply(lambda r: r.astype(str).str.contains(recherche, case=False).any(), axis=1)
                df_edition = df_master[masque].copy()
            else:
                df_edition = df_master.copy()

            total_lignes = len(df)
            st.caption(f"Lignes totales dans le fichier source : {total_lignes} | Prochain index : {total_lignes}")

            df_edite = st.data_editor(
                df_edition,
                use_container_width=True,
                num_rows="dynamic",
                key="admin_data_editor_main"
            )

            ca1, ca2 = st.columns(2)
            with ca1:
                out_ed = io.BytesIO()
                df_edite.to_excel(out_ed, index=False)
                st.download_button("📊 Télécharger l'EDT filtré (Excel)", out_ed.getvalue(), "EDT_Edition.xlsx")
            with ca2:
                st.download_button("📄 Télécharger l'EDT filtré (HTML)", df_edite.to_html(index=False), "EDT_Edition.html", "text/html")

            if st.button("💾 Sauvegarder les modifications et la nouvelle charge"):
                try:
                    if recherche:
                        df_final = pd.concat([df_master[~masque], df_edite], ignore_index=True)
                    else:
                        df_final = df_edite

                    df_final = df_final.dropna(subset=['Enseignements'])
                    df_final = df_final.sort_values(by=["Promotion", "Jours", "Horaire"])
                    df_final.to_excel(NOM_FICHIER_FIXE, index=False)
                    
                    st.success(f"✅ Enregistrement réussi ! Le fichier contient maintenant {len(df_final)} lignes.")
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ Erreur technique lors de la sauvegarde : {e}")

    # ------------------------------------------------------------------
    # PORTAIL : GESTION ADMINISTRATIVE - BORDEREAUX & PVs
    # ------------------------------------------------------------------
    elif portail == "📢 Gestion Administrative - Bordereaux & PVs":
        if not is_admin:
            st.error("🚫 ACCÈS RESTREINT AUX ADMINISTRATEURS.")
            st.stop()

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

        def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
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

        def generer_bordereau_iso(departement, donnees):
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
                r_ref_fixe = footer_p.add_run("Réf : UDL-GEL-ER-004-2027")
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
            
            nom_fichier_logo = "logo.PNG"
            if os.path.exists(nom_fichier_logo):
                p_logo.add_run().add_picture(nom_fichier_logo, width=Inches(0.833))
            else:
                r_alt = p_logo.add_run("[LOGO UNIVERSITÉ]")
                r_alt.font.name = 'Calibri'
                r_alt.font.size = Pt(8)
                r_alt.font.italic = True

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

            p_ref = doc.add_paragraph()
            p_ref.alignment = WD_ALIGN_PARAGRAPH.LEFT
            r_ref = p_ref.add_run(f"N° : {donnees['num_reference']}/ F.G.E/ V.D.E.Q.L.E/2027")
            r_ref.font.size = Pt(10)
            r_ref.font.name = 'Calibri'
            r_ref.bold = True

            doc.add_paragraph("\n")

            p_titre = doc.add_paragraph()
            p_titre.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_titre = p_titre.add_run("BORDEREAU D'ENVOI")
            r_titre.font.name = 'Calibri'
            r_titre.font.size = Pt(36)
            r_titre.italic = True
            r_titre.underline = True
            r_titre.bold = True
            
            doc.add_paragraph("\n")

            p_dest = doc.add_paragraph()
            p_dest.alignment = WD_ALIGN_PARAGRAPH.LEFT
            r_dest = p_dest.add_run(f"A monsieur : {donnees['destinataire']}")
            r_dest.bold = True
            r_dest.font.size = Pt(12)
            r_dest.font.name = 'Calibri'

            doc.add_paragraph("\n")

            liste_pieces = donnees['liste_pieces']
            nb_lignes_totatles = 2 + len(liste_pieces)
            
            table = doc.add_table(rows=nb_lignes_totatles, cols=3)
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
            doc = Document()
            p = doc.add_paragraph()
            run = p.add_run(f"{type_pv} - {departement}\nDocument en cours.")
            run.font.name = 'Calibri'
            return doc

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

else:
    st.error("❌ Aucune donnée disponible. Veuillez vérifier la connexion à Supabase ou la présence du fichier Excel local.")
