import streamlit as st
import pandas as pd
import os
import hashlib
import io
from datetime import datetime
from supabase import create_client

# ==========================================
# CONFIGURATION DE LA PAGE
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

# --- CONNEXION BASE DE DONNÉES ---
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

# --- STYLE CSS DÉTAILLÉ ---
st.markdown("""
<style>
.main-title { 
    color: #1E3A8A; 
    text-align: center; 
    font-family: 'serif'; 
    font-weight: bold; 
    border-bottom: 3px solid #D4AF37; 
    padding-bottom: 12px; 
    font-size: 18px; 
    margin-top: 5px;
}
.portal-badge { 
    background-color: #D4AF37; 
    color: #1E3A8A; 
    padding: 5px 12px; 
    border-radius: 5px; 
    font-weight: bold; 
    text-align: center; 
    margin-bottom: 12px; 
}
.date-badge { 
    background-color: #1E3A8A; 
    color: white; 
    padding: 5px 12px; 
    border-radius: 15px; 
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
    margin: 15px 0; 
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

/* === COULEURS EDT === */
.edt-cours { background-color: #1a5276 !important; color: white !important; }
.edt-td { background-color: #27ae60 !important; color: white !important; }
.edt-tp { background-color: #e67e22 !important; color: white !important; }
.edt-stage { background-color: #9b59b6 !important; color: white !important; }

/* === BOUTONS TÉLÉCHARGEMENT === */
.btn-download {
    padding: 10px 20px;
    border-radius: 10px;
    border: none;
    font-weight: 600;
    cursor: pointer;
    color: white;
    text-decoration: none;
    display: inline-flex;
    align-items: center;
    gap: 8px;
    transition: all 0.2s;
    margin: 5px;
}
.btn-download:hover { transform: scale(1.05); box-shadow: 0 4px 15px rgba(0,0,0,0.2); }
.btn-html { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); }
.btn-pdf { background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%); }
.btn-excel { background: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%); }
.btn-view { background: linear-gradient(135deg, #43e97b 0%, #38f9d7 100%); color: #1a1a2e; }

/* === TABLEAU EDT === */
.edt-table { width: 100%; border-collapse: collapse; table-layout: fixed; margin-top: 10px; background-color: white; }
.edt-table th { background-color: #0f3460 !important; color: white !important; border: 1px solid #000; padding: 8px; text-align: center; font-size: 11px; }
.edt-table td { border: 1px solid #000; padding: 6px !important; vertical-align: top; text-align: center; background-color: white; height: 80px; font-size: 10px; }
.edt-time-col { background-color: #f1f5f9 !important; font-weight: bold; }
.edt-sep { border-top: 3px solid #0f3460 !important; }
.edt-empty { color: #cbd5e1; font-style: italic; }

.separator { border-top: 1px dashed #bbb; margin: 4px 0; }
</style>
""", unsafe_allow_html=True)

# ==========================================
# FONCTIONS UTILITAIRES
# ==========================================
NOM_FICHIER_FIXE = "dataEDT-ELT-S1-2027.xlsx"
NOM_FICHIER_CONTACTS = "Permanents-Vacataires-ELT2-2026-2027.xlsx"

def normalize(s):
    if not s or s == "Non défini": 
        return "vide"
    s = str(s).strip().lower()
    s = s.replace(" ", "").replace("-", "").replace("–", "")
    s = s.replace(":00", "").replace("h00", "h")
    return s

def get_type_seance(code):
    code = str(code).upper()
    if code.startswith('COURS'):
        return 'Cours'
    elif code.startswith('TD'):
        return 'TD'
    elif code.startswith('TP'):
        return 'TP'
    elif code.startswith('STAGE'):
        return 'Stage'
    else:
        return 'Autre'

def get_color_class(type_seance):
    return {
        'Cours': 'edt-cours',
        'TD': 'edt-td', 
        'TP': 'edt-tp',
        'Stage': 'edt-stage',
        'Autre': ''
    }.get(type_seance, '')

def calc_hours(horaire):
    h = str(horaire).strip().replace(' ', '')
    if '1h30' in h or 'h-9h30' in h or 'h30-11h' in h or 'h-12h30' in h or 'h30-14h' in h or 'h-15h30' in h or 'h30-17h' in h:
        return 1.5
    elif '1h' in h and '1h30' not in h:
        return 1.0
    elif '2h' in h or 'h-10h' in h or 'h-16h' in h:
        return 2.0
    elif 'h-9h' in h:
        return 1.0
    elif 'h30-10h30' in h:
        return 1.0
    return 1.5

def map_to_slot(horaire):
    h = horaire.strip().replace(' ', '')
    if h in ['8h-9h30', '8h-9h', '8h-10h']:
        return '8h-9h30'
    elif h in ['9h30-11h', '9h30-10h30', '9h-10h']:
        return '9h30-11h'
    elif h in ['11h-12h30']:
        return '11h-12h30'
    elif h in ['12h30-14h', '13h-14h30']:
        return '12h30-14h'
    elif h in ['14h-15h30', '14h30-16h']:
        return '14h-15h30'
    elif h in ['15h30-17h', '14h-16h']:
        return '15h30-17h'
    return h

# ==========================================
# GÉNÉRATEURS DE FICHIERS (HTML, PDF, EXCEL)
# ==========================================

def generer_edt_html(promo, df_promo, time_slots, days):
    """Génère un fichier HTML complet et coloré pour une promotion"""

    total = len(df_promo)
    cours_count = len(df_promo[df_promo['Type'] == 'Cours'])
    td_count = len(df_promo[df_promo['Type'] == 'TD'])
    tp_count = len(df_promo[df_promo['Type'] == 'TP'])
    stage_count = len(df_promo[df_promo['Type'] == 'Stage'])
    total_hours = sum(df_promo['Horaire'].apply(calc_hours))

    # Récap par jour
    recap_rows = ""
    for day in days:
        day_data = df_promo[df_promo['Jours'] == day]
        c = len(day_data[day_data['Type'] == 'Cours'])
        t = len(day_data[day_data['Type'] == 'TD'])
        p = len(day_data[day_data['Type'] == 'TP'])
        s = len(day_data[day_data['Type'] == 'Stage'])
        h = sum(day_data['Horaire'].apply(calc_hours))
        recap_rows += f"<tr><td><strong>{day}</strong></td><td>{c}</td><td>{t}</td><td>{p}</td><td>{s}</td><td>{c+t+p+s}</td><td>{h:.1f}h</td></tr>"

    # Tableau EDT
    table_rows = ""
    for day in days:
        table_rows += f'<tr class="edt-sep"><td class="edt-time-col"><strong>{day}</strong></td>'
        for slot in time_slots:
            slot_data = df_promo[(df_promo['Jours'] == day) & (df_promo['Slot'] == slot)]
            if len(slot_data) == 0:
                table_rows += '<td class="edt-empty">—</td>'
            else:
                cells = ""
                for _, row in slot_data.iterrows():
                    color_class = get_color_class(row['Type'])
                    dur = "1h30"
                    h = row['Horaire']
                    if '1h' in h and '1h30' not in h:
                        dur = "1h00"
                    elif '2h' in h or 'h-10h' in h or 'h-16h' in h:
                        dur = "2h00"

                    cells += f'<div class="{color_class}" style="border-radius:6px;padding:4px;margin-bottom:3px;box-shadow:0 2px 4px rgba(0,0,0,0.1);">'
                    cells += f'<div style="font-weight:600;font-size:0.85em;">{row["Enseignements"]}</div>'
                    cells += f'<div style="font-size:0.8em;opacity:0.9;">{row["Type"]} — {dur}</div>'
                    cells += f'<div style="display:inline-block;background:rgba(255,255,255,0.25);padding:1px 5px;border-radius:8px;font-size:0.75em;margin-top:2px;">{row["Groupe"]} | {row["Lieu"]} | {row["Enseignants"]}</div>'
                    cells += '</div>'
                table_rows += f'<td>{cells}</td>'
        table_rows += '</tr>'

    html = f"""<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8">
<title>Emploi du Temps — {promo} S2 2027</title>
<style>
* {{ margin:0; padding:0; box-sizing:border-box; }}
body {{ font-family:'Segoe UI',system-ui,sans-serif; background:#f0f2f5; color:#1a1a2e; padding:16px; }}
.container {{ max-width:1600px; margin:0 auto; }}
.header {{ background:linear-gradient(135deg,#1a1a2e 0%,#16213e 50%,#0f3460 100%); color:white; padding:28px; border-radius:14px; text-align:center; margin-bottom:20px; box-shadow:0 8px 32px rgba(0,0,0,0.15); }}
.header h1 {{ font-size:2em; margin-bottom:8px; }}
.header p {{ font-size:1.1em; opacity:0.9; }}
.stats {{ display:flex; justify-content:center; gap:20px; margin-bottom:20px; flex-wrap:wrap; }}
.stat {{ background:white; padding:14px 24px; border-radius:12px; box-shadow:0 2px 10px rgba(0,0,0,0.06); text-align:center; }}
.stat .v {{ font-size:1.8em; font-weight:bold; color:#0f3460; }}
.stat .l {{ font-size:0.8em; color:#6c757d; text-transform:uppercase; }}
.legend {{ display:flex; justify-content:center; gap:20px; margin-bottom:20px; flex-wrap:wrap; }}
.lg {{ display:flex; align-items:center; gap:8px; background:white; padding:8px 16px; border-radius:8px; box-shadow:0 1px 4px rgba(0,0,0,0.06); }}
.lg-d {{ width:18px; height:18px; border-radius:4px; }}
.tw {{ background:white; border-radius:14px; overflow:hidden; box-shadow:0 4px 20px rgba(0,0,0,0.08); overflow-x:auto; }}
table {{ width:100%; border-collapse:collapse; font-size:0.78em; }}
th {{ background:#0f3460; color:white; padding:12px 10px; font-weight:600; text-align:center; position:sticky; top:0; white-space:nowrap; }}
td {{ border:1px solid #e2e8f0; padding:6px; vertical-align:top; min-width:180px; height:70px; }}
td.tcol {{ background:#f1f5f9; font-weight:700; text-align:center; color:#334155; width:90px; min-width:90px; font-size:0.95em; }}
.s {{ border-radius:8px; padding:5px 8px; color:white; font-size:0.82em; line-height:1.3; box-shadow:0 2px 6px rgba(0,0,0,0.12); display:block; margin-bottom:4px; }}
.s .m {{ font-weight:600; font-size:0.95em; }}
.s .t {{ font-size:0.8em; opacity:0.9; }}
.s .g {{ display:inline-block; background:rgba(255,255,255,0.25); padding:1px 6px; border-radius:10px; font-size:0.75em; margin-top:2px; }}
.empty {{ color:#cbd5e1; font-style:italic; text-align:center; padding-top:22px; }}
tr.sep td {{ border-top:3px solid #0f3460; }}
.recap-section {{ background:white; border-radius:14px; padding:24px; margin-bottom:20px; box-shadow:0 2px 10px rgba(0,0,0,0.06); }}
.recap-section h2 {{ color:#0f3460; margin-bottom:16px; font-size:1.4em; }}
.recap-table {{ width:100%; border-collapse:collapse; }}
.recap-table th {{ background:#1a1a2e; padding:10px; font-size:0.85em; }}
.recap-table td {{ height:auto; padding:10px; border-bottom:1px solid #e9ecef; }}
.recap-table tr:nth-child(even) {{ background:#f8f9fa; }}
.download-btns {{ display:flex; justify-content:center; gap:15px; margin-bottom:20px; flex-wrap:wrap; }}
.btn {{ padding:12px 24px; border-radius:10px; border:none; font-size:1em; font-weight:600; cursor:pointer; color:white; text-decoration:none; display:inline-block; transition:transform 0.2s, box-shadow 0.2s; }}
.btn:hover {{ transform:translateY(-2px); box-shadow:0 4px 12px rgba(0,0,0,0.15); }}
.btn-html {{ background:linear-gradient(135deg,#667eea 0%,#764ba2 100%); }}
.btn-pdf {{ background:linear-gradient(135deg,#f093fb 0%,#f5576c 100%); }}
.btn-excel {{ background:linear-gradient(135deg,#4facfe 0%,#00f2fe 100%); }}
@media print {{ body {{ background:white; padding:10px; }} .header {{ border-radius:0; box-shadow:none; }} .tw {{ box-shadow:none; border:1px solid #ccc; }} .download-btns {{ display:none; }} }}
</style>
</head>
<body>
<div class="container">
<div class="header">
<h1>📅 Emploi du Temps — {promo} S2 2027</h1>
<p>Semestre 2 | Dimanche–Jeudi | Électrotechnique</p>
</div>

<div class="download-btns">
<a href="EDT_{promo}_S2_2027.html" class="btn btn-html" download>📄 Télécharger HTML</a>
<a href="EDT_{promo}_S2_2027.pdf" class="btn btn-pdf" download>📑 Télécharger PDF</a>
<a href="EDT_{promo}_S2_2027.xlsx" class="btn btn-excel" download>📊 Télécharger Excel</a>
</div>

<div class="stats">
<div class="stat"><div class="v">{total}</div><div class="l">Séances / sem.</div></div>
<div class="stat"><div class="v">{cours_count}</div><div class="l">Cours</div></div>
<div class="stat"><div class="v">{td_count}</div><div class="l">TD</div></div>
<div class="stat"><div class="v">{tp_count}</div><div class="l">TP</div></div>
<div class="stat"><div class="v">{stage_count}</div><div class="l">Stage</div></div>
<div class="stat"><div class="v">{total_hours:.1f}h</div><div class="l">Total / sem.</div></div>
</div>

<div class="legend">
<div class="lg"><div class="lg-d" style="background:#1a5276"></div><span><strong>Cours</strong> — Commun</span></div>
<div class="lg"><div class="lg-d" style="background:#27ae60"></div><span><strong>TD</strong> — Travaux Dirigés</span></div>
<div class="lg"><div class="lg-d" style="background:#e67e22"></div><span><strong>TP</strong> — Travaux Pratiques</span></div>
<div class="lg"><div class="lg-d" style="background:#9b59b6"></div><span><strong>Stage</strong> — Stage en entreprise</span></div>
</div>

<div class="recap-section">
<h2>📊 Récapitulatif par Jour</h2>
<table class="recap-table">
<thead><tr><th>Jour</th><th>Cours</th><th>TD</th><th>TP</th><th>Stage</th><th>Total Séances</th><th>Heures</th></tr></thead>
<tbody>
{recap_rows}
</tbody>
</table>
</div>

<div class="recap-section">
<h2>📅 Emploi du Temps Détaillé</h2>
<div class="tw">
<table>
<thead><tr><th>Jour / Heure</th>
<th>08h00–09h30</th>
<th>09h30–11h00</th>
<th>11h00–12h30</th>
<th>12h30–14h00</th>
<th>14h30–16h00</th>
<th>16h00–17h30</th>
</tr></thead>
<tbody>
{table_rows}
</tbody>
</table>
</div>
</div>

</div>
</body>
</html>"""
    return html


def generer_edt_excel(promo, df_promo, time_slots, days):
    """Génère un fichier Excel coloré pour une promotion"""
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    colors_excel = {
        'Cours': '1a5276', 'TD': '27ae60', 'TP': 'e67e22', 
        'Stage': '9b59b6', 'Autre': '7f8c8d'
    }
    light_colors = {
        'Cours': 'd4e6f1', 'TD': 'd5f5e3', 'TP': 'fdebd0', 
        'Stage': 'e8daef', 'Autre': 'd5dbdb'
    }

    header_fill = PatternFill(start_color='0f3460', end_color='0f3460', fill_type='solid')
    header_font = Font(color='FFFFFF', bold=True, size=11)
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    thin_border = Border(
        left=Side(style='thin', color='e2e8f0'),
        right=Side(style='thin', color='e2e8f0'),
        top=Side(style='thin', color='e2e8f0'),
        bottom=Side(style='thin', color='e2e8f0')
    )

    wb = Workbook()
    ws = wb.active
    ws.title = f"EDT {promo}"

    # Titre
    ws.merge_cells('A1:G1')
    ws['A1'] = f'Emploi du Temps — {promo} S2 2027'
    ws['A1'].font = Font(size=16, bold=True, color='0f3460')
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 30

    # En-têtes
    headers = ['Jour', '08h-9h30', '9h30-11h', '11h-12h30', '12h30-14h', '14h-15h30', '15h30-17h']
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align
        cell.border = thin_border

    ws.column_dimensions['A'].width = 12
    for col in range(2, 8):
        ws.column_dimensions[get_column_letter(col)].width = 35

    row_idx = 4
    for day in days:
        ws.cell(row=row_idx, column=1, value=day).font = Font(bold=True, size=10)
        ws.cell(row=row_idx, column=1).alignment = center_align
        ws.cell(row=row_idx, column=1).fill = PatternFill(start_color='f1f5f9', end_color='f1f5f9', fill_type='solid')
        ws.cell(row=row_idx, column=1).border = thin_border

        for col_idx, slot in enumerate(time_slots, 2):
            slot_data = df_promo[(df_promo['Jours'] == day) & (df_promo['Slot'] == slot)]
            if len(slot_data) > 0:
                cell = ws.cell(row=row_idx, column=col_idx)
                texts = []
                for _, row in slot_data.iterrows():
                    texts.append(f"{row['Enseignements']}\n({row['Type']}) | {row['Groupe']}\n{row['Lieu']} | {row['Enseignants']}")
                cell.value = '\n---\n'.join(texts)
                cell.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
                cell.border = thin_border
                dominant_type = slot_data.iloc[0]['Type']
                cell.fill = PatternFill(start_color=light_colors.get(dominant_type, 'd5dbdb'), 
                                       end_color=light_colors.get(dominant_type, 'd5dbdb'), fill_type='solid')
            else:
                cell = ws.cell(row=row_idx, column=col_idx, value='—')
                cell.alignment = center_align
                cell.border = thin_border
                cell.font = Font(color='cbd5e1', italic=True)

        ws.row_dimensions[row_idx].height = 80
        row_idx += 1

    # Feuille récap
    ws2 = wb.create_sheet("Récapitulatif")
    ws2['A1'] = f'Récapitulatif — {promo} S2 2027'
    ws2['A1'].font = Font(size=14, bold=True, color='0f3460')

    recap_headers = ['Jour', 'Cours', 'TD', 'TP', 'Stage', 'Total', 'Heures']
    for col, h in enumerate(recap_headers, 1):
        cell = ws2.cell(row=3, column=col, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align
        cell.border = thin_border

    row_idx = 4
    for day in days:
        day_data = df_promo[df_promo['Jours'] == day]
        c = len(day_data[day_data['Type'] == 'Cours'])
        t = len(day_data[day_data['Type'] == 'TD'])
        p = len(day_data[day_data['Type'] == 'TP'])
        s = len(day_data[day_data['Type'] == 'Stage'])
        h = sum(day_data['Horaire'].apply(calc_hours))

        ws2.cell(row=row_idx, column=1, value=day).font = Font(bold=True)
        ws2.cell(row=row_idx, column=2, value=c)
        ws2.cell(row=row_idx, column=3, value=t)
        ws2.cell(row=row_idx, column=4, value=p)
        ws2.cell(row=row_idx, column=5, value=s)
        ws2.cell(row=row_idx, column=6, value=c+t+p+s)
        ws2.cell(row=row_idx, column=7, value=f"{h:.1f}h")

        for col in range(1, 8):
            ws2.cell(row=row_idx, column=col).border = thin_border
            ws2.cell(row=row_idx, column=col).alignment = center_align

        if row_idx % 2 == 0:
            for col in range(1, 8):
                ws2.cell(row=row_idx, column=col).fill = PatternFill(start_color='f8f9fa', end_color='f8f9fa', fill_type='solid')

        row_idx += 1

    for col in range(1, 8):
        ws2.column_dimensions[get_column_letter(col)].width = 15

    # Sauvegarder en mémoire
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def generer_edt_pdf(promo, df_promo, time_slots, days):
    """Génère un fichier PDF pour une promotion"""
    try:
        from reportlab.lib import colors as rl_colors
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.units import cm
        import re

        buffer = io.BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=landscape(A4), 
                               rightMargin=1*cm, leftMargin=1*cm, 
                               topMargin=1*cm, bottomMargin=1*cm)

        styles = getSampleStyleSheet()
        title_style = ParagraphStyle(
            'CustomTitle', parent=styles['Heading1'],
            fontSize=18, textColor=rl_colors.HexColor('#0f3460'),
            spaceAfter=20, alignment=1
        )

        elements = []
        elements.append(Paragraph(f'Emploi du Temps — {promo} S2 2027', title_style))
        elements.append(Spacer(1, 0.3*cm))

        data = [['Jour / Heure', '8h-9h30', '9h30-11h', '11h-12h30', '12h30-14h', '14h-15h30', '15h30-17h']]

        color_map = {
            'Cours': rl_colors.HexColor('#1a5276'),
            'TD': rl_colors.HexColor('#27ae60'),
            'TP': rl_colors.HexColor('#e67e22'),
            'Stage': rl_colors.HexColor('#9b59b6'),
            'Autre': rl_colors.HexColor('#7f8c8d')
        }

        for day in days:
            row = [Paragraph(f'<b>{day}</b>', styles['Normal'])]
            for slot in time_slots:
                slot_data = df_promo[(df_promo['Jours'] == day) & (df_promo['Slot'] == slot)]
                if len(slot_data) > 0:
                    texts = []
                    for _, r in slot_data.iterrows():
                        t_type = r['Type']
                        color = color_map.get(t_type, rl_colors.grey)
                        texts.append(f"<font color='#{color.hexval()[2:]}'><b>{r['Enseignements']}</b></font><br/>({t_type}) {r['Groupe']}<br/><font size=8>{r['Lieu']} | {r['Enseignants']}</font>")
                    row.append(Paragraph('<br/>'.join(texts), styles['Normal']))
                else:
                    row.append(Paragraph('—', styles['Normal']))
            data.append(row)

        t = Table(data, colWidths=[2.5*cm] + [6*cm]*6)
        t.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), rl_colors.HexColor('#0f3460')),
            ('TEXTCOLOR', (0, 0), (-1, 0), rl_colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (0, -1), rl_colors.HexColor('#f1f5f9')),
            ('GRID', (0, 0), (-1, -1), 0.5, rl_colors.HexColor('#e2e8f0')),
            ('LEFTPADDING', (0, 0), (-1, -1), 4),
            ('RIGHTPADDING', (0, 0), (-1, -1), 4),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
            ('FONTSIZE', (0, 1), (-1, -1), 7),
        ]))

        for i in range(1, len(data)):
            if i % 2 == 0:
                t.setStyle(TableStyle([('BACKGROUND', (1, i), (-1, i), rl_colors.HexColor('#f8f9fa'))]))

        elements.append(t)
        elements.append(Spacer(1, 0.5*cm))

        legend_data = [['Légende:', 'Cours', 'TD', 'TP', 'Stage'], ['', '■', '■', '■', '■']]
        legend = Table(legend_data, colWidths=[3*cm, 3*cm, 3*cm, 3*cm, 3*cm])
        legend.setStyle(TableStyle([
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('TEXTCOLOR', (1, 1), (1, 1), rl_colors.HexColor('#1a5276')),
            ('TEXTCOLOR', (2, 1), (2, 1), rl_colors.HexColor('#27ae60')),
            ('TEXTCOLOR', (3, 1), (3, 1), rl_colors.HexColor('#e67e22')),
            ('TEXTCOLOR', (4, 1), (4, 1), rl_colors.HexColor('#9b59b6')),
            ('FONTSIZE', (0, 1), (-1, 1), 20),
        ]))
        elements.append(legend)

        doc.build(elements)
        return buffer.getvalue()
    except Exception as e:
        st.error(f"Erreur PDF: {e}")
        return None


# ==========================================
# CHARGEMENT DES DONNÉES
# ==========================================
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

# Chargement depuis fichier Excel local
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
    df = charger_donnees_supabase()

# Ajout des colonnes Type et Slot
df['Type'] = df['Code'].apply(get_type_seance)
df['Slot'] = df['Horaire'].apply(map_to_slot)

# Extraction des groupes
import re
def extract_group(enseignement):
    text = str(enseignement)
    match = re.search(r'\\(G(\\d+)\\)', text)
    if match: return f"G{match.group(1)}"
    match = re.search(r'\\(SG(\\d+)\\)', text)
    if match: return f"SG{match.group(1)}"
    match = re.search(r'SG(\\d+)', text)
    if match: return f"SG{match.group(1)}"
    return 'TOUS'

df['Groupe'] = df['Enseignements'].apply(extract_group)

# Chargement des contacts
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
# SYSTÈME D'AUTH
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
        noms_possibles = sorted(df["Enseignants"].unique()) if df is not None else []

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
                        "nom_officiel": new_nom, "email": new_email,
                        "password_hash": hash_pw(new_pass), "role": "enseignant",
                        "statut": statut_user, "telephone": new_phone if statut_user == "Vacataire" else None
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
    st.stop()

# ==========================================
# GARDIEN DE SESSION
# ==========================================
user = st.session_state.get("user_data")
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
        "📖 Emploi du Temps", "📅 Surveillances Examens", 
        "🤖 Générateur Automatique", "👥 Portail Enseignants", 
        "🎓 Portail mise à jour EDT", "📢 Gestion Administrative - Bordereaux & PVs"
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

st.markdown(f"<div class='portal-badge'>MODE ACTIF : {portail.upper()}</div>", unsafe_allow_html=True)

# ==========================================
# CONSTANTES EDT
# ==========================================
horaires_list = [
    "8h - 9h", "8h - 9h30", "8h - 10h", "9h - 10h", "9h30 - 11h", 
    "10h - 11h", "11h - 12h", "11h - 12h30", "12h - 13h", 
    "12h30 - 14h", "13h - 14h", "14h - 15h30", "14h - 16h", "15h30 - 17h"
]
jours_list = ["Dimanche", "Lundi", "Mardi", "Mercredi", "Jeudi"]
map_h = {normalize(h): h for h in horaires_list}
map_j = {normalize(j): j for j in jours_list}

# Time slots standard pour les tableaux EDT
time_slots = ['8h-9h30', '9h30-11h', '11h-12h30', '12h30-14h', '14h-15h30', '15h30-17h']
days = ['Dimanche', 'Lundi', 'Mardi', 'Mercredi', 'Jeudi']

# Promotions ELT
elt_promos = ['L2ELT', 'L3ELT', 'ING3EI', 'ING3RSE', 'ING4', 'M1CE', 'M1ER', 'M1RE']
promo_names = {
    'L2ELT': 'Licence 2 Électrotechnique',
    'L3ELT': 'Licence 3 Électrotechnique',
    'ING3EI': 'ING3 Électrotechnique Industrielle',
    'ING3RSE': 'ING3 Réseaux & Systèmes Électriques',
    'ING4': 'ING4 Électrotechnique',
    'M1CE': 'Master 1 Commande Électrique',
    'M1ER': 'Master 1 Énergies Renouvelables',
    'M1RE': 'Master 1 Réseaux Électriques'
}

# ==========================================
# ESPACE ÉDITEUR AVANCÉ (ADMIN)
# ==========================================
if is_admin and mode_view == "✍️ Éditeur de données":
    # ... [garder le code existant de l'éditeur] ...
    pass

# ==========================================
# ESPACE PROMOTION (ADMIN) - AVEC BOUTONS TÉLÉCHARGEMENT
# ==========================================
if is_admin and mode_view == "Promotion":
    st.subheader("📊 Emplois du Temps par Promotion - Téléchargements")

    # Créer les onglets pour chaque promotion ELT
    tabs = st.tabs([f"🎓 {p}" for p in elt_promos])

    for idx, promo in enumerate(elt_promos):
        with tabs[idx]:
            promo_data = df[df['Promotion'] == promo].copy()

            if len(promo_data) == 0:
                st.warning(f"Aucune donnée pour {promo}")
                continue

            # Stats
            total = len(promo_data)
            cours_count = len(promo_data[promo_data['Type'] == 'Cours'])
            td_count = len(promo_data[promo_data['Type'] == 'TD'])
            tp_count = len(promo_data[promo_data['Type'] == 'TP'])
            stage_count = len(promo_data[promo_data['Type'] == 'Stage'])
            total_hours = sum(promo_data['Horaire'].apply(calc_hours))

            # Affichage des stats
            col1, col2, col3, col4, col5, col6 = st.columns(6)
            col1.metric("Séances", total)
            col2.metric("Cours", cours_count)
            col3.metric("TD", td_count)
            col4.metric("TP", tp_count)
            col5.metric("Stage", stage_count)
            col6.metric("Heures", f"{total_hours:.1f}h")

            st.divider()

            # === BOUTONS DE TÉLÉCHARGEMENT ===
            st.markdown("### 📥 Téléchargements")

            col_dl1, col_dl2, col_dl3 = st.columns(3)

            # 1. HTML
            with col_dl1:
                html_content = generer_edt_html(promo, promo_data, time_slots, days)
                st.download_button(
                    label="📄 Télécharger HTML",
                    data=html_content,
                    file_name=f"EDT_{promo}_S2_2027.html",
                    mime="text/html",
                    use_container_width=True,
                    key=f"dl_html_{promo}"
                )

            # 2. Excel
            with col_dl2:
                excel_content = generer_edt_excel(promo, promo_data, time_slots, days)
                st.download_button(
                    label="📊 Télécharger Excel",
                    data=excel_content,
                    file_name=f"EDT_{promo}_S2_2027.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key=f"dl_excel_{promo}"
                )

            # 3. PDF
            with col_dl3:
                pdf_content = generer_edt_pdf(promo, promo_data, time_slots, days)
                if pdf_content:
                    st.download_button(
                        label="📑 Télécharger PDF",
                        data=pdf_content,
                        file_name=f"EDT_{promo}_S2_2027.pdf",
                        mime="application/pdf",
                        use_container_width=True,
                        key=f"dl_pdf_{promo}"
                    )

            st.divider()

            # === AFFICHAGE DU TABLEAU EDT DANS STREAMLIT ===
            st.markdown("### 📅 Aperçu de l'Emploi du Temps")

            # Construire le tableau HTML pour l'affichage
            def format_edt_cell(rows):
                items = []
                for _, r in rows.iterrows():
                    color_class = get_color_class(r['Type'])
                    items.append(f'<div class="{color_class}" style="border-radius:4px;padding:3px;margin:2px;font-size:9px;">'
                                f'<b>{r["Enseignements"]}</b><br/>'
                                f'{r["Type"]} | {r["Groupe"]}<br/>'
                                f'<small>{r["Lieu"]} | {r["Enseignants"]}</small>'
                                f'</div>')
                return "".join(items)

            grid = promo_data.groupby(['h_norm', 'j_norm']).apply(format_edt_cell, include_groups=False).unstack('j_norm')
            grid = grid.reindex(index=[normalize(h) for h in horaires_list], columns=[normalize(j) for j in jours_list]).fillna("")
            grid.index = [map_h.get(i, i) for i in grid.index]
            grid.columns = [map_j.get(c, c) for c in grid.columns]

            st.write(grid.to_html(escape=False, classes='edt-table'), unsafe_allow_html=True)

            # Légende des couleurs
            st.markdown("""
            <div style="display:flex; justify-content:center; gap:20px; margin-top:10px;">
            <div style="display:flex; align-items:center; gap:5px;"><div style="width:15px;height:15px;background:#1a5276;border-radius:3px;"></div><span>Cours</span></div>
            <div style="display:flex; align-items:center; gap:5px;"><div style="width:15px;height:15px;background:#27ae60;border-radius:3px;"></div><span>TD</span></div>
            <div style="display:flex; align-items:center; gap:5px;"><div style="width:15px;height:15px;background:#e67e22;border-radius:3px;"></div><span>TP</span></div>
            <div style="display:flex; align-items:center; gap:5px;"><div style="width:15px;height:15px;background:#9b59b6;border-radius:3px;"></div><span>Stage</span></div>
            </div>
            """, unsafe_allow_html=True)

# ==========================================
# ESPACE ENSEIGNANT (PERSONNEL)
# ==========================================
elif mode_view == "Personnel" or (is_admin and mode_view == "Enseignant"):
    # ... [garder le code existant de l'affichage individuel avec les boutons de téléchargement] ...
    pass

# ==========================================
# AUTRES ESPACES (garder le code existant)
# ==========================================
# ... [Surveillances, Générateur, Portail Enseignants, etc.] ...
