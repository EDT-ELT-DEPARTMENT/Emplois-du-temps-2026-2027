"""
================================================================================
Module EDT Intelligent — Plateforme ELT/UDL-SBA
Génération automatique : Enseignants & Promotions
Exports : HTML, Excel (openpyxl), PDF (fpdf) avec cellules flexibles
================================================================================
"""

import streamlit as st
import pandas as pd
import io
import math
import os
import re
from datetime import datetime
from pathlib import Path

try:
    from fpdf import FPDF
except ImportError:
    FPDF = None

try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Side, PatternFill, Font
except ImportError:
    Workbook = None

# =============================================================================
# PARTIE 1 — CONSTANTES & CONFIGURATION
# =============================================================================
HORAIRES_STD = [
    "08h00 - 09h30", "09h30 - 11h00", "11h00 - 12h30",
    "12h30 - 14h00", "14h00 - 15h30", "15h30 - 17h00"
]
JOURS_STD = ["Dimanche", "Lundi", "Mardi", "Mercredi", "Jeudi"]

COLORS = {
    "COURS": {"bg": "#dbeafe", "border": "#3b82f6", "text": "#1e40af", "emoji": "📘"},
    "TD":    {"bg": "#dcfce7", "border": "#22c55e", "text": "#166534", "emoji": "📗"},
    "TP":    {"bg": "#fee2e2", "border": "#ef4444", "text": "#991b1b", "emoji": "🔴"},
    "AUTRE": {"bg": "#f3f4f6", "border": "#9ca3af", "text": "#374151", "emoji": "⚪"}
}

# =============================================================================
# PARTIE 2 — FONCTIONS UTILITAIRES
# =============================================================================
def normalize_text(s):
    if not s or str(s).strip().lower() in ["non defini", "nan", "none", "", "non défini"]:
        return ""
    return re.sub(r'\s+', ' ', str(s).strip())

def normalize_horaire(h):
    if not h:
        return ""
    s = str(h).lower().replace(" ", "").replace("-", "").replace("–", "").replace(":", "").replace("h00", "h")
    mapping = {
        "08h00-09h30": "08h00 - 09h30", "08h-09h30": "08h00 - 09h30", "8h-9h30": "08h00 - 09h30",
        "09h30-11h00": "09h30 - 11h00", "9h30-11h": "09h30 - 11h00", "9h30-11h00": "09h30 - 11h00",
        "11h00-12h30": "11h00 - 12h30", "11h-12h30": "11h00 - 12h30",
        "12h30-14h00": "12h30 - 14h00", "12h30-14h": "12h30 - 14h00",
        "14h00-15h30": "14h00 - 15h30", "14h-15h30": "14h00 - 15h30",
        "15h30-17h00": "15h30 - 17h00", "15h30-17h": "15h30 - 17h00",
    }
    return mapping.get(s, str(h).strip())

def get_type_matiere(code):
    c = str(code).upper()
    if "COURS" in c: return "COURS"
    if "TD" in c:    return "TD"
    if "TP" in c:    return "TP"
    return "AUTRE"

# =============================================================================
# PARTIE 3 — FORMATAGE DES CELLULES
# =============================================================================
def format_cell_html(row, mode="enseignant"):
    typ = get_type_matiere(row.get("Code", ""))
    style = COLORS.get(typ, COLORS["AUTRE"])
    lines = [f"<b style=\"font-size:13px;\">{style['emoji']} {normalize_text(row.get('Enseignements', ''))}</b>"]
    if mode == "promotion":
        lines.append(f"<span style=\"font-size:11px;\">👤 {normalize_text(row.get('Enseignants', ''))}</span>")
    else:
        lines.append(f"<span style=\"font-size:11px;\">🎓 {normalize_text(row.get('Promotion', ''))}</span>")
    lines.append(f"<span style=\"font-size:10px;color:#64748b;\">📍 {normalize_text(row.get('Lieu', ''))}</span>")
    content = "<br>".join(lines)
    return f"""<div style="background:{style['bg']};border-left:4px solid {style['border']};border-radius:6px;padding:6px;margin:2px 0;color:{style['text']};font-family:'Segoe UI',Arial,sans-serif;line-height:1.4;">{content}</div>"""

def format_cell_text(row, mode="enseignant"):
    typ = get_type_matiere(row.get("Code", ""))
    emoji = COLORS.get(typ, COLORS["AUTRE"])["emoji"]
    lines = [f"{emoji} {normalize_text(row.get('Enseignements', ''))}"]
    if mode == "promotion":
        lines.append(f"Prof: {normalize_text(row.get('Enseignants', ''))}")
    else:
        lines.append(f"Promo: {normalize_text(row.get('Promotion', ''))}")
    lines.append(f"Salle: {normalize_text(row.get('Lieu', ''))}")
    return "\n".join(lines)

# =============================================================================
# PARTIE 4 — GÉNÉRATION DE LA GRILLE (Jours ↓, Horaires →)
# =============================================================================
def generer_grille(df_filtre, mode="enseignant"):
    if df_filtre is None or df_filtre.empty:
        return pd.DataFrame(), pd.DataFrame()
    df = df_filtre.copy()
    df["h_norm"] = df["Horaire"].apply(normalize_horaire)
    df["j_norm"] = df["Jours"].apply(lambda x: str(x).strip().lower() if x else "")
    horaires_map = {normalize_horaire(h): h for h in HORAIRES_STD}
    jours_map = {j.lower(): j for j in JOURS_STD}
    df = df[df["h_norm"].isin(horaires_map.keys())]
    df = df[df["j_norm"].isin(jours_map.keys())]
    if df.empty:
        return pd.DataFrame(), pd.DataFrame()
    def agg_html(rows):
        return "".join([format_cell_html(r, mode) for _, r in rows.iterrows()])
    def agg_text(rows):
        return "\n\n".join([format_cell_text(r, mode) for _, r in rows.iterrows()])
    grille_html = df.groupby(["j_norm", "h_norm"]).apply(agg_html, include_groups=False).unstack(fill_value="")
    grille_text = df.groupby(["j_norm", "h_norm"]).apply(agg_text, include_groups=False).unstack(fill_value="")
    jours_present = [j for j in [j.lower() for j in JOURS_STD] if j in grille_html.index]
    h_present = [h for h in [normalize_horaire(h) for h in HORAIRES_STD] if h in grille_html.columns]
    if not jours_present or not h_present:
        return pd.DataFrame(), pd.DataFrame()
    grille_html = grille_html.reindex(index=jours_present, columns=h_present).fillna("")
    grille_text = grille_text.reindex(index=jours_present, columns=h_present).fillna("")
    grille_html.index = [jours_map.get(i, i) for i in grille_html.index]
    grille_html.columns = [horaires_map.get(c, c) for c in grille_html.columns]
    grille_text.index = [jours_map.get(i, i) for i in grille_text.index]
    grille_text.columns = [horaires_map.get(c, c) for c in grille_text.columns]
    return grille_html, grille_text

# =============================================================================
# PARTIE 5 — EXPORT HTML
# =============================================================================
def export_html(grille_html, titre, sous_titre=""):
    if grille_html is None or grille_html.empty:
        return "<p>Aucune donnée</p>"
    css = """
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap');
        body { font-family: 'Inter', 'Segoe UI', Arial, sans-serif; background: #f8fafc; margin: 0; padding: 30px; color: #1e293b; }
        .container { max-width: 1400px; margin: auto; background: white; border-radius: 16px; box-shadow: 0 10px 40px rgba(0,0,0,0.08); overflow: hidden; }
        .header { background: linear-gradient(135deg, #1E3A8A 0%, #3B82F6 100%); color: white; padding: 30px; text-align: center; }
        .header h1 { margin: 0; font-size: 22px; }
        .header p { margin: 8px 0 0 0; opacity: 0.9; font-size: 14px; }
        .content { padding: 30px; }
        .meta { color: #64748b; font-size: 12px; margin-bottom: 20px; padding-bottom: 15px; border-bottom: 1px solid #f1f5f9; display: flex; justify-content: space-between; }
        table { width: 100%; border-collapse: separate; border-spacing: 0; table-layout: fixed; }
        th { background: #f1f5f9; color: #1e293b; padding: 14px 6px; font-size: 12px; font-weight: 700; border: 1px solid #e2e8f0; text-align: center; position: sticky; top: 0; }
        td { border: 1px solid #e2e8f0; padding: 8px; vertical-align: top; font-size: 12px; min-height: 70px; word-wrap: break-word; }
        tr:nth-child(even) td { background: #fafafa; }
        .time-col { background: #f8fafc !important; font-weight: 700; text-align: center; vertical-align: middle !important; width: 100px; font-size: 13px; }
        tr:hover td { background-color: #eff6ff !important; transition: background 0.2s; }
        .footer { text-align: center; padding: 20px; color: #94a3b8; font-size: 12px; border-top: 1px solid #f1f5f9; }
        @media print { body { background: white; padding: 0; } .container { box-shadow: none; border-radius: 0; } }
    </style>
    """
    thead = "<tr><th class='time-col'>JOUR / HORAIRE</th>" + "".join([f"<th>{h}</th>" for h in grille_html.columns]) + "</tr>"
    tbody = ""
    for jour, row in grille_html.iterrows():
        tbody += f"<tr><td class='time-col'>{jour}</td>"
        for val in row:
            tbody += f"<td>{val}</td>"
        tbody += "</tr>"
    return f"""<!DOCTYPE html>
<html lang="fr">
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>{titre}</title>{css}</head>
<body>
<div class="container">
    <div class="header">
        <h1>📅 {titre}</h1>
        <p>{sous_titre}</p>
        <span style="display:inline-block;background:#D4AF37;color:#1E3A8A;padding:4px 14px;border-radius:20px;font-size:11px;font-weight:700;margin-top:10px;">EDT Semestre 01 — 2026-2027</span>
    </div>
    <div class="content">
        <div class="meta">
            <span>Généré le {datetime.now().strftime('%d/%m/%Y à %H:%M')}</span>
            <span>{len(grille_html)} jours × {len(grille_html.columns)} créneaux</span>
        </div>
        <table><thead>{thead}</thead><tbody>{tbody}</tbody></table>
    </div>
    <div class="footer">département d'Électrotechnique — Faculté de Génie Électrique — UDL-SBA</div>
</div>
</body></html>"""

# =============================================================================
# PARTIE 6 — EXPORT EXCEL (HAUTEURS FLEXIBLES)
# =============================================================================
def export_excel(grille_text, titre):
    if Workbook is None:
        return None
    wb = Workbook()
    ws = wb.active
    ws.title = "EDT"
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(grille_text.columns)+1)
    ws['A1'] = titre
    ws['A1'].font = Font(bold=True, size=14, color="1E3A8A")
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 30
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(grille_text.columns)+1)
    ws['A2'] = f"Généré le {datetime.now().strftime('%d/%m/%Y à %H:%M')} — UDL-SBA"
    ws['A2'].font = Font(italic=True, size=10, color="64748B")
    ws['A2'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[2].height = 20
    header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    top_align = Alignment(horizontal='left', vertical='top', wrap_text=True)
    ws.cell(row=4, column=1, value="JOUR")
    ws.cell(row=4, column=1).fill = header_fill
    ws.cell(row=4, column=1).font = header_font
    ws.cell(row=4, column=1).alignment = center_align
    ws.cell(row=4, column=1).border = thin_border
    ws.column_dimensions['A'].width = 16
    for col_idx, horaire in enumerate(grille_text.columns, start=2):
        cell = ws.cell(row=4, column=col_idx, value=horaire)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align
        cell.border = thin_border
        ws.column_dimensions[cell.column_letter].width = 30
    for row_idx, (jour, row) in enumerate(grille_text.iterrows(), start=5):
        cell = ws.cell(row=row_idx, column=1, value=jour)
        cell.alignment = center_align
        cell.border = thin_border
        cell.font = Font(bold=True, size=11)
        max_lines = 1
        for col_idx, val in enumerate(row, start=2):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.alignment = top_align
            cell.border = thin_border
            cell.font = Font(size=10)
            if val:
                n_lines = str(val).count('\n') + max(1, len(str(val)) // 28)
                if n_lines > max_lines:
                    max_lines = n_lines
        ws.row_dimensions[row_idx].height = max(35, min(max_lines * 16 + 8, 250))
    ws.freeze_panes = 'B5'
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()

# =============================================================================
# PARTIE 7 — EXPORT PDF (CELLULES AUTO-AJUSTÉES)
# =============================================================================
def export_pdf(grille_text, titre, sous_titre=""):
    if FPDF is None:
        return None

    class EDT_PDF(FPDF):
        def header(self):
            self.set_font('Arial', 'B', 9)
            self.set_text_color(30, 58, 138)
            self.cell(0, 6, "Plateforme EDT — UDL-SBA | Semestre 01 2026-2027", 0, 1, 'C')
            self.set_draw_color(212, 175, 55)
            self.line(10, self.get_y(), self.w - 10, self.get_y())
            self.ln(3)
        def footer(self):
            self.set_y(-15)
            self.set_font('Arial', 'I', 8)
            self.set_text_color(128, 128, 128)
            self.cell(0, 10, f'Page {self.page_no()}', 0, 0, 'C')

    def sanitize(text):
        if not text: return ""
        t = str(text)
        repl = {"'":"'","'":"'","“":"\"","”":"\"","–":"-","—":"-","…":"...","«":"\"","»":"\"","œ":"oe","Œ":"OE",
                "à":"a","â":"a","ä":"a","á":"a","ã":"a","å":"a","è":"e","é":"e","ê":"e","ë":"e","ì":"i","í":"i","î":"i","ï":"i",
                "ò":"o","ó":"o","ô":"o","ö":"o","ù":"u","ú":"u","û":"u","ü":"u","ç":"c","ñ":"n","ÿ":"y","ý":"y",
                "À":"A","Â":"A","Ä":"A","Á":"A","Ã":"A","È":"E","É":"E","Ê":"E","Ë":"E","Ì":"I","Í":"I","Î":"I","Ï":"I",
                "Ò":"O","Ó":"O","Ô":"O","Ö":"O","Ù":"U","Ú":"U","Û":"U","Ü":"U","Ç":"C","Ñ":"N"}
        for old, new in repl.items():
            t = t.replace(old, new)
        return t.encode('latin-1', 'ignore').decode('latin-1')

    def calc_height(pdf, text, col_w, font_size=8, line_height=4.5):
        if not text:
            return 12
        pdf.set_font('Arial', '', font_size)
        lines = sanitize(text).split('\n')
        total_lines = 0
        for line in lines:
            if not line.strip():
                total_lines += 1
                continue
            line_w = pdf.get_string_width(line)
            n_wrap = max(1, math.ceil(line_w / (col_w - 4)))
            total_lines += n_wrap
        return max(12, total_lines * line_height + 4)

    pdf = EDT_PDF(orientation='L', unit='mm', format='A4')
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.set_font('Arial', 'B', 14)
    pdf.set_text_color(30, 58, 138)
    pdf.cell(0, 10, sanitize(titre), 0, 1, 'C')
    if sous_titre:
        pdf.set_font('Arial', 'I', 10)
        pdf.set_text_color(100, 100, 100)
        pdf.cell(0, 6, sanitize(sous_titre), 0, 1, 'C')
    pdf.ln(4)

    if grille_text is None or grille_text.empty:
        pdf.set_font('Arial', '', 11)
        pdf.cell(0, 10, "Aucune donnee a afficher.", 0, 1, 'C')
        return bytes(pdf.output())

    n_cols = len(grille_text.columns)
    page_w = pdf.w - 20
    col_time_w = 25
    col_data_w = (page_w - col_time_w) / n_cols if n_cols > 0 else page_w

    pdf.set_font('Arial', 'B', 8)
    pdf.set_fill_color(30, 58, 138)
    pdf.set_text_color(255, 255, 255)
    pdf.set_draw_color(200, 200, 200)
    h_header = 10
    pdf.cell(col_time_w, h_header, 'JOUR', 1, 0, 'C', True)
    for h in grille_text.columns:
        pdf.cell(col_data_w, h_header, sanitize(h), 1, 0, 'C', True)
    pdf.ln()

    pdf.set_text_color(0, 0, 0)
    pdf.set_font('Arial', '', 8)
    pdf.set_draw_color(180, 180, 180)

    for idx, (jour, row) in enumerate(grille_text.iterrows()):
        heights = [calc_height(pdf, val, col_data_w) for val in row]
        row_h = max(12, max(heights)) if heights else 12
        if pdf.get_y() + row_h > pdf.h - 20:
            pdf.add_page()
            pdf.set_font('Arial', 'B', 8)
            pdf.set_fill_color(30, 58, 138)
            pdf.set_text_color(255, 255, 255)
            pdf.cell(col_time_w, h_header, 'JOUR', 1, 0, 'C', True)
            for h in grille_text.columns:
                pdf.cell(col_data_w, h_header, sanitize(h), 1, 0, 'C', True)
            pdf.ln()
            pdf.set_text_color(0, 0, 0)
            pdf.set_font('Arial', '', 8)

        if idx % 2 == 0:
            pdf.set_fill_color(248, 250, 252)
        else:
            pdf.set_fill_color(255, 255, 255)

        x_start = pdf.get_x()
        y_start = pdf.get_y()
        pdf.rect(x_start, y_start, col_time_w, row_h, 'FD')
        pdf.set_xy(x_start, y_start + (row_h - 6) / 2)
        pdf.set_font('Arial', 'B', 8)
        pdf.cell(col_time_w, 6, sanitize(jour), 0, 0, 'C')
        pdf.set_xy(x_start + col_time_w, y_start)
        pdf.set_font('Arial', '', 8)

        for val in row:
            x = pdf.get_x()
            y = pdf.get_y()
            raw = str(val).upper()
            if "COURS" in raw:
                pdf.set_fill_color(219, 234, 254)
            elif "TD" in raw:
                pdf.set_fill_color(220, 252, 231)
            elif "TP" in raw:
                pdf.set_fill_color(254, 226, 226)
            else:
                if idx % 2 == 0:
                    pdf.set_fill_color(248, 250, 252)
                else:
                    pdf.set_fill_color(255, 255, 255)
            pdf.rect(x, y, col_data_w, row_h, 'FD')
            if val:
                pdf.set_xy(x + 2, y + 2)
                pdf.multi_cell(col_data_w - 4, 4.2, sanitize(val), 0, 'L')
            pdf.set_xy(x + col_data_w, y)
        pdf.ln(row_h)

    return bytes(pdf.output())

# =============================================================================
# PARTIE 8 — INTERFACE STREAMLIT
# =============================================================================
def run_edt_intelligent():
    st.title("🧠 EDT Intelligent")
    st.caption("Génération automatique des emplois du temps individuels et par promotion")

    base_dir = Path(__file__).parent.resolve()
    fichier_edt = str(base_dir / "dataEDT-ELT-S1-2027.xlsx")

    df_edt = pd.DataFrame()
    if os.path.exists(fichier_edt):
        try:
            df_edt = pd.read_excel(fichier_edt)
            df_edt.columns = df_edt.columns.str.strip()
            for col in ['Enseignements', 'Code', 'Enseignants', 'Horaire', 'Jours', 'Lieu', 'Promotion']:
                if col in df_edt.columns:
                    df_edt[col] = df_edt[col].fillna("Non défini").astype(str).str.strip()
        except Exception as e:
            st.error(f"Erreur chargement EDT : {e}")
    else:
        st.warning("Fichier EDT local non trouvé. Veuillez uploader le fichier.")
        uploaded = st.file_uploader("Uploader dataEDT-ELT-S1-2027.xlsx", type=["xlsx", "xls"])
        if uploaded:
            df_edt = pd.read_excel(uploaded)
            df_edt.columns = df_edt.columns.str.strip()

    if df_edt.empty:
        st.error("Aucune donnée EDT disponible.")
        return

    col1, col2 = st.columns([1, 3])
    with col1:
        st.markdown("### ⚙️ Paramètres")
        mode = st.radio("Mode :", ["👤 Enseignant", "🎓 Promotion"], key="edt_mode")
        if mode == "👤 Enseignant":
            liste = sorted([e for e in df_edt["Enseignants"].unique() if e and str(e).strip() not in ["", "nan", "None", "Non defini", "Non défini"]])
            selection = st.selectbox("Sélectionner :", liste, key="edt_ens")
            df_filtre = df_edt[df_edt["Enseignants"].str.contains(selection, case=False, na=False)]
            titre = f"EDT Individuel — {selection}"
            mode_export = "enseignant"
        else:
            liste = sorted([p for p in df_edt["Promotion"].unique() if p and str(p).strip() not in ["", "nan", "None", "Non defini", "Non défini"]])
            selection = st.selectbox("Sélectionner :", liste, key="edt_promo")
            df_filtre = df_edt[df_edt["Promotion"] == selection]
            titre = f"EDT Promotion — {selection}"
            mode_export = "promotion"
        st.divider()
        st.markdown("### 📥 Exports")
        if st.button("🔄 Générer / Rafraîchir", use_container_width=True, type="primary"):
            st.session_state['edt_grille_html'] = None
            st.session_state['edt_grille_text'] = None

    if 'edt_grille_html' not in st.session_state or st.session_state.get('edt_last_selection') != selection:
        grille_html, grille_text = generer_grille(df_filtre, mode_export)
        st.session_state['edt_grille_html'] = grille_html
        st.session_state['edt_grille_text'] = grille_text
        st.session_state['edt_last_selection'] = selection
        st.session_state['edt_titre'] = titre
        st.session_state['edt_mode_export'] = mode_export

    grille_html = st.session_state.get('edt_grille_html')
    grille_text = st.session_state.get('edt_grille_text')
    titre = st.session_state.get('edt_titre', 'EDT')

    with col2:
        if grille_html is not None and not grille_html.empty:
            st.markdown(f"### 📅 {titre}")
            st.write(grille_html.to_html(escape=False), unsafe_allow_html=True)
            c1, c2, c3 = st.columns(3)
            html_data = export_html(grille_html, titre, f"Mode : {mode_export.title()}")
            c1.download_button("🌐 HTML", html_data, f"EDT_{selection.replace(' ', '_')}.html", "text/html", use_container_width=True)
            xlsx_data = export_excel(grille_text, titre)
            if xlsx_data:
                c2.download_button("📊 Excel", xlsx_data, f"EDT_{selection.replace(' ', '_')}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            else:
                c2.button("📊 Excel", disabled=True, use_container_width=True)
            pdf_data = export_pdf(grille_text, titre, f"Mode : {mode_export.title()}")
            if pdf_data:
                c3.download_button("📄 PDF", pdf_data, f"EDT_{selection.replace(' ', '_')}.pdf", "application/pdf", use_container_width=True)
            else:
                c3.button("📄 PDF", disabled=True, use_container_width=True)
        else:
            st.info("Sélectionnez un enseignant ou une promotion pour générer l'EDT.")

if __name__ == "__main__":
    run_edt_intelligent()
