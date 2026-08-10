        # =============================================================================
        # ONGLET 5 : MON EDT (ÉTUDIANT CONNECTÉ UNIQUEMENT)
        # =============================================================================
with tab5:
    st.header("📅 Mon Emploi du Temps")
    st.caption("Consultation et téléchargement de votre EDT hebdomadaire filtré par groupe & sous-groupe")

    if not étudiant_connecte:
        st.info("ℹ️ Cet onglet est réservé aux étudiants connectés. Retournez à l'onglet 📝 Suivi d'Assiduité pour vous authentifier.")
    else:
        promo_etu = str(étudiant_connecte.get("promotion", "")).strip()
        nom_etu   = str(étudiant_connecte.get("nom", "Étudiant")).strip()

        if df_edt.empty:
            st.error("❌ Les données EDT ne sont pas disponibles.")
        else:
            # ─── FONCTION DE NORMALISATION LOCALE (autonome) ───
            def _norm(x):
                if not x or str(x).strip().lower() in ["non defini", "nan", "none", "", "vide"]:
                    return "vide"
                s = str(x).strip().lower().replace(" ", "").replace("-", "").replace("–", "").replace(":", "")
                s = s.replace("h00", "h").replace("h0", "h")
                return s

            # ─── MAPPING PROMOTION ───
            df_edt["Promotion_Mappee"] = df_edt["Promotion"].apply(mapper_promotion)
            promo_mapped = mapper_promotion(promo_etu)

            df_edt_etu = df_edt[df_edt["Promotion_Mappee"] == promo_mapped].copy()

            # Fallback si le mapping échoue
            if df_edt_etu.empty:
                df_edt_etu = df_edt[
                    df_edt["Promotion"].astype(str).str.strip().str.upper() == promo_etu.upper()
                ].copy()

            # ═══════════════════════════════════════════════════════
            # RÉCUPÉRATION GROUPE / SOUS-GROUPE ÉTUDIANT
            # ═══════════════════════════════════════════════════════
            groupe_etu = ""
            sous_groupe_etu = ""
            try:
                cols_map = detecter_colonnes_etudiant(df_etu)
                row_etu = df_etu[
                    df_etu["Nom_Complet"].astype(str).str.strip().str.upper() == nom_etu.upper()
                ]
                if not row_etu.empty:
                    groupe_etu = str(row_etu.iloc[0].get(cols_map.get('groupe', ''), '')).strip().upper()
                    sous_groupe_etu = str(row_etu.iloc[0].get(cols_map.get('sous_groupe', ''), '')).strip().upper()
            except Exception:
                pass

            # ═══════════════════════════════════════════════════════
            # DÉTECTION INTELLIGENTE DANS TOUTES LES COLONNES EDT
            # ═══════════════════════════════════════════════════════
            def extraire_identifiant_groupe(val, patterns):
                if pd.isna(val):
                    return None
                s = str(val).upper()
                for pattern, extracteur in patterns:
                    m = re.search(pattern, s)
                    if m:
                        return extracteur(m)
                return None

            patterns_groupes = [
                (r'\bG(\d+)\b', lambda m: f"G{m.group(1)}"),
                (r'\bGRP(\d+)\b', lambda m: f"G{m.group(1)}"),
                (r'\bGROUPE\s*(\d+)\b', lambda m: f"G{m.group(1)}"),
                (r'\bGROUPE\s*([A-Z])\b', lambda m: f"G{m.group(1)}"),
                (r'\bGR\s*(\d+)\b', lambda m: f"G{m.group(1)}"),
            ]
            patterns_sous_groupes = [
                (r'\bSG(\d+)\b', lambda m: f"SG{m.group(1)}"),
                (r'\bSOUS[-\s]?GROUPE\s*(\d+)\b', lambda m: f"SG{m.group(1)}"),
                (r'\bSOUS[-\s]?GRP\s*(\d+)\b', lambda m: f"SG{m.group(1)}"),
                (r'\bSOUS[-\s]?G\s*(\d+)\b', lambda m: f"SG{m.group(1)}"),
            ]

            colonnes_texte = df_edt_etu.select_dtypes(include=['object']).columns.tolist()
            colonnes_texte = [c for c in colonnes_texte if c not in ['h_norm', 'j_norm', 'Promotion_Mappee', 'Groupe_Detecte', 'SousGroupe_Detecte']]

            df_edt_etu["Groupe_Detecte"] = None
            df_edt_etu["SousGroupe_Detecte"] = None

            for col in colonnes_texte:
                detected_g = df_edt_etu[col].apply(lambda x: extraire_identifiant_groupe(x, patterns_groupes))
                df_edt_etu["Groupe_Detecte"] = df_edt_etu["Groupe_Detecte"].combine_first(detected_g)
                detected_sg = df_edt_etu[col].apply(lambda x: extraire_identifiant_groupe(x, patterns_sous_groupes))
                df_edt_etu["SousGroupe_Detecte"] = df_edt_etu["SousGroupe_Detecte"].combine_first(detected_sg)

            # ═══════════════════════════════════════════════════════
            # FILTRAGE INTELLIGENT
            # ═══════════════════════════════════════════════════════
            grp_etu_norm = groupe_etu.replace(" ", "").replace("-", "").upper() if groupe_etu else ""
            sg_etu_norm = sous_groupe_etu.replace(" ", "").replace("-", "").upper() if sous_groupe_etu else ""

            mask_commun = df_edt_etu["Groupe_Detecte"].isna() & df_edt_etu["SousGroupe_Detecte"].isna()

            if grp_etu_norm or sg_etu_norm:
                mask_mon_groupe = pd.Series([False] * len(df_edt_etu))
                mask_mon_sg = pd.Series([False] * len(df_edt_etu))

                if grp_etu_norm:
                    mask_mon_groupe = df_edt_etu["Groupe_Detecte"].astype(str).str.replace(" ", "").str.upper() == grp_etu_norm
                if sg_etu_norm:
                    mask_mon_sg = df_edt_etu["SousGroupe_Detecte"].astype(str).str.replace(" ", "").str.upper() == sg_etu_norm

                if grp_etu_norm and sg_etu_norm:
                    mask_sg_specific = df_edt_etu["SousGroupe_Detecte"].notna() & mask_mon_sg
                    mask_g_sans_sg = df_edt_etu["Groupe_Detecte"].notna() & df_edt_etu["SousGroupe_Detecte"].isna() & mask_mon_groupe
                    df_edt_etu = df_edt_etu[mask_commun | mask_sg_specific | mask_g_sans_sg].copy()
                elif grp_etu_norm:
                    df_edt_etu = df_edt_etu[mask_commun | mask_mon_groupe].copy()
                elif sg_etu_norm:
                    df_edt_etu = df_edt_etu[mask_commun | mask_mon_sg].copy()
            else:
                st.info("ℹ️ Aucun groupe/sous-groupe détecté dans votre fiche étudiant. Affichage des cours communs.")
                df_edt_etu = df_edt_etu[mask_commun].copy()

            # Info récap
            filtre_info = []
            if groupe_etu: filtre_info.append(f"Groupe **{groupe_etu}**")
            if sous_groupe_etu: filtre_info.append(f"Sous-groupe **{sous_groupe_etu}**")
            if filtre_info:
                st.info(f"🔍 Filtrage appliqué : {' | '.join(filtre_info)}")
            else:
                st.info("🔍 Affichage des cours communs à toute la promotion.")

            if df_edt_etu.empty:
                st.warning(f"⚠️ Aucun cours trouvé pour **{promo_etu}** avec vos critères de groupement.")
            else:
                st.success(f"🎓 {len(df_edt_etu)} séance(s) trouvée(s) pour vous.")

                # ─── GRILLE EDT ───
                df_edt_etu["h_norm"] = df_edt_etu["Horaire"].apply(_norm)
                df_edt_etu["j_norm"] = df_edt_etu["Jours"].apply(_norm)

                horaires_ref = [
                    "8h - 9h30", "9h30 - 11h", "11h - 12h30",
                    "12h30 - 14h", "14h - 15h30", "15h30 - 17h"
                ]
                jours_ref = ["dimanche", "lundi", "mardi", "mercredi", "jeudi"]

                def _fmt_html(rows):
                    items = []
                    for _, r in rows.iterrows():
                        code_up = str(r["Code"]).upper()
                        if "COURS" in code_up:
                            nat, color, bg = "📘", "#1e40af", "#dbeafe"
                        elif "TD" in code_up:
                            nat, color, bg = "📗", "#166534", "#dcfce7"
                        else:
                            nat, color, bg = "🔴", "#991b1b", "#fee2e2"

                        badges = ""
                        if pd.notna(r.get("Groupe_Detecte")):
                            badges += (f"<div style='display:inline-block;background:#7c3aed;"
                                      f"color:white;padding:1px 6px;border-radius:4px;"
                                      f"font-size:10px;font-weight:700;margin-bottom:4px;margin-right:4px;'>"
                                      f"👥 {r['Groupe_Detecte']}</div>")
                        if pd.notna(r.get("SousGroupe_Detecte")):
                            badges += (f"<div style='display:inline-block;background:#059669;"
                                      f"color:white;padding:1px 6px;border-radius:4px;"
                                      f"font-size:10px;font-weight:700;margin-bottom:4px;'>"
                                      f"🔹 {r['SousGroupe_Detecte']}</div>")

                        items.append(
                            f"<div style='margin-bottom:6px;padding:8px;border-left:4px solid {color};"
                            f"background-color:{bg};border-radius:6px;text-align:left;'>"
                            f"{badges}"
                            f"<b style='color:{color};font-size:13px;'>{nat} {r['Enseignements']}</b><br>"
                            f"<span style='font-size:12px;color:#334155;'>👤 {r['Enseignants']}</span><br>"
                            f"<span style='font-size:11px;color:#64748b;'>📍 {r['Lieu']}</span>"
                            f"</div>"
                        )
                    return "".join(items)

                def _fmt_text(rows):
                    items = []
                    for _, r in rows.iterrows():
                        code_up = str(r["Code"]).upper()
                        if "COURS" in code_up: nat = "COURS"
                        elif "TD" in code_up: nat = "TD"
                        else: nat = "TP"
                        g_info = ""
                        if pd.notna(r.get("Groupe_Detecte")): g_info += f" [Grp:{r['Groupe_Detecte']}]"
                        if pd.notna(r.get("SousGroupe_Detecte")): g_info += f" [SG:{r['SousGroupe_Detecte']}]"
                        items.append(f"{nat} – {r['Enseignements']}{g_info}\n{r['Enseignants']}\nSalle: {r['Lieu']}")
                    return "\n────────\n".join(items)

                grouped_html = df_edt_etu.groupby(["j_norm", "h_norm"]).apply(_fmt_html, include_groups=False)
                grouped_text = df_edt_etu.groupby(["j_norm", "h_norm"]).apply(_fmt_text, include_groups=False)

                grid_html = grouped_html.unstack("j_norm") if not grouped_html.empty else pd.DataFrame()
                grid_text = grouped_text.unstack("j_norm") if not grouped_text.empty else pd.DataFrame()

                jours_present = [j for j in jours_ref if j in grid_html.columns]
                h_present = [h for h in horaires_ref if _norm(h) in grid_html.index]

                if not jours_present or not h_present:
                    st.info("ℹ️ Impossible de construire la grille (données incomplètes après filtrage).")
                    st.dataframe(
                        df_edt_etu[['Jours', 'Horaire', 'Enseignements', 'Enseignants', 'Lieu', 'Groupe_Detecte', 'SousGroupe_Detecte']],
                        use_container_width=True, hide_index=True
                    )
                else:
                    grid_html = grid_html.reindex(index=[_norm(h) for h in h_present], columns=jours_present).fillna("")
                    grid_text = grid_text.reindex(index=[_norm(h) for h in h_present], columns=jours_present).fillna("")

                    h_labels = { _norm(h): h for h in h_present }
                    j_labels = { j: j.capitalize() for j in jours_present }

                    grid_html.index = [h_labels.get(i, i) for i in grid_html.index]
                    grid_html.columns = [j_labels.get(c, c) for c in grid_html.columns]
                    grid_text.index = [h_labels.get(i, i) for i in grid_text.index]
                    grid_text.columns = [j_labels.get(c, c) for c in grid_text.columns]

                    st.markdown("### 📋 Votre emploi du temps hebdomadaire")
                    st.write(grid_html.to_html(escape=False), unsafe_allow_html=True)

                    # ═══════════════════════════════════════════════════════
                    # EXPORTS : HTML + EXCEL + PDF
                    # ═══════════════════════════════════════════════════════
                    groupe_suffix = f"_{groupe_etu}" if groupe_etu else ""
                    sg_suffix = f"_SG{sous_groupe_etu}" if sous_groupe_etu else ""
                    base_filename = f"EDT_{nom_etu.replace(' ', '_')}_{promo_etu}{groupe_suffix}{sg_suffix}"

                    # ─── HTML ───
                    html_doc = f"""<!DOCTYPE html>
<html lang="fr"><head><meta charset="UTF-8"><title>EDT — {nom_etu}</title>
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap');
body{{font-family:'Inter','Segoe UI',Arial,sans-serif;background:linear-gradient(135deg,#f1f5f9 0%,#e2e8f0 100%);margin:0;padding:30px;color:#1e293b;}}
.container{{max-width:1200px;margin:auto;background:white;border-radius:16px;box-shadow:0 10px 40px rgba(0,0,0,0.08);overflow:hidden;}}
.header{{background:linear-gradient(135deg,#1E3A8A 0%,#3B82F6 100%);color:white;padding:30px;text-align:center;}}
.header h1{{margin:0;font-size:22px;}}.header p{{margin:8px 0 0 0;opacity:0.9;font-size:14px;}}
.badge{{display:inline-block;background:#D4AF37;color:#1E3A8A;padding:4px 14px;border-radius:20px;font-size:11px;font-weight:700;margin-top:10px;}}
.content{{padding:30px;}}table{{width:100%;border-collapse:collapse;margin-top:15px;}}
th{{background-color:#0f172a;color:white;padding:14px;text-align:center;font-size:13px;border:1px solid #e2e8f0;position:sticky;top:0;}}
td{{padding:14px;border:1px solid #e2e8f0;vertical-align:top;font-size:12px;}}
tr:nth-child(even){{background-color:#f8fafc;}}
.footer{{text-align:center;padding:20px;color:#94a3b8;font-size:12px;border-top:1px solid #f1f5f9;}}
@media print{{body{{background:white;padding:0;}}.container{{box-shadow:none;border-radius:0;}}}}
</style></head>
<body>
<div class="container">
<div class="header">
<h1>📅 Emploi du Temps Individuel</h1>
<p>{nom_etu} — Promotion {promo_etu}{f' ({groupe_etu})' if groupe_etu else ''}{f' — Sous-groupe {sous_groupe_etu}' if sous_groupe_etu else ''}</p>
<span class="badge">Semestre 01 — 2026-2027</span>
</div>
<div class="content">
<p style="color:#64748b;font-size:13px;">Généré le {datetime.now().strftime('%d/%m/%Y à %H:%M')}</p>
{grid_html.to_html(escape=False)}
</div>
<div class="footer">département d'Électrotechnique — Faculté de Génie Électrique — UDL-SBA</div>
</div>
</body></html>"""

                    # ─── EXCEL ───
                    buf_xl = io.BytesIO()
                    with pd.ExcelWriter(buf_xl, engine='xlsxwriter') as writer:
                        grid_text.to_excel(writer, sheet_name='Mon_EDT', startrow=2)
                        wb = writer.book
                        ws = writer.sheets['Mon_EDT']
                        title_fmt = wb.add_format({'bold': True, 'font_size': 14, 'font_color': '#1E3A8A', 'align': 'center', 'valign': 'vcenter'})
                        info_fmt = wb.add_format({'italic': True, 'align': 'center', 'font_size': 10, 'font_color': '#64748b'})
                        hdr_fmt = wb.add_format({'bold': True, 'bg_color': '#1E3A8A', 'font_color': 'white', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
                        idx_fmt = wb.add_format({'bold': True, 'bg_color': '#f1f5f9', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
                        cell_fmt = wb.add_format({'border': 1, 'valign': 'top', 'text_wrap': True, 'font_size': 10})
                        alt_fmt = wb.add_format({'border': 1, 'valign': 'top', 'text_wrap': True, 'font_size': 10, 'bg_color': '#F8FAFC'})

                        ws.merge_range(0, 0, 0, len(grid_text.columns), f"EDT Individuel — {nom_etu}", title_fmt)
                        ws.merge_range(1, 0, 1, len(grid_text.columns), f"Promotion {promo_etu} | Généré le {datetime.now().strftime('%d/%m/%Y')}", info_fmt)
                        ws.set_column(0, 0, 16)
                        ws.set_column(1, len(grid_text.columns), 28)

                        for col_num, val in enumerate(grid_text.columns, start=1):
                            ws.write(2, col_num, val, hdr_fmt)
                        ws.write(2, 0, "HORAIRE", hdr_fmt)

                        for row_num, (horaire, row) in enumerate(grid_text.iterrows(), start=3):
                            fmt = alt_fmt if row_num % 2 == 0 else cell_fmt
                            ws.write(row_num, 0, horaire, idx_fmt)
                            for col_num, val in enumerate(row, start=1):
                                ws.write(row_num, col_num, val, fmt)
                            n_lines = max([str(v).count('\n') + 1 for v in row] + [1])
                            ws.set_row(row_num, max(40, n_lines * 14))

                        ws.freeze_panes(3, 1)

                    # ─── PDF ───
                    buf_pdf = io.BytesIO()
                    pdf_data = None
                    try:
                        from fpdf import FPDF

                        class EtudiantEDTPDF(FPDF):
                            def header(self):
                                self.set_font('Arial', 'B', 9)
                                self.set_text_color(30, 58, 138)
                                t = "Plateforme EDT — UDL-SBA | Semestre 01 2026-2027".encode('latin-1','ignore').decode('latin-1')
                                self.cell(0, 6, t, 0, 1, 'C')
                                self.set_draw_color(212, 175, 55)
                                self.line(10, self.get_y(), self.w - 10, self.get_y())
                                self.ln(3)
                            def footer(self):
                                self.set_y(-15)
                                self.set_font('Arial', 'I', 8)
                                self.set_text_color(128, 128, 128)
                                self.cell(0, 10, f'Page {self.page_no()}', 0, 0, 'C')

                        def _san(text):
                            if not text: return ""
                            t = str(text)
                            repl = {"'":"'","'":"'","–":"-","—":"-","…":"...","«":"\"","»":"\"","œ":"oe","Œ":"OE",
                                    "à":"a","â":"a","ä":"a","á":"a","ã":"a","å":"a","è":"e","é":"e","ê":"e","ë":"e",
                                    "ì":"i","í":"i","î":"i","ï":"i","ò":"o","ó":"o","ô":"o","ö":"o","ù":"u","ú":"u","û":"u","ü":"u",
                                    "ç":"c","ñ":"n","ÿ":"y","ý":"y","À":"A","Â":"A","Ä":"A","Á":"A","Ã":"A","È":"E","É":"E","Ê":"E","Ë":"E",
                                    "Ì":"I","Í":"I","Î":"I","Ï":"I","Ò":"O","Ó":"O","Ô":"O","Ö":"O","Ù":"U","Ú":"U","Û":"U","Ü":"U","Ç":"C","Ñ":"N"}
                            for o,n in repl.items(): t=t.replace(o,n)
                            return t.encode('latin-1','ignore').decode('latin-1')

                        pdf = EtudiantEDTPDF(orientation='L', unit='mm', format='A4')
                        pdf.set_auto_page_break(auto=True, margin=15)
                        pdf.add_page()

                        pdf.set_font("Arial", "B", 13)
                        pdf.set_text_color(30, 58, 138)
                        pdf.cell(0, 8, _san(f"EDT Individuel — {nom_etu}"), 0, 1, "C")
                        pdf.set_font("Arial", "I", 9)
                        pdf.set_text_color(100, 100, 100)
                        sub = f"Promotion {promo_etu}"
                        if groupe_etu: sub += f" | Groupe {groupe_etu}"
                        if sous_groupe_etu: sub += f" | Sous-groupe {sous_groupe_etu}"
                        pdf.cell(0, 5, _san(sub), 0, 1, "C")
                        pdf.ln(3)

                        n_cols = len(grid_text.columns)
                        page_w = pdf.w - 20
                        col_h = 25
                        col_w = (page_w - col_h) / n_cols if n_cols > 0 else page_w

                        pdf.set_font("Arial", "B", 8)
                        pdf.set_fill_color(30, 58, 138)
                        pdf.set_text_color(255, 255, 255)
                        pdf.cell(col_h, 8, "HORAIRE", 1, 0, "C", True)
                        for h in grid_text.columns:
                            pdf.cell(col_w, 8, _san(h), 1, 0, "C", True)
                        pdf.ln()

                        pdf.set_text_color(0, 0, 0)
                        pdf.set_font("Arial", "", 7)
                        for idx, (horaire, row) in enumerate(grid_text.iterrows()):
                            bg = (248, 250, 252) if idx % 2 == 0 else (255, 255, 255)
                            pdf.set_fill_color(*bg)
                            pdf.set_font("Arial", "B", 7.5)
                            pdf.cell(col_h, 10, _san(horaire), 1, 0, "C", True)
                            pdf.set_font("Arial", "", 7)
                            for val in row:
                                pdf.cell(col_w, 10, _san(val)[:60], 1, 0, "C", True)
                            pdf.ln()

                        pdf_data = bytes(pdf.output())
                    except Exception as e:
                        st.caption(f"ℹ️ Export PDF indisponible : {e}")

                    # ─── BOUTONS DE TÉLÉCHARGEMENT ───
                    st.markdown("### 📥 Télécharger mon EDT")
                    c1, c2, c3 = st.columns(3)

                    c1.download_button(
                        label="🌐 HTML",
                        data=html_doc,
                        file_name=f"{base_filename}.html",
                        mime="text/html",
                        use_container_width=True,
                        key="dl_edt_etudiant_html"
                    )
                    c2.download_button(
                        label="📊 Excel",
                        data=buf_xl.getvalue(),
                        file_name=f"{base_filename}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        key="dl_edt_etudiant_xlsx"
                    )
                    if pdf_data:
                        c3.download_button(
                            label="📄 PDF",
                            data=pdf_data,
                            file_name=f"{base_filename}.pdf",
                            mime="application/pdf",
                            use_container_width=True,
                            key="dl_edt_etudiant_pdf"
                        )
                    else:
                        c3.button("📄 PDF", disabled=True, use_container_width=True)
