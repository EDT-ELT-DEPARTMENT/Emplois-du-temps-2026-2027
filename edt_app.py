# =============================================================================
# VUE ADMIN : EDT INDIVIDUEL ÉTUDIANTS (PAR GROUPE / SOUS-GROUPE)
# =============================================================================
def render_edt_etudiant_admin(df_edt, df_etu):
    """Génère les EDT individuels étudiants avec filtrage intelligent Groupe/Sous-Groupe."""
    import re
    import io
    import zipfile
    from datetime import datetime
    
    st.markdown("""
        <style>
        .etu-card { background: linear-gradient(135deg, #1E3A8A 0%, #3B82F6 100%); 
                    padding: 18px; border-radius: 14px; color: white; margin-bottom: 18px; }
        .etu-card h3 { margin: 0; font-size: 20px; }
        .etu-card p { margin: 6px 0 0 0; opacity: 0.9; font-size: 13px; }
        .badge-grp { display:inline-block; background:#D4AF37; color:#1E3A8A; 
                     padding: 3px 10px; border-radius: 12px; font-size: 11px; font-weight: 700; margin-right: 6px; }
        </style>
    """, unsafe_allow_html=True)

    st.subheader("🎓 Générateur d'EDT Individuel — Filtrage par Groupe & Sous-Groupe")

    # -------------------------------------------------------------------------
    # A. VÉRIFICATION DES DONNÉES
    # -------------------------------------------------------------------------
    if df_edt is None or df_edt.empty:
        st.error("❌ Données EDT non disponibles.")
        return
    if df_etu is None or df_etu.empty:
        st.error("❌ Données étudiants non disponibles.")
        return

    # S'assurer que Nom_Complet existe dans df_etu
    if "Nom_Complet" not in df_etu.columns:
        col_n = next((c for c in df_etu.columns if c.strip().upper() == "NOM"), None)
        col_p = next((c for c in df_etu.columns if c.strip().upper() in ["PRÉNOM", "PRENOM"]), None)
        if col_n and col_p:
            df_etu = df_etu.copy()
            df_etu["Nom_Complet"] = df_etu[col_n].astype(str).str.strip().str.upper() + " " + df_etu[col_p].astype(str).str.strip().str.title()
        else:
            st.error("❌ Impossible de construire les noms complets des étudiants.")
            return

    # -------------------------------------------------------------------------
    # B. DÉTECTION INTELLIGENTE DES COLONNES GROUPE / SOUS-GROUPE
    # -------------------------------------------------------------------------
    cols_map = detecter_colonnes_etudiant(df_etu)
    col_groupe = cols_map.get('groupe', None)
    col_sg = cols_map.get('sous_groupe', None)

    # Fallback si la détection automatique échoue
    if not col_groupe:
        for c in df_etu.columns:
            if str(c).strip().upper() in ["GROUPE", "GRP", "GROUP"]:
                col_groupe = c
                break
    if not col_sg:
        for c in df_etu.columns:
            if str(c).strip().upper() in ["SOUS GROUPE", "SOUS-GROUPE", "SOUSGROUPE", "SOUS_GRP", "SG"]:
                col_sg = c
                break

    # -------------------------------------------------------------------------
    # C. SÉLECTION DE LA PROMOTION
    # -------------------------------------------------------------------------
    promos_dispo = sorted(df_etu["Promotion"].dropna().unique().tolist())
    if not promos_dispo:
        st.warning("⚠️ Aucune promotion trouvée dans le fichier étudiants.")
        return

    promo_sel = st.selectbox("🎓 Sélectionner la promotion :", promos_dispo, key="admin_edt_promo")

    df_etu_promo = df_etu[df_etu["Promotion"].astype(str).str.strip() == promo_sel].copy()
    if df_etu_promo.empty:
        st.warning(f"⚠️ Aucun étudiant trouvé pour {promo_sel}.")
        return

    # -------------------------------------------------------------------------
    # D. AFFICHAGE RÉCAPITULATIF DES GROUPES DÉTECTÉS
    # -------------------------------------------------------------------------
    st.markdown("### 📋 Récapitulatif des groupes détectés")
    recap_cols = ["Nom_Complet"]
    if col_groupe: recap_cols.append(col_groupe)
    if col_sg: recap_cols.append(col_sg)
    
    # Nettoyage pour l'affichage
    df_display = df_etu_promo[[c for c in recap_cols if c in df_etu_promo.columns]].copy()
    st.dataframe(df_display, use_container_width=True, hide_index=True)

    # Stats groupes
    if col_groupe:
        grps = df_etu_promo[col_groupe].astype(str).str.strip().value_counts()
        st.caption("📊 Répartition : " + " | ".join([f"**{k}** : {v} étud." for k, v in grps.items()]))

    # -------------------------------------------------------------------------
    # E. PRÉPARATION DE L'EDT (Extraction Groupe & Sous-Groupe depuis l'EDT)
    # -------------------------------------------------------------------------
    promo_mapped = mapper_promotion(promo_sel)
    df_edt_work = df_edt.copy()
    df_edt_work["Promotion_Mappee"] = df_edt_work["Promotion"].apply(mapper_promotion)
    df_edt_promo = df_edt_work[df_edt_work["Promotion_Mappee"] == promo_mapped].copy()

    if df_edt_promo.empty:
        st.warning(f"⚠️ Aucun cours trouvé dans l'EDT pour la promotion mappée : {promo_mapped}")
        return

    # --- Fonction d'extraction robuste G/SG dans l'EDT ---
    def _extract_gsg(val):
        if pd.isna(val):
            return None, None
        s = str(val).upper()
        # Groupe : G1, G2...
        m_g = re.search(r'\bG(\d+)\b', s)
        g = f"G{m_g.group(1)}" if m_g else None
        # Sous-groupe : SG11, SG12, SG21, SG22...
        m_sg = re.search(r'\bSG(\d+)\b', s)
        sg = f"SG{m_sg.group(1)}" if m_sg else None
        return g, sg

    # On scanne Code, Lieu ET Enseignements pour trouver les groupes/sous-groupes
    g_code, sg_code = zip(*df_edt_promo["Code"].apply(_extract_gsg))
    g_lieu, sg_lieu = zip(*df_edt_promo["Lieu"].apply(_extract_gsg))
    g_ens,  sg_ens  = zip(*df_edt_promo["Enseignements"].apply(_extract_gsg))

    df_edt_promo["Groupe_EDT"] = pd.Series(g_code, index=df_edt_promo.index).fillna(
        pd.Series(g_lieu, index=df_edt_promo.index)).fillna(pd.Series(g_ens, index=df_edt_promo.index))
    df_edt_promo["SG_EDT"] = pd.Series(sg_code, index=df_edt_promo.index).fillna(
        pd.Series(sg_lieu, index=df_edt_promo.index)).fillna(pd.Series(sg_ens, index=df_edt_promo.index))

    # -------------------------------------------------------------------------
    # F. SÉLECTION DE L'ÉTUDIANT
    # -------------------------------------------------------------------------
    etu_liste = ["Tous les étudiants"] + sorted(df_etu_promo["Nom_Complet"].dropna().unique().tolist())
    etu_sel = st.selectbox("👤 Générer l'EDT pour :", etu_liste, key="admin_edt_etu_sel")

    cibles = df_etu_promo["Nom_Complet"].unique().tolist() if etu_sel == "Tous les étudiants" else [etu_sel]

    # -------------------------------------------------------------------------
    # G. CONSTANTES DE LA GRILLE
    # -------------------------------------------------------------------------
    HORAIRES_REF = ["8h - 9h30", "9h30 - 11h", "11h - 12h30", "12h30 - 14h", "14h - 15h30", "15h30 - 17h"]
    JOURS_REF = ["Dimanche", "Lundi", "Mardi", "Mercredi", "Jeudi"]
    
    def _norm_h(x):
        if not x: return ""
        s = str(x).strip().lower().replace(" ", "").replace("–", "-")
        s = s.replace(":00", "").replace("h00", "h")
        for ref in HORAIRES_REF:
            if s == ref.lower().replace(" ", ""): return ref
        return str(x).strip()
    def _norm_j(x):
        if not x: return ""
        s = str(x).strip().lower()
        for ref in JOURS_REF:
            if s == ref.lower(): return ref
        return str(x).strip()

    # -------------------------------------------------------------------------
    # H. FONCTION DE CONSTRUCTION DE LA GRILLE HTML
    # -------------------------------------------------------------------------
    def _build_grid_html(df_source, nom_etu, grp_etu, sg_etu):
        if df_source.empty:
            return None, "<p>Aucun cours trouvé.</p>"

        df_g = df_source.copy()
        df_g["h_norm"] = df_g["Horaire"].apply(_norm_h)
        df_g["j_norm"] = df_g["Jours"].apply(_norm_j)
        df_g = df_g[df_g["h_norm"].isin(HORAIRES_REF)]
        df_g = df_g[df_g["j_norm"].isin(JOURS_REF)]

        if df_g.empty:
            return None, "<p>Aucun cours sur les créneaux standards.</p>"

        def _type_emoji(code):
            c = str(code).upper()
            if "COURS" in c: return "📘", "#dbeafe", "#1e40af"
            if "TD" in c:    return "📗", "#dcfce7", "#166534"
            if "TP" in c:    return "🔴", "#fee2e2", "#991b1b"
            return "⚪", "#f3f4f6", "#374151"

        def _fmt_cell(rows):
            out = []
            for _, r in rows.iterrows():
                em, bg, col = _type_emoji(r.get("Code", ""))
                badge_g = ""
                if pd.notna(r.get("Groupe_EDT")):
                    badge_g = (f"<div style='display:inline-block;background:#7c3aed;color:white;"
                               f"padding:1px 6px;border-radius:4px;font-size:10px;font-weight:700;"
                               f"margin-bottom:4px;'>👥 {r['Groupe_EDT']}"
                               f"{f' / {r['SG_EDT']}' if pd.notna(r.get('SG_EDT')) else ''}</div><br>")
                lines = [
                    f"<b style='color:{col};font-size:12px;'>{em} {r.get('Enseignements','')}</b>",
                    f"<span style='font-size:10px;color:#64748b;'>👤 {r.get('Enseignants','')}</span>",
                    f"<span style='font-size:10px;color:#64748b;'>📍 {r.get('Lieu','')}</span>"
                ]
                out.append(f"<div style='background:{bg};border-left:3px solid {col};border-radius:4px;"
                           f"padding:4px;margin:2px 0;line-height:1.3;'>{badge_g}{'<br>'.join(lines)}</div>")
            return "".join(out)

        grouped = df_g.groupby(["j_norm", "h_norm"]).apply(_fmt_cell, include_groups=False)
        grid = grouped.unstack(fill_value="") if not grouped.empty else pd.DataFrame()

        jours_ok = [j for j in JOURS_REF if j in grid.columns]
        h_ok = [h for h in HORAIRES_REF if h in grid.index]
        if not jours_ok or not h_ok:
            return None, "<p>Données incomplètes pour construire la grille.</p>"

        grid = grid.reindex(index=h_ok, columns=jours_ok).fillna("")

        # HTML complet
        thead = "<tr><th style='background:#1E3A8A;color:white;padding:10px;width:110px;'>HORAIRE</th>" + \
                "".join([f"<th style='background:#1E3A8A;color:white;padding:10px;font-size:12px;'>{j}</th>" for j in grid.columns]) + "</tr>"
        tbody = ""
        for h, row in grid.iterrows():
            tbody += f"<tr><td style='background:#f1f5f9;font-weight:bold;text-align:center;padding:10px;'>{h}</td>"
            for val in row:
                tbody += f"<td style='border:1px solid #e2e8f0;padding:6px;vertical-align:top;'>{val}</td>"
            tbody += "</tr>"

        html_doc = f"""<!DOCTYPE html>
<html lang='fr'><head><meta charset='UTF-8'><title>EDT {nom_etu}</title>
<style>
body{{font-family:'Segoe UI',Arial,sans-serif;background:#f8fafc;padding:20px;margin:0;color:#1e293b;}}
.container{{max-width:1200px;margin:auto;background:white;border-radius:12px;box-shadow:0 4px 12px rgba(0,0,0,0.08);overflow:hidden;}}
.header{{background:linear-gradient(135deg,#1E3A8A,#3B82F6);color:white;padding:20px;text-align:center;}}
.header h1{{margin:0;font-size:20px;}} .header p{{margin:6px 0 0 0;opacity:0.9;font-size:13px;}}
.badge{{display:inline-block;background:#D4AF37;color:#1E3A8A;padding:4px 14px;border-radius:20px;font-size:11px;font-weight:700;margin-top:10px;}}
.content{{padding:20px;}}
table{{width:100%;border-collapse:collapse;table-layout:fixed;}}
th{{position:sticky;top:0;z-index:10;}}
td{{word-wrap:break-word;}}
.footer{{text-align:center;padding:15px;color:#94a3b8;font-size:11px;border-top:1px solid #f1f5f9;}}
@media print{{body{{background:white;padding:0;}} .container{{box-shadow:none;border-radius:0;}}}}
</style></head><body>
<div class='container'>
<div class='header'><h1>📅 EDT Individuel — {nom_etu}</h1>
<p>Promotion {promo_sel}{f' | Groupe {grp_etu}' if grp_etu else ''}{f' | Sous-Groupe {sg_etu}' if sg_etu else ''}</p>
<span class='badge'>Semestre 01 — 2026-2027</span></div>
<div class='content'><p style='color:#64748b;font-size:13px;'>Généré le {datetime.now().strftime('%d/%m/%Y à %H:%M')}</p>
<table><thead>{thead}</thead><tbody>{tbody}</tbody></table></div>
<div class='footer'>département d'Électrotechnique — Faculté de Génie Électrique — UDL-SBA</div>
</div></body></html>"""
        return grid, html_doc

    # -------------------------------------------------------------------------
    # I. GÉNÉRATION POUR UN OU TOUS LES ÉTUDIANTS
    # -------------------------------------------------------------------------
    if st.button("⚙️ Générer l'EDT", type="primary", use_container_width=True):
        zip_buffer = io.BytesIO()
        has_data = False

        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for nom_etu in cibles:
                row_etu = df_etu_promo[df_etu_promo["Nom_Complet"] == nom_etu]
                if row_etu.empty:
                    continue
                row_etu = row_etu.iloc[0]

                # Récupération Groupe / Sous-Groupe
                grp_etu = str(row_etu.get(col_groupe, "")).strip().upper() if col_groupe else ""
                sg_etu = str(row_etu.get(col_sg, "")).strip().upper() if col_sg else ""

                # -----------------------------------------------------------------
                # FILTRAGE INTELLIGENT
                # -----------------------------------------------------------------
                mask_commun = df_edt_promo["Groupe_EDT"].isna() & df_edt_promo["SG_EDT"].isna()
                
                # Cours du groupe (sans sous-groupe précis) OU cours spécifique au sous-groupe
                mask_groupe = pd.Series(False, index=df_edt_promo.index)
                if grp_etu:
                    mask_groupe = (df_edt_promo["Groupe_EDT"] == grp_etu) & (df_edt_promo["SG_EDT"].isna())
                
                mask_sg = pd.Series(False, index=df_edt_promo.index)
                if grp_etu and sg_etu:
                    mask_sg = (df_edt_promo["Groupe_EDT"] == grp_etu) & (df_edt_promo["SG_EDT"] == sg_etu)

                df_edt_etu = df_edt_promo[mask_commun | mask_groupe | mask_sg].copy()

                # Construction
                grid, html_doc = _build_grid_html(df_edt_etu, nom_etu, grp_etu, sg_etu)

                if etu_sel == "Tous les étudiants":
                    # Ajout au ZIP
                    safe_name = nom_etu.replace(" ", "_").replace("/", "-")
                    zf.writestr(f"EDT_{safe_name}.html", html_doc)
                    has_data = True
                else:
                    # Affichage direct + téléchargement unique
                    st.markdown(f"""
                        <div class="etu-card">
                            <h3>👤 {nom_etu}</h3>
                            <p>Promotion : <b>{promo_sel}</b> 
                            {f'<span class="badge-grp">{grp_etu}</span>' if grp_etu else ''}
                            {f'<span class="badge-grp">{sg_etu}</span>' if sg_etu else ''}
                            — <b>{len(df_edt_etu)}</b> séance(s) trouvée(s)</p>
                        </div>
                    """, unsafe_allow_html=True)

                    if grid is not None:
                        st.write(grid.to_html(escape=False), unsafe_allow_html=True)
                    else:
                        st.info(html_doc)

                    # --- BOUTONS DE TÉLÉCHARGEMENT ---
                    c1, c2 = st.columns(2)
                    c1.download_button(
                        "🌐 Télécharger EDT (HTML)", html_doc,
                        f"EDT_{nom_etu.replace(' ', '_')}.html", "text/html",
                        use_container_width=True, key=f"dl_html_{nom_etu.replace(' ', '_')}"
                    )
                    
                    # Excel grille
                    if grid is not None and not grid.empty:
                        buf_xl = io.BytesIO()
                        with pd.ExcelWriter(buf_xl, engine='xlsxwriter') as writer:
                            grid.to_excel(writer, sheet_name='EDT')
                            wb = writer.book
                            ws = writer.sheets['EDT']
                            hdr_fmt = wb.add_format({'bold':True,'bg_color':'#1E3A8A','font_color':'white','border':1,'align':'center','valign':'vcenter'})
                            idx_fmt = wb.add_format({'bold':True,'bg_color':'#f1f5f9','border':1,'align':'center','valign':'vcenter'})
                            cell_fmt = wb.add_format({'border':1,'valign':'top','text_wrap':True,'font_size':10})
                            for col_num, val in enumerate(grid.columns, start=1):
                                ws.write(0, col_num, val, hdr_fmt)
                            ws.write(0, 0, "HORAIRE", hdr_fmt)
                            for row_num, (idx, row) in enumerate(grid.iterrows(), start=1):
                                ws.write(row_num, 0, idx, idx_fmt)
                                for col_num, val in enumerate(row, start=1):
                                    ws.write(row_num, col_num, val, cell_fmt)
                            ws.set_column(0, 0, 16)
                            ws.set_column(1, len(grid.columns), 32)
                            for r in range(1, len(grid)+1):
                                ws.set_row(r, 80)
                        c2.download_button(
                            "📊 Télécharger EDT (Excel)", buf_xl.getvalue(),
                            f"EDT_{nom_etu.replace(' ', '_')}.xlsx",
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True, key=f"dl_xl_{nom_etu.replace(' ', '_')}"
                        )
                    return  # On sort car un seul étudiant

        # ---------------------------------------------------------------------
        # J. TÉLÉCHARGEMENT MASSIF (ZIP)
        # ---------------------------------------------------------------------
        if etu_sel == "Tous les étudiants" and has_data:
            st.success(f"✅ {len(cibles)} EDT individuels générés avec succès.")
            st.download_button(
                "🗜️ Télécharger le pack ZIP (tous les EDT)",
                zip_buffer.getvalue(),
                f"Pack_EDT_{promo_sel.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d')}.zip",
                "application/zip",
                use_container_width=True,
                key="dl_zip_all_edt"
            )
        elif etu_sel == "Tous les étudiants" and not has_data:
            st.warning("⚠️ Aucun EDT n'a pu être généré. Vérifiez que les groupes/sous-groupes dans l'EDT correspondent bien à ceux des étudiants.")
