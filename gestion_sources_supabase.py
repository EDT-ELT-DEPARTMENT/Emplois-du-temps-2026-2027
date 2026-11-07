# -*- coding: utf-8 -*-
"""Gestion CRUD des classeurs sources de la plateforme ELT via Supabase Storage."""
from __future__ import annotations
import io, re, unicodedata
from pathlib import Path
from typing import Any
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

BUCKET_DEFAUT = "elt-sources"
SOURCES_GESTION = {
    "Étudiants": {
        "canonical": "Liste des étudiants_2026-2027.xlsx",
        "aliases": ["Liste des étudiants_2026-2027(5).xlsx", "Liste des étudiants_2026-2027(4).xlsx", "Liste des étudiants_2026-2027(3).xlsx"],
        "sheets": ["Feuille1", "Suivi des emails", "Suivi des téléphones"],
    },
    "Enseignants": {
        "canonical": "Permanents-Vacataires-ELT2-2026-2027.xlsx",
        "aliases": ["Permanents-Vacataires-ELT2-2026-2027(2).xlsx", "Permanents-Vacataires-ELT2-2026-2027(1).xlsx"],
        "sheets": ["Permanents-Vacataires", "Répartiion pour Examen"],
    },
    "Emploi du temps": {
        "canonical": "dataEDT-ELT-S1-2027.xlsx",
        "aliases": ["dataEDT-ELT-S1-2027(3).xlsx", "dataEDT-ELT-S1-2027(1).xlsx"],
        "sheets": ["Sheet1"],
    },
    "Coordonnées des étudiants": {
        "canonical": "Vos coordonnées (réponses).xlsx",
        "aliases": ["Vos coordonnées (réponses)(1).xlsx"],
        "sheets": ["Réponses au formulaire 1"],
    },
}


def _norm(v):
    t = unicodedata.normalize("NFKD", str(v or "")).encode("ascii", "ignore").decode("ascii").strip().lower()
    return re.sub(r"[^a-z0-9]+", " ", t).strip()


def _download(storage, bucket, names):
    for name in names:
        try:
            data = storage.from_(bucket).download(name)
            if hasattr(data, "read"):
                data = data.read()
            if isinstance(data, bytearray):
                data = bytes(data)
            if isinstance(data, bytes) and data:
                return data, name
        except Exception:
            continue
    return None, None


def _upload(storage, bucket, name, payload, content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"):
    opts = {"upsert": "true", "content-type": content_type}
    try:
        storage.from_(bucket).upload(name, payload, file_options=opts)
    except TypeError:
        storage.from_(bucket).update(name, payload, file_options=opts)


def _resolve(storage, bucket, base_dir, cfg):
    names = [cfg["canonical"], *cfg.get("aliases", [])]
    payload, found = _download(storage, bucket, names) if storage else (None, None)
    if payload:
        return payload, found
    for name in names:
        path = Path(base_dir) / name
        if path.is_file():
            return path.read_bytes(), name
    return None, None


def _read_workbook(payload):
    return pd.read_excel(io.BytesIO(payload), sheet_name=None, dtype=object)


def _excel_value(value):
    if value is None:
        return None
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()
    try:
        if pd.isna(value):
            return None
    except Exception:
        pass
    if hasattr(value, "item"):
        try:
            return value.item()
        except Exception:
            pass
    return value


def _replace_sheet(payload, sheet_name, frame):
    wb = load_workbook(io.BytesIO(payload))
    ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.create_sheet(sheet_name)
    if ws.max_row:
        ws.delete_rows(1, ws.max_row)
    clean = frame.copy()
    clean.columns = [str(c) for c in clean.columns]
    for row in dataframe_to_rows(clean, index=False, header=True):
        ws.append([_excel_value(v) for v in row])
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def _validate(label, sheet, frame):
    if frame is None or len(frame.columns) == 0:
        return False, "La feuille doit conserver ses en-têtes de colonnes."
    if label == "Étudiants" and sheet == "Feuille1":
        col = next((c for c in frame.columns if _norm(c) == "mat etudiant"), None)
        if col:
            vals = frame[col].fillna("").astype(str).str.strip()
            vals = vals[vals.ne("")]
            duplicates = vals[vals.duplicated(keep=False)]
            if not duplicates.empty:
                return False, "Mat. Etudiant en doublon : " + ", ".join(duplicates.unique()[:5])
    # Les doublons NOM+PRÉNOM déjà présents dans le répertoire du personnel
    # sont tolérés ; l'application doit distinguer les personnes par nom complet.
    return True, ""


def _save(storage, bucket, base_dir, cfg, payload, sheet, frame):
    ok, message = _validate(cfg["label"], sheet, frame)
    if not ok:
        raise ValueError(message)
    new_payload = _replace_sheet(payload, sheet, frame)
    canonical = cfg["canonical"]
    _upload(storage, bucket, canonical, new_payload)
    (Path(base_dir) / canonical).write_bytes(new_payload)
    return new_payload


def render_admin_data_manager(storage, base_dir, bucket=BUCKET_DEFAUT, supabase_client=None, actor="Administrateur"):
    if not storage:
        st.error("Connexion Supabase indisponible : gestion des sources impossible.")
        return
    c1, c2 = st.columns([5, 1])
    with c1:
        st.markdown("## 🗂️ Gestion générale des données")
    with c2:
        if st.button("↩️ Retour", key="retour_gestion_sources"):
            st.session_state["ouvrir_gestion_sources"] = False
            st.rerun()
    st.caption("Ajoutez, recherchez, modifiez ou supprimez des lignes. Les changements sont enregistrés dans le classeur central de Supabase.")
    label = st.selectbox("Source à gérer", list(SOURCES_GESTION), key="admin_source_choice")
    cfg = dict(SOURCES_GESTION[label]); cfg["label"] = label
    payload, found_name = _resolve(storage, bucket, base_dir, cfg)
    if not payload:
        st.warning(f"La source {cfg['canonical']} n'est pas encore dans Supabase. Importez le classeur original pour l'initialiser.")
        upload = st.file_uploader("Importer le classeur source", type=["xlsx"], key=f"initial_upload_{label}")
        if upload and st.button("☁️ Importer dans Supabase", key=f"initial_save_{label}"):
            raw = upload.getvalue()
            _read_workbook(raw)
            _upload(storage, bucket, cfg["canonical"], raw)
            (Path(base_dir) / cfg["canonical"]).write_bytes(raw)
            st.success("Source initialisée dans Supabase."); st.rerun()
        return
    try:
        sheets = _read_workbook(payload)
    except Exception as exc:
        st.error(f"Impossible de lire le classeur : {exc}"); return
    if not sheets:
        st.warning("Le classeur ne contient aucune feuille."); return
    default_sheet = next((s for s in cfg.get("sheets", []) if s in sheets), list(sheets)[0])
    sheet = st.selectbox("Feuille à modifier", list(sheets), index=list(sheets).index(default_sheet), key=f"sheet_{label}")
    frame = sheets[sheet].copy()
    st.caption(f"Source : `{found_name}` · Feuille : `{sheet}` · {len(frame):,} lignes · {len(frame.columns)} colonnes")
    search_cols = [c for c in frame.columns if any(t in _norm(c) for t in ["mat etudiant", "nom", "prenom", "promotion", "email", "enseignant"])]
    search_col, search_value = "Toutes les lignes", ""
    if search_cols:
        search_col = st.selectbox("Filtrer les lignes (facultatif)", ["Toutes les lignes"] + search_cols, key=f"search_col_{label}_{sheet}")
        search_value = st.text_input("Valeur recherchée", key=f"search_value_{label}_{sheet}", placeholder="Matricule, nom, prénom, promotion…")
    if search_col != "Toutes les lignes" and search_value.strip():
        mask = frame[search_col].fillna("").astype(str).str.contains(re.escape(search_value.strip()), case=False, na=False)
        filtered = frame.loc[mask].copy()
        if len(filtered) > 500:
            st.warning("Plus de 500 résultats : précisez votre recherche."); return
        if filtered.empty:
            st.info("Aucun résultat. Vous pouvez effacer le filtre pour ajouter une ligne.")
    else:
        filtered = frame.copy()
    if len(filtered) > 2500:
        st.info("Le classeur contient plus de 2 500 lignes. Utilisez le filtre de recherche avant de modifier."); return
    st.markdown("#### ✏️ Modifier, ajouter ou supprimer des lignes")
    editor_key = f"editor_{label}_{sheet}_{len(frame)}"
    edited = st.data_editor(filtered, num_rows="dynamic", use_container_width=True, hide_index=True, key=editor_key, height=520)
    if st.button("💾 Enregistrer les modifications dans Supabase", type="primary", use_container_width=True, key=f"save_{label}_{sheet}"):
        try:
            if search_col != "Toutes les lignes" and search_value.strip():
                updated = frame.copy()
                original_ids = list(filtered.index)
                kept_ids = [i for i in original_ids if i in edited.index]
                for idx in kept_ids:
                    updated.loc[idx, edited.columns] = edited.loc[idx, edited.columns]
                new_ids = [i for i in edited.index if i not in original_ids]
                if new_ids:
                    updated = pd.concat([updated, edited.loc[new_ids]], ignore_index=True)
                removed = [i for i in original_ids if i not in edited.index]
                if removed:
                    updated = updated.drop(index=removed, errors="ignore").reset_index(drop=True)
            else:
                updated = edited.copy()
            _save(storage, bucket, Path(base_dir), cfg, payload, sheet, updated)
            if supabase_client is not None:
                try:
                    supabase_client.table("journal_modifications_sources").insert({
                        "utilisateur": str(actor)[:250], "source": cfg["canonical"], "feuille": str(sheet)[:150],
                        "operation": "MODIFICATION_TABLEAU", "nombre_lignes": int(len(updated)),
                        "details": {"colonnes": [str(c) for c in updated.columns]},
                    }).execute()
                except Exception:
                    pass
            st.session_state.pop(editor_key, None)
            st.success(f"{len(updated):,} ligne(s) enregistrée(s) dans Supabase. Les autres feuilles sont conservées.")
            st.rerun()
        except Exception as exc:
            st.error(f"Échec de l'enregistrement : {exc}")
    st.divider()
    st.markdown("#### 📤 Remplacer le classeur complet")
    st.caption("Cette opération remplace toutes les feuilles du classeur. Vérifiez le fichier avant validation.")
    replacement = st.file_uploader("Nouveau classeur complet", type=["xlsx"], key=f"replace_{label}")
    if replacement and st.button("Remplacer le classeur complet dans Supabase", key=f"replace_save_{label}"):
        try:
            raw = replacement.getvalue(); _read_workbook(raw)
            _upload(storage, bucket, cfg["canonical"], raw)
            (Path(base_dir) / cfg["canonical"]).write_bytes(raw)
            if supabase_client is not None:
                try:
                    supabase_client.table("journal_modifications_sources").insert({
                        "utilisateur": str(actor)[:250], "source": cfg["canonical"], "feuille": "*",
                        "operation": "REMPLACEMENT_CLASSEUR", "details": {"nom_fichier_importe": replacement.name},
                    }).execute()
                except Exception:
                    pass
            st.success("Classeur remplacé dans Supabase."); st.rerun()
        except Exception as exc:
            st.error(f"Remplacement impossible : {exc}")
    st.divider()
    st.markdown("#### 📦 Autres fichiers sources")
    st.caption("Le fichier surveillances_2027.xlsx n'est pas encore renseigné : il sera ajouté plus tard. Les PDF, images, CSV et JSON peuvent déjà être archivés ici.")
    other = st.file_uploader("Déposer un autre fichier source", type=["xlsx", "xls", "csv", "json", "pdf", "png", "jpg", "jpeg"], key="other_source_upload")
    if other and st.button("Enregistrer cet autre fichier dans Supabase", key="other_source_save"):
        try:
            _upload(storage, bucket, Path(other.name).name, other.getvalue(), other.type or "application/octet-stream")
            st.success(f"Fichier {other.name} archivé dans Supabase.")
        except Exception as exc:
            st.error(f"Échec de l'archivage : {exc}")
    st.info("Identifiant étudiant : Mat. Etudiant (Mat. BAC n'est pas unique). Enseignant : NOM + PRÉNOM. Les colonnes source sont conservées.")
