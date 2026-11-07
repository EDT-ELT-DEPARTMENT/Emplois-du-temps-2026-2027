# Migration des sources de la plateforme ELT vers Supabase

## Fichiers à déposer dans le même répertoire GitHub
- `edt_app.py` : application complète modifiée.
- `gestion_sources_supabase.py` : module d'administration CRUD des sources.

Les classeurs de données ne sont plus nécessaires dans GitHub après migration et vérification.

## 1. Préparer Supabase
1. Ouvrir Supabase > SQL Editor.
2. Exécuter `INSTALLATION_SUPABASE_SOURCES.sql`.
3. Dans Streamlit Cloud > Settings > Secrets, vérifier `SUPABASE_URL` et `SUPABASE_SERVICE_ROLE_KEY`.
4. Garder la clé Service Role secrète : ne jamais la mettre dans le code ou GitHub.

Le bucket `elt-sources` est privé. L'application utilise la clé serveur pour lire et écrire les classeurs.

## 2. Déployer le code
Committer ensemble `edt_app.py` et `gestion_sources_supabase.py` dans le même dossier. Ne pas supprimer les classeurs GitHub avant l'import initial et les tests.

## 3. Import initial des sources
1. Se connecter comme administrateur.
2. Ouvrir **Sources de données — Supabase** dans la barre latérale pour déposer les fichiers, ou **Gestion générale des données** pour initialiser chaque classeur.
3. Les versions de fichiers reconnues sont :
   - `Liste des étudiants_2026-2027(5).xlsx` -> `Liste des étudiants_2026-2027.xlsx`
   - `Permanents-Vacataires-ELT2-2026-2027(2).xlsx` -> `Permanents-Vacataires-ELT2-2026-2027.xlsx`
   - `dataEDT-ELT-S1-2027(3).xlsx` -> `dataEDT-ELT-S1-2027.xlsx`
   - `Vos coordonnées (réponses).xlsx`
4. Choisir une source et une feuille. Rechercher une ligne, modifier les cellules ou utiliser l'éditeur pour ajouter/supprimer des lignes.
5. Cliquer sur **Enregistrer les modifications dans Supabase**.

## 4. Colonnes et feuilles conservées
- Étudiants : les 83 colonnes de `Feuille1`, les 8 colonnes de `Suivi des emails` et les 8 colonnes de `Suivi des téléphones`.
- Enseignants : les 7 colonnes de `Permanents-Vacataires` et les 4 colonnes de `Répartiion pour Examen`.
- EDT : `Enseignements`, `Code`, `Enseignants`, `Horaire`, `Jours`, `Lieu`, `Promotion`, `Département`.
- Coordonnées : les 7 colonnes de `Réponses au formulaire 1`.

La modification d'une feuille conserve les autres feuilles du même classeur et les en-têtes de la feuille modifiée.

## 5. Identifiants et contrôles
- Étudiant : `Mat. Etudiant` sert d'identifiant de gestion. `Mat. BAC` n'est pas une clé unique.
- Enseignant : l'association se fait par `NOM` + `PRÉNOM`, jamais par le seul nom.
- EDT : l'éditeur existant conserve les contrôles de conflits. Les ajouts de séances sont sauvegardés dans le classeur central Supabase Storage ; la table historique `edt_data` est facultative.
- Les opérations du gestionnaire sont journalisées dans `journal_modifications_sources` si la table est accessible.

## 6. Surveillances
`surveillances_2027.xlsx` n'est pas encore renseigné. Aucune colonne n'est inventée. Le module de surveillance sera raccordé après réception du classeur réel.

## 7. Tests obligatoires avant retrait des fichiers GitHub
- rechercher un étudiant par `Mat. Etudiant` et modifier son téléphone ;
- ajouter un étudiant avec un matricule unique ;
- modifier un enseignant (qualité, grade, e-mail ou téléphone) ;
- modifier une séance EDT et vérifier les vues enseignant/étudiant ;
- redémarrer Streamlit et vérifier que les modifications persistent.

Ne retirer les classeurs du dépôt GitHub qu'après validation de ces tests.
