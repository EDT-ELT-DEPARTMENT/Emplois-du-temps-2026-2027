# ⚡ DÉPLOIEMENT EXPRESS EN 5 MINUTES

## 🚀 Version rapide (copy-paste)

Suivez ces 12 étapes exactement comme écrit. Aucune modification!

---

## ÉTAPE 1 : Préparer votre dossier

**Windows (PowerShell) :**
```powershell
mkdir C:\Users\$env:USERNAME\Desktop\mon-app-edt
cd C:\Users\$env:USERNAME\Desktop\mon-app-edt
```

**macOS/Linux (Terminal) :**
```bash
mkdir ~/Desktop/mon-app-edt
cd ~/Desktop/mon-app-edt
```

---

## ÉTAPE 2 : Copier vos fichiers

**Copier ces 4 fichiers DANS le dossier `mon-app-edt` :**
1. `Emploi_du_temps_COMPLET_AVEC_SIDEBAR.py`
2. `requirements.txt`
3. `.gitignore`
4. `README.md`

---

## ÉTAPE 3 : Initialiser Git

```bash
git init
git config --global user.name "Votre Nom"
git config --global user.email "votre_email@gmail.com"
```

---

## ÉTAPE 4 : Premier commit

```bash
git add .
git commit -m "Première version"
```

---

## ÉTAPE 5 : Créer repo GitHub

1. Aller sur : **https://github.com/new**
2. Repository name : `mon-app-edt`
3. Public (important!)
4. Cliquer "Create repository"

---

## ÉTAPE 6 : Ajouter le repo distant

**Remplacer `VotreNom` par votre username GitHub :**

```bash
git remote add origin https://github.com/VotreNom/mon-app-edt.git
git branch -M main
```

---

## ÉTAPE 7 : Uploader sur GitHub

```bash
git push -u origin main
```

**Si demandé :**
- Username : `VotreNom`
- Password : Aller sur https://github.com/settings/tokens/new
  - Cocher ✓ repo, ✓ workflow
  - Generate token
  - Coller le token

---

## ÉTAPE 8 : Créer compte Streamlit Cloud

1. Aller sur : **https://streamlit.io/cloud**
2. "Sign up"
3. Utiliser votre compte GitHub

---

## ÉTAPE 9 : Déployer l'app

1. Dashboard : **https://share.streamlit.io/**
2. "New app"
3. Remplir :
   ```
   Repository : VotreNom/mon-app-edt
   Branch     : main
   File path  : Emploi_du_temps_COMPLET_AVEC_SIDEBAR.py
   ```
4. "Deploy"

---

## ÉTAPE 10 : Attendre

```
Status : "Building Docker image..."
Puis : "🎈 App is launching..."
Enfin : "🎉 Your app is live!"
```

**Durée :** 2-3 minutes

---

## ÉTAPE 11 : Accéder à l'app

```
Votre URL : https://[votre-username]-mon-app-edt.streamlit.app
```

Copier l'URL et ouvrir dans un navigateur! ✅

---

## ÉTAPE 12 : Partager

```
Vous pouvez maintenant partager l'URL avec :
- Vos étudiants
- Vos collègues
- N'importe qui!

L'app fonctionne 24/7 gratuitement!
```

---

## 🔄 Pour les mises à jour

```bash
# Modifier votre code localement
# ...

# Puis faire :
git add .
git commit -m "Description de la modification"
git push

# ✅ L'app est à jour automatiquement!
```

---

## ✅ Checklist

- ✅ Dossier créé
- ✅ Fichiers copiés
- ✅ Git initialisé
- ✅ Commit fait
- ✅ Repo GitHub créé
- ✅ Code poussé
- ✅ Compte Streamlit Cloud
- ✅ App déployée
- ✅ URL accessible
- ✅ Partage avec d'autres

---

## 🎉 C'EST FAIT!

Votre app est en ligne et fonctionne 24/7 gratuitement!

Durée totale : **5-10 minutes** (selon la vitesse du déploiement)

---

## ⚠️ Si ça ne marche pas

### Erreur "Repository not found"
```
✅ Solution : Utiliser un repo PUBLIC (pas private)
```

### Erreur "ModuleNotFoundError"
```
✅ Solution : Ajouter le module à requirements.txt
             git add . && git commit -m "Fix" && git push
```

### L'app ne se met pas à jour
```
✅ Solution : Vérifier git status
             git add . && git commit -m "..." && git push
```

### L'app crash
```
✅ Solution : Voir les logs dans Streamlit Cloud
             Corriger l'erreur et faire git push
```

---

## 📞 Besoin d'aide?

Consultez le guide complet : **GUIDE_DEPLOIEMENT_STREAMLIT_CLOUD.md**

---

**🚀 Vous êtes prêt! Lancez le déploiement! 🚀**
