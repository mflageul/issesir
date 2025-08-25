# 🔧 RÉSOLUTION ERREUR XLRD

## ❌ PROBLÈME RENCONTRÉ
```
ModuleNotFoundError: No module named 'xlrd'
```

## ✅ SOLUTION AUTOMATIQUE

Le fichier `lancer_application.bat` a été mis à jour pour installer automatiquement **xlrd version 2.0.1+** qui supporte les fichiers .xls.

### 🚀 PROCÉDURE CORRIGÉE

1. **Supprimez l'ancien dossier** si vous l'avez déjà extrait
2. **Téléchargez la nouvelle archive** : `ISESIR_Application_Final_CORRECTED.tar.gz`
3. **Extrayez** le nouveau dossier
4. **Double-cliquez** sur `lancer_application.bat`
5. **Attendez** l'installation automatique (première fois seulement)

### 📦 DÉPENDANCES INSTALLÉES AUTOMATIQUEMENT

- ✅ **flask** : Framework web
- ✅ **pandas** : Traitement données Excel
- ✅ **xlrd>=2.0.1** : Support fichiers .xls (VERSION CORRIGÉE)
- ✅ **openpyxl** : Support fichiers .xlsx
- ✅ **matplotlib, seaborn** : Graphiques
- ✅ **reportlab, pillow** : PDF et images
- ✅ **werkzeug** : Sécurité

### 🛠️ EN CAS D'ÉCHEC D'INSTALLATION (SANS DROITS ADMIN)

Si l'installation automatique échoue :

1. **Installation manuelle mode utilisateur** :
   ```cmd
   python -m pip install --user flask pandas "xlrd>=2.0.1" openpyxl matplotlib seaborn reportlab pillow werkzeug
   ```

2. **Vérifier Python** :
   ```cmd
   python --version
   ```

3. **Vérifier les packages installés** :
   ```cmd
   python -c "import flask, pandas, xlrd, openpyxl, matplotlib, seaborn, reportlab, PIL, werkzeug; print('Succès')"
   ```

**IMPORTANT** : L'installation utilise le mode `--user` qui NE NÉCESSITE PAS de droits administrateur.

### 📞 SUPPORT

Si le problème persiste, contactez :
- **Flageul Martin**
- **Service** : Pole QoS et Transformation SI Retail

---

**Note** : Le fichier dupliqué `LANCER_APPLICATION.bat` a été supprimé pour éviter la confusion.