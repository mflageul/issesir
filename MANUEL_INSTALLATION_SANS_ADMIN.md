# 📘 MANUEL D'INSTALLATION SANS DROITS ADMINISTRATEUR

## 🔒 CONTRAINTES D'ENTREPRISE RESPECTÉES

Cette application a été spécialement conçue pour fonctionner **SANS droits administrateur**, respectant les politiques de sécurité d'entreprise strictes.

## 🚀 PROCÉDURE D'INSTALLATION COMPLÈTE

### 1️⃣ PRÉREQUIS SYSTÈME

- **Python 3.7+** installé sur le système
- **Connexion internet** pour télécharger les packages
- **Droits utilisateur standard** (pas d'administrateur requis)

### 2️⃣ INSTALLATION AUTOMATIQUE

1. **Téléchargez** l'archive `ISESIR_Application_Final_CORRECTED.tar.gz`
2. **Extrayez** le dossier `ISESIR_Application_Final_Download`
3. **Double-cliquez** sur `lancer_application.bat`
4. **Attendez** l'installation automatique des dépendances

### 3️⃣ MÉCANISME D'INSTALLATION

L'installation utilise le paramètre `--user` de pip :
```cmd
python -m pip install --user [packages]
```

**Avantages :**
- ✅ Installation dans le dossier utilisateur
- ✅ Aucun droit administrateur requis
- ✅ Packages disponibles pour l'utilisateur connecté
- ✅ Respect des politiques de sécurité d'entreprise

### 4️⃣ PACKAGES INSTALLÉS AUTOMATIQUEMENT

- **flask** : Framework web pour l'interface
- **pandas** : Traitement des données Excel
- **xlrd>=2.0.1** : Support fichiers .xls anciens
- **openpyxl** : Support fichiers .xlsx modernes
- **matplotlib, seaborn** : Génération graphiques
- **reportlab, pillow** : Création PDF et images
- **werkzeug** : Sécurité et gestion fichiers

### 5️⃣ RÉSOLUTION AUTOMATIQUE DES ERREURS

Le launcher intègre une gestion intelligente :
- **Lecture Excel multi-moteurs** (openpyxl, xlrd, auto-détection)
- **Fallback CSV** si Excel échoue
- **Messages d'erreur détaillés** pour diagnostic
- **Vérification automatique** des dépendances

## 🛠️ DÉPANNAGE SANS DROITS ADMIN

### Si l'installation automatique échoue :

1. **Installation manuelle** :
   ```cmd
   python -m pip install --user flask pandas "xlrd>=2.0.1" openpyxl matplotlib seaborn reportlab pillow werkzeug
   ```

2. **Vérification Python** :
   ```cmd
   python --version
   ```

3. **Test des packages** :
   ```cmd
   python -c "import flask, pandas, xlrd, openpyxl, matplotlib, seaborn, reportlab, PIL, werkzeug; print('✅ Succès')"
   ```

### En cas de problème réseau d'entreprise :

1. **Proxy d'entreprise** : Contactez le support IT
2. **Firewall** : Vérifiez l'accès à pypi.org
3. **Certificats** : Problème SSL/TLS avec le proxy

## 📞 SUPPORT TECHNIQUE

**Responsable :**
- **Nom** : Flageul Martin
- **Service** : Pole QoS et Transformation SI Retail
- **Direction** : Exploitation et Système d'Information Retail
- **Entité** : Réseau Clubs Bouygues Telecom

## ✅ VALIDATION FINALE

Une fois l'installation terminée :
1. L'application s'ouvre automatiquement dans votre navigateur
2. Interface ISESIR by Q&T disponible sur http://localhost:5000
3. Chargez vos 4 fichiers Excel requis
4. Générez vos rapports d'analyse

**IMPORTANT :** Cette solution respecte intégralement les contraintes de sécurité d'entreprise en n'utilisant que des droits utilisateur standard.