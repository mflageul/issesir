@echo off
chcp 65001 > nul
title APPLICATION ISESIR by Q&T - Lancement
color 0E

echo.
echo ======================================================================
echo                        -- ISESIR --
echo            by Pole QoS et Transformation SI Retail
echo                Réseau Clubs Bouygues Telecom
echo ======================================================================
echo.
echo 📊 Lancement de l'analyse des enquêtes de satisfaction...
echo.

REM Diagnostic complet Python
echo 🔍 === DIAGNOSTIC PYTHON ===
echo 🔍 Vérification de Python...
python --version 2>&1
set PYTHON_CHECK=%errorlevel%
echo Code retour Python: %PYTHON_CHECK%

if %PYTHON_CHECK% neq 0 (
    echo.
    echo ❌ ERREUR : Python n'est pas installé ou accessible
    echo.
    echo 💡 Solutions :
    echo    1. Installez Python depuis https://python.org
    echo    2. Ajoutez Python au PATH système
    echo    3. Redémarrez l'ordinateur après installation
    echo.
    echo 🔍 Diagnostic avancé...
    where python 2>nul
    if %errorlevel% neq 0 (
        echo    ❌ Python non trouvé dans le PATH système
    ) else (
        echo    ✅ Python trouvé dans le PATH
    )
    echo.
    echo ⏸️  Appuyez sur une touche pour fermer...
    pause >nul
    exit /b 1
)
echo ✅ Python détecté avec succès
echo.

REM Installation des dépendances avec diagnostic
echo 🔧 === INSTALLATION DÉPENDANCES ===
echo 📦 Installation en mode utilisateur (sans droits admin)...
echo Packages: flask pandas xlrd>=2.0.1 openpyxl matplotlib seaborn reportlab pillow werkzeug
echo.

python -m pip install --user flask pandas "xlrd>=2.0.1" openpyxl matplotlib seaborn reportlab pillow werkzeug
set INSTALL_CHECK=%errorlevel%
echo Code retour installation: %INSTALL_CHECK%

if %INSTALL_CHECK% neq 0 (
    echo.
    echo ❌ ERREUR : Échec de l'installation automatique
    echo.
    echo 🔍 Vérification des packages existants...
    python -c "import flask, pandas, xlrd, openpyxl, matplotlib, seaborn, reportlab, PIL, werkzeug; print('✅ Tous les packages disponibles')" 2>nul
    if %errorlevel% neq 0 (
        echo ❌ Certains packages manquent
        echo.
        echo 💡 Solutions alternatives :
        echo    1. Vérifiez votre connexion internet
        echo    2. Contactez le support informatique si proxy d'entreprise
        echo    3. Installation manuelle :
        echo       python -m pip install --user flask pandas "xlrd>=2.0.1" openpyxl matplotlib seaborn reportlab pillow werkzeug
        echo.
        echo ⏸️  Appuyez sur une touche pour fermer...
        pause >nul
        exit /b 1
    ) else (
        echo ✅ Packages déjà disponibles - continuons
    )
) else (
    echo ✅ Installation réussie
)
echo.

REM Créer les dossiers nécessaires
echo 🔧 === PRÉPARATION ENVIRONNEMENT ===
if not exist "uploads" mkdir uploads
if not exist "ressources" mkdir ressources
if not exist "static\images" mkdir static\images
echo ✅ Dossiers créés
echo.

REM Vérification des fichiers requis
echo 🔍 === VÉRIFICATION FICHIERS ===
if not exist "app.py" (
    echo ❌ ERREUR : Fichier app.py introuvable
    echo 💡 Vérifiez que vous êtes dans le bon dossier
    echo ⏸️  Appuyez sur une touche pour fermer...
    pause >nul
    exit /b 1
)
echo ✅ Fichier app.py trouvé
echo.

echo ======================================================================
echo                        -- ISESIR --
echo            by Pole QoS et Transformation SI Retail
echo                Réseau Clubs Bouygues Telecom
echo ======================================================================
echo 📊 Analyse avancée des enquêtes de satisfaction boutiques
echo 🔧 Fonctionnalités intégrées :
echo    ✓ Génération de rapports RCBT complets
echo    ✓ Détection intelligente d'incohérences contextuelles
echo    ✓ Système de validation centralisé avec traçabilité
echo    ✓ Interface web responsive et intuitive
echo    ✓ Rapports individuels par site/collaborateur
echo 🌐 ACCÈS À L'APPLICATION :
echo    ┌─ URL principale : http://localhost:5000
echo    └─ URL alternative : http://127.0.0.1:5000
echo ⚠️  IMPORTANT - NE PAS FERMER CETTE FENÊTRE
echo    Cette console doit rester ouverte pour que l'application
echo    web fonctionne. La fermer arrêtera le serveur RCBT.
echo 📞 Responsable technique :
echo    Contact   : Flageul Martin
echo    Service   : Pole QoS et Transformation SI Retail
echo    Direction : Exploitation et Système d'Information Retail
echo    Entité    : Réseau Clubs Bouygues Telecom
echo 💡 Guide d'utilisation rapide :
echo    1️⃣  Ouvrez votre navigateur à l'adresse ci-dessus
echo    2️⃣  Chargez vos 4 fichiers Excel requis
echo    3️⃣  Générez le rapport global d'analyse
echo    4️⃣  Validez les incohérences détectées
echo    5️⃣  Consultez/générez les rapports individuels
echo ======================================================================
echo 🚀 Initialisation du serveur web en cours...
echo ======================================================================

REM Lancer l'application Python avec gestion d'erreurs
echo 🌟 Démarrage de l'application ISESIR...
python app.py
set APP_CHECK=%errorlevel%

REM Gestion des erreurs de l'application
echo.
if %APP_CHECK% neq 0 (
    echo ❌ L'application s'est arrêtée avec une erreur (code: %APP_CHECK%)
    echo.
    echo 💡 Causes possibles :
    echo    1. Port 5000 déjà utilisé par une autre application
    echo    2. Erreur dans le code Python
    echo    3. Dépendance manquante
    echo.
    echo 🔧 Solutions :
    echo    1. Fermez toute autre instance de l'application
    echo    2. Redémarrez votre ordinateur
    echo    3. Contactez le support technique si le problème persiste
) else (
    echo ℹ️  L'application s'est fermée normalement
)

echo.
echo ⏸️  Appuyez sur une touche pour fermer cette fenêtre...
pause >nul