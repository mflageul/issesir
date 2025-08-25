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

REM Vérifier si Python est installé
python --version > nul 2>&1
if %errorlevel% neq 0 (
    echo ❌ ERREUR : Python n'est pas installé ou accessible
    echo.
    echo 💡 Solutions :
    echo    1. Installez Python depuis https://python.org
    echo    2. Ajoutez Python au PATH système
    echo    3. Redémarrez l'ordinateur après installation
    echo.
    pause
    exit /b 1
)

REM Installation sans droits administrateur
echo 🔧 Installation/vérification des dépendances (mode utilisateur)...
echo    - flask : Framework web
echo    - pandas : Traitement de données Excel
echo    - xlrd : Support fichiers .xls anciens
echo    - openpyxl : Support fichiers .xlsx modernes
echo    - matplotlib, seaborn : Graphiques et visualisations
echo    - reportlab, pillow : Génération PDF et images
echo    - werkzeug : Sécurité fichiers
echo.

REM Installation en mode utilisateur (--user) pour éviter les droits admin
echo 📦 Installation en cours (mode utilisateur - sans droits admin)...
python -m pip install --user flask pandas "xlrd>=2.0.1" openpyxl matplotlib seaborn reportlab pillow werkzeug
if %errorlevel% neq 0 (
    echo.
    echo ❌ ERREUR : Échec de l'installation automatique
    echo.
    echo 💡 Solutions alternatives (SANS droits administrateur) :
    echo    1. Installez manuellement avec cette commande :
    echo       python -m pip install --user flask pandas "xlrd>=2.0.1" openpyxl matplotlib seaborn reportlab pillow werkzeug
    echo    2. Vérifiez votre connexion internet
    echo    3. Contactez le support informatique si problème réseau d'entreprise
    echo    4. IMPORTANT : Ne nécessite PAS de droits administrateur
    echo.
    echo 🔍 Tentative de vérification des packages déjà installés...
    python -c "import flask, pandas, xlrd, openpyxl, matplotlib, seaborn, reportlab, PIL, werkzeug; print('✅ Toutes les dépendances sont disponibles')" 2>nul
    if %errorlevel% neq 0 (
        echo ❌ Certains packages manquent encore
        pause
        exit /b 1
    )
)

echo ✅ Toutes les dépendances installées avec succès

echo ✅ Installation terminée - Application prête
echo.

REM Créer les dossiers nécessaires
if not exist "uploads" mkdir uploads
if not exist "ressources" mkdir ressources
if not exist "session_cache" mkdir session_cache
if not exist "tmp" mkdir tmp

echo 🚀 Démarrage de l'application...
echo.
echo ============================================================
echo 🚀 APPLICATION RCBT - PRÊTE À L'USAGE
echo    Version avec Validation des Incohérences
echo ============================================================
echo 📋 Fonctionnalités disponibles :
echo    ✓ Génération de rapports RCBT complets
echo    ✓ Validation manuelle des incohérences
echo    ✓ Interface web intuitive
echo    ✓ Traçabilité complète des corrections
echo ------------------------------------------------------------
echo 🌐 Accédez à l'application sur :
echo    👉 http://localhost:5000
echo    👉 http://127.0.0.1:5000
echo ------------------------------------------------------------
echo 💡 Utilisation :
echo    1. Chargez vos 4 fichiers Excel
echo    2. Générez le rapport
echo    3. Validez les incohérences si détectées
echo ============================================================
echo.
echo ⚠️  IMPORTANT - NE PAS FERMER CETTE FENÊTRE
echo    Cette console doit rester ouverte pour que l'application
echo    web fonctionne. La fermer arrêtera le serveur RCBT.
echo.
echo 🔄 Démarrage du serveur...
echo.

REM Lancer l'application Python
python app.py

REM Si l'application s'arrête, afficher un message
echo.
echo ============================================================
echo ⚠️  APPLICATION ARRÊTÉE
echo ============================================================
echo 💡 L'application s'est arrêtée. Causes possibles :
echo    • Fermeture manuelle du serveur (Ctrl+C)
echo    • Erreur dans l'application
echo    • Port 5000 déjà utilisé
echo.
echo 🔧 Pour relancer l'application :
echo    • Double-cliquez sur ce fichier à nouveau
echo    • Ou exécutez : python app.py
echo.
echo ============================================================
pause