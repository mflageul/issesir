# 📖 GUIDE UTILISATEUR COMPLET - APPLICATION ISESIR

## 🎯 PRÉSENTATION GÉNÉRALE

**Application ISESIR by Q&T - Analyse des enquêtes de satisfaction**

Application professionnelle développée pour le Pole QoS et Transformation SI Retail du Réseau Clubs Bouygues Telecom. Elle permet l'analyse automatisée des enquêtes de satisfaction avec détection intelligente d'incohérences et génération de rapports détaillés.

## 🚀 DÉMARRAGE RAPIDE

### 1️⃣ LANCEMENT DE L'APPLICATION

1. **Double-cliquez** sur `lancer_application.bat`
2. **Attendez** l'installation automatique (première utilisation)
3. **Accédez** à l'interface web : http://localhost:5000
4. **Ne fermez jamais** la fenêtre CMD (arrêterait l'application)

### 2️⃣ INTERFACE PRINCIPALE

L'interface affiche :
- **Logo Q&T** et **Logo ISESIR** 
- **Titre** : "Application ISESIR by Q&T - Analyse des enquêtes de satisfaction"
- **Zone de chargement** des 4 fichiers Excel requis
- **Boutons d'action** pour génération des rapports

## 📁 FICHIERS REQUIS

### Types de fichiers à charger :

1. **📊 Enquêtes satisfaction** (.xlsx/.xls)
   - Contient les réponses aux enquêtes
   - Colonnes : Dossier RCBT, Questions, Commentaires, etc.

2. **🎫 Tickets clients** (.xlsx/.xls)
   - Informations sur les tickets de support
   - Colonnes : Numéro, Créé par, Compte, Description, etc.

3. **🏢 Référence équipes/sites** (.xlsx/.xls)
   - Mapping Login → Site
   - Colonnes : Login, Site

4. **💼 Comptes clients** (.xlsx/.xls)
   - Informations des comptes
   - Colonnes : Code compte, etc.

### Formats supportés :
- ✅ Fichiers .xlsx (Excel moderne)
- ✅ Fichiers .xls (Excel ancien)
- ✅ Fallback CSV automatique si Excel échoue

## 🔄 WORKFLOW D'UTILISATION

### Étape 1 : Chargement des fichiers
1. **Cliquez** sur chaque zone de chargement
2. **Sélectionnez** vos 4 fichiers Excel
3. **Validez** que tous les fichiers sont chargés (✅ vert)

### Étape 2 : Génération du rapport global
1. **Cliquez** sur "Générer le rapport global"
2. **Attendez** le traitement (masquage automatique des rapports individuels)
3. **Consultez** les résultats dans l'interface

### Étape 3 : Validation des incohérences (si détectées)
1. **Examinez** les incohérences signalées
2. **Validez ou corrigez** chaque incohérence
3. **Enregistrez** vos corrections dans le système

### Étape 4 : Rapports individuels
1. **Sélectionnez** un site ou collaborateur
2. **Générez** le rapport individuel
3. **Téléchargez** le PDF généré

## 📊 CALCULS ET MÉTRIQUES

### Indicateurs calculés :

- **Taux de fermeture** : Tickets fermés / Tickets totaux
- **Taux de satisfaction Q1** : Réponses positives / Total réponses
- **Taux de réponse boutique** : Réponses boutique / Total enquêtes
- **Nb réponses Q1** : Nombre total de réponses à la question 1

### Objectifs de référence :
- 🎯 Taux fermeture : 13%
- 🎯 Satisfaction Q1 : 92%
- 🎯 Réponse boutique : 30%
- 🎯 Jamais de réponse : 70%

## 🔍 DÉTECTION D'INCOHÉRENCES

### Types d'incohérences détectées :

1. **Commentaires négatifs avec notes positives**
   - Mots-clés négatifs : "déçu", "mauvais", "long", "lent", etc.
   - Suggestion automatique de note corrigée

2. **Mots de contraste avec discordance**
   - Mots : "mais", "cependant", "toutefois", "or", etc.
   - Analyse contextuelle fine

3. **Expressions mitigées**
   - "bien accueilli mais déçu par..."
   - "satisfait cependant..."

### Processus de validation :
1. **Examination** du commentaire original
2. **Vérification** de la note suggérée
3. **Validation ou correction** manuelle
4. **Traçabilité** complète des modifications

## 📈 RAPPORTS GÉNÉRÉS

### Rapport global :
- **Statistiques générales** par site et catégorie
- **Graphiques de satisfaction** et fermeture
- **Tableaux détaillés** par boutique
- **Détection automatique** d'incohérences

### Rapports individuels :
- **Analyse spécifique** par site/collaborateur
- **PDF téléchargeable** avec graphiques
- **Données filtrées** et personnalisées

## 🗂️ HISTORIQUE ET TRAÇABILITÉ

### Fonctionnalités :
- **Sauvegarde automatique** de tous les traitements
- **Historique des rapports** avec dates et métadonnées
- **Traçabilité des corrections** d'incohérences
- **Base de données** SQLite intégrée

## ⚙️ CARACTÉRISTIQUES TECHNIQUES

### Installation :
- ✅ **Sans droits administrateur** (mode --user)
- ✅ **Installation automatique** des dépendances
- ✅ **Compatible entreprise** (politiques de sécurité)

### Compatibilité :
- 🖥️ **Windows** 7, 8, 10, 11
- 🐍 **Python** 3.7+ requis
- 🌐 **Navigateurs** : Chrome, Firefox, Edge, Safari

### Performance :
- ⚡ **Traitement rapide** de gros volumes de données
- 💾 **Optimisation mémoire** pour fichiers volumineux
- 🔄 **Gestion d'erreurs** robuste et récupération automatique

## 🆘 DÉPANNAGE

### Problèmes courants :

**❌ "ModuleNotFoundError: xlrd"**
- ✅ Solution : L'installation automatique corrige cela

**❌ "Can't find workbook in OLE2 compound document"**
- ✅ Solution : Système multi-moteurs intégré (openpyxl, xlrd, CSV)

**❌ "Address already in use"**
- ✅ Solution : Fermez les autres instances de l'application

**❌ Fichiers non reconnus**
- ✅ Solution : Vérifiez le format Excel ou resauvegardez en .xlsx

### Messages d'aide intégrés :
- 💬 **Toasts informatifs** lors des actions
- 📱 **Instructions contextuelles** avec emojis
- ⚠️ **Alertes** en cas de problème

## 📞 SUPPORT ET CONTACT

**Responsable technique :**
- **Nom** : Flageul Martin
- **Service** : Pole QoS et Transformation SI Retail
- **Direction** : Exploitation et Système d'Information Retail
- **Entité** : Réseau Clubs Bouygues Telecom

**En cas de problème :**
1. Consultez ce guide utilisateur
2. Vérifiez les messages d'erreur dans la console CMD
3. Contactez le responsable technique si nécessaire

## 🔄 MISES À JOUR

L'application intègre :
- **Corrections automatiques** des erreurs Excel
- **Améliorations UX/UI** continues
- **Optimisations** de performance
- **Nouvelles fonctionnalités** selon les besoins

---

**Version finale** : Application ISESIR prête pour production
**Date** : Août 2025
**Statut** : Déployable sans droits administrateur