# MANUEL D'UTILISATION COMPLET
# APPLICATION ISESIR
## Analyse des Enquêtes de Satisfaction - Réseau Clubs Bouygues Telecom

---

**Version :** 1.0 - Août 2025  
**Développé par :** Pole QoS et Transformation SI Retail  
**Contact :** Flageul Martin - mflageul@rcbt.fr

---

## TABLE DES MATIÈRES

1. [Présentation générale](#1-présentation-générale)
2. [Interface principale](#2-interface-principale)
3. [Rapport global - Page 1 : Statistiques générales](#3-rapport-global---page-1--statistiques-générales)
4. [Rapport global - Page 2 : Analyse des commentaires](#4-rapport-global---page-2--analyse-des-commentaires)
5. [Système de validation des incohérences](#5-système-de-validation-des-incohérences)
6. [Rapports individuels](#6-rapports-individuels)
7. [Historique et traçabilité](#7-historique-et-traçabilité)
8. [Lexique des termes](#8-lexique-des-termes)
9. [Formules de calcul](#9-formules-de-calcul)
10. [Guide de résolution des problèmes](#10-guide-de-résolution-des-problèmes)

---

## 1. PRÉSENTATION GÉNÉRALE

### 1.1 Objectif de l'application
L'application ISESIR permet d'analyser automatiquement les enquêtes de satisfaction des boutiques du Réseau Clubs Bouygues Telecom. Elle traite les données d'enquêtes, détecte les incohérences potentielles et génère des rapports détaillés avec visualisations.

### 1.2 Fichiers requis
L'application nécessite 4 fichiers Excel :

1. **Fichier Enquêtes** (`asmt_metric_result_*.xlsx`)
   - Contient les réponses aux enquêtes de satisfaction
   - Colonnes principales : N° dossier, Note satisfaction, Commentaires

2. **Fichier Tickets** (`sn_customerservice_case_*.xlsx`)
   - Informations sur les tickets client
   - Colonnes principales : N° dossier, Date création, Site, Collaborateur

3. **Fichier Référence** (`reference_equipes_sites_*.xlsx`)
   - Mapping entre collaborateurs et sites
   - Colonnes principales : Collaborateur, Site, Équipe

4. **Fichier Comptes** (`customer_account_*.xlsx`)
   - Informations complémentaires sur les comptes clients
   - Utilisé pour enrichir les analyses

### 1.3 Fonctionnalités principales
- Génération de rapports globaux avec graphiques
- Détection intelligente d'incohérences dans les commentaires
- Système de validation manuelle des incohérences
- Génération de rapports individuels par site ou collaborateur
- Historique complet des traitements
- Export des données et logs de validation

---

## 2. INTERFACE PRINCIPALE

### 2.1 Écran d'accueil
L'interface présente plusieurs sections :

**Zone de chargement des fichiers :**
- 4 zones de glisser-déposer pour les fichiers Excel
- Validation automatique du format et du contenu
- Indicateurs visuels de statut (rouge/vert)

**Boutons d'action principaux :**
- "Générer le rapport global" : Lance l'analyse complète
- "Rapports individuels" : Accès aux analyses par site/collaborateur
- "Historique" : Consultation des traitements précédents

**Zone d'information :**
- Affichage des messages de statut
- Compteurs de données chargées
- Alertes et notifications

### 2.2 Navigation
- Interface responsive s'adaptant à tous les écrans
- Navigation intuitive avec breadcrumbs
- Boutons de retour et d'annulation disponibles
- Sauvegarde automatique de la session

---

## 3. RAPPORT GLOBAL - PAGE 1 : STATISTIQUES GÉNÉRALES

### 3.1 Présentation de la page
La page 1 du rapport global présente une vue d'ensemble des performances avec 6 tableaux principaux et leurs graphiques associés.

### 3.2 Tableau 1 : Vue d'ensemble
**Contenu :**
- Total tickets traités
- Nombre de boutiques analysées  
- Période d'analyse
- Nombre de collaborateurs

**Calculs :**
- **Total tickets** = Nombre total d'enregistrements dans le fichier tickets
- **Boutiques** = Nombre unique de sites dans les données
- **Collaborateurs** = Nombre unique de collaborateurs ayant traité des tickets

### 3.3 Tableau 2 : Répartition par canal
**Contenu :**
- Tickets Boutiques vs Autres canaux
- Pourcentages respectifs

**Calculs :**
- **Tickets Boutiques** = Tickets où Canal = "Boutique" OU Site contient des codes boutique
- **Autres canaux** = Total tickets - Tickets Boutiques
- **% Boutiques** = (Tickets Boutiques / Total tickets) × 100

**Graphique associé :** Camembert de répartition

### 3.4 Tableau 3 : Taux de closure
**Contenu :**
- Tickets fermés vs ouverts
- Taux de closure global

**Calculs :**
- **Tickets fermés** = Tickets avec Statut = "Fermé" ou "Résolu"
- **Tickets ouverts** = Tickets avec Statut ≠ "Fermé" et ≠ "Résolu"
- **Taux de closure** = (Tickets fermés / Total tickets) × 100

**Graphique associé :** Graphique en barres horizontales

### 3.5 Tableau 4 : Enquêtes de satisfaction
**Contenu :**
- Nombre de réponses Q1 (satisfaction)
- Taux de réponse aux enquêtes

**Calculs :**
- **Nb réponses Q1** = Nombre d'enregistrements avec une note de satisfaction renseignée
- **Taux de réponse** = (Nb réponses Q1 / Total tickets) × 100

**Seuils d'évaluation :**
- Vert (Bon) : > 15%
- Orange (Moyen) : 5-15%
- Rouge (Faible) : < 5%

### 3.6 Tableau 5 : Distribution satisfaction
**Contenu :**
- Répartition des 4 niveaux de satisfaction
- Pourcentages pour chaque niveau

**Niveaux de satisfaction :**
1. Très satisfaisant (Note 4)
2. Satisfaisant (Note 3)
3. Peu satisfaisant (Note 2)
4. Très peu satisfaisant (Note 1)

**Calculs :**
- **Nombre par niveau** = Compte des réponses pour chaque note
- **Pourcentage** = (Nombre niveau / Total réponses Q1) × 100

**Graphique associé :** Camembert avec couleurs :
- Vert foncé : Très satisfaisant
- Vert clair : Satisfaisant  
- Orange : Peu satisfaisant
- Rouge : Très peu satisfaisant

### 3.7 Tableau 6 : Taux de satisfaction global
**Contenu :**
- Calcul du taux de satisfaction global
- Évaluation qualitative

**Calculs :**
- **Taux satisfaction** = ((Très satisfaisant + Satisfaisant) / Total réponses Q1) × 100

**Seuils d'évaluation :**
- Excellent : > 90%
- Bon : 75-90%
- Moyen : 60-75%
- Faible : < 60%

---

## 4. RAPPORT GLOBAL - PAGE 2 : ANALYSE DES COMMENTAIRES

### 4.1 Présentation de la page
La page 2 se concentre sur l'analyse qualitative des commentaires clients et la détection d'incohérences.

### 4.2 Tableau 1 : Statistiques des commentaires
**Contenu :**
- Nombre total de commentaires
- Pourcentage de tickets avec commentaires
- Répartition par longueur de commentaire

**Calculs :**
- **Total commentaires** = Nombre d'enregistrements avec commentaire non vide
- **% avec commentaires** = (Total commentaires / Total tickets) × 100
- **Commentaires courts** = Commentaires < 50 caractères
- **Commentaires longs** = Commentaires ≥ 50 caractères

### 4.3 Tableau 2 : Détection d'incohérences
**Contenu :**
- Nombre d'incohérences détectées
- Types d'incohérences identifiées
- Statut de validation

**Types d'incohérences :**

**Type 1 - Amélioration possible :**
- Note faible (1-2) mais commentaire positif
- Suggère une note plus élevée
- Exemple : Note "Très peu satisfaisant" + commentaire "Service excellent"

**Type 2 - Attention requise :**
- Note élevée (3-4) mais commentaire négatif
- Suggère une note plus faible
- Exemple : Note "Très satisfaisant" + commentaire "Déçu du service"

### 4.4 Système de détection automatique
**Mots-clés positifs surveillés :**
- excellent, parfait, super, formidable, remarquable
- satisfait, content, ravi, enchanté
- professionnel, compétent, à l'écoute
- rapide, efficace, aimable

**Mots-clés négatifs surveillés :**
- déçu, mécontent, insatisfait, frustré
- lent, long, attente, retard
- incompétent, désagréable, impoli
- problème, erreur, bug, panne

**Expressions de contraste :**
- "mais", "cependant", "néanmoins", "toutefois"
- "malgré", "bien que", "même si"

### 4.5 Graphiques d'analyse
**Graphique 1 :** Répartition des incohérences par type
**Graphique 2 :** Évolution temporelle des commentaires
**Graphique 3 :** Corrélation notes/sentiment des commentaires

---

## 5. SYSTÈME DE VALIDATION DES INCOHÉRENCES

### 5.1 Processus de validation
Lorsque des incohérences sont détectées, l'application propose un système de validation manuelle.

### 5.2 Interface de validation
**Éléments affichés pour chaque incohérence :**
- Numéro de dossier client
- Note originale vs note suggérée
- Commentaire complet du client
- Contexte (site, collaborateur, date)
- Boutons d'action (Valider/Rejeter/Reporter)

### 5.3 Actions possibles
**Valider l'incohérence :**
- Confirme que la détection est correcte
- La note suggérée remplace la note originale
- Impact sur les statistiques finales

**Rejeter l'incohérence :**
- Indique que la détection est incorrecte
- Maintient la note originale
- Améliore l'apprentissage du système

**Reporter la décision :**
- Marque l'incohérence pour révision ultérieure
- Maintient la note originale temporairement
- Permet traitement différé

### 5.4 Traçabilité
Toutes les actions de validation sont enregistrées :
- Horodatage des décisions
- Utilisateur ayant validé
- Justification des choix
- Impact sur les métriques

---

## 6. RAPPORTS INDIVIDUELS

### 6.1 Types de rapports individuels
L'application permet de générer des rapports ciblés :

**Par site :**
- Analyse dédiée à une boutique spécifique
- Comparaison avec la moyenne réseau
- Évolution dans le temps

**Par collaborateur :**
- Performance individuelle d'un conseiller
- Historique des évaluations
- Points d'amélioration identifiés

### 6.2 Contenu des rapports individuels
**Section 1 : Identification**
- Nom du site/collaborateur
- Période analysée
- Nombre de tickets traités

**Section 2 : Métriques clés**
- Taux de satisfaction personnel
- Nombre d'enquêtes reçues
- Répartition des notes

**Section 3 : Analyse qualitative**
- Commentaires positifs remarquables
- Points d'attention identifiés
- Suggestions d'amélioration

**Section 4 : Comparaison**
- Position par rapport à la moyenne
- Ranking relatif
- Évolution tendancielle

### 6.3 Filtrage et sélection
**Critères de filtrage disponibles :**
- Période : Sélection libre de dates
- Site : Liste déroulante des boutiques
- Collaborateur : Recherche par nom
- Type d'enquête : Satisfaction, réclamation, etc.

---

## 7. HISTORIQUE ET TRAÇABILITÉ

### 7.1 Consultation de l'historique
L'onglet "Historique" présente tous les traitements effectués :

**Colonnes affichées :**
- Date/heure de traitement
- Fichiers utilisés
- Type de rapport généré
- Métriques principales
- Actions de validation
- Liens de téléchargement

### 7.2 Gestion de l'historique
**Actions possibles :**
- Consultation des rapports précédents
- Re-téléchargement des fichiers
- Suppression des entrées obsolètes
- Export des logs de traitement

### 7.3 Audit et conformité
**Informations d'audit :**
- Traçabilité complète des modifications
- Justifications des validations manuelles
- Horodatage précis de toutes les actions
- Identification des utilisateurs

---

## 8. LEXIQUE DES TERMES

### 8.1 Termes métier
**RCBT :** Réseau Clubs Bouygues Telecom

**Boutique :** Point de vente physique du réseau

**Ticket :** Demande client enregistrée dans le système

**Enquête de satisfaction :** Questionnaire envoyé au client après résolution

**Q1 :** Question 1 de l'enquête (note de satisfaction globale)

**Incohérence :** Divergence entre la note attribuée et le sentiment du commentaire

### 8.2 Termes techniques
**Taux de closure :** Pourcentage de tickets fermés/résolus

**Taux de réponse :** Pourcentage de clients ayant répondu à l'enquête

**NPS :** Net Promoter Score (indicateur de recommandation)

**Distribution :** Répartition statistique des valeurs

**Validation :** Processus de vérification manuelle des détections automatiques

### 8.3 Niveaux de satisfaction
**Très satisfaisant (4) :** Client très content, expérience excellente

**Satisfaisant (3) :** Client content, expérience positive

**Peu satisfaisant (2) :** Client mitigé, expérience décevante

**Très peu satisfaisant (1) :** Client mécontent, expérience négative

---

## 9. FORMULES DE CALCUL

### 9.1 Métriques principales

**Taux de satisfaction :**
```
Taux de satisfaction = (Nombre de notes 3 et 4 / Total réponses Q1) × 100
```

**Taux de closure :**
```
Taux de closure = (Tickets fermés / Total tickets) × 100
```

**Taux de réponse aux enquêtes :**
```
Taux de réponse = (Réponses Q1 / Total tickets) × 100
```

**Pourcentage par niveau :**
```
% Niveau = (Nombre ce niveau / Total réponses Q1) × 100
```

### 9.2 Calculs de répartition

**Pourcentage boutiques :**
```
% Boutiques = (Tickets boutiques / Total tickets) × 100
```

**Pourcentage avec commentaires :**
```
% Commentaires = (Tickets avec commentaire / Total tickets) × 100
```

### 9.3 Métriques d'incohérence

**Taux de détection :**
```
Taux détection = (Incohérences trouvées / Total commentaires) × 100
```

**Taux de validation :**
```
Taux validation = (Incohérences validées / Incohérences détectées) × 100
```

### 9.4 Calculs individuels

**Performance relative :**
```
Performance = (Satisfaction individuelle / Satisfaction moyenne) × 100
```

**Écart type :**
```
Mesure la dispersion des notes autour de la moyenne
```

---

## 10. GUIDE DE RÉSOLUTION DES PROBLÈMES

### 10.1 Problèmes de chargement de fichiers

**Erreur : "Format de fichier non reconnu"**
- Vérifier que le fichier est au format .xlsx
- S'assurer que le fichier n'est pas corrompu
- Contrôler la présence des colonnes obligatoires

**Erreur : "Données manquantes"**
- Vérifier que les colonnes clés sont renseignées
- Contrôler la cohérence des identifiants
- S'assurer de la complétude des données

### 10.2 Problèmes de génération de rapport

**Erreur : "Aucune donnée à traiter"**
- Vérifier que les 4 fichiers sont chargés
- Contrôler la période de données
- S'assurer de la cohérence des jointures

**Erreur : "Calcul impossible"**
- Vérifier la qualité des données numériques
- Contrôler les valeurs nulles ou aberrantes
- S'assurer de la cohérence des référentiels

### 10.3 Problèmes de performance

**Lenteur de traitement :**
- Réduire la taille des fichiers si possible
- Fermer les autres applications
- Vérifier l'espace disque disponible

**Blocage de l'interface :**
- Attendre la fin du traitement en cours
- Rafraîchir la page si nécessaire
- Redémarrer l'application en dernier recours

### 10.4 Support technique

**Contact :**
- Email : mflageul@rcbt.fr
- Service : Pole QoS et Transformation SI Retail
- Téléphone : [À renseigner]

**Informations à fournir :**
- Version de l'application
- Description détaillée du problème
- Captures d'écran si possible
- Fichiers de test (si autorisé)

---

**Fin du manuel utilisateur**

*Ce document est régulièrement mis à jour. Vérifiez la version en ligne pour les dernières modifications.*

---

**Copyright © 2025 - Réseau Clubs Bouygues Telecom**  
**Tous droits réservés - Usage interne uniquement**