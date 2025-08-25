import pandas as pd
import numpy as np
from pathlib import Path
import os

class DataProcessor:
    def __init__(self):
        self.objectifs = {
            'closure_rate': 13.0,
            'satisfaction_q1': 92.0,
            'boutique_response': 30.0,
            'never_response': 70.0
        }

    def load_and_merge_data(self, files, enable_fiabilization=False):
        """Load and merge all Excel files with improved format handling"""
        print("=== CHARGEMENT DES DONNÉES ===")
        
        try:
            # Load Excel files with engine detection and fallback options
            def safe_read_excel(file_path, sheet_name=0):
                """Safely read Excel file with enhanced validation and multiple engines"""
                print(f"🔄 Lecture fichier: {os.path.basename(file_path)}")
                
                # First check if file exists and basic validation
                if not os.path.exists(file_path):
                    raise Exception(f"Fichier non trouvé: {file_path}")
                
                file_size = os.path.getsize(file_path)
                print(f"📏 Taille fichier: {file_size} bytes")
                
                if file_size == 0:
                    raise Exception(f"Fichier vide: {os.path.basename(file_path)}")
                
                if file_size < 100:  # Fichier trop petit pour être un Excel valide
                    raise Exception(f"Fichier trop petit pour être un Excel valide: {os.path.basename(file_path)} ({file_size} bytes)")
                
                # Check file signature to validate it's actually an Excel file
                try:
                    with open(file_path, 'rb') as f:
                        header = f.read(8)
                        # Check for common Excel signatures
                        is_xlsx = header.startswith(b'PK\x03\x04')  # ZIP signature for .xlsx
                        is_xls = header.startswith(b'\xd0\xcf\x11\xe0')  # OLE2 signature for .xls
                        
                        if not (is_xlsx or is_xls):
                            print(f"⚠️ Signature fichier non reconnue: {header.hex()}")
                            print(f"⚠️ Le fichier ne semble pas être un vrai fichier Excel")
                except Exception as header_err:
                    print(f"⚠️ Impossible de vérifier l'en-tête: {header_err}")
                
                # Try with openpyxl first (for .xlsx files)
                try:
                    df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
                    print(f"✅ Succès avec openpyxl: {len(df)} lignes")
                    return df
                except Exception as e1:
                    print(f"⚠️ Échec openpyxl: {str(e1)[:100]}...")
                    
                    # Try with xlrd (for .xls files)
                    try:
                        df = pd.read_excel(file_path, sheet_name=sheet_name, engine='xlrd')
                        print(f"✅ Succès avec xlrd: {len(df)} lignes")
                        return df
                    except Exception as e2:
                        print(f"⚠️ Échec xlrd: {str(e2)[:100]}...")
                        
                        # Try without specifying engine (pandas auto-detection)
                        try:
                            df = pd.read_excel(file_path, sheet_name=sheet_name)
                            print(f"✅ Succès avec auto-détection: {len(df)} lignes")
                            return df
                        except Exception as e3:
                            print(f"⚠️ Tous les moteurs Excel ont échoué")
                            print(f"❌ ERREUR CRITIQUE: Le fichier {os.path.basename(file_path)} est corrompu ou dans un format non supporté")
                            print(f"💡 SOLUTION: Ouvrez le fichier dans Excel et resauvegardez-le au format .xlsx")
                            print(f"💡 VÉRIFICATION: Assurez-vous que le fichier n'est pas endommagé")
                            
                            raise Exception(f"FICHIER EXCEL CORROMPU: {os.path.basename(file_path)}. "
                                            f"Le fichier est endommagé ou dans un format non standard. "
                                            f"Veuillez l'ouvrir dans Excel et le resauvegarder au format .xlsx. "
                                            f"Erreurs techniques: openpyxl={str(e1)[:50]}, xlrd={str(e2)[:50]}, auto={str(e3)[:50]}")
            
            # Load all Excel files with safe reading
            df_enq = safe_read_excel(files['enq_file'])
            df_case = safe_read_excel(files['case_file'])
            df_ref = safe_read_excel(files['ref_file'])
            df_acct = safe_read_excel(files['acct_file'])
            
            print(f"✓ Enquêtes: {len(df_enq)} lignes")
            print(f"✓ Tickets: {len(df_case)} lignes")
            print(f"✓ Référence: {len(df_ref)} lignes")
            print(f"✓ Comptes: {len(df_acct)} lignes")
            
            # Rename columns to avoid conflicts
            df_enq = df_enq.rename(columns={'Créé par': 'Créé par enquête'})
            df_case = df_case.rename(columns={'Créé par': 'Créé par ticket'})
            
            # Ensure Code compte is numeric in both datasets
            df_case['Code compte'] = pd.to_numeric(df_case['Code compte'], errors='coerce')
            df_acct['Code compte'] = pd.to_numeric(df_acct['Code compte'], errors='coerce')
            
            # Merge enquetes with tickets
            df_merged = df_enq.merge(
                df_case[['Numéro', 'Créé par ticket', 'Compte', 'Code compte', 'Description courte', 'Créé le', 'Mise à jour', 'Clos par']],
                left_on='Dossier Rcbt',
                right_on='Numéro',
                how='left',
                suffixes=('_enq', '_ticket')
            )
            
            # Ensure Code compte is numeric in merged data
            df_merged['Code compte'] = pd.to_numeric(df_merged['Code compte'], errors='coerce')
            
            # Add site mapping
            if 'Login' in df_ref.columns and 'Site' in df_ref.columns:
                site_mapping = dict(zip(df_ref['Login'], df_ref['Site']))
                df_merged['Site'] = df_merged['Créé par ticket'].map(site_mapping)
                df_merged['Site'] = df_merged['Site'].fillna('Autres')
            else:
                df_merged['Site'] = 'Autres'
            
            # Classify boutiques
            def classify_boutique(code):
                if pd.isna(code):
                    return 'Autres'
                code_str = str(int(code)) if isinstance(code, (int, float)) else str(code)
                prefix = code_str[:3].zfill(3)
                if prefix == '039':
                    return 'Siège'
                elif prefix in ['400', '499', '993']:
                    return prefix
                else:
                    return 'Mini-enseigne'
            
            df_merged['Boutique_categorie'] = df_merged['Code compte'].apply(classify_boutique)
            df_case['Boutique_categorie'] = df_case['Code compte'].apply(classify_boutique)
            df_acct['Categorie'] = df_acct['Code compte'].apply(classify_boutique)
            
            print(f"✓ Fusion réussie: {len(df_merged)} lignes")
            
            # Fiabilisation optionnelle des tickets
            if enable_fiabilization:
                print("🔧 FIABILISATION ACTIVÉE")
                from .ticket_fiabilization import TicketFiabilizer
                
                fiabilizer = TicketFiabilizer(df_case, df_ref)
                
                # Analyser le potentiel de fiabilisation
                analysis = fiabilizer.analyze_fiabilization_potential()
                print(f"📊 Potentiel fiabilisation: {analysis['fiabilisation_possible']} tickets")
                
                # Fiabiliser si possible
                if analysis['fiabilisation_possible'] > 0:
                    df_case_fiabilise = fiabilizer.fiabilize_tickets()
                    
                    # Recalculer les données merged avec les tickets fiabilisés
                    df_merged_fiabilise = self._merge_data_sources(df_enq, df_case_fiabilise, df_ref, df_acct)
                    
                    # Statistiques finales
                    stats = fiabilizer.get_fiabilization_stats(df_case_fiabilise)
                    print(f"✅ Fiabilisation terminée: {stats['tickets_fiabilises']} tickets fiabilisés")
                    
                    return {
                        'merged': df_merged_fiabilise,
                        'case': df_case_fiabilise,
                        'accounts': df_acct,
                        'ref': df_ref,
                        'fiabilization_stats': stats
                    }
            
            # Calculer les métriques page 2 pour récupérer les incohérences
            page2_metrics = self.calculate_page2_enhanced_metrics({
                'merged': df_merged,
                'case': df_case,
                'accounts': df_acct,
                'reference': df_ref
            })
            
            # S'assurer que toutes les clés sont présentes dans le résultat
            result = {
                'merged': df_merged,
                'case': df_case,
                'inconsistencies': page2_metrics.get('inconsistencies', [])
            }
            
            # Ajouter accounts et ref s'ils sont valides
            if df_acct is not None and not df_acct.empty:
                result['accounts'] = df_acct
            if df_ref is not None and not df_ref.empty:
                result['ref'] = df_ref
                
            return result
            
        except Exception as e:
            print(f"❌ Erreur lors du chargement: {e}")
            raise

    def calculate_page1_metrics(self, data):
        """Calculate Page 1 synthesis metrics"""
        print("=== CALCUL PAGE 1 - SYNTHÈSE ===")
        
        df_merged = data['merged']
        df_case = data['case']
        
        # CORRECTION: Basic metrics basés sur le fichier sn_customerservice_case
        total_tickets = int(df_case['Numéro'].nunique())
        
        # CORRECTION: tickets_boutiques = nombre total tickets sans doublons du fichier case
        # Pour rapports globaux: tous les tickets du fichier case
        tickets_boutiques = int(df_case['Numéro'].nunique())  # Nombre total sans doublons
        
        # System tickets
        mask_system = df_case['Clos par'].isna() | (df_case['Clos par'] == 'system')
        tickets_system = int(df_case.loc[mask_system, 'Numéro'].nunique())
        
        # CORRECTION: Taux de clôture = enquêtes répondues / tickets total
        enquetes_repondues = int(df_merged['Dossier Rcbt'].nunique())
        taux_closure = round(enquetes_repondues / tickets_boutiques * 100, 1) if tickets_boutiques > 0 else 0.0
        
        # Satisfaction metrics
        q1_data = df_merged[df_merged['Mesure'].str.contains('Q1', na=False)]
        q1_unique = q1_data.drop_duplicates('Dossier Rcbt')
        satisfaits = int(q1_unique['Valeur de chaîne'].isin(['Très satisfaisant', 'Satisfaisant']).sum())
        total_q1 = int(len(q1_unique))
        taux_sat = round(satisfaits / total_q1 * 100, 1) if total_q1 > 0 else 0.0
        
        # CORRECTION CRITIQUE: Ajouter nb_reponses_q1 pour historique
        result = {
            'total_tickets': total_tickets,
            'tickets_boutiques': tickets_boutiques,
            'tickets_system': tickets_system,
            'nb_reponses_q1': total_q1,  # AJOUT MANQUANT pour historique correct
            'taux_closure': float(taux_closure),
            'satisfaits': satisfaits,
            'insatisfaits': total_q1 - satisfaits,
            'taux_sat': float(taux_sat),
            'closure_ok': bool(taux_closure >= self.objectifs['closure_rate']),
            'sat_ok': bool(taux_sat >= self.objectifs['satisfaction_q1'])
        }
        
        print(f"✓ Taux de clôture: {taux_closure}% ({'✅' if result['closure_ok'] else '⚠️'})")
        print(f"✓ Taux satisfaction Q1: {taux_sat}% ({'✅' if result['sat_ok'] else '⚠️'})")
        
        return result

    def calculate_page2_enhanced_metrics(self, data):
        """Calculate enhanced Page 2 metrics with detailed analysis"""
        print("=== CALCUL PAGE 2 AMÉLIORÉ - COMMENTAIRES ===")
        
        df_merged = data['merged']
        
        # Q1 data with comments (for page 2)
        q1_data_with_comments = df_merged[
            df_merged['Mesure'].str.contains('Q1', na=False) &
            df_merged['Commentaire'].notna() & 
            (df_merged['Commentaire'].str.strip() != '')
        ]
        
        # CORRECTION: All Q1 responses for percentage calculation - utiliser TOUTES les données Q1
        q1_all = df_merged[df_merged['Mesure'].str.contains('Q1', na=False)]
        # On garde la même logique que calculate_basic_metrics pour la cohérence
        q1_unique = q1_all.drop_duplicates('Dossier Rcbt')
        total_q1_responses = int(len(q1_unique))
        total_with_comments = int(len(q1_data_with_comments.drop_duplicates('Dossier Rcbt')))
        
        # Comments percentage
        comments_percentage = round(total_with_comments / total_q1_responses * 100, 1) if total_q1_responses > 0 else 0.0
        
        # Detailed analysis by collaborator
        collaborator_analysis = self._analyze_by_collaborator(q1_data_with_comments)
        
        # Site analysis
        site_analysis = self._analyze_by_site(q1_data_with_comments)
        
        # Satisfaction distribution in comments
        satisfaction_distribution = self._analyze_satisfaction_distribution(q1_data_with_comments)
        
        # Detect inconsistencies (positive words in negative ratings)
        inconsistencies = self._detect_inconsistencies(q1_data_with_comments)
        
        result = {
            'total_q1_responses': total_q1_responses,
            'total_with_comments': total_with_comments,
            'comments_percentage': float(comments_percentage),
            'total_collaborators': int(len(collaborator_analysis)),
            'collaborator_analysis': collaborator_analysis,
            'site_analysis': site_analysis,
            'satisfaction_distribution': satisfaction_distribution,
            'inconsistencies': inconsistencies
        }
        
        print(f"✓ Pourcentage enquêtes avec commentaires: {comments_percentage}%")
        print(f"✓ Collaborateurs analysés: {len(collaborator_analysis)}")
        print(f"✓ Sites analysés: {len(site_analysis)}")
        print(f"✓ Incohérences détectées: {len(inconsistencies)}")
        
        return result

    def _analyze_by_collaborator(self, df_comments):
        """Analyze comments by collaborator"""
        if df_comments.empty:
            return []
        
        # Group by collaborator
        collaborator_stats = []
        
        for collaborator, group in df_comments.groupby('Créé par ticket'):
            if pd.isna(collaborator):
                continue
                
            total_comments = int(len(group))
            
            # Satisfaction breakdown
            satisfait_count = int(group['Valeur de chaîne'].isin(['Très satisfaisant', 'Satisfaisant']).sum())
            insatisfait_count = int(group['Valeur de chaîne'].isin(['Peu satisfaisant', 'Très peu satisfaisant']).sum())
            
            # Percentages
            satisfait_pct = round(satisfait_count / total_comments * 100, 1) if total_comments > 0 else 0.0
            insatisfait_pct = round(insatisfait_count / total_comments * 100, 1) if total_comments > 0 else 0.0
            
            # Site information
            site = group['Site'].iloc[0] if 'Site' in group.columns else 'Non défini'
            
            collaborator_stats.append({
                'collaborator': str(collaborator),
                'site': str(site),
                'total_comments': total_comments,
                'satisfait_count': satisfait_count,
                'insatisfait_count': insatisfait_count,
                'satisfait_pct': float(satisfait_pct),
                'insatisfait_pct': float(insatisfait_pct)
            })
        
        # Sort by total comments descending
        return sorted(collaborator_stats, key=lambda x: x['total_comments'], reverse=True)

    def _analyze_by_site(self, df_comments):
        """Analyze comments by site"""
        if df_comments.empty:
            return []
        
        site_stats = []
        
        for site, group in df_comments.groupby('Site'):
            if pd.isna(site):
                continue
                
            total_comments = int(len(group))
            unique_collaborators = int(group['Créé par ticket'].nunique())
            
            # Satisfaction breakdown
            satisfait_count = int(group['Valeur de chaîne'].isin(['Très satisfaisant', 'Satisfaisant']).sum())
            insatisfait_count = int(group['Valeur de chaîne'].isin(['Peu satisfaisant', 'Très peu satisfaisant']).sum())
            
            # Percentages
            satisfait_pct = round(satisfait_count / total_comments * 100, 1) if total_comments > 0 else 0.0
            
            site_stats.append({
                'site': str(site),
                'total_comments': total_comments,
                'unique_collaborators': unique_collaborators,
                'satisfait_count': satisfait_count,
                'insatisfait_count': insatisfait_count,
                'satisfait_pct': float(satisfait_pct)
            })
        
        # Sort by total comments descending
        return sorted(site_stats, key=lambda x: x['total_comments'], reverse=True)

    def _analyze_satisfaction_distribution(self, df_comments):
        """CORRECTION MAJEURE: Analyze satisfaction distribution - HARMONISÉ AVEC TABLEAU"""
        if df_comments.empty:
            return {}
        
        # CORRECTION INCOHÉRENCE: Utiliser les mêmes données que le tableau (unique par dossier)
        df_satisfaction = df_comments[df_comments['Valeur de chaîne'].notna()]
        
        if df_satisfaction.empty:
            return {}
        
        # CORRECTION CRITQUE: Prendre données uniques par Dossier RCBT comme le tableau
        df_unique = df_satisfaction.drop_duplicates('Dossier Rcbt')
        
        # Count by satisfaction level sur données harmonisées
        satisfaction_counts = df_unique['Valeur de chaîne'].value_counts()
        total = int(len(df_unique))  # CORRECTION: Total = entrées uniques par dossier
        
        print(f"🔍 SATISFACTION DISTRIBUTION HARMONISÉE:")
        print(f"   Total UNIQUE par dossier: {total}")
        
        distribution = {}
        for satisfaction, count in satisfaction_counts.items():
            count = int(count)
            percentage = round((count / total) * 100, 1) if total > 0 else 0.0
            print(f"   {satisfaction}: {count}/{total} = {percentage}%")
            distribution[str(satisfaction)] = {
                'count': count,
                'percentage': percentage
            }
        
        return distribution

    def _detect_inconsistencies(self, df_comments):
        """Detect inconsistencies between ratings and comments - bidirectional"""
        if df_comments.empty:
            return []
        
        # MOTS POSITIFS (forts et contextuels)
        mots_positifs = ['merci', 'parfait', 'top', 'efficace', 'rapide', 'solution', 
                        'clair', 'précis', 'excellent', 'super', 'formidable', 'génial',
                        'impeccable', 'extraordinaire', 'fantastique', 'remarquable',
                        'bien accueilli', 'bien acceuilli', 'professionnel', 'compétent', 'aimable', 'courtois']
        
        # MOTS NÉGATIFS (forts et contextuels) - ÉTENDUS POUR MITIGÉ
        # ATTENTION: "problème/probleme" sera traité contextuellement ci-dessous
        mots_negatifs = ['catastrophique', 'horrible', 'nul', 'inadmissible', 'scandaleux',
                        'inacceptable', 'décevant', 'frustrant', 'agacé', 'énervé', 'furieux',
                        'incompétent', 'déplorable', 'lamentable', 'désastreux', 'déçu', 'déçu par',
                        'insatisfait', 'mécontent', 'erreur', 'mauvais', 'qualité du traitement',
                        'inefficace', 'lent', 'pas bien', 'pas bon', 'pas résolu', 'non résolu',
                        'pas satisfait', 'pas content', 'raté', 'échec', 'faute', 'manque',
                        'insuffisant', 'incomplet', 'inadéquat', 'attendais mieux', 'j\'attendais mieux',
                        'attendre mieux', 'pas parfait', 'n\'était pas parfait', 'ce n\'était pas parfait',
                        'bloquant', 'perdre', 'business', 'fait perdre', 'perte']
        
        # DÉTECTION CONTEXTUELLE DE "PROBLÈME/PROBLEME"
        def is_probleme_negative_context(comment_text):
            """Détermine si 'problème' est dans un contexte négatif ou positif"""
            comment_lower = comment_text.lower()
            
            # Contextes POSITIFS (problème résolu/traité positivement)
            positive_patterns = [
                r'r[ée]solu.{0,10}probl[eè]me?s?', r'probl[eè]me?s?.{0,10}r[ée]solu',
                r'r[ée]gl[ée].{0,10}probl[eè]me?s?', r'probl[eè]me?s?.{0,10}r[ée]gl[ée]', 
                r'solutionn.{0,10}probl[eè]me?s?', r'probl[eè]me?s?.{0,10}solutionn',
                r'solution.{0,10}probl[eè]me?s?', r'probl[eè]me?s?.{0,10}solution',
                r'pas de probl[eè]me?s?', r'aucun probl[eè]me?s?', r'zero probl[eè]me?s?',
                r'sans probl[eè]me?s?', r'r[ée]solution.{0,15}probl[eè]me?s?',
                r'a\s+r[ée]solu\s+.{0,15}probl[eè]me?s?', r'ont\s+r[ée]solu\s+.{0,15}probl[eè]me?s?'
            ]
            
            # Contextes NÉGATIFS (problème persistant/non résolu)  
            negative_patterns = [
                r'gros probl[eè]me', r'grave probl[eè]me', r'sérieux probl[eè]me',
                r'probl[eè]me bloquant', r'probl[eè]me majeur', r'probl[eè]me important',
                r'probl[eè]me non résolu', r'probl[eè]me pas résolu'
            ]
            
            import re
            
            # Vérifier contextes positifs (prioritaire)
            for pattern in positive_patterns:
                if re.search(pattern, comment_lower):
                    return False  # Pas négatif, contexte positif
                    
            # Vérifier contextes négatifs explicites
            for pattern in negative_patterns:
                if re.search(pattern, comment_lower):
                    return True  # Négatif explicite
                    
            # Par défaut, si "problème" seul → considérer comme négatif
            if re.search(r'\bprobl[eè]me?s?\b', comment_lower):
                return True
                
            return False
        
        # MOTS NÉGATIFS CONTEXTUELS (nécessitent une vérification du contexte)
        mots_negatifs_contextuels = {
            'long': ['trop long', 'très long', 'delai long', 'délai long', 'temps long', 'attente long']
        }
        
        # PATTERNS TEMPORELS NÉGATIFS (détection d'attente excessive)
        import re
        patterns_temporels_negatifs = [
            r'\d+\s*mn?\s+d[\'e\s]*attente',     # "18 mn d attente", "15 min d'attente" 
            r'\d+\s*mn?\s+d[\'e\s]*echange',     # "18 mn d echange", "20 min d'échange"
            r'\d+\s*minutes?\s+pour',            # "15 minutes pour"
            r'\d+\s*mn?\s+avant',                # "10 mn avant"
            r'attendre?\s+\d+\s*mn?',            # "attendre 15 mn"
            r'patiente[rz]?\s+\d+\s*mn?'         # "patienter 20 mn"
        ]
        
        # MOTS DE CONTRASTE (détectent les nuances/ambiguïtés) - ÉTENDUS
        mots_contraste = ['mais', 'cependant', 'toutefois', 'néanmoins', 'par contre', 
                         'en revanche', 'malheureusement', 'dommage', 'regret',
                         'même si', 'bien que', 'malgré', 'pourtant',
                         'seulement', 'uniquement', 'juste', 'sauf que']
        
        # MOTS DE CONTRASTE SPÉCIAUX (nécessitent des limites de mots)
        mots_contraste_speciaux = [' or ', 'or,', 'or.', 'or!', 'or?', 'or;']
        
        inconsistencies = []
        
        # TYPE 1: Notes négatives avec commentaires positifs
        negative_ratings = df_comments[
            df_comments['Valeur de chaîne'].isin(['Peu satisfaisant', 'Très peu satisfaisant'])
        ]
        
        for _, row in negative_ratings.iterrows():
            comment = str(row['Commentaire']).lower()
            positive_words_found = [word for word in mots_positifs if word in comment]
            
            # FIABILISATION: Au moins 2 mots positifs ou 1 mot très fort
            mots_forts = ['parfait', 'excellent', 'formidable', 'impeccable', 'extraordinaire']
            mots_forts_trouves = [word for word in mots_forts if word in comment]
            
            if len(positive_words_found) >= 2 or len(mots_forts_trouves) >= 1:
                # CORRECTION PROBLÈME 1: Calculer la suggestion et vérifier si elle diffère
                original_rating = str(row['Valeur de chaîne'])
                suggested_rating = self._get_suggested_rating(original_rating, 'Note négative / Commentaire positif')
                
                # FILTRAGE INTELLIGENT: Ne remonter que si suggestion différente de l'original
                print(f"🔍 TYPE 1 - Dossier {row['Dossier Rcbt']}: '{original_rating}' → '{suggested_rating}' ({'FILTRÉ' if original_rating == suggested_rating else 'REMONTÉ'})")
                if original_rating != suggested_rating:
                    inconsistencies.append({
                        'dossier': str(row['Dossier Rcbt']),
                        'collaborator': str(row['Créé par ticket']),
                        'rating': original_rating,
                        'comment': str(row['Commentaire']),
                        'inconsistency_type': 'Note négative / Commentaire positif',
                        'detected_words': positive_words_found,
                        'suggested_rating': suggested_rating,
                        'impact': 'Ticket classé comme incohérent - Vérification recommandée'
                    })
        
        # TYPE 2: Notes positives avec commentaires négatifs  
        positive_ratings = df_comments[
            df_comments['Valeur de chaîne'].isin(['Très satisfaisant', 'Satisfaisant'])
        ]
        
        for _, row in positive_ratings.iterrows():
            comment = str(row['Commentaire']).lower()
            comment_original = str(row['Commentaire'])
            negative_words_found = [word for word in mots_negatifs if word in comment]
            contraste_words_found = [word for word in mots_contraste if word in comment]
            
            # DÉTECTION CONTEXTUELLE DE "PROBLÈME/PROBLEME"
            if is_probleme_negative_context(comment_original):
                if "probleme_negatif" not in negative_words_found:
                    negative_words_found.append("probleme_negatif")
            
            # CORRECTION: Vérifier les mots de contraste spéciaux avec limites
            contraste_speciaux_found = [word.strip() for word in mots_contraste_speciaux if word in comment]
            contraste_words_found.extend(contraste_speciaux_found)
            
            # VÉRIFICATION CONTEXTUELLE pour les mots négatifs contextuels
            for word, contexts in mots_negatifs_contextuels.items():
                if any(context in comment for context in contexts):
                    if word not in negative_words_found:
                        negative_words_found.append(word)
            
            # VÉRIFICATION PATTERNS TEMPORELS NÉGATIFS (attente excessive)
            for pattern in patterns_temporels_negatifs:
                if re.search(pattern, comment):
                    if "delai_excessif" not in negative_words_found:
                        negative_words_found.append("delai_excessif")
            
            # FIABILISATION: Au moins 1 mot négatif fort OU mot de contraste avec positif+négatif
            detected_words = negative_words_found + contraste_words_found
            
            # Cas 1: Mots négatifs directs
            # Cas 2: Commentaires mitigés (positif + contraste + négatif)
            # Cas 3: NOUVEAU - Commentaires nuancés avec contraste même sans mots négatifs forts
            has_positive = any(word in comment for word in mots_positifs)
            has_contraste = len(contraste_words_found) >= 1
            has_negative = len(negative_words_found) >= 1
            
            # AMÉLIORATION SPÉCIALE: Détecter expressions nuancées
            has_nuanced_negative = any(phrase in comment for phrase in [
                'attendais mieux', 'j\'attendais mieux', 'attendre mieux',
                'pas parfait', 'n\'était pas parfait', 'ce n\'était pas parfait'
            ])
            
            # LOGIQUE AMÉLIORÉE: Détecter "bien accueilli mais déçu" et variations
            # Inclure les expressions nuancées dans les mots détectés
            nuanced_words_found = [phrase for phrase in [
                'attendais mieux', 'j\'attendais mieux', 'attendre mieux',
                'pas parfait', 'n\'était pas parfait', 'ce n\'était pas parfait'
            ] if phrase in comment]
            
            if has_nuanced_negative:
                detected_words.extend(nuanced_words_found)
            
            if has_negative or (has_positive and has_contraste) or (has_contraste and has_nuanced_negative):
                inconsistency_type = 'Note positive / Commentaire mitigé' if has_contraste else 'Note positive / Commentaire négatif'
                
                # CORRECTION PROBLÈME 1: Calculer la suggestion et vérifier si elle diffère
                original_rating = str(row['Valeur de chaîne'])
                suggested_rating = self._get_suggested_rating(original_rating, inconsistency_type)
                
                # FILTRAGE INTELLIGENT: Ne remonter que si suggestion différente de l'original
                print(f"🔍 TYPE 2 - Dossier {row['Dossier Rcbt']}: '{original_rating}' → '{suggested_rating}' ({'FILTRÉ' if original_rating == suggested_rating else 'REMONTÉ'})")
                if original_rating != suggested_rating:
                    inconsistencies.append({
                        'dossier': str(row['Dossier Rcbt']),
                        'collaborator': str(row['Créé par ticket']),
                        'rating': original_rating,
                        'comment': str(row['Commentaire']),
                        'inconsistency_type': inconsistency_type,
                        'detected_words': detected_words,
                        'suggested_rating': suggested_rating,
                        'impact': 'Ticket classé comme incohérent - Vérification recommandée'
                    })
        
        return inconsistencies
    
    def _get_suggested_rating(self, original_rating: str, inconsistency_type: str) -> str:
        """CORRECTION PROBLÈME 1: Suggère une correction basée sur le type d'incohérence"""
        if inconsistency_type == 'Note négative / Commentaire positif':
            if original_rating in ['Très peu satisfaisant', 'Peu satisfaisant']:
                return 'Satisfaisant'
        elif inconsistency_type in ['Note positive / Commentaire négatif', 'Note positive / Commentaire mitigé']:
            # LOGIQUE AJUSTÉE: Pour commentaires mitigés, suggérer une diminution appropriée
            if inconsistency_type == 'Note positive / Commentaire mitigé':
                # Pour les commentaires mitigés (ex: "problème mais solution"), diminution modérée
                if original_rating == 'Très satisfaisant':
                    return 'Satisfaisant'  # Diminution d'un niveau pour nuance
                elif original_rating == 'Satisfaisant':
                    return 'Peu satisfaisant'
            elif original_rating in ['Très satisfaisant', 'Satisfaisant']:
                # Pour commentaires purement négatifs, diminution plus forte
                return 'Peu satisfaisant'
        return original_rating
    
    def get_page2_data(self, data):
        """Get detailed data for page 2 (comments)"""
        df_merged = data['merged']
        
        # Q1 data with comments
        q1_data_with_comments = df_merged[
            df_merged['Mesure'].str.contains('Q1', na=False) &
            df_merged['Commentaire'].notna() & 
            (df_merged['Commentaire'].str.strip() != '')
        ]
        
        return q1_data_with_comments
    
    def get_page3_data(self, data):
        """Get data for page 3 (unsatisfied without comments)"""
        df_merged = data['merged']
        
        # Clients insatisfaits sans commentaire
        insatisfied_no_comment = df_merged[
            df_merged['Mesure'].str.contains('Q1', na=False) &
            df_merged['Valeur de chaîne'].isin(['Peu satisfaisant', 'Très peu satisfaisant']) &
            (df_merged['Commentaire'].isna() | (df_merged['Commentaire'].str.strip() == ''))
        ]
        
        return insatisfied_no_comment
