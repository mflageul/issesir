import os
import base64
import tempfile
from datetime import datetime
from pathlib import Path
import pandas as pd

class OptimizedReportGenerator:
    def __init__(self):
        self.template = self._get_optimized_template()
        
    def generate_optimized_report(self, data):
        """Generate optimized HTML report with all requested features"""
        print("=== GÉNÉRATION DU RAPPORT OPTIMISÉ RCBT ===")
        
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        report_filename = f'rapport_rcbt_optimized_{timestamp}.html'
        report_path = os.path.join(tempfile.gettempdir(), report_filename)
        
        # Get logo
        logo_base64 = self._get_logo_base64()
        
        # Import processor for calculations
        from .data_processor import DataProcessor
        processor = DataProcessor()
        
        # Calculate metrics with debugging
        page1_metrics = processor.calculate_page1_metrics(data)
        print(f"🔍 Metrics calculées: {page1_metrics}")
        
        # Generate all pages data with proper Q1 filtering (no duplicates)
        page1_data = self._get_page1_synthesis(data, page1_metrics)
        page2_data = self._get_page2_comments_enhanced(data)
        page3_data = self._get_page3_unsatisfied_no_comments(data)
        page4_data = self._get_page4_global_analysis_colored(data)
        page5_data = self._get_page5_responding_shops_colored(data)
        page6_data = self._get_page6_never_responding(data)
        page7_data = self._get_page7_shop_ranking(data)
        
        # Load and encode Q&T logo for report header
        import base64
        try:
            with open('static/images/logo_qt.jpg', 'rb') as f:
                logo_base64 = base64.b64encode(f.read()).decode('utf-8')
                print(f"✅ Logo Q&T chargé pour le rapport (taille: {len(logo_base64)} caractères)")
        except Exception as e:
            print(f"❌ Erreur chargement logo: {e}")
            logo_base64 = ""
        
        # Load and encode Centre de Services logo
        try:
            with open('static/images/logo_centre_services.png', 'rb') as f:
                logo_centre_base64 = base64.b64encode(f.read()).decode('utf-8')
                print(f"✅ Logo Centre de Services chargé (taille: {len(logo_centre_base64)} caractères)")
        except Exception as e:
            print(f"❌ Erreur chargement logo centre: {e}")
            logo_centre_base64 = ""
        
        html_content = self.template.format(
            timestamp=datetime.now().strftime('%d/%m/%Y à %H:%M'),
            logo_base64=logo_base64,
            logo_centre_base64=logo_centre_base64,
            **page1_data,
            **page2_data,
            **page3_data,
            **page4_data,
            **page5_data,
            **page6_data,
            **page7_data
        )
        
        # Save report
        with open(report_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        print(f"✅ Rapport optimisé généré: {report_path}")
        return report_path
    
    def _get_logo_base64(self):
        """Get logo as base64"""
        try:
            logo_path = Path(__file__).parent.parent / 'static' / 'images' / 'logo.png'
            if logo_path.exists():
                with open(logo_path, 'rb') as f:
                    return base64.b64encode(f.read()).decode('utf-8')
        except Exception as e:
            print(f"⚠️ Logo non trouvé: {e}")
        return ""
    
    def _get_page1_synthesis(self, data, metrics):
        """Page 1 - Synthèse avec taux corrects"""
        df_merged = data['merged']
        
        # Force recalculation to ensure correct values
        q1_data = df_merged[df_merged['Mesure'].str.contains('Q1', na=False)]
        q1_unique = q1_data.drop_duplicates('Dossier Rcbt')  # Remove duplicates per RCBT folder
        
        if len(q1_unique) > 0:
            satisf_responses = q1_unique['Valeur de chaîne'].isin(['Très satisfaisant', 'Satisfaisant']).sum()
            total_responses = len(q1_unique)
            taux_satisfaction = (satisf_responses / total_responses * 100) if total_responses > 0 else 0.0
        else:
            satisf_responses = 0
            total_responses = 0
            taux_satisfaction = 0.0
            
        # CORRECTION: Closure rate calculation basé sur fichier case
        df_case = data['case']
        total_tickets = df_case['Numéro'].nunique()  # Total tickets sans doublons du fichier case
        tickets_boutiques = total_tickets  # Pour rapport global = tous les tickets du fichier case
        enquetes_repondues = q1_unique['Dossier Rcbt'].nunique()  # Enquêtes ayant répondu
        taux_closure = (enquetes_repondues / tickets_boutiques * 100) if tickets_boutiques > 0 else 0.0
        
        print(f"✅ Taux recalculés - Clôture: {taux_closure:.1f}%, Satisfaction: {taux_satisfaction:.1f}%")
        
        # Generate individual pie charts for satisfaction/dissatisfaction by site
        df_merged = data['merged']
        q1_data = df_merged[df_merged['Mesure'].str.contains('Q1', na=False)]
        q1_unique = q1_data.drop_duplicates('Dossier Rcbt')
        
        # Individual pie charts per site (satisfaction vs dissatisfaction)
        site_pie_charts = ""
        sites = q1_unique['Site'].unique()
        
        for site in sites:
            site_data = q1_unique[q1_unique['Site'] == site]
            satisf_count = site_data['Valeur de chaîne'].isin(['Très satisfaisant', 'Satisfaisant']).sum()
            total_count = len(site_data)
            dissatisf_count = total_count - satisf_count
            
            if total_count > 0:
                satisf_pct = (satisf_count / total_count * 100)
                dissatisf_pct = (dissatisf_count / total_count * 100)
                
                # Create pie chart data
                pie_data = pd.Series([satisf_pct, dissatisf_pct], 
                                   index=['Satisfait', 'Insatisfait'])
                
                site_pie_chart = self._generate_pie_chart_base64(
                    pie_data,
                    title=f"Site {site} (n={total_count})",
                    colors=['#28a745', '#dc3545']
                )
                
                site_pie_charts += f'<div class="site-pie-chart"><img src="data:image/png;base64,{site_pie_chart}" alt="Satisfaction site {site}"></div>'
        
        # Satisfaction by measure type (from column "Mesure")
        ticket_type_charts = ""
        if 'Mesure' in q1_unique.columns:
            for mesure_type in q1_unique['Mesure'].unique():
                if pd.notna(mesure_type):
                    type_data = q1_unique[q1_unique['Mesure'] == mesure_type]
                    satisf_count = type_data['Valeur de chaîne'].isin(['Très satisfaisant', 'Satisfaisant']).sum()
                    total_count = len(type_data)
                    dissatisf_count = total_count - satisf_count
                    
                    if total_count > 0:
                        satisf_pct = (satisf_count / total_count * 100)
                        dissatisf_pct = (dissatisf_count / total_count * 100)
                        
                        pie_data = pd.Series([satisf_pct, dissatisf_pct], 
                                           index=['Satisfait', 'Insatisfait'])
                        
                        type_pie_chart = self._generate_pie_chart_base64(
                            pie_data,
                            title=f"{mesure_type} (n={total_count})",
                            colors=['#28a745', '#dc3545']
                        )
                        
                        ticket_type_charts += f'<div class="measure-pie-chart"><img src="data:image/png;base64,{type_pie_chart}" alt="Satisfaction {mesure_type}"></div>'
        else:
            ticket_type_charts = '<div class="info-box"><p>⚠️ Colonne "Mesure" non trouvée dans les données.</p></div>'
        
        return {
            'page1_taux_closure': f"{taux_closure:.1f}%",
            'page1_taux_satisfaction': f"{taux_satisfaction:.1f}%",
            'page1_closure_color': 'success' if taux_closure >= 13.0 else 'warning',
            'page1_satisfaction_color': 'success' if taux_satisfaction >= 92.0 else 'warning',
            'page1_closure_icon': '✅' if taux_closure >= 13.0 else '⚠️',
            'page1_satisfaction_icon': '✅' if taux_satisfaction >= 92.0 else '⚠️',
            'page1_tickets_boutiques': f"{tickets_boutiques:,}",
            'page1_total_tickets': f"{total_tickets:,}",
            'page1_satisfaits': f"{satisf_responses:,}",
            'page1_insatisfaits': f"{total_responses - satisf_responses:,}",
            'page1_total_q1': f"{total_responses:,}",
            'page1_site_pie_charts': site_pie_charts,
            'page1_ticket_type_charts': ticket_type_charts
        }
    
    def _get_page2_comments_enhanced(self, data):
        """Page 2 - Commentaires avec analyse améliorée et graphiques"""
        df_merged = data['merged']
        
        # Q1 unique data only (no duplicates per RCBT folder)
        q1_data = df_merged[df_merged['Mesure'].str.contains('Q1', na=False)]
        q1_unique = q1_data.drop_duplicates('Dossier Rcbt')
        
        # Comments analysis
        with_comments = q1_unique[q1_unique['Commentaire'].notna() & (q1_unique['Commentaire'].str.strip() != '')]
        total_q1_responses = len(q1_unique)
        total_with_comments = len(with_comments)
        comments_percentage = (total_with_comments / total_q1_responses * 100) if total_q1_responses > 0 else 0.0
        
        # Use the correct date columns as confirmed by user:
        # 'Créé le' = creation date
        # Date of survey response = closure date (from ticket data)
        
        # Build aggregation dict with correct columns
        agg_dict = {
            'Dossier Rcbt': 'count',
            'Valeur de chaîne': lambda x: (x.isin(['Très satisfaisant', 'Satisfaisant'])).sum(),
            'Site': 'first'  # Get site for each collaborator
        }
        
        # Add the confirmed date columns
        if 'Créé le' in with_comments.columns:
            agg_dict['Créé le'] = 'first'
        
        # Look for closure date column (various possible names)
        closure_date_col = None
        for col in with_comments.columns:
            if any(term in col.lower() for term in ['clos', 'closure', 'fermé', 'closed']):
                agg_dict[col] = 'first'
                closure_date_col = col
                break
        
        # Detailed analysis by collaborator with site info
        collaborator_analysis = with_comments.groupby('Créé par ticket').agg(agg_dict).reset_index()
        
        # Build column names - keep it simple and predictable
        new_columns = ['Collaborateur', 'Total_commentaires', 'Commentaires_satisfaits', 'Site']
        
        # Add date columns in the order they were added to agg_dict
        if 'Créé le' in agg_dict:
            new_columns.append('Date_creation')
        
        if closure_date_col:
            new_columns.append('Date_cloture')
        
        collaborator_analysis.columns = new_columns
        
        # Convert to numeric BEFORE any calculations to prevent dtype errors
        collaborator_analysis['Commentaires_satisfaits'] = pd.to_numeric(collaborator_analysis['Commentaires_satisfaits'], errors='coerce').fillna(0)
        collaborator_analysis['Total_commentaires'] = pd.to_numeric(collaborator_analysis['Total_commentaires'], errors='coerce').fillna(0)
        
        collaborator_analysis['Commentaires_insatisfaits'] = collaborator_analysis['Total_commentaires'] - collaborator_analysis['Commentaires_satisfaits']
        
        # Calculate percentage with safe division
        collaborator_analysis['Pourcentage_satisfait'] = 0.0
        mask = collaborator_analysis['Total_commentaires'] > 0
        collaborator_analysis.loc[mask, 'Pourcentage_satisfait'] = ((collaborator_analysis.loc[mask, 'Commentaires_satisfaits'] / collaborator_analysis.loc[mask, 'Total_commentaires']) * 100).round(1)
        
        # Add satisfaction color class
        def get_satisfaction_color_class(pct):
            if pct >= 92:
                return 'success'
            elif pct >= 80:
                return 'warning'
            else:
                return 'danger'
        
        collaborator_analysis['satisfaction_class'] = collaborator_analysis['Pourcentage_satisfait'].apply(get_satisfaction_color_class)
        
        # Site analysis with comment percentage
        all_tickets_by_site = q1_unique.groupby('Site').size().reset_index(name='Total_enquetes')
        site_comments = with_comments.groupby('Site').agg({
            'Dossier Rcbt': 'count',
            'Valeur de chaîne': lambda x: (x.isin(['Très satisfaisant', 'Satisfaisant'])).sum()
        }).reset_index()
        
        # Rename columns properly first
        site_comments.columns = ['Site', 'Total_commentaires', 'Commentaires_satisfaits']
        site_comments['Commentaires_insatisfaits'] = site_comments['Total_commentaires'] - site_comments['Commentaires_satisfaits']
        
        # Merge to get comment percentage per site
        site_analysis = all_tickets_by_site.merge(site_comments, on='Site', how='left').fillna(0).infer_objects(copy=False)
        
        # Convert to numeric BEFORE any calculations to prevent dtype errors
        site_analysis['Total_commentaires'] = pd.to_numeric(site_analysis['Total_commentaires'], errors='coerce').fillna(0)
        site_analysis['Total_enquetes'] = pd.to_numeric(site_analysis['Total_enquetes'], errors='coerce').fillna(0)
        site_analysis['Commentaires_satisfaits'] = pd.to_numeric(site_analysis['Commentaires_satisfaits'], errors='coerce').fillna(0)
        
        # Calculate percentages with safe division
        site_analysis['Pourcentage_commentaires'] = 0.0
        mask1 = site_analysis['Total_enquetes'] > 0
        site_analysis.loc[mask1, 'Pourcentage_commentaires'] = ((site_analysis.loc[mask1, 'Total_commentaires'] / site_analysis.loc[mask1, 'Total_enquetes']) * 100).round(1)
        
        # Add satisfaction percentage with comments
        site_analysis['Pourcentage_satisfaction_commentaires'] = 0.0
        mask2 = site_analysis['Total_commentaires'] > 0
        site_analysis.loc[mask2, 'Pourcentage_satisfaction_commentaires'] = ((site_analysis.loc[mask2, 'Commentaires_satisfaits'] / site_analysis.loc[mask2, 'Total_commentaires']) * 100).round(1)
        site_analysis['satisfaction_class'] = site_analysis['Pourcentage_satisfaction_commentaires'].apply(get_satisfaction_color_class)
        
        # Build enhanced tables with account codes, shop names, measure and closed by
        comments_table = ""
        # CORRECTION CRITIQUE: Afficher TOUTES les données pour cohérence avec camembert
        for _, row in with_comments.iterrows():
            # Get account code without decimals and shop name
            code_compte = int(float(row.get('Code compte', 0))) if pd.notna(row.get('Code compte')) else 'N/A'
            
            # Get shop name from available columns
            nom_boutique = 'N/A'
            for name_col in ['Compte', 'Account', 'Nom']:
                if name_col in row.index and pd.notna(row.get(name_col)) and str(row.get(name_col)).strip():
                    nom_boutique = row.get(name_col)
                    break
            
            q1_class = self._get_q1_satisfaction_class(row.get('Valeur de chaîne', ''))
            
            # NOUVEAU: Vérifier les incohérences traitées avec traçabilité
            dossier = row.get('Dossier Rcbt', '')
            inconsistent_badge = ''
            
            # Vérifier si cette ligne a été traitée pour incohérence
            if 'Validation_Applied' in row.index and row.get('Validation_Applied') == True:
                original_rating = row.get('Original_Rating', 'N/A')
                new_rating = row.get('Valeur de chaîne', 'N/A')
                validation_reason = row.get('Validation_Reason', 'Correction appliquée')
                
                # Vérifier si c'est une validation "conservée" (original = validé)
                is_conserved = (original_rating == new_rating) or any(pattern in validation_reason.lower() for pattern in [
                    'conserver l\'original', 'conserver original', 'conservé à l\'original', 
                    'original conservé', 'garder original', 'maintenir original', 'conservé', 'conserver'
                ])
                if is_conserved:
                    tooltip_text = f"INCOHÉRENCE CONTRÔLÉE: {original_rating} - Raison: {validation_reason}"
                    inconsistent_badge = f' <span class="badge bg-success text-white ms-1" style="cursor: help;" title="{tooltip_text}" data-bs-toggle="tooltip">🔎 CONTRÔLÉ</span>'
                else:
                    tooltip_text = f"INCOHÉRENCE CORRIGÉE: {original_rating} → {new_rating} - Raison: {validation_reason}"
                    inconsistent_badge = f' <span class="badge bg-info text-white ms-1" style="cursor: help;" title="{tooltip_text}" data-bs-toggle="tooltip">⚠️ CORRIGÉ</span>'
            else:
                # Détecter les incohérences non traitées
                is_inconsistent = self._is_comment_inconsistent(row.get('Valeur de chaîne', ''), row.get('Commentaire', ''))
                if is_inconsistent:
                    inconsistency_details = self._get_inconsistency_details(row.get('Valeur de chaîne', ''), row.get('Commentaire', ''))
                    tooltip_text = f"INCOHÉRENCE: {inconsistency_details['type']} - Mots: {', '.join(inconsistency_details['words'][:3])} - {inconsistency_details['impact']}"
                    inconsistent_badge = f' <span class="badge bg-warning text-dark ms-1" style="cursor: help;" title="{tooltip_text}" data-bs-toggle="tooltip">⚠️ INCOHÉRENCE</span>'
            
            # Get ticket number and dates
            numero_ticket = row.get('Dossier Rcbt', 'N/A')
            date_creation = 'N/A'
            date_cloture = 'N/A'
            
            if 'Créé le' in row.index and pd.notna(row['Créé le']):
                try:
                    if hasattr(row['Créé le'], 'strftime'):
                        date_creation = row['Créé le'].strftime('%d/%m/%Y')
                    else:
                        date_creation = str(row['Créé le'])[:10]
                except:
                    date_creation = str(row['Créé le'])[:10] if row['Créé le'] else 'N/A'
            
            # Look for closure date column
            for col in ['Mise à jour', 'Clos le', 'Date cloture']:
                if col in row.index and pd.notna(row[col]):
                    try:
                        if hasattr(row[col], 'strftime'):
                            date_cloture = row[col].strftime('%d/%m/%Y')
                        else:
                            date_cloture = str(row[col])[:10]
                        break
                    except:
                        continue
            
            # Ajouter classe CSS pour lignes incohérentes
            row_class = " class='inconsistent-row'" if is_inconsistent else ""
            
            comments_table += f"""
            <tr{row_class}>
                <td>{code_compte}</td>
                <td>{nom_boutique}</td>
                <td>{row.get('Site', 'N/A')}</td>
                <td>{row.get('Créé par ticket', 'N/A')}</td>
                <td>{row.get('Mesure', 'N/A')}</td>
                <td>{numero_ticket}</td>
                <td>{row.get('Clos par', 'N/A')}</td>
                <td class="{q1_class}">{row.get('Valeur de chaîne', 'N/A')}{inconsistent_badge}</td>
                <td>{date_creation}</td>
                <td>{date_cloture}</td>
                <td style="max-width: 300px; word-wrap: break-word;">{str(row.get('Commentaire', ''))[:200]}...</td>
            </tr>"""
        
        # Collaborator analysis table
        table_html = ""
        for _, row in collaborator_analysis.head(50).iterrows():
            # Format dates - use the confirmed column structure
            date_creation = 'N/A'
            date_cloture = 'N/A'
            
            # Creation date from 'Créé le' column
            if 'Date_creation' in row.index and pd.notna(row['Date_creation']):
                try:
                    if hasattr(row['Date_creation'], 'strftime'):
                        date_creation = row['Date_creation'].strftime('%d/%m/%Y')
                    else:
                        date_creation = str(row['Date_creation'])[:10]
                except:
                    date_creation = str(row['Date_creation'])[:10] if row['Date_creation'] else 'N/A'
            
            # Closure date (survey response date)
            if 'Date_cloture' in row.index and pd.notna(row['Date_cloture']):
                try:
                    if hasattr(row['Date_cloture'], 'strftime'):
                        date_cloture = row['Date_cloture'].strftime('%d/%m/%Y')
                    else:
                        date_cloture = str(row['Date_cloture'])[:10]
                except:
                    date_cloture = str(row['Date_cloture'])[:10] if row['Date_cloture'] else 'N/A'
            
            # Calculate comment percentage for this collaborator in Page 2
            q1_all_collab = q1_unique[q1_unique['Créé par ticket'] == row['Collaborateur']]
            total_q1_collab = len(q1_all_collab)
            comment_pct = (int(row['Total_commentaires']) / total_q1_collab * 100) if total_q1_collab > 0 else 0
            
            # Color coding for comment percentage
            if comment_pct >= 20:
                comment_class = 'success'
            elif comment_pct >= 15:
                comment_class = 'warning'
            else:
                comment_class = 'danger'
            
            table_html += f"""
            <tr>
                <td>{row['Collaborateur']}</td>
                <td>{row['Site']}</td>
                <td>{int(row['Total_commentaires'])}</td>
                <td>{int(row['Commentaires_satisfaits'])}</td>
                <td>{int(row['Commentaires_insatisfaits'])}</td>
                <td class="text-{row['satisfaction_class']}">{row['Pourcentage_satisfait']:.1f}%</td>
                <td class="text-{comment_class}">{comment_pct:.1f}%</td>
            </tr>"""
        
        # Site comment synthesis - fix column references
        site_comment_synthesis = ""
        for _, row in site_analysis.iterrows():
            sat_rate = row.get('Pourcentage_satisfaction_commentaires', 0) if pd.notna(row.get('Pourcentage_satisfaction_commentaires')) else 0
            sat_class = row.get('satisfaction_class', 'secondary')
            total_comments = int(row.get('Total_commentaires', 0))
            site_comment_synthesis += f"""
            <tr>
                <td>{row['Site']}</td>
                <td>{int(row['Total_enquetes'])}</td>
                <td>{total_comments}</td>
                <td>{row['Pourcentage_commentaires']:.1f}%</td>
                <td class="text-{sat_class}">{sat_rate:.1f}%</td>
            </tr>"""
        
        # CORRECTION CRITIQUE: Utiliser les mêmes données que le data_processor pour cohérence
        from utils.data_processor import DataProcessor
        processor = DataProcessor()
        page2_data = processor.calculate_page2_enhanced_metrics(data)
        
        # Utiliser la distribution satisfaction harmonisée
        satisfaction_dist = page2_data.get('satisfaction_distribution', {})
        if satisfaction_dist:
            # FORCER L'ORDRE EXACT : Très satisfaisant, Satisfaisant, Peu satisfaisant, Très peu satisfaisant
            ordre_exact = ['Très satisfaisant', 'Satisfaisant', 'Peu satisfaisant', 'Très peu satisfaisant']
            satisfaction_data = {}
            for satisfaction in ordre_exact:
                satisfaction_data[satisfaction] = satisfaction_dist.get(satisfaction, {}).get('count', 0)
            satisfaction_series = pd.Series(satisfaction_data)
            pie_chart_comments = self._generate_pie_chart_base64(
                satisfaction_series,
                title="Synthèse des commentaires par satisfaction", 
                colors=['#1e7e34', '#28a745', '#fd7e14', '#dc3545']  # Ordre souhaité: Très satisfaisant=vert foncé, Satisfaisant=vert moyen, Peu satisfaisant=orange, Très peu satisfaisant=rouge
            )
        else:
            # Fallback vers ancienne méthode
            pie_chart_comments = self._generate_pie_chart_base64(
                with_comments['Valeur de chaîne'].value_counts(),
                title="Synthèse des commentaires par satisfaction",
                colors=['#1e7e34', '#28a745', '#fd7e14', '#dc3545']  # Fallback ordre souhaité
            )
        
        # Generate pie chart by site
        pie_chart_sites = self._generate_pie_chart_base64(
            site_comments.set_index('Site')['Total_commentaires'],
            title="Répartition des commentaires par site",
            colors=['#007bff', '#6c757d', '#fd7e14', '#e83e8c', '#20c997']
        )
        
        return {
            'page2_comments_count': total_with_comments,
            'page2_total_responses': total_q1_responses,
            'page2_comments_percentage': f"{comments_percentage:.1f}%",
            'page2_comments_table': comments_table,
            'page2_collab_analysis_table': table_html,
            'page2_site_comment_synthesis': site_comment_synthesis,
            'page2_pie_chart_comments': pie_chart_comments,
            'page2_pie_chart_sites': pie_chart_sites
        }
    
    def _get_page3_unsatisfied_no_comments(self, data):
        """Page 3 - Insatisfaits sans commentaires avec codes compte et noms boutiques"""
        df_merged = data['merged']
        
        # Q1 unique data only
        q1_data = df_merged[df_merged['Mesure'].str.contains('Q1', na=False)]
        q1_unique = q1_data.drop_duplicates('Dossier Rcbt')
        
        # Unsatisfied without comments
        unsatisfied_no_comments = q1_unique[
            (q1_unique['Valeur de chaîne'].isin(['Très peu satisfaisant', 'Peu satisfaisant'])) &
            (q1_unique['Commentaire'].isna() | (q1_unique['Commentaire'].str.strip() == ''))
        ]
        
        table_html = ""
        for _, row in unsatisfied_no_comments.head(50).iterrows():
            # Get account code without decimals and shop name
            code_compte = int(float(row.get('Code compte', 0))) if pd.notna(row.get('Code compte')) else 'N/A'
            
            # Get shop name from available columns
            nom_boutique = 'N/A'
            for name_col in ['Compte', 'Account', 'Nom']:
                if name_col in row.index and pd.notna(row.get(name_col)) and str(row.get(name_col)).strip():
                    nom_boutique = row.get(name_col)
                    break
            
            q1_class = self._get_q1_satisfaction_class(row.get('Valeur de chaîne', ''))
            
            # Get ticket number and dates like in page 2
            numero_ticket = row.get('Dossier Rcbt', 'N/A')
            date_creation = 'N/A'
            date_cloture = 'N/A'
            
            if 'Créé le' in row.index and pd.notna(row['Créé le']):
                try:
                    if hasattr(row['Créé le'], 'strftime'):
                        date_creation = row['Créé le'].strftime('%d/%m/%Y')
                    else:
                        date_creation = str(row['Créé le'])[:10]
                except:
                    date_creation = str(row['Créé le'])[:10] if row['Créé le'] else 'N/A'
            
            # Look for closure date column
            for col in ['Mise à jour', 'Clos le', 'Date cloture']:
                if col in row.index and pd.notna(row[col]):
                    try:
                        if hasattr(row[col], 'strftime'):
                            date_cloture = row[col].strftime('%d/%m/%Y')
                        else:
                            date_cloture = str(row[col])[:10]
                        break
                    except:
                        continue
            
            table_html += f"""
            <tr>
                <td>{code_compte}</td>
                <td>{nom_boutique}</td>
                <td>{row.get('Site', 'N/A')}</td>
                <td>{row.get('Créé par ticket', 'N/A')}</td>
                <td>{row.get('Mesure', 'N/A')}</td>
                <td>{numero_ticket}</td>
                <td>{row.get('Clos par', 'N/A')}</td>
                <td class="{q1_class}">{row.get('Valeur de chaîne', 'N/A')}</td>
                <td>{date_creation}</td>
                <td>{date_cloture}</td>
            </tr>"""
        
        return {
            'page3_count': len(unsatisfied_no_comments),
            'page3_table': table_html
        }
    
    def _get_page4_global_analysis_colored(self, data):
        """Page 4 - Analyse globale avec codes couleur pour satisfaction
        CORRECTION: Inclut TOUTES les enquêtes Q1 (avec ET sans commentaires)"""
        df_merged = data['merged']
        
        # Q1 unique data only
        q1_data = df_merged[df_merged['Mesure'].str.contains('Q1', na=False)]
        q1_unique = q1_data.drop_duplicates('Dossier Rcbt')
        
        # Q1 response distribution
        q1_counts = q1_unique['Valeur de chaîne'].value_counts()
        q1_table = ""
        for response, count in q1_counts.items():
            if pd.notna(response):
                q1_class = self._get_q1_satisfaction_class(response)
                q1_table += f'<tr><td class="{q1_class}">{response}</td><td>{count}</td></tr>'
        
        # Collaborator analysis with colored satisfaction results and comment percentage
        df_merged = data['merged']
        comments_data = df_merged[df_merged['Commentaire'].notna() & (df_merged['Commentaire'].str.strip() != '')]
        
        collab_detail = []
        
        # CORRECTION: Le problème vient du fait que les collaborateurs dans asmt ('Créé par') 
        # ne correspondent pas aux collaborateurs dans case ('Créé par')
        # Il faut utiliser la colonne correcte selon la structure des données
        
        # Identifier la colonne collaborateur dans merged data
        collab_column = 'Créé par ticket' if 'Créé par ticket' in df_merged.columns else 'Créé par'
        print(f"🔍 Colonne collaborateur utilisée: '{collab_column}'")
        
        for (collaborateur, site), group_data in q1_unique.groupby([collab_column, 'Site']):
            # Analyser TOUS les collaborateurs (sans limitation de nombre de réponses)
            # group_data = TOUTES les enquêtes Q1 de ce collaborateur (avec ET sans commentaires)
            satisf_counts = group_data['Valeur de chaîne'].value_counts()
            total = len(group_data)
            satisf = satisf_counts.get('Très satisfaisant', 0) + satisf_counts.get('Satisfaisant', 0)
            satisfaction_rate = (satisf / total * 100) if total > 0 else 0
                
            # CORRECTION: Calculate comment percentage for this collaborator based on ALL Q1 responses
            # Utiliser toutes les enquêtes Q1 du collaborateur (pas seulement celles avec commentaires)
            collab_all_q1 = group_data  # group_data contient déjà toutes les Q1 de ce collaborateur
            collab_comments_only = comments_data[comments_data[collab_column] == collaborateur]
            collab_q1_comments = collab_comments_only[collab_comments_only['Mesure'].str.contains('Q1', na=False)]
            collab_unique_comments = collab_q1_comments.drop_duplicates('Dossier Rcbt')
            
            # Pourcentage de commentaires = enquêtes avec commentaires / toutes les enquêtes Q1
            comment_percentage = (len(collab_unique_comments) / total * 100) if total > 0 else 0
            
            # CORRECTION DU CALCUL TAUX DE RETOUR
            # D'après l'analyse : le collaborateur 'alneves' devrait avoir 47/543 = 8.7%
            # La logique correcte: réponses Q1 / tickets créés par le collaborateur
            
            # Utiliser la colonne 'Créé par ticket' après le merge pour trouver les tickets originaux
            if 'case' in data:
                # Chercher dans les données case avec la colonne renommée 'Créé par ticket'  
                case_data = data['case']
                case_col = 'Créé par ticket' if 'Créé par ticket' in case_data.columns else 'Créé par'
                
                collab_created_tickets = case_data[case_data[case_col] == collaborateur]
                total_tickets_created = len(collab_created_tickets['Numéro'].drop_duplicates()) if len(collab_created_tickets) > 0 else 0
                
                if total_tickets_created > 0:
                    # Nombre de réponses Q1 uniques pour ce collaborateur
                    q1_responses_count = len(group_data['Dossier Rcbt'].drop_duplicates())
                    return_rate = (q1_responses_count / total_tickets_created * 100)
                    print(f"🔍 {collaborateur}: {q1_responses_count} réponses Q1 / {total_tickets_created} tickets créés = {return_rate:.1f}%")
                else:
                    # Ce collaborateur répond mais n'a pas créé de tickets
                    return_rate = 0.0  # Ou marquer comme "répondant seulement"
                    print(f"🔍 {collaborateur}: Répondant uniquement (0 tickets créés)")
            else:
                # Fallback si pas de données case
                return_rate = 0.0
                print(f"🔍 {collaborateur} (pas de données case): 0.0%")
            
            collab_detail.append({
                'name': collaborateur,
                'site': site,
                'tres_satisf': satisf_counts.get('Très satisfaisant', 0),
                'satisf': satisf_counts.get('Satisfaisant', 0),
                'peu_satisf': satisf_counts.get('Peu satisfaisant', 0),
                'tres_peu_satisf': satisf_counts.get('Très peu satisfaisant', 0),
                'total': total,
                'satisfaction_rate': satisfaction_rate,
                'comment_percentage': comment_percentage,
                'return_rate': return_rate
            })
        
        # Build collaborator table with color coding including comment percentage
        collab_table = ""
        for collab in sorted(collab_detail, key=lambda x: x['satisfaction_rate'], reverse=True):
            sat_class = 'success' if collab['satisfaction_rate'] >= 92 else 'warning'
            
            # Comment percentage color coding
            comment_pct = collab['comment_percentage']
            if comment_pct >= 20:
                comment_class = 'success'
            elif comment_pct >= 15:
                comment_class = 'warning' 
            else:
                comment_class = 'danger'
            
            # Return rate color coding
            return_rate = collab.get('return_rate', 0)
            if return_rate >= 15:
                return_class = 'success'
            elif return_rate >= 10:
                return_class = 'warning'
            else:
                return_class = 'danger'
            
            collab_table += f"""
            <tr>
                <td>{collab['name']}</td>
                <td>{collab['site']}</td>
                <td class="q1-tres-satisfaisant">{collab['tres_satisf']}</td>
                <td class="q1-satisfaisant">{collab['satisf']}</td>
                <td class="q1-peu-satisfaisant">{collab['peu_satisf']}</td>
                <td class="q1-tres-peu-satisfaisant">{collab['tres_peu_satisf']}</td>
                <td>{collab['total']}</td>
                <td class="text-{sat_class}">{collab['satisfaction_rate']:.1f}%</td>
                <td class="text-{comment_class}">{comment_pct:.1f}%</td>
                <td class="text-{return_class}">{return_rate:.1f}%</td>
            </tr>"""
        
        return {
            'page4_q1_table': q1_table,
            'page4_collab_table': collab_table,
            'page4_collab_count': len(collab_detail)
        }
    
    def _get_page5_responding_shops_colored(self, data):
        """Page 5 - Boutiques répondues avec codes couleur pour objectifs"""
        df_merged = data['merged']
        df_accounts = data['accounts']
        
        # Get unique responding shops from Q1 data
        q1_data = df_merged[df_merged['Mesure'].str.contains('Q1', na=False)]
        responding_shops = q1_data['Code compte'].unique()
        
        # CORRECTION: Utiliser seulement les boutiques qui ont ouvert au moins un ticket
        df_case = data['case']
        case_boutiques = df_case[df_case['Boutique_categorie'] != 'Autres']
        boutiques_with_tickets = case_boutiques['Code compte'].unique()
        
        # Category analysis with objectives coloring
        category_stats = []
        for category in ['400', '499', '993', 'Mini-enseigne', 'Siège']:
            # Filtrer par boutiques qui ont ouvert des tickets ET appartiennent à cette catégorie
            cat_accounts = df_accounts[df_accounts['Categorie'] == category]
            cat_with_tickets = cat_accounts[cat_accounts['Code compte'].isin(boutiques_with_tickets)]
            
            total_open = len(cat_with_tickets)  # Seulement les boutiques avec tickets
            responded = len(cat_with_tickets[cat_with_tickets['Code compte'].isin(responding_shops)])
            never_responded = total_open - responded
            
            response_rate = (responded / total_open * 100) if total_open > 0 else 0
            never_rate = (never_responded / total_open * 100) if total_open > 0 else 0
            
            objective_response_ok = response_rate >= 30
            objective_never_ok = never_rate <= 70
            
            # Compter les boutiques totales dans cette catégorie (fichier accounts)
            cat_all_accounts = df_accounts[df_accounts['Categorie'] == category]
            total_open_all = len(cat_all_accounts)
            
            category_stats.append({
                'category': category,
                'total_open': total_open_all,
                'total_with_tickets': total_open,
                'responded': responded,
                'response_rate': response_rate,
                'never_responded': never_responded,
                'never_rate': never_rate,
                'objective_response_ok': objective_response_ok,
                'objective_never_ok': objective_never_ok
            })
        
        # Build category table with objective color coding and correct column structure
        category_table = ""
        for stat in category_stats:
            response_class = 'success' if stat['objective_response_ok'] else 'warning'
            never_class = 'success' if stat['objective_never_ok'] else 'warning'
            
            category_table += f"""
            <tr>
                <td>{stat['category']}</td>
                <td>{stat['total_open']}</td>
                <td>{stat['total_with_tickets']}</td>
                <td>{stat['responded']}</td>
                <td class="text-{response_class}">{stat['response_rate']:.1f}%</td>
                <td>{stat['never_responded']}</td>
                <td class="text-{never_class}">{stat['never_rate']:.1f}%</td>
            </tr>"""
        
        # Responding shops list - fix the N/A issue
        responding_accounts = df_accounts[df_accounts['Code compte'].isin(responding_shops)]
        shops_table = ""
        for _, row in responding_accounts.iterrows():
            code_compte = int(float(row.get('Code compte', 0))) if pd.notna(row.get('Code compte')) else 'N/A'
            
            # Get account name from available columns  
            account_name = 'N/A'
            for name_col in ['Compte', 'Account', 'Nom']:
                if name_col in row.index and pd.notna(row.get(name_col)) and str(row.get(name_col)).strip():
                    account_name = row.get(name_col)
                    break
            
            # If still N/A, try from merged data
            if account_name == 'N/A':
                matching_merged = df_merged[df_merged['Code compte'] == row.get('Code compte')]
                if len(matching_merged) > 0:
                    for name_col in ['Compte', 'Account', 'Nom']:
                        if name_col in matching_merged.columns and pd.notna(matching_merged.iloc[0].get(name_col)):
                            account_name = matching_merged.iloc[0].get(name_col)
                            break
            
            shops_table += f"""
            <tr>
                <td>{code_compte}</td>
                <td>{account_name}</td>
                <td>{row.get('Categorie', 'N/A')}</td>
            </tr>"""
        
        # Generate pie chart for shop response rate by category
        response_data = pd.Series({stat['category']: stat['response_rate'] for stat in category_stats})
        pie_chart_response_rate = self._generate_pie_chart_base64(
            response_data,
            title="Taux de réponse par catégorie (%)",
            colors=['#28a745', '#17a2b8', '#ffc107', '#dc3545', '#6f42c1']
        )
        
        # Calculate global response rate - CORRECTION: uniquement boutiques avec tickets
        accounts_with_tickets = df_accounts[df_accounts['Code compte'].isin(boutiques_with_tickets)]
        total_open_shops = len(accounts_with_tickets)
        total_responding_shops = len(responding_accounts)
        global_response_rate = (total_responding_shops / total_open_shops * 100) if total_open_shops > 0 else 0
        print(f"🔍 Page 5: {total_responding_shops} boutiques répondues / {total_open_shops} boutiques avec tickets = {global_response_rate:.1f}%")
        
        global_objective_class = 'success' if global_response_rate >= 30 else 'danger'
        global_objective_icon = '✅' if global_response_rate >= 30 else '❌'
        
        # Generate individual pie charts for satisfaction by category
        category_pie_charts = ""
        for category in ['400', '499', '993', 'Siège']:
            # Get Q1 data for this category - join with accounts to get category info
            if 'Categorie' in df_accounts.columns:
                cat_accounts = df_accounts[df_accounts['Categorie'] == category]['Code compte'].values
                cat_q1_data = q1_data[q1_data['Code compte'].isin(cat_accounts)]
                
                if not cat_q1_data.empty:
                    satisf_counts = cat_q1_data['Valeur de chaîne'].value_counts()
                    total_cat = len(cat_q1_data)
                    satisf_cat = satisf_counts.get('Très satisfaisant', 0) + satisf_counts.get('Satisfaisant', 0)
                    insatisf_cat = total_cat - satisf_cat
                    
                    if total_cat > 0:
                        satisf_pct = (satisf_cat / total_cat * 100)
                        insatisf_pct = (insatisf_cat / total_cat * 100)
                        
                        pie_data = pd.Series([satisf_pct, insatisf_pct], 
                                           index=['Satisfait', 'Insatisfait'])
                        
                        cat_pie_chart = self._generate_pie_chart_base64(
                            pie_data,
                            title=f"Satisfaction {category} (n={total_cat})",
                            colors=['#28a745', '#dc3545']
                        )
                        
                        category_pie_charts += f'<div class="site-pie-chart"><img src="data:image/png;base64,{cat_pie_chart}" alt="Satisfaction {category}"></div>'

        return {
            'page5_category_table': category_table,
            'page5_shops_table': shops_table,
            'page5_responding_count': len(responding_accounts),
            'page5_category_pie_charts': category_pie_charts,
            'page5_global_response_rate': f"{global_response_rate:.1f}",
            'page5_global_objective_class': global_objective_class,
            'page5_global_objective_icon': global_objective_icon
        }
    
    def _get_page6_never_responding(self, data):
        """Page 6 - Boutiques jamais répondues avec liste détaillée (CONTEXTUALISÉ)"""
        df_merged = data['merged']
        df_case = data['case'] 
        df_accounts = data['accounts']
        
        # CORRECTION: Filtrage selon contexte (collaborateur/site)
        filter_type = data.get('filter_type')
        filter_value = data.get('filter_value')
        
        # Appliquer le filtrage contextuel aux données case
        context_case = df_case.copy()
        if filter_type == 'collaborator' and filter_value:
            context_case = context_case[context_case['Créé par ticket'] == filter_value]
        elif filter_type == 'site' and filter_value:
            # Ajouter le mapping site si nécessaire
            if 'Site' not in context_case.columns and 'ref' in data:
                df_ref = data['ref']
                if not df_ref.empty and 'Login' in df_ref.columns and 'Site' in df_ref.columns:
                    site_mapping = dict(zip(df_ref['Login'], df_ref['Site']))
                    context_case['Site'] = context_case['Créé par ticket'].map(site_mapping)
                    context_case['Site'] = context_case['Site'].fillna('Autres')
            context_case = context_case[context_case.get('Site', '') == filter_value]
        
        # Get shops that have opened tickets in this CONTEXT
        boutiques_with_tickets = context_case['Code compte'].unique()
        
        # Get shops that have responded to Q1 in this context  
        q1_data = df_merged[df_merged['Mesure'].str.contains('Q1', na=False)]
        responding_shops = q1_data['Code compte'].unique()
        
        # Never responding shops = shops with tickets but no responses IN THIS CONTEXT
        never_responding_shop_codes = set(boutiques_with_tickets) - set(responding_shops)
        
        # Get detailed shop info for never responding shops (FILTER ACCOUNTS TOO)
        never_responding_shops = df_accounts[df_accounts['Code compte'].isin(never_responding_shop_codes)]
        
        # CORRECTION: Utiliser seulement les boutiques qui ont ouvert au moins un ticket  
        df_case = data['case']
        case_boutiques = df_case[df_case['Boutique_categorie'] != 'Autres']
        boutiques_with_tickets = case_boutiques['Code compte'].unique()
        
        # Category analysis for never responding shops
        category_stats = []
        for category in ['400', '499', '993', 'Mini-enseigne', 'Siège']:
            cat_accounts = df_accounts[df_accounts['Categorie'] == category]
            cat_with_tickets = cat_accounts[cat_accounts['Code compte'].isin(boutiques_with_tickets)]
            
            total_open = len(cat_with_tickets)  # Seulement les boutiques avec tickets
            never_responded = len(cat_with_tickets[~cat_with_tickets['Code compte'].isin(responding_shops)])
            never_rate = (never_responded / total_open * 100) if total_open > 0 else 0
            
            # Compter les boutiques totales dans cette catégorie (fichier accounts)
            total_open_all = len(cat_accounts)
            
            category_stats.append({
                'category': category,
                'total_open': total_open_all,
                'total_with_tickets': total_open,
                'never_responded': never_responded,
                'never_rate': never_rate
            })
        
        # Build category table for never responding with correct structure
        never_category_table = ""
        for stat in category_stats:
            never_category_table += f"""
            <tr>
                <td>{stat['category']}</td>
                <td>{stat['total_open']}</td>
                <td>{stat['total_with_tickets']}</td>
                <td>{stat['never_responded']}</td>
                <td>{stat['never_rate']:.1f}%</td>
            </tr>"""
        
        # Réutiliser les stats déjà calculées pour éviter la duplication
        final_category_stats = []
        for stat in category_stats:
            final_category_stats.append({
                'category': stat['category'],
                'total_open': stat['total_open'],
                'total_with_tickets': stat['total_with_tickets'],
                'never_responded': stat['never_responded'],
                'never_rate': stat['never_rate']
            })
        
        # Category summary table
        category_table_html = ""
        for stat in final_category_stats:
            objective_class = 'success' if stat['never_rate'] <= 70 else 'warning'
            category_table_html += f"""
            <tr>
                <td>{stat['category']}</td>
                <td>{stat['total_open']}</td>
                <td>{stat['total_with_tickets']}</td>
                <td>{stat['never_responded']}</td>
                <td class="text-{objective_class}">{stat['never_rate']:.1f}%</td>
            </tr>"""
        
        # Detailed list of never responding shops
        never_shops_table = ""
        for _, row in never_responding_shops.iterrows():
            code_compte = int(float(row.get('Code compte', 0))) if pd.notna(row.get('Code compte')) else 'N/A'
            
            # Get account name from available columns
            account_name = 'N/A'
            for name_col in ['Compte', 'Account', 'Nom']:
                if name_col in row.index and pd.notna(row.get(name_col)) and str(row.get(name_col)).strip():
                    account_name = row.get(name_col)
                    break
            
            never_shops_table += f"""
            <tr>
                <td>{code_compte}</td>
                <td>{account_name}</td>
                <td>{row.get('Categorie', 'N/A')}</td>
            </tr>"""
        
        # Calculate global never response rate - CORRECTION: uniquement boutiques avec tickets
        accounts_with_tickets = df_accounts[df_accounts['Code compte'].isin(boutiques_with_tickets)]
        never_responding_with_tickets = accounts_with_tickets[~accounts_with_tickets['Code compte'].isin(responding_shops)]
        
        total_open_shops = len(accounts_with_tickets)
        total_never_shops = len(never_responding_with_tickets)
        global_never_rate = (total_never_shops / total_open_shops * 100) if total_open_shops > 0 else 0
        print(f"🔍 Page 6: {total_never_shops} boutiques jamais répondues / {total_open_shops} boutiques avec tickets = {global_never_rate:.1f}%")
        
        global_objective_class = 'success' if global_never_rate <= 70 else 'danger'
        global_objective_icon = '✅' if global_never_rate <= 70 else '❌'
        
        return {
            'page6_category_table': category_table_html,
            'page6_never_table': never_shops_table,
            'page6_never_count': len(never_responding_shops),
            'page6_global_never_rate': f"{global_never_rate:.1f}",
            'page6_global_objective_class': global_objective_class,
            'page6_global_objective_icon': global_objective_icon
        }
    
    def _get_page7_shop_ranking(self, data):
        """Page 7 - Classement par boutique (tickets et enquêtes)"""
        df_merged = data['merged']
        df_case = data['case']
        df_accounts = data['accounts']
        
        # Get all shops with ticket counts
        shop_ticket_counts = df_case.groupby('Code compte').agg({
            'Numéro': 'nunique',
            'État': lambda x: (x == 'Clos').sum()
        }).reset_index()
        shop_ticket_counts.columns = ['Code compte', 'Total_tickets', 'Tickets_clos']
        shop_ticket_counts['Taux_closure'] = (shop_ticket_counts['Tickets_clos'] / shop_ticket_counts['Total_tickets'] * 100).round(1)
        
        # Get survey response counts
        q1_data = df_merged[df_merged['Mesure'].str.contains('Q1', na=False)]
        survey_counts = q1_data.groupby('Code compte').agg({
            'Dossier Rcbt': 'nunique',
            'Valeur de chaîne': lambda x: (x.isin(['Très satisfaisant', 'Satisfaisant'])).sum()
        }).reset_index()
        survey_counts.columns = ['Code compte', 'Enquetes_repondues', 'Enquetes_satisfaites']
        survey_counts['Taux_satisfaction'] = (survey_counts['Enquetes_satisfaites'] / survey_counts['Enquetes_repondues'] * 100).round(1)
        
        # Merge all data
        shop_ranking = shop_ticket_counts.merge(survey_counts, on='Code compte', how='left')
        
        # Check available columns in accounts and adapt
        account_cols = ['Code compte']
        if 'Compte' in df_accounts.columns:
            account_cols.append('Compte')
        elif 'Account' in df_accounts.columns:
            account_cols.append('Account')
        elif 'Nom' in df_accounts.columns:
            account_cols.append('Nom')
        else:
            # Use first text column after Code compte
            for col in df_accounts.columns:
                if col != 'Code compte' and df_accounts[col].dtype == 'object':
                    account_cols.append(col)
                    break
        
        if 'Categorie' in df_accounts.columns:
            account_cols.append('Categorie')
        elif 'Category' in df_accounts.columns:
            account_cols.append('Category')
        
        shop_ranking = shop_ranking.merge(df_accounts[account_cols], on='Code compte', how='left')
        shop_ranking['Enquetes_repondues'] = shop_ranking['Enquetes_repondues'].fillna(0)
        shop_ranking['Enquetes_satisfaites'] = shop_ranking['Enquetes_satisfaites'].fillna(0)
        shop_ranking['Taux_satisfaction'] = shop_ranking['Taux_satisfaction'].fillna(0)
        
        # Sort by total tickets (descending) - CRITÈRE DE CLASSEMENT
        shop_ranking = shop_ranking.sort_values('Total_tickets', ascending=False)
        
        # Display all shops that have responded to surveys, not just top 50
        all_responding_shops = shop_ranking[shop_ranking['Enquetes_repondues'] > 0]
        
        # Build ranking table
        ranking_table = ""
        # Determine account name column
        name_col = 'Compte' if 'Compte' in shop_ranking.columns else ('Account' if 'Account' in shop_ranking.columns else 'Nom')
        cat_col = 'Categorie' if 'Categorie' in shop_ranking.columns else 'Category'
        
        for i, (_, row) in enumerate(all_responding_shops.iterrows(), 1):
            code_compte = int(float(row.get('Code compte', 0))) if pd.notna(row.get('Code compte')) else 'N/A'
            account_name = row.get(name_col, 'N/A') if name_col in shop_ranking.columns else 'N/A'
            
            # Color coding for satisfaction rate and display
            enquetes_repondues = int(row.get('Enquetes_repondues', 0))
            sat_rate = row.get('Taux_satisfaction', 0)
            
            if enquetes_repondues == 0:
                sat_display = 'NC'
                sat_class = 'secondary'
            else:
                sat_display = f'{sat_rate:.1f}%'
                if sat_rate >= 92:
                    sat_class = 'success'
                elif sat_rate >= 80:
                    sat_class = 'warning'
                else:
                    sat_class = 'danger'
            
            ranking_table += f"""
            <tr>
                <td>{i}</td>
                <td>{code_compte}</td>
                <td>{account_name}</td>
                <td>{row.get(cat_col, 'N/A') if cat_col in shop_ranking.columns else 'N/A'}</td>
                <td>{int(row.get('Tickets_clos', 0))}</td>
                <td class="text-{('success' if (row.get('Enquetes_repondues', 0) / row.get('Total_tickets', 1) * 100) >= 13.0 else 'danger')}">{(row.get('Enquetes_repondues', 0) / row.get('Total_tickets', 1) * 100):.1f}%</td>
                <td>{enquetes_repondues}</td>
                <td>{int(row.get('Enquetes_satisfaites', 0))}</td>
                <td class="text-{sat_class}">{sat_display}</td>
            </tr>"""
        
        return {
            'page7_ranking_table': ranking_table,
            'page7_total_shops': len(all_responding_shops)
        }
    
    def _get_q1_satisfaction_class(self, value):
        """Get CSS class for Q1 satisfaction coloring"""
        if value == 'Très satisfaisant':
            return 'q1-tres-satisfaisant'
        elif value == 'Satisfaisant':
            return 'q1-satisfaisant'
        elif value == 'Peu satisfaisant':
            return 'q1-peu-satisfaisant'
        elif value == 'Très peu satisfaisant':
            return 'q1-tres-peu-satisfaisant'
        return ''
    
    def _is_comment_inconsistent(self, satisfaction, comment):
        """Détecter si un commentaire est incohérent avec la note de satisfaction - bidirectionnel"""
        if not comment or pd.isna(comment):
            return False
        
        comment_lower = str(comment).lower()
        
        # MOTS POSITIFS (forts et contextuels)
        positive_words = ['merci', 'parfait', 'top', 'efficace', 'rapide', 'solution', 
                         'clair', 'précis', 'excellent', 'super', 'formidable', 'génial',
                         'impeccable', 'extraordinaire', 'fantastique', 'remarquable']
        
        # MOTS NÉGATIFS (forts et contextuels) 
        negative_words = ['catastrophique', 'horrible', 'nul', 'inadmissible', 'scandaleux',
                         'inacceptable', 'décevant', 'frustrant', 'agacé', 'énervé', 'furieux',
                         'incompétent', 'déplorable', 'lamentable', 'désastreux']
        
        # TYPE 1: Notes négatives avec commentaires positifs (fiabilisé)
        if satisfaction in ['Peu satisfaisant', 'Très peu satisfaisant']:
            positive_found = [word for word in positive_words if word in comment_lower]
            strong_words = ['parfait', 'excellent', 'formidable', 'impeccable', 'extraordinaire']
            strong_found = [word for word in strong_words if word in comment_lower]
            # Au moins 2 mots positifs ou 1 mot très fort
            return len(positive_found) >= 2 or len(strong_found) >= 1
        
        # TYPE 2: Notes positives avec commentaires négatifs
        if satisfaction in ['Très satisfaisant', 'Satisfaisant']:
            negative_found = [word for word in negative_words if word in comment_lower]
            # Au moins 1 mot négatif fort
            return len(negative_found) >= 1
        
        return False

    def _get_inconsistency_details(self, satisfaction, comment):
        """Get details about the inconsistency for tooltip"""
        if not comment or pd.isna(comment):
            return {'type': 'Aucune', 'words': [], 'impact': 'Aucun impact'}
        
        comment_lower = str(comment).lower()
        
        positive_words = ['merci', 'parfait', 'top', 'efficace', 'rapide', 'solution', 
                         'clair', 'précis', 'excellent', 'super', 'formidable', 'génial',
                         'impeccable', 'extraordinaire', 'fantastique', 'remarquable']
        
        negative_words = ['catastrophique', 'horrible', 'nul', 'inadmissible', 'scandaleux',
                         'inacceptable', 'décevant', 'frustrant', 'agacé', 'énervé', 'furieux',
                         'incompétent', 'déplorable', 'lamentable', 'désastreux']
        
        if satisfaction in ['Peu satisfaisant', 'Très peu satisfaisant']:
            found_words = [word for word in positive_words if word in comment_lower]
            return {
                'type': 'Note négative avec commentaire positif',
                'words': found_words,
                'impact': 'Ticket signalé pour vérification manuelle'
            }
        elif satisfaction in ['Très satisfaisant', 'Satisfaisant']:
            found_words = [word for word in negative_words if word in comment_lower]
            return {
                'type': 'Note positive avec commentaire négatif', 
                'words': found_words,
                'impact': 'Ticket signalé pour vérification manuelle'
            }
        
        return {'type': 'Aucune incohérence', 'words': [], 'impact': 'Aucun impact'}
    
    def _generate_pie_chart_base64(self, data, title, colors):
        """Generate pie chart as base64 image with percentages"""
        try:
            import matplotlib.pyplot as plt
            import io
            
            # Create pie chart
            fig, ax = plt.subplots(figsize=(8, 6))
            result = ax.pie(
                data.values, 
                labels=data.index, 
                autopct='%1.1f%%',
                colors=colors[:len(data)],
                startangle=90
            )
            
            # Unpack based on result length
            if len(result) == 3:
                wedges, texts, autotexts = result
            else:
                wedges, texts = result
                autotexts = []
            
            # Style the chart
            ax.set_title(title, fontsize=14, fontweight='bold')
            if autotexts:
                plt.setp(autotexts, size=10, weight="bold")
            plt.setp(texts, size=9)
            
            # Convert to base64
            buffer = io.BytesIO()
            plt.savefig(buffer, format='png', bbox_inches='tight', dpi=150)
            buffer.seek(0)
            chart_base64 = base64.b64encode(buffer.getvalue()).decode()
            plt.close()
            
            return chart_base64
            
        except Exception as e:
            print(f"⚠️ Erreur génération graphique: {e}")
            return ""
    
    def _get_optimized_template(self):
        """Get optimized HTML template with all features"""
        return '''<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Rapport RCBT Optimisé - {timestamp}</title>
    <style>
        /* Professional RCBT styling */
        :root {{
            --rcbt-primary: #1e3a8a;
            --rcbt-secondary: #3b82f6;
            --rcbt-success: #22c55e;
            --rcbt-warning: #f59e0b;
            --rcbt-danger: #ef4444;
            --rcbt-light: #f8fafc;
        }}
        
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 0;
            padding: 0;
            background-color: var(--rcbt-light);
            color: #374151;
            line-height: 1.6;
        }}
        
        .header {{
            background: linear-gradient(135deg, var(--rcbt-primary), var(--rcbt-secondary));
            color: white;
            padding: 20px;
            text-align: center;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }}
        
        .logo {{
            max-height: 60px;
            margin-bottom: 10px;
        }}
        
        .nav-tabs {{
            display: flex;
            background: white;
            border-bottom: 3px solid var(--rcbt-primary);
            overflow-x: auto;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }}
        
        .tab {{
            padding: 15px 20px;
            background: white;
            border: none;
            cursor: pointer;
            font-size: 14px;
            font-weight: 600;
            color: var(--rcbt-primary);
            border-bottom: 3px solid transparent;
            transition: all 0.3s ease;
            white-space: nowrap;
        }}
        
        .tab:hover {{
            background-color: #f1f5f9;
            border-bottom-color: var(--rcbt-secondary);
        }}
        
        .tab.active {{
            background-color: var(--rcbt-primary);
            color: white;
            border-bottom-color: var(--rcbt-warning);
        }}
        
        .tab-content {{
            display: none;
            padding: 30px;
            max-width: 1400px;
            margin: 0 auto;
        }}
        
        .tab-content.active {{
            display: block;
        }}
        
        .section-title {{
            color: var(--rcbt-primary);
            font-size: 24px;
            margin-bottom: 20px;
            border-bottom: 2px solid var(--rcbt-secondary);
            padding-bottom: 10px;
        }}
        
        .metric-highlight {{
            background: linear-gradient(135deg, #fef3c7, #fbbf24);
            padding: 15px;
            border-radius: 8px;
            margin: 20px 0;
            border-left: 5px solid var(--rcbt-warning);
            font-weight: 600;
        }}
        
        .kpi-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 20px;
            margin: 20px 0;
        }}
        
        .kpi-card {{
            background: white;
            padding: 20px;
            border-radius: 10px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
            text-align: center;
            border-top: 4px solid var(--rcbt-secondary);
        }}
        
        .kpi-value {{
            font-size: 32px;
            font-weight: bold;
            margin: 10px 0;
        }}
        
        .kpi-label {{
            color: #6b7280;
            font-size: 14px;
            text-transform: uppercase;
            letter-spacing: 1px;
        }}
        
        .text-success {{ color: var(--rcbt-success) !important; }}
        .text-warning {{ color: var(--rcbt-warning) !important; }}
        .text-danger {{ color: var(--rcbt-danger) !important; }}
        
        /* Q1 Satisfaction Colors - CORRECTED */
        .q1-tres-satisfaisant {{ background-color: #1e7e34 !important; color: white; }}  /* VERT FONCÉ - très satisfaisant */
        .q1-satisfaisant {{ background-color: #28a745 !important; color: white; }}  /* VERT MOYEN - satisfaisant */
        .q1-peu-satisfaisant {{ background-color: #fd7e14 !important; color: #92400e; }}  /* ORANGE - peu satisfaisant */
        .q1-tres-peu-satisfaisant {{ background-color: #dc3545 !important; color: white; }}  /* ROUGE - très peu satisfaisant */
        
        /* Advanced filtering system */
        .search-container {{
            background: white;
            padding: 20px;
            border-radius: 8px;
            margin-bottom: 20px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }}
        
        .filter-label {{
            font-weight: 600;
            color: var(--rcbt-primary);
            margin-bottom: 10px;
            display: block;
        }}
        
        .search-input, .column-filter {{
            width: 100%;
            padding: 10px;
            border: 2px solid #e5e7eb;
            border-radius: 6px;
            margin-bottom: 10px;
            font-size: 14px;
            transition: border-color 0.3s;
        }}
        
        .search-input:focus, .column-filter:focus {{
            outline: none;
            border-color: var(--rcbt-secondary);
        }}
        
        .column-filters {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
            gap: 10px;
            margin-top: 15px;
        }}
        
        .clear-filters {{
            background: var(--rcbt-warning);
            color: white;
            border: none;
            padding: 10px 20px;
            border-radius: 6px;
            cursor: pointer;
            font-weight: 600;
            margin-left: 10px;
            transition: background 0.3s;
        }}
        
        .clear-filters:hover {{
            background: #d97706;
        }}
        
        .table-wrapper {{
            background: white;
            border-radius: 8px;
            overflow: hidden;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
            margin-bottom: 30px;
        }}
        
        table {{
            width: 100%;
            border-collapse: collapse;
            font-size: 13px;
        }}
        
        th {{
            background: var(--rcbt-primary);
            color: white;
            padding: 12px 8px;
            text-align: left;
            font-weight: 600;
            cursor: pointer;
            user-select: none;
            position: relative;
        }}
        
        th:hover {{
            background: var(--rcbt-secondary);
        }}
        
        th::after {{
            content: '⇅';
            position: absolute;
            right: 8px;
            opacity: 0.7;
        }}
        
        td {{
            padding: 10px 8px;
            border-bottom: 1px solid #e5e7eb;
        }}
        
        tbody tr:hover {{
            background-color: #f8fafc;
        }}
        
        .no-results {{
            text-align: center;
            padding: 40px;
            color: #6b7280;
            font-style: italic;
            display: none;
        }}
        
        .chart-container {{
            text-align: center;
            margin: 20px 0;
            background: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }}
        
        .chart-container img {{
            max-width: 100%;
            height: auto;
            border-radius: 8px;
        }}
        
        .pie-charts-grid {{
            display: flex;
            flex-wrap: wrap;
            gap: 20px;
            margin: 20px 0;
            justify-content: center;
        }}
        
        .site-pie-chart, .measure-pie-chart {{
            flex: 0 0 300px;
            text-align: center;
            background: white;
            border-radius: 8px;
            padding: 15px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            margin-bottom: 20px;
        }}
        
        .site-pie-chart img, .measure-pie-chart img {{
            max-width: 280px;
            height: auto;
            display: block;
            margin: 0 auto;
        }}
        
        .info-box {{
            background: #f8f9fa;
            border-left: 4px solid var(--rcbt-primary);
            padding: 15px;
            margin: 15px 0;
            border-radius: 5px;
        }}
        
        .text-secondary {{
            color: #6c757d !important;
        }}
        
        @media (max-width: 768px) {{
            .nav-tabs {{
                flex-direction: column;
            }}
            .tab {{
                text-align: center;
            }}
            .kpi-grid {{
                grid-template-columns: 1fr;
            }}
            .column-filters {{
                grid-template-columns: 1fr;
            }}
        }}
        
        @media print {{
            .nav-tabs {{ display: none; }}
            .tab-content {{ display: block !important; padding: 10px; }}
            .search-container {{ display: none; }}
        }}
    </style>
</head>
<body>
    <div class="header">
        <div style="text-align: center; margin-bottom: 20px;">
            <img src="data:image/png;base64,{logo_centre_base64}" style="height: 60px;" alt="Centre de Services">
        </div>
        <h1>📊 Application ISESIR by <img src="data:image/jpeg;base64,{logo_base64}" style="height: 24px; margin-left: 8px; vertical-align: middle;" alt="Q&T Logo"> - Rapport RCBT Optimisé</h1>
        <p>Généré le {timestamp}</p>
    </div>
    
    <div class="nav-tabs">
        <button class="tab active" onclick="showTab('page1')">📋 Synthèse</button>
        <button class="tab" onclick="showTab('page2')">💬 Commentaires</button>
        <button class="tab" onclick="showTab('page3')">😞 Sans commentaires</button>
        <button class="tab" onclick="showTab('page4')">📊 Analyse globale</button>
        <button class="tab" onclick="showTab('page5')">✅ Boutiques répondues</button>
        <button class="tab" onclick="showTab('page6')">❌ Jamais répondues</button>
        <button class="tab" onclick="showTab('page7')">🏆 Classement boutiques</button>
    </div>
    
    <div id="page1" class="tab-content active">
        <h2 class="section-title">Page 1 - Synthèse mensuelle</h2>
        
        <div class="kpi-grid">
            <div class="kpi-card">
                <div class="kpi-label">Taux de clôture boutiques</div>
                <div class="kpi-value text-{page1_closure_color}">{page1_taux_closure} {page1_closure_icon}</div>
                <small>Objectif: ≥ 13%</small>
            </div>
            <div class="kpi-card">
                <div class="kpi-label">Satisfaction Q1</div>
                <div class="kpi-value text-{page1_satisfaction_color}">{page1_taux_satisfaction} {page1_satisfaction_icon}</div>
                <small>Objectif: ≥ 92%</small>
            </div>
            <div class="kpi-card">
                <div class="kpi-label">Tickets boutiques</div>
                <div class="kpi-value">{page1_tickets_boutiques}</div>
                <small>Sur {page1_total_tickets} total</small>
            </div>
            <div class="kpi-card">
                <div class="kpi-label">Réponses Q1</div>
                <div class="kpi-value">{page1_total_q1}</div>
                <small>{page1_satisfaits} satisfaits, {page1_insatisfaits} insatisfaits</small>
            </div>
        </div>
        
        <h3>📊 Satisfaction Q1 par site</h3>
        <div class="pie-charts-grid">
            {page1_site_pie_charts}
        </div>
        
        <h3>🎫 Satisfaction par type d'enquête</h3>
        <div class="pie-charts-grid">
            {page1_ticket_type_charts}
        </div>
    </div>
    
    <div id="page2" class="tab-content">
        <h2 class="section-title">Page 2 - Tickets avec commentaires</h2>
        <div class="metric-highlight">
            <strong>{page2_comments_count} tickets avec commentaires</strong><br>
            Soit {page2_comments_percentage} des {page2_total_responses} enquêtes Q1 répondues
        </div>
        
        <h3>📊 Synthèse des commentaires par satisfaction</h3>
        <div class="chart-container">
            <img src="data:image/png;base64,{page2_pie_chart_comments}" alt="Synthèse commentaires">
        </div>
        
        <h3>🏢 Répartition par site</h3>
        <div class="chart-container">
            <img src="data:image/png;base64,{page2_pie_chart_sites}" alt="Répartition par site">
        </div>
        
        <h3>📊 Pourcentage commentaires par site</h3>
        <div class="table-wrapper">
            <table>
                <thead>
                    <tr>
                        <th>Site</th>
                        <th>Total enquêtes</th>
                        <th>Avec commentaires</th>
                        <th>% Commentaires</th>
                        <th>% Satisfaction avec commentaires</th>
                    </tr>
                </thead>
                <tbody>
                    {page2_site_comment_synthesis}
                </tbody>
            </table>
        </div>
        
        <h3>👥 Analyse par collaborateur</h3>
        <div class="search-container">
            <div class="filter-label">🔍 Filtrer l'analyse des collaborateurs</div>
            <input type="text" class="search-input" placeholder="Rechercher un collaborateur..." onkeyup="filterTable('collab-analysis-table', this.value)">
            <button class="clear-filters" onclick="clearAllFilters('page2-collab')">Effacer filtres</button>
        </div>
        
        <div class="table-wrapper">
            <table id="collab-analysis-table">
                <thead>
                    <tr>
                        <th onclick="sortTable('collab-analysis-table', 0)">Collaborateur</th>
                        <th onclick="sortTable('collab-analysis-table', 1)">Site</th>
                        <th onclick="sortTable('collab-analysis-table', 2)">Total commentaires</th>
                        <th onclick="sortTable('collab-analysis-table', 3)">Satisfaits</th>
                        <th onclick="sortTable('collab-analysis-table', 4)">Insatisfaits</th>
                        <th onclick="sortTable('collab-analysis-table', 5)">% Satisfait</th>
                        <th onclick="sortTable('collab-analysis-table', 6)">% Enquêtes avec commentaire</th>
                    </tr>
                </thead>
                <tbody>
                    {page2_collab_analysis_table}
                </tbody>
            </table>
            <div class="no-results" id="no-results-collab-analysis">
                Aucun résultat trouvé.
            </div>
        </div>
        
        <h3>💬 Liste des commentaires détaillés</h3>
        <div class="search-container">
            <div class="filter-label">🔍 Rechercher dans les commentaires</div>
            <input type="text" class="search-input" placeholder="Rechercher par code, boutique, site, collaborateur..." onkeyup="filterTable('comments-table', this.value)">
            <button class="clear-filters" onclick="clearAllFilters('page2')">Effacer filtres</button>
            <div class="column-filters">
                <input type="text" class="column-filter" placeholder="Code compte..." onkeyup="filterColumn('comments-table', 0, this.value)">
                <input type="text" class="column-filter" placeholder="Nom boutique..." onkeyup="filterColumn('comments-table', 1, this.value)">
                <input type="text" class="column-filter" placeholder="Site..." onkeyup="filterColumn('comments-table', 2, this.value)">
                <input type="text" class="column-filter" placeholder="Collaborateur..." onkeyup="filterColumn('comments-table', 3, this.value)">
                <input type="text" class="column-filter" placeholder="Type ticket..." onkeyup="filterColumn('comments-table', 4, this.value)">
                <input type="text" class="column-filter" placeholder="Numero ticket..." onkeyup="filterColumn('comments-table', 5, this.value)">
                <select class="column-filter" onchange="filterColumn('comments-table', 7, this.value)">
                    <option value="">Toutes satisfactions</option>
                    <option value="Très satisfaisant">Très satisfaisant</option>
                    <option value="Satisfaisant">Satisfaisant</option>
                    <option value="Peu satisfaisant">Peu satisfaisant</option>
                    <option value="Très peu satisfaisant">Très peu satisfaisant</option>
                </select>
                <input type="text" class="column-filter" placeholder="Commentaire..." onkeyup="filterColumn('comments-table', 10, this.value)">
            </div>
        </div>
        
        <div class="table-wrapper">
            <table id="comments-table">
                <thead>
                    <tr>
                        <th onclick="sortTable('comments-table', 0)">Code compte</th>
                        <th onclick="sortTable('comments-table', 1)">Nom boutique</th>
                        <th onclick="sortTable('comments-table', 2)">Site</th>
                        <th onclick="sortTable('comments-table', 3)">Collaborateur</th>
                        <th onclick="sortTable('comments-table', 4)">Type de ticket</th>
                        <th onclick="sortTable('comments-table', 5)">Numéro ticket</th>
                        <th onclick="sortTable('comments-table', 6)">Clos par</th>
                        <th onclick="sortTable('comments-table', 7)">Q1 Satisfaction</th>
                        <th onclick="sortTable('comments-table', 8)">Date création</th>
                        <th onclick="sortTable('comments-table', 9)">Date clôture</th>
                        <th onclick="sortTable('comments-table', 10)">Commentaire</th>
                    </tr>
                </thead>
                <tbody>
                    {page2_comments_table}
                </tbody>
            </table>
            <div class="no-results" id="no-results-comments">
                Aucun résultat trouvé.
            </div>
        </div>
    </div>
    
    <div id="page3" class="tab-content">
        <h2 class="section-title">Page 3 - Insatisfaits sans commentaires</h2>
        <div class="metric-highlight">
            <strong>{page3_count} tickets insatisfaits sans commentaire</strong>
        </div>
        
        <div class="search-container">
            <div class="filter-label">🔍 Rechercher dans les insatisfaits</div>
            <input type="text" class="search-input" placeholder="Rechercher par code, boutique, site, collaborateur..." onkeyup="filterTable('unsatisfied-table', this.value)">
            <button class="clear-filters" onclick="clearAllFilters('page3')">Effacer filtres</button>
            <div class="column-filters">
                <input type="text" class="column-filter" placeholder="Code compte..." onkeyup="filterColumn('unsatisfied-table', 0, this.value)">
                <input type="text" class="column-filter" placeholder="Nom boutique..." onkeyup="filterColumn('unsatisfied-table', 1, this.value)">
                <input type="text" class="column-filter" placeholder="Site..." onkeyup="filterColumn('unsatisfied-table', 2, this.value)">
                <input type="text" class="column-filter" placeholder="Collaborateur..." onkeyup="filterColumn('unsatisfied-table', 3, this.value)">
                <input type="text" class="column-filter" placeholder="Type ticket..." onkeyup="filterColumn('unsatisfied-table', 4, this.value)">
                <input type="text" class="column-filter" placeholder="Numero ticket..." onkeyup="filterColumn('unsatisfied-table', 5, this.value)">
                <select class="column-filter" onchange="filterColumn('unsatisfied-table', 7, this.value)">
                    <option value="">Toutes insatisfactions</option>
                    <option value="Peu satisfaisant">Peu satisfaisant</option>
                    <option value="Très peu satisfaisant">Très peu satisfaisant</option>
                </select>
                <input type="date" class="column-filter" onchange="filterColumn('unsatisfied-table', 8, this.value)">
            </div>
        </div>
        
        <div class="table-wrapper">
            <table id="unsatisfied-table">
                <thead>
                    <tr>
                        <th onclick="sortTable('unsatisfied-table', 0)">Code compte</th>
                        <th onclick="sortTable('unsatisfied-table', 1)">Nom boutique</th>
                        <th onclick="sortTable('unsatisfied-table', 2)">Site</th>
                        <th onclick="sortTable('unsatisfied-table', 3)">Collaborateur</th>
                        <th onclick="sortTable('unsatisfied-table', 4)">Type de ticket</th>
                        <th onclick="sortTable('unsatisfied-table', 5)">Numéro ticket</th>
                        <th onclick="sortTable('unsatisfied-table', 6)">Clos par</th>
                        <th onclick="sortTable('unsatisfied-table', 7)">Q1 Satisfaction</th>
                        <th onclick="sortTable('unsatisfied-table', 8)">Date création</th>
                        <th onclick="sortTable('unsatisfied-table', 9)">Date clôture</th>
                    </tr>
                </thead>
                <tbody>
                    {page3_table}
                </tbody>
            </table>
            <div class="no-results" id="no-results-unsatisfied">
                Aucun résultat trouvé.
            </div>
        </div>
    </div>
    
    <div id="page4" class="tab-content">
        <h2 class="section-title">Page 4 - Analyse globale</h2>
        
        <h3>📊 Répartition des réponses Q1</h3>
        <div class="table-wrapper">
            <table>
                <thead>
                    <tr>
                        <th>Satisfaction Q1</th>
                        <th>Nombre</th>
                    </tr>
                </thead>
                <tbody>
                    {page4_q1_table}
                </tbody>
            </table>
        </div>
        
        <h3>👥 Performance par collaborateur ({page4_collab_count} collaborateurs analysés)</h3>
        <div class="search-container">
            <div class="filter-label">🔍 Filtrer les collaborateurs</div>
            <input type="text" class="search-input" placeholder="Rechercher un collaborateur ou site..." onkeyup="filterTable('collaborators-table', this.value)">
            <button class="clear-filters" onclick="clearAllFilters('page4')">Effacer filtres</button>
            <div class="column-filters">
                <input type="text" class="column-filter" placeholder="Collaborateur..." onkeyup="filterColumn('collaborators-table', 0, this.value)">
                <input type="text" class="column-filter" placeholder="Site..." onkeyup="filterColumn('collaborators-table', 1, this.value)">
                <input type="number" class="column-filter" placeholder="Min très satisf..." onkeyup="filterNumberColumn('collaborators-table', 2, this.value, 'min')">
                <input type="number" class="column-filter" placeholder="Min satisf..." onkeyup="filterNumberColumn('collaborators-table', 3, this.value, 'min')">
                <input type="number" class="column-filter" placeholder="Max peu satisf..." onkeyup="filterNumberColumn('collaborators-table', 4, this.value, 'max')">
                <input type="number" class="column-filter" placeholder="Max très peu..." onkeyup="filterNumberColumn('collaborators-table', 5, this.value, 'max')">
                <input type="number" class="column-filter" placeholder="Min total..." onkeyup="filterNumberColumn('collaborators-table', 6, this.value, 'min')">
                <input type="number" class="column-filter" placeholder="Min %..." onkeyup="filterNumberColumn('collaborators-table', 7, this.value, 'min')">
                <input type="number" class="column-filter" placeholder="Min % comm..." onkeyup="filterNumberColumn('collaborators-table', 8, this.value, 'min')">
            </div>
        </div>
        
        <div class="table-wrapper">
            <table id="collaborators-table">
                <thead>
                    <tr>
                        <th onclick="sortTable('collaborators-table', 0)">Collaborateur</th>
                        <th onclick="sortTable('collaborators-table', 1)">Site</th>
                        <th onclick="sortTable('collaborators-table', 2)">Très satisf.</th>
                        <th onclick="sortTable('collaborators-table', 3)">Satisf.</th>
                        <th onclick="sortTable('collaborators-table', 4)">Peu satisf.</th>
                        <th onclick="sortTable('collaborators-table', 5)">Très peu satisf.</th>
                        <th onclick="sortTable('collaborators-table', 6)">Total</th>
                        <th onclick="sortTable('collaborators-table', 7)">% Satisfaction</th>
                        <th onclick="sortTable('collaborators-table', 8)">% Commentaires</th>
                        <th onclick="sortTable('collaborators-table', 9)">Taux retour</th>
                    </tr>
                </thead>
                <tbody>
                    {page4_collab_table}
                </tbody>
            </table>
            <div class="no-results" id="no-results-collaborators">
                Aucun résultat trouvé.
            </div>
        </div>
    </div>
    
    <div id="page5" class="tab-content">
        <h2 class="section-title">Page 5 - Boutiques ayant répondu</h2>
        <div class="metric-highlight">
            <strong>Objectif:</strong> ≥ 30% de boutiques ayant répondu<br>
            <strong>Indicateur global:</strong> <span class="text-{page5_global_objective_class}">{page5_global_response_rate}% {page5_global_objective_icon}</span>
        </div>
        
        <h3>📊 Synthèse par catégorie</h3>
        <div class="search-container">
            <div class="filter-label">🔍 Filtrer les catégories</div>
            <input type="text" class="search-input" placeholder="Rechercher une catégorie..." onkeyup="filterTable('category-table', this.value)">
            <button class="clear-filters" onclick="clearAllFilters('page5-category')">Effacer filtres</button>
        </div>
        
        <div class="table-wrapper">
            <table id="category-table">
                <thead>
                    <tr>
                        <th onclick="sortTable('category-table', 0)">Catégorie</th>
                        <th onclick="sortTable('category-table', 1)">Boutiques ouvertes</th>
                        <th onclick="sortTable('category-table', 2)">Dont ayant ouvert un ticket sur M</th>
                        <th onclick="sortTable('category-table', 3)">Boutiques répondues</th>
                        <th onclick="sortTable('category-table', 4)">% répondues</th>
                        <th onclick="sortTable('category-table', 5)">Jamais répondues</th>
                        <th onclick="sortTable('category-table', 6)">% jamais répondues</th>


                    </tr>
                </thead>
                <tbody>
                    {page5_category_table}
                </tbody>
            </table>
        </div>
        
        <h3>📊 Satisfaction par catégorie</h3>
        <div class="info-box">
            <p><strong>Ces graphiques montrent :</strong> Le taux de satisfaction pour chaque catégorie de boutique ayant répondu aux enquêtes Q1.</p>
            <p><strong>Calcul :</strong> Pour chaque catégorie, on présente le pourcentage de réponses satisfaisantes vs insatisfaisantes.</p>
            <p><strong>Objectif :</strong> Analyser la satisfaction par type de boutique.</p>
        </div>
        <div class="pie-charts-grid">
            {page5_category_pie_charts}
        </div>
        
        <h3>🏪 Liste des boutiques ayant répondu ({page5_responding_count})</h3>
        <div class="search-container">
            <div class="filter-label">🔍 Rechercher dans les boutiques</div>
            <input type="text" class="search-input" placeholder="Rechercher par code ou nom..." onkeyup="filterTable('shops-table', this.value)">
            <button class="clear-filters" onclick="clearAllFilters('page5-shops')">Effacer filtres</button>
            <div class="column-filters">
                <input type="text" class="column-filter" placeholder="Code compte..." onkeyup="filterColumn('shops-table', 0, this.value)">
                <input type="text" class="column-filter" placeholder="Nom compte..." onkeyup="filterColumn('shops-table', 1, this.value)">
                <select class="column-filter" onchange="filterColumn('shops-table', 2, this.value)">
                    <option value="">Toutes catégories</option>
                    <option value="400">400</option>
                    <option value="499">499</option>
                    <option value="993">993</option>
                    <option value="Mini-enseigne">Mini-enseigne</option>
                    <option value="Siège">Siège</option>
                </select>
            </div>
        </div>
        
        <div class="table-wrapper">
            <table id="shops-table">
                <thead>
                    <tr>
                        <th onclick="sortTable('shops-table', 0)">Code compte</th>
                        <th onclick="sortTable('shops-table', 1)">Compte</th>
                        <th onclick="sortTable('shops-table', 2)">Catégorie</th>
                    </tr>
                </thead>
                <tbody>
                    {page5_shops_table}
                </tbody>
            </table>
            <div class="no-results" id="no-results-shops">
                Aucun résultat trouvé.
            </div>
        </div>
    </div>
    
    <div id="page6" class="tab-content">
        <h2 class="section-title">Page 6 - Boutiques n'ayant jamais répondu</h2>
        <div class="metric-highlight">
            <strong>Objectif:</strong> ≤ 70% de boutiques n'ayant jamais répondu<br>
            <strong>Indicateur global:</strong> <span class="text-{page6_global_objective_class}">{page6_global_never_rate}% {page6_global_objective_icon}</span>
        </div>
        
        <div class="search-container">
            <div class="filter-label">🔍 Filtrer les boutiques jamais répondues</div>
            <input type="text" class="search-input" placeholder="Rechercher une catégorie..." onkeyup="filterTable('never-table', this.value)">
            <button class="clear-filters" onclick="clearAllFilters('page6')">Effacer filtres</button>
            <div class="column-filters">
                <input type="text" class="column-filter" placeholder="Catégorie..." onkeyup="filterColumn('never-table', 0, this.value)">
                <input type="number" class="column-filter" placeholder="Min ouvertes..." onkeyup="filterNumberColumn('never-table', 1, this.value, 'min')">
                <input type="number" class="column-filter" placeholder="Min jamais..." onkeyup="filterNumberColumn('never-table', 2, this.value, 'min')">
                <input type="number" class="column-filter" placeholder="Max %..." onkeyup="filterNumberColumn('never-table', 3, this.value, 'max')">
            </div>
        </div>
        
        <div class="table-wrapper">
            <table id="never-table">
                <thead>
                    <tr>
                        <th onclick="sortTable('never-table', 0)">Catégorie</th>
                        <th onclick="sortTable('never-table', 1)">Boutiques ouvertes</th>
                        <th onclick="sortTable('never-table', 2)">Dont ayant ouvert un ticket sur M</th>
                        <th onclick="sortTable('never-table', 3)">Jamais répondues</th>
                        <th onclick="sortTable('never-table', 4)">% jamais répondues</th>
                    </tr>
                </thead>
                <tbody>
                    {page6_category_table}
                </tbody>
            </table>
            <div class="no-results" id="no-results-never">
                Aucun résultat trouvé.
            </div>
        </div>
        
        <h3>🏪 Liste détaillée des boutiques n'ayant jamais répondu ({page6_never_count})</h3>
        <div class="search-container">
            <div class="filter-label">🔍 Rechercher dans les boutiques</div>
            <input type="text" class="search-input" placeholder="Rechercher par code ou nom..." onkeyup="filterTable('never-shops-table', this.value)">
            <button class="clear-filters" onclick="clearAllFilters('page6-shops')">Effacer filtres</button>
            <div class="column-filters">
                <input type="text" class="column-filter" placeholder="Code compte..." onkeyup="filterColumn('never-shops-table', 0, this.value)">
                <input type="text" class="column-filter" placeholder="Nom boutique..." onkeyup="filterColumn('never-shops-table', 1, this.value)">
                <select class="column-filter" onchange="filterColumn('never-shops-table', 2, this.value)">
                    <option value="">Toutes catégories</option>
                    <option value="400">400</option>
                    <option value="499">499</option>
                    <option value="993">993</option>
                    <option value="Mini-enseigne">Mini-enseigne</option>
                    <option value="Siège">Siège</option>
                </select>
            </div>
        </div>
        
        <div class="table-wrapper">
            <table id="never-shops-table">
                <thead>
                    <tr>
                        <th onclick="sortTable('never-shops-table', 0)">Code compte</th>
                        <th onclick="sortTable('never-shops-table', 1)">Nom boutique</th>
                        <th onclick="sortTable('never-shops-table', 2)">Catégorie</th>
                    </tr>
                </thead>
                <tbody>
                    {page6_never_table}
                </tbody>
            </table>
            <div class="no-results" id="no-results-never-shops">
                Aucun résultat trouvé.
            </div>
        </div>
    </div>
    
    <div id="page7" class="tab-content">
        <h2 class="section-title">Page 7 - Classement par boutique</h2>
        <div class="info-box">
            <p><strong>🏆 Critère de classement :</strong> Nombre total de tickets ouverts (ordre décroissant)</p>
            <p><strong>📊 Logique :</strong> Les boutiques ayant le plus de tickets clients sont classées en premier, indépendamment de leur taux de satisfaction aux enquêtes.</p>
        </div>
        <p><strong>Top {page7_total_shops} boutiques</strong> classées par volume d'activité (nombre de tickets)</p>
        
        <div class="search-container">
            <div class="filter-label">🔍 Rechercher dans le classement</div>
            <input type="text" class="search-input" placeholder="Rechercher par code, nom ou catégorie..." onkeyup="filterTable('ranking-table', this.value)">
            <button class="clear-filters" onclick="clearAllFilters('page7')">Effacer filtres</button>
            <div class="column-filters">
                <input type="text" class="column-filter" placeholder="Code compte..." onkeyup="filterColumn('ranking-table', 1, this.value)">
                <input type="text" class="column-filter" placeholder="Nom boutique..." onkeyup="filterColumn('ranking-table', 2, this.value)">
                <select class="column-filter" onchange="filterColumn('ranking-table', 3, this.value)">
                    <option value="">Toutes catégories</option>
                    <option value="400">400</option>
                    <option value="499">499</option>
                    <option value="993">993</option>
                    <option value="Mini-enseigne">Mini-enseigne</option>
                    <option value="Siège">Siège</option>
                </select>
                <input type="number" class="column-filter" placeholder="Min tickets..." onkeyup="filterNumberColumn('ranking-table', 4, this.value, 'min')">
                <input type="number" class="column-filter" placeholder="Min satisf..." onkeyup="filterNumberColumn('ranking-table', 9, this.value, 'min')">
            </div>
        </div>
        
        <div class="table-wrapper">
            <table id="ranking-table">
                <thead>
                    <tr>
                        <th onclick="sortTable('ranking-table', 0)">Rang</th>
                        <th onclick="sortTable('ranking-table', 1)">Code compte</th>
                        <th onclick="sortTable('ranking-table', 2)">Nom boutique</th>
                        <th onclick="sortTable('ranking-table', 3)">Catégorie</th>
                        <th onclick="sortTable('ranking-table', 4)">Tickets clos</th>
                        <th onclick="sortTable('ranking-table', 5)">% Taux retour</th>
                        <th onclick="sortTable('ranking-table', 6)">Enquêtes répondues</th>
                        <th onclick="sortTable('ranking-table', 7)">Enquêtes satisfaites</th>
                        <th onclick="sortTable('ranking-table', 8)">% Satisfaction</th>
                    </tr>
                </thead>
                <tbody>
                    {page7_ranking_table}
                </tbody>
            </table>
            <div class="no-results" id="no-results-ranking">
                Aucun résultat trouvé.
            </div>
        </div>
    </div>
    
    <script>
        function showTab(tabName) {{
            // Hide all tab contents
            document.querySelectorAll('.tab-content').forEach(content => {{
                content.classList.remove('active');
            }});
            // Remove active class from all tabs
            document.querySelectorAll('.tab').forEach(tab => {{
                tab.classList.remove('active');
            }});
            // Show selected tab content
            document.getElementById(tabName).classList.add('active');
            // Add active class to clicked tab
            event.target.classList.add('active');
        }}
        
        // Global search across all table columns
        function filterTable(tableId, searchValue) {{
            const table = document.getElementById(tableId);
            if (!table) return;
            
            const tbody = table.querySelector('tbody');
            const rows = tbody.querySelectorAll('tr');
            const noResultsDiv = document.getElementById('no-results-' + tableId.replace('-table', ''));
            let visibleRows = 0;
            
            searchValue = searchValue.toLowerCase().trim();
            
            rows.forEach(row => {{
                if (searchValue === '') {{
                    row.style.display = '';
                    visibleRows++;
                    return;
                }}
                
                const cells = row.querySelectorAll('td');
                let rowText = '';
                cells.forEach(cell => {{
                    rowText += ' ' + cell.textContent.toLowerCase();
                }});
                
                if (rowText.includes(searchValue)) {{
                    row.style.display = '';
                    visibleRows++;
                }} else {{
                    row.style.display = 'none';
                }}
            }});
            
            // Show/hide no results message
            if (noResultsDiv) {{
                noResultsDiv.style.display = visibleRows === 0 ? 'block' : 'none';
            }}
        }}
        
        // Column-specific filtering
        function filterColumn(tableId, columnIndex, filterValue) {{
            const table = document.getElementById(tableId);
            if (!table) return;
            
            const tbody = table.querySelector('tbody');
            const rows = tbody.querySelectorAll('tr');
            const noResultsDiv = document.getElementById('no-results-' + tableId.replace('-table', ''));
            let visibleRows = 0;
            
            filterValue = filterValue.toLowerCase().trim();
            
            rows.forEach(row => {{
                const cell = row.querySelectorAll('td')[columnIndex];
                if (!cell) return;
                
                const cellText = cell.textContent.toLowerCase().trim();
                
                if (filterValue === '' || cellText.includes(filterValue)) {{
                    // Check if row is hidden by other active filters
                    const otherFiltersActive = checkOtherActiveFilters(row, tableId, columnIndex, filterValue);
                    
                    if (!otherFiltersActive) {{
                        row.style.display = '';
                        visibleRows++;
                    }}
                }} else {{
                    row.style.display = 'none';
                }}
            }});
            
            if (noResultsDiv) {{
                noResultsDiv.style.display = visibleRows === 0 ? 'block' : 'none';
            }}
        }}
        
        // Numerical column filtering (min/max)
        function filterNumberColumn(tableId, columnIndex, filterValue, type) {{
            const table = document.getElementById(tableId);
            if (!table) return;
            
            const tbody = table.querySelector('tbody');
            const rows = tbody.querySelectorAll('tr');
            const noResultsDiv = document.getElementById('no-results-' + tableId.replace('-table', ''));
            let visibleRows = 0;
            
            const numValue = parseFloat(filterValue);
            if (isNaN(numValue) && filterValue.trim() !== '') return;
            
            rows.forEach(row => {{
                const cell = row.querySelectorAll('td')[columnIndex];
                if (!cell) return;
                
                const cellValue = parseFloat(cell.textContent.replace('%', '').replace(',', '.'));
                
                if (filterValue.trim() === '' || isNaN(numValue)) {{
                    row.style.display = '';
                    visibleRows++;
                }} else {{
                    let showRow = false;
                    if (type === 'min') {{
                        showRow = cellValue >= numValue;
                    }} else if (type === 'max') {{
                        showRow = cellValue <= numValue;
                    }}
                    
                    if (showRow) {{
                        row.style.display = '';
                        visibleRows++;
                    }} else {{
                        row.style.display = 'none';
                    }}
                }}
            }});
            
            if (noResultsDiv) {{
                noResultsDiv.style.display = visibleRows === 0 ? 'block' : 'none';
            }}
        }}
        
        // Table sorting functionality
        function sortTable(tableId, columnIndex) {{
            const table = document.getElementById(tableId);
            const tbody = table.querySelector('tbody');
            const rows = Array.from(tbody.querySelectorAll('tr'));
            
            // Determine if we should sort ascending or descending
            const header = table.querySelector('th:nth-child(' + (columnIndex + 1) + ')');
            const isAscending = !header.classList.contains('sort-desc');
            
            // Remove existing sort classes
            table.querySelectorAll('th').forEach(th => {{
                th.classList.remove('sort-asc', 'sort-desc');
            }});
            
            // Add appropriate sort class
            header.classList.add(isAscending ? 'sort-asc' : 'sort-desc');
            
            // Sort the rows
            rows.sort((a, b) => {{
                const aVal = a.querySelectorAll('td')[columnIndex].textContent.trim();
                const bVal = b.querySelectorAll('td')[columnIndex].textContent.trim();
                
                // Try to parse as numbers
                const aNum = parseFloat(aVal.replace('%', '').replace(',', '.'));
                const bNum = parseFloat(bVal.replace('%', '').replace(',', '.'));
                
                let comparison = 0;
                if (!isNaN(aNum) && !isNaN(bNum)) {{
                    comparison = aNum - bNum;
                }} else {{
                    comparison = aVal.localeCompare(bVal);
                }}
                
                return isAscending ? comparison : -comparison;
            }});
            
            // Re-append sorted rows
            rows.forEach(row => tbody.appendChild(row));
        }}
        
        // Check other active filters (simplified version)
        function checkOtherActiveFilters(row, tableId, currentColumnIndex, currentValue) {{
            return false; // Simplified for now
        }}
        
        // Clear all filters for a specific section
        function clearAllFilters(section) {{
            // Clear search inputs
            const searchInputs = document.querySelectorAll('#' + section + ' .search-input, #' + section + ' .column-filter');
            searchInputs.forEach(input => {{
                input.value = '';
            }});
            
            // Clear select dropdowns
            const selects = document.querySelectorAll('#' + section + ' select.column-filter');
            selects.forEach(select => {{
                select.value = '';
            }});
            
            // Show all table rows
            const tables = document.querySelectorAll('#' + section + ' table tbody tr');
            tables.forEach(row => {{
                row.style.display = '';
            }});
            
            // Hide no results messages
            const noResultsDivs = document.querySelectorAll('#' + section + ' .no-results');
            noResultsDivs.forEach(div => {{
                div.style.display = 'none';
            }});
        }}
        
        // Apply Q1 response colors when page loads
        document.addEventListener('DOMContentLoaded', function() {{
            document.querySelectorAll('td').forEach(cell => {{
                const text = cell.textContent.trim();
                if (text === 'Très satisfaisant') {{
                    cell.classList.add('q1-tres-satisfaisant');
                }} else if (text === 'Satisfaisant') {{
                    cell.classList.add('q1-satisfaisant');
                }} else if (text === 'Peu satisfaisant') {{
                    cell.classList.add('q1-peu-satisfaisant');
                }} else if (text === 'Très peu satisfaisant') {{
                    cell.classList.add('q1-tres-peu-satisfaisant');
                }}
            }});
            
            // Initialize all no-results as hidden
            document.querySelectorAll('.no-results').forEach(div => {{
                div.style.display = 'none';
            }});
        }});
    </script>
    
    <!-- Footer -->
    <footer style="background: #343a40; color: white; text-align: center; padding: 20px; margin-top: 40px;">
        <p style="margin-bottom: 10px;">Application conçue et développée par le pôle Q&T de la DESIR RCBT</p>
        <p style="margin: 0; font-size: 0.9em;">
            <a href="mailto:mflageul@rcbt.fr" style="color: #007bff; text-decoration: underline;">
                En cas d'évolutions ou anomalies, contactez-nous
            </a>
        </p>
    </footer>
</body>
</html>'''