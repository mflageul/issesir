import os
import tempfile
import pandas as pd
from datetime import datetime
from utils.optimized_html_generator import OptimizedReportGenerator
from utils.data_processor import DataProcessor

class IndividualReportGenerator(OptimizedReportGenerator):
    """Generate individual reports filtered by site or collaborator"""
    
    def __init__(self):
        super().__init__()
    
    def generate_individual_report(self, filtered_data, filter_type, filter_value):
        """Generate an individual HTML report for a specific site or collaborator"""
        
        print(f"=== GÉNÉRATION RAPPORT INDIVIDUEL: {filter_type.upper()} - {filter_value} ===")
        
        # Use the parent class's report generation logic
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        # Sanitize filename - nettoyer tous caractères de contrôle
        safe_value = filter_value.replace(' ', '_').replace('/', '_').replace('\\', '_')
        safe_value = safe_value.replace('\x0c', '').replace('\x0d', '').replace('\r', '').replace('\n', '')
        report_filename = f"rapport_individuel_{filter_type}_{safe_value}_{timestamp}.html"
        
        # Utiliser chemin temporaire simple et propre
        temp_dir = tempfile.gettempdir()
        report_path = os.path.join(temp_dir, report_filename)
        
        # Nettoyer le chemin final de tous caractères de contrôle
        report_path = report_path.replace('\x0c', '').replace('\x0d', '').replace('\r', '').replace('\n', '').strip()
        
        # Use DataProcessor to calculate metrics like in the main app
        processor = DataProcessor()
        
        # Calculate all pages data using filtered dataset - use CORRECTED method from parent class
        # Calculate page1 metrics first
        page1_metrics = processor.calculate_page1_metrics(filtered_data)
        page1_data = self._get_page1_synthesis(filtered_data, page1_metrics)
        page2_data = self._get_page2_comments_enhanced(filtered_data) 
        page3_data = self._get_page3_unsatisfied_no_comments(filtered_data)
        page4_data = self._get_page4_global_analysis_colored(filtered_data)
        
        # Use parent class methods for individual reports 
        page5_data = self._get_page5_responding_shops_colored(filtered_data)
        page6_data = self._get_page6_never_responding(filtered_data)
        
        # Store metrics for history saving
        self.last_page1_metrics = page1_data
        self.last_page2_metrics = page2_data
        # Use parent class method for page 7
        page7_data = self._get_page7_shop_ranking(filtered_data)
        
        # Customize template for individual report after getting the base template
        self._customize_template_for_individual(filter_type, filter_value)
        
        # Load logos
        import base64
        try:
            with open('static/images/logo_qt.jpg', 'rb') as f:
                logo_base64 = base64.b64encode(f.read()).decode('utf-8')
        except Exception as e:
            print(f"❌ Erreur chargement logo Q&T: {e}")
            logo_base64 = ""
        
        try:
            with open('static/images/logo_centre_services.png', 'rb') as f:
                logo_centre_base64 = base64.b64encode(f.read()).decode('utf-8')
        except Exception as e:
            print(f"❌ Erreur chargement logo centre: {e}")
            logo_centre_base64 = ""
        
        # Generate HTML content
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
        
        # Write report
        with open(report_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        print(f"✅ Rapport individuel généré: {report_path}")
        return report_path
    
    def _customize_template_for_individual(self, filter_type, filter_value):
        """Customize the HTML template for individual reports"""
        # Modify the title to reflect the individual nature
        filter_display = "Site" if filter_type == "site" else "Collaborateur"
        self.template = self.template.replace(
            "Rapport RCBT Optimisé",
            f"Rapport RCBT - {filter_display}: {filter_value}"
        )
        
        # Add a filter indicator at the top
        filter_indicator = f'''
        <div style="background: #e3f2fd; border: 1px solid #2196f3; padding: 10px; margin: 20px 0; border-radius: 5px; text-align: center;">
            <strong>📋 Rapport filtré par {filter_display.lower()}: {filter_value}</strong>
        </div>
        '''
        
        # Insert after the header div - be more flexible with the replacement
        if '</div>\n    \n    <div class="nav-tabs">' in self.template:
            self.template = self.template.replace(
                '</div>\n    \n    <div class="nav-tabs">',
                f'</div>{filter_indicator}\n    \n    <div class="nav-tabs">'
            )
        elif '<div class="nav-tabs">' in self.template:
            self.template = self.template.replace(
                '<div class="nav-tabs">',
                f'{filter_indicator}\n    <div class="nav-tabs">'
            )
    
    def _get_page1_synthesis_data(self, filtered_data, filter_type, filter_value):
        """Calculate Page 1 data for individual report - CORRIGÉ selon spécifications"""
        import pandas as pd
        
        df_merged = filtered_data['merged']
        df_case = filtered_data['case']
        
        print(f"📊 Page 1 individuelle CORRIGÉE: {filter_type} = {filter_value}")
        
        # CORRECTION FINALE: df_case est déjà filtré dans filter_data_for_individual_report
        # Donc context_tickets = df_case (qui est déjà le contexte filtré)
        context_tickets = df_case
        print(f"   - Tickets dans le contexte {filter_type} = {filter_value}: {len(context_tickets)} lignes")
        
        # CORRECTION FINALE Page 1: "Tickets boutiques" = nombre de tickets sans doublons du fichier case dans le contexte
        tickets_boutiques_contexte = context_tickets['Numéro'].nunique()  # Tickets sans doublons dans le contexte
        # Pour l'affichage, tickets_total = tickets du contexte (pas global)
        tickets_total = tickets_boutiques_contexte  # Dans le contexte individuel, total = contexte
        
        # CORRECTION FINALE Page 1: "Taux de clôture boutiques" = enquêtes répondues / tickets boutiques contexte
        context_ticket_nums = set(context_tickets['Numéro'].unique())
        q1_data = df_merged[df_merged['Mesure'].str.contains('Q1', na=False)]
        enquetes_repondues_contexte = q1_data[q1_data['Dossier Rcbt'].isin(context_ticket_nums)]
        nb_enquetes_repondues = enquetes_repondues_contexte['Dossier Rcbt'].nunique()  # Enquêtes sans doublons
        
        # Taux de clôture = enquêtes répondues / tickets boutiques dans le contexte
        taux_cloture_boutiques = (nb_enquetes_repondues / tickets_boutiques_contexte * 100) if tickets_boutiques_contexte > 0 else 0
        
        print(f"   - Tickets boutiques (contexte): {tickets_boutiques_contexte}/{tickets_total}")
        print(f"   - Enquêtes répondues contexte: {nb_enquetes_repondues}")
        print(f"   - Taux clôture boutiques: {taux_cloture_boutiques:.1f}%")
        
        # CORRECTION 2-3: CALCUL SATISFACTION Q1 dans le contexte - MÉTHODOLOGIE CORRIGÉE
        if len(enquetes_repondues_contexte) > 0:
            # Chercher la colonne de satisfaction dans les données Q1
            satisfaction_data = enquetes_repondues_contexte[enquetes_repondues_contexte['Valeur de chaîne'].notna()]
            
            if len(satisfaction_data) > 0:
                # Compter les satisfaits (Très satisfaisant + Satisfaisant)
                satisfaits_contexte = satisfaction_data['Valeur de chaîne'].isin(['Très satisfaisant', 'Satisfaisant']).sum()
                # Compter les insatisfaits (tout le reste des réponses)
                insatisfaits_contexte = len(satisfaction_data) - satisfaits_contexte
                # Taux = satisfaits / total réponses Q1
                taux_satisfaction = (satisfaits_contexte / len(satisfaction_data) * 100) if len(satisfaction_data) > 0 else 0
                
                print(f"   🎯 SATISFACTION CORRIGÉE: {satisfaits_contexte} satisfaits / {len(satisfaction_data)} réponses = {taux_satisfaction:.1f}%")
                print(f"   📊 Répartition: {satisfaits_contexte} satisfaits, {insatisfaits_contexte} insatisfaits")
            else:
                satisfaits_contexte = 0
                insatisfaits_contexte = 0
                taux_satisfaction = 0
        else:
            satisfaits_contexte = 0
            insatisfaits_contexte = 0
            taux_satisfaction = 0

        # Toutes les variables nécessaires pour le template Page 1
        closure_ok = taux_cloture_boutiques >= 13
        sat_ok = taux_satisfaction >= 92
        
        return {
            'total_tickets': tickets_total,
            'page1_total_tickets': tickets_total,  # Variable manquante ajoutée
            'page1_total_q1': nb_enquetes_repondues,  # Variable Q1 manquante ajoutée
            'tickets_boutiques': tickets_boutiques_contexte,  # CORRIGÉ: tickets sans doublons du contexte
            'page1_tickets_boutiques': tickets_boutiques_contexte,  # CORRIGÉ: tickets sans doublons du contexte  
            'tickets_boutiques_display': f"{tickets_boutiques_contexte}/{tickets_total}",
            'tickets_system': tickets_total - tickets_boutiques_contexte,
            'taux_cloture_boutiques': f"{taux_cloture_boutiques:.1f}%",
            'taux_closure': taux_cloture_boutiques,
            'enquetes_repondues_contexte': nb_enquetes_repondues,
            
            # Variables pour les couleurs et icônes (format template)
            'page1_closure_color': 'success' if closure_ok else 'warning',
            'page1_closure_icon': '✅' if closure_ok else '⚠️',
            'page1_taux_closure': f"{taux_cloture_boutiques:.1f}%",
            'page1_taux_satisfaction': f"{taux_satisfaction:.1f}%",
            'closure_ok': closure_ok,
            
            # Variables satisfaction (calculées selon les enquêtes contexte)
            'satisfaits': satisfaits_contexte,
            'page1_satisfaits': satisfaits_contexte,  # Variable manquante ajoutée
            'insatisfaits': insatisfaits_contexte,
            'page1_insatisfaits': insatisfaits_contexte,  # Variable manquante ajoutée
            'taux_sat': taux_satisfaction,
            'sat_ok': sat_ok,
            'page1_sat_color': 'success' if sat_ok else 'warning',
            'page1_sat_icon': '✅' if sat_ok else '⚠️',
            'page1_satisfaction_color': 'success' if sat_ok else 'warning',
            'page1_satisfaction_icon': '✅' if sat_ok else '⚠️',
            
            # CORRECTION 4: Variables graphiques avec génération camemberts
            'page1_site_pie_charts': self._generate_page1_pie_charts(filtered_data, filter_type, filter_value),
            'page1_satisfaction_chart': '',  # Variable graphique manquante ajoutée
            'page1_response_chart': '',  # Variable graphique manquante ajoutée
            'page1_ticket_type_charts': '',  # Variable graphique manquante ajoutée
            'page1_metrics': f"<strong>{satisfaits_contexte}</strong> satisfaits, <strong>{insatisfaits_contexte}</strong> insatisfaits",
            'page1_taux_cloture': f"{taux_cloture_boutiques:.1f}%",  # Variable duplicata nécessaire
            'page1_taux_cloture_icon': '✅' if closure_ok else '⚠️',  # Variable icône manquante
            'page1_taux_cloture_status': 'Objectif atteint' if closure_ok else 'Objectif non atteint',  # Variable status manquante
            'page1_taux_satisfaction_icon': '✅' if sat_ok else '⚠️',  # Variable icône manquante
            'page1_taux_satisfaction_status': 'Objectif atteint' if sat_ok else 'Objectif non atteint'  # Variable status manquante
        }
    
    def _generate_page1_pie_charts(self, filtered_data, filter_type, filter_value):
        """CORRECTION 4: Génère les camemberts pour la page 1 des rapports individuels"""
        try:
            from utils.visualizations import VisualizationGenerator
            viz_gen = VisualizationGenerator()
            
            # Générer camembert satisfaction pour ce contexte
            df_merged = filtered_data['merged']
            
            # CORRECTION: Chercher données Q1 satisfaction dans les bonnes colonnes
            if 'Valeur de chaîne' in df_merged.columns:
                satisfaction_data = df_merged[df_merged['Valeur de chaîne'].notna()]
                
                if len(satisfaction_data) > 0:
                    # Préparer distribution satisfaction contextualisée
                    satisfaction_counts = satisfaction_data['Valeur de chaîne'].value_counts()
                    total = len(satisfaction_data)
                    
                    distribution = {}
                    for satisfaction, count in satisfaction_counts.items():
                        percentage = round((count / total) * 100, 1)
                        distribution[str(satisfaction)] = {
                            'count': int(count),
                            'percentage': percentage
                        }
                    
                    satisfaction_chart = viz_gen.create_satisfaction_pie_chart(distribution)
                    if satisfaction_chart:
                        return f'<div class="text-center mb-3"><img src="data:image/png;base64,{satisfaction_chart}" alt="Répartition Satisfaction" class="img-fluid" style="max-width: 400px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);"></div>'
            
            return '<div class="text-center text-muted p-3" style="background: #f8f9fa; border-radius: 8px;">Aucune donnée de satisfaction disponible pour ce filtre</div>'
                
        except Exception as e:
            print(f"❌ Erreur génération camembert page 1: {e}")
            import traceback
            traceback.print_exc()
            return '<div class="text-center text-warning p-3" style="background: #fff3cd; border-radius: 8px;">Erreur génération graphique</div>'
    
    def _generate_page5_pie_charts(self, filtered_data, filter_type, filter_value):
        """CORRECTION: Génère les camemberts pour la page 5 des rapports individuels"""
        try:
            from utils.visualizations import VisualizationGenerator
            viz_gen = VisualizationGenerator()
            
            df_merged = filtered_data['merged']
            df_accounts = filtered_data['accounts']
            
            # Analyser distribution par catégorie des boutiques répondantes
            if len(df_merged) > 0 and 'Code compte' in df_merged.columns:
                # Merger avec comptes pour avoir catégories
                merged_with_accounts = df_merged.merge(
                    df_accounts[['Code compte', 'Catégorie']], 
                    on='Code compte', 
                    how='left'
                )
                
                if len(merged_with_accounts) > 0:
                    category_counts = merged_with_accounts['Catégorie'].value_counts()
                    total = len(merged_with_accounts)
                    
                    # Préparer distribution catégories
                    category_distribution = {}
                    for category, count in category_counts.items():
                        if pd.notna(category):
                            percentage = round((count / total) * 100, 1)
                            category_distribution[str(category)] = {
                                'count': int(count),
                                'percentage': percentage
                            }
                    
                    if category_distribution:
                        category_chart = viz_gen.create_category_pie_chart(category_distribution)
                        if category_chart:
                            return f'<div class="text-center mb-3"><img src="data:image/png;base64,{category_chart}" alt="Répartition par Catégorie" class="img-fluid" style="max-width: 400px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);"></div>'
            
            return '<div class="text-center text-muted p-3" style="background: #f8f9fa; border-radius: 8px;">Aucune donnée de catégorie disponible</div>'
                
        except Exception as e:
            print(f"❌ Erreur génération camembert page 5: {e}")
            import traceback
            traceback.print_exc()
            return '<div class="text-center text-warning p-3" style="background: #fff3cd; border-radius: 8px;">Erreur génération graphique page 5</div>'
    
    def _find_satisfaction_column(self, df):
        """Trouve la colonne de satisfaction dans les données Q1"""
        satisfaction_columns = ['Satisfaction', 'Q1', 'Réponse', 'Response']
        for col in satisfaction_columns:
            if col in df.columns:
                return col
        # Si aucune colonne spécifique trouvée, chercher par contenu
        for col in df.columns:
            if df[col].dtype == 'object':
                unique_vals = df[col].dropna().unique()
                if any('satisf' in str(val).lower() for val in unique_vals):
                    return col
        return None

    def _get_individual_page5_responding_shops(self, data, filter_type, filter_value):
        """Page 5 adaptée pour rapports individuels - CALCULS CORRIGÉS"""
        import pandas as pd
        
        df_merged = data['merged']
        df_accounts = data['accounts']
        df_case = data['case']
        
        print(f"🔍 Page 5 individuelle CORRIGÉE: {filter_type} = {filter_value}")
        
        # CORRECTION MAJEURE: Adapter les calculs au contexte individuel
        # Pour collaborateur: tickets ouverts PAR le collaborateur
        # Pour site: tickets ouverts PAR des collaborateurs du site
        
        if filter_type == 'collaborator':
            # Tickets créés par ce collaborateur SEULEMENT
            context_tickets = df_case[df_case['Créé par ticket'] == filter_value]
            print(f"   - Tickets créés par collaborateur {filter_value}: {len(context_tickets)}")
            
            # Enquêtes Q1 correspondant aux tickets de ce collaborateur
            context_ticket_nums = set(context_tickets['Numéro'].unique())
            q1_data = df_merged[
                (df_merged['Mesure'].str.contains('Q1', na=False)) & 
                (df_merged['Dossier Rcbt'].isin(context_ticket_nums))
            ]
            
        elif filter_type == 'site':
            # Assurer le mapping des sites si nécessaire
            if 'Site' not in df_case.columns:
                df_ref = data.get('ref', pd.DataFrame())
                if 'Login' in df_ref.columns and 'Site' in df_ref.columns:
                    site_mapping = dict(zip(df_ref['Login'], df_ref['Site']))
                    df_case['Site'] = df_case['Créé par ticket'].map(site_mapping)
                    df_case['Site'] = df_case['Site'].fillna('Autres')
            
            # Tickets créés par des collaborateurs de ce site
            context_tickets = df_case[df_case['Site'] == filter_value] if 'Site' in df_case.columns else df_case
            print(f"   - Tickets créés par site {filter_value}: {len(context_tickets)}")
            
            # Enquêtes Q1 correspondant aux tickets de ce site
            context_ticket_nums = set(context_tickets['Numéro'].unique())
            q1_data = df_merged[
                (df_merged['Mesure'].str.contains('Q1', na=False)) & 
                (df_merged['Dossier Rcbt'].isin(context_ticket_nums))
            ]
        else:
            context_tickets = df_case
            q1_data = df_merged[df_merged['Mesure'].str.contains('Q1', na=False)]
        
        # Boutiques ayant ouvert des tickets dans ce contexte
        context_boutiques_with_tickets = context_tickets['Code compte'].unique()
        print(f"   - Boutiques dans contexte: {len(context_boutiques_with_tickets)}")
        
        # Boutiques ayant répondu aux enquêtes dans ce contexte
        responding_shops_context = q1_data['Code compte'].unique()
        print(f"   - Boutiques ayant répondu dans contexte: {len(responding_shops_context)}")
        
        # Category analysis with context-aware calculations
        category_stats = []
        for category in ['400', '499', '993', 'Mini-enseigne', 'Siège']:
            # Boutiques de cette catégorie
            cat_accounts = df_accounts[df_accounts['Categorie'] == category]
            
            # DANS LE CONTEXTE: Boutiques qui ont ouvert tickets avec ce collaborateur/site
            cat_with_tickets_context = cat_accounts[cat_accounts['Code compte'].isin(context_boutiques_with_tickets)]
            total_with_tickets_context = len(cat_with_tickets_context)
            
            # DANS LE CONTEXTE: Boutiques qui ont répondu aux enquêtes
            cat_responded_context = cat_with_tickets_context[cat_with_tickets_context['Code compte'].isin(responding_shops_context)]
            responded = len(cat_responded_context)
            
            # Jamais répondues dans le contexte
            never_responded = total_with_tickets_context - responded
            
            # Calculs des pourcentages basés sur le contexte
            response_rate = (responded / total_with_tickets_context * 100) if total_with_tickets_context > 0 else 0
            never_rate = (never_responded / total_with_tickets_context * 100) if total_with_tickets_context > 0 else 0
            
            # Total boutiques dans la catégorie (pour information)
            total_open_all = len(cat_accounts)
            
            objective_response_ok = response_rate >= 30
            objective_never_ok = never_rate <= 70
            
            category_stats.append({
                'category': category,
                'total_open': total_open_all,
                'total_with_tickets': total_with_tickets_context,  # CORRIGÉ: basé sur contexte
                'responded': responded,  # CORRIGÉ: basé sur contexte
                'response_rate': response_rate,
                'never_responded': never_responded,
                'never_rate': never_rate,
                'objective_response_ok': objective_response_ok,
                'objective_never_ok': objective_never_ok
            })
            
            print(f"   - Catégorie {category}: {responded}/{total_with_tickets_context} réponses ({response_rate:.1f}%)")
        
        # Build category table with corrected headers
        category_table = ""
        for stat in category_stats:
            response_class = 'success' if stat['objective_response_ok'] else 'warning'
            never_class = 'success' if stat['objective_never_ok'] else 'warning'
            
            category_table += f"""
            <tr>
                <td>{stat['category']}</td>
                <td>{stat['total_open']}</td>
                <td><strong>{stat['total_with_tickets']}</strong></td>
                <td><strong>{stat['responded']}</strong></td>
                <td class="text-{response_class}"><strong>{stat['response_rate']:.1f}%</strong></td>
                <td>{stat['never_responded']}</td>
                <td class="text-{never_class}">{stat['never_rate']:.1f}%</td>
            </tr>"""
        
        # CORRECTION: Boutiques ayant répondu dans le contexte avec bonnes données
        context_responding_accounts = df_accounts[
            df_accounts['Code compte'].isin(responding_shops_context)
        ]
        
        shops_table = ""
        for _, row in context_responding_accounts.iterrows():
            code_compte = int(float(row.get('Code compte', 0))) if pd.notna(row.get('Code compte')) else 'N/A'
            
            # CORRECTION: Get account name from correct column avec vérifications
            account_name = 'N/A'
            for name_col in ['Compte', 'Account', 'Nom', 'Nom boutique', 'Shop Name']:
                if name_col in row.index and pd.notna(row.get(name_col)):
                    name_value = str(row.get(name_col)).strip()
                    if name_value and name_value != 'nan':
                        account_name = name_value
                        break
            
            # CORRECTION: Get category with proper verification
            category_name = 'N/A'
            for cat_col in ['Categorie', 'Category', 'Type', 'Catégorie']:
                if cat_col in row.index and pd.notna(row.get(cat_col)):
                    cat_value = str(row.get(cat_col)).strip()
                    if cat_value and cat_value != 'nan':
                        category_name = cat_value
                        break
            
            shops_table += f"""
            <tr>
                <td>{code_compte}</td>
                <td>{account_name}</td>
                <td>{category_name}</td>
            </tr>"""
        
        # Global response rate (contextualisé)
        total_context_shops = len(context_boutiques_with_tickets)
        total_responding_context = len(context_responding_accounts)
        global_response_rate = (total_responding_context / total_context_shops * 100) if total_context_shops > 0 else 0
        
        return {
            'page5_category_table': category_table,
            'page5_shops_table': shops_table,
            'page5_responding_count': total_responding_context,
            'page5_category_pie_charts': self._generate_page5_pie_charts(filtered_data, filter_type, filter_value),  # CORRECTION: Camemberts page 5 ajoutés
            'page5_global_response_rate': f"{global_response_rate:.1f}",
            'page5_global_objective_class': 'success' if global_response_rate >= 30 else 'warning',
            'page5_global_objective_icon': '✅' if global_response_rate >= 30 else '⚠️'
        }
    
    def _get_individual_page6_never_responding(self, data, filter_type, filter_value):
        """Page 6 adaptée pour rapports individuels"""
        import pandas as pd
        
        df_merged = data['merged']
        df_accounts = data['accounts']
        df_case = data['case']
        
        # Get responding shops
        q1_data = df_merged[df_merged['Mesure'].str.contains('Q1', na=False)]
        responding_shops = q1_data['Code compte'].unique()
        
        # Context filtering like Page 5
        if filter_type == 'collaborator':
            context_tickets = df_case[df_case['Créé par ticket'] == filter_value]
        elif filter_type == 'site':
            context_tickets = df_case[df_case['Site'] == filter_value] if 'Site' in df_case.columns else df_case
        else:
            context_tickets = df_case
            
        context_boutiques_with_tickets = context_tickets['Code compte'].unique()
        
        # Category table for never responding shops
        category_table = ""
        never_shops = []
        
        for category in ['400', '499', '993', 'Mini-enseigne', 'Siège']:
            cat_accounts = df_accounts[df_accounts['Categorie'] == category]
            cat_with_tickets_context = cat_accounts[cat_accounts['Code compte'].isin(context_boutiques_with_tickets)]
            
            # Shops that never responded (in context)
            never_responded_cat = cat_with_tickets_context[~cat_with_tickets_context['Code compte'].isin(responding_shops)]
            
            total_open_all = len(cat_accounts)
            total_with_tickets_context = len(cat_with_tickets_context)
            never_count = len(never_responded_cat)
            never_rate = (never_count / total_with_tickets_context * 100) if total_with_tickets_context > 0 else 0
            
            never_class = 'success' if never_rate <= 70 else 'warning'
            
            category_table += f"""
            <tr>
                <td>{category}</td>
                <td>{total_open_all}</td>
                <td>{total_with_tickets_context}</td>
                <td>{never_count}</td>
                <td class="text-{never_class}">{never_rate:.1f}%</td>
            </tr>"""
            
            # Add to never shops list
            for _, row in never_responded_cat.iterrows():
                never_shops.append(row)
        
        # Detailed never responding shops table
        never_shops_table = ""
        for shop in never_shops:
            code_compte = int(float(shop.get('Code compte', 0))) if pd.notna(shop.get('Code compte')) else 'N/A'
            account_name = 'N/A'
            for name_col in ['Compte', 'Account', 'Nom']:
                if name_col in shop.index and pd.notna(shop.get(name_col)) and str(shop.get(name_col)).strip():
                    account_name = shop.get(name_col)
                    break
            
            never_shops_table += f"""
            <tr>
                <td>{code_compte}</td>
                <td>{account_name}</td>
                <td>{shop.get('Categorie', 'N/A')}</td>
            </tr>"""
        
        total_never_rate = (len(never_shops) / len(context_boutiques_with_tickets) * 100) if len(context_boutiques_with_tickets) > 0 else 0
        
        return {
            'page6_category_table': category_table,
            'page6_never_shops_table': never_shops_table,
            'page6_never_table': never_shops_table,  # Alias for template compatibility
            'page6_never_count': len(never_shops),
            'page6_global_never_rate': f"{total_never_rate:.1f}",
            'page6_global_objective_class': 'success' if total_never_rate <= 70 else 'warning',
            'page6_global_objective_icon': '✅' if total_never_rate <= 70 else '⚠️'
        }
        
        # Count boutique tickets (same logic as global report but with filtered data)
        boutique_tickets = len(df_merged[df_merged['Boutique_categorie'] != 'Autres'])
        
        # Calculate closure rate
        closed_tickets = len(df_merged[df_merged['Clos par'].notna()])
        taux_closure = round(closed_tickets / total_tickets * 100, 1) if total_tickets > 0 else 0.0
        
        # Calculate satisfaction from Q1 responses
        q1_responses = df_merged[df_merged['Valeur de chaîne'].notna()]
        satisfied = len(q1_responses[q1_responses['Valeur de chaîne'].isin(['Très satisfaisant', 'Satisfaisant'])])
        total_responses = len(q1_responses)
        taux_satisfaction = round(satisfied / total_responses * 100, 1) if total_responses > 0 else 0.0
        
        # Status indicators
        closure_ok = taux_closure >= 13.0
        sat_ok = taux_satisfaction >= 92.0
        
        return {
            'page1_total_tickets': total_tickets,
            'page1_tickets_boutiques': boutique_tickets,
            'page1_taux_closure': f"{taux_closure}%",
            'page1_taux_satisfaction': f"{taux_satisfaction}%",
            'page1_closure_color': 'success' if closure_ok else 'warning',
            'page1_satisfaction_color': 'success' if sat_ok else 'warning',
            'page1_closure_icon': '✅' if closure_ok else '⚠️',
            'page1_satisfaction_icon': '✅' if sat_ok else '⚠️'
        }