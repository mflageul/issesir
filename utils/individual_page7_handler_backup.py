"""
Gestionnaire spécialisé pour la Page 7 des rapports individuels
Adapte les calculs selon le collaborateur ou site sélectionné
"""

import pandas as pd

class IndividualPage7Handler:
    def __init__(self, filter_type, filter_value):
        self.filter_type = filter_type  # 'site' or 'collaborator'
        self.filter_value = filter_value
    
    def get_individual_shop_ranking(self, data):
        """Page 7 adaptée pour rapports individuels"""
        df_merged = data['merged']
        df_case = data['case']
        df_accounts = data['accounts']
        
        print(f"🔍 Page 7 individuelle: {self.filter_type} = {self.filter_value}")
        
        if self.filter_type == 'collaborator':
            return self._get_collaborator_shop_ranking(data)
        elif self.filter_type == 'site':
            return self._get_site_shop_ranking(data)
        else:
            # Fallback to global ranking
            return self._get_global_shop_ranking(data)
    
    def _get_collaborator_shop_ranking(self, data):
        """Page 7 pour un collaborateur spécifique - LOGIQUE CORRIGÉE"""
        import pandas as pd
        df_merged = data['merged']
        df_case = data['case']
        df_accounts = data['accounts']
        collaborator = self.filter_value
        
        print(f"📊 Calculs Page 7 CORRIGÉS pour collaborateur: {collaborator}")
        
        # NOUVELLE LOGIQUE: Utiliser les tickets CRÉÉS par le collaborateur
        # Car un collaborateur "possède" les tickets qu'il a créés
        collaborator_tickets = df_case[df_case['Créé par ticket'] == collaborator]
        print(f"   - Tickets créés par {collaborator}: {len(collaborator_tickets)}")
        
        if len(collaborator_tickets) == 0:
            print(f"   ⚠️ Aucun ticket créé par {collaborator}")
            return {
                'page7_shop_ranking_table': '<tr><td colspan="9">Aucun ticket créé par ce collaborateur</td></tr>',
                'page7_ranking_table': '<tr><td colspan="9">Aucun ticket créé par ce collaborateur</td></tr>',
                'page7_total_shops': 0
            }
        
        # 1. STATS PAR BOUTIQUE basées sur les tickets créés par le collaborateur
        q1_data = df_merged[df_merged['Mesure'].str.contains('Q1', na=False)]
        shop_stats = []
        
        for code_compte in collaborator_tickets['Code compte'].unique():
            # Tickets de cette boutique créés par le collaborateur
            boutique_tickets = collaborator_tickets[collaborator_tickets['Code compte'] == code_compte]
            total_tickets = len(boutique_tickets)
            tickets_clos = len(boutique_tickets[boutique_tickets['État'] == 'Clos'])
            taux_closure = (tickets_clos / total_tickets * 100) if total_tickets > 0 else 0
            
            # Réponses Q1 pour cette boutique sur les tickets de ce collaborateur
            ticket_nums = set(boutique_tickets['Numéro'].unique())
            q1_responses = q1_data[
                (q1_data['Code compte'] == code_compte) & 
                (q1_data['Dossier Rcbt'].isin(ticket_nums))
            ]
            
            q1_count = len(q1_responses)
            taux_retour = (q1_count / total_tickets * 100) if total_tickets > 0 else 0
            
            # Satisfaction sur ces réponses
            if len(q1_responses) > 0:
                satisfaits = q1_responses['Valeur de chaîne'].isin(['Très satisfaisant', 'Satisfaisant']).sum()
                satisfaction_rate = (satisfaits / len(q1_responses) * 100)
            else:
                satisfaits = 0
                satisfaction_rate = 0
            
            # Info boutique
            shop_info = df_accounts[df_accounts['Code compte'] == code_compte]
            nom_boutique = shop_info.iloc[0].get('Compte', 'N/A') if len(shop_info) > 0 else 'N/A'
            categorie = shop_info.iloc[0].get('Categorie', 'N/A') if len(shop_info) > 0 else 'N/A'
            
            shop_stats.append({
                'Code compte': int(code_compte),
                'Nom boutique': nom_boutique,
                'Catégorie': categorie,
                'Total_tickets': total_tickets,
                'Tickets_clos': tickets_clos,
                'Taux_closure': taux_closure,
                'Q1_responses': q1_count,
                'Satisfaits': satisfaits,
                'Satisfaction_rate': satisfaction_rate,
                'Taux_retour': taux_retour
            })
            
            print(f"   - Boutique {code_compte}: {q1_count}/{total_tickets} réponses ({taux_retour:.1f}%), satisfaction {satisfaction_rate:.1f}%")
        
        # Trier par nombre de tickets décroissant
        shop_stats.sort(key=lambda x: x['Total_tickets'], reverse=True)
        
        # Construire le tableau HTML
        shop_ranking_table = ""
        for shop in shop_stats:
            # Classes CSS pour les couleurs
            closure_class = 'success' if shop['Taux_closure'] >= 13 else ('warning' if shop['Taux_closure'] >= 10 else 'danger')
            return_class = 'success' if shop['Taux_retour'] >= 15 else ('warning' if shop['Taux_retour'] >= 10 else 'danger')
            satisfaction_class = 'success' if shop['Satisfaction_rate'] >= 92 else ('warning' if shop['Satisfaction_rate'] >= 85 else 'danger')
            
            shop_ranking_table += f"""
            <tr>
                <td>{shop['Code compte']}</td>
                <td>{shop['Nom boutique']}</td>
                <td>{shop['Catégorie']}</td>
                <td>{shop['Total_tickets']}</td>
                <td class="text-{closure_class}">{shop['Taux_closure']:.1f}%</td>
                <td>{shop['Q1_responses']}</td>
                <td>{shop['Satisfaits']}</td>
                <td class="text-{satisfaction_class}">{shop['Satisfaction_rate']:.1f}%</td>
                <td class="text-{return_class}">{shop['Taux_retour']:.1f}%</td>
            </tr>"""
        
        return {
            'page7_shop_ranking_table': shop_ranking_table,
            'page7_ranking_table': shop_ranking_table,  # Alias for template compatibility
            'page7_total_shops': len(shop_stats)
        }

    
    def _get_site_shop_ranking(self, data):
        """Page 7 pour un site spécifique - CORRIGÉE pour contexte"""
        import pandas as pd
        
        df_merged = data['merged']
        df_case = data['case']
        df_accounts = data['accounts']
        site = self.filter_value
        
        print(f"📊 Calculs Page 7 CORRIGÉS pour site: {site}")
        
        # Assurer le mapping des sites si nécessaire
        if 'Site' not in df_case.columns:
            df_ref = data.get('ref', pd.DataFrame())
            if 'Login' in df_ref.columns and 'Site' in df_ref.columns:
                site_mapping = dict(zip(df_ref['Login'], df_ref['Site']))
                df_case['Site'] = df_case['Créé par ticket'].map(site_mapping)
                df_case['Site'] = df_case['Site'].fillna('Autres')
        
        # CORRECTION: Tickets créés par des collaborateurs de ce site SEULEMENT
        site_tickets = df_case[df_case['Site'] == site] if 'Site' in df_case.columns else df_case
        print(f"   - Tickets créés par le site {site}: {len(site_tickets)}")
        
        # Boutiques ayant ouvert des tickets AVEC CE SITE
        boutiques_site = site_tickets['Code compte'].unique()
        print(f"   - Boutiques du contexte {site}: {len(boutiques_site)}")
        
        # Stats par boutique DANS LE CONTEXTE du site
        shop_stats = []
        for code_compte in boutiques_site:
            # Tickets de cette boutique créés par le site
            boutique_tickets_site = site_tickets[site_tickets['Code compte'] == code_compte]
            total_tickets = len(boutique_tickets_site)
            tickets_clos = len(boutique_tickets_site[boutique_tickets_site['État'] == 'Clos'])
            taux_closure = (tickets_clos / total_tickets * 100) if total_tickets > 0 else 0
            
            # Réponses Q1 pour cette boutique SUR LES TICKETS DU SITE
            q1_data = df_merged[df_merged['Mesure'].str.contains('Q1', na=False)]
            tickets_site_nums = set(boutique_tickets_site['Numéro'].unique())
            q1_responses_site = q1_data[
                (q1_data['Code compte'] == code_compte) & 
                (q1_data['Dossier Rcbt'].isin(tickets_site_nums))
            ]
            
            q1_count = len(q1_responses_site)
            taux_retour = (q1_count / total_tickets * 100) if total_tickets > 0 else 0
            
            # Satisfaction sur les réponses du contexte
            if len(q1_responses_site) > 0:
                # Try different column names for satisfaction
                satisfaction_col = None
                for col in ['valuation', 'Valeur de chaîne', 'evaluation', 'satisfaction']:
                    if col in q1_responses_site.columns:
                        satisfaction_col = col
                        break
                
                if satisfaction_col:
                    satisfaits = q1_responses_site[satisfaction_col].isin(['Très satisfaisant', 'Satisfaisant']).sum()
                    satisfaction_rate = (satisfaits / len(q1_responses_site) * 100)
                else:
                    satisfaits = 0
                    satisfaction_rate = 0
            else:
                satisfaits = 0
                satisfaction_rate = 0
            
            # Info boutique
            shop_info = df_accounts[df_accounts['Code compte'] == code_compte]
            nom_boutique = shop_info.iloc[0].get('Compte', 'N/A') if len(shop_info) > 0 else 'N/A'
            categorie = shop_info.iloc[0].get('Categorie', 'N/A') if len(shop_info) > 0 else 'N/A'
            
            shop_stats.append({
                'Code compte': int(code_compte),
                'Nom boutique': nom_boutique,
                'Catégorie': categorie,
                'Total_tickets': total_tickets,
                'Tickets_clos': tickets_clos,
                'Taux_closure': taux_closure,
                'Q1_responses': q1_count,
                'Satisfaits': satisfaits,
                'Satisfaction_rate': satisfaction_rate,
                'Taux_retour': taux_retour
            })
        
        # Trier par taux de retour décroissant
        shop_stats.sort(key=lambda x: x['Taux_retour'], reverse=True)
        
        # Construire le tableau HTML
        shop_ranking_table = ""
        for shop in shop_stats:
            # Classes CSS pour les couleurs
            closure_class = 'success' if shop['Taux_closure'] >= 13 else ('warning' if shop['Taux_closure'] >= 10 else 'danger')
            return_class = 'success' if shop['Taux_retour'] >= 15 else ('warning' if shop['Taux_retour'] >= 10 else 'danger')
            satisfaction_class = 'success' if shop['Satisfaction_rate'] >= 92 else ('warning' if shop['Satisfaction_rate'] >= 85 else 'danger')
            
            shop_ranking_table += f"""
            <tr>
                <td>{shop['Code compte']}</td>
                <td>{shop['Nom boutique']}</td>
                <td>{shop['Catégorie']}</td>
                <td>{shop['Total_tickets']}</td>
                <td class="text-{closure_class}">{shop['Taux_closure']:.1f}%</td>
                <td>{shop['Q1_responses']}</td>
                <td>{shop['Satisfaits']}</td>
                <td class="text-{satisfaction_class}">{shop['Satisfaction_rate']:.1f}%</td>
                <td class="text-{return_class}">{shop['Taux_retour']:.1f}%</td>
            </tr>"""
        
        return {
            'page7_shop_ranking_table': shop_ranking_table,
            'page7_ranking_table': shop_ranking_table,  # Alias for template compatibility
            'page7_total_shops': len(shop_stats)
        }
    
    def _get_global_shop_ranking(self, data):
        """Page 7 logique globale standard"""
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
        
        # Add account info
        account_cols = ['Code compte']
        if 'Compte' in df_accounts.columns:
            account_cols.append('Compte')
        if 'Categorie' in df_accounts.columns:
            account_cols.append('Categorie')
            
        shop_ranking = shop_ranking.merge(df_accounts[account_cols], on='Code compte', how='left')
        shop_ranking['Enquetes_repondues'] = shop_ranking['Enquetes_repondues'].fillna(0)
        shop_ranking['Enquetes_satisfaites'] = shop_ranking['Enquetes_satisfaites'].fillna(0)
        shop_ranking['Taux_satisfaction'] = shop_ranking['Taux_satisfaction'].fillna(0)
        
        # Sort by total tickets
        shop_ranking = shop_ranking.sort_values('Total_tickets', ascending=False)
        
        filter_info = f"{self.filter_type} {self.filter_value}" if self.filter_type else "Global"
        return self._build_ranking_table(shop_ranking, filter_info)
    
    def _build_ranking_table(self, shop_ranking, context_info):
        """Construire le tableau HTML du classement"""
        # Display all shops that have data
        all_shops = shop_ranking[shop_ranking['Total_tickets'] > 0]
        
        ranking_table = ""
        name_col = 'Compte' if 'Compte' in shop_ranking.columns else 'Account'
        cat_col = 'Categorie' if 'Categorie' in shop_ranking.columns else 'Category'
        
        for i, (_, row) in enumerate(all_shops.iterrows(), 1):
            code_compte = int(float(row.get('Code compte', 0))) if pd.notna(row.get('Code compte')) else 'N/A'
            account_name = row.get(name_col, 'N/A') if name_col in shop_ranking.columns else 'N/A'
            
            # Données adaptées selon le contexte
            total_tickets = int(row.get('Total_tickets', 0))
            tickets_clos = int(row.get('Tickets_clos', 0))
            taux_closure = row.get('Taux_closure', 0)
            
            # Taux de retour (adapté pour collaborateur)
            if hasattr(row, 'Taux_retour') and 'Taux_retour' in row.index:
                taux_retour = row.get('Taux_retour', 0)
            else:
                # Calcul standard pour sites
                enquetes_repondues = int(row.get('Enquetes_repondues', 0))
                taux_retour = (enquetes_repondues / total_tickets * 100) if total_tickets > 0 else 0
            
            enquetes_repondues = int(row.get('Enquetes_repondues', 0))
            taux_satisfaction = row.get('Taux_satisfaction', 0)
            
            # Color coding
            if enquetes_repondues == 0:
                sat_display = 'NC'
                sat_class = 'secondary'
            else:
                sat_display = f"{taux_satisfaction:.1f}%"
                sat_class = 'success' if taux_satisfaction >= 92 else ('warning' if taux_satisfaction >= 80 else 'danger')
            
            closure_class = 'success' if taux_closure >= 13 else ('warning' if taux_closure >= 10 else 'danger')
            return_class = 'success' if taux_retour >= 30 else ('warning' if taux_retour >= 20 else 'danger')
            
            ranking_table += f"""
            <tr>
                <td>{i}</td>
                <td>{code_compte}</td>
                <td>{account_name}</td>
                <td>{row.get(cat_col, 'N/A') if cat_col in shop_ranking.columns else 'N/A'}</td>
                <td>{total_tickets}</td>
                <td>{tickets_clos}</td>
                <td class="text-{closure_class}">{taux_closure:.1f}%</td>
                <td class="text-{return_class}">{taux_retour:.1f}%</td>
                <td>{enquetes_repondues}</td>
                <td>{int(row.get('Enquetes_satisfaites', 0))}</td>
                <td class="text-{sat_class}">{sat_display}</td>
            </tr>"""
        
        print(f"📊 {context_info}: {len(all_shops)} boutiques classées")
        
        return {
            'page7_ranking_table': ranking_table,
            'page7_total_shops': len(all_shops),
            'page7_context': context_info
        }