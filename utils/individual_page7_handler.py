"""
Gestionnaire spécialisé pour la Page 7 des rapports individuels
Adapte les calculs selon le collaborateur ou site sélectionné - VERSION CORRIGÉE
"""

import pandas as pd

class IndividualPage7Handler:
    def __init__(self, filter_type, filter_value):
        self.filter_type = filter_type  # 'site' or 'collaborator'
        self.filter_value = filter_value
    
    def get_individual_shop_ranking(self, data):
        """Page 7 adaptée pour rapports individuels - ENTIÈREMENT CORRIGÉE"""
        df_merged = data['merged']
        df_case = data['case']
        df_accounts = data['accounts']
        
        print(f"🔍 Page 7 individuelle CORRIGÉE: {self.filter_type} = {self.filter_value}")
        
        if self.filter_type == 'collaborator':
            return self._get_collaborator_shop_ranking_fixed(data)
        elif self.filter_type == 'site':
            return self._get_site_shop_ranking_fixed(data)
        else:
            # Fallback to global ranking
            return self._get_global_shop_ranking(data)
    
    def _get_collaborator_shop_ranking_fixed(self, data):
        """Page 7 pour un collaborateur spécifique - LOGIQUE CORRIGÉE"""
        df_merged = data['merged']
        df_case = data['case']
        df_accounts = data['accounts']
        collaborator = self.filter_value
        
        print(f"📊 Calculs Page 7 CORRIGÉS pour collaborateur: {collaborator}")
        
        # NOUVELLE LOGIQUE: Utiliser les tickets CRÉÉS par le collaborateur
        collaborator_tickets = df_case[df_case['Créé par ticket'] == collaborator]
        print(f"   - Tickets créés par {collaborator}: {len(collaborator_tickets)}")
        
        if len(collaborator_tickets) == 0:
            print(f"   ⚠️ Aucun ticket créé par {collaborator}")
            return {
                'page7_shop_ranking_table': '<tr><td colspan="9">Aucun ticket créé par ce collaborateur</td></tr>',
                'page7_ranking_table': '<tr><td colspan="9">Aucun ticket créé par ce collaborateur</td></tr>',
                'page7_total_shops': 0
            }
        
        # Q1 data pour filtrage
        q1_data = df_merged[df_merged['Mesure'].str.contains('Q1', na=False)]
        
        # Stats par boutique basées sur les tickets créés par le collaborateur
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
            
            # Satisfaction sur ces réponses avec colonne flexible
            if len(q1_responses) > 0:
                satisfaction_col = self._get_satisfaction_column(q1_responses)
                if satisfaction_col:
                    satisfaits = q1_responses[satisfaction_col].isin(['Très satisfaisant', 'Satisfaisant']).sum()
                    satisfaction_rate = (satisfaits / len(q1_responses) * 100)
                else:
                    satisfaits = 0
                    satisfaction_rate = 0
            else:
                satisfaits = 0
                satisfaction_rate = 0
            
            # Info boutique CORRIGÉE avec gestion des colonnes
            nom_boutique, categorie = self._get_shop_info(df_accounts, code_compte)
            
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
            
            print(f"   - Boutique {code_compte} ({nom_boutique}): {q1_count}/{total_tickets} réponses ({taux_retour:.1f}%), satisfaction {satisfaction_rate:.1f}%")
        
        # Trier par nombre de tickets décroissant
        shop_stats.sort(key=lambda x: x['Total_tickets'], reverse=True)
        
        # Construire le tableau HTML
        shop_ranking_table = self._build_shop_table_html(shop_stats)
        
        return {
            'page7_shop_ranking_table': shop_ranking_table,
            'page7_ranking_table': shop_ranking_table,  # Alias for template compatibility
            'page7_total_shops': len(shop_stats)
        }
    
    def _get_site_shop_ranking_fixed(self, data):
        """Page 7 pour un site spécifique - CORRIGÉE pour contexte"""
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
        
        if len(site_tickets) == 0:
            print(f"   ⚠️ Aucun ticket pour le site {site}")
            return {
                'page7_shop_ranking_table': '<tr><td colspan="9">Aucun ticket pour ce site</td></tr>',
                'page7_ranking_table': '<tr><td colspan="9">Aucun ticket pour ce site</td></tr>',
                'page7_total_shops': 0
            }
        
        # Boutiques ayant ouvert des tickets AVEC CE SITE
        boutiques_site = site_tickets['Code compte'].unique()
        print(f"   - Boutiques du contexte {site}: {len(boutiques_site)}")
        
        # Q1 data pour filtrage
        q1_data = df_merged[df_merged['Mesure'].str.contains('Q1', na=False)]
        
        # Stats par boutique DANS LE CONTEXTE du site
        shop_stats = []
        for code_compte in boutiques_site:
            # Tickets de cette boutique créés par le site
            boutique_tickets_site = site_tickets[site_tickets['Code compte'] == code_compte]
            total_tickets = len(boutique_tickets_site)
            tickets_clos = len(boutique_tickets_site[boutique_tickets_site['État'] == 'Clos'])
            taux_closure = (tickets_clos / total_tickets * 100) if total_tickets > 0 else 0
            
            # Réponses Q1 pour cette boutique SUR LES TICKETS DU SITE
            tickets_site_nums = set(boutique_tickets_site['Numéro'].unique())
            q1_responses_site = q1_data[
                (q1_data['Code compte'] == code_compte) & 
                (q1_data['Dossier Rcbt'].isin(tickets_site_nums))
            ]
            
            q1_count = len(q1_responses_site)
            taux_retour = (q1_count / total_tickets * 100) if total_tickets > 0 else 0
            
            # Satisfaction sur les réponses du contexte
            if len(q1_responses_site) > 0:
                satisfaction_col = self._get_satisfaction_column(q1_responses_site)
                if satisfaction_col:
                    satisfaits = q1_responses_site[satisfaction_col].isin(['Très satisfaisant', 'Satisfaisant']).sum()
                    satisfaction_rate = (satisfaits / len(q1_responses_site) * 100)
                else:
                    satisfaits = 0
                    satisfaction_rate = 0
            else:
                satisfaits = 0
                satisfaction_rate = 0
            
            # Info boutique CORRIGÉE
            nom_boutique, categorie = self._get_shop_info(df_accounts, code_compte)
            
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
            
            print(f"   - Boutique {code_compte} ({nom_boutique}): {q1_count}/{total_tickets} réponses ({taux_retour:.1f}%), satisfaction {satisfaction_rate:.1f}%")
        
        # Trier par taux de retour décroissant
        shop_stats.sort(key=lambda x: x['Taux_retour'], reverse=True)
        
        # Construire le tableau HTML
        shop_ranking_table = self._build_shop_table_html(shop_stats)
        
        return {
            'page7_shop_ranking_table': shop_ranking_table,
            'page7_ranking_table': shop_ranking_table,
            'page7_total_shops': len(shop_stats)
        }
    
    def _get_shop_info(self, df_accounts, code_compte):
        """Récupère les informations de boutique avec gestion des colonnes"""
        shop_info = df_accounts[df_accounts['Code compte'] == code_compte]
        
        if len(shop_info) > 0:
            # Vérifier les différentes colonnes possibles pour le nom
            nom_boutique = 'N/A'
            for name_col in ['Compte', 'Account', 'Nom', 'Nom boutique', 'Shop Name']:
                if name_col in shop_info.columns and pd.notna(shop_info.iloc[0].get(name_col)):
                    nom_boutique = str(shop_info.iloc[0].get(name_col)).strip()
                    if nom_boutique and nom_boutique != 'nan':
                        break
            
            # Vérifier les différentes colonnes possibles pour la catégorie
            categorie = 'N/A'
            for cat_col in ['Categorie', 'Category', 'Type', 'Catégorie']:
                if cat_col in shop_info.columns and pd.notna(shop_info.iloc[0].get(cat_col)):
                    categorie = str(shop_info.iloc[0].get(cat_col)).strip()
                    if categorie and categorie != 'nan':
                        break
        else:
            nom_boutique = 'N/A'
            categorie = 'N/A'
        
        return nom_boutique, categorie
    
    def _get_satisfaction_column(self, df):
        """Trouve la colonne de satisfaction dans le DataFrame"""
        for col in ['Valeur de chaîne', 'valuation', 'evaluation', 'satisfaction', 'value']:
            if col in df.columns:
                return col
        return None
    
    def _build_shop_table_html(self, shop_stats):
        """Construit le tableau HTML des boutiques"""
        table_html = ""
        for shop in shop_stats:
            # Classes CSS pour les couleurs
            closure_class = 'success' if shop['Taux_closure'] >= 13 else ('warning' if shop['Taux_closure'] >= 10 else 'danger')
            return_class = 'success' if shop['Taux_retour'] >= 15 else ('warning' if shop['Taux_retour'] >= 10 else 'danger')
            satisfaction_class = 'success' if shop['Satisfaction_rate'] >= 92 else ('warning' if shop['Satisfaction_rate'] >= 85 else 'danger')
            
            table_html += f"""
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
        
        return table_html
    
    def _get_global_shop_ranking(self, data):
        """Fallback vers le ranking global"""
        return {
            'page7_shop_ranking_table': '<tr><td colspan="9">Mode global non implémenté dans handler individuel</td></tr>',
            'page7_ranking_table': '<tr><td colspan="9">Mode global non implémenté dans handler individuel</td></tr>',
            'page7_total_shops': 0
        }