"""
Page 7 pour rapports individuels - LOGIQUE ENTIÈREMENT CORRIGÉE
Calculs selon spécifications exactes de l'utilisateur
"""

import pandas as pd

class Page7IndividualCalculator:
    def __init__(self, filter_type, filter_value):
        self.filter_type = filter_type  # 'collaborator' or 'site'
        self.filter_value = filter_value
    
    def calculate_page7_data(self, data):
        """
        Calcule la Page 7 selon les spécifications exactes :
        1ère colonne "Rang" = classement par nombre tickets incidents ouverts
        2ème colonne "Code compte" = colonne A du fichier customer_account
        3ème colonne "Nom boutique" = colonne C du fichier customer_account  
        4ème colonne "Catégorie" = type boutique (400, 499, 993, siège)
        5ème colonne "tickets clos" = nombre tickets générés par la boutique
        6ème colonne "% Taux de retour" = % tickets avec réponse enquête vs tickets ouverts
        7ème colonne "enquetes repondues" = nombre réponses enquête satisfaction
        8ème colonne "enquetes satisfaites" = volume tickets satisfaits/très satisfaits
        9ème colonne "% Satisfaction" = % tickets satisfaits vs volume tickets ouverts
        """
        
        df_case = data['case']  # sn_customerservice_case.xlsx
        df_asmt = data['merged']  # asmt_metric_result.xlsx (dans merged)
        df_accounts = data['accounts']  # customer_account.xlsx
        
        print(f"📊 Calcul Page 7 CORRIGÉ pour {self.filter_type}: {self.filter_value}")
        
        # ÉTAPE 1: Filtrer les tickets selon le contexte (collaborateur ou site)
        if self.filter_type == 'collaborator':
            # Tickets créés par ce collaborateur spécifique
            context_tickets = df_case[df_case['Créé par ticket'] == self.filter_value]
            print(f"   - Tickets créés par collaborateur {self.filter_value}: {len(context_tickets)}")
        elif self.filter_type == 'site':
            # Assurer le mapping site si nécessaire
            if 'Site' not in df_case.columns:
                df_ref = data.get('ref', pd.DataFrame())
                if not df_ref.empty and 'Login' in df_ref.columns and 'Site' in df_ref.columns:
                    site_mapping = dict(zip(df_ref['Login'], df_ref['Site']))
                    df_case['Site'] = df_case['Créé par ticket'].map(site_mapping)
                    df_case['Site'] = df_case['Site'].fillna('Autres')
            
            # Tickets créés par des collaborateurs de ce site
            context_tickets = df_case[df_case['Site'] == self.filter_value] if 'Site' in df_case.columns else df_case
            print(f"   - Tickets créés par site {self.filter_value}: {len(context_tickets)}")
        else:
            context_tickets = df_case
        
        if len(context_tickets) == 0:
            return self._empty_result(f"Aucun ticket trouvé pour {self.filter_type} {self.filter_value}")
        
        # ÉTAPE 2: Identifier les boutiques dans le contexte
        boutiques_contexte = context_tickets['Code compte'].unique()
        print(f"   - Boutiques dans le contexte: {len(boutiques_contexte)}")
        
        # ÉTAPE 3: Filtrer les enquêtes Q1 selon le contexte
        q1_data = df_asmt[df_asmt['Mesure'].str.contains('Q1', na=False)]
        context_ticket_nums = set(context_tickets['Numéro'].unique())
        q1_contexte = q1_data[q1_data['Dossier Rcbt'].isin(context_ticket_nums)]
        
        print(f"   - Enquêtes Q1 dans le contexte: {len(q1_contexte)}")
        
        # ÉTAPE 4: Calculer les statistiques par boutique (FILTRÉES - seulement celles avec tickets)
        boutique_stats = []
        
        for code_compte in boutiques_contexte:
            # Tickets de cette boutique dans le contexte
            tickets_boutique = context_tickets[context_tickets['Code compte'] == code_compte]
            nb_tickets_boutique = len(tickets_boutique)
            
            # Enquêtes répondues pour cette boutique dans le contexte
            enquetes_boutique = q1_contexte[q1_contexte['Code compte'] == code_compte]
            nb_enquetes_repondues = len(enquetes_boutique)
            
            # Calcul du taux de retour (spécification 6ème colonne)
            taux_retour = (nb_enquetes_repondues / nb_tickets_boutique * 100) if nb_tickets_boutique > 0 else 0
            
            # Enquêtes satisfaites (spécification 8ème colonne)
            if len(enquetes_boutique) > 0:
                satisfaction_col = self._find_satisfaction_column(enquetes_boutique)
                if satisfaction_col:
                    enquetes_satisfaites = enquetes_boutique[satisfaction_col].isin(['Très satisfaisant', 'Satisfaisant']).sum()
                else:
                    enquetes_satisfaites = 0
            else:
                enquetes_satisfaites = 0
            
            # Pourcentage satisfaction (spécification 9ème colonne)
            # % de tickets satisfaits vs volume tickets ouverts par la boutique
            pourcentage_satisfaction = (enquetes_satisfaites / nb_tickets_boutique * 100) if nb_tickets_boutique > 0 else 0
            
            # Informations boutique (spécifications colonnes 2, 3, 4)
            code_compte_display, nom_boutique, categorie = self._get_boutique_info(df_accounts, code_compte)
            
            # FILTRE: Ne garder que les boutiques avec des tickets traités par le collaborateur/site
            if nb_tickets_boutique > 0:
                boutique_stats.append({
                    'code_compte': code_compte_display,  # Colonne A customer_account
                    'nom_boutique': nom_boutique,        # Colonne C customer_account
                    'categorie': categorie,              # Type 400/499/993/siège
                    'nb_tickets': nb_tickets_boutique,   # Nombre tickets période
                    'taux_retour': taux_retour,          # % tickets avec réponse enquête
                    'enquetes_repondues': nb_enquetes_repondues,  # Nombre réponses enquête
                    'enquetes_satisfaites': enquetes_satisfaites, # Volume tickets satisfaits
                    'pourcentage_satisfaction': pourcentage_satisfaction  # % satisfaction vs tickets
                })
                
                print(f"   - Boutique {code_compte_display} ({nom_boutique}): {nb_tickets_boutique} tickets, {nb_enquetes_repondues} réponses ({taux_retour:.1f}%), {enquetes_satisfaites} satisfaites ({pourcentage_satisfaction:.1f}%)")
            else:
                print(f"   - Boutique {code_compte_display} FILTRÉE (aucun ticket traité)")
        
        # ÉTAPE 5: Trier par nombre de tickets (spécification 1ère colonne "Rang")
        boutique_stats.sort(key=lambda x: x['nb_tickets'], reverse=True)
        
        # ÉTAPE 6: Ajouter les rangs
        for i, stat in enumerate(boutique_stats, 1):
            stat['rang'] = i
        
        # ÉTAPE 7: Construire le tableau HTML
        table_html = self._build_table_html(boutique_stats)
        
        return {
            'page7_shop_ranking_table': table_html,
            'page7_ranking_table': table_html,
            'page7_total_shops': len(boutique_stats)
        }
    
    def _get_boutique_info(self, df_accounts, code_compte):
        """Récupère les informations boutique selon spécifications"""
        boutique_info = df_accounts[df_accounts['Code compte'] == code_compte]
        
        if len(boutique_info) > 0:
            row = boutique_info.iloc[0]
            
            # Code compte - colonne A (spécification 2ème colonne)
            code_display = int(code_compte) if pd.notna(code_compte) else 'N/A'
            
            # Nom boutique - colonne C (spécification 3ème colonne)
            # Essayer différents noms de colonnes possibles
            nom_boutique = 'N/A'
            for col in ['Compte', 'Account', 'Nom', 'Nom boutique', 'C']:
                if col in row.index and pd.notna(row.get(col)):
                    nom_value = str(row.get(col)).strip()
                    if nom_value and nom_value != 'nan':
                        nom_boutique = nom_value
                        break
            
            # Catégorie (spécification 4ème colonne)
            categorie = 'N/A'
            for col in ['Categorie', 'Category', 'Type', 'Catégorie']:
                if col in row.index and pd.notna(row.get(col)):
                    cat_value = str(row.get(col)).strip()
                    if cat_value and cat_value != 'nan':
                        categorie = cat_value
                        break
        else:
            code_display = int(code_compte) if pd.notna(code_compte) else 'N/A'
            nom_boutique = 'N/A'
            categorie = 'N/A'
        
        return code_display, nom_boutique, categorie
    
    def _find_satisfaction_column(self, df):
        """Trouve la colonne contenant les valeurs de satisfaction"""
        for col in ['Valeur de chaîne', 'valuation', 'evaluation', 'satisfaction', 'value']:
            if col in df.columns:
                return col
        return None
    
    def _build_table_html(self, stats):
        """Construit le tableau HTML selon les spécifications"""
        table_html = ""
        
        for stat in stats:
            # Classes CSS pour les couleurs
            taux_retour_class = 'success' if stat['taux_retour'] >= 15 else ('warning' if stat['taux_retour'] >= 10 else 'danger')
            satisfaction_class = 'success' if stat['pourcentage_satisfaction'] >= 92 else ('warning' if stat['pourcentage_satisfaction'] >= 85 else 'danger')
            
            table_html += f"""
            <tr>
                <td>{stat['rang']}</td>
                <td>{stat['code_compte']}</td>
                <td>{stat['nom_boutique']}</td>
                <td>{stat['categorie']}</td>
                <td>{stat['nb_tickets']}</td>
                <td class="text-{taux_retour_class}">{stat['taux_retour']:.1f}%</td>
                <td>{stat['enquetes_repondues']}</td>
                <td>{stat['enquetes_satisfaites']}</td>
                <td class="text-{satisfaction_class}">{stat['pourcentage_satisfaction']:.1f}%</td>
            </tr>"""
        
        return table_html
    
    def _empty_result(self, message):
        """Retourne un résultat vide avec message"""
        return {
            'page7_shop_ranking_table': f'<tr><td colspan="9">{message}</td></tr>',
            'page7_ranking_table': f'<tr><td colspan="9">{message}</td></tr>',
            'page7_total_shops': 0
        }