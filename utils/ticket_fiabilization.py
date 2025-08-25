"""
Module de fiabilisation des tickets "Autres"
Permet de rattacher les tickets non référencés à des sites/collaborateurs
basé sur différentes colonnes du fichier case
"""

import pandas as pd
from typing import Dict, Tuple, Set

class TicketFiabilizer:
    """Classe pour fiabiliser le rattachement des tickets orphelins"""
    
    def __init__(self, df_case: pd.DataFrame, df_ref: pd.DataFrame):
        self.df_case = df_case.copy()
        self.df_ref = df_ref.copy()
        self.site_mapping = dict(zip(df_ref['Login'], df_ref['Site'])) if not df_ref.empty else {}
        self.collab_ref = set(df_ref['Login'].unique()) if not df_ref.empty else set()
        
    def analyze_fiabilization_potential(self) -> Dict:
        """Analyse le potentiel de fiabilisation des tickets"""
        results = {
            'total_tickets': len(self.df_case),
            'tickets_orphelins': 0,
            'colonnes_analysees': {},
            'fiabilisation_possible': 0,
            'strategie_recommandee': []
        }
        
        # Identifier les tickets orphelins
        df_with_site = self._add_site_mapping()
        tickets_orphelins = df_with_site[df_with_site['Site'] == 'Autres']
        results['tickets_orphelins'] = len(tickets_orphelins)
        
        print(f"🔍 ANALYSE FIABILISATION: {len(tickets_orphelins)} tickets orphelins à analyser")
        
        # Analyser chaque colonne potentielle
        colonnes_fiabilisation = [
            'Résolu par', 'Pris en charge par', 'Mis à jour par', 
            'Clos par', 'Mis a traiter par', "Groupe d'affectation"
        ]
        
        for col in colonnes_fiabilisation:
            if col in self.df_case.columns:
                fiabilisation_col = self._analyze_column_for_fiabilization(tickets_orphelins, col)
                results['colonnes_analysees'][col] = fiabilisation_col
                results['fiabilisation_possible'] += fiabilisation_col['tickets_fiabilisables']
        
        # Stratégie recommandée
        if results['fiabilisation_possible'] > 0:
            results['strategie_recommandee'].append(f"Fiabilisation de {results['fiabilisation_possible']} tickets possible")
        
        return results
    
    def _analyze_column_for_fiabilization(self, tickets_orphelins: pd.DataFrame, column: str) -> Dict:
        """Analyse une colonne pour la fiabilisation"""
        result = {
            'colonne': column,
            'valeurs_uniques': 0,
            'collaborateurs_references': 0,
            'tickets_fiabilisables': 0,
            'exemple_fiabilisation': []
        }
        
        if column not in tickets_orphelins.columns:
            return result
            
        # Valeurs uniques dans la colonne
        values = tickets_orphelins[column].dropna().unique()
        result['valeurs_uniques'] = len(values)
        
        # Intersection avec collaborateurs référencés
        intersection = set(values) & self.collab_ref
        result['collaborateurs_references'] = len(intersection)
        
        # Tickets fiabilisables via cette colonne
        if len(intersection) > 0:
            tickets_fiabilisables = tickets_orphelins[tickets_orphelins[column].isin(intersection)]
            result['tickets_fiabilisables'] = len(tickets_fiabilisables)
            
            # Exemples de fiabilisation
            for _, row in tickets_fiabilisables.head(3).iterrows():
                collaborateur_fiabilisant = row[column]
                site_destination = self.site_mapping.get(collaborateur_fiabilisant, 'Non trouvé')
                result['exemple_fiabilisation'].append({
                    'ticket': row['Numéro'],
                    'cree_par': row['Créé par ticket'],
                    'fiabilise_par': collaborateur_fiabilisant,
                    'site_destination': site_destination
                })
        
        return result
    
    def fiabilize_tickets(self, colonnes_priorite: list = None) -> pd.DataFrame:
        """
        Fiabilise les tickets en utilisant les colonnes spécifiées
        
        Args:
            colonnes_priorite: Liste des colonnes à utiliser par ordre de priorité
                              Par défaut: ['Résolu par', 'Pris en charge par', 'Mis à jour par']
        """
        if colonnes_priorite is None:
            colonnes_priorite = ['Résolu par', 'Pris en charge par', 'Mis à jour par']
        
        df_fiabilise = self._add_site_mapping()
        tickets_fiabilises = 0
        
        print(f"🔧 FIABILISATION EN COURS avec colonnes: {colonnes_priorite}")
        
        for colonne in colonnes_priorite:
            if colonne not in df_fiabilise.columns:
                continue
                
            # Identifier les tickets encore orphelins
            tickets_orphelins = df_fiabilise[df_fiabilise['Site'] == 'Autres'].copy()
            
            if len(tickets_orphelins) == 0:
                break
                
            # Fiabiliser via cette colonne
            mask_fiabilisation = (
                tickets_orphelins[colonne].isin(self.collab_ref) & 
                tickets_orphelins[colonne].notna()
            )
            
            tickets_a_fiabiliser = tickets_orphelins[mask_fiabilisation]
            
            if len(tickets_a_fiabiliser) > 0:
                print(f"   📌 {colonne}: {len(tickets_a_fiabiliser)} tickets fiabilisés")
                
                for _, ticket in tickets_a_fiabiliser.iterrows():
                    collaborateur_fiabilisant = ticket[colonne]
                    site_destination = self.site_mapping.get(collaborateur_fiabilisant)
                    
                    if site_destination:
                        # Mettre à jour le site ET le collaborateur dans df_fiabilise
                        mask_update = df_fiabilise['Numéro'] == ticket['Numéro']
                        df_fiabilise.loc[mask_update, 'Site'] = site_destination
                        df_fiabilise.loc[mask_update, 'Créé par ticket'] = collaborateur_fiabilisant  # ATTRIBUTION AU COLLABORATEUR
                        df_fiabilise.loc[mask_update, 'Fiabilise_par'] = colonne
                        df_fiabilise.loc[mask_update, 'Collaborateur_fiabilisant'] = collaborateur_fiabilisant
                        df_fiabilise.loc[mask_update, 'Createur_original'] = ticket['Créé par ticket']  # TRAÇABILITE ORIGINALE
                        tickets_fiabilises += 1
        
        print(f"✅ FIABILISATION TERMINÉE: {tickets_fiabilises} tickets fiabilisés")
        
        # Ajouter les colonnes de traçabilité
        if 'Fiabilise_par' not in df_fiabilise.columns:
            df_fiabilise['Fiabilise_par'] = 'Original'
            df_fiabilise['Collaborateur_fiabilisant'] = df_fiabilise['Créé par ticket']
            df_fiabilise['Createur_original'] = df_fiabilise['Créé par ticket']
        
        return df_fiabilise
    
    def _add_site_mapping(self) -> pd.DataFrame:
        """Ajoute le mapping site initial"""
        df_with_site = self.df_case.copy()
        df_with_site['Site'] = df_with_site['Créé par ticket'].map(self.site_mapping)
        df_with_site['Site'] = df_with_site['Site'].fillna('Autres')
        return df_with_site
    
    def get_fiabilization_stats(self, df_fiabilise: pd.DataFrame) -> Dict:
        """Retourne les statistiques de fiabilisation"""
        stats = {
            'total_tickets': len(df_fiabilise),
            'tickets_originaux': len(df_fiabilise[df_fiabilise.get('Fiabilise_par', 'Original') == 'Original']),
            'tickets_fiabilises': len(df_fiabilise[df_fiabilise.get('Fiabilise_par', 'Original') != 'Original']),
            'repartition_sites': {},
            'methodes_fiabilisation': {}
        }
        
        # Répartition par site après fiabilisation
        for site in df_fiabilise['Site'].unique():
            stats['repartition_sites'][site] = len(df_fiabilise[df_fiabilise['Site'] == site])
        
        # Méthodes de fiabilisation utilisées
        if 'Fiabilise_par' in df_fiabilise.columns:
            for methode in df_fiabilise['Fiabilise_par'].unique():
                stats['methodes_fiabilisation'][methode] = len(
                    df_fiabilise[df_fiabilise['Fiabilise_par'] == methode]
                )
        
        return stats