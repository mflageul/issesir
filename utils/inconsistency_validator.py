"""
Module pour la validation et traitement des incohérences détectées
Permet la révision manuelle avant génération des rapports finaux
"""
import pandas as pd
import json
import os
from typing import List, Dict, Any
from dataclasses import dataclass
from datetime import datetime


@dataclass
class InconsistencyItem:
    """Structure pour un élément d'incohérence"""
    dossier: str
    collaborator: str
    original_rating: str
    comment: str
    inconsistency_type: str
    detected_words: List[str]
    suggested_rating: str = ""
    validated_rating: str = ""
    validation_status: str = "pending"  # pending, validated, ignored
    validation_reason: str = ""
    validator: str = ""
    validation_date: str = ""


class InconsistencyValidator:
    """Gestionnaire de validation des incohérences"""
    
    def __init__(self):
        self.inconsistencies: List[InconsistencyItem] = []
        self.suggestions_rules = {
            # TYPE 1: Note négative avec commentaire positif
            'Note négative / Commentaire positif': {
                'Très peu satisfaisant': 'Satisfaisant',
                'Peu satisfaisant': 'Satisfaisant'
            },
            # TYPE 2: Note positive avec commentaire négatif  
            'Note positive / Commentaire négatif': {
                'Très satisfaisant': 'Peu satisfaisant',
                'Satisfaisant': 'Peu satisfaisant'
            }
        }
    
    def load_inconsistencies(self, inconsistencies_data: List[Dict]) -> None:
        """Charge les incohérences détectées pour validation"""
        self.inconsistencies = []
        
        for item in inconsistencies_data:
            # CORRECTION: Utiliser la suggestion du DataProcessor plutôt que recalculer
            rating_value = item.get('rating', '')
            suggested_rating = item.get('suggested_rating', rating_value)  # Utiliser la suggestion du DataProcessor
            
            inconsistency = InconsistencyItem(
                dossier=item.get('dossier', ''),
                collaborator=item.get('collaborator', ''),
                original_rating=rating_value,
                comment=item.get('comment', ''),
                inconsistency_type=item.get('inconsistency_type', ''),
                detected_words=item.get('detected_words', []),
                suggested_rating=suggested_rating
            )
            
            self.inconsistencies.append(inconsistency)
        
        # Charger les validations existantes si elles existent
        self._load_existing_validations()
    
    def _suggest_correction(self, original_rating: str, inconsistency_type: str) -> str:
        """Suggère une correction basée sur le type d'incohérence"""
        if inconsistency_type in self.suggestions_rules:
            return self.suggestions_rules[inconsistency_type].get(
                original_rating, original_rating
            )
        return original_rating
    
    def validate_inconsistency(self, dossier: str, validated_rating: str, 
                             reason: str = "", validator: str = "User") -> bool:
        """Valide une incohérence spécifique"""
        for inconsistency in self.inconsistencies:
            if inconsistency.dossier == dossier:
                inconsistency.validated_rating = validated_rating
                inconsistency.validation_status = "validated"
                inconsistency.validation_reason = reason
                inconsistency.validator = validator
                inconsistency.validation_date = datetime.now().strftime('%d/%m/%Y %H:%M')
                
                # Sauvegarder immédiatement la validation
                self._save_validations()
                return True
        return False
    
    def ignore_inconsistency(self, dossier: str, reason: str = "", validator: str = "User") -> bool:
        """Ignore une incohérence (conserver note originale)"""
        for inconsistency in self.inconsistencies:
            if inconsistency.dossier == dossier:
                inconsistency.validated_rating = inconsistency.original_rating
                inconsistency.validation_status = "ignored"
                inconsistency.validation_reason = reason
                inconsistency.validator = validator
                inconsistency.validation_date = datetime.now().strftime('%d/%m/%Y %H:%M')
                
                # Sauvegarder immédiatement la validation
                self._save_validations()
                return True
        return False
    

    
    def get_validated_inconsistencies(self) -> List[InconsistencyItem]:
        """Retourne les incohérences validées"""
        return [inc for inc in self.inconsistencies if inc.validation_status != "pending"]
    
    def get_pending_inconsistencies(self) -> List[InconsistencyItem]:
        """Retourne les incohérences en attente de validation"""
        return [inc for inc in self.inconsistencies if inc.validation_status == "pending"]
    
    def _load_existing_validations(self):
        """Charge les validations existantes depuis le fichier JSON"""
        try:
            os.makedirs('session_cache', exist_ok=True)
            if os.path.exists(self.validation_file):
                with open(self.validation_file, 'r', encoding='utf-8') as f:
                    validations = json.load(f)
                    
                # Appliquer les validations existantes
                for validation in validations:
                    for inconsistency in self.inconsistencies:
                        if inconsistency.dossier == validation.get('dossier'):
                            inconsistency.validated_rating = validation.get('validated_rating', '')
                            inconsistency.validation_status = validation.get('validation_status', 'pending')
                            inconsistency.validation_reason = validation.get('validation_reason', '')
                            inconsistency.validator = validation.get('validator', '')
                            inconsistency.validation_date = validation.get('validation_date', '')
                            break
                            
                print(f"✅ Validations existantes restaurées depuis {self.validation_file}")
        except Exception as e:
            print(f"⚠️ Erreur chargement validations: {e}")
    
    def _save_validations(self):
        """Sauvegarde les validations dans un fichier JSON persistant"""
        try:
            os.makedirs('session_cache', exist_ok=True)
            
            validations = []
            for inconsistency in self.inconsistencies:
                if inconsistency.validation_status != "pending":
                    validations.append({
                        'dossier': inconsistency.dossier,
                        'original_rating': inconsistency.original_rating,
                        'validated_rating': inconsistency.validated_rating,
                        'validation_status': inconsistency.validation_status,
                        'validation_reason': inconsistency.validation_reason,
                        'validator': inconsistency.validator,
                        'validation_date': inconsistency.validation_date,
                        'inconsistency_type': inconsistency.inconsistency_type
                    })
            
            with open(self.validation_file, 'w', encoding='utf-8') as f:
                json.dump(validations, f, ensure_ascii=False, indent=2)
                
            print(f"✅ {len(validations)} validations sauvegardées dans {self.validation_file}")
        except Exception as e:
            print(f"⚠️ Erreur sauvegarde validations: {e}")
    
    def apply_validations_to_dataframe(self, df_merged: pd.DataFrame) -> pd.DataFrame:
        """Applique les validations au DataFrame tout en conservant traçabilité"""
        df_modified = df_merged.copy()
        
        # Ajouter colonnes de traçabilité si pas présentes
        if 'Original_Rating' not in df_modified.columns:
            df_modified['Original_Rating'] = df_modified['Valeur de chaîne']
        if 'Validation_Applied' not in df_modified.columns:
            df_modified['Validation_Applied'] = False
        if 'Validation_Reason' not in df_modified.columns:
            df_modified['Validation_Reason'] = ''
        if 'Validation_Date' not in df_modified.columns:
            df_modified['Validation_Date'] = ''
        
        # Appliquer les validations
        modifications_count = 0
        for inconsistency in self.inconsistencies:
            if inconsistency.validation_status == "validated":
                # Trouver les lignes correspondantes
                mask = df_modified['Dossier Rcbt'] == inconsistency.dossier
                matching_rows = df_modified[mask]
                
                if not matching_rows.empty:
                    # Appliquer la nouvelle note validée
                    df_modified.loc[mask, 'Valeur de chaîne'] = inconsistency.validated_rating
                    df_modified.loc[mask, 'Validation_Applied'] = True
                    df_modified.loc[mask, 'Validation_Reason'] = inconsistency.validation_reason
                    df_modified.loc[mask, 'Validation_Date'] = inconsistency.validation_date
                    modifications_count += len(matching_rows)
        
        print(f"✅ Validations appliquées: {modifications_count} lignes modifiées")
        return df_modified
    
    def get_validation_summary(self) -> Dict[str, Any]:
        """Retourne un résumé des validations"""
        total = len(self.inconsistencies)
        validated = len([inc for inc in self.inconsistencies if inc.validation_status == "validated"])
        ignored = len([inc for inc in self.inconsistencies if inc.validation_status == "ignored"])
        pending = len([inc for inc in self.inconsistencies if inc.validation_status == "pending"])
        
        return {
            'total_detected': total,
            'validated_corrected': validated,
            'validated_ignored': ignored,
            'pending_validation': pending,
            'completion_rate': round((validated + ignored) / total * 100, 1) if total > 0 else 0
        }
    
    def export_validation_log(self) -> Dict[str, Any]:
        """Exporte le log des validations pour audit"""
        return {
            'export_date': datetime.now().strftime('%d/%m/%Y %H:%M:%S'),
            'summary': self.get_validation_summary(),
            'validations': [
                {
                    'dossier': inc.dossier,
                    'collaborator': inc.collaborator,
                    'original_rating': inc.original_rating,
                    'validated_rating': inc.validated_rating,
                    'inconsistency_type': inc.inconsistency_type,
                    'detected_words': inc.detected_words,
                    'validation_status': inc.validation_status,
                    'validation_reason': inc.validation_reason,
                    'validator': inc.validator,
                    'validation_date': inc.validation_date
                }
                for inc in self.inconsistencies
            ]
        }