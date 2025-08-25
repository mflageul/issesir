"""
Système de base de données centralisée pour la validation des incohérences
Garantit la persistance des validations entre rapports globaux et individuels
"""

import sqlite3
import json
import os
from datetime import datetime
from typing import Dict, List, Any, Optional
import pandas as pd

class ValidationDatabase:
    """Gestionnaire de base de données pour les validations d'incohérences"""
    
    def __init__(self, db_path: str = 'rcbt_validations.db'):
        self.db_path = db_path
        self._initialize_database()
    
    def _initialize_database(self):
        """Initialise la base de données avec les tables nécessaires"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # Table pour les validations d'incohérences
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS inconsistency_validations (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    dossier TEXT NOT NULL,
                    original_rating TEXT NOT NULL,
                    validated_rating TEXT,
                    validation_status TEXT DEFAULT 'pending',
                    validation_reason TEXT,
                    validator TEXT,
                    validation_date TIMESTAMP,
                    inconsistency_type TEXT,
                    detected_words TEXT,
                    session_id TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    UNIQUE(dossier, session_id)
                )
            ''')
            
            # Table pour les sessions de validation
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS validation_sessions (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    session_id TEXT UNIQUE NOT NULL,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    status TEXT DEFAULT 'active',
                    total_inconsistencies INTEGER DEFAULT 0,
                    validated_count INTEGER DEFAULT 0,
                    metadata TEXT
                )
            ''')
            
            conn.commit()
            conn.close()
            print(f"✅ Base de données de validation initialisée: {self.db_path}")
            
        except Exception as e:
            print(f"⚠️ Erreur initialisation base de données: {e}")
    
    def create_validation_session(self, inconsistencies_data: List[Dict]) -> str:
        """Crée une nouvelle session de validation"""
        try:
            session_id = f"session_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # Créer la session
            cursor.execute('''
                INSERT INTO validation_sessions (session_id, total_inconsistencies, metadata)
                VALUES (?, ?, ?)
            ''', (session_id, len(inconsistencies_data), json.dumps({'source': 'global_report'})))
            
            # Insérer les incohérences
            for item in inconsistencies_data:
                cursor.execute('''
                    INSERT OR REPLACE INTO inconsistency_validations 
                    (dossier, original_rating, inconsistency_type, detected_words, session_id)
                    VALUES (?, ?, ?, ?, ?)
                ''', (
                    item.get('dossier', ''),
                    item.get('rating', ''),
                    item.get('inconsistency_type', ''),
                    json.dumps(item.get('detected_words', [])),
                    session_id
                ))
            
            conn.commit()
            conn.close()
            
            print(f"✅ Session de validation créée: {session_id} avec {len(inconsistencies_data)} incohérences")
            return session_id
            
        except Exception as e:
            print(f"⚠️ Erreur création session: {e}")
            return ""
    
    def save_validation(self, dossier: str, validated_rating: str, 
                       validation_status: str, reason: str = "", 
                       validator: str = "User") -> bool:
        """Sauvegarde une validation d'incohérence"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # Mettre à jour la validation
            cursor.execute('''
                UPDATE inconsistency_validations 
                SET validated_rating = ?, validation_status = ?, validation_reason = ?,
                    validator = ?, validation_date = ?, updated_at = CURRENT_TIMESTAMP
                WHERE dossier = ?
            ''', (validated_rating, validation_status, reason, validator, 
                  datetime.now().isoformat(), dossier))
            
            if cursor.rowcount > 0:
                conn.commit()
                conn.close()
                print(f"✅ Validation sauvegardée pour dossier {dossier}")
                return True
            else:
                conn.close()
                print(f"⚠️ Aucune incohérence trouvée pour dossier {dossier}")
                return False
                
        except Exception as e:
            print(f"⚠️ Erreur sauvegarde validation: {e}")
            return False
    
    def get_active_validations(self) -> List[Dict]:
        """Récupère toutes les validations actives"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute('''
                SELECT dossier, original_rating, validated_rating, validation_status,
                       validation_reason, validator, validation_date, inconsistency_type,
                       detected_words
                FROM inconsistency_validations
                WHERE validation_status IN ('validated', 'ignored')
                ORDER BY updated_at DESC
            ''')
            
            validations = []
            for row in cursor.fetchall():
                validations.append({
                    'dossier': row[0],
                    'original_rating': row[1],
                    'validated_rating': row[2],
                    'validation_status': row[3],
                    'validation_reason': row[4],
                    'validator': row[5],
                    'validation_date': row[6],
                    'inconsistency_type': row[7],
                    'detected_words': json.loads(row[8]) if row[8] else []
                })
            
            conn.close()
            return validations
            
        except Exception as e:
            print(f"⚠️ Erreur récupération validations: {e}")
            return []
    
    def apply_validations_to_dataframe(self, df_merged: pd.DataFrame) -> pd.DataFrame:
        """Applique TOUTES les validations actives au DataFrame"""
        df_modified = df_merged.copy()
        
        # Ajouter colonnes de traçabilité
        if 'Original_Rating' not in df_modified.columns:
            df_modified['Original_Rating'] = df_modified['Valeur de chaîne']
        if 'Validation_Applied' not in df_modified.columns:
            df_modified['Validation_Applied'] = False
        if 'Validation_Reason' not in df_modified.columns:
            df_modified['Validation_Reason'] = ''
        if 'Validation_Date' not in df_modified.columns:
            df_modified['Validation_Date'] = ''
        
        # Récupérer toutes les validations actives
        validations = self.get_active_validations()
        
        modifications_count = 0
        for validation in validations:
            if validation['validation_status'] in ['validated', 'ignored']:
                # Trouver les lignes correspondantes
                mask = df_modified['Dossier Rcbt'] == validation['dossier']
                matching_rows = df_modified[mask]
                
                if not matching_rows.empty:
                    # Pour les validations "ignored", conserver l'original mais marquer comme traité
                    if validation['validation_status'] == 'ignored':
                        # Note reste identique, mais marquer comme contrôlé
                        validation_reason = validation['validation_reason'] or "Conserver l'original après contrôle"
                    else:
                        # Appliquer la nouvelle note validée
                        df_modified.loc[mask, 'Valeur de chaîne'] = validation['validated_rating']
                        validation_reason = validation['validation_reason'] or "Correction appliquée"
                    
                    df_modified.loc[mask, 'Validation_Applied'] = True
                    df_modified.loc[mask, 'Validation_Reason'] = validation_reason
                    df_modified.loc[mask, 'Validation_Date'] = validation['validation_date']
                    modifications_count += len(matching_rows)
        
        print(f"✅ Validations centralisées appliquées: {modifications_count} lignes modifiées")
        return df_modified
    
    def get_validation_summary(self) -> Dict[str, Any]:
        """Retourne un résumé des validations actives"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute('''
                SELECT validation_status, COUNT(*) 
                FROM inconsistency_validations 
                GROUP BY validation_status
            ''')
            
            status_counts = dict(cursor.fetchall())
            total = sum(status_counts.values())
            
            conn.close()
            
            return {
                'total_detected': total,
                'validated_corrected': status_counts.get('validated', 0),
                'validated_ignored': status_counts.get('ignored', 0),
                'pending_validation': status_counts.get('pending', 0),
                'completion_rate': round((status_counts.get('validated', 0) + status_counts.get('ignored', 0)) / total * 100, 1) if total > 0 else 0
            }
            
        except Exception as e:
            print(f"⚠️ Erreur calcul résumé: {e}")
            return {
                'total_detected': 0,
                'validated_corrected': 0,
                'validated_ignored': 0,
                'pending_validation': 0,
                'completion_rate': 0
            }
    
    def clear_old_sessions(self, days_old: int = 7):
        """Nettoie les anciennes sessions de validation"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute('''
                DELETE FROM inconsistency_validations 
                WHERE created_at < datetime('now', '-{} days')
            '''.format(days_old))
            
            cursor.execute('''
                DELETE FROM validation_sessions 
                WHERE created_at < datetime('now', '-{} days')
            '''.format(days_old))
            
            conn.commit()
            conn.close()
            
            print(f"✅ Anciennes sessions supprimées (plus de {days_old} jours)")
            
        except Exception as e:
            print(f"⚠️ Erreur nettoyage: {e}")

# Instance globale pour l'application
validation_db = ValidationDatabase()