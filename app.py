from flask import Flask, render_template, request, jsonify, send_file, flash, redirect, url_for, session
import pandas as pd
import xlrd  # Support pour fichiers .xls
import os
from werkzeug.utils import secure_filename
from pathlib import Path
import sqlite3
from datetime import datetime
import tempfile
import shutil
from utils.data_processor import DataProcessor
from utils.report_generator import ReportGenerator
from utils.html_report_generator import HTMLReportGenerator
from utils.inconsistency_validator import InconsistencyValidator
import traceback
import json

app = Flask(__name__)
app.secret_key = os.getenv('SECRET_KEY', 'rcbt-app-secret-key-2025')

# AJOUT: Message d'aide pour CMD si interface fermée
def display_interface_help():
    """Affiche un message d'aide si l'interface web est fermée"""
    print("\n" + "="*60)
    print("🌐 INTERFACE WEB FERMÉE")
    print("="*60)
    print("💡 Pour rouvrir l'interface de l'application :")
    print("   👉 Ouvrez votre navigateur")
    print("   👉 Allez sur : http://localhost:5000")
    print("   👉 Ou sur : http://127.0.0.1:5000")
    print("="*60)
    print("⚠️  IMPORTANT :")
    print("   • NE FERMEZ PAS cette fenêtre CMD")
    print("   • Votre session sera perdue si vous fermez la CMD")
    print("   • L'application continue de fonctionner en arrière-plan")
    print("="*60)
    print("🔄 Application toujours active - Attendant connexions...")

# Configuration
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}
MAX_CONTENT_LENGTH = 50 * 1024 * 1024  # 50MB max file size

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = MAX_CONTENT_LENGTH

# Ensure upload directory exists
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs('ressources', exist_ok=True)

# Global cache for uploaded files (to fix session issue)
uploaded_files_cache = {}

# Global validator instance for inconsistency handling
inconsistency_validator = InconsistencyValidator()

# Session persistence system
SESSION_CACHE_DIR = 'session_cache'
os.makedirs(SESSION_CACHE_DIR, exist_ok=True)

def get_session_id():
    """Obtenir ou créer un ID de session unique"""
    if 'session_id' not in session:
        session['session_id'] = f"session_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{os.urandom(4).hex()}"
    return session['session_id']

def save_session_data(data, data_type='files'):
    """Sauvegarder les données de session sur disque"""
    try:
        session_id = get_session_id()
        cache_file = os.path.join(SESSION_CACHE_DIR, f"{session_id}_{data_type}.pkl")
        import pickle
        with open(cache_file, 'wb') as f:
            pickle.dump(data, f)
        return True
    except Exception as e:
        print(f"Erreur sauvegarde session: {e}")
        return False

def load_session_data(data_type='files'):
    """Charger les données de session depuis le disque"""
    try:
        session_id = get_session_id()
        cache_file = os.path.join(SESSION_CACHE_DIR, f"{session_id}_{data_type}.pkl")
        if os.path.exists(cache_file):
            import pickle
            with open(cache_file, 'rb') as f:
                return pickle.load(f)
    except Exception as e:
        print(f"Erreur chargement session: {e}")
    return {}

def copy_report_to_resources(temp_path):
    """Copier un rapport temporaire vers le dossier ressources pour persistance"""
    try:
        if os.path.exists(temp_path):
            # Nettoyer le nom de fichier des caractères de contrôle
            filename = os.path.basename(temp_path).replace('\r', '').replace('\n', '').strip()
            
            # Créer le dossier ressources s'il n'existe pas
            resources_dir = 'ressources'
            if not os.path.exists(resources_dir):
                os.makedirs(resources_dir)
            
            resource_path = os.path.join(resources_dir, filename)
            shutil.copy2(temp_path, resource_path)
            
            # Retourner le chemin propre
            return resource_path.replace('\\', '/').replace('\r', '').replace('\n', '')
        return None
    except Exception as e:
        print(f"Erreur copie rapport: {e}")
        return None

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def init_database():
    """Initialize SQLite database for storing history"""
    conn = sqlite3.connect('rcbt_history.db')
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS reports (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            timestamp TEXT NOT NULL,
            filename TEXT NOT NULL,
            total_tickets INTEGER,
            tickets_boutiques INTEGER,
            taux_closure REAL,
            taux_satisfaction REAL,
            comments_with_percentage REAL,
            file_path TEXT
        )
    ''')
    
    # Add new columns if they don't exist (for database migration)
    try:
        cursor.execute('ALTER TABLE reports ADD COLUMN report_type TEXT DEFAULT "global"')
    except:
        pass  # Column already exists
    try:
        cursor.execute('ALTER TABLE reports ADD COLUMN filter_type TEXT')
    except:
        pass  # Column already exists
    try:
        cursor.execute('ALTER TABLE reports ADD COLUMN filter_value TEXT')
    except:
        pass  # Column already exists
    conn.commit()
    conn.close()

@app.route('/')
def index():
    # Restaurer automatiquement les fichiers depuis la session si ils existent
    global uploaded_files_cache
    session_files = None
    if not uploaded_files_cache:
        session_files = load_session_data('files')
        if session_files:
            uploaded_files_cache = session_files
            print(f"✅ Session restaurée: {len(session_files)} fichiers récupérés")
    
    # Vérifier s'il y a des incohérences en attente de validation
    pending_inconsistencies = load_session_data('inconsistencies')
    has_pending_validation = bool(pending_inconsistencies)
    
    # Vérifier s'il y a des données pour les rapports individuels
    individual_data_available = bool(uploaded_files_cache)
    session_restored = bool(session_files)
    
    return render_template('index.html', 
                         files_loaded=individual_data_available,
                         session_restored=session_restored,
                         has_pending_validation=has_pending_validation)

@app.route('/upload', methods=['POST'])
def upload_files():
    try:
        # Check if all required files are present
        required_files = ['enq_file', 'case_file', 'ref_file', 'acct_file']
        uploaded_files = {}
        
        for file_key in required_files:
            if file_key not in request.files:
                return jsonify({'success': False, 'error': f'Fichier manquant: {file_key}'})
            
            file = request.files[file_key]
            if file.filename == '':
                return jsonify({'success': False, 'error': f'Aucun fichier sélectionné pour {file_key}'})
            
            if not allowed_file(file.filename):
                return jsonify({'success': False, 'error': f'Type de fichier non autorisé: {file.filename}'})
            
            # Save file
            filename = secure_filename(file.filename) if file.filename else 'unknown'
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f"{timestamp}_{filename}"
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)
            uploaded_files[file_key] = filepath
        
        # Sauvegarder dans le cache global ET la session persistante
        global uploaded_files_cache
        uploaded_files_cache = uploaded_files
        save_session_data(uploaded_files, 'files')
        
        return jsonify({
            'success': True, 
            'message': 'Tous les fichiers ont été téléchargés avec succès',
            'files': uploaded_files
        })
        
    except Exception as e:
        return jsonify({'success': False, 'error': f'Erreur lors du téléchargement: {str(e)}'})

@app.route('/generate_report', methods=['POST'])
def generate_report():
    try:
        # Charger les fichiers depuis la session si le cache est vide
        global uploaded_files_cache
        if not uploaded_files_cache:
            uploaded_files_cache = load_session_data('files')
            
        if not uploaded_files_cache:
            return jsonify({'success': False, 'error': 'Aucun fichier chargé. Veuillez recharger vos fichiers.'})
            
        data = request.get_json()
        files = data.get('files', {})
        
        if not files or len(files) != 4:
            return jsonify({'success': False, 'error': 'Tous les fichiers sont requis'})
        
        # Process data
        processor = DataProcessor()
        merged_data = processor.load_and_merge_data(files)
        
        # Détecter les incohérences dans les données
        inconsistencies_data = merged_data.get('inconsistencies', [])
        
        if inconsistencies_data and len(inconsistencies_data) > 0:
            # Sauvegarder les fichiers dans le cache global pour génération finale
            uploaded_files_cache.update(files)
            
            # Charger les incohérences dans le validateur ET la base centralisée
            inconsistency_validator.load_inconsistencies(inconsistencies_data)
            
            # Créer session centralisée pour persistance garantie
            from utils.validation_database import validation_db
            session_id = validation_db.create_validation_session(inconsistencies_data)
            
            return jsonify({
                'success': True,
                'inconsistencies_detected': len(inconsistencies_data),
                'requires_validation': True,
                'message': f'{len(inconsistencies_data)} incohérences détectées nécessitent une validation manuelle avant génération.',
                'redirect_to': '/validate_inconsistencies'
            })
        
        # Générer rapport directement si aucune incohérence
        from utils.optimized_html_generator import OptimizedReportGenerator
        optimized_gen = OptimizedReportGenerator()
        report_path = optimized_gen.generate_optimized_report(merged_data)
        
        # Copier le rapport vers ressources pour persistance
        persistent_path = copy_report_to_resources(report_path)
        final_path = persistent_path if persistent_path else report_path
        
        # Calculate metrics for history
        page1_metrics = processor.calculate_page1_metrics(merged_data)
        page2_metrics = processor.calculate_page2_enhanced_metrics(merged_data)
        
        # SUPPRESSION: Maintenant nb_reponses_q1 est calculé directement dans calculate_page1_metrics
        # Plus besoin de ce calcul manuel car c'est maintenant dans page1_metrics
        
        # Sauvegarder dans l'historique avec le chemin persistant
        save_to_history(final_path, page1_metrics, page2_metrics)
        
        return jsonify({
            'success': True,
            'message': 'Rapport généré avec succès',
            'report_path': report_path,
            'metrics': {
                'page1': page1_metrics,
                'page2': page2_metrics
            }
        })
        
    except Exception as e:
        error_msg = f'Erreur lors de la génération du rapport: {str(e)}'
        print(f"Error: {error_msg}")
        print(traceback.format_exc())
        return jsonify({'success': False, 'error': error_msg})

@app.route('/check_session_data')
def check_session_data():
    """Vérifier si des données session existent pour rapports individuels"""
    try:
        global uploaded_files_cache
        session_files = load_session_data('files')
        
        # Restaurer dans le cache global si session existe mais cache vide
        if session_files and not uploaded_files_cache:
            uploaded_files_cache = session_files
            print(f"✅ Session restaurée: {len(session_files)} fichiers récupérés")
        
        if session_files:
            return jsonify({
                'success': True, 
                'has_files': True,
                'files': session_files
            })
        else:
            return jsonify({
                'success': True,
                'has_files': False
            })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})

@app.route('/get_available_data', methods=['POST'])
def get_available_data():
    try:
        data = request.get_json()
        files = data.get('files', {})
        
        if not files or len(files) != 4:
            return jsonify({'success': False, 'error': 'Tous les fichiers sont requis'})
        
        # Process data to get available sites and collaborators
        processor = DataProcessor()
        merged_data = processor.load_and_merge_data(files)
        
        # Extract unique sites and collaborators
        sites = list(merged_data['merged']['Site'].dropna().unique())
        sites = [str(site) for site in sites if site != 'Autres']
        sites.sort()
        
        collaborators = list(merged_data['merged']['Créé par ticket'].dropna().unique())
        collaborators = [str(collab) for collab in collaborators]
        collaborators.sort()
        
        return jsonify({
            'success': True,
            'data': {
                'sites': sites,
                'collaborators': collaborators
            }
        })
        
    except Exception as e:
        error_msg = f'Erreur lors du chargement des données: {str(e)}'
        print(f"Error: {error_msg}")
        return jsonify({'success': False, 'error': error_msg})

@app.route('/generate_individual_report', methods=['POST'])
def generate_individual_report():
    try:
        # Charger les fichiers depuis la session si le cache est vide
        global uploaded_files_cache
        if not uploaded_files_cache:
            uploaded_files_cache = load_session_data('files')
            
        # CORRECTION 1: Récupérer data depuis JSON comme dans main.js
        data = request.get_json()
        if not data:
            return jsonify({'success': False, 'error': 'Données JSON requises'})
            
        report_type = data.get('report_type')
        report_target = data.get('report_target')
        files = data.get('files', uploaded_files_cache)
        
        print(f"🔍 Génération rapport individuel: {report_type} - {report_target}")
        
        if not report_type or not report_target:
            return jsonify({'success': False, 'error': 'Type et cible du rapport requis'})
        
        # Vérifier que les fichiers sont fournis
        if not files or len(files) != 4:
            return jsonify({'success': False, 'error': 'Tous les fichiers sont requis'})
        
        # Process data
        processor = DataProcessor()
        merged_data = processor.load_and_merge_data(files)
        
        # Filter data based on report type and target
        filtered_data = filter_data_for_individual_report(merged_data, report_type, report_target)
        
        # Ensure numeric columns are properly typed to prevent dtype errors
        if 'Code compte' in filtered_data['merged'].columns:
            filtered_data['merged']['Code compte'] = pd.to_numeric(filtered_data['merged']['Code compte'], errors='coerce')
        if 'Code compte' in filtered_data['case'].columns:
            filtered_data['case']['Code compte'] = pd.to_numeric(filtered_data['case']['Code compte'], errors='coerce')
        if 'Code compte' in filtered_data['accounts'].columns:
            filtered_data['accounts']['Code compte'] = pd.to_numeric(filtered_data['accounts']['Code compte'], errors='coerce')
        
        # Generate individual HTML report
        from utils.individual_report_generator import IndividualReportGenerator
        individual_gen = IndividualReportGenerator()
        report_path = individual_gen.generate_individual_report(filtered_data, report_type, report_target)
        
        # Calculer les métriques pour l'historique des rapports individuels
        page1_metrics = processor.calculate_page1_metrics(filtered_data)
        page2_metrics = processor.calculate_page2_enhanced_metrics(filtered_data)
        
        # Copier le rapport vers le dossier ressources pour historique  
        ressources_dir = os.path.join(os.getcwd(), 'ressources')
        os.makedirs(ressources_dir, exist_ok=True)
        
        # Nom de fichier propre pour ressources
        clean_filename = os.path.basename(report_path)
        clean_filename = clean_filename.replace('\x0c', '').replace('\x0d', '').replace('\r', '').replace('\n', '').strip()
        ressources_path = os.path.join(ressources_dir, clean_filename)
        
        try:
            import shutil
            shutil.copy2(report_path, ressources_path)
            print(f"✅ Rapport copié vers ressources: {ressources_path}")
        except Exception as e:
            print(f"⚠️ Erreur copie vers ressources: {e}")
            ressources_path = report_path
        
        # CORRECTION PROBLÈME 2: Ne plus recalculer nb_reponses_q1 ici car déjà fait dans calculate_page1_metrics
        # Cette duplication causait des incohérences dans l'historique
        print(f"✅ nb_reponses_q1 calculé dans calculate_page1_metrics: {page1_metrics.get('nb_reponses_q1', 'N/A')}")
        
        # Sauvegarder dans l'historique avec le chemin ressources
        save_to_history(
            clean_filename,  # Juste le nom pour l'historique
            page1_metrics, 
            page2_metrics, 
            report_type='individual',
            filter_type=report_type,
            filter_value=report_target
        )
        
        return jsonify({
            'success': True,
            'message': f'Rapport individuel généré avec succès: {report_type} - {report_target}',
            'report_path': report_path
        })
        
    except Exception as e:
        error_msg = f'Erreur lors de la génération du rapport individuel: {str(e)}'
        print(f"Error: {error_msg}")
        print(traceback.format_exc())
        return jsonify({'success': False, 'error': error_msg})

@app.route('/download_report/<path:filename>')
def download_report(filename):
    try:
        # Handle encoded paths from frontend
        decoded_filename = filename.replace('%2F', '/').replace('%2f', '/')
        
        # Nettoyer le nom de fichier et chercher dans plusieurs emplacements
        clean_filename = Path(decoded_filename).name
        temp_dir = tempfile.gettempdir()
        
        # Nettoyer le nom de fichier de caractères de contrôle
        clean_filename = clean_filename.replace('\r', '').replace('\n', '').strip()
        
        possible_paths = [
            # Priorité 1: ressources local
            os.path.join('ressources', clean_filename),
            # Priorité 2: répertoire temp système 
            os.path.join(temp_dir, clean_filename),
            # Priorité 3: chemins directs si pas corrompus
            decoded_filename if not '\r' in decoded_filename else None,
            filename if not '\r' in filename else None,
            # Priorité 4: chemins absolus
            os.path.abspath(os.path.join('ressources', clean_filename)),
            os.path.abspath(os.path.join(temp_dir, clean_filename))
        ]
        
        # Filtrer les chemins None
        possible_paths = [p for p in possible_paths if p is not None]
        
        actual_file_path = None
        for path in possible_paths:
            if os.path.exists(path):
                actual_file_path = path
                break
        
        if not actual_file_path:
            print(f"File not found in any location: {possible_paths}")
            return jsonify({'success': False, 'error': 'Fichier non trouvé'}), 404
        
        print(f"Serving file: {actual_file_path}")
        
        # Determine MIME type based on file extension
        if actual_file_path.endswith('.html'):
            return send_file(actual_file_path, as_attachment=True, download_name=Path(actual_file_path).name, mimetype='text/html')
        elif actual_file_path.endswith('.pdf'):
            return send_file(actual_file_path, as_attachment=True, download_name=Path(actual_file_path).name, mimetype='application/pdf')
        else:
            return send_file(actual_file_path, as_attachment=True, download_name=Path(actual_file_path).name)
    except Exception as e:
        print(f"Download error: {str(e)}")
        return jsonify({'success': False, 'error': f'Erreur lors du téléchargement: {str(e)}'}), 500

@app.route('/history')
def get_history():
    try:
        conn = sqlite3.connect('rcbt_history.db')
        cursor = conn.cursor()
        cursor.execute('''
            SELECT id, timestamp, filename, total_tickets, tickets_boutiques, nb_reponses_q1,
                   taux_closure, taux_satisfaction, comments_with_percentage, file_path, report_type, filter_type, filter_value
            FROM reports 
            ORDER BY timestamp DESC 
            LIMIT 50
        ''')
        rows = cursor.fetchall()
        conn.close()
        
        history = []
        for row in rows:
            # Gérer les anciennes entrées qui n'ont pas nb_reponses_q1
            if len(row) >= 10:
                nb_reponses_q1 = row[5] if row[5] is not None else 0
                report_type = row[10] if len(row) > 10 else 'global'
                filter_type = row[11] if len(row) > 11 else None
                filter_value = row[12] if len(row) > 12 else None
            else:
                nb_reponses_q1 = row[4]  # Fallback vers tickets_boutiques pour ancien format
                report_type = 'global'
                filter_type = None
                filter_value = None
            
            history.append({
                'id': row[0],
                'timestamp': row[1],
                'filename': row[2],
                'total_tickets': row[3],
                'tickets_boutiques': row[4],
                'nb_reponses_q1': nb_reponses_q1,
                'taux_closure': row[6] if len(row) > 6 else (row[5] if len(row) > 5 else 0),
                'taux_satisfaction': row[7] if len(row) > 7 else (row[6] if len(row) > 6 else 0),
                'comments_with_percentage': row[8] if len(row) > 8 else (row[7] if len(row) > 7 else 0),
                'file_path': row[9] if len(row) > 9 else (row[8] if len(row) > 8 else ''),
                'report_type': report_type,
                'filter_type': filter_type,
                'filter_value': filter_value
            })
        
        return jsonify({'success': True, 'history': history})
        
    except Exception as e:
        return jsonify({'success': False, 'error': f'Erreur lors de la récupération de l\'historique: {str(e)}'})

def filter_data_for_individual_report(merged_data, report_type, report_target):
    """Filter data for individual reports by site or collaborator"""
    
    # CORRECTION: Appliquer les validations d'incohérences centralisées AVANT le filtrage
    from utils.validation_database import validation_db
    
    # Appliquer TOUTES les validations actives depuis la base centralisée
    try:
        merged_with_validations = validation_db.apply_validations_to_dataframe(merged_data['merged'])
        print(f"✅ Validations centralisées appliquées aux rapports individuels")
    except Exception as e:
        print(f"⚠️ Pas de validations centralisées à appliquer: {e}")
        merged_with_validations = merged_data['merged']
    
    filtered_merged = merged_with_validations.copy()
    filtered_case = merged_data['case'].copy()
    
    # Add site mapping to case data if needed
    if report_type == 'site' and 'Site' not in filtered_case.columns:
        # Vérifier si 'ref' existe dans merged_data
        if 'ref' in merged_data:
            df_ref = merged_data['ref']
            if not df_ref.empty and 'Login' in df_ref.columns and 'Site' in df_ref.columns:
                site_mapping = dict(zip(df_ref['Login'], df_ref['Site']))
                filtered_case['Site'] = filtered_case['Créé par ticket'].map(site_mapping)
                filtered_case['Site'] = filtered_case['Site'].fillna('Autres')
        else:
            # Si pas de données ref, utiliser le site depuis merged qui contient déjà cette info
            if 'Site' in filtered_merged.columns:
                site_mapping = dict(zip(filtered_merged['Créé par ticket'], filtered_merged['Site']))
                filtered_case['Site'] = filtered_case['Créé par ticket'].map(site_mapping)
                filtered_case['Site'] = filtered_case['Site'].fillna('Autres')
    
    if report_type == 'site':
        # Filter by site
        filtered_merged = filtered_merged[filtered_merged['Site'] == report_target]
        filtered_case = filtered_case[filtered_case['Site'] == report_target] if 'Site' in filtered_case.columns else filtered_case
    elif report_type == 'collaborator':
        # Filter by collaborator
        filtered_merged = filtered_merged[filtered_merged['Créé par ticket'] == report_target]
        filtered_case = filtered_case[filtered_case['Créé par ticket'] == report_target]
    else:
        raise ValueError(f"Type de rapport non supporté: {report_type}")
    
    if filtered_merged.empty:
        raise ValueError(f"Aucune donnée trouvée pour {report_type}: {report_target}")
    
    print(f"🔍 FILTRAGE APPLIQUÉ: {report_type} = {report_target}")
    print(f"   - Données merged filtrées: {len(filtered_merged)} lignes")
    print(f"   - Données case filtrées: {len(filtered_case)} lignes")
    
    # Construire le dictionnaire de retour avec vérification des clés
    result = {
        'merged': filtered_merged,
        'case': filtered_case,  # CORRIGÉ: case est maintenant filtré selon le contexte
        'filter_type': report_type,
        'filter_value': report_target
    }
    
    # Ajouter les clés optionnelles seulement si elles existent
    if 'accounts' in merged_data:
        result['accounts'] = merged_data['accounts']
    if 'ref' in merged_data:
        result['ref'] = merged_data['ref']
    
    return result

def save_to_history(report_path, page1_metrics, page2_metrics, report_type='global', filter_type=None, filter_value=None):
    """Save report metrics to history database"""
    try:
        conn = sqlite3.connect('rcbt_history.db')
        cursor = conn.cursor()
        
        timestamp = datetime.now().isoformat()
        filename = Path(report_path).name
        
        # Nettoyer le chemin de fichier avant sauvegarde pour éviter caractères de contrôle
        clean_report_path = report_path.replace('\r', '').replace('\n', '').strip()
        clean_filename = filename.replace('\r', '').replace('\n', '').strip()
        
        # Ajouter colonne nb_reponses_q1 si elle n'existe pas
        try:
            cursor.execute('ALTER TABLE reports ADD COLUMN nb_reponses_q1 INTEGER DEFAULT 0')
        except sqlite3.OperationalError:
            pass  # Column already exists
        
        cursor.execute('''
            INSERT INTO reports (timestamp, filename, report_type, filter_type, filter_value,
                               total_tickets, tickets_boutiques, nb_reponses_q1, taux_closure, taux_satisfaction, 
                               comments_with_percentage, file_path)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            timestamp,
            clean_filename,
            report_type,
            filter_type,
            filter_value,
            page1_metrics.get('total_tickets', 0),
            page1_metrics.get('tickets_boutiques', 0),
            page1_metrics.get('nb_reponses_q1', 0),
            page1_metrics.get('taux_closure', 0),
            page1_metrics.get('taux_sat', 0),
            page2_metrics.get('comments_percentage', 0),
            clean_report_path
        ))
        
        conn.commit()
        conn.close()
        
    except Exception as e:
        print(f"Error saving to history: {e}")

@app.route('/reports_history')
def get_reports_history():
    """Get all reports history"""
    try:
        # S'assurer que la base est à jour
        ensure_database_schema()
        
        conn = sqlite3.connect('rcbt_history.db')
        cursor = conn.cursor()
        cursor.execute('''
            SELECT id, timestamp, filename, total_tickets, tickets_boutiques, 
                   COALESCE(nb_reponses_q1, 0) as nb_reponses_q1, taux_closure, taux_satisfaction, 
                   comments_with_percentage, file_path, report_type, filter_type, filter_value
            FROM reports ORDER BY timestamp DESC
        ''')
        reports = cursor.fetchall()
        conn.close()
        
        history = []
        for report in reports:
            history.append({
                'id': report[0],
                'timestamp': report[1],
                'filename': report[2],
                'total_tickets': report[3],
                'tickets_boutiques': report[4],
                'nb_reponses_q1': report[5],  # CORRECTION: Position corrigée
                'taux_closure': report[6],
                'taux_satisfaction': report[7],
                'comments_percentage': report[8],
                'file_path': report[9],
                'report_type': report[10] or 'global',
                'filter_type': report[11],
                'filter_value': report[12]
            })
        
        return jsonify({'success': True, 'history': history})
        
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})

@app.route('/delete_history/<int:report_id>', methods=['DELETE'])
def delete_history_item(report_id):
    """Delete a specific report from history"""
    try:
        conn = sqlite3.connect('rcbt_history.db')
        cursor = conn.cursor()
        
        # Get file path before deletion
        cursor.execute('SELECT file_path FROM reports WHERE id = ?', (report_id,))
        result = cursor.fetchone()
        
        if result:
            file_path = result[0]
            
            # Delete from database
            cursor.execute('DELETE FROM reports WHERE id = ?', (report_id,))
            conn.commit()
            
            # Delete physical file if it exists
            if file_path and os.path.exists(file_path):
                try:
                    os.remove(file_path)
                except Exception as e:
                    print(f"Warning: Could not delete file {file_path}: {e}")
            
            conn.close()
            return jsonify({'success': True, 'message': 'Rapport supprimé avec succès'})
        else:
            conn.close()
            return jsonify({'success': False, 'error': 'Rapport non trouvé'})
            
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})

@app.route('/status')
def status():
    return jsonify({
        'status': 'running',
        'version': '2.0.0',
        'timestamp': datetime.now().isoformat()
    })

# Routes pour la validation des incohérences
@app.route('/validate_inconsistencies')
def validate_inconsistencies():
    """Interface de validation des incohérences détectées"""
    try:
        # Récupérer les incohérences en attente
        pending_inconsistencies = inconsistency_validator.get_pending_inconsistencies()
        all_inconsistencies = inconsistency_validator.inconsistencies
        summary = inconsistency_validator.get_validation_summary()
        
        return render_template(
            'validate_inconsistencies.html',
            inconsistencies=all_inconsistencies,
            summary=summary
        )
    except Exception as e:
        flash(f'Erreur lors du chargement des incohérences: {str(e)}')
        return redirect(url_for('index'))

@app.route('/validate_inconsistency', methods=['POST'])
def validate_inconsistency():
    """Valider une incohérence spécifique"""
    try:
        data = request.get_json()
        dossier = data.get('dossier')
        validated_rating = data.get('validated_rating')
        reason = data.get('reason', '')
        
        # Validation dans l'ancien système ET le nouveau système centralisé
        success_old = inconsistency_validator.validate_inconsistency(
            dossier, validated_rating, reason
        )
        
        # Validation dans le système centralisé pour persistance garantie
        from utils.validation_database import validation_db
        success_new = validation_db.save_validation(
            dossier, validated_rating, "validated", reason, "User"
        )
        
        if success_old or success_new:
            return jsonify({'success': True, 'message': 'Validation sauvegardée dans les deux systèmes'})
        else:
            return jsonify({'success': False, 'error': 'Dossier non trouvé'})
            
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})

@app.route('/ignore_inconsistency', methods=['POST'])
def ignore_inconsistency():
    """Ignorer une incohérence (conserver note originale)"""
    try:
        data = request.get_json()
        dossier = data.get('dossier')
        reason = data.get('reason', '')
        
        # Ignorer dans l'ancien système ET le nouveau système centralisé
        success_old = inconsistency_validator.ignore_inconsistency(
            dossier, reason
        )
        
        # Ignorer dans le système centralisé pour persistance garantie
        from utils.validation_database import validation_db
        # Pour ignorer, on garde la note originale
        original_rating = next((inc.original_rating for inc in inconsistency_validator.inconsistencies if inc.dossier == dossier), "")
        success_new = validation_db.save_validation(
            dossier, original_rating, "ignored", reason, "User"
        )
        
        if success_old or success_new:
            return jsonify({'success': True, 'message': 'Incohérence ignorée dans les deux systèmes'})
        else:
            return jsonify({'success': False, 'error': 'Dossier non trouvé'})
            
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})

@app.route('/generate_final_reports')
def generate_final_reports():
    """Générer les rapports finaux avec validations appliquées"""
    try:
        # Vérifier qu'il n'y a plus d'incohérences en attente
        pending = inconsistency_validator.get_pending_inconsistencies()
        if len(pending) > 0:
            flash(f'{len(pending)} incohérences restent à traiter avant génération finale.')
            return redirect(url_for('validate_inconsistencies'))
        
        # Appliquer les validations aux données et générer rapports finaux
        # Récupérer les fichiers depuis le cache global ou session pour la génération finale
        global uploaded_files_cache
        if not uploaded_files_cache:
            uploaded_files_cache = load_session_data('files')
            
        if uploaded_files_cache:
            processor = DataProcessor()
            merged_data = processor.load_and_merge_data(uploaded_files_cache)
            
            # Appliquer les validations centralisées aux données avec traçabilité garantie
            from utils.validation_database import validation_db
            validated_data = validation_db.apply_validations_to_dataframe(merged_data['merged'])
            merged_data['merged'] = validated_data
            
            # Sauvegarder les validations dans la session pour traçabilité
            save_session_data({
                'validations': inconsistency_validator.inconsistencies,
                'validated_data': True
            }, 'validations')
            
            # Générer rapport final avec validations appliquées
            from utils.optimized_html_generator import OptimizedReportGenerator
            optimized_gen = OptimizedReportGenerator()
            report_path = optimized_gen.generate_optimized_report(merged_data)
            
            # Calculer métriques pour historique
            page1_metrics = processor.calculate_page1_metrics(merged_data)
            page2_metrics = processor.calculate_page2_enhanced_metrics(merged_data)
            
            # Copier le rapport vers ressources pour persistance
            persistent_path = copy_report_to_resources(report_path)
            final_path = persistent_path if persistent_path else report_path
            
            # Nettoyer le chemin avant sauvegarde pour éviter caractères de contrôle
            clean_final_path = final_path.replace('\r', '').replace('\n', '').strip()
            
            # Sauvegarder dans l'historique avec le chemin persistant nettoyé
            save_to_history(clean_final_path, page1_metrics, page2_metrics)
            
            # Préparer les données pour la page de succès
            summary = inconsistency_validator.get_validation_summary()
            validation_date = datetime.now().strftime('%d/%m/%Y à %H:%M')
            
            return render_template(
                'validation_success.html',
                summary=summary,
                report_path=os.path.basename(report_path),
                validation_date=validation_date
            )
        else:
            flash('Erreur: données de fichiers non trouvées. Veuillez recharger les fichiers.', 'error')
            return redirect(url_for('index'))
        
    except Exception as e:
        flash(f'Erreur lors de la génération finale: {str(e)}')
        return redirect(url_for('validate_inconsistencies'))

@app.route('/export_validation_log')
def export_validation_log():
    """Exporter le log de validation en JSON"""
    try:
        validation_log = inconsistency_validator.export_validation_log()
        
        # Sauvegarder dans ressources
        log_filename = f"validation_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        log_path = os.path.join('ressources', log_filename)
        
        with open(log_path, 'w', encoding='utf-8') as f:
            json.dump(validation_log, f, ensure_ascii=False, indent=2)
        
        return send_file(log_path, as_attachment=True)
        
    except Exception as e:
        flash(f'Erreur lors de l\'export: {str(e)}')
        return redirect(url_for('validate_inconsistencies'))

@app.route('/download/<filename>')
def download_file(filename):
    """Télécharger un fichier depuis le dossier ressources"""
    try:
        file_path = os.path.join('ressources', filename)
        if os.path.exists(file_path):
            return send_file(file_path, as_attachment=True)
        else:
            flash('Fichier non trouvé', 'error')
            return redirect(url_for('index'))
    except Exception as e:
        flash(f'Erreur lors du téléchargement: {str(e)}', 'error')
        return redirect(url_for('index'))

def ensure_database_schema():
    """Assurer la compatibilité de la base de données"""
    try:
        conn = sqlite3.connect('rcbt_history.db')
        cursor = conn.cursor()
        
        # Vérifier si la colonne nb_reponses_q1 existe
        cursor.execute("PRAGMA table_info(reports)")
        columns = [col[1] for col in cursor.fetchall()]
        
        if 'nb_reponses_q1' not in columns:
            print("📊 Migration base de données: ajout colonne nb_reponses_q1")
            cursor.execute('ALTER TABLE reports ADD COLUMN nb_reponses_q1 INTEGER DEFAULT 0')
            conn.commit()
            print("✅ Migration terminée")
        
        conn.close()
    except Exception as e:
        print(f"⚠️ Erreur migration base: {e}")

if __name__ == '__main__':
    init_database()
    ensure_database_schema()
    print("\n" + "="*60)
    print("🚀 APPLICATION RCBT - PRÊTE À L'USAGE")
    print("   Version avec Validation des Incohérences")
    print("="*60)
    print("📋 Fonctionnalités disponibles :")
    print("   ✓ Génération de rapports RCBT complets")
    print("   ✓ Validation manuelle des incohérences")  
    print("   ✓ Interface web intuitive")
    print("   ✓ Traçabilité complète des corrections")
    print("-"*60)
    print("🌐 Accédez à l'application sur :")
    print("   👉 http://localhost:5000")
    print("   👉 http://127.0.0.1:5000")
    print("-"*60)
    print("💡 Utilisation :")
    print("   1. Chargez vos 4 fichiers Excel")
    print("   2. Générez le rapport")
    print("   3. Validez les incohérences si détectées")
    print("="*60)
    print("🔄 Démarrage du serveur...")
    print()
    
@app.route('/archive')
def download_clean_archive():
    """Route pour télécharger l'archive propre de l'application"""
    try:
        archive_name = "RCBT_Application_PROPRE_14082025.tar.gz"
        file_path = os.path.join(os.getcwd(), archive_name)
        
        if os.path.exists(file_path):
            return send_file(file_path, as_attachment=True, download_name=archive_name)
        else:
            return f"<h1>Archive non trouvée</h1><p>Le fichier {archive_name} n'existe pas.</p>", 404
            
    except Exception as e:
        return f"<h1>Erreur</h1><p>Erreur lors du téléchargement: {str(e)}</p>", 500

def display_startup_banner():
    """Affiche le message de démarrage professionnel unique"""
    print("\n" + "="*70)
    print("                       -- ISESIR --")
    print("           by Pole QoS et Transformation SI Retail")
    print("               Réseau Clubs Bouygues Telecom")
    print("="*70)
    print("📊 Analyse avancée des enquêtes de satisfaction boutiques")
    print("")
    print("🔧 Fonctionnalités intégrées :")
    print("   ✓ Génération de rapports RCBT complets")
    print("   ✓ Détection intelligente d'incohérences contextuelles")
    print("   ✓ Système de validation centralisé avec traçabilité")
    print("   ✓ Interface web responsive et intuitive")
    print("   ✓ Rapports individuels par site/collaborateur")
    print("   ✓ Interface web intuitive")
    print("   ✓ Traçabilité complète des corrections")
    print("")
    print("🌐 ACCÈS À L'APPLICATION :")
    print("   ┌─ URL principale : http://localhost:5000")
    print("   └─ URL alternative : http://127.0.0.1:5000")
    print("")
    print("⚠️  IMPORTANT - NE PAS FERMER CETTE FENÊTRE")
    print("   Cette console doit rester ouverte pour que l'application")
    print("   web fonctionne. La fermer arrêtera le serveur RCBT.")
    print("")
    print("📞 Responsable technique :")
    print("   Contact   : Flageul Martin")
    print("   Service   : Pole QoS et Transformation SI Retail")
    print("   Direction : Exploitation et Système d'Information Retail")
    print("   Entité    : Réseau Clubs Bouygues Telecom")
    print("")
    print("💡 Guide d'utilisation rapide :")
    print("   1️⃣  Ouvrez votre navigateur à l'adresse ci-dessus")
    print("   2️⃣  Chargez vos 4 fichiers Excel requis")
    print("   3️⃣  Générez le rapport global d'analyse")
    print("   4️⃣  Validez les incohérences détectées")
    print("   5️⃣  Consultez/générez les rapports individuels")
    print("="*70)

if __name__ == '__main__':
    try:
        # Initialiser la base de données silencieusement
        ensure_database_schema()
        
        # S'assurer que les dossiers requis existent
        for folder in [UPLOAD_FOLDER, 'ressources', SESSION_CACHE_DIR, 'tmp']:
            if not os.path.exists(folder):
                os.makedirs(folder)
        
        # Afficher le banner professionnel
        print("\n" + "="*70)
        print("                       -- ISESIR --")
        print("           by Pole QoS et Transformation SI Retail")
        print("               Réseau Clubs Bouygues Telecom")
        print("="*70)
        print("📊 Analyse avancée des enquêtes de satisfaction boutiques")
        print("🔧 Fonctionnalités intégrées :")
        print("   ✓ Génération de rapports RCBT complets")
        print("   ✓ Détection intelligente d'incohérences contextuelles")
        print("   ✓ Système de validation centralisé avec traçabilité")
        print("   ✓ Interface web responsive et intuitive")
        print("   ✓ Rapports individuels par site/collaborateur")
        print("🌐 ACCÈS À L'APPLICATION :")
        print("   ┌─ URL principale : http://localhost:5000")
        print("   └─ URL alternative : http://127.0.0.1:5000")
        print("⚠️  IMPORTANT - NE PAS FERMER CETTE FENÊTRE")
        print("   Cette console doit rester ouverte pour que l'application")
        print("   web fonctionne. La fermer arrêtera le serveur RCBT.")
        print("📞 Responsable technique :")
        print("   Contact   : Flageul Martin")
        print("   Service   : Pole QoS et Transformation SI Retail") 
        print("   Direction : Exploitation et Système d'Information Retail")
        print("   Entité    : Réseau Clubs Bouygues Telecom")
        print("💡 Guide d'utilisation rapide :")
        print("   1️⃣  Ouvrez votre navigateur à l'adresse ci-dessus")
        print("   2️⃣  Chargez vos 4 fichiers Excel requis")
        print("   3️⃣  Générez le rapport global d'analyse")
        print("   4️⃣  Validez les incohérences détectées")
        print("   5️⃣  Consultez/générez les rapports individuels")
        print("="*70)
        print("🚀 Initialisation du serveur web en cours...")
        print("="*70)
        
        app.run(host='0.0.0.0', port=5000, debug=False)
        
    except KeyboardInterrupt:
        print("\n\n🛑 Arrêt de l'application demandé par l'utilisateur")
        print("✅ Application fermée proprement")
    except Exception as e:
        print(f"\n❌ ERREUR DE LANCEMENT : {e}")
        print("\n💡 Solutions possibles :")
        print("   • Vérifiez que le port 5000 n'est pas utilisé")
        print("   • Redémarrez l'application")
        print("   • Utilisez le fichier lancer_application.bat")
        print("   • Contactez le support technique")
        print("\n" + "="*50)
        input("Appuyez sur Entrée pour quitter...")
    finally:
        print("\n🔄 Redémarrage possible via lancer_application.bat")

