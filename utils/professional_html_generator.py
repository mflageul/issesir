import os
import base64
import tempfile
from datetime import datetime
from pathlib import Path
import pandas as pd

class ProfessionalReportGenerator:
    def __init__(self):
        self.template = self._get_professional_template()
        
    def generate_professional_report(self, data):
        """Generate professional HTML report with 7 pages like the example"""
        print("=== GÉNÉRATION DU RAPPORT PROFESSIONNEL RCBT ===")
        
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        report_filename = f'rapport_rcbt_{timestamp}.html'
        report_path = os.path.join(tempfile.gettempdir(), report_filename)
        
        # Get logo
        logo_base64 = self._get_logo_base64()
        
        # Import processor for calculations
        from .data_processor import DataProcessor
        processor = DataProcessor()
        
        # Calculate all metrics with proper Q1 filtering
        page1_metrics = processor.calculate_page1_metrics(data)
        
        # Debug: Print metrics to verify calculations
        print(f"DEBUG - Page 1 Metrics: {page1_metrics}")
        
        # Get detailed data for all 7 pages
        page1_data = self._get_page1_data(data, page1_metrics)
        page2_data = self._get_page2_data(data)  # Comments
        page3_data = self._get_page3_data(data)  # No comments unsatisfied
        page4_data = self._get_page4_data(data)  # Global analysis
        page5_data = self._get_page5_data(data)  # Responding shops
        page6_data = self._get_page6_data(data)  # Never responding shops
        page7_data = self._get_page7_data()      # History
        
        # Build complete HTML
        html_content = self.template.format(
            timestamp=datetime.now().strftime('%d/%m/%Y à %H:%M'),
            logo_base64=logo_base64,
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
        
        print(f"✓ Rapport professionnel généré: {report_path}")
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
    
    def _get_page1_data(self, data, metrics):
        """Page 1 - Synthèse mensuelle"""
        df_merged = data['merged']
        
        # KPI Cards
        taux_cloture = metrics.get('taux_cloture_boutiques', 0)
        taux_satisfaction = metrics.get('taux_satisfaction_q1', 0)
        
        taux_cloture_status = "success" if taux_cloture >= 13 else "warning"
        taux_satisfaction_status = "success" if taux_satisfaction >= 92 else "warning"
        
        taux_cloture_icon = "✅" if taux_cloture >= 13 else "⚠️"
        taux_satisfaction_icon = "✅" if taux_satisfaction >= 92 else "⚠️"
        
        return {
            'page1_taux_cloture': f"{taux_cloture:.1f}%",
            'page1_taux_cloture_status': taux_cloture_status,
            'page1_taux_cloture_icon': taux_cloture_icon,
            'page1_taux_satisfaction': f"{taux_satisfaction:.1f}%",
            'page1_taux_satisfaction_status': taux_satisfaction_status,
            'page1_taux_satisfaction_icon': taux_satisfaction_icon,
            'page1_tickets_boutiques': metrics.get('tickets_boutiques', 0),
            'page1_total_tickets': len(df_merged),
            'page1_satisfaits': metrics.get('satisfaits_q1', 0),
            'page1_total_q1': metrics.get('total_q1_responses', 0)
        }
    
    def _get_page2_data(self, data):
        """Page 2 - Commentaires"""
        df_merged = data['merged']
        
        # Tickets with comments
        with_comments = df_merged[df_merged['Commentaire'].notna() & (df_merged['Commentaire'].str.strip() != '')]
        
        # Build comments table
        comments_table = ""
        for _, row in with_comments.head(50).iterrows():  # Limit to first 50
            comments_table += f"""
            <tr>
                <td>{row.get('Dossier Rcbt', 'N/A')}</td>
                <td>{row.get('Site', 'N/A')}</td>
                <td>{row.get('Créé par ticket', 'N/A')}</td>
                <td>{row.get('Boutique_categorie', 'N/A')}</td>
                <td>{row.get('Valeur de chaîne', 'N/A')}</td>
                <td style="max-width: 300px; word-wrap: break-word;">{str(row.get('Commentaire', ''))[:200]}...</td>
            </tr>"""
        
        # Inconsistencies (negative comments with positive ratings)
        inconsistent = with_comments[
            (with_comments['Valeur de chaîne'].isin(['Très peu satisfaisant', 'Peu satisfaisant'])) &
            (with_comments['Commentaire'].str.len() > 20)
        ]
        
        inconsistent_table = ""
        for _, row in inconsistent.head(10).iterrows():
            q1_class = self._get_q1_class(row.get('Valeur de chaîne', ''))
            inconsistent_table += f"""
            <tr>
                <td>{row.get('Dossier Rcbt', 'N/A')}</td>
                <td>{row.get('Site', 'N/A')}</td>
                <td>{row.get('Créé par ticket', 'N/A')}</td>
                <td style="max-width: 200px; word-wrap: break-word;">{str(row.get('Commentaire', ''))[:150]}...</td>
                <td class="{q1_class}">{row.get('Valeur de chaîne', 'N/A')}</td>
            </tr>"""
        
        return {
            'page2_comments_count': len(with_comments),
            'page2_comments_table': comments_table,
            'page2_inconsistent_count': len(inconsistent),
            'page2_inconsistent_table': inconsistent_table
        }
    
    def _get_page3_data(self, data):
        """Page 3 - Insatisfaits sans commentaires"""
        df_merged = data['merged']
        
        # Unsatisfied without comments
        unsatisfied_no_comments = df_merged[
            (df_merged['Valeur de chaîne'].isin(['Très peu satisfaisant', 'Peu satisfaisant'])) &
            (df_merged['Commentaire'].isna() | (df_merged['Commentaire'].str.strip() == ''))
        ]
        
        table_html = ""
        for _, row in unsatisfied_no_comments.head(30).iterrows():
            q1_class = self._get_q1_class(row.get('Valeur de chaîne', ''))
            table_html += f"""
            <tr>
                <td>{row.get('Dossier Rcbt', 'N/A')}</td>
                <td>{row.get('Site', 'N/A')}</td>
                <td>{row.get('Créé par ticket', 'N/A')}</td>
                <td>{row.get('Boutique_categorie', 'N/A')}</td>
                <td class="{q1_class}">{row.get('Valeur de chaîne', 'N/A')}</td>
                <td>{str(row.get('Créé le', ''))[:10]}</td>
            </tr>"""
        
        return {
            'page3_count': len(unsatisfied_no_comments),
            'page3_table': table_html
        }
    
    def _get_page4_data(self, data):
        """Page 4 - Analyse globale"""
        df_merged = data['merged']
        
        # Q1 volumes
        q1_counts = df_merged['Valeur de chaîne'].value_counts()
        q1_table = ""
        for response, count in q1_counts.items():
            if pd.notna(response):
                q1_table += f"<tr><td>{response}</td><td>{count}</td></tr>"
        
        # Collaborator details
        collab_stats = df_merged.groupby(['Créé par ticket', 'Site']).agg({
            'Valeur de chaîne': 'count'
        }).reset_index()
        
        # Calculate satisfaction by collaborator
        collab_detail = []
        for (collaborateur, site), group_data in df_merged.groupby(['Créé par ticket', 'Site']):
            if len(group_data) >= 5:  # Only collaborators with 5+ responses
                satisf_counts = group_data['Valeur de chaîne'].value_counts()
                total = len(group_data)
                satisf = satisf_counts.get('Très satisfaisant', 0) + satisf_counts.get('Satisfaisant', 0)
                satisfaction_rate = (satisf / total * 100) if total > 0 else 0
                
                collab_detail.append({
                    'name': collaborateur,
                    'site': site,
                    'tres_satisf': satisf_counts.get('Très satisfaisant', 0),
                    'satisf': satisf_counts.get('Satisfaisant', 0),
                    'peu_satisf': satisf_counts.get('Peu satisfaisant', 0),
                    'tres_peu_satisf': satisf_counts.get('Très peu satisfaisant', 0),
                    'total': total,
                    'satisfaction_rate': satisfaction_rate
                })
        
        # Sort by total responses
        collab_detail.sort(key=lambda x: x['total'], reverse=True)
        
        collab_table = ""
        for collab in collab_detail[:20]:  # Top 20
            collab_table += f"""
            <tr>
                <td>{collab['name']}</td>
                <td>{collab['site']}</td>
                <td>{collab['tres_satisf']}</td>
                <td>{collab['satisf']}</td>
                <td>{collab['peu_satisf']}</td>
                <td>{collab['tres_peu_satisf']}</td>
                <td>{collab['total']}</td>
                <td>{collab['satisfaction_rate']:.1f}</td>
            </tr>"""
        
        return {
            'page4_q1_table': q1_table,
            'page4_collab_table': collab_table
        }
    
    def _get_page5_data(self, data):
        """Page 5 - Boutiques répondues"""
        df_accounts = data['accounts']
        df_merged = data['merged']
        
        # Get responding shops
        responding_accounts = set(df_merged['Code compte'].dropna())
        
        # Category analysis
        category_stats = []
        for category in df_accounts['Categorie'].unique():
            if pd.notna(category):
                category_accounts = df_accounts[df_accounts['Categorie'] == category]
                total_open = len(category_accounts)
                responded = len([acc for acc in category_accounts['Code compte'] if acc in responding_accounts])
                never_responded = total_open - responded
                
                response_rate = (responded / total_open * 100) if total_open > 0 else 0
                never_rate = (never_responded / total_open * 100) if total_open > 0 else 0
                
                category_stats.append({
                    'category': category,
                    'total_open': total_open,
                    'responded': responded,
                    'response_rate': response_rate,
                    'never_responded': never_responded,
                    'never_rate': never_rate,
                    'objective_response_ok': response_rate >= 30,
                    'objective_never_ok': never_rate <= 70
                })
        
        category_table = ""
        for stat in category_stats:
            category_table += f"""
            <tr>
                <td>{stat['category']}</td>
                <td>{stat['total_open']}</td>
                <td>{stat['responded']}</td>
                <td>{stat['response_rate']:.1f}</td>
                <td>{stat['never_responded']}</td>
                <td>{stat['never_rate']:.1f}</td>
                <td>{'True' if stat['objective_response_ok'] else 'False'}</td>
                <td>{'True' if stat['objective_never_ok'] else 'False'}</td>
            </tr>"""
        
        # Responding shops list
        responding_shops = df_accounts[df_accounts['Code compte'].isin(responding_accounts)]
        shops_table = ""
        for _, shop in responding_shops.head(50).iterrows():
            shops_table += f"""
            <tr>
                <td>{shop.get('Code compte', 'N/A')}</td>
                <td>{shop.get('Compte', 'N/A')}</td>
                <td>{shop.get('Categorie', 'N/A')}</td>
            </tr>"""
        
        return {
            'page5_category_table': category_table,
            'page5_responding_count': len(responding_shops),
            'page5_shops_table': shops_table
        }
    
    def _get_page6_data(self, data):
        """Page 6 - Boutiques jamais répondues"""
        df_accounts = data['accounts']
        df_merged = data['merged']
        
        responding_accounts = set(df_merged['Code compte'].dropna())
        never_responding = df_accounts[~df_accounts['Code compte'].isin(responding_accounts)]
        
        # Category analysis for never responding
        category_stats = []
        for category in never_responding['Categorie'].unique():
            if pd.notna(category):
                category_data = never_responding[never_responding['Categorie'] == category]
                total_category = len(df_accounts[df_accounts['Categorie'] == category])
                never_count = len(category_data)
                never_rate = (never_count / total_category * 100) if total_category > 0 else 0
                
                category_stats.append({
                    'category': category,
                    'total_open': total_category,
                    'never_responded': never_count,
                    'never_rate': never_rate
                })
        
        never_table = ""
        for stat in category_stats:
            never_table += f"""
            <tr>
                <td>{stat['category']}</td>
                <td>{stat['total_open']}</td>
                <td>{stat['never_responded']}</td>
                <td>{stat['never_rate']:.1f}</td>
            </tr>"""
        
        return {
            'page6_never_table': never_table
        }
    
    def _get_page7_data(self):
        """Page 7 - Historique"""
        return {
            'page7_content': "Évolution des indicateurs clés disponible dans la base de données locale"
        }
    
    def _get_q1_class(self, value):
        """Get CSS class for Q1 responses"""
        if value == 'Très satisfaisant':
            return 'q1-tres-satisfaisant'
        elif value == 'Satisfaisant':
            return 'q1-satisfaisant'
        elif value == 'Peu satisfaisant':
            return 'q1-peu-satisfaisant'
        elif value == 'Très peu satisfaisant':
            return 'q1-tres-peu-satisfaisant'
        return ''
    
    def _get_professional_template(self):
        """Professional HTML template matching the example"""
        return """<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Rapport RCBT - {timestamp}</title>
    <style>
        :root {{
            --bt-blue: #0072C6;
            --bt-orange: #F7941D;
            --bt-green: #00A551;
        }}
        body {{ 
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; 
            margin: 0; 
            padding: 20px; 
            background: #f5f5f5; 
        }}
        .container {{ 
            max-width: 1200px; 
            margin: 0 auto; 
            background: white; 
            padding: 30px; 
            border-radius: 12px; 
            box-shadow: 0 4px 20px rgba(0,0,0,0.1); 
        }}
        .header {{ 
            text-align: center; 
            margin-bottom: 40px; 
            padding-bottom: 20px;
            border-bottom: 3px solid var(--bt-blue);
        }}
        .header h1 {{ 
            color: var(--bt-blue); 
            margin: 10px 0; 
            font-size: 2.2rem;
            font-weight: 700;
        }}
        .header img {{
            max-width: 150px;
            height: auto;
            margin-bottom: 20px;
            border-radius: 8px;
        }}
        .tabs {{ 
            display: flex; 
            flex-wrap: wrap;
            border-bottom: 2px solid var(--bt-blue); 
            margin-bottom: 30px; 
            gap: 2px;
        }}
        .tab {{ 
            padding: 12px 18px; 
            cursor: pointer; 
            border: none; 
            background: #f8f9fa; 
            font-size: 14px; 
            font-weight: 600;
            border-radius: 6px 6px 0 0;
            transition: all 0.2s ease;
            color: #495057;
            white-space: nowrap;
        }}
        .tab:hover {{
            background: #e9ecef;
            color: var(--bt-blue);
        }}
        .tab.active {{ 
            background: var(--bt-blue); 
            color: white; 
        }}
        .tab-content {{ 
            display: none; 
            animation: fadeIn 0.3s ease-in;
        }}
        .tab-content.active {{ 
            display: block; 
        }}
        @keyframes fadeIn {{
            from {{ opacity: 0; transform: translateY(10px); }}
            to {{ opacity: 1; transform: translateY(0); }}
        }}
        .kpi-grid {{ 
            display: grid; 
            grid-template-columns: 1fr 1fr; 
            gap: 25px; 
            margin-bottom: 40px; 
        }}
        .kpi-card {{ 
            background: linear-gradient(135deg, #fff 0%, #f9f9f9 100%); 
            padding: 25px; 
            border-radius: 12px; 
            border-left: 6px solid var(--bt-blue);
            box-shadow: 0 2px 15px rgba(0,0,0,0.08);
            transition: transform 0.2s ease;
        }}
        .kpi-card:hover {{
            transform: translateY(-2px);
        }}
        .kpi-card h3 {{
            margin: 0 0 15px 0;
            color: #333;
            font-size: 1.1rem;
            font-weight: 600;
        }}
        .kpi-value {{ 
            font-size: 2.5rem; 
            font-weight: 700; 
            margin: 15px 0; 
            line-height: 1;
        }}
        .kpi-objective {{ 
            font-size: 0.9rem; 
            color: #666; 
            margin-bottom: 10px;
        }}
        .success {{ color: var(--bt-green); }}
        .warning {{ color: var(--bt-orange); }}
        .danger {{ color: #dc3545; }}
        .indicator {{ font-size: 1.2em; margin-left: 8px; }}
        
        table {{ 
            width: 100%; 
            border-collapse: collapse; 
            margin: 25px 0; 
            font-size: 0.9rem;
            border-radius: 8px;
            overflow: hidden;
            box-shadow: 0 2px 10px rgba(0,0,0,0.05);
        }}
        th, td {{ 
            padding: 12px 15px; 
            text-align: left; 
            border-bottom: 1px solid #ddd; 
        }}
        th {{ 
            background-color: var(--bt-blue); 
            color: white; 
            font-weight: 600;
            font-size: 0.85rem;
        }}
        tr:nth-child(even) {{ 
            background-color: #f9f9f9; 
        }}
        tr:hover {{
            background-color: rgba(0, 114, 198, 0.05);
        }}
        .q1-tres-satisfaisant {{ background-color: #d4edda !important; color: #155724; font-weight: 600; }}
        .q1-satisfaisant {{ background-color: #d1ecf1 !important; color: #0c5460; font-weight: 600; }}
        .q1-peu-satisfaisant {{ background-color: #fff3cd !important; color: #856404; font-weight: 600; }}
        .q1-tres-peu-satisfaisant {{ background-color: #f8d7da !important; color: #721c24; font-weight: 600; }}
        
        /* Advanced filtering styles */
        .search-container {{
            margin-bottom: 25px;
            padding: 20px;
            background: linear-gradient(135deg, #f8f9fa, #e9ecef);
            border-radius: 12px;
            border: 1px solid #dee2e6;
        }}
        .search-input {{
            width: 100%;
            max-width: 400px;
            padding: 12px 16px;
            border: 2px solid #e9ecef;
            border-radius: 8px;
            font-size: 14px;
            margin-bottom: 15px;
        }}
        .search-input:focus {{
            outline: none;
            border-color: var(--bt-blue);
            box-shadow: 0 0 0 3px rgba(0, 114, 198, 0.1);
        }}
        .column-filters {{
            display: flex;
            flex-wrap: wrap;
            gap: 10px;
            margin-top: 15px;
        }}
        .column-filter {{
            min-width: 120px;
            padding: 8px 12px;
            border: 1px solid #ced4da;
            border-radius: 6px;
            font-size: 13px;
            background: white;
        }}
        .column-filter:focus {{
            outline: none;
            border-color: var(--bt-blue);
        }}
        .filter-label {{
            font-weight: 600;
            color: #495057;
            margin-bottom: 10px;
            font-size: 14px;
        }}
        .table-wrapper {{
            position: relative;
            overflow-x: auto;
            margin-top: 20px;
        }}
        .no-results {{
            text-align: center;
            padding: 40px 20px;
            color: #6c757d;
            font-style: italic;
            background: #f8f9fa;
            border-radius: 8px;
            margin: 20px 0;
            display: none;
        }}
        .clear-filters {{
            background: var(--bt-orange);
            color: white;
            border: none;
            padding: 10px 20px;
            border-radius: 6px;
            cursor: pointer;
            font-weight: 600;
            margin-left: 15px;
            transition: all 0.2s ease;
        }}
        .clear-filters:hover {{
            background: #e8850f;
            transform: translateY(-1px);
        }}
        
        .section-title {{
            color: var(--bt-blue);
            font-size: 1.8rem;
            font-weight: 700;
            margin: 30px 0 20px 0;
            border-bottom: 2px solid var(--bt-blue);
            padding-bottom: 10px;
        }}
        
        .metric-highlight {{
            background: linear-gradient(135deg, var(--bt-blue), #0052a3);
            color: white;
            padding: 15px 20px;
            border-radius: 8px;
            margin: 15px 0;
            font-weight: 600;
        }}
        
        @media print {{
            body {{ padding: 0; }}
            .container {{ box-shadow: none; padding: 15px; }}
            .tab {{ display: none; }}
            .tab-content {{ display: block !important; }}
        }}
        
        @media (max-width: 968px) {{
            .tabs {{ 
                flex-direction: column; 
                align-items: stretch;
            }}
            .tab {{ 
                margin-bottom: 2px; 
                border-radius: 6px;
            }}
            .kpi-grid {{ 
                grid-template-columns: 1fr; 
            }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <img src="data:image/png;base64,{logo_base64}" alt="Logo RCBT">
            <h1>📊 Rapport RCBT - Analyse des enquêtes de satisfaction</h1>
            <p>Généré le {timestamp}</p>
        </div>
        
        <div class="tabs">
            <button class="tab active" onclick="showTab('page1')">Page 1 - Synthèse</button>
            <button class="tab" onclick="showTab('page2')">Page 2 - Commentaires</button>
            <button class="tab" onclick="showTab('page3')">Page 3 - Sans commentaires</button>
            <button class="tab" onclick="showTab('page4')">Page 4 - Analyse globale</button>
            <button class="tab" onclick="showTab('page5')">Page 5 - Boutiques répondues</button>
            <button class="tab" onclick="showTab('page6')">Page 6 - Boutiques non répondues</button>
            <button class="tab" onclick="showTab('page7')">Page 7 - Historique</button>
        </div>
        
        <div id="page1" class="tab-content active">
            <h2 class="section-title">Page 1 - Synthèse mensuelle</h2>
            <div class="kpi-grid">
                <div class="kpi-card">
                    <h3>Taux de clôture boutiques</h3>
                    <div class="kpi-value {page1_taux_cloture_status}">{page1_taux_cloture}</div>
                    <div class="kpi-objective">Objectif: ≥ 13% <span class="indicator">{page1_taux_cloture_icon}</span></div>
                    <p>Tickets boutiques: {page1_tickets_boutiques} / Total: {page1_total_tickets}</p>
                </div>
                <div class="kpi-card">
                    <h3>Taux de satisfaction Q1</h3>
                    <div class="kpi-value {page1_taux_satisfaction_status}">{page1_taux_satisfaction}</div>
                    <div class="kpi-objective">Objectif: ≥ 92% <span class="indicator">{page1_taux_satisfaction_icon}</span></div>
                    <p>Satisfaits: {page1_satisfaits} / Total Q1: {page1_total_q1}</p>
                </div>
            </div>
        </div>
        
        <div id="page2" class="tab-content">
            <h2 class="section-title">Page 2 - Retours avec commentaires</h2>
            <div class="metric-highlight">
                <strong>{page2_comments_count}</strong> tickets avec commentaires
            </div>
            
            <div class="search-container">
                <div class="filter-label">🔍 Recherche et filtrage</div>
                <input type="text" class="search-input" id="search-comments" placeholder="Rechercher dans tous les champs..." onkeyup="filterTable('comments-table', this.value)">
                <button class="clear-filters" onclick="clearAllFilters('page2')">Effacer tous les filtres</button>
                <div class="column-filters">
                    <input type="text" class="column-filter" placeholder="Filtrer Dossier..." onkeyup="filterColumn('comments-table', 0, this.value)">
                    <input type="text" class="column-filter" placeholder="Filtrer Site..." onkeyup="filterColumn('comments-table', 1, this.value)">
                    <input type="text" class="column-filter" placeholder="Filtrer Créé par..." onkeyup="filterColumn('comments-table', 2, this.value)">
                    <input type="text" class="column-filter" placeholder="Filtrer Catégorie..." onkeyup="filterColumn('comments-table', 3, this.value)">
                    <select class="column-filter" onchange="filterColumn('comments-table', 4, this.value)">
                        <option value="">Toutes réponses Q1</option>
                        <option value="Très satisfaisant">Très satisfaisant</option>
                        <option value="Satisfaisant">Satisfaisant</option>
                        <option value="Peu satisfaisant">Peu satisfaisant</option>
                        <option value="Très peu satisfaisant">Très peu satisfaisant</option>
                    </select>
                    <input type="text" class="column-filter" placeholder="Filtrer Commentaire..." onkeyup="filterColumn('comments-table', 5, this.value)">
                </div>
            </div>
            
            <div class="table-wrapper">
                <table id="comments-table">
                    <thead>
                        <tr>
                            <th>Dossier Rcbt</th>
                            <th>Site</th>
                            <th>Créé par</th>
                            <th>Catégorie</th>
                            <th>Réponse Q1</th>
                            <th>Commentaire</th>
                        </tr>
                    </thead>
                    <tbody>
                        {page2_comments_table}
                    </tbody>
                </table>
                <div class="no-results" id="no-results-comments">
                    Aucun résultat trouvé pour les critères de recherche actuels.
                </div>
            </div>
            
            <h3>⚠️ Incohérences commentaire ↔ Q1 détectées ({page2_inconsistent_count})</h3>
            
            <div class="search-container">
                <div class="filter-label">🔍 Recherche dans les incohérences</div>
                <input type="text" class="search-input" placeholder="Rechercher dans les incohérences..." onkeyup="filterTable('inconsistent-table', this.value)">
                <button class="clear-filters" onclick="clearAllFilters('inconsistent')">Effacer filtres</button>
            </div>
            
            <div class="table-wrapper">
                <table id="inconsistent-table">
                    <thead>
                        <tr>
                            <th>Dossier Rcbt</th>
                            <th>Site</th>
                            <th>Créé par</th>
                            <th>Commentaire</th>
                            <th>Réponse Q1</th>
                        </tr>
                    </thead>
                    <tbody>
                        {page2_inconsistent_table}
                    </tbody>
                </table>
                <div class="no-results" id="no-results-inconsistent">
                    Aucun résultat trouvé.
                </div>
            </div>
        </div>
        
        <div id="page3" class="tab-content">
            <h2 class="section-title">Page 3 - Retours insatisfaisants sans commentaires</h2>
            <div class="metric-highlight">
                <strong>{page3_count}</strong> tickets insatisfaits sans commentaires
            </div>
            
            <div class="search-container">
                <div class="filter-label">🔍 Recherche et filtrage</div>
                <input type="text" class="search-input" placeholder="Rechercher..." onkeyup="filterTable('unsatisfied-table', this.value)">
                <button class="clear-filters" onclick="clearAllFilters('page3')">Effacer filtres</button>
                <div class="column-filters">
                    <input type="text" class="column-filter" placeholder="Filtrer Dossier..." onkeyup="filterColumn('unsatisfied-table', 0, this.value)">
                    <input type="text" class="column-filter" placeholder="Filtrer Site..." onkeyup="filterColumn('unsatisfied-table', 1, this.value)">
                    <input type="text" class="column-filter" placeholder="Filtrer Créé par..." onkeyup="filterColumn('unsatisfied-table', 2, this.value)">
                    <input type="text" class="column-filter" placeholder="Filtrer Catégorie..." onkeyup="filterColumn('unsatisfied-table', 3, this.value)">
                    <select class="column-filter" onchange="filterColumn('unsatisfied-table', 4, this.value)">
                        <option value="">Toutes réponses</option>
                        <option value="Peu satisfaisant">Peu satisfaisant</option>
                        <option value="Très peu satisfaisant">Très peu satisfaisant</option>
                    </select>
                </div>
            </div>
            
            <div class="table-wrapper">
                <table id="unsatisfied-table">
                    <thead>
                        <tr>
                            <th>Dossier Rcbt</th>
                            <th>Site</th>
                            <th>Créé par</th>
                            <th>Catégorie</th>
                            <th>Réponse Q1</th>
                            <th>Date création</th>
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
            <h3>Volumes par catégorie Q1</h3>
            <table>
                <thead>
                    <tr>
                        <th>Réponse Q1</th>
                        <th>Nombre</th>
                    </tr>
                </thead>
                <tbody>
                    {page4_q1_table}
                </tbody>
            </table>
            
            <h3>Détail par collaborateur</h3>
            
            <div class="search-container">
                <div class="filter-label">🔍 Recherche dans les collaborateurs</div>
                <input type="text" class="search-input" placeholder="Rechercher un collaborateur..." onkeyup="filterTable('collab-table', this.value)">
                <button class="clear-filters" onclick="clearAllFilters('collab')">Effacer filtres</button>
                <div class="column-filters">
                    <input type="text" class="column-filter" placeholder="Nom collaborateur..." onkeyup="filterColumn('collab-table', 0, this.value)">
                    <input type="text" class="column-filter" placeholder="Site..." onkeyup="filterColumn('collab-table', 1, this.value)">
                    <input type="number" class="column-filter" placeholder="Min Total..." onkeyup="filterNumberColumn('collab-table', 6, this.value, 'min')">
                    <input type="number" class="column-filter" placeholder="Min Satisfaction %..." onkeyup="filterNumberColumn('collab-table', 7, this.value, 'min')">
                </div>
            </div>
            
            <div class="table-wrapper">
                <table id="collab-table">
                    <thead>
                        <tr>
                            <th>Collaborateur</th>
                            <th>Site</th>
                            <th>Très satisfaisant</th>
                            <th>Satisfaisant</th>
                            <th>Peu satisfaisant</th>
                            <th>Très peu satisfaisant</th>
                            <th>Total retours Q1</th>
                            <th>Taux satisfaction (%)</th>
                        </tr>
                    </thead>
                    <tbody>
                        {page4_collab_table}
                    </tbody>
                </table>
                <div class="no-results" id="no-results-collab">
                    Aucun résultat trouvé.
                </div>
            </div>
        </div>
        
        <div id="page5" class="tab-content">
            <h2 class="section-title">Page 5 - Boutiques ayant répondu</h2>
            <div class="metric-highlight">
                <strong>Objectif:</strong> ≥ 30% de boutiques ayant répondu
            </div>
            
            <h3>Synthèse par catégorie</h3>
            <div class="search-container">
                <div class="filter-label">🔍 Filtrer les catégories</div>
                <input type="text" class="search-input" placeholder="Rechercher une catégorie..." onkeyup="filterTable('category-table', this.value)">
                <button class="clear-filters" onclick="clearAllFilters('category')">Effacer filtres</button>
            </div>
            
            <div class="table-wrapper">
                <table id="category-table">
                    <thead>
                        <tr>
                            <th>Catégorie</th>
                            <th>Boutiques ouvertes</th>
                            <th>Boutiques répondues</th>
                            <th>% répondues</th>
                            <th>Jamais répondues</th>
                            <th>% jamais répondues</th>
                            <th>Objectif_reponse_ok</th>
                            <th>Objectif_jamais_ok</th>
                        </tr>
                    </thead>
                    <tbody>
                        {page5_category_table}
                    </tbody>
                </table>
            </div>
            
            <h3>Liste des boutiques ayant répondu ({page5_responding_count})</h3>
            <div class="search-container">
                <div class="filter-label">🔍 Rechercher dans les boutiques</div>
                <input type="text" class="search-input" placeholder="Rechercher par code ou nom..." onkeyup="filterTable('shops-table', this.value)">
                <button class="clear-filters" onclick="clearAllFilters('shops')">Effacer filtres</button>
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
                            <th>Code compte</th>
                            <th>Compte</th>
                            <th>Catégorie</th>
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
                <strong>Objectif:</strong> ≤ 70% de boutiques n'ayant jamais répondu
            </div>
            
            <div class="search-container">
                <div class="filter-label">🔍 Filtrer les boutiques jamais répondues</div>
                <input type="text" class="search-input" placeholder="Rechercher une catégorie..." onkeyup="filterTable('never-table', this.value)">
                <button class="clear-filters" onclick="clearAllFilters('never')">Effacer filtres</button>
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
                            <th>Catégorie</th>
                            <th>Boutiques ouvertes</th>
                            <th>Jamais répondues</th>
                            <th>% jamais répondues</th>
                        </tr>
                    </thead>
                    <tbody>
                        {page6_never_table}
                    </tbody>
                </table>
                <div class="no-results" id="no-results-never">
                    Aucun résultat trouvé.
                </div>
            </div>
        </div>
        
        <div id="page7" class="tab-content">
            <h2 class="section-title">Page 7 - Historique et tendances</h2>
            <div class="metric-highlight">
                Évolution des indicateurs clés sur plusieurs mois
            </div>
            <p>{page7_content}</p>
            <p>Historique disponible dans la base de données locale</p>
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
                    // Check if row is also hidden by other filters
                    const isHiddenByOtherFilters = row.style.display === 'none' && 
                        !cellText.includes(filterValue) && filterValue !== '';
                    
                    if (!isHiddenByOtherFilters) {{
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
</body>
</html>"""