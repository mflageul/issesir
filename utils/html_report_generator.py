import pandas as pd
from datetime import datetime
from pathlib import Path
import os
import tempfile
import base64
from utils.data_processor import DataProcessor
from utils.visualizations import VisualizationGenerator

class HTMLReportGenerator:
    def __init__(self):
        self.template = self._get_html_template()
        
    def generate_interactive_report(self, data):
        """Generate interactive HTML report with filters and tables"""
        print("=== GÉNÉRATION DU RAPPORT HTML INTERACTIF ===")
        
        # Create temporary file for the report
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        report_filename = f'rapport_rcbt_interactif_{timestamp}.html'
        report_path = os.path.join(tempfile.gettempdir(), report_filename)
        
        # Load and encode logo
        logo_base64 = self._get_logo_base64()
        
        # Process data
        processor = DataProcessor()
        page1_metrics = processor.calculate_page1_metrics(data)
        page2_metrics = processor.calculate_page2_enhanced_metrics(data)
        
        # Generate visualizations
        viz_generator = VisualizationGenerator()
        visualizations = viz_generator.create_all_visualizations(page2_metrics)
        
        # Get detailed data for interactive tables
        page2_data = processor.get_page2_data(data)
        page3_data = processor.get_page3_data(data)
        
        # Build HTML content
        html_content = self._build_html_report(
            page1_metrics, page2_metrics, visualizations, 
            page2_data, page3_data, logo_base64
        )
        
        # Write HTML file
        with open(report_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        print(f"✓ Rapport HTML généré: {report_path}")
        return report_path
    
    def _build_html_report(self, page1_metrics, page2_metrics, visualizations, page2_data, page3_data, logo_base64):
        """Build complete HTML report content"""
        
        # Generate interactive tables
        page1_table = self._create_page1_table(page1_metrics)
        collaborator_table = self._create_collaborator_table(page2_metrics['collaborator_analysis'])
        site_table = self._create_site_table(page2_metrics['site_analysis'])
        comments_table = self._create_comments_table(page2_data)
        unsatisfied_table = self._create_unsatisfied_table(page3_data)
        
        # Generate charts HTML
        charts_html = self._create_charts_section(visualizations)
        
        # Build metrics cards
        metrics_cards = self._create_metrics_cards(page1_metrics, page2_metrics)
        
        # Replace placeholders in template
        html_content = self.template.format(
            timestamp=datetime.now().strftime('%d/%m/%Y à %H:%M'),
            logo_base64=logo_base64,
            metrics_cards=metrics_cards,
            page1_table=page1_table,
            charts_html=charts_html,
            collaborator_table=collaborator_table,
            site_table=site_table,
            comments_table=comments_table,
            unsatisfied_table=unsatisfied_table,
            total_collaborators=page2_metrics['total_collaborators'],
            comments_percentage=page2_metrics['comments_percentage'],
            total_with_comments=page2_metrics['total_with_comments'],
            total_q1_responses=page2_metrics['total_q1_responses']
        )
        
        return html_content
    
    def _get_logo_base64(self):
        """Get logo as base64 encoded string"""
        try:
            logo_path = Path(__file__).parent.parent / 'static' / 'images' / 'logo.png'
            if logo_path.exists():
                with open(logo_path, 'rb') as f:
                    return base64.b64encode(f.read()).decode('utf-8')
        except Exception as e:
            print(f"⚠️ Impossible de charger le logo: {e}")
        
        # Return empty placeholder if logo not found
        return ""
    
    def _create_metrics_cards(self, page1_metrics, page2_metrics):
        """Create metrics cards HTML"""
        closure_class = "success" if page1_metrics['closure_ok'] else "warning"
        satisfaction_class = "success" if page1_metrics['sat_ok'] else "warning"
        
        return f"""
        <div class="row mb-4">
            <div class="col-md-3">
                <div class="card metric-card bg-{closure_class}">
                    <div class="card-body text-white text-center">
                        <h2 class="display-4">{page1_metrics['taux_closure']}%</h2>
                        <p class="mb-0">Taux de clôture</p>
                        <small>Objectif: 13%</small>
                    </div>
                </div>
            </div>
            <div class="col-md-3">
                <div class="card metric-card bg-{satisfaction_class}">
                    <div class="card-body text-white text-center">
                        <h2 class="display-4">{page1_metrics['taux_sat']}%</h2>
                        <p class="mb-0">Satisfaction Q1</p>
                        <small>Objectif: 92%</small>
                    </div>
                </div>
            </div>
            <div class="col-md-3">
                <div class="card metric-card bg-info">
                    <div class="card-body text-white text-center">
                        <h2 class="display-4">{page2_metrics['comments_percentage']}%</h2>
                        <p class="mb-0">Avec commentaires</p>
                        <small>{page2_metrics['total_with_comments']} sur {page2_metrics['total_q1_responses']}</small>
                    </div>
                </div>
            </div>
            <div class="col-md-3">
                <div class="card metric-card bg-primary">
                    <div class="card-body text-white text-center">
                        <h2 class="display-4">{page2_metrics['total_collaborators']}</h2>
                        <p class="mb-0">Collaborateurs</p>
                        <small>Analysés</small>
                    </div>
                </div>
            </div>
        </div>
        """
    
    def _create_page1_table(self, metrics):
        """Create Page 1 synthesis table"""
        return f"""
        <table class="table table-striped table-hover">
            <thead class="table-primary">
                <tr>
                    <th>Métrique</th>
                    <th>Valeur</th>
                    <th>Objectif</th>
                    <th>Statut</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td><strong>Total tickets</strong></td>
                    <td>{metrics['total_tickets']:,}</td>
                    <td>-</td>
                    <td><span class="badge bg-success">✓</span></td>
                </tr>
                <tr>
                    <td><strong>Tickets avec enquête</strong></td>
                    <td>{metrics['tickets_boutiques']:,}</td>
                    <td>-</td>
                    <td><span class="badge bg-success">✓</span></td>
                </tr>
                <tr>
                    <td><strong>Tickets system</strong></td>
                    <td>{metrics['tickets_system']:,}</td>
                    <td>-</td>
                    <td><span class="badge bg-info">ℹ️</span></td>
                </tr>
                <tr>
                    <td><strong>Taux de clôture</strong></td>
                    <td>{metrics['taux_closure']}%</td>
                    <td>13%</td>
                    <td><span class="badge bg-{'success' if metrics['closure_ok'] else 'warning'}">{'✅' if metrics['closure_ok'] else '⚠️'}</span></td>
                </tr>
                <tr>
                    <td><strong>Clients satisfaits</strong></td>
                    <td>{metrics['satisfaits']:,}</td>
                    <td>-</td>
                    <td><span class="badge bg-success">✓</span></td>
                </tr>
                <tr>
                    <td><strong>Clients insatisfaits</strong></td>
                    <td>{metrics['insatisfaits']:,}</td>
                    <td>-</td>
                    <td><span class="badge bg-warning">⚠️</span></td>
                </tr>
                <tr>
                    <td><strong>Taux satisfaction Q1</strong></td>
                    <td>{metrics['taux_sat']}%</td>
                    <td>92%</td>
                    <td><span class="badge bg-{'success' if metrics['sat_ok'] else 'warning'}">{'✅' if metrics['sat_ok'] else '⚠️'}</span></td>
                </tr>
            </tbody>
        </table>
        """
    
    def _create_collaborator_table(self, collaborator_analysis):
        """Create interactive collaborator analysis table"""
        if not collaborator_analysis:
            return "<p>Aucune donnée de collaborateur disponible.</p>"
        
        rows = ""
        for i, collab in enumerate(collaborator_analysis, 1):
            performance_class = "success" if collab['satisfait_pct'] >= 92 else "warning" if collab['satisfait_pct'] >= 70 else "danger"
            rows += f"""
            <tr>
                <td>{i}</td>
                <td>{collab['collaborator']}</td>
                <td>{collab['site']}</td>
                <td>{collab['total_comments']}</td>
                <td>{collab['satisfait_count']}</td>
                <td>{collab['insatisfait_count']}</td>
                <td><span class="badge bg-{performance_class}">{collab['satisfait_pct']}%</span></td>
            </tr>
            """
        
        return f"""
        <div class="table-responsive">
            <table class="table table-striped table-hover" id="collaboratorTable">
                <thead class="table-success">
                    <tr>
                        <th>Rang</th>
                        <th>Collaborateur <button class="btn btn-sm btn-outline-secondary" onclick="sortTable('collaboratorTable', 1)">↕️</button></th>
                        <th>Site <button class="btn btn-sm btn-outline-secondary" onclick="sortTable('collaboratorTable', 2)">↕️</button></th>
                        <th>Commentaires <button class="btn btn-sm btn-outline-secondary" onclick="sortTable('collaboratorTable', 3)">↕️</button></th>
                        <th>Satisfaits</th>
                        <th>Insatisfaits</th>
                        <th>% Satisfaction <button class="btn btn-sm btn-outline-secondary" onclick="sortTable('collaboratorTable', 6)">↕️</button></th>
                    </tr>
                </thead>
                <tbody>
                    {rows}
                </tbody>
            </table>
        </div>
        """
    
    def _create_site_table(self, site_analysis):
        """Create interactive site analysis table"""
        if not site_analysis:
            return "<p>Aucune donnée de site disponible.</p>"
        
        rows = ""
        for site in site_analysis:
            performance_class = "success" if site['satisfait_pct'] >= 92 else "warning" if site['satisfait_pct'] >= 70 else "danger"
            rows += f"""
            <tr>
                <td>{site['site']}</td>
                <td>{site['total_comments']}</td>
                <td>{site['unique_collaborators']}</td>
                <td>{site['satisfait_count']}</td>
                <td>{site['insatisfait_count']}</td>
                <td><span class="badge bg-{performance_class}">{site['satisfait_pct']}%</span></td>
            </tr>
            """
        
        return f"""
        <div class="table-responsive">
            <table class="table table-striped table-hover" id="siteTable">
                <thead class="table-info">
                    <tr>
                        <th>Site <button class="btn btn-sm btn-outline-secondary" onclick="sortTable('siteTable', 0)">↕️</button></th>
                        <th>Commentaires <button class="btn btn-sm btn-outline-secondary" onclick="sortTable('siteTable', 1)">↕️</button></th>
                        <th>Collaborateurs</th>
                        <th>Satisfaits</th>
                        <th>Insatisfaits</th>
                        <th>% Satisfaction <button class="btn btn-sm btn-outline-secondary" onclick="sortTable('siteTable', 5)">↕️</button></th>
                    </tr>
                </thead>
                <tbody>
                    {rows}
                </tbody>
            </table>
        </div>
        """
    
    def _create_comments_table(self, page2_data):
        """Create interactive comments table"""
        if page2_data.empty:
            return "<p>Aucun commentaire disponible.</p>"
        
        rows = ""
        for _, row in page2_data.head(50).iterrows():  # Limit to first 50 for performance
            satisfaction_class = "success" if row.get('Valeur de chaîne', '') in ['Très satisfaisant', 'Satisfaisant'] else "danger"
            rows += f"""
            <tr>
                <td>{row.get('Dossier Rcbt', 'N/A')}</td>
                <td>{row.get('Site', 'N/A')}</td>
                <td>{str(row.get('Créé par ticket', 'N/A'))[:30]}...</td>
                <td>{str(row.get('Compte', 'N/A'))[:40]}...</td>
                <td><span class="badge bg-{satisfaction_class}">{row.get('Valeur de chaîne', 'N/A')}</span></td>
                <td class="comment-cell">{str(row.get('Commentaire', 'N/A'))[:100]}...</td>
            </tr>
            """
        
        return f"""
        <div class="mb-3">
            <input type="text" id="commentsFilter" class="form-control" placeholder="Filtrer les commentaires..." onkeyup="filterTable('commentsTable', 'commentsFilter')">
        </div>
        <div class="table-responsive">
            <table class="table table-striped table-hover table-sm" id="commentsTable">
                <thead class="table-warning">
                    <tr>
                        <th>Dossier</th>
                        <th>Site</th>
                        <th>Collaborateur</th>
                        <th>Compte</th>
                        <th>Évaluation</th>
                        <th>Commentaire</th>
                    </tr>
                </thead>
                <tbody>
                    {rows}
                </tbody>
            </table>
        </div>
        <p class="text-muted"><small>Affichage des 50 premiers commentaires. Total: {len(page2_data)} commentaires.</small></p>
        """
    
    def _create_unsatisfied_table(self, page3_data):
        """Create unsatisfied customers table"""
        if page3_data.empty:
            return "<p class='text-success'>✅ Aucun client insatisfait sans commentaire détecté.</p>"
        
        rows = ""
        for _, row in page3_data.iterrows():
            rows += f"""
            <tr>
                <td>{row.get('Dossier Rcbt', 'N/A')}</td>
                <td>{row.get('Site', 'N/A')}</td>
                <td>{str(row.get('Créé par ticket', 'N/A'))[:30]}...</td>
                <td>{str(row.get('Compte', 'N/A'))[:40]}...</td>
                <td>{row.get('Code compte', 'N/A')}</td>
                <td><span class="badge bg-danger">{row.get('Valeur de chaîne', 'N/A')}</span></td>
            </tr>
            """
        
        return f"""
        <div class="table-responsive">
            <table class="table table-striped table-hover" id="unsatisfiedTable">
                <thead class="table-danger">
                    <tr>
                        <th>Dossier</th>
                        <th>Site</th>
                        <th>Collaborateur</th>
                        <th>Compte</th>
                        <th>Code Compte</th>
                        <th>Évaluation</th>
                    </tr>
                </thead>
                <tbody>
                    {rows}
                </tbody>
            </table>
        </div>
        <p class="text-muted mt-2"><small>📊 Total: {len(page3_data)} clients insatisfaits sans commentaire nécessitant un suivi</small></p>
        """
    
    def _create_charts_section(self, visualizations):
        """Create charts section with base64 embedded images"""
        charts_html = ""
        
        if visualizations.get('satisfaction_pie'):
            charts_html += f"""
            <div class="col-lg-6 mb-4">
                <div class="card h-100">
                    <div class="card-header bg-primary text-white">
                        <h5 class="mb-0">Répartition par satisfaction</h5>
                    </div>
                    <div class="card-body text-center">
                        <img src="data:image/png;base64,{visualizations['satisfaction_pie']}" 
                             class="img-fluid" alt="Graphique satisfaction">
                    </div>
                </div>
            </div>
            """
        
        if visualizations.get('site_pie'):
            charts_html += f"""
            <div class="col-lg-6 mb-4">
                <div class="card h-100">
                    <div class="card-header bg-info text-white">
                        <h5 class="mb-0">Répartition par site</h5>
                    </div>
                    <div class="card-body text-center">
                        <img src="data:image/png;base64,{visualizations['site_pie']}" 
                             class="img-fluid" alt="Graphique sites">
                    </div>
                </div>
            </div>
            """
        
        if visualizations.get('collaborator_bar'):
            charts_html += f"""
            <div class="col-12 mb-4">
                <div class="card">
                    <div class="card-header bg-success text-white">
                        <h5 class="mb-0">Volume par collaborateur</h5>
                    </div>
                    <div class="card-body text-center">
                        <img src="data:image/png;base64,{visualizations['collaborator_bar']}" 
                             class="img-fluid" alt="Graphique collaborateurs">
                    </div>
                </div>
            </div>
            """
        
        return charts_html
    
    def _get_html_template(self):
        """Get the HTML template with navigation between pages"""
        return """<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Rapport RCBT Interactif - 3 Pages</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>
        body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }}
        .metric-card {{ border-radius: 15px; box-shadow: 0 4px 8px rgba(0,0,0,0.1); }}
        .comment-cell {{ max-width: 300px; word-wrap: break-word; overflow: hidden; }}
        .table-responsive {{ max-height: 600px; }}
        .navbar {{ box-shadow: 0 2px 4px rgba(0,0,0,0.1); }}
        h1, h2 {{ color: #2c3e50; }}
        .card-header {{ font-weight: 600; }}
        .page-section {{ display: none; }}
        .page-section.active {{ display: block; }}
        .page-indicator {{ padding: 15px; margin: 10px 0; border-left: 4px solid #007bff; background: #f8f9fa; }}
        .nav-pills .nav-link.active {{ background-color: #007bff; }}
        @media print {{ 
            .no-print {{ display: none !important; }} 
            .page-section {{ display: block !important; }}
            .table {{ font-size: 11px; }}
            .card {{ margin-bottom: 30px; page-break-inside: avoid; }}
        }}
    </style>
</head>
<body>
    <nav class="navbar navbar-dark bg-primary no-print">
        <div class="container-fluid">
            <div class="d-flex align-items-center">
                <img src="data:image/png;base64,{logo_base64}" 
                     alt="Logo RCBT" class="me-3 logo-navbar">
                <div>
                    <span class="navbar-brand mb-0">📊 Rapport RCBT Interactif</span>
                    <small class="text-light d-block">{timestamp}</small>
                </div>
            </div>
            <div>
                <button class="btn btn-outline-light me-2" onclick="window.print()">🖨️ Imprimer</button>
                <button class="btn btn-outline-light" onclick="showAllPages()">👁️ Vue complète</button>
            </div>
        </div>
    </nav>

    <!-- Navigation Pills -->
    <div class="container-fluid mt-3 no-print">
        <ul class="nav nav-pills justify-content-center" id="page-nav">
            <li class="nav-item">
                <a class="nav-link active" href="#" onclick="showPage(1)">📋 Page 1: Synthèse</a>
            </li>
            <li class="nav-item">
                <a class="nav-link" href="#" onclick="showPage(2)">👥 Page 2: Collaborateurs</a>
            </li>
            <li class="nav-item">
                <a class="nav-link" href="#" onclick="showPage(3)">⚠️ Page 3: Actions</a>
            </li>
        </ul>
    </div>

    <div class="container-fluid mt-4">
        <!-- Metrics Cards - Always visible -->
        <div class="row no-print">
            {metrics_cards}
        </div>
        
        <!-- PAGE 1: SYNTHESE GENERALE -->
        <div class="page-section active" id="page1">
            <div class="page-indicator">
                <h1>📋 PAGE 1 - SYNTHÈSE GÉNÉRALE</h1>
                <p class="mb-0">Vue d'ensemble des performances et métriques principales</p>
            </div>
            
            <div class="row">
                <div class="col-md-8">
                    <div class="card">
                        <div class="card-header bg-primary text-white">
                            <h3 class="mb-0">📊 Tableau de Synthèse</h3>
                        </div>
                        <div class="card-body">
                            {page1_table}
                        </div>
                    </div>
                </div>
                <div class="col-md-4">
                    <div class="card h-100">
                        <div class="card-header bg-secondary text-white">
                            <h3 class="mb-0">📈 Graphiques</h3>
                        </div>
                        <div class="card-body">
                            {charts_html}
                        </div>
                    </div>
                </div>
            </div>
        </div>
        
        <!-- PAGE 2: ANALYSE COLLABORATIVE -->
        <div class="page-section" id="page2">
            <div class="page-indicator">
                <h1>👥 PAGE 2 - ANALYSE COLLABORATIVE</h1>
                <p class="mb-0">Performances par collaborateur et site | {total_collaborators} collaborateurs | {comments_percentage}% avec commentaires</p>
            </div>
            
            <div class="row">
                <div class="col-12 mb-4">
                    <div class="card">
                        <div class="card-header bg-success text-white">
                            <h3 class="mb-0">📋 Performance par Collaborateur</h3>
                            <small>Triable par colonne - Cliquez sur les en-têtes</small>
                        </div>
                        <div class="card-body">
                            {collaborator_table}
                        </div>
                    </div>
                </div>
                
                <div class="col-md-6">
                    <div class="card">
                        <div class="card-header bg-info text-white">
                            <h3 class="mb-0">🏢 Analyse par Site</h3>
                        </div>
                        <div class="card-body">
                            {site_table}
                        </div>
                    </div>
                </div>
                
                <div class="col-md-6">
                    <div class="card">
                        <div class="card-header bg-warning text-dark">
                            <h3 class="mb-0">💬 Commentaires Détaillés</h3>
                            <small>{total_with_comments} sur {total_q1_responses} enquêtes</small>
                        </div>
                        <div class="card-body">
                            {comments_table}
                        </div>
                    </div>
                </div>
            </div>
        </div>
        
        <!-- PAGE 3: PLAN D'ACTION -->
        <div class="page-section" id="page3">
            <div class="page-indicator">
                <h1>⚠️ PAGE 3 - PLAN D'ACTION</h1>
                <p class="mb-0">Clients insatisfaits nécessitant un suivi et actions correctives</p>
            </div>
            
            <div class="card">
                <div class="card-header bg-danger text-white">
                    <h3 class="mb-0">🚨 Clients Insatisfaits sans Commentaire - Actions Requises</h3>
                    <small>Ces clients nécessitent un contact immédiat pour comprendre leur insatisfaction</small>
                </div>
                <div class="card-body">
                    <div class="alert alert-warning">
                        <strong>Action requise :</strong> Contacter ces clients pour obtenir leur feedback détaillé 
                        et mettre en place des mesures correctives.
                    </div>
                    <input type="text" class="form-control mb-3" id="unsatisfiedFilter" 
                           placeholder="🔍 Filtrer par nom, ticket ou collaborateur..." 
                           onkeyup="filterTable('unsatisfiedTable', 'unsatisfiedFilter')">
                    {unsatisfied_table}
                </div>
            </div>
            
            <div class="row mt-4">
                <div class="col-md-6">
                    <div class="card">
                        <div class="card-header bg-warning text-dark">
                            <h4 class="mb-0">📈 Recommandations d'Amélioration</h4>
                        </div>
                        <div class="card-body">
                            <ul class="list-unstyled">
                                <li class="mb-2">🎯 <strong>Taux de clôture :</strong> Objectif 13% - Actions sur les tickets ouverts</li>
                                <li class="mb-2">😊 <strong>Satisfaction :</strong> Maintenir >92% - Suivi des insatisfaits</li>
                                <li class="mb-2">💬 <strong>Commentaires :</strong> Analyser les retours négatifs</li>
                                <li class="mb-2">👥 <strong>Formation :</strong> Accompagner les collaborateurs en difficulté</li>
                            </ul>
                        </div>
                    </div>
                </div>
                <div class="col-md-6">
                    <div class="card">
                        <div class="card-header bg-info text-white">
                            <h4 class="mb-0">📅 Suivi Recommandé</h4>
                        </div>
                        <div class="card-body">
                            <ul class="list-unstyled">
                                <li class="mb-2">✅ <strong>Immédiat :</strong> Contacter les insatisfaits sans commentaire</li>
                                <li class="mb-2">📞 <strong>Cette semaine :</strong> Suivi des tickets non clos</li>
                                <li class="mb-2">📋 <strong>Ce mois :</strong> Bilan avec les équipes</li>
                                <li class="mb-2">🔄 <strong>Récurrent :</strong> Monitoring mensuel des KPIs</li>
                            </ul>
                        </div>
                    </div>
                </div>
            </div>
        </div>
        
        <footer class="text-center text-muted mt-5 mb-4">
            <small>Rapport généré le {timestamp} - Application RCBT v2.0</small>
        </footer>
    </div>

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
    <script>
        // Page navigation
        function showPage(pageNum) {{
            // Hide all pages
            document.querySelectorAll('.page-section').forEach(page => {{
                page.classList.remove('active');
            }});
            
            // Remove active class from nav links
            document.querySelectorAll('.nav-link').forEach(link => {{
                link.classList.remove('active');
            }});
            
            // Show selected page
            document.getElementById('page' + pageNum).classList.add('active');
            
            // Add active class to current nav link
            event.target.classList.add('active');
        }}
        
        function showAllPages() {{
            document.querySelectorAll('.page-section').forEach(page => {{
                page.classList.add('active');
            }});
        }}
        
        // Table sorting
        function sortTable(tableId, columnIndex) {{
            const table = document.getElementById(tableId);
            if (!table) return;
            
            const tbody = table.querySelector('tbody');
            const rows = Array.from(tbody.querySelectorAll('tr'));
            
            const sorted = rows.sort((a, b) => {{
                const aText = a.cells[columnIndex].textContent.trim();
                const bText = b.cells[columnIndex].textContent.trim();
                
                const aNum = parseFloat(aText.replace(/[^0-9.-]/g, ''));
                const bNum = parseFloat(bText.replace(/[^0-9.-]/g, ''));
                
                if (!isNaN(aNum) && !isNaN(bNum)) {{
                    return bNum - aNum;
                }}
                
                return aText.localeCompare(bText);
            }});
            
            tbody.innerHTML = '';
            sorted.forEach(row => tbody.appendChild(row));
        }}
        
        // Table filtering
        function filterTable(tableId, filterId) {{
            const filter = document.getElementById(filterId).value.toLowerCase();
            const table = document.getElementById(tableId);
            if (!table) return;
            
            const rows = table.querySelectorAll('tbody tr');
            
            rows.forEach(row => {{
                const text = row.textContent.toLowerCase();
                row.style.display = text.includes(filter) ? '' : 'none';
            }});
        }}
        
        // Comment cell expansion
        document.addEventListener('DOMContentLoaded', function() {{
            document.querySelectorAll('.comment-cell').forEach(cell => {{
                cell.addEventListener('mouseenter', function() {{
                    this.style.whiteSpace = 'normal';
                    this.style.maxWidth = '600px';
                    this.style.overflow = 'visible';
                }});
                cell.addEventListener('mouseleave', function() {{
                    this.style.whiteSpace = 'nowrap';
                    this.style.maxWidth = '300px';
                    this.style.overflow = 'hidden';
                }});
            }});
        }});
    </script>
</body>
</html>"""