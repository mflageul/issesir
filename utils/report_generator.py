from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from datetime import datetime
from pathlib import Path
import os
import tempfile
import base64
from io import BytesIO
from PIL import Image as PILImage
from utils.data_processor import DataProcessor
from utils.visualizations import VisualizationGenerator

class ReportGenerator:
    def __init__(self):
        self.styles = getSampleStyleSheet()
        self.setup_custom_styles()
        
    def setup_custom_styles(self):
        """Setup custom paragraph styles"""
        # Title style
        self.title_style = ParagraphStyle(
            'CustomTitle',
            parent=self.styles['Heading1'],
            fontSize=16,
            spaceAfter=30,
            textColor=colors.HexColor('#0072C6'),
            alignment=TA_CENTER,
            fontName='Helvetica-Bold'
        )
        
        # Subtitle style
        self.subtitle_style = ParagraphStyle(
            'CustomSubtitle',
            parent=self.styles['Heading2'],
            fontSize=14,
            spaceAfter=20,
            textColor=colors.HexColor('#2c3e50'),
            fontName='Helvetica-Bold'
        )
        
        # Body style
        self.body_style = ParagraphStyle(
            'CustomBody',
            parent=self.styles['Normal'],
            fontSize=11,
            spaceAfter=12,
            textColor=colors.black
        )
        
        # Metric style
        self.metric_style = ParagraphStyle(
            'MetricStyle',
            parent=self.styles['Normal'],
            fontSize=12,
            spaceAfter=8,
            fontName='Helvetica-Bold'
        )

    def generate_enhanced_report(self, data):
        """Generate the complete enhanced PDF report"""
        print("=== GÉNÉRATION DU RAPPORT AMÉLIORÉ ===")
        
        # Create temporary file for the report
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        report_filename = f'rapport_rcbt_ameliore_{timestamp}.pdf'
        report_path = os.path.join(tempfile.gettempdir(), report_filename)
        
        # Initialize document
        doc = SimpleDocTemplate(
            report_path,
            pagesize=A4,
            rightMargin=72,
            leftMargin=72,
            topMargin=72,
            bottomMargin=18
        )
        
        # Process data
        processor = DataProcessor()
        page1_metrics = processor.calculate_page1_metrics(data)
        page2_metrics = processor.calculate_page2_enhanced_metrics(data)
        
        # Generate visualizations
        viz_generator = VisualizationGenerator()
        visualizations = viz_generator.create_all_visualizations(page2_metrics)
        
        # Build document content
        story = []
        
        # Cover page
        story.extend(self._create_cover_page())
        story.append(PageBreak())
        
        # Page 1: Synthesis
        story.extend(self._create_page1_synthesis(page1_metrics))
        story.append(PageBreak())
        
        # Page 2: Enhanced Comments Analysis
        story.extend(self._create_enhanced_page2(page2_metrics, visualizations))
        story.append(PageBreak())
        
        # Page 3: Unsatisfied without comments
        page3_data = processor.get_page3_data(data)
        story.extend(self._create_page3_unsatisfied(page3_data))
        story.append(PageBreak())
        
        # Additional analysis pages
        story.extend(self._create_collaborator_analysis_page(page2_metrics))
        story.append(PageBreak())
        
        story.extend(self._create_site_analysis_page(page2_metrics))
        
        # Build PDF
        doc.build(story)
        
        print(f"✓ Rapport généré: {report_path}")
        return report_path

    def _create_cover_page(self):
        """Create cover page"""
        story = []
        
        # Try to add logo if available
        logo_path = Path('ressources/logo.png')
        if logo_path.exists():
            try:
                logo = Image(str(logo_path), width=2*inch, height=2*inch)
                story.append(logo)
                story.append(Spacer(1, 0.5*inch))
            except:
                pass
        
        # Title
        title = Paragraph("RAPPORT D'ANALYSE RCBT", self.title_style)
        story.append(title)
        story.append(Spacer(1, 0.5*inch))
        
        # Subtitle
        subtitle = Paragraph("Analyse des enquêtes de satisfaction - Version Améliorée", self.subtitle_style)
        story.append(subtitle)
        story.append(Spacer(1, 1*inch))
        
        # Date
        date_str = datetime.now().strftime('%d/%m/%Y à %H:%M')
        date_para = Paragraph(f"Généré le {date_str}", self.body_style)
        story.append(date_para)
        story.append(Spacer(1, 0.5*inch))
        
        # Features list
        features_text = """
        <b>Nouvelles fonctionnalités de cette version :</b><br/>
        • Analyse détaillée par collaborateur avec métriques individuelles<br/>
        • Graphiques en secteurs (camemberts) avec pourcentages<br/>
        • Répartition des commentaires par site avec visualisations<br/>
        • Tableaux récapitulatifs avec indicateurs de performance<br/>
        • Détection automatique des incohérences<br/>
        • Métriques visuelles et comparaisons avec objectifs<br/>
        """
        features = Paragraph(features_text, self.body_style)
        story.append(features)
        
        return story

    def _create_page1_synthesis(self, metrics):
        """Create Page 1 - Synthesis"""
        story = []
        
        # Page title
        title = Paragraph("Page 1 - Synthèse Générale", self.title_style)
        story.append(title)
        story.append(Spacer(1, 0.3*inch))
        
        # Metrics table
        data = [
            ['Métrique', 'Valeur', 'Objectif', 'Statut'],
            ['Total tickets', str(metrics['total_tickets']), '-', '✓'],
            ['Tickets avec enquête', str(metrics['tickets_boutiques']), '-', '✓'],
            ['Tickets system', str(metrics['tickets_system']), '-', 'ℹ️'],
            ['Taux de clôture', f"{metrics['taux_closure']}%", '13%', '✅' if metrics['closure_ok'] else '⚠️'],
            ['Clients satisfaits', str(metrics['satisfaits']), '-', '✓'],
            ['Clients insatisfaits', str(metrics['insatisfaits']), '-', '⚠️'],
            ['Taux satisfaction Q1', f"{metrics['taux_sat']}%", '92%', '✅' if metrics['sat_ok'] else '⚠️']
        ]
        
        table = Table(data, colWidths=[3*inch, 1.5*inch, 1*inch, 1*inch])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#0072C6')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 12),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('FONTSIZE', (0, 1), (-1, -1), 10),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey])
        ]))
        
        story.append(table)
        story.append(Spacer(1, 0.3*inch))
        
        # Summary text
        closure_status = "conforme" if metrics['closure_ok'] else "non conforme"
        sat_status = "conforme" if metrics['sat_ok'] else "non conforme"
        
        summary_text = f"""
        <b>Résumé :</b><br/>
        Sur {metrics['total_tickets']} tickets au total, {metrics['tickets_boutiques']} ont fait l'objet d'une enquête de satisfaction,
        soit un taux de clôture de {metrics['taux_closure']}% ({closure_status} à l'objectif de 13%).<br/><br/>
        
        Parmi les enquêtes Q1, {metrics['satisfaits']} clients se déclarent satisfaits et {metrics['insatisfaits']} insatisfaits,
        soit un taux de satisfaction de {metrics['taux_sat']}% ({sat_status} à l'objectif de 92%).
        """
        
        summary = Paragraph(summary_text, self.body_style)
        story.append(summary)
        
        return story

    def _create_enhanced_page2(self, metrics, visualizations):
        """Create enhanced Page 2 - Comments Analysis"""
        story = []
        
        # Page title
        title = Paragraph("Page 2 - Analyse Améliorée des Commentaires", self.title_style)
        story.append(title)
        story.append(Spacer(1, 0.2*inch))
        
        # Global statistics
        global_stats = f"""
        <b>Statistiques globales :</b><br/>
        {metrics['comments_percentage']}% des enquêtes répondues ont un commentaire 
        ({metrics['total_with_comments']} sur {metrics['total_q1_responses']} enquêtes)<br/>
        {metrics['total_collaborators']} collaborateurs analysés sur {len(metrics['site_analysis'])} sites
        """
        story.append(Paragraph(global_stats, self.body_style))
        story.append(Spacer(1, 0.2*inch))
        
        # Satisfaction distribution pie chart
        if visualizations.get('satisfaction_pie'):
            story.append(Paragraph("Répartition par niveau de satisfaction", self.subtitle_style))
            img_data = base64.b64decode(visualizations['satisfaction_pie'])
            img = Image(BytesIO(img_data), width=5*inch, height=3.75*inch)
            story.append(img)
            story.append(Spacer(1, 0.2*inch))
        
        # Site distribution pie chart
        if visualizations.get('site_pie'):
            story.append(Paragraph("Répartition des commentaires par site", self.subtitle_style))
            img_data = base64.b64decode(visualizations['site_pie'])
            img = Image(BytesIO(img_data), width=6*inch, height=4.8*inch)
            story.append(img)
            story.append(Spacer(1, 0.2*inch))
        
        # Top collaborators table
        if metrics['collaborator_analysis']:
            story.append(Paragraph("Top 10 Collaborateurs par volume de commentaires", self.subtitle_style))
            
            # Create table data
            table_data = [['Collaborateur', 'Site', 'Total', 'Satisfaits', 'Insatisfaits', '% Satisfaction']]
            
            for collab in metrics['collaborator_analysis'][:10]:
                table_data.append([
                    collab['collaborator'][:30] + '...' if len(collab['collaborator']) > 30 else collab['collaborator'],
                    collab['site'][:15] + '...' if len(collab['site']) > 15 else collab['site'],
                    str(collab['total_comments']),
                    str(collab['satisfait_count']),
                    str(collab['insatisfait_count']),
                    f"{collab['satisfait_pct']}%"
                ])
            
            table = Table(table_data, colWidths=[2.5*inch, 1.5*inch, 0.7*inch, 0.8*inch, 0.8*inch, 0.9*inch])
            table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2ecc71')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 10),
                ('FONTSIZE', (0, 1), (-1, -1), 8),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey])
            ]))
            
            story.append(table)
        
        return story

    def _create_page3_unsatisfied(self, page3_data):
        """Create Page 3 - Unsatisfied without comments"""
        story = []
        
        # Page title
        title = Paragraph("Page 3 - Clients Insatisfaits sans Commentaire", self.title_style)
        story.append(title)
        story.append(Spacer(1, 0.3*inch))
        
        if page3_data.empty:
            no_data = Paragraph("✅ Aucun client insatisfait sans commentaire détecté.", self.body_style)
            story.append(no_data)
            return story
        
        # Summary
        summary_text = f"<b>{len(page3_data)} clients insatisfaits sans commentaire détectés :</b>"
        story.append(Paragraph(summary_text, self.metric_style))
        story.append(Spacer(1, 0.2*inch))
        
        # Create table with data
        columns = ['Dossier RCBT', 'Site', 'Collaborateur', 'Compte', 'Code Compte', 'Évaluation']
        
        # Prepare table data
        table_data = [columns]
        
        for _, row in page3_data.iterrows():
            table_data.append([
                str(row.get('Dossier Rcbt', 'N/A')),
                str(row.get('Site', 'N/A')),
                str(row.get('Créé par ticket', 'N/A'))[:20] + '...' if len(str(row.get('Créé par ticket', 'N/A'))) > 20 else str(row.get('Créé par ticket', 'N/A')),
                str(row.get('Compte', 'N/A'))[:25] + '...' if len(str(row.get('Compte', 'N/A'))) > 25 else str(row.get('Compte', 'N/A')),
                str(row.get('Code compte', 'N/A')),
                str(row.get('Valeur de chaîne', 'N/A'))
            ])
        
        # Create table
        table = Table(table_data, colWidths=[1.2*inch, 1*inch, 1.5*inch, 2*inch, 1*inch, 1.3*inch])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#e74c3c')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 9),
            ('FONTSIZE', (0, 1), (-1, -1), 8),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey]),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')
        ]))
        
        story.append(table)
        
        return story

    def _create_collaborator_analysis_page(self, metrics):
        """Create detailed collaborator analysis page"""
        story = []
        
        # Page title
        title = Paragraph("Analyse Détaillée par Collaborateur", self.title_style)
        story.append(title)
        story.append(Spacer(1, 0.2*inch))
        
        if not metrics['collaborator_analysis']:
            no_data = Paragraph("Aucune donnée de collaborateur disponible.", self.body_style)
            story.append(no_data)
            return story
        
        # Statistics
        total_collabs = len(metrics['collaborator_analysis'])
        high_performers = sum(1 for c in metrics['collaborator_analysis'] if c['satisfait_pct'] >= 92)
        avg_satisfaction = sum(c['satisfait_pct'] for c in metrics['collaborator_analysis']) / total_collabs
        
        stats_text = f"""
        <b>Statistiques par collaborateur :</b><br/>
        • Total collaborateurs analysés: {total_collabs}<br/>
        • Collaborateurs atteignant l'objectif (≥92%): {high_performers} ({round(high_performers/total_collabs*100, 1)}%)<br/>
        • Satisfaction moyenne: {round(avg_satisfaction, 1)}%<br/>
        """
        story.append(Paragraph(stats_text, self.body_style))
        story.append(Spacer(1, 0.2*inch))
        
        # Detailed table for all collaborators
        story.append(Paragraph("Tableau détaillé de tous les collaborateurs", self.subtitle_style))
        
        # Create comprehensive table
        table_data = [['Rang', 'Collaborateur', 'Site', 'Commentaires', '% Satisfaction', 'Performance']]
        
        for i, collab in enumerate(metrics['collaborator_analysis'], 1):
            performance = "🟢 Excellent" if collab['satisfait_pct'] >= 92 else \
                         "🟡 Moyen" if collab['satisfait_pct'] >= 75 else \
                         "🔴 À améliorer"
            
            table_data.append([
                str(i),
                collab['collaborator'][:25] + '...' if len(collab['collaborator']) > 25 else collab['collaborator'],
                collab['site'][:15] + '...' if len(collab['site']) > 15 else collab['site'],
                str(collab['total_comments']),
                f"{collab['satisfait_pct']}%",
                performance
            ])
        
        # Limit to first 20 for space
        if len(table_data) > 21:  # header + 20 rows
            table_data = table_data[:21]
            story.append(Paragraph(f"<i>Affichage des 20 premiers collaborateurs sur {total_collabs}</i>", self.body_style))
        
        table = Table(table_data, colWidths=[0.5*inch, 2.2*inch, 1.3*inch, 0.8*inch, 1*inch, 1.5*inch])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#3498db')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 9),
            ('FONTSIZE', (0, 1), (-1, -1), 8),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey]),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')
        ]))
        
        story.append(table)
        
        return story

    def _create_site_analysis_page(self, metrics):
        """Create detailed site analysis page"""
        story = []
        
        # Page title
        title = Paragraph("Analyse Détaillée par Site", self.title_style)
        story.append(title)
        story.append(Spacer(1, 0.2*inch))
        
        if not metrics['site_analysis']:
            no_data = Paragraph("Aucune donnée de site disponible.", self.body_style)
            story.append(no_data)
            return story
        
        # Site statistics
        total_sites = len(metrics['site_analysis'])
        total_comments_all = sum(s['total_comments'] for s in metrics['site_analysis'])
        avg_comments_per_site = total_comments_all / total_sites if total_sites > 0 else 0
        
        stats_text = f"""
        <b>Statistiques par site :</b><br/>
        • Total sites analysés: {total_sites}<br/>
        • Total commentaires: {total_comments_all}<br/>
        • Moyenne commentaires par site: {round(avg_comments_per_site, 1)}<br/>
        """
        story.append(Paragraph(stats_text, self.body_style))
        story.append(Spacer(1, 0.2*inch))
        
        # Site table
        story.append(Paragraph("Répartition détaillée par site", self.subtitle_style))
        
        table_data = [['Rang', 'Site', 'Commentaires', 'Collaborateurs', '% Satisfaction', 'Performance']]
        
        for i, site in enumerate(metrics['site_analysis'], 1):
            performance = "🟢 Excellent" if site['satisfait_pct'] >= 92 else \
                         "🟡 Moyen" if site['satisfait_pct'] >= 75 else \
                         "🔴 À améliorer"
            
            table_data.append([
                str(i),
                site['site'][:25] + '...' if len(site['site']) > 25 else site['site'],
                str(site['total_comments']),
                str(site['unique_collaborators']),
                f"{site['satisfait_pct']}%",
                performance
            ])
        
        table = Table(table_data, colWidths=[0.5*inch, 2.5*inch, 1*inch, 1*inch, 1*inch, 1.3*inch])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#9b59b6')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 9),
            ('FONTSIZE', (0, 1), (-1, -1), 8),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey]),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')
        ]))
        
        story.append(table)
        story.append(Spacer(1, 0.2*inch))
        
        # Inconsistencies if any
        if metrics.get('inconsistencies'):
            story.append(Paragraph("Incohérences Détectées", self.subtitle_style))
            story.append(Paragraph(f"<b>{len(metrics['inconsistencies'])} incohérence(s) détectée(s)</b> (évaluations négatives avec mots positifs):", self.body_style))
            
            for inconsistency in metrics['inconsistencies'][:5]:  # Show first 5
                incons_text = f"• Dossier {inconsistency['dossier']} - {inconsistency['collaborator']} - Rating: {inconsistency['rating']}"
                story.append(Paragraph(incons_text, self.body_style))
        
        return story

## file_path: ressources/.gitkeep
