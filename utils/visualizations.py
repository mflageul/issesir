import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
from io import BytesIO
import base64
import matplotlib
matplotlib.use('Agg')  # Use non-interactive backend

class VisualizationGenerator:
    def __init__(self):
        # Set style
        plt.style.use('default')
        sns.set_palette("husl")
        
        # Configure matplotlib for better appearance
        plt.rcParams.update({
            'font.size': 10,
            'axes.titlesize': 12,
            'axes.labelsize': 10,
            'xtick.labelsize': 9,
            'ytick.labelsize': 9,
            'legend.fontsize': 9,
            'figure.titlesize': 14
        })

    def create_satisfaction_pie_chart(self, satisfaction_distribution):
        """Create pie chart for satisfaction distribution"""
        if not satisfaction_distribution:
            return None
        
        # Prepare data
        labels = []
        sizes = []
        colors = []
        
        # CORRECTION COULEURS FORÇÉE: Vert foncé pour "Très satisfaisant", vert clair pour "Satisfaisant" 
        color_map = {
            'Très satisfaisant': '#1e7e34',      # Vert TRÈS FONCÉ pour très satisfaisant
            'Satisfaisant': '#28a745',           # Vert MOYEN pour satisfaisant 
            'Peu satisfaisant': '#fd7e14',       # ORANGE pour "peu satisfaisant" 
            'Très peu satisfaisant': '#dc3545'   # ROUGE pour "très peu satisfaisant"
        }
        
        print(f"🎨 COULEURS CAMEMBERT APPLIQUÉES: {color_map}")  # Debug pour vérifier
        
        for satisfaction, data in satisfaction_distribution.items():
            labels.append(f"{satisfaction}\n({data['percentage']}%)")
            sizes.append(data['count'])
            color_used = color_map.get(satisfaction, '#95a5a6')
            colors.append(color_used)
            print(f"   - {satisfaction}: {color_used}")  # Debug chaque couleur appliquée
        
        # Create pie chart with optimized legend placement
        fig, ax = plt.subplots(figsize=(10, 6))
        
        # Create pie chart with labels outside
        wedges, texts, autotexts = ax.pie(
            sizes, 
            labels=None,  # Remove labels from pie to avoid overlap
            colors=colors,
            autopct='%1.1f%%',
            startangle=90,
            textprops={'fontsize': 10, 'weight': 'bold'},
            pctdistance=0.85
        )
        
        # Add legend with better positioning
        ax.legend(wedges, labels, 
                 title="Niveaux de satisfaction",
                 loc="center left", 
                 bbox_to_anchor=(1, 0, 0.5, 1),
                 fontsize=9)
        
        ax.set_title('Répartition des commentaires par niveau de satisfaction', 
                    fontsize=12, fontweight='bold', pad=20)
        
        plt.tight_layout()
        return self._fig_to_base64(fig)

    def create_site_distribution_pie_chart(self, site_analysis):
        """Create pie chart for site distribution"""
        if not site_analysis:
            return None
        
        # Prepare data (top 8 sites + others)
        sorted_sites = sorted(site_analysis, key=lambda x: x['total_comments'], reverse=True)
        
        labels = []
        sizes = []
        
        # Top 7 sites
        for site_data in sorted_sites[:7]:
            labels.append(f"{site_data['site']}\n({site_data['total_comments']} commentaires)")
            sizes.append(site_data['total_comments'])
        
        # Combine remaining sites as "Autres"
        if len(sorted_sites) > 7:
            others_count = sum(site['total_comments'] for site in sorted_sites[7:])
            if others_count > 0:
                labels.append(f"Autres\n({others_count} commentaires)")
                sizes.append(others_count)
        
        # Generate colors
        colors = plt.cm.Set3(np.linspace(0, 1, len(labels)))
        
        # Create pie chart with optimized legend placement
        fig, ax = plt.subplots(figsize=(12, 8))
        
        # Create pie chart with legend outside
        wedges, texts = ax.pie(
            sizes, 
            labels=None,  # Remove labels from pie
            colors=colors,
            autopct='%1.1f%%',
            startangle=90,
            textprops={'fontsize': 10, 'weight': 'bold'},
            pctdistance=0.85
        )
        
        # Add legend with better positioning  
        ax.legend(wedges, labels,
                 title="Sites et commentaires",
                 loc="center left",
                 bbox_to_anchor=(1, 0, 0.5, 1),
                 fontsize=9)
        
        ax.set_title('Répartition des commentaires par site', 
                    fontsize=12, fontweight='bold', pad=20)
        
        plt.tight_layout()
        return self._fig_to_base64(fig)
    
    def create_category_pie_chart(self, category_distribution):
        """Create pie chart for category distribution in individual reports"""
        if not category_distribution:
            return None
        
        # Prepare data
        labels = []
        sizes = []
        colors = []
        
        # Color map for categories
        color_map = {
            'Très petite entreprise': '#3498db',     # Bleu
            'Petite entreprise': '#2ecc71',          # Vert
            'Moyenne entreprise': '#f39c12',         # Orange
            'Grande entreprise': '#9b59b6',          # Violet
            'Entreprise de taille intermédiaire': '#e74c3c'  # Rouge
        }
        
        for category, data in category_distribution.items():
            labels.append(f"{category}\n({data['percentage']}%)")
            sizes.append(data['count'])
            colors.append(color_map.get(category, '#95a5a6'))
        
        # Create pie chart with optimized legend placement
        fig, ax = plt.subplots(figsize=(10, 6))
        
        # Create pie chart with legend outside
        wedges, texts = ax.pie(
            sizes, 
            labels=None,  # Remove labels from pie
            colors=colors,
            autopct='%1.1f%%',
            startangle=90,
            textprops={'fontsize': 10, 'weight': 'bold'},
            pctdistance=0.85
        )
        
        # Add legend with better positioning  
        ax.legend(wedges, labels,
                 title="Catégories d'entreprises",
                 loc="center left",
                 bbox_to_anchor=(1, 0, 0.5, 1),
                 fontsize=9)
        
        ax.set_title('Répartition par catégorie d\'entreprise', 
                    fontsize=12, fontweight='bold', pad=20)
        
        plt.tight_layout()
        return self._fig_to_base64(fig)

    def create_collaborator_bar_chart(self, collaborator_analysis):
        """Create bar chart for collaborator comment volume"""
        if not collaborator_analysis:
            return None
        
        # Top 15 collaborators
        top_collaborators = collaborator_analysis[:15]
        
        names = [collab['collaborator'] for collab in top_collaborators]
        comments = [collab['total_comments'] for collab in top_collaborators]
        
        # Create bar chart
        fig, ax = plt.subplots(figsize=(12, 8))
        bars = ax.bar(range(len(names)), comments, color='#3498db', alpha=0.7)
        
        # Customize chart
        ax.set_xlabel('Collaborateurs', fontweight='bold')
        ax.set_ylabel('Nombre de commentaires', fontweight='bold')
        ax.set_title('Volume de commentaires par collaborateur (Top 15)', 
                    fontsize=12, fontweight='bold', pad=20)
        
        # Rotate x-axis labels
        ax.set_xticks(range(len(names)))
        ax.set_xticklabels(names, rotation=45, ha='right')
        
        # Add value labels on bars
        for bar in bars:
            height = bar.get_height()
            ax.text(bar.get_x() + bar.get_width()/2., height + 0.1,
                   f'{int(height)}', ha='center', va='bottom', fontsize=9)
        
        plt.tight_layout()
        return self._fig_to_base64(fig)

    def create_site_satisfaction_chart(self, site_analysis):
        """Create chart showing satisfaction rate by site"""
        if not site_analysis:
            return None
        
        # Filter sites with at least 3 comments
        filtered_sites = [site for site in site_analysis if site['total_comments'] >= 3]
        if not filtered_sites:
            return None
        
        # Sort by satisfaction percentage
        sorted_sites = sorted(filtered_sites, key=lambda x: x['satisfait_pct'], reverse=True)
        
        # Take top 10
        top_sites = sorted_sites[:10]
        
        sites = [site['site'] for site in top_sites]
        satisfait_pct = [site['satisfait_pct'] for site in top_sites]
        total_comments = [site['total_comments'] for site in top_sites]
        
        # Create horizontal bar chart
        fig, ax = plt.subplots(figsize=(10, 8))
        bars = ax.barh(range(len(sites)), satisfait_pct, color='#27ae60', alpha=0.7)
        
        # Customize chart
        ax.set_xlabel('Pourcentage de satisfaction (%)', fontweight='bold')
        ax.set_ylabel('Sites', fontweight='bold')
        ax.set_title('Taux de satisfaction par site (≥3 commentaires)', 
                    fontsize=12, fontweight='bold', pad=20)
        
        # Set y-axis labels
        ax.set_yticks(range(len(sites)))
        ax.set_yticklabels([f"{site} ({total_comments[i]})" for i, site in enumerate(sites)])
        
        # Add percentage labels
        for i, bar in enumerate(bars):
            width = bar.get_width()
            ax.text(width + 1, bar.get_y() + bar.get_height()/2,
                   f'{satisfait_pct[i]}%', ha='left', va='center', fontsize=9)
        
        # Add reference line at 92%
        ax.axvline(x=92, color='red', linestyle='--', alpha=0.7, label='Objectif (92%)')
        ax.legend()
        
        plt.tight_layout()
        return self._fig_to_base64(fig)

    def create_collaborator_satisfaction_heatmap(self, collaborator_analysis):
        """Create heatmap showing collaborator performance"""
        if not collaborator_analysis or len(collaborator_analysis) < 5:
            return None
        
        # Filter collaborators with at least 3 comments
        filtered_collabs = [collab for collab in collaborator_analysis if collab['total_comments'] >= 3]
        if len(filtered_collabs) < 5:
            return None
        
        # Take top 15 by comment volume
        top_collabs = filtered_collabs[:15]
        
        # Prepare data for heatmap
        names = [collab['collaborator'][:20] + '...' if len(collab['collaborator']) > 20 
                else collab['collaborator'] for collab in top_collabs]
        satisfait_pct = [collab['satisfait_pct'] for collab in top_collabs]
        total_comments = [collab['total_comments'] for collab in top_collabs]
        
        # Create data matrix (1 row per collaborator, 2 columns: satisfaction%, total comments)
        data_matrix = np.array([satisfait_pct, total_comments]).T
        
        # Normalize total comments for better visualization
        normalized_comments = [(c - min(total_comments)) / (max(total_comments) - min(total_comments)) * 100 
                              for c in total_comments]
        data_matrix[:, 1] = normalized_comments
        
        # Create heatmap
        fig, ax = plt.subplots(figsize=(8, 10))
        
        im = ax.imshow(data_matrix, cmap='RdYlGn', aspect='auto', vmin=0, vmax=100)
        
        # Set ticks and labels
        ax.set_xticks([0, 1])
        ax.set_xticklabels(['Satisfaction (%)', 'Volume (normalisé)'])
        ax.set_yticks(range(len(names)))
        ax.set_yticklabels(names)
        
        # Add text annotations
        for i in range(len(names)):
            ax.text(0, i, f'{satisfait_pct[i]:.1f}%', ha='center', va='center', 
                   color='white' if satisfait_pct[i] < 50 else 'black', fontweight='bold')
            ax.text(1, i, f'{total_comments[i]}', ha='center', va='center', 
                   color='white' if normalized_comments[i] < 50 else 'black', fontweight='bold')
        
        ax.set_title('Performance des collaborateurs (Top 15)', 
                    fontsize=12, fontweight='bold', pad=20)
        
        # Add colorbar
        cbar = plt.colorbar(im, ax=ax, shrink=0.8)
        cbar.set_label('Score (%)', rotation=270, labelpad=20)
        
        plt.tight_layout()
        return self._fig_to_base64(fig)

    def _fig_to_base64(self, fig):
        """Convert matplotlib figure to base64 string"""
        buffer = BytesIO()
        fig.savefig(buffer, format='png', dpi=150, bbox_inches='tight', 
                   facecolor='white', edgecolor='none')
        buffer.seek(0)
        image_base64 = base64.b64encode(buffer.getvalue()).decode('utf-8')
        plt.close(fig)
        return image_base64

    def create_all_visualizations(self, page2_metrics):
        """Create all visualizations for Page 2"""
        visualizations = {}
        
        # Satisfaction pie chart
        print("🔍 GÉNÉRATION VISUALISATIONS - Vérification satisfaction_distribution")
        if page2_metrics.get('satisfaction_distribution'):
            print("✓ satisfaction_distribution trouvée, génération du camembert...")
            visualizations['satisfaction_pie'] = self.create_satisfaction_pie_chart(
                page2_metrics['satisfaction_distribution']
            )
        else:
            print("❌ satisfaction_distribution manquante dans page2_metrics")
        
        # Site distribution pie chart
        if page2_metrics.get('site_analysis'):
            visualizations['site_pie'] = self.create_site_distribution_pie_chart(
                page2_metrics['site_analysis']
            )
        
        # Collaborator bar chart
        if page2_metrics.get('collaborator_analysis'):
            visualizations['collaborator_bar'] = self.create_collaborator_bar_chart(
                page2_metrics['collaborator_analysis']
            )
        
        # Site satisfaction chart
        if page2_metrics.get('site_analysis'):
            visualizations['site_satisfaction'] = self.create_site_satisfaction_chart(
                page2_metrics['site_analysis']
            )
        
        # Collaborator heatmap
        if page2_metrics.get('collaborator_analysis'):
            visualizations['collaborator_heatmap'] = self.create_collaborator_satisfaction_heatmap(
                page2_metrics['collaborator_analysis']
            )
        
        return visualizations
