class RCBTApp {
    constructor() {
        this.uploadedFiles = {};
        this.reportPath = null;
        this.availableData = null;
        this.workflowInProgress = false;  // AJOUT: État du workflow global
        this.init();
    }

    init() {
        // Initialize Feather icons
        feather.replace();
        
        // Setup event listeners
        this.setupEventListeners();
        
        // Bind delete function to window for HTML onclick access
        window.app = this;
        
        // Try to load logo
        this.loadLogo();
        
        // Initialize log
        this.log('Application RCBT initialisée', 'info');
        
        // Vérifier automatiquement s'il y a des données session disponibles
        this.checkForSessionData();
    }

    setupEventListeners() {
        // File upload form
        document.getElementById('uploadForm').addEventListener('submit', (e) => {
            e.preventDefault();
            this.startGlobalWorkflow();  // MODIFIÉ: Déclencher workflow global
            this.uploadFiles();
        });

        // Generate report button
        document.getElementById('generateBtn').addEventListener('click', () => {
            this.startGlobalWorkflow();  // MODIFIÉ: Déclencher workflow global
            this.generateReport();
        });

        // Download report button
        document.getElementById('downloadBtn').addEventListener('click', () => {
            this.downloadReport();
        });

        // Individual reports event listeners
        document.getElementById('reportType').addEventListener('change', (e) => {
            this.onReportTypeChange(e.target.value);
        });

        document.getElementById('reportTarget').addEventListener('change', (e) => {
            this.onReportTargetChange(e.target.value);
        });

        document.getElementById('generateIndividualBtn').addEventListener('click', () => {
            this.generateIndividualReport();
        });

        // File input change events
        ['enq_file', 'case_file', 'ref_file', 'acct_file'].forEach(fileId => {
            document.getElementById(fileId).addEventListener('change', (e) => {
                this.validateFile(e.target);
            });
        });
    }

    loadLogo() {
        const logoImg = document.getElementById('logo');
        logoImg.onerror = () => {
            logoImg.style.display = 'none';
        };
        logoImg.onload = () => {
            logoImg.style.display = 'block';
            this.log('Logo chargé avec succès', 'success');
        };
        // Try to load logo
        logoImg.src = '/static/images/logo.png';
    }

    validateFile(input) {
        const file = input.files[0];
        if (!file) return;

        const allowedTypes = ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'application/vnd.ms-excel'];
        const maxSize = 50 * 1024 * 1024; // 50MB

        if (!allowedTypes.includes(file.type)) {
            this.log(`Type de fichier non autorisé: ${file.name}`, 'error');
            input.value = '';
            return false;
        }

        if (file.size > maxSize) {
            this.log(`Fichier trop volumineux: ${file.name} (${this.formatFileSize(file.size)})`, 'error');
            input.value = '';
            return false;
        }

        this.log(`Fichier validé: ${file.name} (${this.formatFileSize(file.size)})`, 'success');
        return true;
    }

    formatFileSize(bytes) {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }

    async uploadFiles() {
        const formData = new FormData();
        const fileInputs = ['enq_file', 'case_file', 'ref_file', 'acct_file'];
        
        // Validate all files are selected
        for (const fileId of fileInputs) {
            const input = document.getElementById(fileId);
            if (!input.files[0]) {
                this.log(`Veuillez sélectionner le fichier: ${fileId}`, 'error');
                return;
            }
            formData.append(fileId, input.files[0]);
        }

        this.showProgress('Téléchargement des fichiers...', 20);
        this.log('Début du téléchargement des fichiers', 'info');

        try {
            const response = await fetch('/upload', {
                method: 'POST',
                body: formData
            });

            const result = await response.json();

            if (result.success) {
                this.uploadedFiles = result.files;
                this.log('Tous les fichiers téléchargés avec succès', 'success');
                this.updateProgress('Fichiers téléchargés', 50);
                
                // Enable generate button
                document.getElementById('generateBtn').disabled = false;
                document.getElementById('generateBtn').classList.add('pulse');
                
                setTimeout(() => {
                    this.hideProgress();
                }, 1000);
            } else {
                throw new Error(result.error);
            }
        } catch (error) {
            this.log(`Erreur lors du téléchargement: ${error.message}`, 'error');
            this.hideProgress();
        }
    }

    async generateReport() {
        if (!this.uploadedFiles || Object.keys(this.uploadedFiles).length !== 4) {
            this.log('Veuillez d\'abord télécharger tous les fichiers', 'error');
            return;
        }

        this.showProgress('Génération du rapport amélioré...', 10);
        this.log('Début de la génération du rapport', 'info');

        try {
            // Disable generate button
            document.getElementById('generateBtn').disabled = true;
            document.getElementById('generateBtn').classList.remove('pulse');

            const response = await fetch('/generate_report', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({
                    files: this.uploadedFiles
                })
            });

            const result = await response.json();

            if (result.success) {
                // Vérifier si des incohérences nécessitent une validation
                if (result.requires_validation) {
                    this.hideProgress();
                    this.log(`${result.inconsistencies_detected} incohérences détectées`, 'warning');
                    
                    // Afficher un message et rediriger vers la validation
                    const validationMessage = `
                        <div class="alert alert-warning" role="alert">
                            <h5><i data-feather="alert-triangle"></i> Incohérences détectées</h5>
                            <p>${result.message}</p>
                            <button class="btn btn-warning" onclick="window.location.href='${result.redirect_to}'">
                                <i data-feather="check-square"></i> Valider les incohérences
                            </button>
                        </div>
                    `;
                    
                    // Afficher le message dans la zone de log existante
                    const logOutput = document.getElementById('logOutput');
                    if (logOutput) {
                        logOutput.innerHTML += validationMessage;
                    } else {
                        // Fallback: afficher une alerte
                        alert(result.message);
                    }
                    feather.replace();
                    
                    // Re-enable generate button
                    document.getElementById('generateBtn').disabled = false;
                    return;
                }
                
                // Génération normale si pas d'incohérences
                this.reportPath = result.report_path;
                this.log('Rapport généré avec succès', 'success');
                
                // Update progress through stages
                this.updateProgress('Chargement des données...', 30);
                await this.delay(500);
                this.updateProgress('Analyse des métriques...', 50);
                await this.delay(500);
                this.updateProgress('Génération des graphiques...', 70);
                await this.delay(500);
                this.updateProgress('Création du rapport HTML...', 90);
                await this.delay(500);
                this.updateProgress('Rapport terminé', 100);

                // Show results
                this.showResults(result.metrics);
                
                // AJOUT: Terminer le workflow global et réafficher les rapports individuels
                this.endGlobalWorkflow();
                
                setTimeout(() => {
                    this.hideProgress();
                }, 1000);
            } else {
                throw new Error(result.error);
            }
        } catch (error) {
            this.log(`Erreur lors de la génération: ${error.message}`, 'error');
            this.hideProgress();
            document.getElementById('generateBtn').disabled = false;
            // AJOUT: Terminer le workflow même en cas d'erreur
            this.endGlobalWorkflow();
        }
    }

    showResults(metrics) {
        const resultsSection = document.getElementById('resultsSection');
        const metricsPreview = document.getElementById('metricsPreview');
        
        const page1 = metrics.page1;
        const page2 = metrics.page2;
        
        metricsPreview.innerHTML = `
            <div class="row g-3">
                <div class="col-md-3">
                    <div class="metric-card ${page1.closure_ok ? 'metric-positive' : 'metric-warning'}">
                        <div class="metric-value">${page1.taux_closure}%</div>
                        <div class="metric-label">Taux de clôture</div>
                    </div>
                </div>
                <div class="col-md-3">
                    <div class="metric-card ${page1.sat_ok ? 'metric-positive' : 'metric-warning'}">
                        <div class="metric-value">${page1.taux_sat}%</div>
                        <div class="metric-label">Satisfaction Q1</div>
                    </div>
                </div>
                <div class="col-md-3">
                    <div class="metric-card metric-info">
                        <div class="metric-value">${page2.comments_percentage}%</div>
                        <div class="metric-label">Enquêtes avec commentaires</div>
                    </div>
                </div>
                <div class="col-md-3">
                    <div class="metric-card metric-info">
                        <div class="metric-value">${page2.total_collaborators}</div>
                        <div class="metric-label">Collaborateurs analysés</div>
                    </div>
                </div>
            </div>

        `;
        
        resultsSection.style.display = 'block';
        resultsSection.classList.add('fade-in-up');
        
        // Show individual reports section
        this.showIndividualReportsSection();
        
        // Re-initialize feather icons for new content
        feather.replace();
        
        this.log('Résultats affichés avec succès', 'success');
    }

    async downloadReport() {
        if (!this.reportPath) {
            this.log('Aucun rapport à télécharger', 'error');
            return;
        }

        try {
            this.log('Téléchargement du rapport en cours...', 'info');
            window.open(`/download_report/${encodeURIComponent(this.reportPath)}`, '_blank');
            this.log('Téléchargement initié', 'success');
        } catch (error) {
            this.log(`Erreur lors du téléchargement: ${error.message}`, 'error');
        }
    }

    async showHistory() {
        try {
            const response = await fetch('/reports_history');
            const result = await response.json();

            if (result.success) {
                this.populateHistoryTable(result.history);
                const modal = new bootstrap.Modal(document.getElementById('historyModal'));
                modal.show();
            } else {
                this.log(`Erreur lors du chargement de l'historique: ${result.error}`, 'error');
            }
        } catch (error) {
            this.log(`Erreur lors du chargement de l'historique: ${error.message}`, 'error');
        }
    }

    showIndividualReportsSection() {
        const section = document.getElementById('individualReportsSection');
        section.style.display = 'block';
        section.classList.add('fade-in-up');
        
        // Load available data for dropdowns
        this.loadAvailableData();
    }

    async loadAvailableData() {
        if (!this.uploadedFiles || Object.keys(this.uploadedFiles).length === 0) {
            return;
        }

        try {
            const response = await fetch('/get_available_data', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({ files: this.uploadedFiles })
            });

            const result = await response.json();
            if (result.success) {
                this.availableData = result.data;
                this.log('Données disponibles chargées pour rapports individuels', 'success');
            } else {
                this.log(`Erreur chargement données: ${result.error}`, 'error');
            }
        } catch (error) {
            this.log(`Erreur lors du chargement des données: ${error.message}`, 'error');
        }
    }

    async checkForSessionData() {
        // Vérifier si une session active existe avec des fichiers
        try {
            const response = await fetch('/check_session_data');
            const result = await response.json();
            
            if (result.success && result.has_files) {
                // Restaurer les fichiers dans l'interface
                this.uploadedFiles = result.files;
                
                // Vérifier si on revient de validation d'incohérences
                const validationCompleted = sessionStorage.getItem('validation_completed');
                if (validationCompleted) {
                    this.log('Retour après validation - Session restaurée automatiquement', 'success');
                    sessionStorage.removeItem('validation_completed');
                } else if (!this.uploadedFiles) {
                    this.log('Session détectée - Données restaurées pour rapports individuels', 'info');
                }
                
                // Afficher la section rapports individuels
                this.showIndividualReportsSection();
            }
        } catch (error) {
            console.log('Pas de session disponible:', error.message);
        }
    }

    onReportTypeChange(type) {
        const targetSelect = document.getElementById('reportTarget');
        const generateBtn = document.getElementById('generateIndividualBtn');
        
        targetSelect.innerHTML = '<option value="">Chargement...</option>';
        targetSelect.disabled = true;
        generateBtn.disabled = true;

        if (!type || !this.availableData) {
            targetSelect.innerHTML = '<option value="">Choisir d\'abord le type</option>';
            return;
        }

        let options = '<option value="">Sélectionner...</option>';
        
        if (type === 'site' && this.availableData.sites) {
            this.availableData.sites.forEach(site => {
                options += `<option value="${site}">${site}</option>`;
            });
        } else if (type === 'collaborator' && this.availableData.collaborators) {
            this.availableData.collaborators.forEach(collab => {
                options += `<option value="${collab}">${collab}</option>`;
            });
        }

        targetSelect.innerHTML = options;
        targetSelect.disabled = false;
    }

    onReportTargetChange(target) {
        const generateBtn = document.getElementById('generateIndividualBtn');
        generateBtn.disabled = !target;
    }

    async generateIndividualReport() {
        const reportType = document.getElementById('reportType').value;
        const reportTarget = document.getElementById('reportTarget').value;

        if (!reportType || !reportTarget) {
            this.log('Veuillez sélectionner le type et la cible du rapport', 'error');
            return;
        }

        try {
            this.showIndividualProgress();
            this.log(`Génération du rapport individuel: ${reportType} - ${reportTarget}`, 'info');

            const response = await fetch('/generate_individual_report', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({
                    files: this.uploadedFiles,
                    report_type: reportType,
                    report_target: reportTarget
                })
            });

            const result = await response.json();

            if (result.success) {
                this.log(`Rapport individuel généré: ${result.report_path}`, 'success');
                this.hideIndividualProgress();
                
                // Open the individual report
                window.open(`/download_report/${encodeURIComponent(result.report_path)}`, '_blank');
                
                // Recharger les données disponibles pour garder les menus
                this.loadAvailableData();
            } else {
                throw new Error(result.error);
            }
        } catch (error) {
            this.log(`Erreur génération rapport individuel: ${error.message}`, 'error');
            this.hideIndividualProgress();
        }
    }

    showIndividualProgress() {
        const progress = document.getElementById('individualProgress');
        const progressBar = progress.querySelector('.progress-bar');
        const progressText = document.getElementById('individualProgressText');
        
        progress.style.display = 'block';
        progressBar.style.width = '0%';
        progressText.textContent = 'Génération en cours...';
        
        // Simulate progress
        let width = 0;
        const interval = setInterval(() => {
            width += Math.random() * 15;
            if (width >= 90) {
                clearInterval(interval);
                width = 90;
            }
            progressBar.style.width = width + '%';
        }, 200);
    }

    hideIndividualProgress() {
        const progress = document.getElementById('individualProgress');
        const progressBar = progress.querySelector('.progress-bar');
        
        progressBar.style.width = '100%';
        setTimeout(() => {
            progress.style.display = 'none';
            progressBar.style.width = '0%';
        }, 500);
    }

    populateHistoryTable(history) {
        const tbody = document.getElementById('historyTableBody');
        tbody.innerHTML = '';

        history.forEach(record => {
            const row = document.createElement('tr');
            const date = new Date(record.timestamp).toLocaleString('fr-FR');
            
            const reportTypeLabel = record.report_type === 'global' ? 'Global' : 'Individuel';
            const filterDisplay = record.filter_type && record.filter_value ? 
                `${record.filter_type}: ${record.filter_value}` : '-';
            
            row.innerHTML = `
                <td>${record.id}</td>
                <td>${date}</td>
                <td><span class="badge ${record.report_type === 'global' ? 'bg-primary' : 'bg-info'}">${reportTypeLabel}</span></td>
                <td>${filterDisplay}</td>
                <td>${record.filename}</td>
                <td>${record.total_tickets || 'N/A'}</td>
                <td>${record.nb_reponses_q1 || record.tickets_boutiques || 'N/A'}</td>
                <td>
                    <span class="status-badge ${record.taux_closure >= 13 ? 'status-success' : 'status-warning'}">
                        ${record.taux_closure || 0}%
                    </span>
                </td>
                <td>
                    <span class="status-badge ${record.taux_satisfaction >= 92 ? 'status-success' : 'status-warning'}">
                        ${record.taux_satisfaction || 0}%
                    </span>
                </td>
                <td>
                    <button class="btn btn-sm btn-primary me-1" onclick="app.downloadHistoryReport('${record.file_path}')" title="Télécharger">
                        <i data-feather="download"></i>
                    </button>
                    <button class="btn btn-sm btn-danger" onclick="app.deleteHistoryItem(${record.id})" title="Supprimer">
                        <i data-feather="trash-2"></i>
                    </button>
                </td>
            `;
            tbody.appendChild(row);
        });

        // Re-initialize feather icons
        feather.replace();
    }

    async deleteHistoryItem(reportId) {
        if (!confirm('Êtes-vous sûr de vouloir supprimer ce rapport ?')) {
            return;
        }
        
        try {
            const response = await fetch(`/delete_history/${reportId}`, {
                method: 'DELETE'
            });
            const data = await response.json();
            
            if (data.success) {
                this.log('Rapport supprimé avec succès', 'success');
                this.showHistory(); // Recharger la liste
            } else {
                this.log('Erreur lors de la suppression: ' + data.error, 'error');
            }
        } catch (error) {
            this.log('Erreur lors de la suppression: ' + error.message, 'error');
        }
    }

    downloadHistoryReport(filePath) {
        window.open(`/download_report/${encodeURIComponent(filePath)}`, '_blank');
    }

    showProgress(text, percentage) {
        const progressSection = document.getElementById('progressSection');
        const progressBar = progressSection.querySelector('.progress-bar');
        const progressText = document.getElementById('progressText');
        
        progressSection.style.display = 'block';
        progressBar.style.width = `${percentage}%`;
        progressText.textContent = text;
    }

    updateProgress(text, percentage) {
        const progressBar = document.querySelector('.progress-bar');
        const progressText = document.getElementById('progressText');
        
        progressBar.style.width = `${percentage}%`;
        progressText.textContent = text;
    }

    hideProgress() {
        document.getElementById('progressSection').style.display = 'none';
    }

    log(message, type = 'info') {
        const logOutput = document.getElementById('logOutput');
        const timestamp = new Date().toLocaleTimeString('fr-FR');
        
        const logEntry = document.createElement('div');
        logEntry.className = 'log-entry';
        
        const icon = this.getLogIcon(type);
        logEntry.innerHTML = `
            <span class="log-timestamp">[${timestamp}]</span>
            <span class="log-${type}"> ${icon} ${message}</span>
        `;
        
        logOutput.appendChild(logEntry);
        logOutput.scrollTop = logOutput.scrollHeight;
        
        console.log(`[${timestamp}] ${message}`);
    }

    getLogIcon(type) {
        const icons = {
            'success': '✅',
            'error': '❌',
            'warning': '⚠️',
            'info': 'ℹ️'
        };
        return icons[type] || 'ℹ️';
    }

    // AJOUT: Gestion de l'état du workflow global
    startGlobalWorkflow() {
        this.workflowInProgress = true;
        this.hideIndividualReports();
        this.log('🔄 Workflow global démarré - Rapports individuels masqués', 'info');
    }

    endGlobalWorkflow() {
        this.workflowInProgress = false;
        this.showIndividualReports();
        this.log('✅ Workflow global terminé - Rapports individuels disponibles', 'success');
    }

    hideIndividualReports() {
        const section = document.getElementById('individualReportsSection');
        if (section) {
            section.style.display = 'none';
        }
    }

    showIndividualReports() {
        const section = document.getElementById('individualReportsSection');
        if (section && this.availableData) {
            section.style.display = 'block';
        }
    }

    delay(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }
}

// Global functions
window.showHistory = function() {
    app.showHistory();
};

window.showUpdates = function() {
    const modal = new bootstrap.Modal(document.getElementById('updatesModal'));
    modal.show();
    setTimeout(() => feather.replace(), 100);
};

window.showContact = function() {
    const modal = new bootstrap.Modal(document.getElementById('contactModal'));
    modal.show();
    setTimeout(() => feather.replace(), 100);
};

window.app = null;

// Initialize app when DOM is loaded
document.addEventListener('DOMContentLoaded', () => {
    window.app = new RCBTApp();
});
