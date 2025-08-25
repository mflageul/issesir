# RCBT Application

## Overview
The RCBT (Rapport Client Bout Tiers) Application is a Flask-based web application designed to analyze customer satisfaction surveys and generate comprehensive reports. It processes multiple Excel files (survey data, ticket information, reference data, account details) to create detailed satisfaction analysis reports with visualizations and statistics. The system calculates key performance indicators like closure rates, satisfaction percentages, and response rates, generating PDF reports with charts and tables, storing processing history, and providing a user-friendly web interface for file uploads and report generation. The project aims to provide automated individual reports for specific sites or collaborators, moving beyond global analysis.

## Recent Changes (August 13, 2025)
### UX/UI Improvements Delivered
- **Workflow State Management**: Individual report filters are now automatically hidden during global report generation and restored upon completion
- **CMD Interface Help**: Added helpful messaging system when web interface is accidentally closed, with clear reopening instructions  
- **Enhanced Validation Dialogs**: Improved user-friendly dialog boxes without technical references (localhost:5000 removed), featuring custom toasts and contextual messages with emojis

### Critical Bug Fixes Delivered
- **Inconsistency Detection**: Fixed false positive detection by implementing intelligent filtering that only reports inconsistencies when suggested rating differs from original rating
- **Data Analysis**: Confirmed "Nb réponses Q1" calculation is correct (523 responses with current data) - all proposed calculation methods yield identical results
- **History Display**: Fixed database history where "Nb réponses Q1" column incorrectly showed total tickets value (5975) instead of actual Q1 responses (523) - corrected column mapping in get_reports_history function
- **Database Error Fix**: Resolved "no such column: nb_reponses_q1" error by correcting SQL query column order and adding COALESCE for NULL value handling in reports_history function

### Advanced Inconsistency Detection Improvements (August 13, 2025 - Evening)
- **Enhanced Mitigated Comments Detection**: Successfully improved detection of subtle patterns like "J'ai été bien accueilli mais déçu par la qualité du traitement" with positive ratings
- **Word Boundary Fixes**: Eliminated false positives from "or" appearing within words like "collaborateur", "rapport" by implementing contextual detection (detects only " or ", "or,", etc.)
- **Contextual Analysis**: Added intelligent filtering for words like "long" - detects "trop long" (negative) but ignores "tout le long" (positive expression)
- **Reduced False Positives**: Optimized from 12 detections (many false positives) to 4 precise, valid inconsistencies
- **Extended Vocabulary**: Added 15+ negative expressions and 8+ contrast words for better nuanced comment analysis

## User Preferences
Preferred communication style: Simple, everyday language.

## System Architecture

### Frontend Architecture
The application uses a traditional server-side rendered architecture with Flask and Jinja2 templates for HTML rendering. It leverages Bootstrap 5.3.0 for responsive design, vanilla JavaScript for client-side functionality, and Feather icons for UI elements. HTML5 file inputs are used with client-side validation for Excel files.

### Backend Architecture
The backend follows a modular Flask application structure. `app.py` handles routing, file uploads, and request processing. Data processing, including loading, merging, and business logic calculations, is managed by `utils/data_processor.py`. Report generation uses `utils/report_generator.py` (ReportLab) and visualizations are created via `utils/visualizations.py` (matplotlib/seaborn). Configuration is environment-based with secure file upload restrictions.

### Data Storage Solutions
A local SQLite database (`rcbt_history.db`) stores report generation history and metadata. The file system is used for temporary uploads, storing generated reports and assets in a `resources` directory, and static files (CSS, JS, images).

### Authentication and Authorization
The application employs a basic security model with Flask's built-in session management. Security measures include secure filename handling, restriction of uploads to Excel files only, and a 50MB file size limit.

### Data Processing Pipeline
The system implements a multi-stage data processing workflow:
1.  **File Validation**: Ensures all required Excel files are present and valid.
2.  **Data Loading**: Reads Excel files using pandas with error handling.
3.  **Data Merging**: Joins survey data with ticket information.
4.  **Site Classification**: Maps users to sites and classifies channels.
5.  **Metrics Calculation**: Computes satisfaction rates, closure rates, and response percentages.
6.  **Report Generation**: Creates formatted PDF reports with embedded visualizations, including detection and flagging of inconsistencies in survey responses.

### Error Handling and Logging
The application includes comprehensive try-catch blocks for exception management, provides user feedback via Flask flash messages, and uses detailed logging for troubleshooting. Automatic cleanup of temporary files is performed after processing.

## External Dependencies

### Core Framework Dependencies
-   **Flask**: Web application framework.
-   **Werkzeug**: WSGI utilities for secure file handling.

### Data Processing Libraries
-   **pandas**: Excel file reading, data manipulation.
-   **numpy**: Numerical computations.
-   **sqlite3**: Database connectivity (Python standard library).

### Report Generation Libraries
-   **ReportLab**: PDF generation.
-   **matplotlib**: Chart and graph generation.
-   **seaborn**: Statistical data visualization.
-   **PIL (Pillow)**: Image processing for chart embedding.

### Frontend Libraries (CDN)
-   **Bootstrap 5.3.0**: CSS framework.
-   **Feather Icons**: Icon library.

### File and Utility Libraries
-   **pathlib**: Modern path handling.
-   **tempfile**: Temporary file and directory management.
-   **shutil**: High-level file operations.
-   **datetime**: Date and time handling.

### Development and Configuration
-   **os**: Environment variable access.
-   **base64**: Data encoding.
-   **traceback**: Error tracking.