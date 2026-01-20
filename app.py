#!/usr/bin/env python3
"""
KCSE Enterprise System - Web Application
=========================================
Flask-based web interface for KCSE results extraction.
"""

from flask import Flask, render_template, request, jsonify, send_file, flash, redirect, url_for
from werkzeug.utils import secure_filename
import os
import logging
from pathlib import Path
from datetime import datetime
import tempfile
import zipfile

from kcse_modules.ocr_extractor import OCRExtractor
from kcse_modules.excel_processor import ExcelProcessor
from kcse_modules.config import Config
from kcse_modules.logger import setup_logger

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'dev-secret-key-change-in-production')
app.config['MAX_CONTENT_LENGTH'] = 500 * 1024 * 1024  # 500MB max file size
# Use project directory for templates
PROJECT_ROOT = Path(__file__).parent
app.config['TEMPLATE_FOLDER'] = PROJECT_ROOT / 'excel_templates'
# Use project directory for outputs (more accessible)
PROJECT_ROOT = Path(__file__).parent
app.config['UPLOAD_FOLDER'] = PROJECT_ROOT / 'uploads'
app.config['OUTPUT_FOLDER'] = PROJECT_ROOT / 'outputs'

# Create upload/output directories
app.config['UPLOAD_FOLDER'].mkdir(parents=True, exist_ok=True)
app.config['OUTPUT_FOLDER'].mkdir(parents=True, exist_ok=True)

# Initialize system components
config = Config()
logger = setup_logger('INFO', None)
ocr_extractor = OCRExtractor(logger, config)
excel_processor = ExcelProcessor(logger, config)

# Allowed file extensions
ALLOWED_IMAGE_EXTENSIONS = {'png', 'jpg', 'jpeg', 'PNG', 'JPG', 'JPEG'}
ALLOWED_EXCEL_EXTENSIONS = {'xlsx', 'xls'}


def allowed_file(filename, extensions):
    """Check if file extension is allowed."""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in extensions


@app.route('/')
def index():
    """Home page."""
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload_files():
    """Handle file upload and processing."""
    try:
        # Check if files were uploaded
        if 'screenshots' not in request.files or 'template' not in request.files:
            return jsonify({'error': 'Missing files. Please upload both screenshots and template.'}), 400
        
        screenshots = request.files.getlist('screenshots')
        template = request.files['template']
        
        if not screenshots or not template:
            return jsonify({'error': 'Please select files to upload.'}), 400
        
        # Validate template file
        if not allowed_file(template.filename, ALLOWED_EXCEL_EXTENSIONS):
            return jsonify({'error': 'Template must be an Excel file (.xlsx or .xls)'}), 400
        
        # Save template
        template_path = app.config['UPLOAD_FOLDER'] / secure_filename(template.filename)
        template.save(str(template_path))
        
        # Process screenshots
        extracted_data = {}
        processed_count = 0
        failed_count = 0
        errors = []
        
        for screenshot in screenshots:
            if screenshot.filename == '':
                continue
            
            if not allowed_file(screenshot.filename, ALLOWED_IMAGE_EXTENSIONS):
                failed_count += 1
                errors.append(f"{screenshot.filename}: Invalid file type")
                continue
            
            # Save screenshot
            screenshot_path = app.config['UPLOAD_FOLDER'] / secure_filename(screenshot.filename)
            screenshot.save(str(screenshot_path))
            
            # Extract data
            try:
                result = ocr_extractor.extract_from_image(screenshot_path)
                if result and result.get('index_number'):
                    index_num = result['index_number']
                    extracted_data[index_num] = result
                    processed_count += 1
                else:
                    failed_count += 1
                    errors.append(f"{screenshot.filename}: Could not extract index number")
            except Exception as e:
                failed_count += 1
                errors.append(f"{screenshot.filename}: {str(e)}")
                logger.error(f"Error processing {screenshot.filename}: {e}", exc_info=True)
        
        if not extracted_data:
            return jsonify({
                'error': 'No data extracted from screenshots. Please check image quality.',
                'errors': errors
            }), 400
        
        # Update Excel template
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_filename = f"kcse_results_{timestamp}.xlsx"
        output_path = app.config['OUTPUT_FOLDER'] / output_filename
        
        excel_processor.update_template(template_path, extracted_data, output_path)
        
        file_size = output_path.stat().st_size if output_path.exists() else 0
        
        return jsonify({
            'success': True,
            'message': f'Successfully processed {processed_count} screenshots',
            'processed': processed_count,
            'failed': failed_count,
            'errors': errors[:10],  # Limit errors shown
            'output_file': output_filename,
            'download_url': url_for('download_file', filename=output_filename),
            'file_path': str(output_path),
            'file_size': file_size,
            'output_directory': str(app.config['OUTPUT_FOLDER'])
        })
        
    except Exception as e:
        logger.error(f"Error in upload: {e}", exc_info=True)
        return jsonify({'error': f'Processing error: {str(e)}'}), 500


@app.route('/download/<filename>')
def download_file(filename):
    """Download processed Excel file."""
    try:
        file_path = app.config['OUTPUT_FOLDER'] / secure_filename(filename)
        if not file_path.exists():
            return jsonify({'error': 'File not found'}), 404
        
        return send_file(
            str(file_path),
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        logger.error(f"Error downloading file: {e}", exc_info=True)
        return jsonify({'error': f'Error downloading file: {str(e)}'}), 500


@app.route('/api/health')
def health_check():
    """Health check endpoint."""
    return jsonify({'status': 'healthy', 'service': 'KCSE Enterprise System'})


@app.route('/api/stats')
def stats():
    """Get system statistics."""
    upload_count = len(list(app.config['UPLOAD_FOLDER'].glob('*')))
    output_count = len(list(app.config['OUTPUT_FOLDER'].glob('*.xlsx')))
    
    return jsonify({
        'uploads': upload_count,
        'outputs': output_count,
        'version': '1.0.0'
    })


if __name__ == '__main__':
    # Development server
    app.run(debug=True, host='0.0.0.0', port=5000)
