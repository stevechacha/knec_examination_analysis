document.addEventListener('DOMContentLoaded', function() {
    const form = document.getElementById('uploadForm');
    const submitBtn = document.getElementById('submitBtn');
    const progressSection = document.getElementById('progressSection');
    const resultSection = document.getElementById('resultSection');
    const errorSection = document.getElementById('errorSection');
    const progressFill = document.getElementById('progressFill');
    const progressPercentage = document.getElementById('progressPercentage');
    const progressText = document.getElementById('progressText');
    const resultContent = document.getElementById('resultContent');
    const errorContent = document.getElementById('errorContent');
    const downloadBtn = document.getElementById('downloadBtn');

    // File input handlers
    setupFileInput('template', 'templateSelected', 'templateFileName');
    setupFileInput('screenshots', 'screenshotsSelected', null, true);

    // Drag and drop functionality
    setupDragAndDrop();

    form.addEventListener('submit', async function(e) {
        e.preventDefault();
        
        // Hide previous results
        resultSection.style.display = 'none';
        errorSection.style.display = 'none';
        
        // Show progress
        progressSection.style.display = 'block';
        updateProgress(0, 'Initializing...');
        
        // Disable submit button
        const btnContent = submitBtn.querySelector('.btn-content');
        const btnLoader = submitBtn.querySelector('.btn-loader');
        btnContent.style.display = 'none';
        btnLoader.style.display = 'flex';
        submitBtn.disabled = true;
        
        // Create FormData
        const formData = new FormData();
        
        const template = document.getElementById('template').files[0];
        const screenshots = document.getElementById('screenshots').files;
        
        if (!template || screenshots.length === 0) {
            showError('Please select both template and screenshot files.');
            resetForm();
            return;
        }
        
        formData.append('template', template);
        for (let i = 0; i < screenshots.length; i++) {
            formData.append('screenshots', screenshots[i]);
        }
        
        // Simulate progress updates
        updateProgress(20, 'Uploading files...');
        
        setTimeout(() => updateProgress(40, 'Processing screenshots with OCR...'), 500);
        
        try {
            const response = await fetch('/upload', {
                method: 'POST',
                body: formData
            });
            
            updateProgress(70, 'Extracting data...');
            
            const data = await response.json();
            
            if (response.ok && data.success) {
                updateProgress(100, 'Complete!');
                
                // Show success
                setTimeout(() => {
                    progressSection.style.display = 'none';
                    showSuccess(data);
                }, 800);
            } else {
                showError(data.error || 'Processing failed', data.errors);
                resetForm();
            }
        } catch (error) {
            showError('Network error: ' + error.message);
            resetForm();
        }
    });
    
    function setupFileInput(inputId, selectedId, fileNameId, multiple = false) {
        const input = document.getElementById(inputId);
        const selectedDiv = document.getElementById(selectedId);
        const uploadArea = input.closest('.file-upload-wrapper').querySelector('.file-upload-area');
        
        input.addEventListener('change', function() {
            if (multiple) {
                handleMultipleFiles(input, selectedDiv, uploadArea);
            } else {
                handleSingleFile(input, selectedDiv, fileNameId, uploadArea);
            }
        });
    }

    function handleSingleFile(input, selectedDiv, fileNameId, uploadArea) {
        if (input.files.length > 0) {
            const file = input.files[0];
            document.getElementById(fileNameId).textContent = file.name;
            uploadArea.querySelector('.file-upload-content').style.display = 'none';
            selectedDiv.style.display = 'flex';
        }
    }

    function handleMultipleFiles(input, selectedDiv, uploadArea) {
        if (input.files.length > 0) {
            const filesList = document.getElementById('screenshotsList');
            filesList.innerHTML = '';
            
            Array.from(input.files).forEach((file, index) => {
                const fileItem = document.createElement('div');
                fileItem.className = 'file-item';
                fileItem.innerHTML = `
                    <i class="fas fa-image"></i>
                    <span>${file.name}</span>
                    <span style="color: var(--text-light); font-size: 0.85em;">(${(file.size / 1024).toFixed(1)} KB)</span>
                `;
                filesList.appendChild(fileItem);
            });
            
            uploadArea.querySelector('.file-upload-content').style.display = 'none';
            selectedDiv.style.display = 'block';
        }
    }

    function setupDragAndDrop() {
        const uploadAreas = document.querySelectorAll('.file-upload-area');
        
        uploadAreas.forEach(area => {
            ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
                area.addEventListener(eventName, preventDefaults, false);
            });

            ['dragenter', 'dragover'].forEach(eventName => {
                area.addEventListener(eventName, () => area.classList.add('drag-over'), false);
            });

            ['dragleave', 'drop'].forEach(eventName => {
                area.addEventListener(eventName, () => area.classList.remove('drag-over'), false);
            });

            area.addEventListener('drop', handleDrop, false);
        });
    }

    function preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }

    function handleDrop(e) {
        const dt = e.dataTransfer;
        const files = dt.files;
        const target = e.currentTarget.getAttribute('data-target');
        const input = document.getElementById(target);
        
        if (target === 'screenshots') {
            input.files = files;
            handleMultipleFiles(input, document.getElementById('screenshotsSelected'), e.currentTarget);
        } else {
            input.files = files;
            handleSingleFile(input, document.getElementById('templateSelected'), 'templateFileName', e.currentTarget);
        }
    }

    function updateProgress(percentage, text) {
        progressFill.style.width = percentage + '%';
        progressPercentage.textContent = percentage + '%';
        progressText.textContent = text;
    }
    
    function showSuccess(data) {
        const fileSize = data.file_size ? `(${(data.file_size / 1024).toFixed(1)} KB)` : '';
        
        let html = `
            <div style="background: linear-gradient(135deg, rgba(16, 185, 129, 0.1) 0%, rgba(16, 185, 129, 0.05) 100%); padding: 20px; border-radius: 12px; margin-bottom: 20px;">
                <p style="font-size: 1.3em; font-weight: 700; margin-bottom: 10px; color: #065f46;">
                    <i class="fas fa-check-circle"></i> Processing Complete!
                </p>
                <p style="font-size: 1.1em; font-weight: 600; margin-bottom: 15px;">
                    Successfully processed <span style="color: #059669; font-size: 1.2em;">${data.processed}</span> screenshot(s)
                </p>
                ${data.output_file ? `
                    <div style="background: white; padding: 15px; border-radius: 8px; margin-top: 15px; border: 2px solid #10b981;">
                        <p style="margin: 0; font-weight: 600; color: #065f46;">
                            <i class="fas fa-file-excel"></i> Output File: <span style="color: #059669;">${data.output_file}</span> ${fileSize}
                        </p>
                        <p style="margin: 8px 0 0 0; font-size: 0.9em; color: #6b7280;">
                            Click the download button below to save the Excel file with filled grades
                        </p>
                    </div>
                ` : ''}
            </div>
        `;
        
        if (data.failed > 0) {
            html += `<div style="background: rgba(245, 158, 11, 0.1); padding: 15px; border-radius: 8px; margin: 15px 0; border-left: 4px solid #f59e0b;">
                <p class="warning" style="margin: 0; color: #92400e;">
                    <i class="fas fa-exclamation-triangle"></i> ${data.failed} screenshot(s) failed to process
                </p>
            </div>`;
        }
        
        if (data.errors && data.errors.length > 0) {
            html += '<div style="margin-top: 20px;"><strong style="color: #6b7280;">Processing Details:</strong><ul style="margin-top: 10px;">';
            data.errors.forEach(error => {
                html += `<li style="margin: 5px 0; color: #6b7280;">${error}</li>`;
            });
            html += '</ul></div>';
        }
        
        resultContent.innerHTML = html;
        
        if (data.download_url) {
            downloadBtn.href = data.download_url;
            downloadBtn.style.display = 'inline-flex';
            document.getElementById('downloadHint').style.display = 'block';
            
            // Auto-trigger download after a short delay
            setTimeout(() => {
                // Create a temporary link and trigger download
                const link = document.createElement('a');
                link.href = data.download_url;
                link.download = data.output_file || 'kcse_results.xlsx';
                link.style.display = 'none';
                document.body.appendChild(link);
                link.click();
                
                // Clean up after a delay
                setTimeout(() => {
                    document.body.removeChild(link);
                }, 100);
            }, 800);
        }
        
        resultSection.style.display = 'block';
        resultSection.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
        resetForm();
    }
    
    function showError(message, errors = null) {
        let html = `<p style="font-size: 1.1em; font-weight: 600; margin-bottom: 15px;">${message}</p>`;
        
        if (errors && errors.length > 0) {
            html += '<div style="margin-top: 15px;"><strong>Details:</strong><ul>';
            errors.forEach(error => {
                html += `<li>${error}</li>`;
            });
            html += '</ul></div>';
        }
        
        errorContent.innerHTML = html;
        errorSection.style.display = 'block';
        errorSection.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    }
    
    function resetForm() {
        const btnContent = submitBtn.querySelector('.btn-content');
        const btnLoader = submitBtn.querySelector('.btn-loader');
        btnContent.style.display = 'flex';
        btnLoader.style.display = 'none';
        submitBtn.disabled = false;
        updateProgress(0, '');
    }

    // Global function for clearing files
    window.clearFile = function(inputId) {
        const input = document.getElementById(inputId);
        input.value = '';
        
        if (inputId === 'screenshots') {
            document.getElementById('screenshotsSelected').style.display = 'none';
            document.getElementById('screenshotsList').innerHTML = '';
            input.closest('.file-upload-wrapper').querySelector('.file-upload-content').style.display = 'block';
        } else {
            document.getElementById('templateSelected').style.display = 'none';
            input.closest('.file-upload-wrapper').querySelector('.file-upload-content').style.display = 'block';
        }
    };
});
