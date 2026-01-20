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
        let html = `<p style="font-size: 1.2em; font-weight: 600; margin-bottom: 15px;">
            <i class="fas fa-check-circle"></i> Successfully processed ${data.processed} screenshot(s)
        </p>`;
        
        if (data.failed > 0) {
            html += `<p class="warning" style="margin: 15px 0;">
                <i class="fas fa-exclamation-triangle"></i> ${data.failed} screenshot(s) failed to process
            </p>`;
        }
        
        if (data.errors && data.errors.length > 0) {
            html += '<div style="margin-top: 20px;"><strong>Details:</strong><ul>';
            data.errors.forEach(error => {
                html += `<li>${error}</li>`;
            });
            html += '</ul></div>';
        }
        
        resultContent.innerHTML = html;
        
        if (data.download_url) {
            downloadBtn.href = data.download_url;
            downloadBtn.style.display = 'inline-flex';
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
