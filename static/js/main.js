document.addEventListener('DOMContentLoaded', function() {
    const form = document.getElementById('uploadForm');
    const submitBtn = document.getElementById('submitBtn');
    const progressSection = document.getElementById('progressSection');
    const resultSection = document.getElementById('resultSection');
    const errorSection = document.getElementById('errorSection');
    const progressFill = document.getElementById('progressFill');
    const progressText = document.getElementById('progressText');
    const resultContent = document.getElementById('resultContent');
    const errorContent = document.getElementById('errorContent');
    const downloadBtn = document.getElementById('downloadBtn');

    form.addEventListener('submit', async function(e) {
        e.preventDefault();
        
        // Hide previous results
        resultSection.style.display = 'none';
        errorSection.style.display = 'none';
        
        // Show progress
        progressSection.style.display = 'block';
        progressFill.style.width = '0%';
        progressText.textContent = 'Uploading files...';
        
        // Disable submit button
        submitBtn.disabled = true;
        submitBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Processing...';
        
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
        
        // Update progress
        progressFill.style.width = '30%';
        progressText.textContent = 'Processing screenshots with OCR...';
        
        try {
            const response = await fetch('/upload', {
                method: 'POST',
                body: formData
            });
            
            const data = await response.json();
            
            if (response.ok && data.success) {
                progressFill.style.width = '100%';
                progressText.textContent = 'Complete!';
                
                // Show success
                setTimeout(() => {
                    progressSection.style.display = 'none';
                    showSuccess(data);
                }, 500);
            } else {
                showError(data.error || 'Processing failed', data.errors);
                resetForm();
            }
        } catch (error) {
            showError('Network error: ' + error.message);
            resetForm();
        }
    });
    
    function showSuccess(data) {
        let html = `<p><strong>Successfully processed ${data.processed} screenshot(s)</strong></p>`;
        
        if (data.failed > 0) {
            html += `<p class="warning">${data.failed} screenshot(s) failed to process</p>`;
        }
        
        if (data.errors && data.errors.length > 0) {
            html += '<ul>';
            data.errors.forEach(error => {
                html += `<li>${error}</li>`;
            });
            html += '</ul>';
        }
        
        resultContent.innerHTML = html;
        
        if (data.download_url) {
            downloadBtn.href = data.download_url;
            downloadBtn.style.display = 'inline-block';
        }
        
        resultSection.style.display = 'block';
        resetForm();
    }
    
    function showError(message, errors = null) {
        let html = `<p>${message}</p>`;
        
        if (errors && errors.length > 0) {
            html += '<ul>';
            errors.forEach(error => {
                html += `<li>${error}</li>`;
            });
            html += '</ul>';
        }
        
        errorContent.innerHTML = html;
        errorSection.style.display = 'block';
    }
    
    function resetForm() {
        submitBtn.disabled = false;
        submitBtn.innerHTML = '<i class="fas fa-upload"></i> Process Screenshots';
        progressFill.style.width = '0%';
    }
});
