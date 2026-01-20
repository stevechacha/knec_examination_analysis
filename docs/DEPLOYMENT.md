# Deployment Guide - KCSE Enterprise System

## Deployment Options

The KCSE Enterprise System can be deployed in multiple ways:

### 1. Web Application (Recommended) 🌐

**Best for:** Multi-user access, easy updates, cloud deployment

#### Local Development Server

```bash
# Install dependencies
pip install -r requirements.txt

# Run the web application
python app.py
# or
python run.py
```

Access at: `http://localhost:5000`

#### Production Deployment

**Option A: Using Gunicorn (Linux/Mac)**

```bash
pip install gunicorn
gunicorn -w 4 -b 0.0.0.0:5000 app:app
```

**Option B: Using Docker**

```dockerfile
FROM python:3.9-slim

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

EXPOSE 5000
CMD ["gunicorn", "-w", "4", "-b", "0.0.0.0:5000", "app:app"]
```

**Option C: Cloud Platforms**

- **Heroku**: Add `Procfile` with `web: gunicorn app:app`
- **AWS Elastic Beanstalk**: Deploy as Python application
- **Google Cloud Run**: Containerize and deploy
- **Azure App Service**: Deploy as Python web app

### 2. Desktop Application 💻

**Best for:** Offline use, single-user, no internet required

#### Option A: Electron Wrapper

```bash
npm install electron
# Wrap Flask app in Electron
```

#### Option B: PyInstaller (Standalone Executable)

```bash
pip install pyinstaller
pyinstaller --onefile --windowed app.py
```

#### Option C: Tkinter GUI (Python Native)

Create a desktop GUI using Tkinter or PyQt.

### 3. Command Line Interface (CLI) 📟

**Best for:** Automation, scripting, batch processing

```bash
python kcse_enterprise_system.py \
    --screenshots ./screenshots \
    --template templates/Students\ Upload\ KCSE\ Results\ -\ Template.xlsx \
    --output results.xlsx
```

## Quick Start - Web Application

### Installation

1. **Install Python 3.7+**
2. **Install Tesseract OCR**
   - macOS: `brew install tesseract`
   - Ubuntu: `sudo apt-get install tesseract-ocr`
   - Windows: Download from GitHub
3. **Install Python dependencies**
   ```bash
   pip install -r requirements.txt
   ```

### Running

```bash
# Development mode
python app.py

# Production mode
python run.py
```

### Access

Open your browser and navigate to:
- Local: `http://localhost:5000`
- Network: `http://YOUR_IP:5000`

## Features

- ✅ **Web Interface**: Modern, responsive UI
- ✅ **File Upload**: Drag & drop or click to upload
- ✅ **Progress Tracking**: Real-time processing status
- ✅ **Error Handling**: Detailed error messages
- ✅ **Download Results**: Processed Excel files ready to download
- ✅ **Multi-file Support**: Process multiple screenshots at once

## Configuration

### Environment Variables

```bash
export SECRET_KEY="your-secret-key-here"
export FLASK_ENV="production"
export MAX_UPLOAD_SIZE="500MB"
```

### Custom Port

```bash
# Change port in app.py or run.py
app.run(host='0.0.0.0', port=8080)
```

## Security Considerations

1. **Change SECRET_KEY** in production
2. **Use HTTPS** in production (nginx reverse proxy)
3. **Set file size limits** appropriately
4. **Implement authentication** for multi-user scenarios
5. **Sanitize file uploads** (already implemented)

## Performance

- **Single-threaded**: Suitable for small to medium workloads
- **Multi-worker**: Use Gunicorn with multiple workers for production
- **Caching**: Consider Redis for session management
- **CDN**: Serve static files via CDN for better performance

## Troubleshooting

### OCR Not Working
- Verify Tesseract is installed: `tesseract --version`
- Check image quality (should be clear and readable)

### Port Already in Use
- Change port in `app.py` or `run.py`
- Kill existing process: `lsof -ti:5000 | xargs kill`

### File Upload Errors
- Check file size limits
- Verify file permissions
- Check disk space

## Next Steps

1. **Add Authentication**: Implement user login
2. **Database Integration**: Store processing history
3. **API Endpoints**: Create REST API for integrations
4. **Monitoring**: Add logging and monitoring
5. **Scaling**: Deploy to cloud with auto-scaling
