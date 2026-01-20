# Quick Start Guide - KCSE Enterprise System

## Quick Setup

1. **Install Tesseract OCR:**
   ```bash
   # macOS
   brew install tesseract
   
   # Ubuntu/Debian
   sudo apt-get install tesseract-ocr
   ```

2. **Install Python Dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

## Basic Usage

Process screenshots in the current directory:

```bash
python kcse_enterprise_system.py \
    --screenshots . \
    --template "Students Upload KCSE Results - Template.xlsx"
```

## Example Workflow

1. **Prepare your screenshots:**
   - Place all KCSE result screenshots in a folder (e.g., `./screenshots`)
   - Supported formats: PNG, JPG, JPEG

2. **Run the system:**
   ```bash
   python kcse_enterprise_system.py \
       --screenshots ./screenshots \
       --template "Students Upload KCSE Results - Template.xlsx" \
       --output results_filled.xlsx \
       --verbose
   ```

3. **Check the output:**
   - Filled Excel file will be created with timestamp
   - Review log file for any issues
   - Verify extracted marks in the Excel file

## What the System Does

1. ✅ Scans directory for screenshot images
2. ✅ Extracts text using OCR (Tesseract)
3. ✅ Identifies student index numbers
4. ✅ Extracts subject grades and mean grades
5. ✅ Matches index numbers to Excel template
6. ✅ Populates grades into correct columns
7. ✅ Saves filled template with clean formatting

## Output Files

- **Excel File**: `template_filled_YYYYMMDD_HHMMSS.xlsx`
- **Log File**: `kcse_system.log` (if configured)

## Troubleshooting

**No images found?**
- Check directory path is correct
- Verify image files have .png, .jpg, or .jpeg extension

**OCR not working?**
- Verify Tesseract is installed: `tesseract --version`
- Check image quality (should be clear and readable)

**Index numbers not matching?**
- Review log file for detailed matching information
- Check template has INDEXNO column

## Need More Help?

See `README_ENTERPRISE.md` for detailed documentation.
