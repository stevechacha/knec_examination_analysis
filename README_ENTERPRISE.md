# KCSE Results Extraction Enterprise System

An enterprise-grade system for extracting KCSE (Kenya Certificate of Secondary Education) results from screenshot images and automatically populating student upload templates with marks.

## Features

- **Automated OCR Processing**: Extract student data from screenshot images using Tesseract OCR
- **Directory-Based Processing**: Process entire directories of screenshots automatically
- **Excel Template Population**: Automatically fill student upload templates with extracted marks
- **Comprehensive Logging**: Detailed logging for debugging and audit trails
- **Error Handling**: Robust error handling with detailed error messages
- **Progress Tracking**: Real-time progress updates during processing
- **Modular Architecture**: Clean, maintainable codebase with separated concerns

## Installation

### Prerequisites

- Python 3.7 or higher
- Tesseract OCR installed on your system

### Install Tesseract OCR

**macOS:**
```bash
brew install tesseract
```

**Ubuntu/Debian:**
```bash
sudo apt-get install tesseract-ocr
```

**Windows:**
Download and install from: https://github.com/UB-Mannheim/tesseract/wiki

### Install Python Dependencies

```bash
pip install -r requirements.txt
```

## Usage

### Basic Usage

```bash
python kcse_enterprise_system.py \
    --screenshots /path/to/screenshots \
    --template "Students Upload KCSE Results - Template.xlsx"
```

### Advanced Usage

```bash
python kcse_enterprise_system.py \
    --screenshots ./screenshots \
    --template template.xlsx \
    --output results_filled.xlsx \
    --config config.json \
    --verbose
```

### Command Line Arguments

- `--screenshots`: **Required**. Directory containing screenshot images (PNG, JPG, JPEG)
- `--template`: **Required**. Path to Excel template file
- `--output`: Optional. Output path for filled Excel file (default: `template_filled_TIMESTAMP.xlsx`)
- `--config`: Optional. Path to JSON configuration file
- `--verbose`, `-v`: Enable verbose logging

## Configuration

Create a `config.json` file to customize system behavior:

```json
{
  "ocr_psm_mode": 6,
  "ocr_lang": "eng",
  "index_column": "INDEXNO",
  "name_column": "NAME",
  "log_level": "INFO",
  "log_file": "kcse_system.log",
  "subject_mappings": {
    "ENGLISH": "ENG",
    "KISWAHILI": "KIS",
    "MATHEMATICS": "MAT"
  }
}
```

## System Architecture

```
kcse_enterprise_system.py     # Main entry point
├── kcse_modules/
│   ├── __init__.py           # Module initialization
│   ├── config.py             # Configuration management
│   ├── logger.py             # Logging setup
│   ├── ocr_extractor.py     # OCR extraction logic
│   └── excel_processor.py    # Excel processing logic
```

## Workflow

1. **Screenshot Processing**: System scans the specified directory for image files
2. **OCR Extraction**: Each image is processed using Tesseract OCR to extract text
3. **Data Parsing**: Extracted text is parsed to identify:
   - Student index numbers
   - Subject grades
   - Mean grades
4. **Template Matching**: Index numbers are matched against the Excel template
5. **Data Population**: Grades are populated into the appropriate columns
6. **Output Generation**: Filled template is saved with timestamp

## Supported Image Formats

- PNG (.png)
- JPEG (.jpg, .jpeg)
- Case-insensitive matching

## Output

The system generates:
- **Filled Excel File**: Template populated with extracted marks
- **Log File**: Detailed processing log (if configured)
- **Console Output**: Progress updates and summary statistics

## Error Handling

The system handles:
- Missing files or directories
- Invalid image formats
- OCR extraction failures
- Excel template mismatches
- Duplicate index numbers

## Logging

Logs are written to:
- **Console**: INFO level and above
- **Log File**: DEBUG level and above (if configured)

Log format includes:
- Timestamp
- Log level
- Module name
- Function name and line number
- Message

## Examples

### Example 1: Process Current Directory

```bash
python kcse_enterprise_system.py \
    --screenshots . \
    --template "Students Upload KCSE Results - Template.xlsx"
```

### Example 2: Custom Output Location

```bash
python kcse_enterprise_system.py \
    --screenshots ./kcse_screenshots \
    --template ./templates/kcse_template.xlsx \
    --output ./output/results_2025.xlsx
```

### Example 3: Verbose Mode with Custom Config

```bash
python kcse_enterprise_system.py \
    --screenshots ./screenshots \
    --template template.xlsx \
    --config ./config/custom_config.json \
    --verbose
```

## Troubleshooting

### OCR Not Working

1. Verify Tesseract is installed: `tesseract --version`
2. Check image quality (should be clear and readable)
3. Try adjusting `ocr_psm_mode` in config

### Index Numbers Not Matching

1. Verify index number format in template matches extracted format
2. Check for extra spaces or formatting differences
3. Review log file for detailed matching information

### Excel File Errors

1. Ensure template file is not open in another program
2. Verify template has required columns (INDEXNO, subject columns)
3. Check file permissions

## Development

### Project Structure

```
.
├── kcse_enterprise_system.py    # Main system
├── kcse_modules/                 # Core modules
│   ├── config.py
│   ├── logger.py
│   ├── ocr_extractor.py
│   └── excel_processor.py
├── requirements.txt              # Dependencies
└── README_ENTERPRISE.md         # This file
```

### Adding New Features

1. Create new module in `kcse_modules/`
2. Import and integrate in `kcse_enterprise_system.py`
3. Update configuration if needed
4. Add tests and documentation

## License

This system is designed for educational and administrative use.

## Support

For issues or questions, please review the log files for detailed error information.
