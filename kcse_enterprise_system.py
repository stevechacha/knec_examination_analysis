#!/usr/bin/env python3
"""
KCSE Results Extraction Enterprise System
=========================================
An enterprise-grade system for extracting KCSE results from screenshots
and populating student upload templates with marks.

Features:
- Directory-based screenshot processing
- OCR-based data extraction
- Automatic Excel template population
- Comprehensive logging and error handling
- Progress tracking and reporting
"""

import argparse
import logging
import sys
from pathlib import Path
from typing import Dict, List, Optional
from datetime import datetime

try:
    import pandas as pd
    from PIL import Image
    import pytesseract
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter
except ImportError as e:
    print(f"Error: Missing required package. {e}")
    print("Please install required packages:")
    print("  pip3 install pandas openpyxl pillow pytesseract")
    sys.exit(1)

try:
    from kcse_modules.ocr_extractor import OCRExtractor
    from kcse_modules.excel_processor import ExcelProcessor
    from kcse_modules.config import Config
    from kcse_modules.logger import setup_logger
except ImportError:
    # Fallback for direct execution
    import sys
    sys.path.insert(0, str(Path(__file__).parent))
    from kcse_modules.ocr_extractor import OCRExtractor
    from kcse_modules.excel_processor import ExcelProcessor
    from kcse_modules.config import Config
    from kcse_modules.logger import setup_logger


class KCSEEnterpriseSystem:
    """Main enterprise system for KCSE results processing."""
    
    def __init__(self, config_path: Optional[Path] = None):
        """Initialize the system with configuration."""
        if config_path:
            self.config = Config.from_file(config_path)
        else:
            self.config = Config()
        self.logger = setup_logger(self.config.log_level, self.config.log_file)
        self.ocr_extractor = OCRExtractor(self.logger, self.config)
        self.excel_processor = ExcelProcessor(self.logger, self.config)
        
    def process_screenshots_directory(self, screenshots_dir: Path) -> Dict:
        """
        Process all screenshots in a directory.
        
        Args:
            screenshots_dir: Path to directory containing screenshot images
            
        Returns:
            Dictionary containing extracted data
        """
        self.logger.info(f"Processing screenshots from directory: {screenshots_dir}")
        
        if not screenshots_dir.exists():
            self.logger.error(f"Directory does not exist: {screenshots_dir}")
            raise FileNotFoundError(f"Directory not found: {screenshots_dir}")
        
        # Find all image files
        image_extensions = ['.png', '.jpg', '.jpeg', '.PNG', '.JPG', '.JPEG']
        image_files = []
        for ext in image_extensions:
            image_files.extend(list(screenshots_dir.glob(f"*{ext}")))
        
        if not image_files:
            self.logger.warning(f"No image files found in {screenshots_dir}")
            return {}
        
        self.logger.info(f"Found {len(image_files)} image files")
        
        # Process each image
        extracted_data = {}
        processed_count = 0
        failed_count = 0
        
        for i, image_path in enumerate(sorted(image_files), 1):
            self.logger.info(f"Processing {i}/{len(image_files)}: {image_path.name}")
            
            try:
                result = self.ocr_extractor.extract_from_image(image_path)
                if result and result.get('index_number'):
                    index_num = result['index_number']
                    extracted_data[index_num] = result
                    processed_count += 1
                    self.logger.info(f"  ✓ Extracted index: {index_num}")
                else:
                    failed_count += 1
                    self.logger.warning(f"  ✗ Failed to extract data from {image_path.name}")
            except Exception as e:
                failed_count += 1
                self.logger.error(f"  ✗ Error processing {image_path.name}: {e}", exc_info=True)
        
        self.logger.info(f"Processing complete: {processed_count} successful, {failed_count} failed")
        return extracted_data
    
    def update_excel_template(self, template_path: Path, extracted_data: Dict, 
                             output_path: Optional[Path] = None) -> Path:
        """
        Update Excel template with extracted data.
        
        Args:
            template_path: Path to Excel template file
            extracted_data: Dictionary of extracted data keyed by index number
            output_path: Optional output path (defaults to template_path with _filled suffix)
            
        Returns:
            Path to the updated Excel file
        """
        self.logger.info(f"Updating Excel template: {template_path}")
        
        if not template_path.exists():
            raise FileNotFoundError(f"Template file not found: {template_path}")
        
        if output_path is None:
            output_path = template_path.parent / f"{template_path.stem}_filled_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        
        return self.excel_processor.update_template(template_path, extracted_data, output_path)
    
    def run(self, screenshots_dir: Path, template_path: Path, 
            output_path: Optional[Path] = None) -> Path:
        """
        Run the complete processing pipeline.
        
        Args:
            screenshots_dir: Directory containing screenshot images
            template_path: Path to Excel template file
            output_path: Optional output path for filled template
            
        Returns:
            Path to the output Excel file
        """
        self.logger.info("=" * 60)
        self.logger.info("KCSE Enterprise System - Starting Processing")
        self.logger.info("=" * 60)
        
        start_time = datetime.now()
        
        # Step 1: Extract data from screenshots
        self.logger.info("Step 1: Extracting data from screenshots...")
        extracted_data = self.process_screenshots_directory(screenshots_dir)
        
        if not extracted_data:
            self.logger.error("No data extracted from screenshots. Exiting.")
            raise ValueError("No data extracted from screenshots")
        
        self.logger.info(f"Extracted data for {len(extracted_data)} students")
        
        # Step 2: Update Excel template
        self.logger.info("Step 2: Updating Excel template...")
        output_file = self.update_excel_template(template_path, extracted_data, output_path)
        
        # Step 3: Generate summary report
        elapsed_time = (datetime.now() - start_time).total_seconds()
        self.logger.info("=" * 60)
        self.logger.info("Processing Complete!")
        self.logger.info(f"Output file: {output_file}")
        self.logger.info(f"Students processed: {len(extracted_data)}")
        self.logger.info(f"Processing time: {elapsed_time:.2f} seconds")
        self.logger.info("=" * 60)
        
        return output_file


def main():
    """Main entry point for CLI."""
    parser = argparse.ArgumentParser(
        description="KCSE Results Extraction Enterprise System",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Process screenshots in current directory
  python kcse_enterprise_system.py --screenshots . --template templates/template.xlsx
  
  # Specify output file
  python kcse_enterprise_system.py --screenshots ./screenshots --template templates/template.xlsx --output results.xlsx
  
  # Use custom config
  python kcse_enterprise_system.py --screenshots . --template templates/template.xlsx --config config.json
        """
    )
    
    parser.add_argument(
        '--screenshots',
        type=Path,
        required=True,
        help='Directory containing screenshot images'
    )
    
    parser.add_argument(
        '--template',
        type=Path,
        required=True,
        help='Path to Excel template file'
    )
    
    parser.add_argument(
        '--output',
        type=Path,
        default=None,
        help='Output path for filled Excel file (default: template_filled_TIMESTAMP.xlsx)'
    )
    
    parser.add_argument(
        '--config',
        type=Path,
        default=None,
        help='Path to configuration file (optional)'
    )
    
    parser.add_argument(
        '--verbose',
        '-v',
        action='store_true',
        help='Enable verbose logging'
    )
    
    args = parser.parse_args()
    
    # Initialize system
    try:
        system = KCSEEnterpriseSystem(config_path=args.config)
        
        if args.verbose:
            system.logger.setLevel(logging.DEBUG)
        
        # Run processing
        output_file = system.run(
            screenshots_dir=args.screenshots,
            template_path=args.template,
            output_path=args.output
        )
        
        print(f"\n✓ Success! Output saved to: {output_file}")
        sys.exit(0)
        
    except Exception as e:
        print(f"\n✗ Error: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
