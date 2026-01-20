"""
OCR Extraction Module
======================
Handles OCR extraction from screenshot images.
"""

import re
import logging
from pathlib import Path
from typing import Dict, Optional, Tuple
from PIL import Image
import pytesseract


class OCRExtractor:
    """Extract KCSE results data from images using OCR."""
    
    def __init__(self, logger: logging.Logger, config=None):
        """
        Initialize OCR extractor.
        
        Args:
            logger: Logger instance
            config: Optional configuration object
        """
        self.logger = logger
        self.config = config
        
        # Default subject mappings
        self.subject_mappings = {
            'ENGLISH': 'ENG',
            'KISWAHILI': 'KIS',
            'MATHEMATICS': 'MAT',
            'BIOLOGY': 'BIO',
            'PHYSICS': 'PHY',
            'CHEMISTRY': 'CHE',
            'GEOGRAPHY': 'GEO',
            'HISTORY': 'HIS',
            'HISTORY AND GOVERNMENT': 'HIS',
            'CHRISTIAN RELIGIOUS EDUCATION': 'CRE',
            'CHRISTIAN RELIGIOUS': 'CRE',
            'AGRICULTURE': 'AGR',
            'BUSINESS STUDIES': 'BST',
            'BUSINESS': 'BST',
            'COMPUTER STUDIES': 'COM',
            'COMPUTER': 'COM',
        }
        
        if config and hasattr(config, 'subject_mappings'):
            self.subject_mappings.update(config.subject_mappings)
    
    def extract_text_from_image(self, image_path: Path) -> Optional[str]:
        """
        Extract text from image using OCR.
        
        Args:
            image_path: Path to image file
            
        Returns:
            Extracted text or None if extraction fails
        """
        try:
            image = Image.open(image_path)
            psm_mode = self.config.ocr_psm_mode if self.config else 6
            text = pytesseract.image_to_string(image, config=f'--psm {psm_mode}')
            return text
        except Exception as e:
            self.logger.error(f"Error processing {image_path}: {e}", exc_info=True)
            return None
    
    def extract_index_number(self, text: str) -> Optional[str]:
        """
        Extract index number from OCR text.
        
        Args:
            text: OCR extracted text
            
        Returns:
            Index number string or None
        """
        index_patterns = [
            r'(\d{10,11}[/\-]\d{2,4})',  # Format: 44748019001/2019
            r'(\d{10,11})',  # Just 10-11 digit numbers
            r'(\d{6,10}[/\-]\d{2,4})',  # Format: 123456/2024
            r'(\d{6,10})',  # Just numbers
        ]
        
        if self.config and hasattr(self.config, 'index_patterns'):
            index_patterns = self.config.index_patterns
        
        best_match = None
        best_length = 0
        best_score = -1
        
        # Try pattern: index followed by dash and capital letters (student name)
        name_after_index_pattern = r'(\d{10,11})\s*-\s*[A-Z][A-Z\s]+'
        name_match = re.search(name_after_index_pattern, text)
        if name_match:
            potential_index = name_match.group(1)
            if not potential_index.startswith('123456'):
                best_match = potential_index
                best_length = len(potential_index)
                best_score = 100
        
        # Try other patterns
        if best_score < 100:
            for pattern in index_patterns:
                index_matches = re.findall(pattern, text)
                for match in index_matches:
                    match_str = str(match).split('/')[0]
                    if match_str.startswith('123456'):
                        continue
                    score = 10 if match_str.startswith('4474801') else 5
                    match_length = len(match_str)
                    if score > best_score or (score == best_score and match_length > best_length):
                        best_match = str(match)
                        best_length = match_length
                        best_score = score
        
        if best_match:
            index_number = best_match
            # Add year suffix if missing
            if '/' not in index_number and len(index_number) >= 6:
                year_match = re.search(r'(19|20)\d{2}', text)
                if year_match:
                    index_number = f"{index_number}/{year_match.group()}"
            return index_number
        
        # Fallback: hard search for 11-digit numbers
        hard_match = re.search(r'4474801[789]\d{3}', text)
        if hard_match:
            return hard_match.group()
        
        return None
    
    def extract_subject_grades(self, text: str) -> Tuple[Dict[str, str], Optional[str]]:
        """
        Extract subject-grade pairs and mean grade from OCR text.
        
        Args:
            text: OCR extracted text
            
        Returns:
            Tuple of (subject_grades dict, mean_grade string)
        """
        subject_grades = {}
        mean_grade = None
        
        # Extract subject-grade pairs line by line
        subject_matches = []
        lines = text.split('\n')
        
        for line in lines:
            line = line.strip()
            if not line:
                continue
            
            # Pattern: number, code, subject words, then grade
            pattern = r'^(\d+)\s+(\d+)\s+(.+?)\s+([A-E][+-]?|X)(?:\s*\([^)]+\))?\s*$'
            match = re.match(pattern, line, re.IGNORECASE)
            if match:
                subject_name = ' '.join(match.group(3).split()).strip()
                grade = match.group(4).strip()
                
                if grade in ['A', 'A-', 'B+', 'B', 'B-', 'C+', 'C', 'C-', 'D+', 'D', 'D-', 'E', 'X']:
                    subject_matches.append((subject_name, grade))
        
        # Fallback pattern if line-by-line didn't work
        if len(subject_matches) < 3:
            pattern = r'(\d+)\s+(\d+)\s+([A-Z][A-Z\s]+?)\s+([A-E][+-]?|X)(?:\s*\([^)]+\))?'
            matches = re.findall(pattern, text, re.IGNORECASE)
            if matches and len(matches) >= 3:
                cleaned_matches = []
                for m in matches:
                    subject_name = ' '.join(m[2].split()).strip()
                    grade = m[3].strip()
                    if grade in ['A', 'A-', 'B+', 'B', 'B-', 'C+', 'C', 'C-', 'D+', 'D', 'D-', 'E', 'X']:
                        if not subject_name.upper().endswith(grade[0]):
                            cleaned_matches.append((subject_name, grade))
                if cleaned_matches:
                    subject_matches = cleaned_matches
        
        # Map subject names to column names
        for subject_name, grade in subject_matches:
            subject_name = subject_name.strip().upper()
            mapped = False
            for key, col_name in self.subject_mappings.items():
                if key in subject_name:
                    if grade != 'X' and grade.upper() != 'X':
                        subject_grades[col_name] = grade
                    mapped = True
                    break
        
        # Extract mean grade
        mean_grade_patterns = [
            r'Mean\s+Grade[:\s]+([A-E][+-]?)\s*\(',
            r'Mean\s+Grade[:\s]+([A-E][+-]?)\s',
            r'Mean\s+Grade[:\s]+([A-E][+-]?)$',
            r'Mean[:\s]+([A-E][+-]?)\s',
        ]
        
        for pattern in mean_grade_patterns:
            mean_match = re.search(pattern, text, re.IGNORECASE)
            if mean_match:
                mean_grade = mean_match.group(1).strip()
                if mean_grade in ['A', 'A-', 'B+', 'B', 'B-', 'C+', 'C', 'C-', 'D+', 'D', 'D-', 'E']:
                    break
        
        return subject_grades, mean_grade
    
    def normalize_index_number(self, index_num: str) -> Optional[str]:
        """
        Normalize index number format.
        
        Args:
            index_num: Index number string
            
        Returns:
            Normalized index number or None
        """
        if not index_num:
            return None
        
        index_str = str(index_num).strip()
        # Remove .0 suffix if present
        if index_str.endswith('.0'):
            index_str = index_str[:-2]
        # Remove year suffix
        index_str = re.sub(r'/\d{4}$', '', index_str)
        # Remove whitespace
        index_str = re.sub(r'\s+', '', index_str)
        
        return index_str if index_str else None
    
    def extract_from_image(self, image_path: Path) -> Optional[Dict]:
        """
        Extract all data from an image.
        
        Args:
            image_path: Path to image file
            
        Returns:
            Dictionary with extracted data or None
        """
        self.logger.debug(f"Extracting data from {image_path.name}")
        
        text = self.extract_text_from_image(image_path)
        if not text:
            return None
        
        index_number = self.extract_index_number(text)
        if not index_number:
            self.logger.warning(f"Could not extract index number from {image_path.name}")
            return None
        
        # Skip placeholder indices
        if '12345678' in index_number or '1234567' in index_number:
            self.logger.warning(f"Skipping placeholder index: {index_number}")
            return None
        
        subject_grades, mean_grade = self.extract_subject_grades(text)
        
        normalized_idx = self.normalize_index_number(index_number)
        
        return {
            'index_number': normalized_idx or index_number,
            'original_index': index_number,
            'subject_grades': subject_grades,
            'mean_grade': mean_grade,
            'source_file': image_path.name,
        }
