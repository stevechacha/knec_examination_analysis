"""
Configuration Management Module
================================
Handles system configuration and settings.
"""

import json
import logging
from pathlib import Path
from typing import Dict, Any
from dataclasses import dataclass, asdict


@dataclass
class Config:
    """System configuration."""
    
    # OCR Settings
    ocr_psm_mode: int = 6  # Page segmentation mode for Tesseract
    ocr_lang: str = 'eng'  # OCR language
    
    # Index Number Patterns
    index_patterns: list = None
    
    # Subject Mappings
    subject_mappings: Dict[str, str] = None
    
    # Excel Settings
    index_column: str = 'INDEXNO'
    name_column: str = 'NAME'
    
    # Logging
    log_level: str = 'INFO'
    log_file: str = 'kcse_system.log'
    
    # File Paths
    default_template: str = 'Students Upload KCSE Results - Template.xlsx'
    
    def __post_init__(self):
        """Initialize default values."""
        if self.index_patterns is None:
            self.index_patterns = [
                r'(\d{10,11}[/\-]\d{2,4})',  # Format: 44748019001/2019
                r'(\d{10,11})',  # Just 10-11 digit numbers
                r'(\d{6,10}[/\-]\d{2,4})',  # Format: 123456/2024
                r'(\d{6,10})',  # Just numbers
            ]
        
        if self.subject_mappings is None:
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
    
    @classmethod
    def from_file(cls, config_path: Path) -> 'Config':
        """Load configuration from JSON file."""
        if not config_path.exists():
            return cls()
        
        try:
            with open(config_path, 'r') as f:
                data = json.load(f)
            return cls(**data)
        except Exception as e:
            logging.warning(f"Failed to load config from {config_path}: {e}. Using defaults.")
            return cls()
    
    def to_file(self, config_path: Path):
        """Save configuration to JSON file."""
        with open(config_path, 'w') as f:
            json.dump(asdict(self), f, indent=2)
    
    def get_log_level(self) -> int:
        """Get logging level as integer."""
        level_map = {
            'DEBUG': logging.DEBUG,
            'INFO': logging.INFO,
            'WARNING': logging.WARNING,
            'ERROR': logging.ERROR,
            'CRITICAL': logging.CRITICAL,
        }
        return level_map.get(self.log_level.upper(), logging.INFO)
