"""
Configuration settings and constants
"""

import os
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# API Configuration
ANTHROPIC_API_KEY = os.getenv('ANTHROPIC_API_KEY')
CLAUDE_MODEL = "claude-3-5-sonnet-20241022"

# Processing Configuration
DEFAULT_CHUNK_SIZE = 50
DEFAULT_OVERLAP = 10
MAX_RETRIES = 5
BASE_DELAY = 1

# Column Analysis Configuration
MIN_TEXT_LENGTH_FOR_QUESTIONS = 100
SHORT_TEXT_THRESHOLD = 20
QUESTION_COLUMN_THRESHOLD = 0.3
ANSWER_COLUMN_THRESHOLD = 0.3

# Filling Configuration
DEFAULT_POSITIVE_PERCENTAGE = 70
DEFAULT_NEGATIVE_PERCENTAGE = 15
DEFAULT_PARTIAL_PERCENTAGE = 15

# File Configuration
MAX_SAMPLE_ROWS = 20
MAX_SAMPLE_COLUMNS = 25
MAX_TEXT_CONTENT_ITEMS = 100
MAX_SECTIONS = 20

# Output Configuration
TIMESTAMP_FORMAT = "%Y%m%d_%H%M%S"
OUTPUT_DIR_PREFIX = "enhanced_extraction_"
FILLED_FILE_PREFIX = "enhanced_filled_"

# Validation Configuration
SUPPORTED_EXTENSIONS = ['.xlsx', '.xls']
MIN_ROWS_FOR_PROCESSING = 2
MIN_COLUMNS_FOR_PROCESSING = 2

def validate_environment():
    """Validate required environment variables"""
    if not ANTHROPIC_API_KEY:
        raise ValueError(
            "ANTHROPIC_API_KEY environment variable is required. "
            "Please set it in your .env file or environment."
        )
    
    return True