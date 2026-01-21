"""
Configuration Module
====================
This module manages application settings and environment variables.
Reads API keys from .env file and defines constants used throughout the application.
"""

import os
from pathlib import Path
from dotenv import load_dotenv

# Load .env file
load_dotenv()

# =============================================================================
# API Configuration
# =============================================================================

# Google Gemini API key
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "")

# Gemini model name (Vision supported - latest model)
GEMINI_MODEL = "gemini-2.5-flash"

# Gemini system instruction - concise summary with multiple choice question
GEMINI_SYSTEM_INSTRUCTION = """Analyze this educational image and provide:

1. **Summary:** A brief, concise summary (2-3 sentences max) of the main concept shown.

2. **Multiple Choice Question:** Create a multiple choice question with 4 options (A, B, C, D) to test understanding. Mark the correct answer.

Format your response exactly as:
**Summary:**
[Your concise summary here]

**Question:**
[Your question here]
A) [Option A]
B) [Option B]
C) [Option C]
D) [Option D]

**Correct Answer:** [Letter]"""

# =============================================================================
# File Paths
# =============================================================================

# Project root directory
BASE_DIR = Path(__file__).parent.absolute()

# Screenshots folder
SCREENSHOTS_DIR = BASE_DIR / "screenshots"

# PowerPoint output folder
OUTPUT_DIR = BASE_DIR / "output"

# =============================================================================
# Application Settings
# =============================================================================

# Hotkey combination - with AI analysis
HOTKEY = "ctrl+v"

# Hotkey combination - direct capture (no AI)
HOTKEY_DIRECT = "ctrl+b"

# Screenshot format
SCREENSHOT_FORMAT = "png"

# Screenshot quality (for JPEG, 1-100)
SCREENSHOT_QUALITY = 95

# =============================================================================
# PowerPoint Settings
# =============================================================================

# Slide dimensions (in inches)
SLIDE_WIDTH_INCHES = 13.333  # Widescreen 16:9
SLIDE_HEIGHT_INCHES = 7.5

# Image size (on slide)
IMAGE_WIDTH_INCHES = 6.0
IMAGE_HEIGHT_INCHES = 5.0

# Text box size
TEXT_BOX_WIDTH_INCHES = 6.0
TEXT_BOX_HEIGHT_INCHES = 5.5

# =============================================================================
# Helper Functions
# =============================================================================

def validate_config() -> tuple[bool, str]:
    """
    Validates configuration settings.
    
    Returns:
        tuple: (is_valid, error_message)
    """
    if not GEMINI_API_KEY:
        return False, "GEMINI_API_KEY environment variable is not set. Please check your .env file."
    
    return True, ""


def ensure_directories():
    """
    Ensures required directories exist.
    Creates them if they don't exist.
    """
    SCREENSHOTS_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


# Create directories when module is loaded
ensure_directories()
