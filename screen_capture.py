"""
Screen Capture Module
=====================
This module handles screen capture operations using the Pillow (PIL) library.
Captured images are saved to the screenshots folder with timestamps.
"""

import logging
from datetime import datetime
from pathlib import Path
from PIL import ImageGrab

from config import SCREENSHOTS_DIR, SCREENSHOT_FORMAT, SCREENSHOT_QUALITY

# Setup logger
logger = logging.getLogger(__name__)


def capture_screen() -> Path | None:
    """
    Captures a full screenshot of the current screen and saves it to a file.
    
    The image is saved with a unique timestamp-based filename.
    Example: screenshot_2024-01-15_14-30-45.png
    
    Returns:
        Path | None: File path of the saved image or None on error
    
    Raises:
        Any exception is caught and logged.
    """
    try:
        # Create unique filename with timestamp
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        filename = f"screenshot_{timestamp}.{SCREENSHOT_FORMAT}"
        filepath = SCREENSHOTS_DIR / filename
        
        logger.info(f"Capturing screenshot: {filename}")
        
        # Capture screenshot (entire screen)
        screenshot = ImageGrab.grab(all_screens=False)
        
        # Save image
        if SCREENSHOT_FORMAT.lower() == "png":
            screenshot.save(filepath, format="PNG", optimize=True)
        elif SCREENSHOT_FORMAT.lower() in ("jpg", "jpeg"):
            screenshot.save(
                filepath, 
                format="JPEG", 
                quality=SCREENSHOT_QUALITY, 
                optimize=True
            )
        else:
            screenshot.save(filepath)
        
        logger.info(f"Screenshot saved: {filepath}")
        logger.info(f"Image size: {screenshot.size[0]}x{screenshot.size[1]} pixels")
        
        return filepath
        
    except Exception as e:
        logger.error(f"Failed to capture screenshot: {e}")
        return None


def capture_region(bbox: tuple[int, int, int, int]) -> Path | None:
    """
    Captures a specific region of the screen.
    
    Args:
        bbox: Region coordinates (left, top, right, bottom)
        
    Returns:
        Path | None: File path of the saved image or None on error
    """
    try:
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        filename = f"screenshot_region_{timestamp}.{SCREENSHOT_FORMAT}"
        filepath = SCREENSHOTS_DIR / filename
        
        logger.info(f"Capturing region: {bbox}")
        
        # Capture the specified region
        screenshot = ImageGrab.grab(bbox=bbox)
        screenshot.save(filepath, format=SCREENSHOT_FORMAT.upper())
        
        logger.info(f"Region screenshot saved: {filepath}")
        
        return filepath
        
    except Exception as e:
        logger.error(f"Failed to capture region: {e}")
        return None


def get_screen_size() -> tuple[int, int]:
    """
    Returns the screen dimensions.
    
    Returns:
        tuple: (width, height) in pixels
    """
    try:
        screenshot = ImageGrab.grab()
        size = screenshot.size
        screenshot.close()
        return size
    except Exception as e:
        logger.error(f"Failed to get screen size: {e}")
        return (1920, 1080)  # Default value


def cleanup_old_screenshots(max_age_days: int = 7) -> int:
    """
    Removes screenshots older than a specified age.
    
    Args:
        max_age_days: Maximum file age (in days)
        
    Returns:
        int: Number of deleted files
    """
    import time
    
    deleted_count = 0
    current_time = time.time()
    max_age_seconds = max_age_days * 24 * 60 * 60
    
    try:
        for filepath in SCREENSHOTS_DIR.glob(f"*.{SCREENSHOT_FORMAT}"):
            file_age = current_time - filepath.stat().st_mtime
            if file_age > max_age_seconds:
                filepath.unlink()
                deleted_count += 1
                logger.info(f"Deleted old file: {filepath}")
                
        if deleted_count > 0:
            logger.info(f"Cleaned up {deleted_count} old files.")
            
    except Exception as e:
        logger.error(f"File cleanup error: {e}")
        
    return deleted_count


# =============================================================================
# Test / Demo
# =============================================================================

if __name__ == "__main__":
    # Setup logging
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"
    )
    
    print("Screen capture test...")
    print(f"Screen size: {get_screen_size()}")
    
    filepath = capture_screen()
    if filepath:
        print(f"Success! File: {filepath}")
    else:
        print("Error occurred!")
