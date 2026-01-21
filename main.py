"""
Educational Slide Automation Application
=========================================
This application automatically creates educational slides from video content.

Usage:
    1. Start the application: python main.py
    2. Open a video or educational content on your screen
    3. Press Ctrl+V to capture the screen
    4. The application will automatically:
       - Capture the screen
       - Analyze it with Gemini AI
       - Create a PowerPoint slide
    5. Press Ctrl+C to close the application

Requirements:
    - Python 3.10+
    - GEMINI_API_KEY must be defined in .env file
    - Required packages: pip install -r requirements.txt

Author: AI Assistant
Date: 2024
"""

import asyncio
import logging
import sys
import winsound
from pathlib import Path

import keyboard

from config import HOTKEY, HOTKEY_DIRECT, validate_config
from screen_capture import capture_screen
from gemini_service import analyze_image_async
from slide_generator import add_slide, add_slide_direct, get_current_filepath, export_to_pdf

# =============================================================================
# Logging Configuration
# =============================================================================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler("app.log", encoding="utf-8")
    ]
)

logger = logging.getLogger(__name__)

# =============================================================================
# Global Variables
# =============================================================================

# Prevent new triggers while processing
_processing = False

# Event loop reference
_loop: asyncio.AbstractEventLoop = None


# =============================================================================
# Sound Notifications
# =============================================================================

def play_capture_sound():
    """Screen capture sound notification."""
    try:
        winsound.Beep(800, 150)  # High short beep
    except Exception:
        pass


def play_success_sound():
    """Operation successful sound notification."""
    try:
        winsound.Beep(600, 100)
        winsound.Beep(800, 100)
        winsound.Beep(1000, 150)
    except Exception:
        pass


def play_error_sound():
    """Error sound notification."""
    try:
        winsound.Beep(300, 300)
    except Exception:
        pass


# =============================================================================
# Main Workflow
# =============================================================================

async def process_capture():
    """
    Main processing workflow:
    1. Capture screenshot
    2. Analyze with Gemini
    3. Create PowerPoint slide
    """
    global _processing
    
    if _processing:
        logger.warning("Previous operation still in progress, please wait...")
        return
    
    _processing = True
    
    try:
        logger.info("=" * 50)
        logger.info("New capture started...")
        
        # 1. Capture screenshot
        logger.info("Capturing screenshot...")
        play_capture_sound()
        
        image_path = capture_screen()
        
        if not image_path or not image_path.exists():
            logger.error("Failed to capture screenshot!")
            play_error_sound()
            return
        
        logger.info(f"Screenshot saved: {image_path.name}")
        
        # 2. Analyze with Gemini
        logger.info("Analyzing with Gemini AI...")
        
        result = await analyze_image_async(image_path)
        
        if not result.get("success", False):
            logger.error(f"Gemini analysis failed: {result.get('summary', 'Unknown error')}")
            play_error_sound()
            return
        
        summary = result.get("summary", "No summary available.")
        question = result.get("question", "No question generated.")
        
        logger.info("Gemini analysis complete.")
        logger.info(f"   Summary length: {len(summary)} characters")
        
        # 3. Create PowerPoint slide
        logger.info("Creating PowerPoint slide...")
        
        slide_num = add_slide(image_path, summary, question)
        
        logger.info(f"Slide #{slide_num} added successfully!")
        logger.info(f"   Presentation file: {get_current_filepath()}")
        
        play_success_sound()
        
        logger.info("=" * 50)
        
    except Exception as e:
        logger.error(f"Processing error: {e}")
        play_error_sound()
        
    finally:
        _processing = False


def on_hotkey_pressed():
    """
    Called when the AI analysis hotkey (Ctrl+V) is pressed.
    Adds the async task to the event loop.
    """
    global _loop
    
    if _loop is not None and _loop.is_running():
        # Add async task to event loop
        asyncio.run_coroutine_threadsafe(process_capture(), _loop)
    else:
        logger.error("Event loop is not running!")


def on_direct_hotkey_pressed():
    """
    Called when the direct capture hotkey (Ctrl+B) is pressed.
    Captures screen and adds to presentation without AI analysis.
    """
    global _loop
    
    if _loop is not None and _loop.is_running():
        # Add async task to event loop
        asyncio.run_coroutine_threadsafe(process_direct_capture(), _loop)
    else:
        logger.error("Event loop is not running!")


async def process_direct_capture():
    """
    Direct capture workflow (no AI):
    1. Capture screenshot
    2. Add directly to PowerPoint
    """
    global _processing
    
    if _processing:
        logger.warning("Previous operation still in progress, please wait...")
        return
    
    _processing = True
    
    try:
        logger.info("=" * 50)
        logger.info("Direct capture started...")
        
        # 1. Capture screenshot
        logger.info("Capturing screenshot...")
        play_capture_sound()
        
        image_path = capture_screen()
        
        if not image_path or not image_path.exists():
            logger.error("Failed to capture screenshot!")
            play_error_sound()
            return
        
        logger.info(f"Screenshot saved: {image_path.name}")
        
        # 2. Add directly to PowerPoint (no AI)
        logger.info("Adding directly to PowerPoint (no AI analysis)...")
        
        slide_num = add_slide_direct(image_path)
        
        logger.info(f"Slide #{slide_num} added successfully (direct capture)!")
        logger.info(f"   Presentation file: {get_current_filepath()}")
        
        play_success_sound()
        
        logger.info("=" * 50)
        
    except Exception as e:
        logger.error(f"Processing error: {e}")
        play_error_sound()
        
    finally:
        _processing = False


# =============================================================================
# Application Startup
# =============================================================================

def print_banner():
    """Prints the application banner."""
    banner = """
+===============================================================+
|                                                               |
|        EDUCATIONAL SLIDE AUTOMATION APPLICATION               |
|                                                               |
|   Automatically creates educational slides from video content |
|                                                               |
+===============================================================+
|                                                               |
|   Ctrl+V  : Capture with AI analysis (summary + question)     |
|   Ctrl+B  : Direct capture (screenshot only, no AI)           |
|   Ctrl+C  : Close the application                             |
|                                                               |
+===============================================================+
"""
    print(banner)


async def main():
    """
    Main application loop.
    """
    global _loop
    
    # Print banner
    print_banner()
    
    # Validate configuration
    logger.info("Checking configuration...")
    is_valid, error_msg = validate_config()
    
    if not is_valid:
        logger.error(f"Configuration error: {error_msg}")
        logger.error("Please create a .env file and set GEMINI_API_KEY.")
        logger.error("Example: GEMINI_API_KEY=your_api_key_here")
        return
    
    logger.info("Configuration valid.")
    
    # Store event loop reference
    _loop = asyncio.get_event_loop()
    
    # Register hotkeys
    logger.info(f"Registering hotkeys: {HOTKEY.upper()} (AI), {HOTKEY_DIRECT.upper()} (Direct)")
    keyboard.add_hotkey(HOTKEY, on_hotkey_pressed)
    keyboard.add_hotkey(HOTKEY_DIRECT, on_direct_hotkey_pressed)
    logger.info(f"{HOTKEY.upper()} and {HOTKEY_DIRECT.upper()} hotkeys active!")
    
    # Presentation file info
    logger.info(f"Presentation file: {get_current_filepath()}")
    
    print("\n" + "=" * 60)
    print("Application ready!")
    print("  Ctrl+V = Capture with AI analysis")
    print("  Ctrl+B = Direct capture (no AI)")
    print("=" * 60 + "\n")
    
    # Startup sound
    play_success_sound()
    
    try:
        # Infinite loop - keep application running
        while True:
            await asyncio.sleep(0.1)
            
    except KeyboardInterrupt:
        logger.info("\nClosing application...")
        
    finally:
        # Remove hotkeys
        keyboard.unhook_all()
        logger.info("Hotkeys removed.")
        
        # Export to PDF
        print("\nExporting presentation to PDF...")
        logger.info("Exporting presentation to PDF...")
        pdf_path = export_to_pdf()
        
        if pdf_path:
            logger.info(f"PDF exported: {pdf_path}")
            print(f"PDF saved: {pdf_path}")
        else:
            logger.warning("PDF export failed or skipped.")
            print("PDF export failed (PowerPoint may not be installed).")
        
        # Final info
        logger.info(f"Created presentation: {get_current_filepath()}")
        print(f"\nPresentation saved: {get_current_filepath()}")
        print("Application closed successfully. Goodbye!")


# =============================================================================
# Entry Point
# =============================================================================

if __name__ == "__main__":
    try:
        # Set asyncio policy for Windows
        if sys.platform == "win32":
            asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
        
        # Start application
        asyncio.run(main())
        
    except KeyboardInterrupt:
        print("\n\nGoodbye!")
        
    except Exception as e:
        logger.error(f"Critical error: {e}")
        sys.exit(1)
