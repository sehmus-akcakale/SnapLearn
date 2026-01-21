"""
Gemini API Service
==================
This module handles communication with the Google Gemini API for image analysis.
Uses asynchronous processing to prevent the application from freezing.
"""

import asyncio
import logging
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor

from google import genai
from PIL import Image

from config import GEMINI_API_KEY, GEMINI_MODEL, GEMINI_SYSTEM_INSTRUCTION

# Setup logger
logger = logging.getLogger(__name__)

# Thread pool executor (for running sync API calls asynchronously)
_executor = ThreadPoolExecutor(max_workers=2)

# Global client instance
_client = None


def _get_client():
    """
    Returns the Gemini API client.
    """
    global _client
    if _client is None:
        if not GEMINI_API_KEY:
            raise ValueError("GEMINI_API_KEY is not set!")
        _client = genai.Client(api_key=GEMINI_API_KEY)
        logger.info("Gemini API client created.")
    return _client


def _analyze_image_sync(image_path: Path) -> dict:
    """
    Analyzes an image synchronously (for internal use).
    
    Args:
        image_path: Path to the image file to analyze
        
    Returns:
        dict: Analysis results (contains summary and question)
    """
    try:
        # Get client
        client = _get_client()
        
        logger.info(f"Loading image: {image_path}")
        
        # Load image with PIL
        image = Image.open(image_path)
        
        logger.info("Sending request to Gemini API...")
        
        # Create prompt
        prompt = f"""{GEMINI_SYSTEM_INSTRUCTION}

Please analyze this educational content image:"""
        
        # Analyze image - new API allows PIL Image directly
        response = client.models.generate_content(
            model=GEMINI_MODEL,
            contents=[prompt, image]
        )
        
        # Get response
        response_text = response.text
        logger.info("Gemini API response received.")
        
        # Parse response
        result = _parse_response(response_text)
        
        return result
        
    except Exception as e:
        logger.error(f"Gemini API error: {e}")
        return {
            "summary": f"Error during image analysis: {str(e)}",
            "question": "What do you see in this image?\nA) Option A\nB) Option B\nC) Option C\nD) Option D",
            "raw_response": str(e),
            "success": False
        }


def _parse_response(response_text: str) -> dict:
    """
    Parses the Gemini API response and extracts summary/question sections.
    
    Args:
        response_text: Raw text response from the API
        
    Returns:
        dict: Parsed results
    """
    summary = ""
    question = ""
    
    # Find **Summary:** and **Question:** sections
    text_lower = response_text.lower()
    
    # Find Summary section
    sum_markers = ["**summary:**", "summary:"]
    quest_markers = ["**question:**", "**multiple choice question:**", "question:"]
    
    sum_start = -1
    quest_start = -1
    
    for marker in sum_markers:
        idx = text_lower.find(marker)
        if idx != -1:
            sum_start = idx + len(marker)
            break
    
    for marker in quest_markers:
        idx = text_lower.find(marker)
        if idx != -1:
            quest_start = idx + len(marker)
            break
    
    # Split text into sections
    if sum_start != -1 and quest_start != -1:
        if sum_start < quest_start:
            # Summary comes first
            summary = response_text[sum_start:quest_start].strip()
            # Clean question markers from summary
            for marker in quest_markers:
                summary = summary.replace(marker, "").replace(marker.upper(), "")
            summary = summary.strip().rstrip("*").strip()
            question = response_text[quest_start:].strip()
        else:
            # Question comes first
            question = response_text[quest_start:sum_start].strip()
            summary = response_text[sum_start:].strip()
    elif sum_start != -1:
        summary = response_text[sum_start:].strip()
        question = "What is the main concept shown?\nA) Option A\nB) Option B\nC) Option C\nD) Option D"
    elif quest_start != -1:
        question = response_text[quest_start:].strip()
        summary = response_text[:quest_start].strip()
    else:
        # No markers found, use entire text as summary
        summary = response_text.strip()
        question = "What is the main idea in this image?\nA) Option A\nB) Option B\nC) Option C\nD) Option D"
    
    # Clean up
    summary = summary.strip().strip("*").strip()
    question = question.strip().strip("*").strip()
    
    # Remove "Correct Answer" from question display (keep it separate or at end)
    # Keep the full question with correct answer for educational purposes
    
    return {
        "summary": summary,
        "question": question,
        "raw_response": response_text,
        "success": True
    }


async def analyze_image_async(image_path: Path) -> dict:
    """
    Analyzes an image asynchronously.
    
    This function runs the synchronous API call in a thread pool
    to avoid blocking the main event loop.
    
    Args:
        image_path: Path to the image file to analyze
        
    Returns:
        dict: Analysis results
            - summary: Brief summary of the image content
            - question: Multiple choice test question
            - raw_response: Raw API response
            - success: Whether the operation succeeded
    """
    logger.info(f"Starting async image analysis: {image_path}")
    
    # Run sync function asynchronously
    loop = asyncio.get_event_loop()
    result = await loop.run_in_executor(
        _executor,
        _analyze_image_sync,
        image_path
    )
    
    return result


def analyze_image(image_path: Path) -> dict:
    """
    Analyzes an image synchronously (for simple usage).
    
    Args:
        image_path: Path to the image file to analyze
        
    Returns:
        dict: Analysis results
    """
    return _analyze_image_sync(image_path)


# =============================================================================
# Test / Demo
# =============================================================================

if __name__ == "__main__":
    import sys
    
    # Setup logging
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"
    )
    
    # Get file path from command line
    if len(sys.argv) > 1:
        test_image = Path(sys.argv[1])
        if test_image.exists():
            print(f"Test image: {test_image}")
            result = analyze_image(test_image)
            print("\n=== Result ===")
            print(f"Success: {result.get('success', False)}")
            print(f"\nSummary:\n{result.get('summary', 'N/A')}")
            print(f"\nQuestion:\n{result.get('question', 'N/A')}")
        else:
            print(f"File not found: {test_image}")
    else:
        print("Usage: python gemini_service.py <image_file>")
