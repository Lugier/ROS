import os
import json
import time
import logging
import tempfile
import re
from aspose.slides import Presentation
from google import genai
from google.genai import types
from dotenv import load_dotenv
import PIL.Image

load_dotenv()
logger = logging.getLogger(__name__)


def extract_json_from_text(text):
    """
    Extrahiert JSON auch wenn das Modell Markdown oder Text drumherum schreibt.
    """
    text = text.strip()

    # 1. Versuche Markdown Code Blocks zu finden (```json ... ```)
    code_block_pattern = r"```(?:json)?\s*(\{.*?\})\s*```"
    match = re.search(code_block_pattern, text, re.DOTALL)
    if match:
        return match.group(1)

    # 2. Fallback: Finde das erste '{' und das letzte '}'
    start_idx = text.find('{')
    end_idx = text.rfind('}')

    if start_idx != -1 and end_idx != -1:
        return text[start_idx : end_idx + 1]

    return text


def analyze_slide_and_research(pptx_path, user_prompt):
    """
    Analyzes the first slide of a PPTX file using Gemini Vision,
    researches content using Google Search, and returns JSON instructions.
    
    Uses the new google-genai SDK with proper Google Search grounding support.
    
    Args:
        pptx_path: Path to the PPTX template file
        user_prompt: User's prompt for content adaptation
        
    Returns:
        dict: JSON instructions with replacements, charts, and think_cell_replacements
    """
    # Step A: Load PPTX and render first slide as PNG
    step_start = time.time()
    logger.info("    [Step A] Loading PPTX file...")
    pres = Presentation(pptx_path)
    slide = pres.slides[0]
    logger.info(f"    [Step A] ✓ Loaded presentation with {len(pres.slides)} slide(s)")
    
    # Create temporary PNG file
    temp_png = tempfile.NamedTemporaryFile(delete=False, suffix='.png')
    temp_png_path = temp_png.name
    temp_png.close()
    
    # Render slide to PNG
    logger.info("    [Step A] Rendering first slide to PNG...")
    slide.get_thumbnail(1.0, 1.0).save(temp_png_path)
    step_time = time.time() - step_start
    logger.info(f"    [Step A] ✓ Slide rendered to PNG in {step_time:.2f}s ({os.path.getsize(temp_png_path)} bytes)")
    
    # Step B: Initialize Gemini Client with new SDK
    step_start = time.time()
    logger.info("    [Step B] Initializing Gemini Client (google-genai SDK)...")
    api_key = os.getenv('GEMINI_API_KEY')
    if not api_key:
        raise ValueError("GEMINI_API_KEY not found in environment variables")
    
    client = genai.Client(api_key=api_key)
    logger.info("    [Step B] ✓ Client initialized")
    
    # Step C: Configure Google Search tool
    logger.info("    [Step C] Configuring Google Search tool...")
    
    try:
        # Use the new SDK's Tool and GoogleSearch types
        tool = types.Tool(
            google_search=types.GoogleSearch()
        )
        
        config = types.GenerateContentConfig(
            tools=[tool],
            response_mime_type="application/json"
        )
        
        logger.info("    [Step C] ✓ Google Search tool configured")
        use_google_search = True
        
    except (AttributeError, ImportError, Exception) as e:
        logger.warning(f"    [Step C] Google Search tool config failed: {e}")
        logger.info("    [Step C] Using config without explicit tools (Search may be built-in)")
        
        # Fallback: config without explicit tools
        config = types.GenerateContentConfig(
            response_mime_type="application/json"
        )
        use_google_search = False
    
    step_time = time.time() - step_start
    logger.info(f"    [Step B+C] ✓ Configuration completed in {step_time:.2f}s")
    
    # Step D: Prepare system prompt
    system_prompt = f"""You are a Presentation Architect. 

1. Analyze the image visually. Identify colors and chart styles. 

2. Use Google Search to find real-time data for: '{user_prompt}'. 

3. Return a JSON object with: 

   - 'replacements': list of {{'old_text_snippet': str, 'new_text': str}}. 

   - 'charts': list of {{'type': str (bar/column/line), 'title': str, 'data': {{'categories': [str], 'series': [{{'name': str, 'values': [float], 'color_hex': str}}]}}}}. 

   - 'think_cell_replacements': boolean (if the chart in the image looks like Think-Cell/Waterfall, mark it true).

4. Ensure numbers are accurate based on search.

Return ONLY valid JSON, no markdown formatting, no code blocks."""

    # Step E: Prepare and send to Gemini
    step_start = time.time()
    logger.info("    [Step D] Preparing image and prompt for Gemini...")
    
    # WICHTIG: Wir öffnen das Bild, nutzen es und schließen es sicher
    image = PIL.Image.open(temp_png_path)
    
    try:
        image_size = f"{image.size[0]}x{image.size[1]}"
        logger.info(f"    [Step D] ✓ Image loaded: {image_size} pixels")
        
        # Model name - Gemini 3 Pro Preview for best quality
        model_name = "gemini-3-pro-preview"
        
        logger.info(f"    [Step D] Using model: {model_name}")
        logger.info("    [Step D] Sending request to Gemini (this may take 30-90s)...")
        logger.info("    [Step D] → Performing vision analysis...")
        if use_google_search:
            logger.info("    [Step D] → Executing Google Search for research...")
        logger.info("    [Step D] → Generating JSON instructions...")
        
        # Send request using new SDK
        response = client.models.generate_content(
            model=model_name,
            contents=[system_prompt, image],
            config=config
        )
        
    finally:
        # WICHTIG: Bild explizit schließen, damit Windows den File-Handle freigibt
        image.close()
    
    step_time = time.time() - step_start
    logger.info(f"    [Step D] ✓ Gemini response received in {step_time:.2f}s")
    
    # Jetzt kann die Datei sicher gelöscht werden
    try:
        os.unlink(temp_png_path)
        logger.info("    [Step D] ✓ Cleaned up temporary PNG file")
    except Exception as e:
        logger.warning(f"    [Step D] ⚠️ Could not delete temp PNG (ignoring): {e}")
    
    # --- ROBUST JSON PARSING (FIXED) ---
    logger.info("    [Step D] Parsing JSON response...")
    response_text = response.text.strip()

    try:
        # 1. Clean string mit der Funktion (entfernt Markdown-Code-Blöcke)
        json_str = extract_json_from_text(response_text)
        
        # 2. Parse
        json_instructions = json.loads(json_str)
        
        # 3. Validate structure
        if "replacements" not in json_instructions and "charts" not in json_instructions:
            logger.warning("    [Step D] ⚠️ JSON valid but keys missing (replacements/charts)")
        
        logger.info(f"    [Step D] ✓ JSON parsed successfully ({len(str(json_instructions))} chars)")
        return json_instructions

    except json.JSONDecodeError as e:
        logger.error(f"    [Step D] ✗ JSON parsing failed: {e}")
        logger.error(f"    [Step D] Extracted text: {json_str[:500] if 'json_str' in locals() else 'N/A'}...")
        logger.error(f"    [Step D] Raw text start: {response_text[:500]}...")
        
        # Fallback empty
        return {"replacements": [], "charts": []}
    
    except Exception as e:
        logger.error(f"    [Step D] ✗ General error parsing response: {e}")
        return {"replacements": [], "charts": []}
