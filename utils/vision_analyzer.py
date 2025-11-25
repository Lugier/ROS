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
        # Tool Definition
        tool = types.Tool(google_search=types.GoogleSearch())
        
        # --- FIX: JSON Mode für Flash DEAKTIVIEREN ---
        # Wir entfernen response_mime_type="application/json", 
        # weil Flash das nicht zusammen mit Tools unterstützt.
        config = types.GenerateContentConfig(
            tools=[tool],
            # response_mime_type="application/json"  <-- ENTFERNT für Flash-Kompatibilität!
        )
        
        logger.info("    [Step C] ✓ Google Search configured (JSON Mode disabled for Flash compatibility)")
        use_google_search = True
        
    except (AttributeError, ImportError, Exception) as e:
        logger.warning(f"    [Step C] Config failed: {e}")
        logger.info("    [Step C] Using config without explicit tools (Search may be built-in)")
        
        # Fallback: config without explicit tools (auch ohne JSON Mode)
        config = types.GenerateContentConfig()
        use_google_search = False
    
    step_time = time.time() - step_start
    logger.info(f"    [Step B+C] ✓ Configuration completed in {step_time:.2f}s")
    
    # Step D: Prepare system prompt (Stricter JSON rules)
    system_prompt = f"""You are a Presentation Architect. 

TASK:

1. Visually analyze the slide layout, colors, and charts.

2. Research specific data for: '{user_prompt}'.

3. Generate a JSON response to adapt the slide.

CRITICAL JSON RULES:

- Use valid JSON syntax.
- Escape double quotes inside strings (e.g., "The \\"Big 4\\" firms").
- Do not add trailing commas.
- Return ONLY the JSON object. No explanations, no markdown code blocks.

OUTPUT SCHEMA:

{{
   "replacements": [
      {{"old_text_snippet": "exact text from slide", "new_text": "adapted text"}}
   ],
   "charts": [
      {{
         "type": "bar/column/line", 
         "title": "Chart Title", 
         "data": {{
            "categories": ["2023", "2024"], 
            "series": [
               {{"name": "Series 1", "values": [10, 20], "color_hex": "#FF0000"}}
            ]
         }}
      }}
   ],
   "think_cell_replacements": true
}}"""

    # Step E: Prepare and send to Gemini
    step_start = time.time()
    logger.info("    [Step D] Preparing image and prompt for Gemini...")
    
    # WICHTIG: Wir öffnen das Bild, nutzen es und schließen es sicher
    image = PIL.Image.open(temp_png_path)
    
    try:
        image_size = f"{image.size[0]}x{image.size[1]}"
        logger.info(f"    [Step D] ✓ Image loaded: {image_size} pixels")
        
        # Model name - Gemini 2.5 Flash for better availability and speed
        model_name = "gemini-2.5-flash"
        
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
    
    # --- ROBUST JSON PARSING WITH REPAIR ATTEMPTS ---
    logger.info("    [Step D] Parsing JSON response...")
    response_text = response.text.strip()

    try:
        # 1. Clean string mit der Funktion (entfernt Markdown-Code-Blöcke)
        json_str = extract_json_from_text(response_text)
        
        # 2. Try to parse
        json_instructions = json.loads(json_str)
        
        # 3. Validate structure
        if "replacements" not in json_instructions and "charts" not in json_instructions:
            logger.warning("    [Step D] ⚠️ JSON valid but keys missing (replacements/charts)")
        
        logger.info(f"    [Step D] ✓ JSON parsed successfully ({len(str(json_instructions))} chars)")
        
        # --- NEW: Print JSON to Console ---
        print("\n" + "="*60)
        print("🟢 GEMINI AI RESPONSE (DEBUG VIEW)")
        print("="*60)
        json_preview = json.dumps(json_instructions, indent=2, ensure_ascii=False)
        # Truncate if too long for console
        if len(json_preview) > 1000:
            print(json_preview[:1000] + "... (truncated)")
        else:
            print(json_preview)
        print("="*60 + "\n")
        # ----------------------------------
        
        # --- DEBUG: Save JSON to inspect ---
        debug_json_path = pptx_path + ".debug.json"
        try:
            with open(debug_json_path, "w", encoding="utf-8") as f:
                json.dump(json_instructions, f, indent=2, ensure_ascii=False)
            logger.info(f"    [DEBUG] Full Gemini Response saved to: {debug_json_path}")
        except Exception as e:
            logger.warning(f"    [DEBUG] Could not save debug JSON: {e}")
        
        return json_instructions

    except json.JSONDecodeError as e:
        logger.error(f"    [Step D] ✗ JSON parsing failed: {e}")
        
        # Try simple JSON repair: remove trailing commas
        try:
            # Remove trailing commas before } or ]
            repaired = re.sub(r',(\s*[}\]])', r'\1', json_str)
            json_instructions = json.loads(repaired)
            logger.info(f"    [Step D] ✓ JSON repaired (removed trailing commas) and parsed successfully")
            
            # Print and save repaired version
            print("\n" + "="*60)
            print("🟢 GEMINI AI RESPONSE (REPAIRED)")
            print("="*60)
            json_preview = json.dumps(json_instructions, indent=2, ensure_ascii=False)
            if len(json_preview) > 1000:
                print(json_preview[:1000] + "... (truncated)")
            else:
                print(json_preview)
            print("="*60 + "\n")
            
            return json_instructions
        except:
            pass  # Repair failed, continue to error handling
        
        # Dump full error response to file for inspection
        err_file = pptx_path + ".error.txt"
        try:
            with open(err_file, "w", encoding="utf-8") as f:
                f.write("="*80 + "\n")
                f.write("FULL GEMINI RESPONSE (ERROR CASE)\n")
                f.write("="*80 + "\n\n")
                f.write("RAW RESPONSE:\n")
                f.write(response_text)
                f.write("\n\n" + "="*80 + "\n")
                f.write("EXTRACTED JSON STRING:\n")
                f.write(json_str if 'json_str' in locals() else "N/A")
                f.write("\n\n" + "="*80 + "\n")
                f.write(f"ERROR: {str(e)}\n")
            logger.error(f"    [Step D] Full error response saved to: {err_file}")
        except Exception as save_err:
            logger.error(f"    [Step D] Could not save error file: {save_err}")
        
        logger.error(f"    [Step D] Extracted text preview: {json_str[:500] if 'json_str' in locals() else 'N/A'}...")
        logger.error(f"    [Step D] Raw text preview: {response_text[:500]}...")
        
        # Fallback empty
        return {"replacements": [], "charts": []}
    
    except Exception as e:
        logger.error(f"    [Step D] ✗ General error parsing response: {e}")
        
        # Also save error response in general error case
        err_file = pptx_path + ".error.txt"
        try:
            with open(err_file, "w", encoding="utf-8") as f:
                f.write("="*80 + "\n")
                f.write("FULL GEMINI RESPONSE (GENERAL ERROR)\n")
                f.write("="*80 + "\n\n")
                f.write("RAW RESPONSE:\n")
                f.write(response_text)
                f.write("\n\n" + "="*80 + "\n")
                f.write(f"ERROR: {str(e)}\n")
            logger.error(f"    [Step D] Full error response saved to: {err_file}")
        except:
            pass
        
        return {"replacements": [], "charts": []}
