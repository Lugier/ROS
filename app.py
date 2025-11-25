import os
import sys
import time
import tempfile
import logging
import io
from datetime import datetime, timedelta
from flask import Flask, render_template, request, send_file, jsonify
from utils.vision_analyzer import analyze_slide_and_research
from utils.slide_renderer import process_slide

# Configure detailed logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S',
    stream=sys.stdout
)
logger = logging.getLogger(__name__)

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size
app.config['UPLOAD_FOLDER'] = tempfile.gettempdir()


@app.route('/')
def index():
    """Render the upload form."""
    return render_template('index.html')


@app.route('/process', methods=['POST'])
def process_presentation():
    """
    Process the uploaded PPTX file:
    1. Save uploaded file
    2. Analyze with Gemini Vision + Research
    3. Apply replacements and chart replacements
    4. Return modified file for download
    """
    start_time = time.time()
    logger.info("=" * 80)
    logger.info("NEW REQUEST: Presentation Processing Started")
    logger.info("=" * 80)
    
    try:
        # Check if file was uploaded
        if 'pptx_file' not in request.files:
            logger.error("ERROR: No file uploaded")
            return jsonify({'error': 'No file uploaded'}), 400
        
        file = request.files['pptx_file']
        if file.filename == '':
            logger.error("ERROR: No file selected")
            return jsonify({'error': 'No file selected'}), 400
        
        if not file.filename.lower().endswith('.pptx'):
            logger.error(f"ERROR: Invalid file type - {file.filename}")
            return jsonify({'error': 'File must be a .pptx file'}), 400
        
        # Get user prompt
        user_prompt = request.form.get('prompt', '')
        if not user_prompt:
            logger.error("ERROR: No prompt provided")
            return jsonify({'error': 'Prompt is required'}), 400
        
        logger.info(f"File: {file.filename}")
        logger.info(f"Prompt: {user_prompt}")
        logger.info(f"File size: {len(file.read())} bytes")
        file.seek(0)  # Reset file pointer
        
        # Phase 0: Save uploaded file
        phase_start = time.time()
        logger.info("-" * 80)
        logger.info("PHASE 0/3: Saving uploaded file...")
        input_temp = tempfile.NamedTemporaryFile(delete=False, suffix='.pptx')
        input_path = input_temp.name
        file.save(input_path)
        input_temp.close()
        phase_time = time.time() - phase_start
        logger.info(f"✓ File saved to: {input_path} ({phase_time:.2f}s)")
        
        # Create output temporary file
        output_temp = tempfile.NamedTemporaryFile(delete=False, suffix='.pptx')
        output_path = output_temp.name
        output_temp.close()
        
        # Phase 1: Analyze slide and research content
        phase_start = time.time()
        logger.info("-" * 80)
        logger.info("PHASE 1/3: Vision Analysis & Research (Gemini 2.5 Flash)")
        logger.info("  → Loading PPTX and rendering slide to PNG...")
        logger.info("  → Sending to Gemini 2.5 Flash with Google Search...")
        logger.info("  → This may take 30-90 seconds depending on research complexity...")
        
        elapsed = time.time() - start_time
        estimated_total = 120  # Initial estimate in seconds
        eta = datetime.now() + timedelta(seconds=estimated_total - elapsed)
        logger.info(f"  → Estimated completion: {eta.strftime('%H:%M:%S')} (ETA: ~{estimated_total - elapsed:.0f}s)")
        
        json_instructions = analyze_slide_and_research(input_path, user_prompt)
        
        phase_time = time.time() - phase_start
        logger.info(f"✓ Phase 1 completed in {phase_time:.2f}s")
        logger.info(f"  → Found {len(json_instructions.get('replacements', []))} text replacements")
        logger.info(f"  → Found {len(json_instructions.get('charts', []))} charts to process")
        logger.info(f"  → Think-Cell replacements: {json_instructions.get('think_cell_replacements', False)}")
        
        # Phase 2: Process slide with replacements
        phase_start = time.time()
        logger.info("-" * 80)
        logger.info("PHASE 2/3: Applying replacements and chart modifications...")
        logger.info("  → Replacing text while preserving formatting...")
        logger.info("  → Replacing OLE objects with native charts...")
        
        elapsed = time.time() - start_time
        estimated_remaining = 10  # Estimate for processing
        eta = datetime.now() + timedelta(seconds=estimated_remaining)
        logger.info(f"  → Estimated completion: {eta.strftime('%H:%M:%S')} (ETA: ~{estimated_remaining:.0f}s)")
        
        process_slide(input_path, output_path, json_instructions)
        
        phase_time = time.time() - phase_start
        logger.info(f"✓ Phase 2 completed in {phase_time:.2f}s")
        
        # Phase 3: Load to RAM and Cleanup
        logger.info("-" * 80)
        logger.info("PHASE 3/3: Finalizing (Memory Load Strategy)...")

        # 1. Load into Memory
        with open(output_path, 'rb') as f:
            file_data = io.BytesIO(f.read())

        # 2. Delete files from disk NOW (safe because they are closed)
        try:
            if os.path.exists(input_path):
                os.remove(input_path)
            if os.path.exists(output_path):
                os.remove(output_path)
            logger.info(f"✓ Cleaned up temporary files successfully")
        except Exception as e:
            logger.warning(f"⚠️ Warning during cleanup: {e}")

        total_time = time.time() - start_time
        logger.info(f"✓ SUCCESS: Processing completed in {total_time:.2f}s")
        logger.info("=" * 80)

        # 3. Send from Memory
        file_data.seek(0)
        return send_file(
            file_data,
            as_attachment=True,
            download_name='adapted_presentation.pptx',
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
        )
        
    except Exception as e:
        total_time = time.time() - start_time
        logger.error("=" * 80)
        logger.error(f"✗ ERROR after {total_time:.2f}s: {str(e)}", exc_info=True)
        logger.error("=" * 80)
        
        # Fallback cleanup im Fehlerfall
        try:
            if os.path.exists(output_path):
                os.remove(output_path)
                logger.info(f"✓ Cleaned up output file after error: {output_path}")
        except Exception as cleanup_error:
            logger.error(f"Error during cleanup: {cleanup_error}")
        
        return jsonify({'error': f'Processing failed: {str(e)}'}), 500

if __name__ == '__main__':
    # Debug=True für Debugger, aber use_reloader=False damit langlaufende Requests nicht abgebrochen werden
    app.run(debug=True, use_reloader=False, host='localhost', port=5000)

