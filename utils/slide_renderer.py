import os
import time
import logging
import re
import aspose.slides as slides
from aspose.slides import IAutoShape, IGroupShape, ITable
import aspose.pydrawing as drawing

logger = logging.getLogger(__name__)

def load_aspose_license_if_available():
    """Versucht, eine Aspose.Slides Lizenz-Datei zu laden."""
    license_paths = [
        "Aspose.Slides.lic",
        "aspose.lic",
        os.path.join(os.path.dirname(__file__), "Aspose.Slides.lic"),
        os.path.join(os.path.dirname(__file__), "aspose.lic"),
    ]
    
    for lic_path in license_paths:
        if os.path.exists(lic_path):
            try:
                license_obj = slides.License()
                license_obj.set_license(lic_path)
                logger.info(f"      ✓ Aspose license loaded from: {lic_path}")
                return True
            except Exception as e:
                logger.warning(f"      ⚠️ Could not load license from {lic_path}: {e}")
    
    logger.warning("      ⚠️ No Aspose license found - running in evaluation mode (text may be truncated)")
    return False

def normalize_string(s):
    """
    Normalisiert einen String aggressiv für den Vergleich.
    Entfernt Bulletpoints (», •, -), Satzzeichen und macht alles lowercase.
    Dies löst das Problem, dass '» Aufgrund' nicht mit 'Aufgrund' matcht.
    """
    if not s:
        return ""
    # Entferne alles was kein Buchstabe oder Zahl ist (behält Spaces)
    s_clean = re.sub(r'[^\w\s]', '', str(s))
    return " ".join(s_clean.split()).lower().strip()

def replace_text_in_slide(slide, replacements):
    """Ersetzt Text auf einer Slide basierend auf Gemini's Vorschlägen."""
    logger.info(f"      → Processing {len(replacements)} text replacements...")
    replacement_count = 0

    clean_replacements = []
    for r in replacements:
        old = r.get('old_text_snippet', '')
        new = r.get('new_text', '')
        if old and new:
            clean_old = normalize_string(old)
            # Wir akzeptieren auch kürzere Snippets (ab 5 chars), da die Normalisierung vieles bereinigt
            if len(clean_old) > 5: 
                clean_replacements.append((clean_old, new, old))
    
    logger.info(f"      → Prepared {len(clean_replacements)} valid replacements for matching")
    
    def get_visible_text_prefix(text_frame, max_chars=150):
        """Gets the visible text prefix from a text frame (safe for evaluation mode)."""
        if not text_frame or not text_frame.paragraphs:
            return ""
        
        first_para = text_frame.paragraphs[0] if text_frame.paragraphs and text_frame.paragraphs.count > 0 else None
        if not first_para:
            return ""
        
        visible_text = ""
        try:
            if hasattr(first_para, 'text'):
                visible_text = first_para.text
                # Filter Evaluation Messages aggressive
                if "Evaluation" in visible_text or "Created with" in visible_text:
                     # Versuche echten Content zu finden (oft vor der Message oder danach)
                     parts = visible_text.split("Evaluation")
                     if len(parts) > 0 and len(parts[0].strip()) > 3:
                         visible_text = parts[0]
            
            if not visible_text and first_para.portions and first_para.portions.count > 0:
                portion_texts = []
                for i in range(min(5, first_para.portions.count)):
                    if hasattr(first_para.portions[i], 'text'):
                        portion_texts.append(first_para.portions[i].text)
                visible_text = "".join(portion_texts)
            
            if len(visible_text) > max_chars:
                visible_text = visible_text[:max_chars]
                
        except Exception:
            pass
        
        return visible_text.strip()
    
    def process_text_frame(text_frame):
        nonlocal replacement_count
        if not text_frame or not text_frame.paragraphs:
            return
        
        visible_prefix = get_visible_text_prefix(text_frame, max_chars=200)
        clean_visible = normalize_string(visible_prefix)
        
        if len(clean_visible) < 3:
            return
        
        for clean_old, new_text, orig_old in clean_replacements:
            is_match = False
            match_type = None
            
            # 1. Direct Substring Match (Aggressive Normalization macht das möglich)
            # "aufgrund" in "aufgrund im internationalen..."
            if clean_old in clean_visible or clean_visible in clean_old:
                is_match = True
                match_type = "direct_substring"
            
            # 2. Keyword Set Match (Best for "Evaluation Mode" fragments)
            # Wenn > 60% der Wörter aus dem Suchtext im sichtbaren Text vorkommen
            if not is_match:
                old_words = set([w for w in clean_old.split() if len(w) > 3]) # Nur signifikante Wörter
                vis_words = set([w for w in clean_visible.split() if len(w) > 3])
                
                if len(old_words) > 0:
                    common = old_words.intersection(vis_words)
                    # Wenn wir wenig sehen (Evaluation mode), reicht 1 starkes Keyword
                    if len(common) >= 1 and len(clean_visible) < 30: 
                        is_match = True
                        match_type = "keyword_match_short"
                    # Sonst brauchen wir mehr Overlap
                    elif len(common) / len(old_words) > 0.5:
                        is_match = True
                        match_type = "keyword_match_strong"
            
            if is_match:
                logger.info(f"        ✓ Match found ({match_type})!\n          Search: '{orig_old[:40]}...'\n          Found in: '{visible_prefix[:40]}...'")
                
                # Replacement Logic
                if text_frame.paragraphs.count > 0:
                    # Wir löschen erst alle Paragraphen außer dem ersten
                    while text_frame.paragraphs.count > 1:
                        text_frame.paragraphs.remove_at(1)
                    
                    # Dann setzen wir den Text im ersten Paragraphen
                    para = text_frame.paragraphs[0]
                    para.portions.clear() # Alle alten Formatierungen weg
                    para.text = new_text  # Neuen Text setzen
                    
                    replacement_count += 1
                    return # Done with this frame
    
    def process_shape(shape):
        if hasattr(shape, "text_frame") and shape.text_frame:
            process_text_frame(shape.text_frame)
        if isinstance(shape, IGroupShape):
            for child in shape.shapes:
                process_shape(child)
        if isinstance(shape, ITable):
            for row in shape.rows:
                for cell in row:
                    if cell.text_frame:
                        process_text_frame(cell.text_frame)
    
    for shape in slide.shapes:
        process_shape(shape)
    
    logger.info(f"      ✓ Completed {replacement_count} text replacements")

def hex_to_argb(hex_color):
    """Converts hex color string to ARGB tuple."""
    if not hex_color or not isinstance(hex_color, str):
        return (255, 0, 0, 0)
    hex_color = hex_color.lstrip('#').strip()
    try:
        if len(hex_color) == 6:
            return (255, int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16))
        return (255, 0, 0, 0)
    except:
        return (255, 0, 0, 0)

def replace_ole_with_chart(slide, shape, chart_data):
    """Replaces a shape with a new native PowerPoint chart."""
    chart_type_str = chart_data.get('type', 'column')
    chart_title = chart_data.get('title', 'Untitled')
    
    logger.info(f"      → Creating new {chart_type_str} chart: '{chart_title}'...")
    x, y, width, height = shape.x, shape.y, shape.width, shape.height
    
    # Remove old shape
    slide.shapes.remove(shape)
    # Map Type
    c_type_lower = str(chart_type_str).lower()
    if 'pie' in c_type_lower: chart_type = slides.charts.ChartType.PIE
    elif 'bar' in c_type_lower: chart_type = slides.charts.ChartType.CLUSTERED_BAR
    elif 'line' in c_type_lower: chart_type = slides.charts.ChartType.LINE
    else: chart_type = slides.charts.ChartType.CLUSTERED_COLUMN
    
    # Create Chart
    chart = slide.shapes.add_chart(chart_type, x, y, width, height)
    wb = chart.chart_data.chart_data_workbook
    chart.chart_data.series.clear()
    chart.chart_data.categories.clear()
    
    # Data
    categories = chart_data.get('data', {}).get('categories', [])
    series_list = chart_data.get('data', {}).get('series', [])
    # 1. Categories
    for i, cat in enumerate(categories):
        chart.chart_data.categories.add(wb.get_cell(0, i + 1, 0, str(cat)))
    
    # 2. Series & Values
    for series_idx, s_data in enumerate(series_list):
        s_name = s_data.get('name', f'Series {series_idx+1}')
        s_vals = s_data.get('values', [])
        s_color = s_data.get('color_hex', '#000000')
        
        # Add Series
        series = chart.chart_data.series.add(wb.get_cell(0, 0, series_idx + 1, str(s_name)), chart_type)
        
        # Add DataPoints
        for cat_idx, val in enumerate(s_vals):
            if cat_idx < len(categories):
                try: v = float(val) 
                except: v = 0.0
                cell = wb.get_cell(0, cat_idx + 1, series_idx + 1, v)
                
                if chart_type == slides.charts.ChartType.PIE:
                    series.data_points.add_data_point_for_pie_series(cell)
                elif chart_type == slides.charts.ChartType.LINE:
                    series.data_points.add_data_point_for_line_series(cell)
                else:
                    series.data_points.add_data_point_for_bar_series(cell)
        # Color
        try:
            series.format.fill.fill_type = slides.FillType.SOLID
            a, r, g, b = hex_to_argb(s_color)
            series.format.fill.solid_fill_color.color = drawing.Color.from_argb(a, r, g, b)
        except: pass
    
    # Title
    if chart_data.get('title'):
        chart.has_title = True
        chart.chart_title.add_text_frame_for_overriding(chart_data.get('title'))
        
    # Style
    try:
        if hasattr(slides.charts, 'ChartStyle'):
            chart.style = slides.charts.ChartStyle.STYLE_11
    except: pass

def update_native_chart_data(chart, chart_data):
    """Updates data of an EXISTING native chart."""
    logger.info(f"      → Updating existing native chart: '{chart_data.get('title')}'")
    
    wb = chart.chart_data.chart_data_workbook
    chart.chart_data.series.clear()
    chart.chart_data.categories.clear()
    
    # Categories
    cats = chart_data.get('data', {}).get('categories', [])
    for i, cat in enumerate(cats):
        chart.chart_data.categories.add(wb.get_cell(0, i + 1, 0, str(cat)))
        
    # Series
    series_list = chart_data.get('data', {}).get('series', [])
    for i, s_data in enumerate(series_list):
        s_name = s_data.get('name', f'Series {i+1}')
        vals = s_data.get('values', [])
        
        series = chart.chart_data.series.add(wb.get_cell(0, 0, i + 1, str(s_name)), chart.type)
        
        for j, val in enumerate(vals):
            if j < len(cats):
                try: v = float(val) 
                except: v = 0.0
                cell = wb.get_cell(0, j + 1, i + 1, v)
                
                if chart.type == slides.charts.ChartType.PIE:
                    series.data_points.add_data_point_for_pie_series(cell)
                elif chart.type == slides.charts.ChartType.LINE:
                    series.data_points.add_data_point_for_line_series(cell)
                else:
                    series.data_points.add_data_point_for_bar_series(cell)
    
    if chart_data.get('title') and chart.has_title:
        chart.chart_title.add_text_frame_for_overriding(chart_data.get('title'))

def apply_global_substitutions(slide, substitutions):
    """Global replace safety net."""
    if not substitutions: return 0
    count = 0
    
    def check_replace(tf):
        nonlocal count
        if not tf: return
        for p in tf.paragraphs:
            for port in p.portions:
                if hasattr(port, 'text'):
                    txt = port.text
                    for old, new in substitutions.items():
                        if old in txt:
                            port.text = txt.replace(old, new)
                            count += 1
    
    def scan_shape(shape):
        if hasattr(shape, "text_frame"): check_replace(shape.text_frame)
        if isinstance(shape, IGroupShape): 
            for child in shape.shapes: scan_shape(child)
        if isinstance(shape, ITable):
            for r in shape.rows: 
                for c in r: 
                    if c.text_frame: check_replace(c.text_frame)
    
    for s in slide.shapes: scan_shape(s)
    logger.info(f"      ✓ Global safety net replaced {count} occurrences")
    return count

def process_slide(pptx_path, output_path, json_instructions):
    step_start = time.time()
    load_aspose_license_if_available()
    
    logger.info("      [Step 1] Loading presentation...")
    pres = slides.Presentation(pptx_path)
    slide = pres.slides[0]
    # Text
    replacements = json_instructions.get('replacements', [])
    if replacements:
        replace_text_in_slide(slide, replacements)
    # Charts
    charts = json_instructions.get('charts', [])
    if charts:
        logger.info(f"      → Processing {len(charts)} chart replacements...")
        candidates = []
        
        # FIX: Robust Chart Detection - Use try/except to safely detect charts
        for shape in slide.shapes:
            # 1. OLE Objects (Think-Cell) - Check first
            if hasattr(shape, 'ole_format') or isinstance(shape, slides.OleObjectFrame):
                candidates.append(shape)
            # 2. Native Charts - Try to access chart_data (safest method)
            else:
                try:
                    # If shape has chart_data, it's a native chart
                    if hasattr(shape, 'chart_data'):
                        _ = shape.chart_data  # Try to access it
                        candidates.append(shape)
                except:
                    pass  # Not a chart, skip
        # Sort: Top->Bottom, Left->Right
        candidates.sort(key=lambda s: (int(s.y / 50) * 1000 + s.x))
        
        logger.info(f"      → Found {len(candidates)} chart candidate(s)")
        for i, shape in enumerate(candidates):
            if i < len(charts):
                chart_data = charts[i]
                # Check if Native Chart (has chart_data and can access it)
                is_native = False
                try:
                    if hasattr(shape, 'chart_data'):
                        _ = shape.chart_data
                        is_native = True
                except:
                    pass
                
                if is_native:
                    try:
                        update_native_chart_data(shape, chart_data)
                    except Exception as e:
                        logger.error(f"Failed update, replacing: {e}")
                        replace_ole_with_chart(slide, shape, chart_data)
                else:
                    replace_ole_with_chart(slide, shape, chart_data)
    
    # NO GLOBAL SUBSTITUTIONS - Only individual text replacements as specified
    # Each text is processed individually through replace_text_in_slide()
    
    pres.save(output_path, slides.export.SaveFormat.PPTX)
    logger.info(f"      [Step 2] ✓ Saved ({os.path.getsize(output_path)} bytes)")
