import os
import time
import logging
import aspose.slides as slides
from aspose.slides import IAutoShape, IGroupShape, ITable
import aspose.pydrawing as drawing

logger = logging.getLogger(__name__)


def normalize_string(s):
    """Entfernt Whitespace und macht lowercase für besseren Vergleich."""
    if not s:
        return ""
    return " ".join(str(s).split()).lower().strip()


def replace_text_in_slide(slide, replacements):
    """
    Replaces text using tolerant fuzzy matching.
    """
    logger.info(f"      → Processing {len(replacements)} text replacements...")
    replacement_count = 0
    
    # Debug: Was soll ersetzt werden?
    clean_replacements = []
    for r in replacements:
        old = r.get('old_text_snippet', '')
        new = r.get('new_text', '')
        if old and new:
            clean_old = normalize_string(old)
            # Wir speichern nur Snippets die lang genug sind, um False Positives zu vermeiden
            if len(clean_old) > 10: 
                clean_replacements.append((clean_old, new, old))  # Speichern auch das Original für Debug
    
    def process_text_frame(text_frame):
        nonlocal replacement_count
        if not text_frame or not text_frame.paragraphs:
            return
        for paragraph in text_frame.paragraphs:
            full_text = paragraph.text
            clean_full_text = normalize_string(full_text)
            
            # Debug: Logge gefundene Texte (nur wenn lang genug)
            if len(clean_full_text) >= 10:
                logger.debug(f"        [DEBUG] Found text in slide: '{full_text[:50]}...'")
            
            # Wenn der Paragraph leer ist, überspringen
            if len(clean_full_text) < 5: 
                continue
            
            # Prüfen gegen alle Ersetzungen
            for clean_old, new_text, orig_old in clean_replacements:
                
                # Match Logic:
                # 1. Exakter (normalisierter) Match
                # 2. Substring Match (wenn der Suchtext im Absatz vorkommt)
                # 3. Similarity Match (Token Overlap - für harte Fälle)
                
                is_match = False
                
                if clean_old in clean_full_text:
                    is_match = True
                else:
                    # Token Match: Wenn > 80% der Wörter vorkommen (in beliebiger Reihenfolge)
                    old_tokens = set(clean_old.split())
                    para_tokens = set(clean_full_text.split())
                    if len(old_tokens) > 0:
                        common = old_tokens.intersection(para_tokens)
                        overlap = len(common) / len(old_tokens)
                        if overlap > 0.8: 
                            is_match = True
                
                if is_match:
                    logger.info(f"        ✓ Match found!\n          Search: '{orig_old[:30]}...'\n          Found in: '{full_text[:30]}...'")
                    
                    # Replace Logic
                    if len(paragraph.portions) > 0:
                        paragraph.portions[0].text = new_text
                        # Remove others
                        while len(paragraph.portions) > 1:
                            paragraph.portions.remove_at(1)
                            
                        replacement_count += 1
                        return  # Nur eine Ersetzung pro Absatz, break loop
    
    def process_shape(shape):
        # 1. Text Frames
        if hasattr(shape, "text_frame"):
            process_text_frame(shape.text_frame)
        # 2. Groups
        if isinstance(shape, IGroupShape):
            for child in shape.shapes:
                process_shape(child)
        # 3. Tables
        if isinstance(shape, ITable):
            for row in shape.rows:
                for cell in row:
                    if cell.text_frame:
                        process_text_frame(cell.text_frame)
    
    # Iterate
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
        elif len(hex_color) == 8:
            a = int(hex_color[0:2], 16)
            r = int(hex_color[2:4], 16)
            g = int(hex_color[4:6], 16)
            b = int(hex_color[6:8], 16)
            return (a, r, g, b)
        return (255, 0, 0, 0)
    except:
        return (255, 0, 0, 0)


def replace_ole_with_chart(slide, shape, chart_data):
    """Replaces an OLE object with a native PowerPoint chart."""
    chart_type_str = chart_data.get('type', 'column')
    chart_title = chart_data.get('title', 'Untitled')
    logger.info(f"      → Replacing OLE object with {chart_type_str} chart: '{chart_title}'...")

    # 1. Geometrie
    x, y, width, height = shape.x, shape.y, shape.width, shape.height
    if width <= 0 or height <= 0:
        # Fallback Size
        slide_width = slide.presentation.slide_size.size.width
        slide_height = slide.presentation.slide_size.size.height
        width = slide_width * 0.8
        height = slide_height * 0.6
        x = (slide_width - width) / 2
        y = (slide_height - height) / 2
        logger.info(f"        ⚠️ Using fallback size: {width:.0f}x{height:.0f}")

    logger.info(f"        Final Chart Rect: X={x:.0f}, Y={y:.0f}, W={width:.0f}, H={height:.0f}")

    # 2. Typ
    c_type_lower = str(chart_type_str).lower()
    if 'bar' in c_type_lower:
        chart_type = slides.charts.ChartType.CLUSTERED_BAR
    elif 'line' in c_type_lower:
        chart_type = slides.charts.ChartType.LINE
    else:
        chart_type = slides.charts.ChartType.CLUSTERED_COLUMN

    # 3. Delete & Create
    slide.shapes.remove(shape)
    chart = slide.shapes.add_chart(chart_type, x, y, width, height)

    # 4. Data Fill
    chart_data_obj = chart.chart_data
    workbook = chart_data_obj.chart_data_workbook

    chart_data_obj.series.clear()
    chart_data_obj.categories.clear()

    # Debug Data
    categories = chart_data.get('data', {}).get('categories', [])
    series_list = chart_data.get('data', {}).get('series', [])
    logger.info(f"        Data: {len(categories)} categories, {len(series_list)} series")

    if not categories or not series_list:
        logger.warning("        ⚠️ No categories or series data provided!")

    # Aspose Chart Data Structure:
    # Excel-like layout:
    #      | Col 0      | Col 1 (Ser1) | Col 2 (Ser2)
    # Row 0| (empty)    | Series 1     | Series 2
    # Row 1| Category 1 | Value 1.1     | Value 2.1
    # Row 2| Category 2 | Value 1.2     | Value 2.2
    #
    # Categories: Row 1..N, Col 0
    # Series Names: Row 0, Col 1..N
    # Values: Row 1..N, Col 1..N

    # Categories (Row 1..N, Col 0)
    for i, cat in enumerate(categories):
        # Row = i+1 (start at 1), Col = 0
        cell = workbook.get_cell(0, i + 1, 0, str(cat))
        chart_data_obj.categories.add(cell)
        logger.debug(f"        Category {i}: '{cat}' at (row {i+1}, col 0)")

    # Series
    for series_idx, s_data in enumerate(series_list):
        s_name = s_data.get('name', f'Series {series_idx + 1}')
        s_vals = s_data.get('values', [])
        s_color = s_data.get('color_hex', '#000000')

        # Series Name: Row 0, Col = series_idx + 1
        series_name_cell = workbook.get_cell(0, 0, series_idx + 1, str(s_name))
        series = chart_data_obj.series.add(series_name_cell, chart_type)
        logger.debug(f"        Series {series_idx}: '{s_name}' at (row 0, col {series_idx + 1})")

        # Add Values: Row = cat_idx + 1, Col = series_idx + 1
        for cat_idx, val in enumerate(s_vals):
            if cat_idx < len(categories):
                try:
                    val_float = float(val)
                except (ValueError, TypeError):
                    val_float = 0.0
                    logger.warning(f"        ⚠️ Invalid value '{val}' for category {cat_idx}, using 0.0")

                # Value Cell: Row = cat_idx + 1, Col = series_idx + 1
                data_cell = workbook.get_cell(0, cat_idx + 1, series_idx + 1, val_float)
                logger.debug(f"        Value: {val_float} at (row {cat_idx+1}, col {series_idx+1})")

                # Add Point based on chart type
                if chart_type == slides.charts.ChartType.LINE:
                    series.data_points.add_data_point_for_line_series(data_cell)
                else:
                    series.data_points.add_data_point_for_bar_series(data_cell)

        # Color
        series.format.fill.fill_type = slides.FillType.SOLID
        a, r, g, b = hex_to_argb(s_color)
        series.format.fill.solid_fill_color.color = drawing.Color.from_argb(a, r, g, b)
        logger.debug(f"        Color applied: {s_color} -> ARGB({a}, {r}, {g}, {b})")

    # Title
    if chart_data.get('title'):
        chart.has_title = True
        chart.chart_title.add_text_frame_for_overriding(chart_data.get('title'))
        logger.info(f"        ✓ Chart title set: '{chart_data.get('title')}'")

    # FIX: Chart Style Override (damit es nicht "Default Blau" ist)
    try:
        # Versuche einen moderneren Style zu setzen (falls verfügbar)
        if hasattr(slides.charts, 'ChartStyle'):
            # Style 11 ist meist ein modernerer, farbenfroherer Style
            chart.style = slides.charts.ChartStyle.STYLE_11
            logger.debug("        Chart style set to STYLE_11")
    except (AttributeError, Exception) as e:
        # Falls ChartStyle nicht verfügbar ist, ignorieren wir es
        logger.debug(f"        Could not set chart style (ignoring): {e}")

    logger.info(f"      ✓ Chart created successfully with {len(series_list)} series")


def process_slide(pptx_path, output_path, json_instructions):
    """Main orchestrator that applies all replacements and chart replacements."""
    step_start = time.time()
    logger.info("      [Step 1] Loading presentation...")
    pres = slides.Presentation(pptx_path)
    slide = pres.slides[0]
    logger.info(f"      [Step 1] ✓ Presentation loaded ({len(pres.slides)} slide(s))")

    # Text
    replacements = json_instructions.get('replacements', [])
    if replacements:
        replace_text_in_slide(slide, replacements)
    else:
        logger.info("      → No text replacements to apply")

    # Charts
    charts = json_instructions.get('charts', [])
    if charts:
        logger.info(f"      → Processing {len(charts)} chart replacements...")
        ole_shapes = []
        for shape in slide.shapes:
            if hasattr(slides, 'OleObjectFrame') and isinstance(shape, slides.OleObjectFrame):
                ole_shapes.append(shape)
            elif hasattr(shape, 'ole_format'):
                ole_shapes.append(shape)

        logger.info(f"      → Found {len(ole_shapes)} OLE object(s)")

        for i, ole_shape in enumerate(ole_shapes):
            if i < len(charts):
                replace_ole_with_chart(slide, ole_shape, charts[i])
    else:
        logger.info("      → No chart replacements to apply")

    logger.info("      [Step 2] Saving modified presentation...")
    pres.save(output_path, slides.export.SaveFormat.PPTX)
    step_time = time.time() - step_start
    file_size = os.path.getsize(output_path)
    logger.info(f"      [Step 2] ✓ Presentation saved in {step_time:.2f}s ({file_size:,} bytes)")
