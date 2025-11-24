import os
import time
import logging
import aspose.slides as slides
from aspose.slides import IAutoShape, IGroupShape, ITable
import aspose.pydrawing as drawing

logger = logging.getLogger(__name__)


def replace_text_in_slide(slide, replacements):
    """
    Replaces text in slide while preserving formatting.
    Handles Shapes, Groups, and Tables.
    """
    logger.info(f"      → Processing {len(replacements)} text replacements...")
    replacement_count = 0

    # Helper: Normalize text for comparison (ignore extra spaces)
    def normalize(s):
        return ' '.join(s.split()).strip()

    def process_shape(shape):
        nonlocal replacement_count

        # 1. Handle Text Frames (AutoShapes)
        if hasattr(shape, "text_frame") and shape.text_frame:
            for paragraph in shape.text_frame.paragraphs:
                for portion in paragraph.portions:
                    original_text = portion.text
                    norm_original = normalize(original_text)

                    for replacement in replacements:
                        old = replacement.get('old_text_snippet', '')
                        new = replacement.get('new_text', '')
                        norm_old = normalize(old)

                        # Check: exact match OR normalized match OR substring match
                        if old and (old in original_text or norm_old in norm_original):
                            # Perform replacement on the original text to keep formatting
                            portion.text = original_text.replace(old, new)
                            if portion.text == original_text:  # Fallback regex-like replace if simple replace fails
                                portion.text = new

                            replacement_count += 1
                            logger.info(f"        ✓ Replaced: '{old[:20]}...'")

        # 2. Handle Groups (Recursion)
        if isinstance(shape, IGroupShape):
            for child_shape in shape.shapes:
                process_shape(child_shape)

        # 3. Handle Tables
        if isinstance(shape, ITable):
            for row in shape.rows:
                for cell in row:
                    if cell.text_frame:
                        for paragraph in cell.text_frame.paragraphs:
                            for portion in paragraph.portions:
                                original_text = portion.text
                                for replacement in replacements:
                                    old = replacement.get('old_text_snippet', '')
                                    new = replacement.get('new_text', '')
                                    if old and old in original_text:
                                        portion.text = original_text.replace(old, new)
                                        replacement_count += 1

    for shape in slide.shapes:
        process_shape(shape)

    logger.info(f"      ✓ Completed {replacement_count} text replacements")


def hex_to_argb(hex_color):
    """
    Converts hex color string (e.g., '#FF5733') to ARGB tuple.
    
    Args:
        hex_color: Hex color string with or without '#'
        
    Returns:
        tuple: (A, R, G, B) values
    """
    # Robust handling for None, empty strings, or invalid types
    if not hex_color or not isinstance(hex_color, str):
        logger.warning(f"Invalid hex_color provided: {hex_color}, using fallback black")
        return (255, 0, 0, 0)  # Fallback Schwarz
    
    hex_color = hex_color.lstrip('#').strip()
    
    # Check if empty after stripping
    if not hex_color:
        logger.warning("Empty hex_color after stripping, using fallback black")
        return (255, 0, 0, 0)
    
    try:
        if len(hex_color) == 6:
            r = int(hex_color[0:2], 16)
            g = int(hex_color[2:4], 16)
            b = int(hex_color[4:6], 16)
            return (255, r, g, b)
        elif len(hex_color) == 8:
            a = int(hex_color[0:2], 16)
            r = int(hex_color[2:4], 16)
            g = int(hex_color[4:6], 16)
            b = int(hex_color[6:8], 16)
            return (a, r, g, b)
        else:
            # Default to black if invalid length
            logger.warning(f"Invalid hex_color length: {len(hex_color)}, using fallback black")
            return (255, 0, 0, 0)
    except (ValueError, IndexError) as e:
        # Handle invalid hex characters
        logger.warning(f"Error parsing hex_color '{hex_color}': {e}, using fallback black")
        return (255, 0, 0, 0)


def replace_ole_with_chart(slide, shape, chart_data):
    """
    Replaces an OLE object (Think-Cell chart) with a native PowerPoint chart.
    
    Args:
        slide: Aspose.Slides ISlide object
        shape: The OLE object shape to replace
        chart_data: Dict with 'type', 'title', 'data' keys
    """
    chart_type_str = chart_data.get('type', 'column')
    chart_title = chart_data.get('title', 'Untitled')
    logger.info(f"      → Replacing OLE object with {chart_type_str} chart: '{chart_title}'...")
    
    # 1. Position & Size
    x = shape.x
    y = shape.y
    width = shape.width
    height = shape.height

    # Fallback for 0x0 size (Think-Cell specific issue)
    if width <= 0 or height <= 0:
        logger.warning("        ⚠️ OLE Object has 0x0 size. Using default fallback size.")
        slide_width = slide.presentation.slide_size.size.width
        slide_height = slide.presentation.slide_size.size.height
        width = slide_width * 0.8
        height = slide_height * 0.6
        x = (slide_width - width) / 2
        y = (slide_height - height) / 2

    logger.info(f"        Final Chart Rect: X={x:.0f}, Y={y:.0f}, W={width:.0f}, H={height:.0f}")
    
    # 2. Chart Type
    c_type_lower = str(chart_type_str).lower()
    if 'bar' in c_type_lower:
        chart_type = slides.charts.ChartType.CLUSTERED_BAR
    elif 'line' in c_type_lower:
        chart_type = slides.charts.ChartType.LINE
    else:
        chart_type = slides.charts.ChartType.CLUSTERED_COLUMN
    
    # 3. Remove OLE
    slide.shapes.remove(shape)

    # 4. Create Native Chart
    chart = slide.shapes.add_chart(chart_type, x, y, width, height)

    # 5. Fill Data
    chart_data_obj = chart.chart_data
    workbook = chart_data_obj.chart_data_workbook

    chart_data_obj.series.clear()
    chart_data_obj.categories.clear()

    categories = chart_data.get('data', {}).get('categories', [])
    series_list = chart_data.get('data', {}).get('series', [])

    for i, category in enumerate(categories):
        cell = workbook.get_cell(0, 0, i + 1, str(category))
        chart_data_obj.categories.add(cell)

    for series_idx, series_data in enumerate(series_list):
        series_name = series_data.get('name', f'Series {series_idx + 1}')
        values = series_data.get('values', [])
        color_hex = series_data.get('color_hex', '#000000')

        series_name_cell = workbook.get_cell(0, series_idx + 1, 0, str(series_name))
        series = chart_data_obj.series.add(series_name_cell, chart_type)

        for cat_idx, value in enumerate(values):
            if cat_idx < len(categories):
                try:
                    val_float = float(value)
                except:
                    val_float = 0.0

                value_cell = workbook.get_cell(0, series_idx + 1, cat_idx + 1, val_float)

                if chart_type == slides.charts.ChartType.LINE:
                    series.data_points.add_data_point_for_line_series(value_cell)
                else:
                    series.data_points.add_data_point_for_bar_series(value_cell)

        # Color
        series_format = series.format
        series_format.fill.fill_type = slides.FillType.SOLID
        a, r, g, b = hex_to_argb(color_hex)
        series_format.fill.solid_fill_color.color = drawing.Color.from_argb(a, r, g, b)

    # --- FIX: Chart Title Handling ---
    if chart_data.get('title'):
        chart.has_title = True
        chart.chart_title.add_text_frame_for_overriding(chart_data.get('title'))


def process_slide(pptx_path, output_path, json_instructions):
    """
    Main orchestrator that applies all replacements and chart replacements.
    
    Args:
        pptx_path: Path to input PPTX file
        output_path: Path to save modified PPTX file
        json_instructions: Dict with 'replacements', 'charts', 'think_cell_replacements'
    """
    # Load presentation
    step_start = time.time()
    logger.info("      [Step 1] Loading presentation...")
    pres = slides.Presentation(pptx_path)
    slide = pres.slides[0]
    logger.info(f"      [Step 1] ✓ Presentation loaded ({len(pres.slides)} slide(s))")
    
    # Apply text replacements
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
    
    # Save the modified presentation
    logger.info("      [Step 2] Saving modified presentation...")
    pres.save(output_path, slides.export.SaveFormat.PPTX)
    step_time = time.time() - step_start
    file_size = os.path.getsize(output_path)
    logger.info(f"      [Step 2] ✓ Presentation saved in {step_time:.2f}s ({file_size:,} bytes)")

