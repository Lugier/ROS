import os
import time
import logging
import aspose.slides as slides
from aspose.slides import IAutoShape, IGroupShape, ITable
import aspose.pydrawing as drawing

logger = logging.getLogger(__name__)


def load_aspose_license_if_available():
    """
    Versucht, eine Aspose.Slides Lizenz-Datei zu laden, falls vorhanden.
    
    WICHTIG: Mit einer gültigen Lizenz würde die Text-Truncation nicht auftreten.
    Diese Funktion sucht nach einer .lic Datei in mehreren Standard-Pfaden.
    
    Falls keine Lizenz gefunden wird, läuft Aspose im Evaluationsmodus:
    - Text wird nach ~10-20 Zeichen abgeschnitten
    - Watermark wird in generierte Dateien eingefügt
    - Aber: Die Funktion funktioniert trotzdem (mit unseren Workarounds)
    
    Returns:
        bool: True wenn Lizenz geladen wurde, False sonst
    """
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
    Normalisiert einen String für besseren Vergleich.
    
    WARUM: Gemini's Text-Vorschläge können leicht anders formatiert sein als der Text auf der Slide:
    - Unterschiedliche Whitespace (z.B. "Markt Volumen" vs "Marktvolumen")
    - Groß-/Kleinschreibung (z.B. "Markt" vs "markt")
    - Zeilenumbrüche vs. Spaces
    
    Diese Funktion macht beide Strings vergleichbar, indem sie:
    1. Alle Whitespace (auch non-breaking spaces) durch normale Spaces ersetzt
    2. Alles in lowercase konvertiert
    3. Führende/nachfolgende Spaces entfernt
    
    Args:
        s: String zu normalisieren
        
    Returns:
        str: Normalisierter String (lowercase, single spaces)
    """
    if not s:
        return ""
    return " ".join(str(s).split()).lower().strip()


def replace_text_in_slide(slide, replacements):
    """
    Ersetzt Text auf einer Slide basierend auf Gemini's Vorschlägen.
    
    WICHTIG: Diese Funktion ist speziell für Aspose.Slides Evaluationsversion optimiert.
    Da die kostenlose Version Text nach ~10-20 Zeichen abschneidet, verwenden wir:
    
    1. PRÄFIX-EXTRAKTION: Nur die ersten ~150 Zeichen werden extrahiert (vor Truncation)
    2. PRÄFIX-MATCHING: Erste 3-10 Wörter der Gemini-Vorschläge werden gegen sichtbaren Präfix gematcht
    3. TOKEN-OVERLAP: Falls Präfix-Match fehlschlägt, verwenden wir Token-Overlap (40% Threshold)
    4. DIREKTE ERSETZUNG: Text wird ersetzt, ohne vollständigen Text zu lesen
    
    Args:
        slide: Aspose.Slides Slide Objekt
        replacements: Liste von Dicts mit 'old_text_snippet' und 'new_text'
        
    Returns:
        int: Anzahl der durchgeführten Ersetzungen (wird via nonlocal zurückgegeben)
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
    
    logger.info(f"      → Prepared {len(clean_replacements)} valid replacements for matching")
    
    # Collect all text from slide for debugging
    all_texts_found = []
    
    def get_visible_text_prefix(text_frame, max_chars=100):
        """
        Gets the visible text prefix from a text frame (works with evaluation version).
        
        WICHTIG: Aspose.Slides Evaluationsversion (kostenlos) schneidet Text ab, wenn man versucht,
        den vollständigen Text zu lesen. Die Meldung "text has been truncated due to evaluation 
        version limitation" erscheint nach ~10-20 Zeichen.
        
        LÖSUNG: Statt den vollständigen Text zu lesen, extrahieren wir nur den sichtbaren Präfix
        (die ersten N Zeichen VOR der Truncation). Dies reicht aus, um die ersten Wörter zu sehen
        und mit Gemini's Vorschlägen zu matchen.
        
        Args:
            text_frame: Aspose TextFrame Objekt
            max_chars: Maximale Anzahl Zeichen zu extrahieren (default: 100)
            
        Returns:
            str: Sichtbarer Text-Präfix (ohne Truncation-Message)
        """
        if not text_frame or not text_frame.paragraphs:
            return ""
        
        # Get text from first paragraph (usually the most important)
        # WICHTIG: Prüfe mit .count statt len() für Aspose Collections
        first_para = text_frame.paragraphs[0] if text_frame.paragraphs and text_frame.paragraphs.count > 0 else None
        if not first_para:
            return ""
        
        # Try to get text - even if truncated, we get the prefix
        visible_text = ""
        try:
            # Method 1: Try paragraph.text (even if truncated, we get prefix)
            # Die Evaluationsversion gibt uns trotzdem die ersten Zeichen
            if hasattr(first_para, 'text'):
                visible_text = first_para.text
                # Remove truncation message if present
                # Die Truncation-Message kommt nach dem eigentlichen Text, also entfernen wir sie
                if "truncated due to evaluation" in visible_text.lower():
                    visible_text = visible_text.split("...")[0] if "..." in visible_text else visible_text.split("text has been")[0]
            
            # Method 2: If that didn't work, try portions
            # Portions könnten manchmal mehr Text geben, aber auch hier gibt es Limits
            # WICHTIG: Prüfe mit .count statt len() für Aspose Collections
            if not visible_text and first_para.portions and first_para.portions.count > 0:
                # Iteriere über Portions (max 5 für Performance)
                portion_texts = []
                for i in range(min(5, first_para.portions.count)):
                    if hasattr(first_para.portions[i], 'text'):
                        portion_texts.append(first_para.portions[i].text)
                visible_text = "".join(portion_texts)
            
            # Limit to max_chars to avoid processing too much
            if len(visible_text) > max_chars:
                visible_text = visible_text[:max_chars]
                
        except Exception as e:
            logger.debug(f"        [DEBUG] Could not extract text prefix: {e}")
        
        return visible_text.strip()
    
    def process_text_frame(text_frame):
        """
        Verarbeitet einen Text-Frame und versucht, Ersetzungen durchzuführen.
        
        WICHTIG: Da wir nur den sichtbaren Präfix haben (wegen Evaluationsversion),
        verwenden wir Präfix-Matching statt vollständigem Text-Matching.
        
        STRATEGIE:
        1. Extrahiere sichtbaren Präfix (erste ~150 Zeichen)
        2. Matche erste 3-10 Wörter der Gemini-Vorschläge gegen diesen Präfix
        3. Wenn Match gefunden → Ersetze direkt, ohne vollständigen Text zu lesen
        """
        nonlocal replacement_count
        if not text_frame or not text_frame.paragraphs:
            return
        
        # Get visible text prefix (works with evaluation version truncation)
        # Wir nehmen 150 Zeichen, um genug Wörter für Matching zu haben
        visible_prefix = get_visible_text_prefix(text_frame, max_chars=150)
        clean_visible = normalize_string(visible_prefix)
        
        # Debug: Log visible text (auch kurze Präfixe loggen)
        if len(clean_visible) >= 1:
            all_texts_found.append(visible_prefix)
            logger.info(f"        [TEXT] Visible prefix: '{visible_prefix}...' (length: {len(clean_visible)} chars)")
        
        # Skip if too short (weniger als 1 Zeichen = leer)
        if len(clean_visible) < 1:
            return
        
        # Debug: Log what we're trying to match against
        logger.debug(f"        [DEBUG] Trying to match against {len(clean_replacements)} replacements for visible text: '{clean_visible[:50]}...'")
        
        # Try to match against each replacement using visible prefix
        # WICHTIG: Wir haben nur den sichtbaren Präfix, daher verwenden wir Präfix-Matching
        logger.debug(f"        [DEBUG] Checking {len(clean_replacements)} replacements against visible: '{clean_visible[:30]}...'")
        for clean_old, new_text, orig_old in clean_replacements:
            logger.debug(f"        [DEBUG]   Checking: '{clean_old[:30]}...' -> '{new_text[:30]}...'")
            is_match = False
            match_type = None
            matched_paragraph = None
            
            # Extract first words from search text (for prefix matching)
            old_words = clean_old.split()
            
            # Strategy 1: Match first 1-10 words of search text against visible prefix
            # WARUM: Die Evaluationsversion zeigt uns nur die ersten ~10-20 Zeichen (oft nur 1-2 Wörter).
            # Wenn Gemini z.B. sucht: "Die Chemieindustrie als wichtiger Absatzmarkt..."
            # und wir sehen: "Die C..." (abgeschnitten), dann müssen wir auch mit 1-2 Wörtern matchen.
            # Wir versuchen von lang zu kurz, damit wir das beste Match bekommen.
            for prefix_len in [10, 8, 5, 3, 2, 1]:  # Auch 1-2 Wörter, da Präfix sehr kurz sein kann
                if len(old_words) >= prefix_len:
                    search_prefix = " ".join(old_words[:prefix_len])
                    search_prefix_normalized = normalize_string(search_prefix)
                    
                    # Check if this prefix matches the visible text
                    # WICHTIG: Da der sichtbare Präfix sehr kurz ist ("die c"), müssen wir prüfen:
                    # 1. Ob der sichtbare Präfix mit dem Such-Präfix beginnt (z.B. "die c" beginnt mit "die")
                    # 2. Ob der Such-Präfix mit dem sichtbaren Präfix beginnt (z.B. "die" beginnt mit "die c" - nein, aber "die c" beginnt mit "die")
                    # 3. Ob der Such-Präfix im sichtbaren Präfix enthalten ist
                    if (clean_visible.startswith(search_prefix_normalized) or 
                        search_prefix_normalized.startswith(clean_visible) or
                        search_prefix_normalized in clean_visible):
                        is_match = True
                        match_type = f"prefix_match_{prefix_len}_words"
                        logger.info(f"        ✓ Match found! First {prefix_len} words match: '{search_prefix[:50]}...' matches visible '{clean_visible[:50]}...'")
                        break
                    
                    # Zusätzlich: Für sehr kurze Präfixe (1-2 Wörter), prüfe auch die ersten Zeichen
                    # Beispiel: visible="die c", search="die chemieindustrie" -> prüfe ob "die" in "die c"
                    if prefix_len <= 2 and len(clean_visible) >= 2:
                        # Prüfe ob das erste Wort des Suchtexts im sichtbaren Präfix vorkommt
                        first_word = old_words[0].lower() if len(old_words) > 0 else ""
                        if first_word and (first_word in clean_visible or clean_visible.startswith(first_word)):
                            is_match = True
                            match_type = f"first_word_match"
                            logger.info(f"        ✓ Match found! First word '{first_word}' matches visible '{clean_visible[:50]}...'")
                            break
            
            # Strategy 2: Token overlap with visible text
            # WARUM: Manchmal sind die ersten Wörter nicht exakt gleich (z.B. Groß-/Kleinschreibung,
            # oder Gemini hat den Text leicht anders formuliert). Token-Overlap ist robuster.
            # Threshold ist niedrig (30%) weil wir nur den Präfix sehen, nicht den ganzen Text.
            if not is_match and len(old_words) >= 2:  # Auch für 2 Wörter prüfen
                old_tokens = set([w.lower() for w in old_words if len(w) > 1])  # Filter: nur Wörter > 1 Zeichen
                visible_tokens = set([w.lower() for w in clean_visible.split() if len(w) > 1])
                if len(old_tokens) > 0:
                    common = old_tokens.intersection(visible_tokens)
                    overlap = len(common) / len(old_tokens) if len(old_tokens) > 0 else 0
                    # Lower threshold since we only see prefix (nicht 80% wie bei vollem Text)
                    # 30% reicht, wenn mindestens 1 Wort übereinstimmt (für sehr kurze Präfixe)
                    min_common = 1 if len(old_tokens) <= 3 else 2  # Für kurze Texte: 1 Wort reicht
                    if overlap >= 0.3 and len(common) >= min_common:
                        is_match = True
                        match_type = f"token_overlap_{overlap:.2f}"
                        logger.info(f"        ✓ Match found! Token overlap: {overlap:.2%} ({len(common)}/{len(old_tokens)} words) - visible: '{clean_visible[:50]}...'")
            
            # Strategy 3: Try matching each paragraph individually
            # WARUM: Manchmal ist der Text über mehrere Paragraphs verteilt, oder der erste
            # Paragraph enthält nicht den gesuchten Text. Wir prüfen jeden Paragraph einzeln.
            if not is_match:
                for paragraph in text_frame.paragraphs:
                    # Get visible text from this paragraph (auch hier nur Präfix wegen Evaluationsversion)
                    para_visible = ""
                    try:
                        if hasattr(paragraph, 'text'):
                            para_visible = paragraph.text
                            # Remove truncation message (kommt nach dem eigentlichen Text)
                            if "truncated" in para_visible.lower():
                                para_visible = para_visible.split("...")[0] if "..." in para_visible else para_visible.split("text has been")[0]
                    except:
                        pass
                    
                    # Fallback: Versuche Portions (könnte mehr Text geben, aber auch limitiert)
                    # WICHTIG: Prüfe mit .count statt len() für Aspose Collections
                    if not para_visible and paragraph.portions and paragraph.portions.count > 0:
                        try:
                            # Iteriere über Portions (max 3 für Performance)
                            portion_texts = []
                            for i in range(min(3, paragraph.portions.count)):
                                if hasattr(paragraph.portions[i], 'text'):
                                    portion_texts.append(paragraph.portions[i].text)
                            para_visible = "".join(portion_texts)
                        except:
                            pass
                    
                    clean_para_visible = normalize_string(para_visible)
                    
                    # Check if first 1-3 words match (auch für sehr kurze Präfixe)
                    for check_len in [3, 2, 1]:
                        if len(old_words) >= check_len:
                            first_words = normalize_string(" ".join(old_words[:check_len]))
                            # Prüfe ob der Paragraph-Text mit diesen Wörtern beginnt oder sie enthält
                            if (first_words in clean_para_visible or 
                                clean_para_visible.startswith(first_words) or
                                (len(clean_para_visible.split()) >= check_len and 
                                 " ".join(clean_para_visible.split()[:check_len]) == first_words)):
                                is_match = True
                                match_type = f"paragraph_prefix_match_{check_len}_words"
                                matched_paragraph = paragraph
                                logger.info(f"        ✓ Match found in paragraph! First {check_len} words match: '{first_words}' in '{clean_para_visible[:50]}...'")
                                break
                        if is_match:
                            break
                    if is_match:
                        break
            
            # ERSETZUNG: Wenn Match gefunden, ersetze direkt
            # WICHTIG: Wir ersetzen OHNE den vollständigen Text zu lesen (geht ja nicht wegen Evaluationsversion).
            # Wir vertrauen darauf, dass unser Präfix-Match korrekt war.
            if is_match:
                if matched_paragraph:
                    # Replace specific paragraph (wenn wir den spezifischen Paragraph identifiziert haben)
                    # WICHTIG: Aspose Collections verwenden .count statt len()
                    if matched_paragraph.portions.count > 0:
                        # Ersetze erste Portion, entferne andere (behält Formatierung)
                        matched_paragraph.portions[0].text = new_text
                        while matched_paragraph.portions.count > 1:
                            matched_paragraph.portions.remove_at(1)
                    else:
                        # Keine Portions vorhanden, setze Text direkt
                        matched_paragraph.text = new_text
                    replacement_count += 1
                    logger.info(f"        ✓ Match found ({match_type})!\n          Search: '{orig_old[:50]}...'\n          Replaced paragraph")
                    return
                else:
                    # Replace entire text frame (first paragraph)
                    # WICHTIG: Wir ersetzen den ersten Paragraph, weil wir nicht wissen können,
                    # welcher Paragraph genau gemeint war (Text ist abgeschnitten).
                    # In den meisten Fällen ist der erste Paragraph der Haupttext.
                    logger.info(f"        ✓ Match found ({match_type})!\n          Search: '{orig_old[:50]}...'\n          Found in visible: '{visible_prefix[:50]}...'")
                    
                    # Replace first paragraph with new text, remove others
                    # WICHTIG: Aspose Collections verwenden .count statt len()
                    if text_frame.paragraphs.count > 0:
                        first_para = text_frame.paragraphs[0]
                        if first_para.portions.count > 0:
                            # Ersetze erste Portion, entferne andere (behält Formatierung der ersten Portion)
                            first_para.portions[0].text = new_text
                            while first_para.portions.count > 1:
                                first_para.portions.remove_at(1)
                        else:
                            # No portions, set text directly
                            first_para.text = new_text
                        
                        # Remove other paragraphs (da wir nur den ersten ersetzen)
                        while text_frame.paragraphs.count > 1:
                            text_frame.paragraphs.remove_at(1)

                            replacement_count += 1
                        return  # Only one replacement per frame (vermeidet doppelte Ersetzungen)
    
    def process_shape(shape):
        # 1. Text Frames (AutoShapes, TextBoxes, etc.)
        if hasattr(shape, "text_frame") and shape.text_frame:
            process_text_frame(shape.text_frame)
        # 2. Groups (recursive)
        if isinstance(shape, IGroupShape):
            for child in shape.shapes:
                process_shape(child)
        # 3. Tables
        if isinstance(shape, ITable):
            for row in shape.rows:
                for cell in row:
                    if cell.text_frame:
                        process_text_frame(cell.text_frame)
        # 4. AutoShapes (explicit check)
        if isinstance(shape, IAutoShape) and hasattr(shape, "text_frame") and shape.text_frame:
            process_text_frame(shape.text_frame)
    
    # Iterate
    for shape in slide.shapes:
        process_shape(shape)
    
    # Debug: Show what we found
    if len(all_texts_found) > 0:
        logger.info(f"      [DEBUG] Found {len(all_texts_found)} text frame(s) with content")
        for i, text in enumerate(all_texts_found[:5], 1):  # Show first 5
            logger.info(f"      [DEBUG]   {i}. '{text}...'")
    else:
        logger.warning(f"      [DEBUG] ⚠️ No text frames with content found on slide!")

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
    """
    Replaces a shape (OLE object or native chart) with a new native PowerPoint chart.
    
    WICHTIG: Diese Funktion funktioniert sowohl mit:
    - OLE Objects (Think-Cell Charts)
    - Native PowerPoint Charts (IChart)
    
    Args:
        slide: Aspose Slide Objekt
        shape: Shape zu ersetzen (OLE oder IChart)
        chart_data: Dict mit chart data aus Gemini JSON
    """
    chart_type_str = chart_data.get('type', 'column')
    chart_title = chart_data.get('title', 'Untitled')
    
    # Determine shape type for logging
    # WICHTIG: IChart ist nicht direkt importierbar, daher prüfen wir über hasattr
    is_native_chart = hasattr(shape, 'chart_data') and hasattr(shape, 'chart_type')
    shape_type = "native chart" if is_native_chart else "OLE object"
    logger.info(f"      → Replacing {shape_type} with {chart_type_str} chart: '{chart_title}'...")

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
    if 'pie' in c_type_lower or 'donut' in c_type_lower:
        chart_type = slides.charts.ChartType.PIE
    elif 'bar' in c_type_lower:
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

    # Clear existing data
    chart_data_obj.series.clear()
    chart_data_obj.categories.clear()

    # Extract data from JSON
    categories = chart_data.get('data', {}).get('categories', [])
    series_list = chart_data.get('data', {}).get('series', [])
    
    logger.info(f"        Data: {len(categories)} categories, {len(series_list)} series")
    logger.info(f"        Categories: {categories}")
    logger.info(f"        Series: {[s.get('name', 'Unknown') for s in series_list]}")

    if not categories or not series_list:
        logger.warning("        ⚠️ No categories or series data provided!")
        return

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

    # WICHTIG: Zuerst alle Daten in die Workbook-Zellen schreiben, dann Series hinzufügen
    # Dies stellt sicher, dass die Daten korrekt gesetzt werden
    
    # Step 1: Write categories to workbook (Row 1..N, Col 0)
    logger.info(f"        → Writing {len(categories)} categories to workbook...")
    category_cells = []
    for i, cat in enumerate(categories):
        # Row = i+1 (start at 1), Col = 0
        cell = workbook.get_cell(0, i + 1, 0, str(cat))
        category_cells.append(cell)
        logger.info(f"        ✓ Category {i+1}/{len(categories)}: '{cat}' written to (row {i+1}, col 0)")
    
    # Add all categories at once
    for cell in category_cells:
        chart_data_obj.categories.add(cell)
    
    # Step 2: Write series names and values to workbook first
    logger.info(f"        → Writing series data to workbook...")
    series_data = []
    for series_idx, s_data in enumerate(series_list):
        s_name = s_data.get('name', f'Series {series_idx + 1}')
        s_vals = s_data.get('values', [])
        s_color = s_data.get('color_hex', '#000000')
        
        logger.info(f"        → Writing series {series_idx + 1}: '{s_name}' with {len(s_vals)} values")
        logger.info(f"          Values: {s_vals}")
        
        # Write series name: Row 0, Col = series_idx + 1
        series_name_cell = workbook.get_cell(0, 0, series_idx + 1, str(s_name))
        logger.info(f"        ✓ Series name '{s_name}' written to (row 0, col {series_idx + 1})")
        
        # Write all values for this series: Row = cat_idx + 1, Col = series_idx + 1
        value_cells = []
        for cat_idx, val in enumerate(s_vals):
            if cat_idx < len(categories):
                try:
                    val_float = float(val)
                except (ValueError, TypeError):
                    val_float = 0.0
                    logger.warning(f"        ⚠️ Invalid value '{val}' for category {cat_idx}, using 0.0")
                
                # Write value to workbook cell
                data_cell = workbook.get_cell(0, cat_idx + 1, series_idx + 1, val_float)
                value_cells.append(data_cell)
                logger.info(f"        ✓ Value {val_float} written to (row {cat_idx+1}, col {series_idx+1}) for category '{categories[cat_idx]}'")
        
        series_data.append({
            'name': s_name,
            'name_cell': series_name_cell,
            'value_cells': value_cells,
            'color': s_color
        })
    
    # Step 3: Now add series to chart (data is already in workbook)
    logger.info(f"        → Adding series to chart...")
    for series_idx, s_info in enumerate(series_data):
        # Add series using the name cell (data is already in workbook)
        series = chart_data_obj.series.add(s_info['name_cell'], chart_type)
        logger.info(f"        ✓ Series '{s_info['name']}' added to chart")
        
        # Add data points using the pre-written cells
        # WICHTIG: Wir verwenden die Zellen, die bereits in der Workbook geschrieben wurden
        values_added = 0
        for idx, data_cell in enumerate(s_info['value_cells']):
            try:
                # Methode 1: Standard Aspose API
                if chart_type == slides.charts.ChartType.PIE:
                    series.data_points.add_data_point_for_pie_series(data_cell)
                elif chart_type == slides.charts.ChartType.LINE:
                    series.data_points.add_data_point_for_line_series(data_cell)
                else:
                    series.data_points.add_data_point_for_bar_series(data_cell)
                values_added += 1
                logger.debug(f"        → Data point {idx+1} added via API")
            except Exception as e:
                logger.warning(f"        ⚠️ API method failed for data point {idx+1}, trying alternative: {e}")
                # Methode 2: Alternative - direkt über data_points[index]
                try:
                    if series.data_points.count > idx:
                        # Falls bereits ein DataPoint existiert, ersetze ihn
                        dp = series.data_points[idx]
                        dp.value.data = data_cell
                    else:
                        # Sonst füge neuen hinzu
                        if chart_type == slides.charts.ChartType.PIE:
                            series.data_points.add_data_point_for_pie_series(data_cell)
                        elif chart_type == slides.charts.ChartType.LINE:
                            series.data_points.add_data_point_for_line_series(data_cell)
                        else:
                            series.data_points.add_data_point_for_bar_series(data_cell)
                    values_added += 1
                    logger.debug(f"        → Data point {idx+1} added via alternative method")
                except Exception as e2:
                    logger.error(f"        ✗ Both methods failed for data point {idx+1}: {e2}")
        
        logger.info(f"        ✓ Series '{s_info['name']}' completed: {values_added}/{len(s_info['value_cells'])} data points added")
        
        # Verify data was set correctly
        if values_added > 0:
            logger.info(f"        ✓ Verified: Series has {series.data_points.count} data points")
        else:
            logger.error(f"        ✗ WARNING: No data points were added to series '{s_info['name']}'!")
        
        # Set color
        try:
            series.format.fill.fill_type = slides.FillType.SOLID
            a, r, g, b = hex_to_argb(s_info['color'])
            series.format.fill.solid_fill_color.color = drawing.Color.from_argb(a, r, g, b)
            logger.info(f"        ✓ Color applied: {s_info['color']} -> ARGB({a}, {r}, {g}, {b})")
        except Exception as e:
            logger.warning(f"        ⚠️ Could not set color {s_info['color']}: {e}")

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
    
    # Try to load license first
    load_aspose_license_if_available()
    
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
        
        # --- INTELLIGENT MAPPING STRATEGY ---
        # Wir finden alle Kandidaten und sortieren sie nach Position (Oben->Unten, Links->Rechts)
        # Damit matchen wir die Logik aus dem Prompt (Top-Left -> Bottom-Right).
        
        # Find ALL chart candidates: OLE objects (Think-Cell) AND native PowerPoint charts
        chart_candidates = []
        for shape in slide.shapes:
            # Check for OLE objects (Think-Cell charts)
            if hasattr(slides, 'OleObjectFrame') and isinstance(shape, slides.OleObjectFrame):
                chart_candidates.append(shape)
                logger.debug(f"      → Found OLE object (Think-Cell chart) at (X={shape.x:.0f}, Y={shape.y:.0f})")
            elif hasattr(shape, 'ole_format'):
                chart_candidates.append(shape)
                logger.debug(f"      → Found OLE object (via ole_format) at (X={shape.x:.0f}, Y={shape.y:.0f})")
            # Check for native PowerPoint charts
            # WICHTIG: IChart ist nicht direkt importierbar, daher prüfen wir über hasattr
            elif hasattr(shape, 'chart_data') and hasattr(shape, 'chart_type'):
                chart_candidates.append(shape)
                logger.debug(f"      → Found native PowerPoint chart at (X={shape.x:.0f}, Y={shape.y:.0f})")

        logger.info(f"      → Found {len(chart_candidates)} chart candidate(s) on slide (OLE + native)")
        
        # --- POSITION-BASED SORTING ---
        # Sortieren: Zuerst nach Y (Zeile), dann nach X (Spalte)
        # Wir nutzen eine einfache "Reading Order" Logik: Y*1000 + X
        # Dies sortiert von oben nach unten, und bei gleicher Y-Position von links nach rechts
        chart_candidates.sort(key=lambda s: (int(s.y / 50) * 1000 + s.x))
        
        logger.info(f"      → Sorted {len(chart_candidates)} charts by position (Top-Left -> Bottom-Right)")
        for idx, candidate in enumerate(chart_candidates):
            logger.debug(f"        Chart {idx+1}: Position (X={candidate.x:.0f}, Y={candidate.y:.0f})")

        # Match and replace charts
        # WICHTIG: Die Charts sind jetzt sortiert (Top-Left -> Bottom-Right)
        # Die AI-Daten sollten auch in dieser Reihenfolge sein (laut Prompt)
        # Falls mehr Charts gefunden werden als in JSON, verwenden wir nur die ersten N
        for i, chart_shape in enumerate(chart_candidates):
            if i < len(charts):
                chart_data = charts[i]
                position_hint = chart_data.get('position_hint', 'unknown')
                logger.info(f"      → Replacing chart {i+1}/{len(chart_candidates)} (AI hint: '{position_hint}') with data from AI...")
                replace_ole_with_chart(slide, chart_shape, chart_data)
            else:
                logger.warning(f"      ⚠️ More charts found ({len(chart_candidates)}) than AI provided ({len(charts)})")
    else:
        logger.info("      → No chart replacements to apply")

    logger.info("      [Step 2] Saving modified presentation...")
    pres.save(output_path, slides.export.SaveFormat.PPTX)
    step_time = time.time() - step_start
    file_size = os.path.getsize(output_path)
    logger.info(f"      [Step 2] ✓ Presentation saved in {step_time:.2f}s ({file_size:,} bytes)")
