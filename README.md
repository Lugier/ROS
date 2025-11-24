# PowerPoint Automation Agent

Ein KI-gestütztes Präsentations-Automatisierungs-System, das aus bestehenden PowerPoint-Templates vollständig neue, personalisierte Slidesets generiert – ähnlich wie ein **"Junior Consultant, der 80% der Arbeit übernimmt"**.

## 🎯 Vision

Ein automatisierter **"PowerPoint-First-Draft-Generator"** für Beratungen, der:
- **80% der Folienqualität** eines Junior Consultants liefert
- **Templates respektiert** und Layouts strikt erhält
- **Think-Cell-Charts automatisch ersetzt** durch native PowerPoint Charts
- **Automatisch recherchiert** und aktualisiert (Marktdaten, KPIs, CAGR, etc.)
- **Enorme Zeit spart** pro Slide

Dieses System dient als Basis für ein komplettes Agentic Tool, das später:
- Ganze Decks statt einzelner Slides generiert
- Finanzen analysiert
- Marktanalysen automatisiert
- Reportings skaliert produziert

## 🚀 Was das System tut

### 1. User lädt PowerPoint-Templates hoch

Consulting-Firmen wie **Deloitte, PwC, BCG, Bain, Strategy&, Roland Berger** usw. haben interne Standard-Slides:
- Marketslides
- Finanzgrafiken
- Summary-Slides

Diese werden per Upload bereitgestellt (eine oder mehrere PPTX-Dateien).

### 2. User gibt eine neue Aufgabe ein

**Beispiel:**
> "Erstelle eine Markt-Slide für NVIDIA zur Halbleiterentwicklung in der DACH-Region. Ersetze Mercedes-Benz durch NVIDIA."

### 3. System analysiert die Template-Slide visuell

**PPTX → PNG-Rendering**

**Gemini 3 Pro (Vision + Web Search)** versteht:
- Texte, Headlines, Strukturen, Layout
- Think-Cell Charts
- Welche Elemente ersetzt werden müssen

### 4. System recherchiert automatisch relevante Inhalte

- Marktgröße
- CAGR
- Wettbewerbsübersicht
- Relevante KPIs
- Strukturierte Daten fürs Chart

### 5. System liefert strukturierte JSON-Instruktionen

JSON enthält u. a.:
- `replacements` (alte → neue Texte)
- `charts` (vollständige Datenreihen)
- Flag, ob Think-Cell ersetzt werden soll

### 6. System rendert eine neue PPTX

Mit **Aspose.Slides** (local + evaluation mode):
- Texte ersetzen (alte Firma → neue Firma)
- Bulletpoints neu setzen
- Think-Cell-Charts erkennen (OLE Frames)
- Think-Cell entfernen
- Neues natives PowerPoint-Chart einfügen
- Serien, Kategorien, Farben, Titel setzen
- **Layout strikt erhalten**

### 7. User erhält neue PPTX-Datei zum Download

Mit aktualisierten Daten, Grafiken, Texten – im **Template-Look der Consulting-Firma**.

## 🛠️ Technischer Stack

### Framework & Core
- **Flask** - Web App für Upload-Interface
- **Gemini 3 Pro Preview** - Vision + Web Search für Analyse & Recherche
- **Aspose.Slides** (Python via .NET) - PPTX-Manipulation

### Hauptkomponenten

#### `app.py`
- Flask Web App
- Upload: PPTX + Prompt
- Download: Generierte PPTX

#### `utils/vision_analyzer.py`
- Gemini Vision + Research
- PPTX → PNG Rendering
- Slide-Verständnis
- Textanalyse
- Marktrecherche
- JSON-Output

#### `utils/slide_renderer.py`
- Aspose PPTX Editor
- Text-Replace (mit Formatierungserhaltung)
- Chart-Rebuild
- Think-Cell-Replacements
- OLE Frame Detection & Removal
- Native Chart Creation

#### `templates/index.html`
- Upload UI
- Multi-File Support
- Progress Tracking

## 📋 Projektstruktur

```
ROS/
├── app.py                      # Flask Web Application
├── requirements.txt            # Python Dependencies
├── .env                       # Environment Variables (GEMINI_API_KEY)
├── README.md                  # Diese Datei
├── utils/
│   ├── __init__.py
│   ├── vision_analyzer.py    # Gemini Vision + Research
│   └── slide_renderer.py     # Aspose PPTX Editor
└── templates/
    └── index.html             # Upload UI
```

## 🔑 Features

### ✅ Implementiert
- [x] PPTX Template Upload
- [x] Gemini 3 Pro Preview Integration
- [x] Vision-basierte Slide-Analyse
- [x] Google Search Integration für Recherche
- [x] Text-Ersetzung mit Formatierungserhaltung
- [x] Think-Cell Chart Detection & Replacement
- [x] Native PowerPoint Chart Creation
- [x] Detailliertes Logging mit ETA
- [x] Error Handling

### 🚧 Geplant / Roadmap
- [ ] Multi-File Upload (mehrere Templates gleichzeitig)
- [ ] Live Progress Updates (Server-Side Events)
- [ ] Template-Vorschau
- [ ] Beispiel-Prompts für typische Use Cases
- [ ] Batch-Processing für ganze Decks
- [ ] Finanzanalyse-Integration
- [ ] Marktanalyse-Automatisierung
- [ ] Reporting-Skalierung

## 🚀 Installation & Setup

### 1. Dependencies installieren

```bash
pip install -r requirements.txt
```

**Wichtig:** Die neue `google-genai` Bibliothek wird verwendet (nicht mehr `google-generativeai`). 
Falls du die alte Bibliothek noch installiert hast, wird sie durch die neue ersetzt.

### 2. Environment Variables

Erstelle eine `.env` Datei:

```env
GEMINI_API_KEY=dein_api_key_hier
```

### 3. Server starten

```bash
python app.py
```

Die App läuft dann auf `http://localhost:5000`

## 📝 Verwendung

1. **Template hochladen**: PPTX-Datei im Browser auswählen
2. **Prompt eingeben**: Beschreibung der gewünschten Anpassung
3. **Verarbeitung**: System analysiert, recherchiert und generiert
4. **Download**: Angepasste PPTX-Datei herunterladen

### Beispiel-Prompts

- "Erstelle eine Markt-Slide für NVIDIA zur Halbleiterentwicklung in der DACH-Region. Ersetze Mercedes-Benz durch NVIDIA."
- "Adaptiere diese Finanzgrafik für Q4 2024 mit aktuellen Zahlen für die Automobilindustrie."
- "Erstelle eine Wettbewerbsübersicht für SaaS-Unternehmen im B2B-Bereich."

## 🔧 Technische Details

### Text-Ersetzung
- **Formatierungserhaltung**: Text wird über `paragraphs` und `portions` ersetzt, nicht direkt über `text_frame.text`
- **Farben & Styles**: Bold, Italic, Farben bleiben erhalten

### Think-Cell Replacement
- **OLE Frame Detection**: Automatische Erkennung von Think-Cell Charts
- **Position Preservation**: X, Y, Width, Height werden exakt übernommen
- **Data Injection**: Vollständige Datenreihen aus Gemini Research
- **Color Mapping**: Hex-Farben aus Vision-Analyse werden angewendet

### Chart Types
- Bar Charts
- Column Charts
- Line Charts
- Weitere Typen können erweitert werden

## 📊 Logging

Das System bietet detailliertes Logging mit:
- **Phasen-basierte Fortschrittsanzeige** (Phase 0/3, 1/3, 2/3, 3/3)
- **ETA-Berechnungen** mit geschätzter Fertigstellungszeit
- **Schritt-für-Schritt Details** für jeden Verarbeitungsschritt
- **Fehlerbehandlung** mit vollständigem Traceback

## 🎯 Use Cases

### Consulting-Firmen
- **Template-basierte Slide-Generierung** für Kundenpräsentationen
- **Marktanalysen** mit aktuellen Daten
- **Finanzgrafiken** mit automatischer Recherche
- **Wettbewerbsübersichten** mit Live-Daten

### Agentic Automation
- Basis für vollautomatische Deck-Generierung
- Skalierbare Reporting-Produktion
- Konsistente Template-Nutzung

## 🔒 Wichtige Hinweise

- **Aspose.Slides**: Läuft im Evaluation Mode (Watermark in generierten Dateien)
- **Gemini API**: Benötigt gültigen API-Key mit Zugriff auf Gemini 3 Pro Preview
- **Think-Cell**: Erfordert Think-Cell Charts in den Templates (werden als OLE Objects erkannt)

## 📄 License

Dieses Projekt ist für interne Nutzung in Consulting-Firmen konzipiert.

## 🤝 Contributing

Dieses System ist als Basis für weitere Agentic Tools gedacht. Erweiterungen sind willkommen!

---

**Status**: ✅ Funktionsfähig - Ready for Testing

**Version**: 1.0.0

**Letzte Aktualisierung**: November 2025

