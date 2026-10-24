# MD Converter Studio 📑

A modern, high-performance Python desktop application to convert virtually any file format or web/video URL into clean, structured **Markdown (`.md`)** files.

Powered by **CustomTkinter** for a sleek modern UI and **Microsoft MarkItDown** + **MoviePy** for conversion.

---

## 🌟 Key Features

1. **Modern UI Layout**:
   - **Top Navigation Bar**: App branding, version tag, and instant **Dark / Light Mode** theme toggle.
   - **Format Selector Dropdown**: Categorized format selection (PDF, Word, PPTX, Excel, Images, Audio, Video, HTML, Data, ZIP, URLs, EPubs).
   - **Input Box & Browse Button**: Select a local file via native file picker dialog or paste any web/YouTube URL.
   - **Custom Output Filename Field**: Enter any custom file name (e.g. `report.md`). Auto-completes the `.md` extension automatically.
   - **Storage Destination Field**: Dedicated folder picker and path box to choose exactly where the Markdown file is stored.
   - **Convert & Clear Actions**:
     - **Convert to .md**: Asynchronously processes the file without freezing the UI.
     - **Clear All Fields**: Instantly wipes all inputs, resets dropdown to default, and clears st    - **Live Status & Notification**:
      - Real-time conversion feedback.
      - Confirmation banner displaying the storage destination folder.
      - Quick-access buttons:
        - **📄 Open Markdown File**: Opens the converted `.md` file in the system default text editor.
        - **📂 Open Stored Folder**: Reveals the destination directory in File Explorer.
        - **📤 Share .md**: Launches a dedicated modal to share the file directly to **WhatsApp** or **Microsoft Teams** (only these two options are presented), with automated clipboard synchronization for instant pasting (Ctrl+V).
2. **Tesseract OCR Image-to-Markdown Engine**:
   - Converts PNG, JPG, JPEG, BMP, TIFF, and WEBP images into structured Markdown documents.
   - Intelligent Pillow image preprocessing (grayscale and contrast optimization).
   - Extracts printed text using Tesseract OCR with automatic fallback to native Tesseract CLI.
   - Enriches Markdown with image dimension specs and EXIF camera metadata.
3. **Dedicated Excel Spreadsheet Engine**:
   - Native pure-Python `.xlsx` (via `openpyxl`) and `.xls` (via `xlrd`) parsing.
   - Transforms multi-sheet workbooks into clean GitHub Flavored Markdown (GFM) tables without failing on missing/blocked pandas dependencies.

---

## 📂 Supported Formats

| Category | Formats | Processing Method |
| :--- | :--- | :--- |
| **Images (OCR)** | `.png`, `.jpg`, `.jpeg`, `.bmp`, `.tiff`, `.webp` | **Tesseract OCR Engine** + Pillow preprocessing & EXIF extraction |
| **PDF Documents** | `.pdf` | Structured text, tables, and heading extraction via `pdfminer` / `pdfplumber` |
| **Word Documents** | `.docx`, `.doc` | Modern `.docx` via `mammoth`/`markitdown`; Legacy `.doc` via 3-Tier OLE parser |
| **PowerPoint** | `.pptx`, `.ppt` | Modern `.pptx` via `python-pptx`; Legacy `.ppt` via 3-Tier OLE parser |
| **Excel Sheets** | `.xlsx`, `.xls` | Pure-Python tabular extraction into Markdown tables via `openpyxl` & `xlrd` |
| **Audio Files** | `.mp3`, `.wav`, `.m4a` | Speech transcription into text and EXIF audio metadata |
| **Video Files** | `.mp4`, `.mkv`, `.mov`, `.avi` | Audio track extraction via `moviepy` followed by speech transcription |
| **HTML / Web** | `.html`, `.htm` | HTML parsing, DOM cleanup, and conversion via `markdownify` & `beautifulsoup4` |
| **Text & Data** | `.csv`, `.json`, `.xml`, `.txt` | Formatted Markdown code blocks and table generation |
| **ZIP Archives** | `.zip` | Extracts and iterates over all files inside the archive |
| **YouTube & Web URLs** | `https://...` | Transcribes YouTube video speech and scrapes web article text |
| **E-Books** | `.epub` | Chapters, chapters headings, and formatted narrative extraction |

---

## 🚀 How to Run the Desktop App

### Prerequisites & Installation
1. Install Python 3.10+ on your system.
2. Install the required dependencies:
   ```bash
   pip install -r requirements.txt
   ```
> 📖 **Full System Setup & Workflow Guide**: See [SETUP_AND_WORKFLOW_GUIDE.md](SETUP_AND_WORKFLOW_GUIDE.md) for cross-platform instructions (Windows, macOS, Linux), virtual environment setup, and pipeline details.
> 🚀 **Future Roadmap & Optimizations**: See [FUTURE_UPDATES_AND_OPTIMIZATIONS.md](FUTURE_UPDATES_AND_OPTIMIZATIONS.md) for the strategic feature and performance roadmap.

### Method 1: Double-Click (Easiest)
Simply double-click the `run.bat` file in the folder:
```
run.bat
```

### Method 2: Command Line (PowerShell or Command Prompt)
Open your terminal in the project directory (`c:\Users\sasi_\Downloads\MD Convert`) and execute:

```powershell
python app.py
```

### Method 3: PowerShell Script
```powershell
.\run.ps1
```

---

## 📤 Sharing Converted Files (WhatsApp & Microsoft Teams)

When conversion finishes, click the purple **📤 Share .md** button. The Share modal provides **strictly two options**:

1. **💬 WhatsApp**:
   - Copies the complete Markdown content to your clipboard.
   - Automatically opens WhatsApp (Desktop application `whatsapp://` or WhatsApp Web `https://web.whatsapp.com/`).
   - Allows instant pasting (`Ctrl+V`) of the Markdown text into any chat.
2. **👥 Microsoft Teams**:
   - Copies the complete Markdown content to your clipboard.
   - Automatically opens Microsoft Teams (`msteams:/` or Teams Web Share).
   - Allows instant pasting (`Ctrl+V`) of the Markdown document into any channel or direct message.

---

## ⚙️ How It Works (Architecture & Pipeline)

```
[ User Input (File / URL) ]
            │
            ▼
┌──────────────────────────────────────────────┐
│       app.py (CustomTkinter GUI)             │
│  - Formats dropdown & path validations       │
│  - Background Thread Worker (No UI freeze)   │
└──────────────────────┬───────────────────────┘
                       │
                       ▼
┌──────────────────────────────────────────────┐
│         converter.py (Engine)                │
│  1. Check Source Type (URL vs File)          │
│  2. If Image: Tesseract OCR + Pillow         │
│  3. If Excel: openpyxl / xlrd Markdown table │
│  4. If Video: MoviePy audio -> Speech-to-Text│
│  5. If Office/PDF: MarkItDown / OLE Parsers  │
│  6. Fallback formatting for CSV/JSON/XML     │
└──────────────────────┬───────────────────────┘
                       │
                       ▼
┌──────────────────────────────────────────────┐
│           Output File Generation             │
│  - Appends output_dir + filename.md          │
│  - Writes UTF-8 encoded Markdown file        │
│  - Reveals: Open File, Open Folder, Share    │
└──────────────────────────────────────────────┘
```

---

## 🛠️ Technologies Used

| Technology | Purpose |
| :--- | :--- |
| **Python 3.11** | Core programming language runtime |
| **CustomTkinter** | Modern UI framework with native dark/light styling and high-DPI scaling |
| **Tesseract OCR & pytesseract** | Optical Character Recognition for image-to-text conversion |
| **openpyxl & xlrd** | Pure-Python Excel spreadsheet to Markdown table converters |
| **Microsoft MarkItDown** | State-of-the-art document conversion engine |
| **MoviePy & imageio-ffmpeg** | Video audio stream extraction and processing |
| **Pillow (PIL)** | Image contrast/grayscale preprocessing and EXIF parsing |
| **BeautifulSoup4 & Markdownify** | Web and HTML page to Markdown converter |
| **pdfplumber & pdfminer.six** | High-precision PDF layout and text parsing |
| **Mammoth & python-docx & python-pptx** | Microsoft Word and PowerPoint document parsers |
| **SpeechRecognition & pydub** | Audio speech-to-text extraction |
| **YouTube Transcript API** | YouTube video caption and transcript extractor |

