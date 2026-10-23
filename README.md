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
     - **Clear All Fields**: Instantly wipes all inputs, resets dropdown to default, and clears status.
   - **Live Status & Notification**:
     - Real-time conversion feedback.
     - Confirmation banner displaying the storage destination folder.
     - Quick-access buttons: **Open Markdown File** and **Open Stored Folder**.

---

## 📂 Supported Formats

| Category | Formats | Processing Method |
| :--- | :--- | :--- |
| **PDF Documents** | `.pdf` | Structured text, tables, and heading extraction via `pdfminer` / `pdfplumber` |
| **Word Documents** | `.docx`, `.doc` | Headings, lists, styles, and tables via `mammoth` |
| **PowerPoint** | `.pptx`, `.ppt` | Slide text, outlines, and notes via `python-pptx` |
| **Excel Sheets** | `.xlsx`, `.xls` | Spreadsheets converted to Markdown tables via `openpyxl` & `pandas` |
| **Images** | `.jpg`, `.jpeg`, `.png`, `.webp` | EXIF metadata extraction and OCR parsing |
| **Audio Files** | `.mp3`, `.wav`, `.m4a` | Speech transcription into text and EXIF audio metadata |
| **Video Files** | `.mp4`, `.mkv`, `.mov`, `.avi` | Audio track extraction via `moviepy` followed by speech transcription |
| **HTML / Web** | `.html`, `.htm` | HTML parsing, DOM cleanup, and conversion via `markdownify` & `beautifulsoup4` |
| **Text & Data** | `.csv`, `.json`, `.xml`, `.txt` | Formatted Markdown code blocks and table generation |
| **ZIP Archives** | `.zip` | Extracts and iterates over all files inside the archive |
| **YouTube & Web URLs** | `https://...` | Transcribes YouTube video speech and scrapes web article text |
| **E-Books** | `.epub` | Chapters, chapters headings, and formatted narrative extraction |

---

## 🚀 How to Run the Desktop App

### Prerequisites
Make sure Python 3.10+ is installed on your system.

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
│  2. If Video (.mp4): Extract audio via       │
│     MoviePy -> Transcribe speech             │
│  3. If Document/Audio/Image/ZIP:             │
│     Process via Microsoft MarkItDown engine  │
│  4. Fallback formatting for CSV/JSON/XML     │
└──────────────────────┬───────────────────────┘
                       │
                       ▼
┌──────────────────────────────────────────────┐
│           Output File Generation             │
│  - Appends output_dir + filename.md          │
│  - Writes UTF-8 encoded Markdown file        │
│  - Updates GUI with Success & Folder Paths   │
└──────────────────────────────────────────────┘
```

1. **User Interaction**:
   - The user selects the file type from the dropdown, chooses a file or pastes a link, names the file, and chooses the target folder.
2. **Background Thread Execution**:
   - Upon clicking **Convert**, `app.py` spawns a background `threading.Thread`. This prevents the Windows window from entering "Not Responding" state during long OCR or transcription tasks.
3. **Format Engine**:
   - `converter.py` routes the file to the appropriate parser:
     - For video files, it uses `moviepy` to extract the sound track into a temporary audio buffer and passes it to the speech-to-text pipeline.
     - For URLs, it fetches transcripts or parses HTML into markdown.
     - For documents (PDF, DOCX, XLSX, PPTX), it preserves headings, bold text, lists, and tables.
4. **Saving**:
   - Combines the storage path and filename (`os.path.join(output_dir, filename)`), ensures the `.md` extension, and saves the file.
5. **Notification**:
   - The UI shows a green success banner with the complete path and reveals **Open Markdown File** and **Open Stored Folder** buttons for instant access.
6. **Clear Action**:
   - Clicking **Clear** wipes all inputs, resets dropdown to "All Supported Formats", and clears status.

---

## 🛠️ Technologies Used

| Technology | Purpose |
| :--- | :--- |
| **Python 3.11** | Core programming language runtime |
| **CustomTkinter** | Modern UI framework with native dark/light styling and high-DPI scaling |
| **Microsoft MarkItDown** | State-of-the-art document conversion engine |
| **MoviePy & imageio-ffmpeg** | Video audio stream extraction and processing |
| **Pillow (PIL)** | Image EXIF metadata parsing and dimensional inspection |
| **BeautifulSoup4 & Markdownify** | Web and HTML page to Markdown converter |
| **pdfplumber & pdfminer.six** | High-precision PDF layout and text parsing |
| **Mammoth & python-pptx & openpyxl** | Microsoft Office suite (Word, PowerPoint, Excel) parsers |
| **SpeechRecognition & pydub** | Audio speech-to-text extraction |
| **YouTube Transcript API** | YouTube video caption and transcript extractor |
