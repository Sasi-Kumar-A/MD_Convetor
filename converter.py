"""
Markdown Converter Engine
Handles conversion of various file formats and URLs to Markdown (.md)
using Microsoft's MarkItDown and complementary tools (MoviePy, OCR, OLE parsers, etc.).
"""

import os
import sys

# Prevent Windows Application Control policy DLL blocking errors when pytesseract or markitdown try to import pandas
if "pandas" not in sys.modules:
    try:
        import pandas
    except (ImportError, Exception):
        sys.modules["pandas"] = None

import tempfile
import uuid
import struct
import re
import shutil
import subprocess
from typing import Callable, Optional, Tuple
from markitdown import MarkItDown

# OLE compound file reader for legacy Office formats (.doc, .ppt)
try:
    import olefile
except ImportError:
    olefile = None

# Video conversion helper
try:
    from moviepy import VideoFileClip
except ImportError:
    try:
        from moviepy.editor import VideoFileClip
    except ImportError:
        VideoFileClip = None

# Tesseract OCR and Image processing
try:
    from PIL import Image, ImageOps, ImageFilter
except ImportError:
    Image = None

try:
    import pytesseract
except ImportError:
    pytesseract = None

# Excel parsing libraries (pure Python, zero-dependency fallbacks)
try:
    import openpyxl
except ImportError:
    openpyxl = None

try:
    import xlrd
except ImportError:
    xlrd = None

TESSERACT_EXE_PATH = None


def _configure_tesseract_path():
    """Locate and set the Tesseract OCR executable path for pytesseract and direct CLI fallback."""
    global TESSERACT_EXE_PATH
    tess_path = None

    # Check common Windows installation paths
    if sys.platform == "win32":
        candidate_paths = [
            r"C:\Program Files\Tesseract-OCR\tesseract.exe",
            r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
            os.path.expanduser(r"~\AppData\Local\Programs\Tesseract-OCR\tesseract.exe")
        ]
        for path in candidate_paths:
            if os.path.exists(path):
                tess_path = path
                break

    # Fall back to system PATH detection
    if not tess_path:
        tess_path = shutil.which("tesseract")

    if tess_path:
        TESSERACT_EXE_PATH = tess_path
        if pytesseract is not None:
            pytesseract.pytesseract.tesseract_cmd = tess_path


_configure_tesseract_path()


class MarkdownConverter:
    """Core conversion engine supporting a wide range of document and media formats."""

    VIDEO_EXTENSIONS = {".mp4", ".mkv", ".mov", ".avi", ".wmv", ".flv", ".webm"}
    AUDIO_EXTENSIONS = {".mp3", ".wav", ".m4a", ".aac", ".ogg", ".flac"}
    IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tiff", ".gif"}
    DOC_EXTENSIONS = {".pdf", ".docx", ".doc", ".pptx", ".ppt", ".xlsx", ".xls", ".epub"}
    TEXT_EXTENSIONS = {".html", ".htm", ".txt", ".csv", ".json", ".xml", ".tsv"}
    ARCHIVE_EXTENSIONS = {".zip"}

    def __init__(self):
        self.md = MarkItDown()

    def is_url(self, target: str) -> bool:
        """Check if the provided input is a web URL."""
        if not target:
            return False
        clean = target.strip().lower()
        return clean.startswith("http://") or clean.startswith("https://") or clean.startswith("www.")

    def convert_source(
            self,
            source: str,
            progress_callback: Optional[Callable[[str], None]] = None
    ) -> str:
        """
        Converts the source (file path or URL) to a Markdown string.

        Args:
            source: Local file path or web/YouTube URL.
            progress_callback: Optional callable for status updates.

        Returns:
            Extracted Markdown string.
        """
        source = source.strip()
        if not source:
            raise ValueError("Input file path or URL cannot be empty.")

        # Check if URL
        if self.is_url(source):
            if not (source.startswith("http://") or source.startswith("https://")):
                source = "https://" + source
            if progress_callback:
                progress_callback(f"Fetching and parsing URL: {source}...")
            result = self.md.convert(source)
            content = result.text_content.strip() if result and result.text_content else ""
            if not content:
                content = f"# Webpage Content: {source}\n\n*No readable text could be extracted from this URL.*"
            return content

        # Local file validation
        if not os.path.exists(source):
            raise FileNotFoundError(f"File not found: '{source}'")

        _, ext = os.path.splitext(source)
        ext_lower = ext.lower()

        # Dedicated handling for Image files using Tesseract OCR
        if ext_lower in self.IMAGE_EXTENSIONS:
            return self._convert_image_ocr(source, progress_callback=progress_callback)

        # Dedicated handling for Excel spreadsheets (.xlsx, .xls)
        if ext_lower in {".xlsx", ".xls"}:
            return self._convert_excel(source, progress_callback=progress_callback)

        # Special handling for Video files: Extract audio and convert
        if ext_lower in self.VIDEO_EXTENSIONS:
            if VideoFileClip is None:
                raise RuntimeError(
                    f"Video conversion requires 'moviepy' and FFmpeg to extract audio. "
                    f"Please install moviepy or select an audio/document file."
                )

            if progress_callback:
                progress_callback("Extracting audio stream from video...")

            temp_audio = os.path.join(tempfile.gettempdir(), f"audio_extract_{uuid.uuid4().hex[:8]}.mp3")
            try:
                clip = VideoFileClip(source)
                if clip.audio is None:
                    clip.close()
                    return f"# Video Analysis: {os.path.basename(source)}\n\n*This video file has no audio track to transcribe.*"

                clip.audio.write_audiofile(temp_audio, logger=None)
                clip.close()

                if progress_callback:
                    progress_callback("Transcribing extracted audio to Markdown...")

                result = self.md.convert(temp_audio)
                text = result.text_content.strip() if result and result.text_content else ""

                header = f"# Video Transcription: {os.path.basename(source)}\n\n"
                if not text:
                    return header + "*No speech or transcript was detected in this video file.*"
                return header + text
            finally:
                if os.path.exists(temp_audio):
                    try:
                        os.remove(temp_audio)
                    except OSError:
                        pass

        # Special handling for legacy Word .doc files
        if ext_lower == ".doc":
            return self._convert_doc(source, progress_callback=progress_callback)

        # Special handling for legacy PowerPoint .ppt files
        if ext_lower == ".ppt":
            return self._convert_ppt(source, progress_callback=progress_callback)

        # Standard conversion via MarkItDown
        if progress_callback:
            progress_callback(f"Converting {os.path.basename(source)} to Markdown...")

        try:
            result = self.md.convert(source)
            content = result.text_content.strip() if result and result.text_content else ""
        except Exception as e:
            if ext_lower in {".txt", ".csv", ".json", ".xml", ".tsv"}:
                content = self._fallback_text_convert(source, ext_lower)
            else:
                raise e

        if not content:
            if ext_lower in {".txt", ".csv", ".json", ".xml", ".tsv"}:
                content = self._fallback_text_convert(source, ext_lower)
            else:
                content = f"# {os.path.basename(source)}\n\n*(Conversion finished, but no text or captions were found in the file)*"

        return content

    # ------------------ IMAGE OCR (TESSERACT) CONVERSION ------------------

    def _convert_image_ocr(
            self,
            file_path: str,
            progress_callback: Optional[Callable[[str], None]] = None
    ) -> str:
        """
        Extracts textual content from images using Tesseract OCR,
        along with structural image dimensions and EXIF metadata.
        """
        filename = os.path.basename(file_path)
        if progress_callback:
            progress_callback(f"Running Tesseract OCR on image: {filename}...")

        extracted_text = ""
        ocr_error = None

        # 1. Primary Method: pytesseract with Pillow image preprocessing
        if pytesseract is not None and Image is not None:
            try:
                with Image.open(file_path) as img:
                    # Convert palette or transparency modes to RGB
                    if img.mode in ("P", "RGBA", "LA", "CMYK"):
                        processed_img = img.convert("RGB")
                    else:
                        processed_img = img.copy()

                    # Enhance contrast and optimize grayscale for OCR
                    grayscale = ImageOps.grayscale(processed_img)
                    enhanced_img = ImageOps.autocontrast(grayscale)

                    # Extract text using Tesseract
                    extracted_text = pytesseract.image_to_string(enhanced_img).strip()

                    # If enhanced image yielded empty text, try raw image directly
                    if not extracted_text:
                        extracted_text = pytesseract.image_to_string(processed_img).strip()

            except Exception as e:
                ocr_error = str(e)

        # 2. Resilient Fallback: Direct Tesseract binary invocation via subprocess
        if not extracted_text and TESSERACT_EXE_PATH and os.path.exists(TESSERACT_EXE_PATH):
            try:
                if progress_callback:
                    progress_callback(f"Running direct Tesseract OCR binary on: {filename}...")
                cli_res = subprocess.run(
                    [TESSERACT_EXE_PATH, os.path.abspath(file_path), "stdout", "-l", "eng"],
                    capture_output=True,
                    text=True,
                    timeout=30
                )
                if cli_res.stdout and cli_res.stdout.strip():
                    extracted_text = cli_res.stdout.strip()
                    ocr_error = None
            except Exception as cli_err:
                if not ocr_error:
                    ocr_error = str(cli_err)

        if not extracted_text and not ocr_error and pytesseract is None and not TESSERACT_EXE_PATH:
            ocr_error = "Tesseract OCR engine is not found or not installed on this system."

        # Build output Markdown
        sections = [f"# Image OCR: {filename}", "", f"![{filename}]({file_path})", ""]

        if extracted_text:
            sections.append("## Extracted Text (OCR)\n")
            sections.append(extracted_text)
            sections.append("\n---")
        elif ocr_error:
            sections.append(f"> **OCR Notice:** Text extraction failed or was skipped: `{ocr_error}`\n")
        else:
            sections.append("*No text detected in this image.*\n")

        # Include standard image metadata and EXIF
        metadata_block = self._get_image_metadata_block(file_path)
        if metadata_block:
            sections.append(metadata_block)

        return "\n".join(sections)

    # ------------------ EXCEL (.xlsx, .xls) CONVERSION ------------------

    def _convert_excel(
            self,
            source: str,
            progress_callback: Optional[Callable[[str], None]] = None
    ) -> str:
        """
        Converts Excel spreadsheets (.xlsx, .xls) into formatted Markdown tables.
        Uses openpyxl for .xlsx and xlrd for .xls as resilient zero-dependency parsers.
        """
        filename = os.path.basename(source)
        _, ext = os.path.splitext(source)
        ext_lower = ext.lower()

        if progress_callback:
            progress_callback(f"Converting Excel spreadsheet: {filename}...")

        # 1. Modern Excel (.xlsx) using openpyxl
        if ext_lower == ".xlsx" and openpyxl is not None:
            try:
                wb = openpyxl.load_workbook(source, data_only=True)
                sections = [f"# Excel Spreadsheet: {filename}\n"]

                for sheet_name in wb.sheetnames:
                    ws = wb[sheet_name]
                    rows = list(ws.iter_rows(values_only=True))
                    # Filter out trailing empty rows
                    valid_rows = [r for r in rows if any(c is not None and str(c).strip() != "" for c in r)]
                    if not valid_rows:
                        continue

                    sections.append(f"## Sheet: {sheet_name}\n")
                    max_cols = max(len(r) for r in valid_rows)
                    if max_cols == 0:
                        continue

                    # Header row
                    header = [str(c).replace("\n", " ").replace("|", "\\|") if c is not None else "" for c in valid_rows[0]]
                    while len(header) < max_cols:
                        header.append("")
                    sections.append("| " + " | ".join(header) + " |")
                    sections.append("| " + " | ".join(["---"] * max_cols) + " |")

                    # Data rows
                    for r in valid_rows[1:]:
                        row_cells = [str(c).replace("\n", " ").replace("|", "\\|") if c is not None else "" for c in r]
                        while len(row_cells) < max_cols:
                            row_cells.append("")
                        sections.append("| " + " | ".join(row_cells) + " |")
                    sections.append("")

                if len(sections) > 1:
                    return "\n".join(sections).strip()
            except Exception as e:
                if progress_callback:
                    progress_callback(f"openpyxl notice: {e}, falling back to alternative parsers...")

        # 2. Legacy Excel (.xls) using xlrd
        if ext_lower == ".xls" and xlrd is not None:
            try:
                wb = xlrd.open_workbook(source)
                sections = [f"# Excel Spreadsheet: {filename}\n"]

                for sheet_name in wb.sheet_names():
                    ws = wb.sheet_by_name(sheet_name)
                    if ws.nrows == 0 or ws.ncols == 0:
                        continue

                    sections.append(f"## Sheet: {sheet_name}\n")
                    header = [str(ws.cell_value(0, col)).replace("\n", " ").replace("|", "\\|") for col in range(ws.ncols)]
                    sections.append("| " + " | ".join(header) + " |")
                    sections.append("| " + " | ".join(["---"] * ws.ncols) + " |")

                    for row in range(1, ws.nrows):
                        row_cells = [str(ws.cell_value(row, col)).replace("\n", " ").replace("|", "\\|") for col in range(ws.ncols)]
                        sections.append("| " + " | ".join(row_cells) + " |")
                    sections.append("")

                if len(sections) > 1:
                    return "\n".join(sections).strip()
            except Exception:
                pass

        # 3. MarkItDown fallback
        try:
            res = self.md.convert(source)
            if res and res.text_content and res.text_content.strip():
                return res.text_content.strip()
        except Exception:
            pass

        return f"# Excel Spreadsheet: {filename}\n\n*(Could not extract tabular data from this spreadsheet)*"

    def _get_image_metadata_block(self, file_path: str) -> str:
        """Generates Markdown metadata lines for an image."""
        if Image is None:
            return ""

        lines = ["### Image Specifications"]
        try:
            from PIL.ExifTags import TAGS
            with Image.open(file_path) as img:
                lines.append(f"- **Format:** {img.format}")
                lines.append(f"- **Dimensions:** {img.width} x {img.height} px")
                lines.append(f"- **Color Mode:** {img.mode}")

                exif = img.getexif()
                if exif:
                    lines.append("\n### EXIF Metadata")
                    for tag_id, val in exif.items():
                        tag_name = TAGS.get(tag_id, str(tag_id))
                        val_str = str(val)
                        if len(val_str) < 150:
                            lines.append(f"- **{tag_name}:** {val_str}")
            return "\n".join(lines)
        except Exception:
            return ""

    # ------------------ LEGACY WORD (.doc) CONVERSION ------------------

    def _convert_doc(
            self,
            source: str,
            progress_callback: Optional[Callable[[str], None]] = None
    ) -> str:
        if progress_callback:
            progress_callback(f"Converting legacy Word document: {os.path.basename(source)}...")

        if sys.platform == "win32":
            res = self._convert_office_doc_com(source, progress_callback)
            if res:
                return res

        res = self._convert_office_via_libreoffice(source, target_ext=".docx")
        if res:
            return res

        if progress_callback:
            progress_callback("Extracting document structure via built-in OLE parser...")
        return self._parse_doc_pure_python(source)

    def _convert_office_doc_com(
            self,
            source: str,
            progress_callback: Optional[Callable[[str], None]] = None
    ) -> Optional[str]:
        temp_docx = os.path.join(tempfile.gettempdir(), f"doc_conv_{uuid.uuid4().hex[:8]}.docx")
        abs_doc = os.path.abspath(source).replace("'", "''")
        abs_docx = os.path.abspath(temp_docx).replace("'", "''")

        ps_script = f"""
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = [int]0
        try {{
            $doc = $word.Documents.Open('{abs_doc}')
            $doc.SaveAs2('{abs_docx}', [int]16)
            $doc.Close([int]0)
        }} finally {{
            $word.Quit([int]0)
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
        }}
        """
        try:
            subprocess.run(
                ["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", ps_script],
                capture_output=True,
                text=True,
                timeout=35
            )
            if os.path.exists(temp_docx):
                if progress_callback:
                    progress_callback("Parsing converted Word document into Markdown...")
                md_result = self.md.convert(temp_docx)
                return md_result.text_content.strip() if md_result and md_result.text_content else None
        except Exception:
            pass
        finally:
            if os.path.exists(temp_docx):
                try:
                    os.remove(temp_docx)
                except OSError:
                    pass
        return None

    def _parse_doc_pure_python(self, file_path: str) -> str:
        filename = os.path.basename(file_path)

        try:
            with open(file_path, "rb") as f:
                header = f.read(16)
                if header.startswith(b"{\\rtf"):
                    f.seek(0)
                    rtf_raw = f.read().decode("latin1", errors="ignore")
                    text = self._parse_rtf(rtf_raw)
                    return f"# {filename}\n\n{text}"
        except Exception:
            pass

        if olefile is None or not olefile.isOleFile(file_path):
            with open(file_path, "rb") as f:
                raw_bytes = f.read()
            return f"# {filename}\n\n" + self._extract_printable_strings(raw_bytes)

        ole = olefile.OleFileIO(file_path)
        try:
            if not ole.exists("WordDocument"):
                with open(file_path, "rb") as f:
                    raw_bytes = f.read()
                return f"# {filename}\n\n" + self._extract_printable_strings(raw_bytes)

            wd = ole.openstream("WordDocument").read()
            if len(wd) < 0x01A6:
                return f"# {filename}\n\n" + self._extract_printable_strings(wd)

            flags = struct.unpack_from("<H", wd, 0x000A)[0]
            table_name = "1Table" if (flags & 0x0200) else "0Table"

            if not ole.exists(table_name):
                table_name = "0Table" if table_name == "1Table" else "1Table"
                if not ole.exists(table_name):
                    return f"# {filename}\n\n" + self._extract_printable_strings(wd)

            tbl = ole.openstream(table_name).read()
            fcClx, lcbClx = struct.unpack_from("<II", wd, 0x01A2)
            if fcClx + lcbClx > len(tbl):
                return f"# {filename}\n\n" + self._extract_printable_strings(wd)

            clx = tbl[fcClx : fcClx + lcbClx]
            pos = 0
            extracted_parts = []

            while pos < len(clx):
                clx_type = clx[pos]
                pos += 1
                if clx_type == 1:
                    cb = struct.unpack_from("<H", clx, pos)[0]
                    pos += 2 + cb
                elif clx_type == 2:
                    lcb = struct.unpack_from("<I", clx, pos)[0]
                    pos += 4
                    if lcb < 4 or (lcb - 4) % 12 != 0:
                        break
                    n = (lcb - 4) // 12
                    cps = [struct.unpack_from("<I", clx, pos + i * 4)[0] for i in range(n + 1)]
                    pcds_offset = pos + (n + 1) * 4

                    for i in range(n):
                        pcd_offset = pcds_offset + i * 8
                        fc = struct.unpack_from("<I", clx, pcd_offset + 2)[0]
                        is_8bit = bool(fc & 0x40000000)
                        actual_fc = (fc & ~0x40000000) // 2 if is_8bit else fc
                        char_count = cps[i + 1] - cps[i]

                        if is_8bit:
                            raw = wd[actual_fc : actual_fc + char_count]
                            extracted_parts.append(raw.decode("cp1252", errors="replace"))
                        else:
                            raw = wd[actual_fc : actual_fc + char_count * 2]
                            extracted_parts.append(raw.decode("utf-16le", errors="replace"))
                    break
                else:
                    break

            full_text = "".join(extracted_parts)
            if not full_text.strip():
                full_text = self._extract_printable_strings(wd)

            clean_text = self._clean_doc_text(full_text)
            return f"# {filename}\n\n{clean_text}"
        except Exception:
            with open(file_path, "rb") as f:
                raw_bytes = f.read()
            return f"# {filename}\n\n" + self._extract_printable_strings(raw_bytes)
        finally:
            ole.close()

    # ------------------ LEGACY POWERPOINT (.ppt) CONVERSION ------------------

    def _convert_ppt(
            self,
            source: str,
            progress_callback: Optional[Callable[[str], None]] = None
    ) -> str:
        if progress_callback:
            progress_callback(f"Converting legacy PowerPoint presentation: {os.path.basename(source)}...")

        if sys.platform == "win32":
            res = self._convert_office_ppt_com(source, progress_callback)
            if res:
                return res

        res = self._convert_office_via_libreoffice(source, target_ext=".pptx")
        if res:
            return res

        if progress_callback:
            progress_callback("Extracting slides via built-in presentation parser...")
        return self._parse_ppt_pure_python(source)

    def _convert_office_ppt_com(
            self,
            source: str,
            progress_callback: Optional[Callable[[str], None]] = None
    ) -> Optional[str]:
        temp_pptx = os.path.join(tempfile.gettempdir(), f"ppt_conv_{uuid.uuid4().hex[:8]}.pptx")
        abs_ppt = os.path.abspath(source).replace("'", "''")
        abs_pptx = os.path.abspath(temp_pptx).replace("'", "''")

        ps_script = f"""
        $ppt = New-Object -ComObject PowerPoint.Application
        $ppt.DisplayAlerts = [int]1
        try {{
            $pres = $ppt.Presentations.Open('{abs_ppt}', [Microsoft.Office.Core.MsoTriState]::msoTrue, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)
            $pres.SaveAs('{abs_pptx}', [int]24)
            $pres.Close()
        }} finally {{
            $ppt.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
        }}
        """
        try:
            subprocess.run(
                ["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", ps_script],
                capture_output=True,
                text=True,
                timeout=35
            )
            if os.path.exists(temp_pptx):
                if progress_callback:
                    progress_callback("Parsing converted presentation into Markdown...")
                md_result = self.md.convert(temp_pptx)
                return md_result.text_content.strip() if md_result and md_result.text_content else None
        except Exception:
            pass
        finally:
            if os.path.exists(temp_pptx):
                try:
                    os.remove(temp_pptx)
                except OSError:
                    pass
        return None

    def _parse_ppt_pure_python(self, file_path: str) -> str:
        filename = os.path.basename(file_path)

        if olefile is None or not olefile.isOleFile(file_path):
            with open(file_path, "rb") as f:
                raw_bytes = f.read()
            return f"# {filename}\n\n" + self._extract_printable_strings(raw_bytes)

        ole = olefile.OleFileIO(file_path)
        try:
            if not ole.exists("PowerPoint Document"):
                with open(file_path, "rb") as f:
                    raw_bytes = f.read()
                return f"# {filename}\n\n" + self._extract_printable_strings(raw_bytes)

            data = ole.openstream("PowerPoint Document").read()
            slides_text = []
            current_slide_texts = []
            i = 0
            data_len = len(data)

            while i < data_len - 8:
                rec_type = struct.unpack_from("<H", data, i + 2)[0]
                rec_len = struct.unpack_from("<I", data, i + 4)[0]

                if rec_type == 1000:
                    if current_slide_texts:
                        slides_text.append(list(current_slide_texts))
                        current_slide_texts.clear()

                elif rec_type in (4008, 4000) and 0 < rec_len <= 100000 and i + 8 + rec_len <= data_len:
                    enc = "latin1" if rec_type == 4008 else "utf-16le"
                    try:
                        s = data[i + 8 : i + 8 + rec_len].decode(enc, errors="ignore").strip()
                        if s and s not in {"*", "-", "•"} and not s.startswith("___PPT") and not s.startswith("Click to edit"):
                            current_slide_texts.append(s)
                    except Exception:
                        pass

                i += 1

            if current_slide_texts:
                slides_text.append(list(current_slide_texts))

            md_output = []
            slide_num = 1
            for slide_items in slides_text:
                unique_items = [
                    s for s in dict.fromkeys(slide_items)
                    if s not in {"*", "-", "•", ""} and len(s) > 1 and not s.startswith("Click to edit")
                ]
                if not unique_items:
                    continue
                md_output.append(f"<!-- Slide number: {slide_num} -->")
                md_output.append(f"## {unique_items[0]}")
                for body_item in unique_items[1:]:
                    for line in body_item.split("\n"):
                        clean = line.strip()
                        if clean:
                            md_output.append(f"- {clean}" if not clean.startswith("-") else clean)
                md_output.append("")
                slide_num += 1

            if not md_output:
                text_items = []
                j = 0
                while j < data_len - 8:
                    rec_type = struct.unpack_from("<H", data, j + 2)[0]
                    rec_len = struct.unpack_from("<I", data, j + 4)[0]
                    if rec_type in (4008, 4000) and 0 < rec_len <= 100000 and j + 8 + rec_len <= data_len:
                        enc = "latin1" if rec_type == 4008 else "utf-16le"
                        try:
                            s = data[j + 8 : j + 8 + rec_len].decode(enc, errors="ignore").strip()
                            if s and s not in {"*", "-", "•"} and not s.startswith("___PPT") and not s.startswith("Click to edit"):
                                text_items.append(s)
                        except Exception:
                            pass
                    j += 1
                unique = list(dict.fromkeys(text_items))
                if unique:
                    md_output.append(f"# {filename}\n")
                    for item in unique:
                        md_output.append(f"- {item}")
                    return "\n".join(md_output)
                return f"# {filename}\n\n" + self._extract_printable_strings(data)

            return "\n".join(md_output).strip()
        except Exception:
            with open(file_path, "rb") as f:
                raw_bytes = f.read()
            return f"# {filename}\n\n" + self._extract_printable_strings(raw_bytes)
        finally:
            ole.close()

    # ------------------ CROSS-PLATFORM HELPER UTILITIES ------------------

    def _convert_office_via_libreoffice(self, source: str, target_ext: str) -> Optional[str]:
        soffice_bin = shutil.which("libreoffice") or shutil.which("soffice")
        if not soffice_bin:
            return None

        temp_dir = tempfile.mkdtemp()
        try:
            fmt = target_ext.lstrip(".")
            cmd = [
                soffice_bin,
                "--headless",
                "--convert-to",
                fmt,
                "--outdir",
                temp_dir,
                source
            ]
            subprocess.run(cmd, capture_output=True, text=True, timeout=30)
            base_name = os.path.splitext(os.path.basename(source))[0]
            converted_file = os.path.join(temp_dir, f"{base_name}.{fmt}")
            if os.path.exists(converted_file):
                md_res = self.md.convert(converted_file)
                return md_res.text_content.strip() if md_res and md_res.text_content else None
        except Exception:
            pass
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)
        return None

    def _clean_doc_text(self, text: str) -> str:
        text = text.replace('\r\n', '\n').replace('\r', '\n')
        text = text.replace('\x0b', '\n').replace('\x0c', '\n\n---\n\n').replace('\x07', ' | ')
        lines = [line.strip() for line in text.split('\n')]
        md_lines = []
        for line in lines:
            if not line:
                if md_lines and md_lines[-1] != "":
                    md_lines.append("")
                continue
            if line.startswith('\xb7') or line.startswith('\u2022') or line.startswith('- '):
                clean_item = re.sub(r'^[\xb7\u2022\-\*\s]+', '', line).strip()
                md_lines.append(f"- {clean_item}")
            else:
                md_lines.append(line)
        return '\n'.join(md_lines).strip()

    def _parse_rtf(self, rtf_text: str) -> str:
        text = re.sub(r'{\\[^{}]+}', '', rtf_text)
        text = re.sub(r'\\[a-zA-Z]+\-?\d*[ ]?', '', text)
        text = text.replace('{', '').replace('}', '').strip()
        return text

    def _extract_printable_strings(self, data: bytes) -> str:
        ascii_strings = [s.decode('latin1') for s in re.findall(rb'[A-Za-z0-9\s.,;:\'\-\(\)]{4,}', data)]
        utf16_strings = [s.decode('utf-16le', errors='ignore') for s in re.findall(rb'(?:[\x20-\x7e]\x00){4,}', data)]
        all_strings = list(dict.fromkeys(ascii_strings + utf16_strings))
        filtered = [s.strip() for s in all_strings if len(s.strip()) > 3]
        return '\n\n'.join(filtered) if filtered else "*(No readable text could be extracted from this file)*"

    def _fallback_text_convert(self, file_path: str, ext: str) -> str:
        try:
            with open(file_path, "r", encoding="utf-8", errors="replace") as f:
                raw_text = f.read()
        except Exception as e:
            return f"# {os.path.basename(file_path)}\n\n*Error reading text file: {e}*"

        filename = os.path.basename(file_path)
        lang = ext.lstrip(".")
        if lang in {"json", "xml", "csv"}:
            return f"# {filename}\n\n```{lang}\n{raw_text}\n```"
        return f"# {filename}\n\n{raw_text}"


def convert_and_save(
        source: str,
        output_dir: str,
        output_filename: str,
        progress_callback: Optional[Callable[[str], None]] = None
) -> Tuple[bool, str, str]:
    try:
        output_dir = output_dir.strip()
        if not output_dir:
            raise ValueError("Storage folder path cannot be empty.")

        os.makedirs(output_dir, exist_ok=True)

        output_filename = output_filename.strip()
        if not output_filename:
            if "://" in source:
                output_filename = "web_converted.md"
            else:
                base = os.path.splitext(os.path.basename(source))[0]
                output_filename = f"{base}.md"

        if not output_filename.lower().endswith(".md"):
            output_filename += ".md"

        full_output_path = os.path.abspath(os.path.join(output_dir, output_filename))

        converter = MarkdownConverter()
        markdown_text = converter.convert_source(source, progress_callback=progress_callback)

        if progress_callback:
            progress_callback("Writing Markdown file to destination...")

        with open(full_output_path, "w", encoding="utf-8") as out_f:
            out_f.write(markdown_text)

        return True, full_output_path, f"Successfully converted and saved to:\n{full_output_path}"

    except Exception as e:
        return False, "", f"Conversion failed: {str(e)}"