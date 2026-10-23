"""
Markdown Converter Engine
Handles conversion of various file formats and URLs to Markdown (.md)
using Microsoft's MarkItDown and complementary tools (MoviePy, OCR, etc.).
"""

import os
import sys
import tempfile
import uuid
from typing import Callable, Optional, Tuple
from markitdown import MarkItDown

# Video conversion helper
try:
    from moviepy import VideoFileClip
except ImportError:
    try:
        from moviepy.editor import VideoFileClip
    except ImportError:
        VideoFileClip = None


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

        # Standard conversion via MarkItDown
        if progress_callback:
            progress_callback(f"Converting {os.path.basename(source)} to Markdown...")

        try:
            result = self.md.convert(source)
            content = result.text_content.strip() if result and result.text_content else ""
        except Exception as e:
            # Fallback for plain text, JSON, CSV if specific parser fails
            if ext_lower in {".txt", ".csv", ".json", ".xml", ".tsv"}:
                content = self._fallback_text_convert(source, ext_lower)
            else:
                raise e

        # If empty content returned for image or file, provide helpful fallback
        if not content:
            if ext_lower in self.IMAGE_EXTENSIONS:
                content = self._fallback_image_metadata(source)
            elif ext_lower in {".txt", ".csv", ".json", ".xml", ".tsv"}:
                content = self._fallback_text_convert(source, ext_lower)
            else:
                content = f"# {os.path.basename(source)}\n\n*(Conversion finished, but no text or captions were found in the file)*"

        return content

    def _fallback_text_convert(self, file_path: str, ext: str) -> str:
        """Fallback reader for text-based formats."""
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

    def _fallback_image_metadata(self, file_path: str) -> str:
        """Generates markdown for images with dimensions and EXIF metadata."""
        filename = os.path.basename(file_path)
        lines = [f"# Image: {filename}", "", f"![{filename}]({file_path})", ""]
        try:
            from PIL import Image
            from PIL.ExifTags import TAGS

            with Image.open(file_path) as img:
                lines.append("### Image Details")
                lines.append(f"- **Format:** {img.format}")
                lines.append(f"- **Dimensions:** {img.width} x {img.height} px")
                lines.append(f"- **Mode:** {img.mode}")

                exif = img.getexif()
                if exif:
                    lines.append("\n### EXIF Metadata")
                    for tag_id, val in exif.items():
                        tag_name = TAGS.get(tag_id, str(tag_id))
                        # Limit lengthy binary values
                        val_str = str(val)
                        if len(val_str) < 150:
                            lines.append(f"- **{tag_name}:** {val_str}")
        except Exception:
            pass

        return "\n".join(lines)


def convert_and_save(
    source: str,
    output_dir: str,
    output_filename: str,
    progress_callback: Optional[Callable[[str], None]] = None
) -> Tuple[bool, str, str]:
    """
    Executes the conversion and writes the output file.

    Args:
        source: Input file path or URL
        output_dir: Destination folder path
        output_filename: Desired output filename (will ensure .md extension)
        progress_callback: Callback function for status updates

    Returns:
        (success: bool, full_saved_path: str, message: str)
    """
    try:
        # Validate output directory
        output_dir = output_dir.strip()
        if not output_dir:
            raise ValueError("Storage folder path cannot be empty.")

        os.makedirs(output_dir, exist_ok=True)

        # Validate filename
        output_filename = output_filename.strip()
        if not output_filename:
            # Generate default filename from source
            if "://" in source:
                output_filename = "web_converted.md"
            else:
                base = os.path.splitext(os.path.basename(source))[0]
                output_filename = f"{base}.md"

        if not output_filename.lower().endswith(".md"):
            output_filename += ".md"

        full_output_path = os.path.abspath(os.path.join(output_dir, output_filename))

        # Perform conversion
        converter = MarkdownConverter()
        markdown_text = converter.convert_source(source, progress_callback=progress_callback)

        if progress_callback:
            progress_callback("Writing Markdown file to destination...")

        with open(full_output_path, "w", encoding="utf-8") as out_f:
            out_f.write(markdown_text)

        return True, full_output_path, f"Successfully converted and saved to:\n{full_output_path}"

    except Exception as e:
        return False, "", f"Conversion failed: {str(e)}"
