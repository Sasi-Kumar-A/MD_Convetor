"""
Markdown Converter Desktop Application
A modern GUI desktop application built with Python and CustomTkinter to convert
PDF, Word, Excel, PowerPoint, Images, Audio, Video, HTML, CSV/JSON/XML, ZIP, EPubs,
and Web/YouTube URLs into clean Markdown (.md) documents.
"""

import os
import sys
import threading
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk

from converter import MarkdownConverter, convert_and_save

# Configure default appearance
ctk.set_appearance_mode("Dark")  # "System", "Dark", "Light"
ctk.set_default_color_theme("blue")  # "blue", "green", "dark-blue"


class MarkdownConverterApp(ctk.CTk):
    """Main Desktop Application Window."""

    FORMAT_OPTIONS = {
        "All Supported Formats (*.*)": [
            ("All Supported Files", "*.pdf;*.docx;*.pptx;*.xlsx;*.xls;*.jpg;*.jpeg;*.png;*.webp;*.mp3;*.wav;*.mp4;*.html;*.htm;*.txt;*.csv;*.json;*.xml;*.zip;*.epub"),
            ("All Files", "*.*")
        ],
        "PDF Document (*.pdf)": [("PDF Files", "*.pdf"), ("All Files", "*.*")],
        "Word Document (*.docx, *.doc)": [("Word Documents", "*.docx;*.doc"), ("All Files", "*.*")],
        "PowerPoint Presentation (*.pptx, *.ppt)": [("PowerPoint Presentations", "*.pptx;*.ppt"), ("All Files", "*.*")],
        "Excel Spreadsheet (*.xlsx, *.xls)": [("Excel Files", "*.xlsx;*.xls"), ("All Files", "*.*")],
        "Images - OCR & EXIF (*.jpg, *.png, *.webp)": [("Image Files", "*.jpg;*.jpeg;*.png;*.webp;*.bmp;*.tiff"), ("All Files", "*.*")],
        "Audio - Transcription (*.mp3, *.wav)": [("Audio Files", "*.mp3;*.wav;*.m4a;*.flac;*.aac;*.ogg"), ("All Files", "*.*")],
        "Video - Audio Extract (*.mp4, *.mkv)": [("Video Files", "*.mp4;*.mkv;*.mov;*.avi;*.wmv;*.flv"), ("All Files", "*.*")],
        "HTML Webpage (*.html, *.htm)": [("HTML Files", "*.html;*.htm"), ("All Files", "*.*")],
        "Text & Data (*.csv, *.json, *.xml, *.txt)": [("Data Files", "*.csv;*.json;*.xml;*.txt;*.tsv"), ("All Files", "*.*")],
        "ZIP Archive (*.zip)": [("ZIP Archives", "*.zip"), ("All Files", "*.*")],
        "YouTube / Web URL (Online Link)": [("All Files", "*.*")],
        "EPub E-Book (*.epub)": [("EPub Files", "*.epub"), ("All Files", "*.*")],
    }

    def __init__(self):
        super().__init__()

        self.title("MD Converter Studio - Universal Markdown Converter")
        self.geometry("900, 780")
        self.minsize(780, 680)

        # Application state
        self.is_converting = False
        self.last_converted_file = ""
        self.last_output_dir = ""

        # Default paths
        self.default_store_dir = os.path.join(os.path.expanduser("~"), "Documents")
        if not os.path.exists(self.default_store_dir):
            self.default_store_dir = os.path.expanduser("~")

        self._build_ui()

    def _build_ui(self):
        """Constructs the application layout."""
        # Main Grid Layout configuration
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # ----------------- 1. TOP NAVIGATION BAR -----------------
        self.nav_bar = ctk.CTkFrame(self, height=65, corner_radius=0, fg_color=("gray90", "#181b20"))
        self.nav_bar.grid(row=0, column=0, sticky="ew", padx=0, pady=0)
        self.nav_bar.grid_columnconfigure(0, weight=1)

        nav_left = ctk.CTkFrame(self.nav_bar, fg_color="transparent")
        nav_left.grid(row=0, column=0, sticky="w", padx=24, pady=12)

        app_title = ctk.CTkLabel(
            nav_left,
            text="📑 MD Converter Studio",
            font=ctk.CTkFont(size=20, weight="bold")
        )
        app_title.pack(side="left", padx=(0, 10))

        app_badge = ctk.CTkLabel(
            nav_left,
            text="v1.0 • Universal .md Engine",
            font=ctk.CTkFont(size=11),
            text_color=("gray50", "gray60")
        )
        app_badge.pack(side="left")

        # Theme toggle on right of Nav Bar
        nav_right = ctk.CTkFrame(self.nav_bar, fg_color="transparent")
        nav_right.grid(row=0, column=1, sticky="e", padx=24, pady=12)

        self.theme_switch = ctk.CTkOptionMenu(
            nav_right,
            values=["Dark Mode", "Light Mode", "System"],
            command=self._change_theme,
            width=120,
            height=30
        )
        self.theme_switch.set("Dark Mode")
        self.theme_switch.pack(side="right")

        # ----------------- MAIN SCROLLABLE CONTAINER -----------------
        self.content_frame = ctk.CTkScrollableFrame(self, fg_color="transparent")
        self.content_frame.grid(row=1, column=0, sticky="nsew", padx=28, pady=20)
        self.content_frame.grid_columnconfigure(0, weight=1)

        # Card Container for Form Fields
        self.card = ctk.CTkFrame(self.content_frame, corner_radius=14, fg_color=("white", "#1e2229"), border_width=1, border_color=("gray80", "#2c313a"))
        self.card.grid(row=0, column=0, sticky="ew", padx=4, pady=4)
        self.card.grid_columnconfigure(0, weight=1)

        card_padding_x = 24

        # ----------------- FIELD 1: FORMAT DROPDOWN -----------------
        lbl_dropdown = ctk.CTkLabel(
            self.card,
            text="1. Select File Format / Type:",
            font=ctk.CTkFont(size=14, weight="bold"),
            anchor="w"
        )
        lbl_dropdown.pack(fill="x", padx=card_padding_x, pady=(20, 6))

        format_names = list(self.FORMAT_OPTIONS.keys())
        self.format_dropdown = ctk.CTkOptionMenu(
            self.card,
            values=format_names,
            command=self._on_format_changed,
            height=38,
            font=ctk.CTkFont(size=13),
            dropdown_font=ctk.CTkFont(size=12)
        )
        self.format_dropdown.set(format_names[0])
        self.format_dropdown.pack(fill="x", padx=card_padding_x, pady=(0, 16))

        # ----------------- FIELD 2: BROWSE FILE & PATH / URL BOX -----------------
        lbl_file = ctk.CTkLabel(
            self.card,
            text="2. Select File or Paste Web / YouTube Link:",
            font=ctk.CTkFont(size=14, weight="bold"),
            anchor="w"
        )
        lbl_file.pack(fill="x", padx=card_padding_x, pady=(0, 6))

        file_row = ctk.CTkFrame(self.card, fg_color="transparent")
        file_row.pack(fill="x", padx=card_padding_x, pady=(0, 16))
        file_row.grid_columnconfigure(1, weight=1)

        self.btn_browse_file = ctk.CTkButton(
            file_row,
            text="📂 Browse File",
            command=self._browse_file,
            width=135,
            height=38,
            font=ctk.CTkFont(size=13, weight="bold")
        )
        self.btn_browse_file.grid(row=0, column=0, padx=(0, 10), sticky="w")

        self.entry_source = ctk.CTkEntry(
            file_row,
            placeholder_text="Path of selected file, or paste URL (e.g. https://www.youtube.com/...)",
            height=38,
            font=ctk.CTkFont(size=13)
        )
        self.entry_source.grid(row=0, column=1, sticky="ew")

        # ----------------- FIELD 3: OUTPUT FILENAME (.md) -----------------
        lbl_filename = ctk.CTkLabel(
            self.card,
            text="3. Output File Name (will save as .md):",
            font=ctk.CTkFont(size=14, weight="bold"),
            anchor="w"
        )
        lbl_filename.pack(fill="x", padx=card_padding_x, pady=(0, 6))

        self.entry_filename = ctk.CTkEntry(
            self.card,
            placeholder_text="e.g. my_document.md",
            height=38,
            font=ctk.CTkFont(size=13)
        )
        self.entry_filename.pack(fill="x", padx=card_padding_x, pady=(0, 16))

        # ----------------- FIELD 4: STORE PATH (DESTINATION FOLDER) -----------------
        lbl_store = ctk.CTkLabel(
            self.card,
            text="4. Storage Folder (where to store the .md file):",
            font=ctk.CTkFont(size=14, weight="bold"),
            anchor="w"
        )
        lbl_store.pack(fill="x", padx=card_padding_x, pady=(0, 6))

        store_row = ctk.CTkFrame(self.card, fg_color="transparent")
        store_row.pack(fill="x", padx=card_padding_x, pady=(0, 22))
        store_row.grid_columnconfigure(1, weight=1)

        self.btn_browse_folder = ctk.CTkButton(
            store_row,
            text="📁 Browse Folder",
            command=self._browse_folder,
            width=135,
            height=38,
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color=("gray75", "#3a414e"),
            hover_color=("gray65", "#485162")
        )
        self.btn_browse_folder.grid(row=0, column=0, padx=(0, 10), sticky="w")

        self.entry_store_dir = ctk.CTkEntry(
            store_row,
            placeholder_text="Select or type the destination folder path...",
            height=38,
            font=ctk.CTkFont(size=13)
        )
        self.entry_store_dir.insert(0, self.default_store_dir)
        self.entry_store_dir.grid(row=0, column=1, sticky="ew")

        # ----------------- FIELD 5: CONVERT & CLEAR BUTTONS -----------------
        btn_action_row = ctk.CTkFrame(self.card, fg_color="transparent")
        btn_action_row.pack(fill="x", padx=card_padding_x, pady=(4, 24))
        btn_action_row.grid_columnconfigure((0, 1), weight=1)

        self.btn_convert = ctk.CTkButton(
            btn_action_row,
            text="⚡ Convert to .md",
            command=self._start_conversion,
            height=44,
            font=ctk.CTkFont(size=15, weight="bold"),
            fg_color=("#1f6aa5", "#2563eb"),
            hover_color=("#185584", "#1d4ed8")
        )
        self.btn_convert.grid(row=0, column=0, padx=(0, 10), sticky="ew")

        self.btn_clear = ctk.CTkButton(
            btn_action_row,
            text="🔄 Clear All Fields",
            command=self._clear_all_fields,
            height=44,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color=("gray80", "#2c313a"),
            hover_color=("gray70", "#3a414e"),
            text_color=("gray10", "white")
        )
        self.btn_clear.grid(row=0, column=1, padx=(10, 0), sticky="ew")

        # ----------------- 6. LIVE PROGRESS & STATUS NOTIFICATION AREA -----------------
        self.status_card = ctk.CTkFrame(self.content_frame, corner_radius=12, fg_color=("gray95", "#16191f"), border_width=1, border_color=("gray85", "#242933"))
        self.status_card.grid(row=1, column=0, sticky="ew", padx=4, pady=(16, 10))
        self.status_card.grid_columnconfigure(0, weight=1)

        # Progress bar (Indeterminate during conversion)
        self.progress_bar = ctk.CTkProgressBar(self.status_card, height=6, corner_radius=3)
        self.progress_bar.set(0)
        self.progress_bar.pack(fill="x", padx=20, pady=(16, 8))

        # Status Label
        self.lbl_status = ctk.CTkLabel(
            self.status_card,
            text="Ready. Choose a file or paste a link and click 'Convert to .md'.",
            font=ctk.CTkFont(size=13),
            wraplength=760,
            justify="center",
            text_color=("gray40", "gray65")
        )
        self.lbl_status.pack(fill="x", padx=20, pady=(4, 12))

        # Action shortcuts (Open Folder, Open File)
        self.shortcuts_row = ctk.CTkFrame(self.status_card, fg_color="transparent")
        self.shortcuts_row.pack(pady=(0, 16))

        self.btn_open_file = ctk.CTkButton(
            self.shortcuts_row,
            text="📄 Open Markdown File",
            command=self._open_file,
            width=160,
            height=32,
            font=ctk.CTkFont(size=12, weight="bold"),
            fg_color=("#10b981", "#059669"),
            hover_color=("#059669", "#047857")
        )

        self.btn_open_folder = ctk.CTkButton(
            self.shortcuts_row,
            text="📂 Open Stored Folder",
            command=self._open_folder,
            width=160,
            height=32,
            font=ctk.CTkFont(size=12, weight="bold"),
            fg_color=("gray75", "#374151"),
            hover_color=("gray65", "#4b5563")
        )

    # ----------------- EVENT HANDLERS & LOGIC -----------------

    def _on_format_changed(self, choice: str):
        """Updates placeholder and hint based on dropdown selection."""
        if "YouTube" in choice or "URL" in choice:
            self.entry_source.configure(placeholder_text="Paste YouTube URL or Webpage Link here...")
        else:
            self.entry_source.configure(placeholder_text="Path of selected file, or paste URL...")

    def _browse_file(self):
        """Opens native file picker according to selected format."""
        selected_fmt = self.format_dropdown.get()
        filetypes = self.FORMAT_OPTIONS.get(selected_fmt, [("All Files", "*.*")])

        chosen_path = filedialog.askopenfilename(
            title=f"Select File ({selected_fmt})",
            filetypes=filetypes
        )

        if chosen_path:
            chosen_path = os.path.normpath(chosen_path)
            self.entry_source.delete(0, tk.END)
            self.entry_source.insert(0, chosen_path)

            # Auto-suggest filename if filename box is empty or was auto-filled
            current_name = self.entry_filename.get().strip()
            base_name = os.path.splitext(os.path.basename(chosen_path))[0]
            if not current_name or current_name.endswith(".md"):
                self.entry_filename.delete(0, tk.END)
                self.entry_filename.insert(0, f"{base_name}.md")

            # Suggest storing in the same directory as input file if user hasn't modified store dir
            file_dir = os.path.dirname(chosen_path)
            if file_dir and os.path.isdir(file_dir):
                self.entry_store_dir.delete(0, tk.END)
                self.entry_store_dir.insert(0, file_dir)

    def _browse_folder(self):
        """Opens folder picker to select destination directory."""
        current_dir = self.entry_store_dir.get().strip() or self.default_store_dir
        if not os.path.exists(current_dir):
            current_dir = self.default_store_dir

        chosen_dir = filedialog.askdirectory(
            title="Select Storage Folder for Markdown File",
            initialdir=current_dir
        )

        if chosen_dir:
            chosen_dir = os.path.normpath(chosen_dir)
            self.entry_store_dir.delete(0, tk.END)
            self.entry_store_dir.insert(0, chosen_dir)

    def _clear_all_fields(self):
        """Resets all fields, dropdown, and status as requested."""
        if self.is_converting:
            messagebox.showwarning("Busy", "A conversion is currently in progress.")
            return

        # 1. Reset dropdown
        first_format = list(self.FORMAT_OPTIONS.keys())[0]
        self.format_dropdown.set(first_format)
        self._on_format_changed(first_format)

        # 2. Clear selected file / URL
        self.entry_source.delete(0, tk.END)

        # 3. Clear output filename
        self.entry_filename.delete(0, tk.END)

        # 4. Clear storage folder (reset to default)
        self.entry_store_dir.delete(0, tk.END)
        self.entry_store_dir.insert(0, self.default_store_dir)

        # 5. Clear status message & hide action buttons
        self.lbl_status.configure(
            text="All fields cleared. Ready for new file.",
            text_color=("gray40", "gray65")
        )
        self.progress_bar.set(0)
        self.btn_open_file.pack_forget()
        self.btn_open_folder.pack_forget()
        self.last_converted_file = ""
        self.last_output_dir = ""

    def _start_conversion(self):
        """Validates inputs and triggers background conversion thread."""
        if self.is_converting:
            return

        source = self.entry_source.get().strip()
        filename = self.entry_filename.get().strip()
        store_dir = self.entry_store_dir.get().strip()

        # Validation
        if not source:
            messagebox.showwarning(
                "Missing File or URL",
                "Please select a file using 'Browse File' or paste a valid web/YouTube URL."
            )
            return

        if not filename:
            # Auto-generate if empty
            if "://" in source:
                filename = "web_converted.md"
            else:
                base = os.path.splitext(os.path.basename(source))[0]
                filename = f"{base}.md"
            self.entry_filename.delete(0, tk.END)
            self.entry_filename.insert(0, filename)

        if not store_dir:
            store_dir = self.default_store_dir
            self.entry_store_dir.delete(0, tk.END)
            self.entry_store_dir.insert(0, store_dir)

        # Set UI state to converting
        self.is_converting = True
        self.btn_convert.configure(state="disabled", text="⏳ Converting...")
        self.btn_clear.configure(state="disabled")
        self.btn_browse_file.configure(state="disabled")
        self.btn_browse_folder.configure(state="disabled")

        self.btn_open_file.pack_forget()
        self.btn_open_folder.pack_forget()

        self.progress_bar.configure(mode="indeterminate")
        self.progress_bar.start()

        self.lbl_status.configure(
            text="Converting file to Markdown... Please wait.",
            text_color=("#2563eb", "#60a5fa")
        )

        # Launch worker thread so GUI stays fluid
        worker = threading.Thread(
            target=self._run_conversion_worker,
            args=(source, store_dir, filename),
            daemon=True
        )
        worker.start()

    def _run_conversion_worker(self, source: str, store_dir: str, filename: str):
        """Worker thread function executing conversion."""
        def progress_update(msg: str):
            self.after(0, lambda: self.lbl_status.configure(text=msg))

        success, full_path, message = convert_and_save(
            source=source,
            output_dir=store_dir,
            output_filename=filename,
            progress_callback=progress_update
        )

        # Update GUI on main thread
        self.after(0, lambda: self._on_conversion_finished(success, full_path, message, store_dir))

    def _on_conversion_finished(self, success: bool, full_path: str, message: str, store_dir: str):
        """Updates UI after conversion finishes."""
        self.is_converting = False
        self.progress_bar.stop()
        self.progress_bar.configure(mode="determinate")

        self.btn_convert.configure(state="normal", text="⚡ Convert to .md")
        self.btn_clear.configure(state="normal")
        self.btn_browse_file.configure(state="normal")
        self.btn_browse_folder.configure(state="normal")

        if success:
            self.progress_bar.set(1.0)
            self.last_converted_file = full_path
            self.last_output_dir = store_dir

            # User requested message: "converted and store in the respective folder"
            folder_name = os.path.basename(os.path.normpath(store_dir)) or store_dir
            file_name = os.path.basename(full_path)
            self.lbl_status.configure(
                text=f"✓ Converted successfully and stored in the respective folder:\n📁 {full_path}",
                text_color=("#059669", "#34d399")
            )

            # Show quick access buttons
            self.btn_open_file.pack(side="left", padx=8)
            self.btn_open_folder.pack(side="left", padx=8)

        else:
            self.progress_bar.set(0)
            self.lbl_status.configure(
                text=f"⚠ {message}",
                text_color=("#dc2626", "#f87171")
            )
            messagebox.showerror("Conversion Error", message)

    def _open_file(self):
        """Opens the converted Markdown file in system default editor."""
        if self.last_converted_file and os.path.exists(self.last_converted_file):
            try:
                os.startfile(self.last_converted_file)
            except Exception as e:
                messagebox.showerror("Error", f"Could not open file: {e}")

    def _open_folder(self):
        """Opens the folder where the converted .md file is saved."""
        target_dir = self.last_output_dir or self.entry_store_dir.get().strip()
        if target_dir and os.path.exists(target_dir):
            try:
                os.startfile(target_dir)
            except Exception as e:
                messagebox.showerror("Error", f"Could not open directory: {e}")

    def _change_theme(self, choice: str):
        """Changes theme between Dark, Light, and System."""
        mode_map = {
            "Dark Mode": "Dark",
            "Light Mode": "Light",
            "System": "System"
        }
        ctk.set_appearance_mode(mode_map.get(choice, "Dark"))


def main():
    app = MarkdownConverterApp()
    app.mainloop()


if __name__ == "__main__":
    main()
