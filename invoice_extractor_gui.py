#!/usr/bin/env python3
"""
Invoice Extractor - GUI Application
Connects to Gmail, downloads invoice attachments, parses them,
and exports extracted data to an Excel spreadsheet.
"""

import os
import sys
import json
import csv
import re
import threading
import tkinter as tk
from tkinter import ttk
import tkinter.font as tkfont
import time
from datetime import datetime
try:
    from PIL import Image, ImageTk
    PIL_AVAILABLE = True
except Exception:
    PIL_AVAILABLE = False

from gmail_client import GmailClient
from invoice_parser import parse_invoice, OCR_AVAILABLE
from spreadsheet_writer import (
    write_invoice_to_spreadsheet, read_spreadsheet_rows,
    write_validation_result, get_unique_po_numbers
)
from skunexus_client import SkuNexusClient, validate_po_row


def _normalize_vendor_key(name):
    if not name:
        return ''
    s = name.lower().strip()
    s = s.replace('&', 'and')
    s = ''.join(ch for ch in s if ch.isalnum())
    return s


def _split_vendor_aliases(value):
    if not value:
        return []
    parts = re.split(r'[|;]', str(value))
    return [p.strip() for p in parts if p.strip()]


def load_vendor_aliases(base_dir):
    """Load vendor alias map from vendors.csv in base directory."""
    path = os.path.join(base_dir, 'vendors.csv')
    if not os.path.exists(path):
        return {}
    try:
        with open(path, newline='', encoding='utf-8') as f:
            rows = list(csv.reader(f))
    except Exception:
        return {}
    if not rows:
        return {}

    header = [str(c).strip().lower() for c in rows[0]]
    has_header = any(
        h in ('vendor', 'invoice_vendor', 'aliases', 'alias', 'additional_names')
        for h in header
    )
    if not has_header:
        return {}

    def col(row, *names):
        for name in names:
            if name in header:
                idx = header.index(name)
                if idx < len(row):
                    return str(row[idx]).strip()
        return ''

    aliases = {}
    for row in rows[1:]:
        if not row:
            continue
        vendor = col(row, 'vendor')
        if not vendor:
            continue
        alias_value = col(row, 'aliases', 'alias', 'additional_names', 'invoice_vendor')
        if not alias_value and 'skunexus_vendor' in header:
            alias_value = col(row, 'skunexus_vendor')
        alias_list = _split_vendor_aliases(alias_value)
        if not alias_list:
            continue
        key = _normalize_vendor_key(vendor)
        if not key:
            continue
        bucket = aliases.setdefault(key, [])
        for alias in alias_list:
            if alias and alias not in bucket:
                bucket.append(alias)

    return aliases


def _looks_like_sku(value):
    if value is None:
        return False
    raw = str(value).strip()
    if not raw:
        return False
    lower = raw.lower().strip()
    # Explicit non-SKU labels / summary rows
    if 'core' in lower:
        return False
    normalized = re.sub(r'[^a-z0-9]+', '', lower)
    if not normalized:
        return False
    if normalized in {
        'core', 'ere', 'dppdiscount', 'discount',
        'dropship', 'shipping', 'freight',
        'totalamount', 'total', 'subtotal',
    }:
        return False
    # Require at least one digit to avoid matching plain words
    if not any(ch.isdigit() for ch in normalized):
        return False
    # Avoid very short tokens
    return len(normalized) >= 3


def get_base_dir():
    """Get the base directory - works for both script and PyInstaller exe."""
    if getattr(sys, 'frozen', False):
        # Running as PyInstaller bundle - use exe's directory
        return os.path.dirname(sys.executable)
    else:
        # Running as script
        return os.path.dirname(os.path.abspath(__file__))


def get_resource_path(relative_path):
    """Get path to bundled resource (works for PyInstaller onefile)."""
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), relative_path)


class InvoiceExtractorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Invoice Extractor")
        self._set_window_icon()
        self.root.geometry("750x650")
        self.root.resizable(True, True)

        self.base_dir = get_base_dir()
        self.app_dir = os.path.join(self.base_dir, 'app')
        os.makedirs(self.app_dir, exist_ok=True)
        self.invoices_dir = os.path.join(self.app_dir, 'invoices')
        self.output_file = os.path.join(self.base_dir, 'invoices_output.xlsx')
        self.log_file = os.path.join(self.base_dir, 'processed_log.json')

        self.is_running = False
        self.header_label = None
        self.header_image = None
        self.header_src_image = None
        self.header_src_width = None
        self.header_base_width = None
        self.header_current_width = None
        self.header_path = None
        self.header_animating = False
        self.header_shrunken = False
        self._header_anim_token = None
        self.header_base_left = 0

        self.build_ui()

    def _set_window_icon(self):
        """Set the window/taskbar icon (Tk default is the leaf)."""
        try:
            icon_path = get_resource_path('logo.ico')
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
        except Exception:
            # If icon can't be set (e.g., missing file), keep default.
            pass

    def build_ui(self):
        """Build the main application UI."""
        # Main container with padding
        main_frame = ttk.Frame(self.root, padding=15)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Header image (centered)
        self.header_path = get_resource_path('header.png')
        if os.path.exists(self.header_path):
            try:
                self.root.update_idletasks()
                window_width = self.root.winfo_width()
                if window_width <= 1:
                    window_width = 750
                # Base header size (80% width), then reduce by 35%
                target_width = int(window_width * 0.8 * 0.65)
                self.header_image = self._load_header_image(target_width)
                self.header_label = ttk.Label(main_frame, image=self.header_image, anchor='center')
                base_width = self.header_image.width() if self.header_image else None
                self.header_base_left = 0
                self.header_label.pack(fill=tk.X, pady=(0, 5))
                self.header_base_width = base_width
                self.header_current_width = self.header_base_width
            except Exception:
                self.header_image = None

        # Fallback title text if image isn't available
        if self.header_label is None:
            title_label = ttk.Label(
                main_frame, text="Invoice Extractor",
                font=('Segoe UI', 18, 'bold')
            )
            title_label.pack(pady=(0, 5))

        # Subtitle (two lines, with emphasis on the second line)
        subtitle_line1 = ttk.Label(
            main_frame,
            text="Download Invoices from Gmail, Parse them, export to Excel, Validate with SkuNexus,",
            font=('Segoe UI', 9),
            justify='center',
            anchor='center'
        )
        subtitle_line1.pack(pady=(0, 2))

        def _create_spaced_text(parent, text, font_obj, spacing_px=1):
            style = ttk.Style()
            bg = style.lookup('TFrame', 'background') or self.root.cget('bg')
            fg = style.lookup('TLabel', 'foreground') or 'black'
            canvas = tk.Canvas(parent, highlightthickness=0, bd=0, bg=bg)
            x = 0
            for ch in text:
                canvas.create_text(x, 0, text=ch, font=font_obj, anchor='nw', fill=fg)
                x += font_obj.measure(ch) + spacing_px
            width = max(1, x - spacing_px)
            height = font_obj.metrics('linespace')
            canvas.configure(width=width, height=height)
            return canvas

        subtitle_line2_font = tkfont.Font(family='Segoe UI', size=11)
        subtitle_line2 = _create_spaced_text(
            main_frame,
            "Burn and Churn Invoices like a Badass!",
            subtitle_line2_font,
            spacing_px=1
        )
        subtitle_line2.pack(pady=(0, 10))

        # Info frame
        info_frame = ttk.LabelFrame(main_frame, text="Status", padding=8)
        info_frame.pack(fill=tk.X, pady=(0, 10))

        # Credential status
        cred_exists = os.path.exists(os.path.join(self.app_dir, 'client_secret.json'))
        token_exists = os.path.exists(os.path.join(self.app_dir, 'token.pickle'))
        ocr_status = "Available" if OCR_AVAILABLE else "Not available (scanned PDFs will be skipped)"

        self.cred_label = ttk.Label(
            info_frame,
            text=f"Credentials: {'Found' if cred_exists else 'MISSING - place client_secret.json in app folder'}",
            foreground='green' if cred_exists else 'red'
        )
        self.cred_label.pack(anchor=tk.W)

        self.auth_label = ttk.Label(
            info_frame,
            text=f"Authentication: {'Cached (token.pickle found)' if token_exists else 'Will authenticate on first run'}",
            foreground='green' if token_exists else 'gray'
        )
        self.auth_label.pack(anchor=tk.W)

        ttk.Label(
            info_frame,
            text=f"OCR (Tesseract): {ocr_status}",
            foreground='green' if OCR_AVAILABLE else 'orange'
        ).pack(anchor=tk.W)

        # Count existing invoices
        invoice_count = 0
        if os.path.exists(self.invoices_dir):
            invoice_count = len([
                f for f in os.listdir(self.invoices_dir)
                if f.lower().endswith(('.pdf', '.png', '.jpg', '.jpeg', '.tiff'))
            ])
        ttk.Label(
            info_frame,
            text=f"Invoices downloaded: {invoice_count}"
        ).pack(anchor=tk.W)

        # Buttons frame
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(0, 10))

        self.go_button = ttk.Button(
            btn_frame, text="Go", command=self.start_processing,
            style='Accent.TButton'
        )
        self.go_button.pack(side=tk.LEFT, padx=(0, 5))

        self.stop_button = ttk.Button(
            btn_frame, text="Stop", command=self.stop_processing,
            state=tk.DISABLED
        )
        self.stop_button.pack(side=tk.LEFT, padx=(0, 5))

        self.open_excel_button = ttk.Button(
            btn_frame, text="Open Spreadsheet", command=self.open_spreadsheet
        )
        self.open_excel_button.pack(side=tk.LEFT, padx=(0, 5))

        self.open_folder_button = ttk.Button(
            btn_frame, text="Open Invoices Folder", command=self.open_invoices_folder
        )
        self.open_folder_button.pack(side=tk.LEFT, padx=(0, 5))

        self.validate_button = ttk.Button(
            btn_frame, text="Validate POs", command=self.start_validation
        )
        self.validate_button.pack(side=tk.LEFT)

        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            main_frame, variable=self.progress_var,
            maximum=100, mode='determinate'
        )
        self.progress_bar.pack(fill=tk.X, pady=(0, 5))

        self.progress_label = ttk.Label(main_frame, text="Ready", font=('Segoe UI', 9))
        self.progress_label.pack(anchor=tk.W, pady=(0, 5))

        # Status log
        self.log_frame = ttk.LabelFrame(main_frame, text="Log", padding=5)
        self.log_frame.pack(fill=tk.BOTH, expand=True)
        self.log_frame.pack_propagate(False)
        self.log_frame.configure(height=200)

        self.log_text = tk.Text(
            self.log_frame, wrap=tk.WORD, font=('Consolas', 9),
            state=tk.DISABLED, bg='#1e1e1e', fg='#cccccc',
            insertbackground='white'
        )
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(self.log_frame, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)

        # Configure text tags for colored output
        self.log_text.tag_configure('success', foreground='#4ec94e')
        self.log_text.tag_configure('error', foreground='#ff5555')
        self.log_text.tag_configure('warning', foreground='#ffaa00')
        self.log_text.tag_configure('info', foreground='#5599ff')
        self._ensure_log_min_height(200)

    def _init_header_source(self):
        if self.header_src_image is not None or not self.header_path or not os.path.exists(self.header_path):
            return
        try:
            if PIL_AVAILABLE:
                self.header_src_image = Image.open(self.header_path)
                self.header_src_width = self.header_src_image.width
            else:
                self.header_src_image = tk.PhotoImage(file=self.header_path)
                self.header_src_width = self.header_src_image.width()
        except Exception:
            self.header_src_image = None
            self.header_src_width = None

    def _load_header_image(self, target_width):
        self._init_header_source()
        if self.header_src_image is None or not self.header_src_width:
            return None

        target_width = int(target_width) if target_width else self.header_src_width
        if target_width <= 0:
            target_width = self.header_src_width

        scale = min(1.0, target_width / self.header_src_width)

        if PIL_AVAILABLE and isinstance(self.header_src_image, Image.Image):
            img = self.header_src_image
            if scale < 0.999:
                new_size = (
                    max(1, int(round(img.width * scale))),
                    max(1, int(round(img.height * scale)))
                )
                img = img.resize(new_size, Image.LANCZOS)
            return ImageTk.PhotoImage(img)

        img = self.header_src_image
        if scale >= 0.999:
            return img
        denom = 100
        num = max(1, int(round(scale * denom)))
        return img.zoom(num, num).subsample(denom, denom)

    def _update_header_width(self, target_width):
        if not self.header_label:
            return
        img = self._load_header_image(target_width)
        if not img:
            return
        self.header_image = img
        self.header_label.configure(image=self.header_image)
        try:
            self.header_current_width = self.header_image.width()
        except Exception:
            self.header_current_width = target_width

    def _ensure_log_min_height(self, min_height):
        def _apply():
            if not self.log_frame:
                return
            self.root.update_idletasks()
            current = self.log_frame.winfo_height()
            if current < min_height:
                width = self.root.winfo_width()
                height = self.root.winfo_height()
                delta = min_height - current
                self.root.geometry(f"{width}x{height + delta}")
                self.root.update_idletasks()
                self.log_frame.configure(height=min_height)
                self.log_frame.pack_propagate(False)
        self.root.after(0, _apply)

    def _ease_in_out_cubic(self, t):
        if t < 0.5:
            return 4 * t * t * t
        return 1 - pow(-2 * t + 2, 3) / 2

    def _animate_header_shrink(self):
        if not self.header_label or not self.header_base_width:
            return
        if self.header_animating or self.header_shrunken:
            return
        start_width = self.header_current_width or self.header_base_width
        target_width = int(round(self.header_base_width * 0.75))
        if target_width >= start_width:
            return

        self.header_animating = True
        duration_ms = 1750
        start_time = time.perf_counter()
        token = object()
        self._header_anim_token = token

        def step():
            if self._header_anim_token is not token:
                return
            elapsed_ms = (time.perf_counter() - start_time) * 1000.0
            t = min(1.0, elapsed_ms / duration_ms)
            eased = self._ease_in_out_cubic(t)
            width = int(round(start_width + (target_width - start_width) * eased))
            self._update_header_width(max(1, width))
            if t < 1.0:
                self.root.after(16, step)
            else:
                self.header_animating = False
                self.header_shrunken = True
                self.header_current_width = width

        step()

    def log(self, message, tag=None):
        """Add a message to the status log (thread-safe)."""
        def _update():
            self.log_text.config(state=tk.NORMAL)
            timestamp = datetime.now().strftime('%H:%M:%S')
            line = f"[{timestamp}] {message}\n"
            if tag:
                self.log_text.insert(tk.END, line, tag)
            else:
                self.log_text.insert(tk.END, line)
            self.log_text.see(tk.END)
            self.log_text.config(state=tk.DISABLED)

        self.root.after(0, _update)

    def set_progress(self, value, label_text=None):
        """Update progress bar and label (thread-safe)."""
        def _update():
            self.progress_var.set(value)
            if label_text:
                self.progress_label.config(text=label_text)
        self.root.after(0, _update)

    def start_processing(self):
        """Start the full extraction pipeline in a background thread."""
        if self.is_running:
            return

        self._animate_header_shrink()

        # Check for credentials
        if not os.path.exists(os.path.join(self.app_dir, 'client_secret.json')):
            self.log("ERROR: client_secret.json not found!", "error")
            self.log("Place your Google OAuth credentials file in the app folder.", "error")
            return

        self.is_running = True
        self.go_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)

        # Clear log
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)

        self.set_progress(0, "Starting...")

        thread = threading.Thread(target=self.run_pipeline, daemon=True)
        thread.start()

    def stop_processing(self):
        """Signal the pipeline to stop."""
        self.is_running = False
        self.log("Stop requested - will halt after current operation...", "warning")

    def run_pipeline(self):
        """Main processing pipeline - runs in background thread."""
        try:
            # Phase 1: Gmail - download attachments
            self.set_progress(5, "Connecting to Gmail...")
            self.log("=== Phase 1: Fetching emails from Gmail ===", "info")

            client = GmailClient(
                self.base_dir,
                status_callback=self.log,
                data_dir=self.app_dir,
                log_file=self.log_file,
                invoices_dir=self.invoices_dir
            )
            client.authenticate()

            if not self.is_running:
                self.finish("Stopped by user.")
                return

            self.set_progress(10, "Downloading attachments...")
            downloaded_files, total_emails, new_emails = (
                client.fetch_and_download_new_attachments()
            )

            if not self.is_running:
                self.finish("Stopped by user.")
                return

            # Phase 2: Parse invoices
            self.set_progress(40, "Parsing invoices...")
            self.log("", None)
            self.log("=== Phase 2: Parsing invoice files ===", "info")

            log_data = client.load_processed_log()
            processed_invoices = log_data.get("processed_invoices", {})

            # Find all invoice files that haven't been parsed yet
            all_invoice_files = []
            if os.path.exists(self.invoices_dir):
                for f in os.listdir(self.invoices_dir):
                    if f.lower().endswith(('.pdf', '.png', '.jpg', '.jpeg', '.tiff')):
                        if f not in processed_invoices:
                            all_invoice_files.append(f)

            if not all_invoice_files:
                self.log("No new invoice files to parse.", "success")
            else:
                self.log(f"Found {len(all_invoice_files)} new invoice files to parse.")

                success_count = 0
                error_count = 0
                error_files = []

                for i, filename in enumerate(all_invoice_files):
                    if not self.is_running:
                        self.finish("Stopped by user.")
                        return

                    progress = 40 + (50 * (i + 1) / len(all_invoice_files))
                    self.set_progress(
                        progress,
                        f"Parsing invoice {i + 1}/{len(all_invoice_files)}..."
                    )

                    filepath = os.path.join(self.invoices_dir, filename)
                    self.log(f"Processing: {filename}")

                    try:
                        invoice_data = parse_invoice(filepath, self.log)

                        if invoice_data:
                            # Write to spreadsheet
                            write_invoice_to_spreadsheet(
                                self.output_file, invoice_data, self.log
                            )
                            success_count += 1
                        else:
                            error_count += 1
                            error_files.append(filename)

                        # Mark as processed regardless
                        processed_invoices[filename] = datetime.now().isoformat()

                    except Exception as e:
                        self.log(f"  Failed to parse {filename}: {e}", "error")
                        error_count += 1
                        error_files.append(filename)
                        processed_invoices[filename] = datetime.now().isoformat()

                # Save updated processed log
                log_data["processed_invoices"] = processed_invoices
                client.save_processed_log(log_data)

                # Summary
                self.log("")
                self.log("=== Summary ===", "info")
                self.log(f"Emails checked: {total_emails} total, {new_emails} new")
                self.log(f"Attachments downloaded: {len(downloaded_files)}")
                self.log(f"Invoices parsed successfully: {success_count}", "success")
                if error_count:
                    self.log(f"Invoices with errors: {error_count}", "error")
                    for ef in error_files:
                        self.log(f"  - {ef}", "error")

            self.finish("Processing complete!")

        except FileNotFoundError as e:
            self.log(f"File not found: {e}", "error")
            self.finish("Failed - missing file.")
        except Exception as e:
            self.log(f"Unexpected error: {e}", "error")
            self.finish("Failed with error.")

    def finish(self, message):
        """Reset UI state after pipeline completes."""
        def _update():
            self.is_running = False
            self.go_button.config(state=tk.NORMAL)
            self.stop_button.config(state=tk.DISABLED)
            self.set_progress(100, message)
            self.log(message, "info")
        self.root.after(0, _update)

    def open_spreadsheet(self):
        """Open the output spreadsheet in the default application."""
        if os.path.exists(self.output_file):
            os.startfile(self.output_file)
        else:
            self.log("Spreadsheet not found - run extraction first.", "warning")

    def open_invoices_folder(self):
        """Open the invoices folder in Explorer."""
        os.makedirs(self.invoices_dir, exist_ok=True)
        os.startfile(self.invoices_dir)

    def start_validation(self):
        """Start the PO validation pipeline in a background thread."""
        if self.is_running:
            return

        # Check if spreadsheet exists
        if not os.path.exists(self.output_file):
            self.log("ERROR: No spreadsheet found - run extraction first!", "error")
            return

        self.is_running = True
        self.go_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        self.validate_button.config(state=tk.DISABLED)

        # Clear log
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)

        self.set_progress(0, "Starting validation...")

        thread = threading.Thread(target=self.run_validation_pipeline, daemon=True)
        thread.start()

    def run_validation_pipeline(self):
        """Validate POs against SkuNexus - runs in background thread."""
        try:
            self.log("=== SkuNexus PO Validation ===", "info")

            # Load SkuNexus credentials from config file
            config_path = os.path.join(self.base_dir, 'skunexus_config.json')
            if not os.path.exists(config_path):
                config_path = os.path.join(self.app_dir, 'skunexus_config.json')
            if not os.path.exists(config_path):
                self.log("ERROR: skunexus_config.json not found!", "error")
                self.log("Create skunexus_config.json with 'email' and 'password' fields in the app folder.", "error")
                self.finish_validation("Validation failed - missing config file")
                return

            with open(config_path, 'r') as f:
                skunexus_config = json.load(f)

            skunexus_email = skunexus_config.get('email', '')
            skunexus_password = skunexus_config.get('password', '')

            if not skunexus_email or not skunexus_password:
                self.log("ERROR: Invalid skunexus_config.json - missing email or password", "error")
                self.finish_validation("Validation failed - invalid config")
                return

            # Login to SkuNexus
            self.set_progress(5, "Logging into SkuNexus...")
            self.log("Connecting to SkuNexus...")

            client = SkuNexusClient(skunexus_email, skunexus_password)
            success, message = client.login()

            if not success:
                self.log(f"Failed to login to SkuNexus: {message}", "error")
                self.finish_validation("Validation failed - could not login to SkuNexus")
                return

            self.log("Successfully logged into SkuNexus", "success")

            if not self.is_running:
                self.finish_validation("Stopped by user.")
                return

            # Read all rows from spreadsheet
            self.set_progress(10, "Reading spreadsheet...")
            rows = read_spreadsheet_rows(self.output_file)

            vendor_aliases = load_vendor_aliases(self.base_dir)

            if not rows:
                self.log("No data found in spreadsheet.", "warning")
                self.finish_validation("Validation complete - no data to validate")
                return

            self.log(f"Found {len(rows)} rows in spreadsheet")

            # Map Bill No. -> Memo (PO) so we can validate item rows that omit Memo
            memo_by_bill = {}
            for row in rows:
                bill_no = str(row.get('bill_no', '')).strip()
                memo = str(row.get('memo', '')).strip()
                if bill_no and memo:
                    memo_by_bill[bill_no] = memo

            # Group rows by PO number (collect vendor/SKU hints)
            po_groups = {}
            for row in rows:
                memo = row.get('memo', '')
                if not memo:
                    bill_no = str(row.get('bill_no', '')).strip()
                    memo = memo_by_bill.get(bill_no, '')
                if not memo:
                    continue
                group = po_groups.setdefault(memo, {'rows': [], 'vendors': [], 'skus': []})
                group['rows'].append(row)
                vendor = str(row.get('vendor', '')).strip()
                if vendor:
                    group['vendors'].append(vendor)
                sku = str(row.get('product_service', '')).strip()
                if sku:
                    group['skus'].append(sku)

            def _pick_vendor(vendors):
                if not vendors:
                    return ''
                counts = {}
                for v in vendors:
                    counts[v] = counts.get(v, 0) + 1
                return max(counts, key=counts.get)

            self.log(f"Found {len(po_groups)} unique PO numbers to validate")

            if not self.is_running:
                self.finish_validation("Stopped by user.")
                return

            # Validate each PO
            validated_count = 0
            passed_count = 0
            failed_count = 0
            not_found_count = 0
            po_cache = {}  # Cache SkuNexus data by PO number

            total_rows = len(rows)
            for i, row in enumerate(rows):
                if not self.is_running:
                    self.finish_validation("Stopped by user.")
                    return

                progress = 15 + (80 * (i + 1) / total_rows)
                self.set_progress(progress, f"Validating row {i + 1}/{total_rows}...")

                row_num = row['_row_num']
                memo = row.get('memo', '')
                if not memo:
                    bill_no = str(row.get('bill_no', '')).strip()
                    memo = memo_by_bill.get(bill_no, '')
                category = row.get('category', '')

                # Skip rows without PO number
                if not memo:
                    write_validation_result(self.output_file, row_num, None, [])
                    continue

                # Only validate SKU rows (Category/Account = Purchases)
                if category != 'Purchases':
                    write_validation_result(self.output_file, row_num, None, [])
                    continue

                product_service = str(row.get('product_service', '')).strip()
                if not _looks_like_sku(product_service):
                    write_validation_result(self.output_file, row_num, None, [])
                    continue

                # Get SkuNexus data (from cache or fetch)
                if memo not in po_cache:
                    # Extract PO number without "PO" prefix
                    po_number = memo[2:] if memo.upper().startswith('PO') else memo

                    self.log(f"Fetching PO {po_number} from SkuNexus...")
                    group = po_groups.get(memo, {})
                    vendor_hint = _pick_vendor(group.get('vendors', []))
                    sku_hints = group.get('skus', [])
                    sn_data, error = client.get_best_po_with_line_items(
                        po_number,
                        invoice_vendor=vendor_hint,
                        invoice_skus=sku_hints,
                        vendor_aliases=vendor_aliases
                    )

                    if error:
                        self.log(f"  Could not find PO {po_number}: {error}", "warning")
                        po_cache[memo] = None
                    else:
                        po_cache[memo] = sn_data
                        self.log(f"  Found PO with {len(sn_data.get('lineItems', {}).get('rows', []))} line items")

                sn_data = po_cache.get(memo)

                if sn_data is None:
                    # PO not found in SkuNexus
                    write_validation_result(self.output_file, row_num, False, ['PO not found in SkuNexus'])
                    validated_count += 1
                    not_found_count += 1
                    continue

                # Validate this row against SkuNexus data
                is_valid, failed_fields = validate_po_row(sn_data, row, vendor_aliases)

                write_validation_result(self.output_file, row_num, is_valid, failed_fields)
                validated_count += 1

                if is_valid:
                    passed_count += 1
                else:
                    failed_count += 1
                    sku = row.get('product_service', 'N/A')
                    self.log(f"  Row {row_num} (SKU: {sku}) - FAILED: {', '.join(failed_fields)}", "warning")

            # Summary
            self.log("")
            self.log("=== Validation Summary ===", "info")
            self.log(f"Total rows validated: {validated_count}")
            self.log(f"Passed: {passed_count}", "success")
            if failed_count:
                self.log(f"Failed: {failed_count}", "error")
            if not_found_count:
                self.log(f"POs not found: {not_found_count}", "warning")

            self.finish_validation("Validation complete!")

        except Exception as e:
            self.log(f"Validation error: {e}", "error")
            import traceback
            self.log(traceback.format_exc(), "error")
            self.finish_validation("Validation failed with error.")

    def finish_validation(self, message):
        """Reset UI state after validation completes."""
        def _update():
            self.is_running = False
            self.go_button.config(state=tk.NORMAL)
            self.stop_button.config(state=tk.DISABLED)
            self.validate_button.config(state=tk.NORMAL)
            self.set_progress(100, message)
            self.log(message, "info")
        self.root.after(0, _update)


def main():
    root = tk.Tk()

    # Try to use a modern theme
    try:
        style = ttk.Style()
        available_themes = style.theme_names()
        if 'vista' in available_themes:
            style.theme_use('vista')
        elif 'clam' in available_themes:
            style.theme_use('clam')
    except Exception:
        pass

    app = InvoiceExtractorGUI(root)
    root.mainloop()


if __name__ == '__main__':
    main()
