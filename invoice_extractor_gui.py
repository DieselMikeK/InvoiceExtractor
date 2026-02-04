#!/usr/bin/env python3
"""
Invoice Extractor - GUI Application
Connects to Gmail, downloads invoice attachments, parses them,
and exports extracted data to an Excel spreadsheet.
"""

import os
import sys
import json
import threading
import tkinter as tk
from tkinter import ttk
from datetime import datetime

from gmail_client import GmailClient
from invoice_parser import parse_invoice, OCR_AVAILABLE
from spreadsheet_writer import (
    write_invoice_to_spreadsheet, read_spreadsheet_rows,
    write_validation_result, get_unique_po_numbers
)
from skunexus_client import SkuNexusClient, validate_po_row


def get_base_dir():
    """Get the base directory - works for both script and PyInstaller exe."""
    if getattr(sys, 'frozen', False):
        # Running as PyInstaller bundle - use exe's directory
        return os.path.dirname(sys.executable)
    else:
        # Running as script
        return os.path.dirname(os.path.abspath(__file__))


class InvoiceExtractorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Invoice Extractor")
        self.root.geometry("750x650")
        self.root.resizable(True, True)

        self.base_dir = get_base_dir()
        self.invoices_dir = os.path.join(self.base_dir, 'invoices')
        self.output_file = os.path.join(self.base_dir, 'invoices_output.xlsx')
        self.log_file = os.path.join(self.base_dir, 'processed_log.json')

        self.is_running = False

        self.build_ui()

    def build_ui(self):
        """Build the main application UI."""
        # Main container with padding
        main_frame = ttk.Frame(self.root, padding=15)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Title
        title_label = ttk.Label(
            main_frame, text="Invoice Extractor",
            font=('Segoe UI', 18, 'bold')
        )
        title_label.pack(pady=(0, 5))

        # Subtitle
        subtitle = ttk.Label(
            main_frame,
            text="Download invoices from Gmail, parse them, and export to Excel",
            font=('Segoe UI', 9)
        )
        subtitle.pack(pady=(0, 10))

        # Info frame
        info_frame = ttk.LabelFrame(main_frame, text="Status", padding=8)
        info_frame.pack(fill=tk.X, pady=(0, 10))

        # Credential status
        cred_exists = os.path.exists(os.path.join(self.base_dir, 'client_secret.json'))
        token_exists = os.path.exists(os.path.join(self.base_dir, 'token.pickle'))
        ocr_status = "Available" if OCR_AVAILABLE else "Not available (scanned PDFs will be skipped)"

        self.cred_label = ttk.Label(
            info_frame,
            text=f"Credentials: {'Found' if cred_exists else 'MISSING - place client_secret.json in project folder'}",
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
        log_frame = ttk.LabelFrame(main_frame, text="Log", padding=5)
        log_frame.pack(fill=tk.BOTH, expand=True)

        self.log_text = tk.Text(
            log_frame, wrap=tk.WORD, font=('Consolas', 9),
            state=tk.DISABLED, bg='#1e1e1e', fg='#cccccc',
            insertbackground='white'
        )
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)

        # Configure text tags for colored output
        self.log_text.tag_configure('success', foreground='#4ec94e')
        self.log_text.tag_configure('error', foreground='#ff5555')
        self.log_text.tag_configure('warning', foreground='#ffaa00')
        self.log_text.tag_configure('info', foreground='#5599ff')

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

        # Check for credentials
        if not os.path.exists(os.path.join(self.base_dir, 'client_secret.json')):
            self.log("ERROR: client_secret.json not found!", "error")
            self.log("Place your Google OAuth credentials file in the project folder.", "error")
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

            client = GmailClient(self.base_dir, status_callback=self.log)
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
                self.log("ERROR: skunexus_config.json not found!", "error")
                self.log("Create skunexus_config.json with 'email' and 'password' fields.", "error")
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

            if not rows:
                self.log("No data found in spreadsheet.", "warning")
                self.finish_validation("Validation complete - no data to validate")
                return

            self.log(f"Found {len(rows)} rows in spreadsheet")

            # Group rows by PO number
            po_groups = {}
            for row in rows:
                memo = row.get('memo', '')
                if memo:
                    if memo not in po_groups:
                        po_groups[memo] = []
                    po_groups[memo].append(row)

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
                category = row.get('category', '')

                # Skip rows without PO number
                if not memo:
                    write_validation_result(self.output_file, row_num, True, [])
                    continue

                # Skip shipping rows - they pass validation
                if category == 'Freight/Shipping':
                    write_validation_result(self.output_file, row_num, True, [])
                    validated_count += 1
                    passed_count += 1
                    continue

                # Get SkuNexus data (from cache or fetch)
                if memo not in po_cache:
                    # Extract PO number without "PO" prefix
                    po_number = memo[2:] if memo.upper().startswith('PO') else memo

                    self.log(f"Fetching PO {po_number} from SkuNexus...")
                    sn_data, error = client.get_po_with_line_items(po_number)

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
                is_valid, failed_fields = validate_po_row(sn_data, row)

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
