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
import ctypes
import shutil
from datetime import datetime, timedelta
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
try:
    from PIL import Image, ImageTk
    PIL_AVAILABLE = True
except Exception:
    PIL_AVAILABLE = False
try:
    from tkcalendar import Calendar
    TKCALENDAR_AVAILABLE = True
except Exception:
    TKCALENDAR_AVAILABLE = False

from gmail_client import (
    GmailClient,
    PROCESSED_LABEL_NAME,
    WrongAuthorizedAccountError
)
from invoice_parser import parse_invoice, OCR_AVAILABLE
from spreadsheet_writer import (
    COLUMNS, write_invoice_to_spreadsheet, read_spreadsheet_rows,
    write_validation_result, write_validation_results, get_unique_po_numbers
)
from skunexus_client import SkuNexusClient, validate_po_row

# Batch export settings (easy to change)
BATCH_ROW_LIMIT = 1000  # Max data rows per CSV batch (header not counted)
BATCH_INVOICE_LIMIT = 100  # Max invoices per CSV batch
BATCH_FOLDER_PREFIX = "Batch_"
BATCHES_ROOT_NAME = "Batches"
AUTHORIZED_GMAIL_ACCOUNT = "dppautoap@gmail.com"

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


def load_vendor_aliases(preferred_dir):
    """Load vendor alias map from vendors.csv (prefer provided dir, fallback to parent)."""
    path = os.path.join(preferred_dir, 'vendors.csv')
    if not os.path.exists(path):
        parent_dir = os.path.dirname(preferred_dir)
        fallback = os.path.join(parent_dir, 'vendors.csv')
        if os.path.exists(fallback):
            path = fallback
        else:
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


def _get_row_sku(row):
    sku = str(row.get('sku', '')).strip()
    if sku:
        return sku
    return str(row.get('product_service', '')).strip()


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


def _extract_date_tag_from_filename(filename):
    match = re.search(r'(?:invoices_output_|Invoices_Master_)(\d{1,2}-\d{1,2})', filename)
    if match:
        return match.group(1)
    return None


_single_instance_mutex = None


def _bring_existing_window_to_front(window_title):
    if os.name != 'nt':
        return False
    try:
        user32 = ctypes.windll.user32
        hwnd = user32.FindWindowW(None, window_title)
        if hwnd:
            SW_RESTORE = 9
            user32.ShowWindow(hwnd, SW_RESTORE)
            user32.SetForegroundWindow(hwnd)
            return True
    except Exception:
        pass
    return False


def _ensure_single_instance(window_title):
    if os.name != 'nt':
        return True
    try:
        kernel32 = ctypes.windll.kernel32
        mutex = kernel32.CreateMutexW(None, False, "InvoiceExtractor_SingleInstance")
        already_exists = kernel32.GetLastError() == 183  # ERROR_ALREADY_EXISTS
        if already_exists:
            _bring_existing_window_to_front(window_title)
            return False
        global _single_instance_mutex
        _single_instance_mutex = mutex
    except Exception:
        # If mutex fails for some reason, allow app to run.
        return True
    return True


class InvoiceExtractorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Invoice Extractor")
        self._set_window_icon()
        self.root.geometry("750x650")
        self.root.resizable(True, True)

        self.base_dir = get_base_dir()
        if getattr(sys, 'frozen', False):
            self.app_dir = os.path.join(self.base_dir, 'App')
        else:
            # Running from source inside App folder
            self.app_dir = self.base_dir
        self.required_dir = os.path.join(self.app_dir, 'required')
        os.makedirs(self.app_dir, exist_ok=True)
        os.makedirs(self.required_dir, exist_ok=True)
        self.invoices_root = os.path.join(self.base_dir, 'Invoices')
        self.batches_root = os.path.join(self.base_dir, BATCHES_ROOT_NAME)
        os.makedirs(self.invoices_root, exist_ok=True)
        os.makedirs(self.batches_root, exist_ok=True)
        self._migrate_required_files()
        self.output_file, self.invoices_dir = self._get_next_run_paths()

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
        self.last_batches_dir = None
        self._calendar_open = False
        self._calendar_warned = False
        self._calendar_suppress_widget = None
        self._calendar_suppress_until = 0.0

        self.build_ui()

    def _get_next_run_paths(self):
        """Pick the next available output file and invoices folder (same suffix)."""
        date_tag = f"{datetime.now().month}-{datetime.now().day}"
        output_base = f"Invoices_Master_{date_tag}"
        folder_base = f"invoices_{date_tag}"
        ext = '.xlsx'

        output_path = os.path.join(self.base_dir, f"{output_base}{ext}")
        folder_path = os.path.join(self.invoices_root, folder_base)
        if (not os.path.exists(output_path)) and (not os.path.exists(folder_path)):
            return output_path, folder_path

        for i in range(2, 10000):
            output_candidate = os.path.join(self.base_dir, f"{output_base}_{i}{ext}")
            folder_candidate = os.path.join(self.invoices_root, f"{folder_base}_{i}")
            if (not os.path.exists(output_candidate)) and (not os.path.exists(folder_candidate)):
                return output_candidate, folder_candidate

        # Fallback: use timestamp if we somehow hit a huge count
        stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_path = os.path.join(self.base_dir, f"{output_base}_{stamp}{ext}")
        folder_path = os.path.join(self.invoices_root, f"{folder_base}_{stamp}")
        return output_path, folder_path

    def _get_output_files_for_validation(self):
        """Return all output XLSX files in base_dir (sorted oldest to newest)."""
        files = []
        for name in os.listdir(self.base_dir):
            lower = name.lower()
            if not lower.endswith('.xlsx'):
                continue
            if not (lower.startswith('invoices_output') or lower.startswith('invoices_master_')):
                continue
            files.append(os.path.join(self.base_dir, name))
        files.sort(key=lambda p: os.path.getmtime(p))
        return files

    def _find_master_spreadsheets(self):
        """Return all master invoice XLSX files in base_dir (newest first)."""
        files = []
        for name in os.listdir(self.base_dir):
            lower = name.lower()
            if not lower.endswith('.xlsx'):
                continue
            if not (lower.startswith('invoices_output') or lower.startswith('invoices_master_')):
                continue
            if name.startswith('~$'):
                continue
            files.append(os.path.join(self.base_dir, name))
        files.sort(key=lambda p: os.path.getmtime(p), reverse=True)
        return files

    def _select_master_for_batching(self):
        """Pick the best master spreadsheet for batching (prefer today's date tag)."""
        masters = self._find_master_spreadsheets()
        if not masters:
            return None
        date_tag = f"{datetime.now().month}-{datetime.now().day}"
        today_matches = [
            p for p in masters
            if f"Invoices_Master_{date_tag}" in os.path.basename(p)
        ]
        if today_matches:
            return today_matches[0]
        return masters[0]

    def _get_next_batches_dir(self, date_tag):
        base_name = f"{BATCH_FOLDER_PREFIX}{date_tag}"
        candidate = os.path.join(self.batches_root, base_name)
        if not os.path.exists(candidate):
            return candidate
        for i in range(2, 10000):
            candidate = os.path.join(self.batches_root, f"{base_name}_{i}")
            if not os.path.exists(candidate):
                return candidate
        stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        return os.path.join(self.batches_root, f"{base_name}_{stamp}")

    def _migrate_required_files(self):
        files = [
            'client_secret.json',
            'token.pickle',
            'skunexus_config.json',
            'invoice_history.csv',
        ]
        for name in files:
            dest = os.path.join(self.required_dir, name)
            if os.path.exists(dest):
                continue
            legacy_app_dir = os.path.join(self.base_dir, 'app')
            for src_dir in (self.app_dir, legacy_app_dir, self.base_dir):
                src = os.path.join(src_dir, name)
                if os.path.exists(src):
                    try:
                        os.replace(src, dest)
                    except Exception:
                        try:
                            shutil.copy2(src, dest)
                        except Exception:
                            pass
                    break

    def _history_log_path(self):
        return os.path.join(self.required_dir, 'invoice_history.csv')

    def _load_invoice_history(self):
        path = self._history_log_path()
        if not os.path.exists(path):
            return []
        rows = []
        try:
            with open(path, newline='', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    rows.append({k: (v or '').strip() for k, v in row.items()})
        except Exception as e:
            self.log(f"Warning: could not read invoice history ({e})", "warning")
        return rows

    def _append_invoice_history(self, entries):
        if not entries:
            return
        path = self._history_log_path()
        fieldnames = [
            'bill_no', 'po_number', 'vendor', 'invoice_date',
            'downloaded_at', 'source_file'
        ]
        file_exists = os.path.exists(path)
        try:
            with open(path, 'a', newline='', encoding='utf-8') as f:
                writer = csv.DictWriter(f, fieldnames=fieldnames)
                if not file_exists:
                    writer.writeheader()
                for entry in entries:
                    writer.writerow({k: entry.get(k, '') for k in fieldnames})
        except Exception as e:
            self.log(f"Warning: could not update invoice history ({e})", "warning")

    def _apply_duplicate_flags(self, filepath, history_by_po):
        if not os.path.exists(filepath):
            return
        rows = read_spreadsheet_rows(filepath)
        if not rows:
            return

        # Resolve memo (PO) for rows with blank memo by using first memo per bill
        memo_by_bill = {}
        for row in rows:
            bill_no = str(row.get('bill_no', '')).strip()
            memo = str(row.get('memo', '')).strip()
            if bill_no and memo:
                memo_by_bill[bill_no] = memo

        # Track current bill per row and build bill->po mapping
        bill_by_row = {}
        bill_to_po = {}
        first_row_by_bill = {}
        bill_to_index = {}
        invoice_index = -1
        current_bill = ''
        for row in rows:
            row_num = row.get('_row_num')
            bill_no = str(row.get('bill_no', '')).strip()
            memo = str(row.get('memo', '')).strip()
            if bill_no:
                current_bill = bill_no
                if current_bill not in bill_to_index:
                    invoice_index += 1
                    bill_to_index[current_bill] = invoice_index
                if current_bill and current_bill not in first_row_by_bill:
                    first_row_by_bill[current_bill] = row_num
            if not memo and current_bill:
                memo = memo_by_bill.get(current_bill, '')
            if current_bill:
                bill_by_row[row_num] = current_bill
                if memo and current_bill not in bill_to_po:
                    bill_to_po[current_bill] = memo

        # Build PO -> bills mapping for current file
        po_to_bills = {}
        for bill_no, po in bill_to_po.items():
            if not po:
                continue
            po_to_bills.setdefault(po, []).append(bill_no)

        # Compute duplicate status per bill
        status_by_bill = {}
        ref_by_bill = {}
        for bill_no, po in bill_to_po.items():
            if not po:
                continue
            status_parts = []
            ref_parts = []
            bills_with_po = po_to_bills.get(po, [])
            if len(bills_with_po) > 1:
                status_parts.append("Duplicate PO (current run)")
                others = [b for b in bills_with_po if b != bill_no]
                if others:
                    ref_parts.append("Current: " + ", ".join(sorted(others)))
            history_entries = history_by_po.get(po, [])
            if history_entries:
                status_parts.append("Duplicate PO (history)")
                hist_refs = []
                seen_refs = set()
                for entry in history_entries:
                    date_val = str(entry.get('invoice_date', '')).strip()
                    if not date_val:
                        downloaded = str(entry.get('downloaded_at', '')).strip()
                        if downloaded:
                            date_val = downloaded.split(' ')[0]
                    ref = f"{po} - {date_val}" if date_val else po
                    if ref and ref not in seen_refs:
                        hist_refs.append(ref)
                        seen_refs.add(ref)
                if hist_refs:
                    ref_parts.append("History: " + ", ".join(hist_refs[:5]))
            if status_parts:
                status_by_bill[bill_no] = "; ".join(status_parts)
                ref_by_bill[bill_no] = " | ".join(ref_parts)

        if not status_by_bill:
            return

        # Update duplicate columns in spreadsheet
        col_keys = [key for key, _ in COLUMNS]
        try:
            dup_status_col = col_keys.index('duplicate_status') + 1
            dup_ref_col = col_keys.index('duplicate_reference') + 1
        except ValueError:
            return

        wb = load_workbook(filepath)
        ws = wb.active
        dup_bills = set(status_by_bill.keys())
        dup_fill_light = PatternFill(start_color="A8A8A8", end_color="A8A8A8", fill_type='solid')
        dup_fill_dark = PatternFill(start_color="888888", end_color="888888", fill_type='solid')

        for row_num in range(2, ws.max_row + 1):
            bill_no = ws.cell(row=row_num, column=1).value
            bill_no = str(bill_no).strip() if bill_no else ''
            if not bill_no:
                bill_no = bill_by_row.get(row_num, '')
            if not bill_no:
                continue
            is_first = first_row_by_bill.get(bill_no) == row_num
            status = status_by_bill.get(bill_no)
            if is_first and status:
                ws.cell(row=row_num, column=dup_status_col, value=status)
                ws.cell(row=row_num, column=dup_ref_col, value=ref_by_bill.get(bill_no, ''))
            else:
                ws.cell(row=row_num, column=dup_status_col, value='')
                ws.cell(row=row_num, column=dup_ref_col, value='')

            # Match row fill for duplicate columns
            first_cell = ws.cell(row=row_num, column=1)
            is_yellow = False
            if first_cell.fill and first_cell.fill.patternType == 'solid':
                color = first_cell.fill.start_color.rgb or first_cell.fill.start_color.index
                if color and str(color).upper().endswith('FFFF00'):
                    is_yellow = True
            if first_cell.fill and first_cell.fill.patternType == 'solid':
                row_fill = PatternFill(
                    start_color=first_cell.fill.start_color.rgb,
                    end_color=first_cell.fill.end_color.rgb,
                    fill_type='solid'
                )
                ws.cell(row=row_num, column=dup_status_col).fill = row_fill
                ws.cell(row=row_num, column=dup_ref_col).fill = row_fill

            # Override row background for duplicates with darker alternating gray
            if bill_no in dup_bills and not is_yellow:
                idx = bill_to_index.get(bill_no, 0)
                row_fill = dup_fill_dark if (idx % 2 == 0) else dup_fill_light
                for col in range(1, ws.max_column + 1):
                    ws.cell(row=row_num, column=col).fill = row_fill

        wb.save(filepath)

    def _refresh_batch_buttons(self):
        masters = self._find_master_spreadsheets()
        state = tk.NORMAL if masters else tk.DISABLED
        if hasattr(self, 'export_batches_button'):
            self.export_batches_button.config(state=state)
        if hasattr(self, 'validate_button'):
            outputs = self._get_output_files_for_validation()
            validate_state = tk.NORMAL if outputs else tk.DISABLED
            self.validate_button.config(state=validate_state)

    def _set_window_icon(self):
        """Set the window/taskbar icon (Tk default is the leaf)."""
        try:
            icon_path = get_resource_path('logo.ico')
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
        except Exception:
            # If icon can't be set (e.g., missing file), keep default.
            pass

    def _open_date_picker(self, target_var):
        return self._open_date_picker_at(target_var, None)

    def _open_date_picker_at(self, target_var, anchor_widget, title=None):
        if not TKCALENDAR_AVAILABLE:
            if not getattr(self, '_calendar_warned', False):
                self.log("Calendar picker unavailable (tkcalendar not installed).", "warning")
                self._calendar_warned = True
            return
        if getattr(self, '_calendar_open', False):
            return
        # Suppress immediate reopen after close (regardless of widget)
        if time.time() < getattr(self, '_calendar_suppress_until', 0):
            return

        self._calendar_open = True
        top = tk.Toplevel(self.root)
        top.title(title or "Select Date")
        top.transient(self.root)
        top.grab_set()

        cal = Calendar(top, selectmode='day', date_pattern='yyyy/mm/dd')
        # Preselect if existing
        existing = self._parse_date_input(target_var.get().strip())
        if existing:
            cal.selection_set(existing)
        cal.pack(padx=10, pady=10)

        def _close():
            self._calendar_open = False
            self._calendar_suppress_widget = anchor_widget
            self._calendar_suppress_until = time.time() + 0.5
            top.grab_release()
            top.destroy()

        def _set_date():
            target_var.set(cal.get_date())
            _close()

        btn_frame = ttk.Frame(top)
        btn_frame.pack(pady=(0, 10))
        ttk.Button(btn_frame, text="Set Date", command=_set_date).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="Cancel", command=_close).pack(side=tk.LEFT)
        top.protocol("WM_DELETE_WINDOW", _close)

        # Position popup above the input (fallback to below if needed)
        top.update_idletasks()
        if anchor_widget is not None:
            x = anchor_widget.winfo_rootx()
            y = anchor_widget.winfo_rooty() - top.winfo_height() - 8
            if y < 0:
                y = anchor_widget.winfo_rooty() + anchor_widget.winfo_height() + 8
            top.geometry(f"+{x}+{y}")

    def _parse_date_input(self, value):
        if not value:
            return None
        for fmt in ("%Y/%m/%d", "%Y-%m-%d", "%m/%d/%Y"):
            try:
                return datetime.strptime(value, fmt).date()
            except ValueError:
                continue
        return None

    def _build_gmail_query(self):
        """Build Gmail search query based on date filter options."""
        if self.today_filter_var.get():
            today = datetime.now().date()
            after = today.strftime("%Y/%m/%d")
            before = (today + timedelta(days=1)).strftime("%Y/%m/%d")
            return f"after:{after} before:{before}", "today"

        if self.date_filter_var.get():
            from_raw = self.date_from_var.get().strip()
            to_raw = self.date_to_var.get().strip()
            if not from_raw and not to_raw:
                self.log("Date filter is enabled but no dates were provided.", "error")
                return None, None
            from_date = self._parse_date_input(from_raw)
            to_date = self._parse_date_input(to_raw)
            if from_raw and not from_date:
                self.log("Invalid From date. Use YYYY/MM/DD.", "error")
                return None, None
            if to_raw and not to_date:
                self.log("Invalid To date. Use YYYY/MM/DD.", "error")
                return None, None
            if from_date and to_date and from_date > to_date:
                self.log("From date must be on or before To date.", "error")
                return None, None

            parts = []
            if from_date:
                parts.append(f"after:{from_date.strftime('%Y/%m/%d')}")
            if to_date:
                inclusive_before = to_date + timedelta(days=1)
                parts.append(f"before:{inclusive_before.strftime('%Y/%m/%d')}")
            return " ".join(parts), "range"

        return f"-label:{PROCESSED_LABEL_NAME}", "label"

    def _update_date_filter_state(self):
        enabled = self.date_filter_var.get()
        state = tk.NORMAL if enabled else tk.DISABLED
        if hasattr(self, 'date_from_entry'):
            self.date_from_entry.config(state=state)
        if hasattr(self, 'date_to_entry'):
            self.date_to_entry.config(state=state)

    def _on_date_filter_toggle(self):
        if self.date_filter_var.get():
            self.today_filter_var.set(False)
        self._update_date_filter_state()

    def _on_today_filter_toggle(self):
        if self.today_filter_var.get():
            self.date_filter_var.set(False)
        self._update_date_filter_state()

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
        cred_exists = os.path.exists(os.path.join(self.required_dir, 'client_secret.json'))
        token_exists = os.path.exists(os.path.join(self.required_dir, 'token.pickle'))
        ocr_status = "Available" if OCR_AVAILABLE else "Not available (scanned PDFs will be skipped)"

        self.cred_label = ttk.Label(
            info_frame,
            text=(
                f"Credentials: {'Found' if cred_exists else 'MISSING - place client_secret.json in App/required'}"
            ),
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

        # Date filter frame (inside Status)
        filter_frame = ttk.LabelFrame(info_frame, text="Filter Invoices by Date (Optional)", padding=8)
        filter_frame.pack(fill=tk.X, pady=(8, 0))

        self.date_filter_var = tk.BooleanVar(value=False)
        self.today_filter_var = tk.BooleanVar(value=False)
        self.date_from_var = tk.StringVar()
        self.date_to_var = tk.StringVar()

        self.date_filter_check = ttk.Checkbutton(
            filter_frame,
            text="Filter by date range",
            variable=self.date_filter_var,
            command=self._on_date_filter_toggle
        )
        self.date_filter_check.grid(row=0, column=0, sticky='w')

        ttk.Label(filter_frame, text="From (YYYY/MM/DD)").grid(row=0, column=1, padx=(10, 2))
        self.date_from_entry = ttk.Entry(filter_frame, textvariable=self.date_from_var, width=12)
        self.date_from_entry.grid(row=0, column=2, padx=(0, 10))

        ttk.Label(filter_frame, text="To (YYYY/MM/DD)").grid(row=0, column=3, padx=(0, 2))
        self.date_to_entry = ttk.Entry(filter_frame, textvariable=self.date_to_var, width=12)
        self.date_to_entry.grid(row=0, column=4)
        self.date_from_entry.bind(
            "<FocusIn>", lambda _e: self._open_date_picker_at(
                self.date_from_var, self.date_from_entry, "Select FROM Date"
            )
        )
        self.date_to_entry.bind(
            "<FocusIn>", lambda _e: self._open_date_picker_at(
                self.date_to_var, self.date_to_entry, "Select TO Date"
            )
        )

        self.today_filter_check = ttk.Checkbutton(
            filter_frame,
            text="All from Today",
            variable=self.today_filter_var,
            command=self._on_today_filter_toggle
        )
        self.today_filter_check.grid(row=1, column=0, sticky='w', pady=(4, 0))

        ttk.Label(
            filter_frame,
            text="Filtering by date may download already downloaded invoices.",
            foreground='orange'
        ).grid(row=2, column=0, columnspan=5, sticky='w', pady=(4, 0))

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

        ttk.Label(btn_frame, text="-").pack(side=tk.LEFT, padx=(0, 5))

        self.validate_button = ttk.Button(
            btn_frame, text="Validate POs", command=self.start_validation
        )
        self.validate_button.pack(side=tk.LEFT)

        # Batch export buttons (below main actions)
        batch_frame = ttk.Frame(main_frame)
        batch_frame.pack(fill=tk.X, pady=(0, 10))

        self.export_batches_button = ttk.Button(
            batch_frame, text="Export CSV Batches", command=self.export_csv_batches,
            state=tk.DISABLED
        )
        self.export_batches_button.pack(side=tk.LEFT, padx=(0, 5))

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

        # Initialize filter state and batch buttons
        self._update_date_filter_state()
        # Initialize batch buttons based on existing master files
        self._refresh_batch_buttons()
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
        if not os.path.exists(os.path.join(self.required_dir, 'client_secret.json')):
            self.log("ERROR: client_secret.json not found!", "error")
            self.log("Place your Google OAuth credentials file in App/required.", "error")
            return

        self.is_running = True
        self.go_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)

        # Clear log
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)

        # Choose a fresh output file + invoices folder every run
        self.output_file, self.invoices_dir = self._get_next_run_paths()
        self.log(f"Output file: {os.path.basename(self.output_file)}", "info")
        self.log(f"Invoices folder: {os.path.basename(self.invoices_dir)}", "info")

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
                data_dir=self.required_dir,
                invoices_dir=self.invoices_dir,
                expected_email=AUTHORIZED_GMAIL_ACCOUNT
            )
            client.authenticate()

            if not self.is_running:
                self.finish("Stopped by user.")
                return

            self.set_progress(10, "Downloading attachments...")
            query, mode = self._build_gmail_query()
            if query is None:
                self.finish("Failed - invalid date filter.")
                return
            if mode == "label":
                self.log(f"Using Gmail label filter (skipping '{PROCESSED_LABEL_NAME}' emails).", "info")
            elif mode == "today":
                self.log("Downloading all emails from today (label filter ignored).", "warning")
            elif mode == "range":
                self.log("Downloading emails in date range (label filter ignored).", "warning")

            downloaded_files, total_emails, new_emails = (
                client.fetch_and_download_new_attachments(query=query)
            )

            if not self.is_running:
                self.finish("Stopped by user.")
                return

            # Phase 2: Parse invoices
            self.set_progress(40, "Parsing invoices...")
            self.log("", None)
            self.log("=== Phase 2: Parsing invoice files ===", "info")

            def _history_key(bill_no, po_number, vendor, invoice_date):
                parts = [
                    str(bill_no or '').strip(),
                    str(po_number or '').strip(),
                    str(vendor or '').strip(),
                    str(invoice_date or '').strip(),
                ]
                if not any(parts):
                    return ''
                return '|'.join([p.lower() for p in parts])

            history_rows = self._load_invoice_history()
            history_by_po = {}
            history_keys = set()
            for row in history_rows:
                po = str(row.get('po_number', '')).strip()
                if po:
                    history_by_po.setdefault(po, []).append(row)
                key = _history_key(
                    row.get('bill_no', ''),
                    po,
                    row.get('vendor', ''),
                    row.get('invoice_date', '')
                )
                if key:
                    history_keys.add(key)

            new_history_entries = []
            new_history_keys = set()
            folder_name = os.path.basename(self.invoices_dir)
            root_name = os.path.basename(os.path.dirname(self.invoices_dir))

            def _source_file(filename):
                return os.path.join(root_name, folder_name, filename).replace('\\', '/')

            # Find all invoice files to parse
            all_invoice_files = []
            if os.path.exists(self.invoices_dir):
                for f in os.listdir(self.invoices_dir):
                    if f.lower().endswith(('.pdf', '.png', '.jpg', '.jpeg', '.tiff')):
                        all_invoice_files.append(f)

            if not all_invoice_files:
                self.log("No invoice files to parse.", "success")
            else:
                self.log(f"Found {len(all_invoice_files)} invoice files to parse.")

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
                            folder_name = os.path.basename(self.invoices_dir)
                            root_name = os.path.basename(os.path.dirname(self.invoices_dir))
                            invoice_data['source_path'] = os.path.join(
                                root_name, folder_name, filename
                            ).replace('\\', '/')
                            # Write to spreadsheet
                            write_invoice_to_spreadsheet(
                                self.output_file, invoice_data, self.log
                            )
                            success_count += 1
                            bill_no = str(invoice_data.get('invoice_number', '')).strip()
                            po_number = str(invoice_data.get('po_number', '')).strip()
                            vendor = str(invoice_data.get('vendor', '')).strip()
                            invoice_date = str(invoice_data.get('date', '')).strip()
                            key = _history_key(bill_no, po_number, vendor, invoice_date)
                            if key and key not in history_keys and key not in new_history_keys:
                                new_history_entries.append({
                                    'bill_no': bill_no,
                                    'po_number': po_number,
                                    'vendor': vendor,
                                    'invoice_date': invoice_date,
                                    'downloaded_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                                    'source_file': _source_file(filename),
                                })
                                new_history_keys.add(key)
                        else:
                            error_count += 1
                            error_files.append(filename)

                    except Exception as e:
                        self.log(f"  Failed to parse {filename}: {e}", "error")
                        error_count += 1
                        error_files.append(filename)

                # Apply duplicate markers and update history log
                if os.path.exists(self.output_file):
                    self._apply_duplicate_flags(self.output_file, history_by_po)
                if new_history_entries:
                    self._append_invoice_history(new_history_entries)

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
        except WrongAuthorizedAccountError as e:
            self.log(f"Authentication blocked: {e}", "error")
            self.log(
                f"Sign in with {AUTHORIZED_GMAIL_ACCOUNT} and try again.",
                "warning"
            )
            self.finish("Failed - wrong Gmail account.")
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
            self._refresh_batch_buttons()
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

    def export_csv_batches(self):
        """Export the latest master XLSX into CSV batches capped by BATCH_ROW_LIMIT."""
        master_path = self._select_master_for_batching()
        if not master_path or not os.path.exists(master_path):
            self.log("No master spreadsheet found to export.", "warning")
            self._refresh_batch_buttons()
            return

        self.log(f"Exporting CSV batches from {os.path.basename(master_path)}...", "info")
        self.log(
            f"Batch limits: {BATCH_ROW_LIMIT} rows or {BATCH_INVOICE_LIMIT} invoices per file.",
            "info"
        )
        today_tag = f"{datetime.now().month}-{datetime.now().day}"
        if f"Invoices_Master_{today_tag}" not in os.path.basename(master_path):
            self.log(
                "Note: using the most recent master file (not today's date).",
                "warning"
            )

        rows = read_spreadsheet_rows(master_path)
        if not rows:
            self.log("Master spreadsheet has no data rows.", "warning")
            return

        # Group rows by invoice (keep invoices intact)
        invoices = []
        current = []
        current_bill = None
        for row in rows:
            bill_no = str(row.get('bill_no', '')).strip()
            if bill_no:
                if current and bill_no != current_bill:
                    invoices.append(current)
                    current = []
                current_bill = bill_no
            if current_bill is None and not bill_no:
                current_bill = ''
            current.append(row)
        if current:
            invoices.append(current)

        # Build batches without splitting invoices
        batches = []
        batch_rows = 0
        batch_invoices = 0
        current_batch = []
        for inv_rows in invoices:
            inv_count = len(inv_rows)
            if batch_rows and (
                (batch_rows + inv_count) > BATCH_ROW_LIMIT
                or (batch_invoices + 1) > BATCH_INVOICE_LIMIT
            ):
                batches.append(current_batch)
                current_batch = []
                batch_rows = 0
                batch_invoices = 0

            if inv_count > BATCH_ROW_LIMIT and batch_rows == 0:
                # Oversized invoice: put in its own batch
                batches.append(inv_rows)
                continue

            current_batch.extend(inv_rows)
            batch_rows += inv_count
            batch_invoices += 1

        if current_batch:
            batches.append(current_batch)

        date_tag = _extract_date_tag_from_filename(os.path.basename(master_path))
        if not date_tag:
            date_tag = f"{datetime.now().month}-{datetime.now().day}"

        batches_dir = self._get_next_batches_dir(date_tag)
        os.makedirs(batches_dir, exist_ok=True)
        self.last_batches_dir = batches_dir

        headers = [header for _, header in COLUMNS]
        for idx, batch in enumerate(batches, start=1):
            if len(batches) == 1:
                filename = f"{BATCH_FOLDER_PREFIX}{date_tag}.csv"
            else:
                filename = f"{BATCH_FOLDER_PREFIX}{date_tag}_{idx}.csv"
            out_path = os.path.join(batches_dir, filename)
            with open(out_path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(headers)
                for row in batch:
                    writer.writerow([row.get(key, '') for key, _ in COLUMNS])

        # Warn if any invoice exceeded the batch limit
        oversized = [len(inv) for inv in invoices if len(inv) > BATCH_ROW_LIMIT]
        if oversized:
            self.log(
                f"Warning: {len(oversized)} invoice(s) exceeded the batch limit "
                f"({BATCH_ROW_LIMIT} rows) and were exported alone.",
                "warning"
            )

        self.log(
            f"Exported {len(batches)} batch file(s) to {os.path.basename(batches_dir)}",
            "success"
        )
        self._refresh_batch_buttons()

    # Removed open_batches_folder button per UI simplification request.

    def start_validation(self):
        """Start the PO validation pipeline in a background thread."""
        if self.is_running:
            return

        # Check if any output XLSX files exist
        output_files = self._get_output_files_for_validation()
        if not output_files:
            self.log("ERROR: No output XLSX files found - run extraction first!", "error")
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
            config_path = os.path.join(self.required_dir, 'skunexus_config.json')
            if not os.path.exists(config_path):
                config_path = os.path.join(self.app_dir, 'skunexus_config.json')
            if not os.path.exists(config_path):
                config_path = os.path.join(self.base_dir, 'skunexus_config.json')
            if not os.path.exists(config_path):
                self.log("ERROR: skunexus_config.json not found!", "error")
                self.log(
                    "Create skunexus_config.json with 'email' and 'password' in App/required.",
                    "error"
                )
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

            # Collect output files
            self.set_progress(10, "Reading output files...")
            output_files = self._get_output_files_for_validation()
            if not output_files:
                self.log("No output XLSX files found.", "warning")
                self.finish_validation("Validation complete - no data to validate")
                return

            vendor_aliases = load_vendor_aliases(self.app_dir)

            files_info = []
            total_rows = 0
            for filepath in output_files:
                rows = read_spreadsheet_rows(filepath)
                if rows:
                    files_info.append((filepath, rows))
                    total_rows += len(rows)

            if total_rows == 0:
                self.log("No data found in output XLSX files.", "warning")
                self.finish_validation("Validation complete - no data to validate")
                return

            self.log(f"Found {len(files_info)} output file(s) to validate")

            if not self.is_running:
                self.finish_validation("Stopped by user.")
                return

            # Validate rows across all files
            validated_count = 0
            passed_count = 0
            failed_count = 0
            not_found_count = 0
            skipped_count = 0
            already_validated_count = 0
            locked_files_count = 0
            po_cache = {}  # Cache SkuNexus data by PO number

            processed_rows = 0

            def _pick_vendor(vendors):
                if not vendors:
                    return ''
                counts = {}
                for v in vendors:
                    counts[v] = counts.get(v, 0) + 1
                return max(counts, key=counts.get)

            for filepath, rows in files_info:
                if not self.is_running:
                    self.finish_validation("Stopped by user.")
                    return

                basename = os.path.basename(filepath)
                self.log("")
                self.log(f"--- Validating {basename} ---", "info")

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
                    sku = _get_row_sku(row)
                    if sku:
                        group['skus'].append(sku)

                updates = {}
                for row in rows:
                    if not self.is_running:
                        self.finish_validation("Stopped by user.")
                        return

                    processed_rows += 1
                    progress = 15 + (80 * (processed_rows / total_rows))
                    self.set_progress(progress, f"Validating row {processed_rows}/{total_rows}...")

                    row_num = row['_row_num']

                    existing_validation = str(row.get('skunexus_validation', '')).strip()
                    if existing_validation:
                        already_validated_count += 1
                        continue

                    memo = row.get('memo', '')
                    if not memo:
                        bill_no = str(row.get('bill_no', '')).strip()
                        memo = memo_by_bill.get(bill_no, '')
                    category = row.get('category', '')

                    # Skip rows without PO number
                    if not memo:
                        skipped_count += 1
                        continue

                    # Only validate SKU rows (Category/Account = Purchases)
                    if category != 'Purchases':
                        skipped_count += 1
                        continue

                    sku_value = _get_row_sku(row)
                    if not _looks_like_sku(sku_value):
                        skipped_count += 1
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
                        updates[row_num] = (False, ['PO not found in SkuNexus'])
                        validated_count += 1
                        not_found_count += 1
                        continue

                    # Validate this row against SkuNexus data
                    is_valid, failed_fields = validate_po_row(sn_data, row, vendor_aliases)
                    updates[row_num] = (is_valid, failed_fields)
                    validated_count += 1

                    if is_valid:
                        passed_count += 1
                    else:
                        failed_count += 1
                        sku = _get_row_sku(row) or 'N/A'
                        self.log(f"  Row {row_num} (SKU: {sku}) - FAILED: {', '.join(failed_fields)}", "warning")

                if updates:
                    try:
                        write_validation_results(filepath, updates)
                        self.log(f"Updated {len(updates)} row(s) in {basename}", "success")
                    except PermissionError as e:
                        locked_files_count += 1
                        self.log(
                            f"Could not update {basename}: file is locked/open in Excel.",
                            "error"
                        )
                        self.log(f"  {e}", "warning")
                        self.log(
                            "  Close the workbook and run Validate POs again.",
                            "warning"
                        )
                else:
                    self.log(f"No rows needed validation in {basename}", "success")

            # Summary
            self.log("")
            self.log("=== Validation Summary ===", "info")
            self.log(f"Total rows validated: {validated_count}")
            self.log(f"Passed: {passed_count}", "success")
            if failed_count:
                self.log(f"Failed: {failed_count}", "error")
            if not_found_count:
                self.log(f"POs not found: {not_found_count}", "warning")
            if already_validated_count:
                self.log(f"Already validated rows skipped: {already_validated_count}")
            if skipped_count:
                self.log(f"Non-SKU/Non-PO rows skipped: {skipped_count}")
            if locked_files_count:
                self.log(
                    f"Files skipped due to lock/open Excel window: {locked_files_count}",
                    "warning"
                )

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
            self._refresh_batch_buttons()
        self.root.after(0, _update)


def main():
    if not _ensure_single_instance("Invoice Extractor"):
        return
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
