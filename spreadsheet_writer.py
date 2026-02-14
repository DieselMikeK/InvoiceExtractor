"""Spreadsheet writer: writes parsed invoice data to QuickBooks-compatible CSV/Excel format."""
import csv
import os
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill

# Alternating row background color
ALT_ROW_FILL = PatternFill(start_color="EAEAEA", end_color="EAEAEA", fill_type="solid")
SB_DELIVERY_FEE_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")


# QuickBooks Bill Import column definitions (matching Taylor's format)
# Columns marked with * are required
COLUMNS = [
    ('bill_no', 'Bill No.'),           # * Invoice number (e.g., I457156)
    ('vendor', '*Vendor'),              # * Required - vendor name (e.g., S&B)
    ('mailing_address', 'Mailing Address'),
    ('terms', 'Terms'),                 # e.g., Net 30
    ('bill_date', '*Bill Date'),        # * Required - invoice date
    ('due_date', 'Due Date'),
    ('location', 'Location'),
    ('memo', 'Memo'),                   # PO number goes here
    ('type', 'Type'),                   # Always "Category Details" for QuickBooks import
    ('category', 'Category/Account'),   # "Purchases" or "Freight and shipping costs"
    ('product_service', 'Product/Service'),
    ('sku', 'SKU'),
    ('qty', 'Qty'),
    ('rate', 'Rate'),
    ('description', 'Description'),     # Product description
    ('amount', 'Amount'),
    ('billable', 'Billable'),
    ('customer_project', 'Customer/Project'),
    ('tax_rate', 'Tax Rate'),
    ('class_field', 'Class'),
    ('duplicate_status', 'Duplicate Status'),
    ('duplicate_reference', 'Duplicate Reference'),
    ('skunexus_validation', 'SkuNexus Validation'),  # Yes/No
    ('skunexus_failed_fields', 'SkuNexus Failed Fields'),  # Which fields failed
]

# Column indices for validation columns (1-indexed for openpyxl)
VALIDATION_COL = 23  # Column W - SkuNexus Validation
FAILED_FIELDS_COL = 24  # Column X - SkuNexus Failed Fields

PURCHASES_CATEGORY = 'Purchases'
FREIGHT_CATEGORY = 'Freight and shipping costs'
TYPE_CATEGORY = 'Category Details'
TYPE_ITEM = 'Item Details'


def _is_csv(filepath):
    return str(filepath).lower().endswith('.csv')


def _get_csv_writer(filepath):
    file_exists = os.path.exists(filepath) and os.path.getsize(filepath) > 0
    csv_file = open(filepath, 'a', newline='', encoding='utf-8')
    writer = csv.writer(csv_file)
    if not file_exists:
        writer.writerow([header for _, header in COLUMNS])
    return csv_file, writer


def get_or_create_workbook(filepath):
    """Load existing workbook or create a new one with headers."""
    if os.path.exists(filepath):
        wb = load_workbook(filepath)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Bills"
        # Write header row
        for col_idx, (key, header) in enumerate(COLUMNS, 1):
            cell = ws.cell(row=1, column=col_idx, value=header)
            cell.font = cell.font.copy(bold=True)
        # Set reasonable column widths
        widths = {
            'A': 12, 'B': 12, 'C': 25, 'D': 10, 'E': 12,
            'F': 12, 'G': 12, 'H': 15, 'I': 14, 'J': 18,
            'K': 18, 'L': 15, 'M': 6, 'N': 10, 'O': 40, 'P': 12,
            'Q': 10, 'R': 18, 'S': 10, 'T': 12,
            'U': 20, 'V': 40,  # Duplicate columns
            'W': 18, 'X': 40,  # Validation columns
        }
        for col_letter, width in widths.items():
            ws.column_dimensions[col_letter].width = width

    return wb, ws


def count_unique_invoices(ws):
    """Count unique Bill No. values in the worksheet to determine invoice index."""
    bill_nos = set()
    for row in range(2, ws.max_row + 1):  # Skip header row
        bill_no = ws.cell(row=row, column=1).value  # Column A is Bill No.
        if bill_no:
            bill_nos.add(bill_no)
    return len(bill_nos)


def write_invoice_rows(filepath, invoice_data, status_callback=None):
    """Write invoice data as multiple rows (one per line item + shipping).

    Format matches QuickBooks Bill Import:
    - First row: full invoice info with first line item
    - Additional rows: just Bill No., Type, Category, Description, Amount for each additional item
    - Shipping row: Bill No., Category="Freight and shipping costs", Description="Shipping", Amount

    Args:
        filepath: Path to the .xlsx file
        invoice_data: Dict with parsed invoice fields including 'line_items' list
        status_callback: Optional function(msg, tag) for status updates

    Returns:
        int: Number of rows written
    """
    cb = status_callback or (lambda msg, tag=None: None)

    is_csv = _is_csv(filepath)
    wb = None
    ws = None
    csv_file = None
    csv_writer = None
    should_color = False

    if is_csv:
        csv_file, csv_writer = _get_csv_writer(filepath)
    else:
        wb, ws = get_or_create_workbook(filepath)
        # Count existing invoices to determine if this invoice should be colored
        # Even-indexed invoices (0, 2, 4...) get no color, odd-indexed (1, 3, 5...) get gray
        invoice_index = count_unique_invoices(ws)
        should_color = (invoice_index % 2 == 1)

    bill_no = invoice_data.get('invoice_number', '')
    vendor = invoice_data.get('vendor', '')
    mailing_address = invoice_data.get('vendor_address', '')
    terms = invoice_data.get('terms', '')
    bill_date = invoice_data.get('date', '')
    due_date = invoice_data.get('due_date', '')
    memo = invoice_data.get('po_number', '')
    customer = invoice_data.get('customer', '')
    total_amount = invoice_data.get('total', '')

    line_items = invoice_data.get('line_items', [])
    shipping_cost = invoice_data.get('shipping_cost', '')

    rows_written = 0

    def _write_row(row_data):
        nonlocal rows_written
        row_num = None
        row_fill = row_data.get('_row_fill')
        if row_fill is None and should_color:
            row_fill = ALT_ROW_FILL
        if is_csv:
            csv_writer.writerow([row_data.get(key, '') for key, _ in COLUMNS])
        else:
            row_num = ws.max_row + 1
            for col_idx, (key, header) in enumerate(COLUMNS, 1):
                value = row_data.get(key, '')
                cell = ws.cell(row=row_num, column=col_idx, value=value)
                # Apply alternating color per invoice (not per row)
                if row_fill:
                    cell.fill = row_fill
        rows_written += 1
        return row_num

    # Write first row with full invoice header + first line item (if any)

    first_item = line_items[0] if line_items else {}

    def _normalize_qty_value(value):
        if value is None:
            return ''
        s = str(value).strip()
        if not s:
            return ''
        try:
            num = float(s.replace(',', ''))
        except (ValueError, TypeError):
            return s
        # If it's effectively an integer, drop decimals
        if abs(num - round(num)) < 1e-9:
            return str(int(round(num)))
        return s

    def _is_discount(item):
        item_num = str(item.get('item_number', '')).lower()
        desc = str(item.get('description', '')).lower()
        return bool(item.get('is_discount')) or ('discount' in item_num) or ('discount' in desc)
    def _is_core(item):
        item_num = str(item.get('item_number', '')).lower()
        desc = str(item.get('description', '')).lower()
        return item_num == 'core' or item_num.startswith('core ') or item_num.startswith('core-') or desc.startswith('core ')
    def _is_ere(item):
        item_num = str(item.get('item_number', '')).lower().strip()
        desc = str(item.get('description', '')).lower()
        return item_num in ('e.r.e.', 'ere') or 'environmental regulation expense' in desc
    def _normalize_shipping_label(text):
        s = str(text or '').lower()
        if 'drop ship' in s or 'dropship' in s:
            return 'Drop Ship'
        if 'freight' in s or 'frieght' in s:
            return 'Freight'
        if 'ship' in s:
            return 'Shipping'
        return 'Shipping'
    def _core_description(item):
        code = str(item.get('item_number', '')).strip()
        desc = str(item.get('description', '')).strip()
        if code and desc:
            if code.lower() in desc.lower():
                return desc
            return f"{code} {desc}".strip()
        return code or desc
    def _product_service_for_item(item, is_discount, is_core, is_ere, is_freight):
        if is_discount:
            return 'DPP Discount'
        if is_core:
            return 'Core'
        if is_ere:
            return 'E.R.E.'
        if is_freight:
            return _normalize_shipping_label(item.get('description') or item.get('item_number'))
        return 'Inventory Item (Sellable Item)'
    def _sku_for_item(item):
        return str(item.get('item_number', '')).strip()
    def _description_for_item(item, is_discount, is_core, is_freight):
        if is_core:
            return _core_description(item)
        if is_discount:
            return ''
        if is_freight:
            return (str(item.get('description', '')).strip()
                    or str(item.get('item_number', '')).strip())
        return item.get('description', '')

    first_is_freight = bool(first_item.get('is_freight'))
    first_is_discount = _is_discount(first_item)
    first_is_core = _is_core(first_item)
    first_is_ere = _is_ere(first_item)
    first_category = FREIGHT_CATEGORY if first_is_freight else PURCHASES_CATEGORY
    first_type = TYPE_ITEM if first_item else TYPE_CATEGORY
    if first_item:
        first_product_service = _product_service_for_item(first_item, first_is_discount, first_is_core, first_is_ere, first_is_freight)
        first_sku = _sku_for_item(first_item)
    else:
        first_product_service = ''
        first_sku = ''
    row_data = {
        'bill_no': bill_no,
        'vendor': vendor,
        'mailing_address': mailing_address,
        'terms': terms,
        'bill_date': bill_date,
        'due_date': due_date,
        'location': '',
        'memo': memo,
        'type': first_type,
        'category': first_category if first_item else '',
        'product_service': first_product_service,
        'sku': first_sku,
        'qty': _normalize_qty_value(first_item.get('quantity', '')),
        'rate': first_item.get('unit_price', ''),
        'description': _description_for_item(first_item, first_is_discount, first_is_core, first_is_freight),
        'amount': '',
        'billable': '',
        'customer_project': customer,
        'tax_rate': '',
        'class_field': '',
    }
    if first_item and first_item.get('sb_delivery_fee'):
        row_data['_row_fill'] = SB_DELIVERY_FEE_FILL

    first_row_num = _write_row(row_data)
    source_path = invoice_data.get('source_path') or ''
    if (not is_csv) and first_row_num and bill_no and source_path:
        try:
            link_cell = ws.cell(row=first_row_num, column=1)
            link_cell.hyperlink = source_path
        except Exception:
            pass

    # Write additional line items (rows 2+)
    for item in line_items[1:]:
        is_freight = bool(item.get('is_freight'))
        is_discount = _is_discount(item)
        is_core = _is_core(item)
        is_ere = _is_ere(item)
        category = FREIGHT_CATEGORY if is_freight else PURCHASES_CATEGORY
        row_data = {
            'bill_no': '',
            'vendor': '',
            'mailing_address': '',
            'terms': '',
            'bill_date': '',
            'due_date': '',
            'location': '',
            'memo': '',
            'type': TYPE_ITEM,
            'category': category,
            'product_service': _product_service_for_item(item, is_discount, is_core, is_ere, is_freight),
            'sku': _sku_for_item(item),
            'qty': _normalize_qty_value(item.get('quantity', '')),
            'rate': item.get('unit_price', ''),
            'description': _description_for_item(item, is_discount, is_core, is_freight),
            'amount': '',
            'billable': '',
            'customer_project': '',
            'tax_rate': '',
            'class_field': '',
        }
        if item.get('sb_delivery_fee'):
            row_data['_row_fill'] = SB_DELIVERY_FEE_FILL

        _write_row(row_data)

    # Always write shipping row (even if $0), unless freight items already exist
    has_freight_item = any(item.get('is_freight') for item in line_items)

    try:
        shipping_val = float(str(shipping_cost).replace(',', '').replace('$', ''))
    except (ValueError, TypeError):
        shipping_val = 0

    # Format shipping amount - show 0 if no shipping cost
    shipping_rate = shipping_cost if shipping_val > 0 else '0'
    shipping_desc = invoice_data.get('shipping_description', 'Shipping')
    shipping_label = _normalize_shipping_label(shipping_desc)

    if (not has_freight_item) and (shipping_rate or shipping_desc):
        row_data = {
            'bill_no': '',
            'vendor': '',
            'mailing_address': '',
            'terms': '',
            'bill_date': '',
            'due_date': '',
            'location': '',
            'memo': '',
            'type': TYPE_CATEGORY,
            'category': FREIGHT_CATEGORY,
            'product_service': shipping_label,
            'sku': '',
            'qty': '',
            'rate': shipping_rate,
            'description': shipping_desc,
            'amount': '',
            'billable': '',
            'customer_project': '',
            'tax_rate': '',
            'class_field': '',
        }

        _write_row(row_data)

    # Add final total amount row (summary line)
    if total_amount:
        row_data = {
            'bill_no': '',
            'vendor': '',
            'mailing_address': '',
            'terms': '',
            'bill_date': '',
            'due_date': '',
            'location': '',
            'memo': '',
            'type': TYPE_CATEGORY,
            'category': '',
            'product_service': 'Total Amount',
            'sku': '',
            'qty': '',
            'rate': '',
            'description': '',
            'amount': total_amount,
            'billable': '',
            'customer_project': '',
            'tax_rate': '',
            'class_field': '',
        }

        _write_row(row_data)

    if is_csv:
        if csv_file:
            csv_file.close()
    else:
        wb.save(filepath)
    cb(f"  Written {rows_written} row(s) to spreadsheet for invoice {bill_no}", "success")

    return rows_written


# Keep old function for backwards compatibility but redirect to new one
def write_invoice_to_spreadsheet(filepath, invoice_data, status_callback=None):
    """Append invoice data to spreadsheet (wrapper for write_invoice_rows)."""
    return write_invoice_rows(filepath, invoice_data, status_callback)


def read_spreadsheet_rows(filepath):
    """Read all data rows from the spreadsheet.

    Returns:
        list: List of dicts, one per row, with column keys from COLUMNS
    """
    if not os.path.exists(filepath):
        return []

    if _is_csv(filepath):
        rows = []
        with open(filepath, newline='', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            for idx, row in enumerate(reader, start=2):  # Header is row 1
                row_data = {'_row_num': idx}
                for key, header in COLUMNS:
                    row_data[key] = row.get(header, '') or ''
                rows.append(row_data)
        return rows

    wb = load_workbook(filepath)
    ws = wb.active

    rows = []
    for row_num in range(2, ws.max_row + 1):  # Skip header
        row_data = {'_row_num': row_num}
        for col_idx, (key, header) in enumerate(COLUMNS, 1):
            row_data[key] = ws.cell(row=row_num, column=col_idx).value or ''
        rows.append(row_data)

    return rows


def _ensure_validation_headers(ws):
    """Ensure validation header columns exist in the worksheet."""
    if ws.cell(row=1, column=VALIDATION_COL).value != 'SkuNexus Validation':
        ws.cell(row=1, column=VALIDATION_COL, value='SkuNexus Validation')
        ws.cell(row=1, column=VALIDATION_COL).font = ws.cell(row=1, column=VALIDATION_COL).font.copy(bold=True)
    if ws.cell(row=1, column=FAILED_FIELDS_COL).value != 'SkuNexus Failed Fields':
        ws.cell(row=1, column=FAILED_FIELDS_COL, value='SkuNexus Failed Fields')
        ws.cell(row=1, column=FAILED_FIELDS_COL).font = ws.cell(row=1, column=FAILED_FIELDS_COL).font.copy(bold=True)


def _apply_validation_to_ws(ws, row_num, is_valid, failed_fields):
    """Apply validation values to a single row in an open worksheet."""
    validation_cell = ws.cell(row=row_num, column=VALIDATION_COL)
    failed_cell = ws.cell(row=row_num, column=FAILED_FIELDS_COL)
    if is_valid is None:
        validation_cell.value = ''
        failed_cell.value = ''
    else:
        validation_cell.value = 'Yes' if is_valid else 'No'
        failed_cell.value = ', '.join(failed_fields) if failed_fields else ''

    # Apply alternating color to validation cells (match existing row color)
    first_cell = ws.cell(row=row_num, column=1)
    if first_cell.fill and first_cell.fill.patternType == 'solid':
        row_fill = PatternFill(
            start_color=first_cell.fill.start_color.rgb,
            end_color=first_cell.fill.end_color.rgb,
            fill_type='solid'
        )
        validation_cell.fill = row_fill
        failed_cell.fill = row_fill


def _write_validation_results_csv(filepath, updates):
    if not updates or not os.path.exists(filepath):
        return

    with open(filepath, newline='', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        headers = list(reader.fieldnames or [])
        rows = list(reader)

    if not headers:
        headers = [header for _, header in COLUMNS]

    for header in ('SkuNexus Validation', 'SkuNexus Failed Fields'):
        if header not in headers:
            headers.append(header)

    for row_num, (is_valid, failed_fields) in updates.items():
        idx = row_num - 2  # Convert 1-indexed row number (header is row 1)
        if idx < 0 or idx >= len(rows):
            continue
        if is_valid is None:
            validation_value = ''
            failed_value = ''
        else:
            validation_value = 'Yes' if is_valid else 'No'
            failed_value = ', '.join(failed_fields) if failed_fields else ''
        rows[idx]['SkuNexus Validation'] = validation_value
        rows[idx]['SkuNexus Failed Fields'] = failed_value

    with open(filepath, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=headers)
        writer.writeheader()
        for row in rows:
            writer.writerow({h: row.get(h, '') for h in headers})


def _write_validation_results_xlsx(filepath, updates):
    if not updates:
        return
    wb = load_workbook(filepath)
    ws = wb.active
    _ensure_validation_headers(ws)
    for row_num, (is_valid, failed_fields) in updates.items():
        _apply_validation_to_ws(ws, row_num, is_valid, failed_fields)
    wb.save(filepath)


def write_validation_results(filepath, updates):
    """Write multiple validation results in one pass.

    Args:
        filepath: Path to the spreadsheet
        updates: dict {row_num: (is_valid, failed_fields)}
    """
    if _is_csv(filepath):
        _write_validation_results_csv(filepath, updates)
    else:
        _write_validation_results_xlsx(filepath, updates)


def write_validation_result(filepath, row_num, is_valid, failed_fields):
    """Write validation result to a specific row.

    Args:
        filepath: Path to the spreadsheet
        row_num: The row number to update (1-indexed, header is row 1)
        is_valid: Boolean - True for Yes, False for No
        failed_fields: List of field names that failed validation
    """
    write_validation_results(filepath, {row_num: (is_valid, failed_fields)})


def get_unique_po_numbers(filepath):
    """Get unique PO numbers from the spreadsheet.

    Returns:
        dict: {po_number: [list of row_nums with that PO]}
    """
    rows = read_spreadsheet_rows(filepath)
    po_rows = {}

    for row in rows:
        memo = row.get('memo', '')
        if memo:
            # Normalize PO number
            po_num = str(memo).strip()
            if po_num not in po_rows:
                po_rows[po_num] = []
            po_rows[po_num].append(row)

    return po_rows
