"""Spreadsheet writer: writes parsed invoice data to QuickBooks-compatible CSV/Excel format."""
import os
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill

# Alternating row background color
ALT_ROW_FILL = PatternFill(start_color="EAEAEA", end_color="EAEAEA", fill_type="solid")


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
    ('type', 'Type'),                   # "Item Details" for products, empty for fees
    ('category', 'Category/Account'),   # "Purchases" or "Freight and shipping costs"
    ('product_service', 'Product/Service'),
    ('qty', 'Qty'),
    ('rate', 'Rate'),
    ('description', 'Description'),     # Product description
    ('amount', 'Amount'),
    ('billable', 'Billable'),
    ('customer_project', 'Customer/Project'),
    ('tax_rate', 'Tax Rate'),
    ('class_field', 'Class'),
    ('skunexus_validation', 'SkuNexus Validation'),  # Yes/No
    ('skunexus_failed_fields', 'SkuNexus Failed Fields'),  # Which fields failed
]

# Column indices for validation columns (1-indexed for openpyxl)
VALIDATION_COL = 20  # Column T - SkuNexus Validation
FAILED_FIELDS_COL = 21  # Column U - SkuNexus Failed Fields

PURCHASES_CATEGORY = 'Purchases'
FREIGHT_CATEGORY = 'Freight and shipping costs'


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
            'K': 15, 'L': 6, 'M': 10, 'N': 40, 'O': 12,
            'P': 10, 'Q': 18, 'R': 10, 'S': 12,
            'T': 18, 'U': 40,  # Validation columns
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

    # Write first row with full invoice header + first line item (if any)
    new_row = ws.max_row + 1

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
        return item.get('item_number', '')
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
    row_data = {
        'bill_no': bill_no,
        'vendor': vendor,
        'mailing_address': mailing_address,
        'terms': terms,
        'bill_date': bill_date,
        'due_date': due_date,
        'location': '',
        'memo': memo,
        'type': '' if (first_is_freight or first_is_discount) else ('Item Details' if first_item else ''),
        'category': first_category if first_item else '',
        'product_service': _product_service_for_item(first_item, first_is_discount, first_is_core, first_is_ere, first_is_freight),
        'qty': _normalize_qty_value(first_item.get('quantity', '')),
        'rate': first_item.get('unit_price', ''),
        'description': _description_for_item(first_item, first_is_discount, first_is_core, first_is_freight),
        'amount': '',
        'billable': '',
        'customer_project': customer,
        'tax_rate': '',
        'class_field': '',
    }

    for col_idx, (key, header) in enumerate(COLUMNS, 1):
        value = row_data.get(key, '')
        cell = ws.cell(row=new_row, column=col_idx, value=value)
        # Apply alternating color per invoice (not per row)
        if should_color:
            cell.fill = ALT_ROW_FILL

    rows_written += 1

    # Write additional line items (rows 2+)
    for item in line_items[1:]:
        is_freight = bool(item.get('is_freight'))
        is_discount = _is_discount(item)
        is_core = _is_core(item)
        is_ere = _is_ere(item)
        category = FREIGHT_CATEGORY if is_freight else PURCHASES_CATEGORY
        new_row = ws.max_row + 1
        row_data = {
            'bill_no': bill_no,
            'vendor': '',
            'mailing_address': '',
            'terms': '',
            'bill_date': '',
            'due_date': '',
            'location': '',
            'memo': '',
            'type': '' if (is_freight or is_discount) else 'Item Details',
            'category': category,
            'product_service': _product_service_for_item(item, is_discount, is_core, is_ere, is_freight),
            'qty': _normalize_qty_value(item.get('quantity', '')),
            'rate': item.get('unit_price', ''),
            'description': _description_for_item(item, is_discount, is_core, is_freight),
            'amount': '',
            'billable': '',
            'customer_project': '',
            'tax_rate': '',
            'class_field': '',
        }

        for col_idx, (key, header) in enumerate(COLUMNS, 1):
            value = row_data.get(key, '')
            cell = ws.cell(row=new_row, column=col_idx, value=value)
            # Apply alternating color per invoice (not per row)
            if should_color:
                cell.fill = ALT_ROW_FILL

        rows_written += 1

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
        new_row = ws.max_row + 1
        row_data = {
            'bill_no': bill_no,
            'vendor': '',
            'mailing_address': '',
            'terms': '',
            'bill_date': '',
            'due_date': '',
            'location': '',
            'memo': '',
            'type': '',
            'category': FREIGHT_CATEGORY,
            'product_service': shipping_label,
            'qty': '',
            'rate': shipping_rate,
            'description': shipping_desc,
            'amount': '',
            'billable': '',
            'customer_project': '',
            'tax_rate': '',
            'class_field': '',
        }

        for col_idx, (key, header) in enumerate(COLUMNS, 1):
            value = row_data.get(key, '')
            cell = ws.cell(row=new_row, column=col_idx, value=value)
            # Apply alternating color per invoice (not per row)
            if should_color:
                cell.fill = ALT_ROW_FILL

        rows_written += 1

    # Add final total amount row (summary line)
    if total_amount:
        new_row = ws.max_row + 1
        row_data = {
            'bill_no': bill_no,
            'vendor': '',
            'mailing_address': '',
            'terms': '',
            'bill_date': '',
            'due_date': '',
            'location': '',
            'memo': '',
            'type': '',
            'category': '',
            'product_service': 'Total Amount',
            'qty': '',
            'rate': '',
            'description': '',
            'amount': total_amount,
            'billable': '',
            'customer_project': '',
            'tax_rate': '',
            'class_field': '',
        }

        for col_idx, (key, header) in enumerate(COLUMNS, 1):
            value = row_data.get(key, '')
            cell = ws.cell(row=new_row, column=col_idx, value=value)
            if should_color:
                cell.fill = ALT_ROW_FILL

        rows_written += 1

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

    wb = load_workbook(filepath)
    ws = wb.active

    rows = []
    for row_num in range(2, ws.max_row + 1):  # Skip header
        row_data = {'_row_num': row_num}
        for col_idx, (key, header) in enumerate(COLUMNS, 1):
            row_data[key] = ws.cell(row=row_num, column=col_idx).value or ''
        rows.append(row_data)

    return rows


def write_validation_result(filepath, row_num, is_valid, failed_fields):
    """Write validation result to a specific row.

    Args:
        filepath: Path to the spreadsheet
        row_num: The row number to update (1-indexed, header is row 1)
        is_valid: Boolean - True for Yes, False for No
        failed_fields: List of field names that failed validation
    """
    wb = load_workbook(filepath)
    ws = wb.active

    # Ensure validation headers exist
    if ws.cell(row=1, column=VALIDATION_COL).value != 'SkuNexus Validation':
        ws.cell(row=1, column=VALIDATION_COL, value='SkuNexus Validation')
        ws.cell(row=1, column=VALIDATION_COL).font = ws.cell(row=1, column=VALIDATION_COL).font.copy(bold=True)
    if ws.cell(row=1, column=FAILED_FIELDS_COL).value != 'SkuNexus Failed Fields':
        ws.cell(row=1, column=FAILED_FIELDS_COL, value='SkuNexus Failed Fields')
        ws.cell(row=1, column=FAILED_FIELDS_COL).font = ws.cell(row=1, column=FAILED_FIELDS_COL).font.copy(bold=True)

    # Write validation result
    validation_cell = ws.cell(row=row_num, column=VALIDATION_COL)
    failed_cell = ws.cell(row=row_num, column=FAILED_FIELDS_COL)
    if is_valid is None:
        validation_cell.value = ''
        failed_cell.value = ''
    else:
        validation_cell.value = 'Yes' if is_valid else 'No'
        failed_cell.value = ', '.join(failed_fields) if failed_fields else ''

    # Apply alternating color to validation cells (match existing row color)
    # Check if first cell in row has fill - need to copy the fill, not assign the proxy
    first_cell = ws.cell(row=row_num, column=1)
    if first_cell.fill and first_cell.fill.patternType == 'solid':
        # Create a new PatternFill with the same color
        row_fill = PatternFill(
            start_color=first_cell.fill.start_color.rgb,
            end_color=first_cell.fill.end_color.rgb,
            fill_type='solid'
        )
        validation_cell.fill = row_fill
        failed_cell.fill = row_fill

    wb.save(filepath)


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
