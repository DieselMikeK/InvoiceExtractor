"""Invoice parser: extracts structured data from PDF invoices."""
import os
import re
import pdfplumber

# Try to import OCR dependencies (optional)
try:
    import pytesseract
    from PIL import Image

    # Set Tesseract path for Windows if not on PATH
    TESSERACT_PATH = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
    if os.path.exists(TESSERACT_PATH):
        pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False

# Try pypdfium2 for rendering PDF pages to images (no poppler needed)
try:
    import pypdfium2 as pdfium
    PDFIUM_AVAILABLE = True
except ImportError:
    PDFIUM_AVAILABLE = False


# Known vendors for fallback detection (add more as encountered)
KNOWN_VENDORS = ['S&B', 'S & B', 'Diesel Power Products']


def validate_vendor_name(text):
    """Check if extracted text looks like a valid vendor name."""
    if not text or len(text) < 2 or len(text) > 80:
        return False
    # Should have at least some letters
    if not re.search(r'[A-Za-z]', text):
        return False
    # Reject obvious garbage (form fields, underscores, credit card labels, etc.)
    if re.search(r'_{3,}|Credit Card|Type:|Authorize|Please Enter', text, re.IGNORECASE):
        return False
    return True


def extract_text_from_pdf(filepath):
    """Extract text from a PDF using pdfplumber (text-based PDFs)."""
    text = ""
    try:
        with pdfplumber.open(filepath) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"
    except Exception:
        pass
    return text.strip()


def extract_text_with_ocr(filepath):
    """Extract text from a scanned PDF or image using OCR (Tesseract).

    Uses pypdfium2 to render PDF pages (no poppler dependency needed).
    Also handles image files directly (PNG, JPG, TIFF).
    """
    if not OCR_AVAILABLE:
        return ""

    ext = os.path.splitext(filepath)[1].lower()

    # Handle image files directly
    if ext in ('.png', '.jpg', '.jpeg', '.tiff', '.tif', '.bmp'):
        try:
            img = Image.open(filepath)
            return pytesseract.image_to_string(img).strip()
        except Exception:
            return ""

    # Handle PDF files via pypdfium2
    if not PDFIUM_AVAILABLE:
        return ""

    try:
        pdf = pdfium.PdfDocument(filepath)
        text = ""
        for page_index in range(len(pdf)):
            page = pdf[page_index]
            # Render at 300 DPI for good OCR quality
            bitmap = page.render(scale=300 / 72)
            pil_image = bitmap.to_pil()
            page_text = pytesseract.image_to_string(pil_image)
            if page_text:
                text += page_text + "\n"
            page.close()
        pdf.close()
        return text.strip()
    except Exception:
        return ""


def extract_tables_from_pdf(filepath):
    """Extract tables from PDF using pdfplumber's table extraction."""
    tables = []
    try:
        with pdfplumber.open(filepath) as pdf:
            for page in pdf.pages:
                page_tables = page.extract_tables()
                if page_tables:
                    tables.extend(page_tables)
    except Exception:
        pass
    return tables


def parse_field(text, patterns, default=""):
    """Try multiple regex patterns to extract a field value."""
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
        if match:
            return match.group(1).strip()
    return default


def detect_vendor(text):
    """Detect vendor name using priority-based multi-strategy approach.

    Tries detection methods in order of confidence, returns first valid match.
    This is designed to work with various invoice formats, not just S&B.

    Note: "Vendor" = who issued the invoice (who you pay money TO).
    This is different from "Customer" = who the invoice is FOR.
    """

    # Strategy 1: "Make Checks Payable To" (highest confidence - directly identifies payee)
    # Use [^\n] to avoid matching across lines, but allow spaces
    payable_patterns = [
        r'Make\s+Checks\s+Payable\s+To\s+([A-Za-z][A-Za-z0-9 &]+?)(?:\n|\d{4,}|$)',
        r'Make\s+Checks\s+Payable\s+To\s*\n\s*([A-Za-z][A-Za-z0-9 &]+?)(?:\n|$)',
    ]
    for pattern in payable_patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            vendor = match.group(1).strip()
            if validate_vendor_name(vendor):
                return vendor

    # Strategy 2: "Remit To" / "Pay To" (high confidence - payment destination)
    remit_match = re.search(
        r'(?:Remit|Pay)\s+To[:\s]+([A-Za-z][A-Za-z0-9 &\-\.]+?)(?:\n|Address|,)',
        text, re.IGNORECASE
    )
    if remit_match:
        vendor = remit_match.group(1).strip()
        if validate_vendor_name(vendor):
            return vendor

    # Strategy 3: "Bill From" / "Vendor" / "Supplier" / "Sold By" / "Invoice From"
    from_patterns = [
        r'(?:Bill\s+From|Vendor|Supplier|Sold\s+By|Invoice\s+From)[:\s]+([A-Za-z][A-Za-z0-9 &\-\.]+?)(?:\n|Address|Tel|Phone|,)',
    ]
    for pattern in from_patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            vendor = match.group(1).strip()
            if validate_vendor_name(vendor):
                return vendor

    # Strategy 4: Known vendors list (case-insensitive exact match)
    for vendor in KNOWN_VENDORS:
        pattern = r'\b' + re.escape(vendor) + r'\b'
        if re.search(pattern, text, re.IGNORECASE):
            return vendor

    # Strategy 5: Header heuristic - look for capitalized company names in top portion
    top_text = text[:len(text) // 5]  # Top 20%
    company_match = re.search(
        r'^([A-Z][A-Za-z0-9 &\-\.]{2,40}(?:Inc|LLC|Corp|Ltd|Co)?\.?)$',
        top_text, re.MULTILINE
    )
    if company_match:
        vendor = company_match.group(1).strip()
        if validate_vendor_name(vendor):
            return vendor

    return ""


def extract_customer_name(text):
    """Extract customer name (who the invoice is billed TO).

    This is different from vendor (who issued the invoice).
    For S&B invoices, the customer is typically "Diesel Power Products".

    Strategies:
    1. "Customer 505 Diesel Power Products" pattern in remittance section
    2. "Bill To" section - company name after the label
    3. "Billed To" / "Invoice To" patterns
    """

    # Strategy 1: "Customer XXX Name" pattern (most reliable for S&B)
    customer_match = re.search(
        r'Customer\s+\d+\s+([A-Za-z][A-Za-z0-9 &\-\.]+)',
        text, re.IGNORECASE
    )
    if customer_match:
        customer = customer_match.group(1).strip()
        if len(customer) >= 3:
            return customer

    # Strategy 2: "Bill To" section - look for company name
    # Format: "Bill To" followed by company name (often "Diesel Power Products DBA...")
    bill_to_match = re.search(
        r'Bill\s+To[:\s]+(?:FOB\s+)?(?:Shipping.*?\n)?[A-Za-z ]+\s+([A-Za-z][A-Za-z0-9 &]+(?:DBA[A-Za-z0-9 &\.]+)?)',
        text, re.IGNORECASE | re.DOTALL
    )
    if bill_to_match:
        customer = bill_to_match.group(1).strip()
        # Clean up - remove trailing "DBA Power ..." truncation
        customer = re.sub(r'\s+DBA\s+Power\s*\.{3}.*', '', customer)
        if len(customer) >= 3 and customer not in ['FOB', 'Shipping']:
            return customer

    # Strategy 3: "Billed To" / "Invoice To"
    billed_match = re.search(
        r'(?:Billed\s+To|Invoice\s+To)[:\s]+([A-Za-z][A-Za-z0-9 &\-\.]+)',
        text, re.IGNORECASE
    )
    if billed_match:
        customer = billed_match.group(1).strip()
        if len(customer) >= 3:
            return customer

    return ""


def parse_invoice_text(text):
    """Parse extracted text into structured invoice data."""
    data = {}

    # Invoice Number - look for "Invoice #" or "Invoice # I457156" pattern
    data['invoice_number'] = parse_field(text, [
        r'Invoice\s*#\s*([A-Z]?\d+)',      # "Invoice # I457156" or "Invoice # 457156"
        r'Invoice\s+#\s*:?\s*([A-Z]?\d+)',
        r'Invoice\s+Number\s*:?\s*([A-Z]?\d+)',
    ])

    # Vendor detection (use dedicated function)
    data['vendor'] = detect_vendor(text)

    # Vendor address - for S&B, use the known address at top of invoice
    # Look for address after "15461 Slover Avenue"
    address_match = re.search(
        r'(15461\s+Slover\s+Avenue\s*\n\s*Fontana\s+CA\s+\d+)',
        text, re.IGNORECASE
    )
    if address_match:
        data['vendor_address'] = address_match.group(1).strip().replace('\n', ', ')
    else:
        data['vendor_address'] = ''

    # Customer name - who the invoice is billed TO (our customer)
    # Strategy 1: "Customer 505 Diesel Power Products" pattern in remittance section
    # Strategy 2: "Bill To" section
    data['customer'] = extract_customer_name(text)

    # Date - "Date 1/27/2026" format
    data['date'] = parse_field(text, [
        r'^Date\s+(\d{1,2}/\d{1,2}/\d{4})',  # "Date 1/27/2026" at line start
        r'\nDate\s+(\d{1,2}/\d{1,2}/\d{4})',  # "Date 1/27/2026" after newline
        r'Date\s*:?\s*(\d{1,2}/\d{1,2}/\d{4})',
    ])

    # Due Date
    data['due_date'] = parse_field(text, [
        r'Due\s+Date\s+(\d{1,2}/\d{1,2}/\d{4})',
        r'Due\s+Date\s*:?\s*(\d{1,2}/\d{1,2}/\d{4})',
    ])

    # Terms (e.g., "Net 30")
    data['terms'] = parse_field(text, [
        r'Terms\s+(Net\s*\d+)',
        r'Terms\s*:?\s*(Net\s*\d+)',
    ])

    # PO Number
    data['po_number'] = parse_field(text, [
        r'PO\s*#\s*(\d+)',
        r'P\.?O\.?\s*#?\s*:?\s*(\d+)',
    ])

    # Tracking Number
    data['tracking_number'] = parse_field(text, [
        r'Tracking#?\s*(?:Notes:)?\s*\n?\s*(\d+)',
        r'Tracking\s*#?\s*:?\s*(\d+)',
    ])

    # Shipping Method
    data['shipping_method'] = parse_field(text, [
        r'Shipping\s+Method\s+([^\n]+)',
        r'Shipping\s+Method\s*:?\s*([^\n]+)',
    ])

    # Ship Date
    data['ship_date'] = parse_field(text, [
        r'Ship\s+Date\s+(\d{1,2}/\d{1,2}/\d{4})',
        r'Ship\s+Date\s*:?\s*(\d{1,2}/\d{1,2}/\d{4})',
    ])

    # Shipping Tax Code
    data['shipping_tax_code'] = parse_field(text, [
        r'Shipping\s+Tax\s+Code\s+(\S+)',
    ])

    # Shipping Tax Rate
    data['shipping_tax_rate'] = parse_field(text, [
        r'Shipping\s+Tax\s+Rate\s+(\d+)',
    ])

    # Subtotal
    data['subtotal'] = parse_field(text, [
        r'Subtotal\s+(\d+\.?\d*)',
        r'Sub\s*-?\s*total\s*:?\s*\$?([\d,]+\.?\d*)',
    ])

    # Shipping cost - look for "Shipping Cost (FedEx...) 12.00" pattern
    data['shipping_cost'] = parse_field(text, [
        r'Shipping\s+Cost\s*\([^)]+\)\s*(\d+\.?\d*)',
        r'Shipping\s+Cost\s+(\d+\.?\d*)',
        r'Shipping\s*:?\s*\$?([\d,]+\.?\d*)',
    ])

    # Total
    data['total'] = parse_field(text, [
        r'Total\s+\$?([\d,]+\.?\d*)',
        r'Amount\s+Due\s+\$?([\d,]+\.?\d*)',
    ])

    # Line items
    data['line_items'] = extract_line_items_sb(text)

    return data


def extract_line_items(text):
    """Extract line items from invoice text.

    Strategy:
    1. PRIMARY: Find a table structure with headers (Item, Qty, Description, Price, Amount, etc.)
       and extract row data from it
    2. FALLBACK: If no table found, look for repeated patterns with prices

    Maps found data to our standard fields:
    - item_number (Product/Service)
    - quantity
    - units
    - description
    - unit_price
    - amount
    """
    items = []

    # === PRIMARY: Table-based extraction ===
    # Find the items table section by looking for a header row followed by Subtotal

    # Look for header row containing table column names
    # Common headers: Item, Qty/Quantity, Units, Description, Unit Price/Price/Rate, Amount/Total
    header_patterns = [
        r'(Item\s+.*?(?:Amount|Total|Price))\s*\n',  # "Item ... Amount"
        r'((?:SKU|Product|Part)\s+.*?(?:Amount|Total|Price))\s*\n',  # "SKU ... Amount"
        r'(Description\s+.*?(?:Amount|Total|Price))\s*\n',  # "Description ... Amount"
        r'(Qty\s+.*?(?:Amount|Total|Price))\s*\n',  # "Qty ... Amount"
    ]

    items_section = ""
    header_found = False

    for pattern in header_patterns:
        header_match = re.search(pattern, text, re.IGNORECASE)
        if header_match:
            header_found = True
            header_end = header_match.end()

            # Find end of items section (Subtotal, Total $, Shipping Cost, etc.)
            end_match = re.search(
                r'(Subtotal|Sub\s*-?\s*total|^Total\s+\$|Shipping\s+Cost|Tax\s+\d)',
                text[header_end:],
                re.IGNORECASE | re.MULTILINE
            )

            if end_match:
                items_section = text[header_end:header_end + end_match.start()]
            else:
                items_section = text[header_end:header_end + 2000]
            break

    # If we found a table section, parse it
    if items_section.strip():
        items = parse_table_rows(items_section)

    # === FALLBACK: Pattern-based extraction ===
    # If no items found via table, try to find price patterns
    if not items:
        items = extract_items_by_price_patterns(text)

    return items


def parse_table_rows(items_section):
    """Parse rows from a table section.

    Each row typically ends with two decimal numbers (unit_price, amount).
    Format: SKU QTY [UNITS] DESCRIPTION UNIT_PRICE AMOUNT
    """
    items = []
    lines = items_section.strip().split('\n')
    accumulated = ""

    for line in lines:
        line = line.strip()
        if not line:
            continue

        # Skip if this looks like a header or summary line
        if re.match(r'^(item|sku|qty|description|subtotal|total)\s*$', line, re.IGNORECASE):
            continue

        # Accumulate lines (descriptions can wrap to multiple lines)
        accumulated = (accumulated + " " + line).strip() if accumulated else line

        # Check if line ends with price pattern: XX.XX XX.XX (unit_price amount)
        price_match = re.search(r'^(.+?)\s+(\d+\.?\d{2})\s+(\d+\.?\d{2})\s*$', accumulated)

        if price_match:
            content = price_match.group(1)
            unit_price = price_match.group(2)
            amount = price_match.group(3)

            # Parse the content: SKU QTY [UNITS] DESCRIPTION
            item = parse_row_content(content, unit_price, amount)
            if item:
                items.append(item)

            accumulated = ""  # Reset for next row

    return items


def parse_row_content(content, unit_price, amount):
    """Parse row content to extract item_number, quantity, units, description.

    Expected format: SKU QTY [UNITS] DESCRIPTION
    Examples:
    - "75-5068 1 Each 13-18 Dodge Ram 2500 / 3500 L6-6.7L"
    - "WF-1037 1 Each Filter Wrap for KF-1037"
    - "ABC123 2 Widget product description here"
    """
    content = re.sub(r'\s+', ' ', content).strip()

    if len(content) < 3:
        return None

    # Pattern: SKU (first token) + QTY (number) + optional UNITS + DESCRIPTION (rest)
    match = re.match(
        r'^(\S+)\s+(\d+)\s*(Each|EA|each|ea|pc|pcs|units?)?\s*(.*)$',
        content,
        re.IGNORECASE
    )

    if match:
        sku = match.group(1)
        qty = match.group(2)
        units = match.group(3) or 'Each'
        desc = match.group(4).strip()

        # Only accept if we have a description
        if desc and len(desc) >= 3:
            return {
                'item_number': sku,
                'quantity': qty,
                'units': units,
                'description': desc,
                'unit_price': unit_price,
                'amount': amount,
            }

    # Alternative: Maybe no clear SKU, or different order
    # Try: QTY first, then description
    match = re.match(r'^(\d+)\s*(Each|EA|each|ea)?\s+(.+)$', content)
    if match:
        qty = match.group(1)
        desc = match.group(3).strip()

        # Try to extract SKU from start of description
        desc_match = re.match(r'^(\S+)\s+(.+)$', desc)
        if desc_match:
            sku = desc_match.group(1)
            desc = desc_match.group(2)
        else:
            sku = ''

        return {
            'item_number': sku,
            'quantity': qty,
            'units': 'Each',
            'description': desc,
            'unit_price': unit_price,
            'amount': amount,
        }

    return None


def extract_items_by_price_patterns(text):
    """Fallback: Extract items by finding lines with price patterns.

    Looks for any line containing: SOMETHING ... PRICE AMOUNT
    where PRICE and AMOUNT are decimal numbers.
    """
    items = []

    # Find section before "Subtotal" or "Total $"
    end_match = re.search(r'(Subtotal|Total\s+\$)', text, re.IGNORECASE)
    search_text = text[:end_match.start()] if end_match else text

    # Look for lines ending with two decimal numbers
    pattern = r'^(.{10,200}?)\s+(\d+\.?\d{2})\s+(\d+\.?\d{2})\s*$'

    for match in re.finditer(pattern, search_text, re.MULTILINE):
        content = match.group(1).strip()
        unit_price = match.group(2)
        amount = match.group(3)

        # Skip header-like lines
        if re.search(r'^(item|sku|qty|description|price|amount)', content, re.IGNORECASE):
            continue

        item = parse_row_content(content, unit_price, amount)
        if item:
            items.append(item)

    return items


# Keep the old function name as an alias for compatibility
def extract_line_items_sb(text):
    """Alias for extract_line_items for backwards compatibility."""
    return extract_line_items(text)


def parse_invoice(filepath, status_callback=None):
    """Parse a single invoice file and return structured data.

    Args:
        filepath: Path to the invoice file (PDF)
        status_callback: Optional function(msg, tag) for status updates

    Returns:
        dict with invoice data, or None if parsing failed
    """
    cb = status_callback or (lambda msg, tag=None: None)
    filename = os.path.basename(filepath)

    # Step 1: Try text extraction with pdfplumber
    cb(f"  Extracting text from {filename}...")
    text = extract_text_from_pdf(filepath)

    # Step 2: If text is too sparse, try OCR
    if len(text) < 50:
        if OCR_AVAILABLE:
            cb(f"  Text-based extraction sparse, trying OCR for {filename}...")
            text = extract_text_with_ocr(filepath)
            if len(text) < 50:
                cb(f"  Could not extract meaningful text from {filename}", "error")
                return None
        else:
            cb(f"  Could not extract text from {filename} (OCR not available)", "error")
            return None

    # Step 3: Parse the extracted text
    cb(f"  Parsing invoice data from {filename}...")
    data = parse_invoice_text(text)
    data['source_file'] = filename
    data['raw_text'] = text

    # Count extracted fields (excluding empty ones and internal fields)
    skip_keys = {'source_file', 'raw_text', 'line_items', 'vendor_address'}
    filled = sum(
        1 for k, v in data.items()
        if k not in skip_keys and v
    )

    line_item_count = len(data.get('line_items', []))

    if filled == 0:
        cb(f"  No fields could be extracted from {filename}", "warning")
    else:
        cb(f"  Extracted {filled} fields + {line_item_count} line item(s) from {filename}", "success")

    return data
