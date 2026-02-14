"""Invoice parser: extracts structured data from PDF invoices.

Designed to be vendor-agnostic - works with any invoice format by using
multiple extraction strategies (table-based + text-based) and broad
pattern matching for all fields.
"""
import os
import sys
import re
import csv
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


# Known vendors for fallback detection
KNOWN_VENDORS = ['S&B', 'S & B']

# Words/phrases that are NOT vendor names (used to filter letterhead lines)
NON_VENDOR_WORDS = {
    'invoice', 'sales order', 'credit memo', 'statement', 'receipt',
    'bill to', 'ship to', 'sold to', 'page', 'date', 'remittance',
    'terms and conditions', 'powered by', 'www.', 'http', 'warehouse',
}

# Known customer names (to exclude from vendor detection)
KNOWN_CUSTOMERS = [
    'diesel power products', 'power products unlimited',
    'dpp', 'bryan howell',
]

# Non-product line item keywords to filter out
NON_PRODUCT_KEYWORDS = [
    'l.c.', 'lc', 'd.n.a.', 'dna',
    'handling', 'surcharge',
]

FREIGHT_KEYWORDS = [
    'freight', 'shipping', 'drop ship', 'drop-ship', 'drop ship fee',
    'freight out', 'outbound freight',
]


# ---------------------------------------------------------------------------
# Vendor Normalization (from vendors.csv)
# ---------------------------------------------------------------------------

VENDORS_CSV_PATH = os.path.join(os.path.dirname(__file__), 'vendors.csv')
if not os.path.exists(VENDORS_CSV_PATH):
    try:
        candidates = []
        if getattr(sys, 'frozen', False):
            exe_dir = os.path.dirname(sys.executable)
            candidates.append(os.path.join(exe_dir, 'vendors.csv'))
            candidates.append(os.path.join(exe_dir, 'App', 'vendors.csv'))
            candidates.append(os.path.join(exe_dir, 'app', 'vendors.csv'))
        # Current working directory (e.g., running script from repo root)
        candidates.append(os.path.join(os.getcwd(), 'vendors.csv'))
        # Repo root when running from app/ as a script
        parent_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), os.pardir))
        candidates.append(os.path.join(parent_dir, 'vendors.csv'))

        for candidate in candidates:
            if os.path.exists(candidate):
                VENDORS_CSV_PATH = candidate
                break
    except Exception:
        pass


def _normalize_vendor_key(name):
    """Normalize vendor name to a comparable key."""
    if not name:
        return ''
    s = name.lower().strip()
    s = s.replace('&', 'and')
    s = re.sub(r'[^a-z0-9]+', '', s)
    return s


def _split_vendor_aliases(value):
    if not value:
        return []
    parts = re.split(r'[|;]', str(value))
    return [p.strip() for p in parts if p.strip()]


def _load_vendor_data():
    vendors = []
    key_to_canonical = {}
    alias_names = []
    if not os.path.exists(VENDORS_CSV_PATH):
        return vendors, key_to_canonical, alias_names
    try:
        with open(VENDORS_CSV_PATH, newline='', encoding='utf-8') as f:
            rows = list(csv.reader(f))
    except Exception:
        return [], {}, []
    if not rows:
        return vendors, key_to_canonical, alias_names

    header = [str(c).strip().lower() for c in rows[0]]
    has_header = any(
        h in ('vendor', 'invoice_vendor', 'skunexus_vendor', 'aliases', 'alias', 'additional_names')
        for h in header
    )
    seen_aliases = set()

    def add_alias(name):
        if not name:
            return
        key = name.strip().lower()
        if not key or key in seen_aliases:
            return
        seen_aliases.add(key)
        alias_names.append(name)

    if not has_header:
        for row in rows:
            if not row:
                continue
            val = str(row[0]).strip()
            if not val or val.lower() == 'vendor':
                continue
            vendors.append(val)
            key = _normalize_vendor_key(val)
            if key and key not in key_to_canonical:
                key_to_canonical[key] = val
        return vendors, key_to_canonical, alias_names

    def col(row, *names):
        for name in names:
            if name in header:
                idx = header.index(name)
                if idx < len(row):
                    return str(row[idx]).strip()
        return ''

    for row in rows[1:]:
        if not row:
            continue
        vendor = col(row, 'vendor')
        if not vendor:
            continue
        vendors.append(vendor)
        key = _normalize_vendor_key(vendor)
        if key and key not in key_to_canonical:
            key_to_canonical[key] = vendor
        aliases_val = col(row, 'aliases', 'alias', 'additional_names', 'invoice_vendor')
        if not aliases_val and 'skunexus_vendor' in header:
            aliases_val = col(row, 'skunexus_vendor')
        for alias in _split_vendor_aliases(aliases_val):
            if _normalize_vendor_key(alias) == _normalize_vendor_key(vendor):
                continue
            add_alias(alias)
            alias_key = _normalize_vendor_key(alias)
            if alias_key and alias_key not in key_to_canonical:
                key_to_canonical[alias_key] = vendor

    return vendors, key_to_canonical, alias_names


VENDOR_LIST, VENDOR_KEY_TO_CANONICAL, VENDOR_ALIAS_LIST = _load_vendor_data()


def _find_vendor_by_address_alias(text):
    if not text or not VENDOR_ALIAS_LIST:
        return ""
    normalized_text = _normalize_vendor_key(text)
    if not normalized_text:
        return ""
    for alias in VENDOR_ALIAS_LIST:
        if not re.search(r'\d', alias or ''):
            continue
        alias_key = _normalize_vendor_key(alias)
        if alias_key and alias_key in normalized_text:
            return alias
    return ""


def normalize_vendor_name(name):
    """Normalize vendor name to canonical form from vendors.csv."""
    if not name:
        return name
    key = _normalize_vendor_key(name)
    return VENDOR_KEY_TO_CANONICAL.get(key, name)


def _find_vendor_in_text_list(text, vendor_list):
    if not text or not vendor_list:
        return ""

    text_lower = text.lower()
    matches = []

    # Avoid matching known customers as vendors
    customer_phrases = [c for c in KNOWN_CUSTOMERS if len(c) >= 4]

    for vendor in vendor_list:
        v_lower = vendor.lower()
        if v_lower in text_lower:
            if any(cust in v_lower for cust in customer_phrases):
                continue
            pos = text_lower.find(v_lower)
            matches.append((len(vendor), pos, vendor))

    if not matches:
        return ""

    # Prefer the longest name; tie-breaker: earliest occurrence
    matches.sort(key=lambda t: (-t[0], t[1]))
    return matches[0][2]


def _find_vendor_in_text(text):
    """Find a vendor name from vendors.csv that appears in the text."""
    # Column 1 (canonical) first, then aliases (column 2)
    match = _find_vendor_in_text_list(text, VENDOR_LIST)
    if match:
        return match
    return _find_vendor_in_text_list(text, VENDOR_ALIAS_LIST)


def validate_vendor_name(text):
    """Check if extracted text looks like a valid vendor name."""
    if not text or len(text) < 2 or len(text) > 80:
        return False
    if not re.search(r'[A-Za-z]', text):
        return False
    if re.search(r'_{3,}|Credit Card|Type:|Authorize|Please Enter', text, re.IGNORECASE):
        return False
    # Reject if it's a known customer name
    if text.lower().strip() in KNOWN_CUSTOMERS:
        return False
    # Reject common non-vendor words
    if text.lower().strip() in NON_VENDOR_WORDS:
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
    """Extract text from a scanned PDF or image using OCR (Tesseract)."""
    if not OCR_AVAILABLE:
        return ""

    ext = os.path.splitext(filepath)[1].lower()

    if ext in ('.png', '.jpg', '.jpeg', '.tiff', '.tif', '.bmp'):
        try:
            img = Image.open(filepath)
            return pytesseract.image_to_string(img).strip()
        except Exception:
            return ""

    if not PDFIUM_AVAILABLE:
        return ""

    try:
        pdf = pdfium.PdfDocument(filepath)
        text = ""
        for page_index in range(len(pdf)):
            page = pdf[page_index]
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


def _normalize_amount_string(raw):
    if raw is None:
        return ''
    s = str(raw).strip()
    s = s.replace('$', '').replace(',', '')
    return s


def _parse_amount_value(raw):
    s = _normalize_amount_string(raw)
    if not s:
        return None
    negative = False
    if s.startswith('(') and s.endswith(')'):
        negative = True
        s = s[1:-1].strip()
    try:
        value = float(s)
    except ValueError:
        return None
    return -value if negative else value


_AMOUNT_RE = re.compile(
    r'\$?\s*([0-9]{1,3}(?:,[0-9]{3})*(?:\.\d{1,2})?|[0-9]+(?:\.\d{1,2})?)'
)


def _extract_amounts_from_line(line):
    results = []
    for match in _AMOUNT_RE.finditer(line):
        raw = match.group(1)
        value = _parse_amount_value(raw)
        if value is None:
            continue
        results.append((value, raw))
    return results


def _extract_total_amount(text):
    if not text:
        return ''

    strong_patterns = [
        r'\btotal\s+due\b',
        r'\binvoice\s+total\b',
        r'\bgrand\s+total\b',
        r'\btotal\s+amount\b',
        r'\btotal\s+usd\b',
        r'\bamount\s+due\b',
        r'\bbalance\s+due\b',
    ]
    weak_patterns = [
        r'\bsub\s*-?\s*total\b',
        r'\bsubtotal\b',
    ]
    generic_total = r'\btotal\b'
    exclude_words = [
        'tax', 'sales tax', 'shipping', 'freight', 'handling',
        'discount', 'surcharge', 'deposit',
    ]

    strong_candidates = []
    weak_candidates = []

    lines = text.splitlines()
    for idx, line in enumerate(lines):
        line_stripped = line.strip()
        if not line_stripped:
            continue
        lower = line_stripped.lower()

        strength = None
        if any(re.search(pat, lower) for pat in strong_patterns):
            strength = 'strong'
        elif any(re.search(pat, lower) for pat in weak_patterns):
            strength = 'weak'
        elif re.search(generic_total, lower):
            if not any(word in lower for word in exclude_words) and not any(
                re.search(pat, lower) for pat in weak_patterns
            ):
                strength = 'strong'

        if not strength:
            continue

        amounts = _extract_amounts_from_line(line_stripped)
        lookahead = 0
        if not amounts:
            for j in range(idx + 1, min(len(lines), idx + 3)):
                next_line = lines[j].strip()
                if not next_line:
                    continue
                amounts = _extract_amounts_from_line(next_line)
                if amounts:
                    lookahead = j - idx
                    break

        if not amounts:
            continue

        for value, raw in amounts:
            candidate = {
                'value': value,
                'raw': raw,
                'line_index': idx,
                'lookahead': lookahead,
            }
            if strength == 'strong':
                strong_candidates.append(candidate)
            else:
                weak_candidates.append(candidate)

    def pick_best(candidates):
        if not candidates:
            return ''
        candidates.sort(
            key=lambda c: (c['value'], c['line_index'], -c['lookahead']),
            reverse=True
        )
        return _normalize_amount_string(candidates[0]['raw'])

    best = pick_best(strong_candidates)
    if best:
        return best
    return pick_best(weak_candidates)


# ---------------------------------------------------------------------------
# Vendor Detection
# ---------------------------------------------------------------------------

def detect_vendor(text):
    """Detect vendor name from invoice text.

    The vendor is the company that ISSUED the invoice (who you pay money TO).
    This is NOT the customer (who the invoice is billed to).

    Strategy priority:
    1. "Remit To" / "Payment Address" (highest confidence)
    2. Company name from letterhead (first few lines before "Bill To")
    3. Known vendors list
    """

    # Strategy 0: Address alias match (highest confidence)
    address_vendor = _find_vendor_by_address_alias(text)
    if address_vendor:
        return address_vendor

    # Strategy 1: "Remit To" / "Pay To" (RH and similar)
    remit_match = re.search(
        r'(?:Remit|Pay)\s+To\s*:\s*([A-Za-z][A-Za-z0-9 &\-\.,]+?)(?:\n|\d)',
        text, re.IGNORECASE
    )
    if remit_match:
        vendor = remit_match.group(1).strip().rstrip(',.')
        if validate_vendor_name(vendor):
            return vendor

    # Strategy 3: "Payment Address:" section (FL - Fleece Performance)
    payment_addr_match = re.search(
        r'Payment\s+Address\s*:\s*\n?\s*([A-Za-z][A-Za-z0-9 &\-\.,]+?)(?:\n|$)',
        text, re.IGNORECASE
    )
    if payment_addr_match:
        vendor = payment_addr_match.group(1).strip().rstrip(',.')
        if validate_vendor_name(vendor):
            return vendor

    # Strategy 4: Any vendor name found in vendors.csv (scan full text)
    vendor_from_list = _find_vendor_in_text(text)
    if vendor_from_list:
        return vendor_from_list

    # Strategy 5: Company name from letterhead (first lines before "Bill To")
    vendor = _extract_vendor_from_letterhead(text)
    if vendor:
        return vendor

    # Strategy 6: Known vendors list
    for vendor in KNOWN_VENDORS:
        pattern = r'\b' + re.escape(vendor) + r'\b'
        if re.search(pattern, text, re.IGNORECASE):
            return vendor

    # Strategy 7: Extract company name from URLs in the text (e.g., www.Turn14.com)
    DOMAIN_VENDOR_MAP = {
        'turn14': 'Turn 14 Distribution',
        'fleeceperformance': 'Fleece Performance',
        'industrialinjection': 'Industrial Injection Service, Inc.',
    }

    url_match = re.search(r'(?<!@)\b(?:www\.)?([A-Za-z0-9\-]+)\.(com|net|org)\b', text, re.IGNORECASE)
    if url_match:
        domain = url_match.group(1).lower()
        if domain in DOMAIN_VENDOR_MAP:
            return DOMAIN_VENDOR_MAP[domain]
        # Try to find a more complete company name referencing this domain in the text
        full_name_match = re.search(
            r'(' + re.escape(domain[:4]) + r'[A-Za-z0-9 \-]+(?:Distribution|Inc\.?|LLC|Corp\.?|Ltd\.?)?)',
            text, re.IGNORECASE
        )
        if full_name_match:
            vendor = full_name_match.group(1).strip()
            if validate_vendor_name(vendor) and len(vendor) > 3:
                return vendor

    # Strategy 8: "Thank you for choosing X" pattern (CNC)
    thanks_match = re.search(
        r'[Tt]hank\s+you\s+for\s+choosing\s+([A-Za-z][A-Za-z0-9 &\-]+?)(?:\.|$)',
        text
    )
    if thanks_match:
        vendor = thanks_match.group(1).strip()
        if validate_vendor_name(vendor):
            return vendor

    return ""


def infer_vendor_from_filename(filename):
    """Fallback: infer vendor from the filename when the text lacks a vendor name."""
    if not filename:
        return ""

    base = os.path.splitext(os.path.basename(filename))[0]
    match = re.search(r'from_([A-Za-z0-9_\-]+)', base, re.IGNORECASE)
    if match:
        raw = match.group(1)
        raw = re.sub(r'_[0-9]+$', '', raw)
        vendor = raw.replace('_', ' ').replace('-', ' ').strip()
        vendor = re.sub(r'\s+', ' ', vendor)
        vendor = vendor.title()
        if validate_vendor_name(vendor):
            return vendor

    return ""


def _extract_vendor_from_letterhead(text):
    """Extract vendor name from the letterhead area (top of invoice).

    The vendor name is typically in the first few lines of the document,
    before "Bill To" / "Ship To" sections.
    """
    # Get text before "Bill To" or "Ship To" section
    bill_to_pos = re.search(r'(?:Bill|Sold|Ship)\s+To', text, re.IGNORECASE)
    header_text = text[:bill_to_pos.start()] if bill_to_pos else text[:500]

    lines = header_text.split('\n')

    for line in lines[:10]:
        line = line.strip()
        if not line or len(line) < 3:
            continue

        line_lower = line.lower()

        has_company_suffix = bool(re.search(
            r'(?:Inc\.?|LLC|Corp\.?|Ltd\.?|Co\.?|Enterprises|Service|Engineering|Distribution|Distributing|Performance|Motorsports|Fabrication)',
            line, re.IGNORECASE
        ))

        # Skip generic labels and headers (unless the line looks like a company name)
        if any(word in line_lower for word in NON_VENDOR_WORDS) and not has_company_suffix:
            continue
        if line_lower in ('invoice', 'sales order', 'credit memo', 'usa'):
            continue

        # Skip lines starting with field labels
        if re.match(r'^(?:Invoice|Date|PO|P\.O\.|Terms|Page|Customer|Phone|Fax|Tel|Tax|Ship|Due)\b', line, re.IGNORECASE):
            continue

        # Skip addresses: lines starting with numbers + street suffix
        if re.match(r'^\d+\s+', line) and re.search(r'(Ave|Avenue|St|Street|Rd|Road|Blvd|Dr|Drive|Way|Ln|Lane|Ct|Court|Ste|Suite|Commerce|Main|Spencer|Bonsai|Civic|Tournament|Slover)', line, re.IGNORECASE):
            continue

        # Skip city/state/zip lines (e.g., "AUBURN, WA 98001", "North Las Vegas, NV 89030")
        if re.match(r'^[A-Za-z][A-Za-z\s]+,\s*[A-Z]{2}\s+\d{5}', line):
            continue
        # Also: "Pittsboro, IN 46167" mixed into other content
        if re.search(r'[A-Z]{2}\s+\d{5}', line) and ',' in line and len(line) < 40:
            continue
        # Skip city/state/zip without comma (e.g., "Fontana CA 92337")
        if re.match(r'^[A-Za-z][A-Za-z\s]+\s+[A-Z]{2}\s+\d{5}(?:-\d{4})?$', line):
            continue

        # Skip phone numbers and fax
        if re.match(r'^[\d\(\)\-\s]{7,}$', line):
            continue
        if re.search(r'\(\d{3}\)\s*\d{3}[\-\.]\d{4}', line):
            continue
        if re.match(r'^\d{3}[\-\.]\d{3,4}[\-\.]\d{4}', line):
            continue

        # Skip email addresses and URLs
        if '@' in line or 'www.' in line.lower() or 'http' in line.lower() or '.com' in line.lower():
            continue

        # Skip lines that are just a date or contain date patterns
        if re.match(r'^(?:Date\s+)?\d{1,2}/\d{1,2}/\d{2,4}$', line):
            continue

        # Skip lines that are just invoice numbers
        if re.match(r'^[A-Z]?-?\d+$', line):
            continue

        # Skip customer names
        if any(cust in line_lower for cust in KNOWN_CUSTOMERS):
            continue

        # Skip "Customer: XXXXX" lines
        if re.match(r'^Customer\s*:', line, re.IGNORECASE):
            continue

        # Skip "PO #: XXXX" style lines
        if re.match(r'^PO\s*#', line, re.IGNORECASE):
            continue

        if has_company_suffix:
            # Clean trailing labels from the line: "Company, Inc. Date: 2026-01-29" → "Company, Inc."
            clean_line = re.sub(r'\s+(?:Date|Invoice|Page)\s*:?\s*\S*.*$', '', line, flags=re.IGNORECASE).strip()
            # Also: "Company, Inc. Invoice" → "Company, Inc."
            clean_line = re.sub(r'\s+Invoice\s*$', '', clean_line, flags=re.IGNORECASE).strip()
            if validate_vendor_name(clean_line):
                return clean_line

        # Accept prominent capitalized lines that look like company names
        words = line.split()
        if len(words) >= 2:
            alpha_words = [w for w in words if w[0].isalpha()]
            if not alpha_words:
                continue
            is_capitalized = all(w[0].isupper() for w in alpha_words)
            is_upper = line == line.upper() and re.search(r'[A-Z]', line)
            if (is_capitalized or is_upper) and validate_vendor_name(line):
                # Extra check: not a section header
                if not re.match(r'(?:Bill|Ship|Sold|Invoice|Date|Terms|Page|PO|Customer)\s', line, re.IGNORECASE):
                    return line

    return ""


# ---------------------------------------------------------------------------
# Customer Name Detection
# ---------------------------------------------------------------------------

def _extract_name_after_customer(line):
    """Extract the trailing name after a known customer phrase."""
    if not line:
        return ""
    line_lower = line.lower()
    customer_phrases = [
        'diesel power products',
        'power products unlimited',
        'dpp',
    ]

    # Prefer longest match first
    customer_phrases.sort(key=len, reverse=True)

    for cust in customer_phrases:
        if cust in line_lower:
            idx = line_lower.find(cust)
            trailing = line[idx + len(cust):].strip()
            trailing = re.sub(r'^[^A-Za-z]+', '', trailing)
            trailing = re.sub(r'\s+', ' ', trailing).strip()
            # Remove repeated customer/DBA prefixes
            while True:
                cleaned = re.sub(
                    r'^(diesel\s+power\s+products|diesel\s+power|power\s+products|dpp)\b\s*',
                    '', trailing, flags=re.IGNORECASE
                ).strip()
                cleaned = re.sub(
                    r'^(inc\.?|llc|l\.l\.c\.|corp\.?|co\.?|company)\b[,\s]*',
                    '', cleaned, flags=re.IGNORECASE
                ).strip()
                cleaned = re.sub(r'^[^A-Za-z]+', '', cleaned).strip()
                # If the line includes "Dealer To:" / "Ship To:" / "To:", keep only the name after the last To:
                to_match = re.search(r'(?:Dealer|Ship|Bill)?\s*To\s*:\s*(.+)$', cleaned, re.IGNORECASE)
                if to_match:
                    cleaned = to_match.group(1).strip()
                if cleaned == trailing:
                    break
                trailing = cleaned
            if trailing:
                return trailing
    return ""


def _extract_contact_before_customer(line):
    """Extract leading contact name before known customer phrase."""
    if not line:
        return ""
    line_lower = line.lower()
    customer_phrases = [
        'diesel power products',
        'power products unlimited',
        'dpp',
    ]
    earliest_idx = None
    for cust in customer_phrases:
        idx = line_lower.find(cust)
        if idx > 0 and (earliest_idx is None or idx < earliest_idx):
            earliest_idx = idx
    if earliest_idx is not None:
        candidate = line[:earliest_idx].strip()
        candidate = re.sub(r'^(?:ship\s*to|bill\s*to)\s*:?','', candidate, flags=re.IGNORECASE).strip()
        if candidate and len(candidate.split()) >= 2:
            return candidate
    return ""


def _sanitize_customer_candidate(candidate):
    """Remove Bill To / customer prefixes and return the actual customer name."""
    if not candidate:
        return ""
    cleaned = re.sub(r'^(?:Dealer|Ship|Bill)?\s*To\s*:\s*', '', candidate, flags=re.IGNORECASE).strip()
    # Ignore obvious label-only values
    if re.match(r'^(invoice|invoice\s*#|invoice\s+number|customer|bill\s*to|ship\s*to|sold\s*to)$', cleaned, re.IGNORECASE):
        return ""
    # If multiple "To:" segments exist, take the last one (e.g., "To: Power Products Unlimited Dealer To: John Strong")
    to_match = re.search(r'(?:Dealer|Ship|Bill)?\s*To\s*:\s*(.+)$', cleaned, re.IGNORECASE)
    if to_match:
        cleaned = to_match.group(1).strip()
    # Trim address-like tails (start with digits or separated by wide spacing)
    cleaned = re.split(r'\s{2,}|\t', cleaned)[0].strip()
    cleaned = re.sub(r'\s+\d.*$', '', cleaned).strip()
    trailing = _extract_name_after_customer(cleaned)
    return trailing if trailing else cleaned


def extract_ship_to_name(text):
    """Extract ship-to contact/name (used for Customer/Project)."""
    if not text:
        return ""

    lines = text.split('\n')

    # 1) Inline "Ship To: Name" on the same line
    for line in lines:
        match = re.search(
            r'Ship\s*To\s*:?\s*([A-Za-z][A-Za-z0-9 &\-\./,]+?)(?:\s{2,}|$)',
            line, re.IGNORECASE
        )
        if match:
            name = _sanitize_customer_candidate(match.group(1).strip())
            name = re.split(r'\s+(?:Bill|Sold|Customer)\s+To', name, flags=re.IGNORECASE)[0].strip()
            if re.match(r'^(bill|ship)\s+to\b', name, re.IGNORECASE):
                continue
            if name and name.lower() not in KNOWN_CUSTOMERS:
                return name

    # 2) "Bill To Ship To" header, then next non-empty line usually has both names
    for idx, line in enumerate(lines):
        if re.search(r'Bill\s+To\s+Ship\s+To', line, re.IGNORECASE):
            # Prefer any subsequent line that includes a known customer phrase
            for j in range(idx + 1, min(idx + 6, len(lines))):
                candidate = lines[j].strip()
                if not candidate:
                    continue
                name = _extract_name_after_customer(candidate)
                if name:
                    return name
            # Fallback: first non-empty line after the header
            for j in range(idx + 1, min(idx + 6, len(lines))):
                candidate = lines[j].strip()
                if not candidate:
                    continue
                name = _sanitize_customer_candidate(candidate)
                if name:
                    return name
                break

    # 2b) "Ship To Bill To" header (S&B format), then next non-empty line
    for idx, line in enumerate(lines):
        if re.search(r'Ship\s+To\s+Bill\s+To', line, re.IGNORECASE):
            for j in range(idx + 1, min(idx + 6, len(lines))):
                candidate = lines[j].strip()
                if not candidate:
                    continue
                if re.match(r'^(shipping|ship\s+date|ship\s+via|tracking|tax|gst|po\s*#|invoice|date|terms|currency|notes)\b', candidate, re.IGNORECASE):
                    continue
                contact = _extract_contact_before_customer(candidate)
                if contact:
                    return contact
                name = _sanitize_customer_candidate(candidate)
                if name:
                    return name
                break

    # 3) Any line that starts with a known customer phrase and has trailing name
    for line in lines:
        name = _extract_name_after_customer(line)
        if name:
            return name

    return ""


def extract_customer_name(text):
    """Extract customer name (who the invoice is billed TO)."""

    ship_to_name = extract_ship_to_name(text)
    if ship_to_name and ship_to_name.lower() not in KNOWN_CUSTOMERS:
        return ship_to_name

    # Strategy 1: "Bill To" section - find the next line(s) after "Bill To"
    # Many invoices have "Bill To" and "Ship To" on the same line as headers,
    # with the actual names on subsequent lines. Handle both formats:
    #   Format A: "Bill To\nCompanyName\n..."
    #   Format B: "Bill To Ship To\nCompany1 Company2\n..."
    #   Format C: "Bill To:\nCompanyName"
    #   Format D: "BILL TO\nDIESEL POWER PRODUCTS"
    bill_to_match = re.search(
        r'(?:BILL|Bill)\s*(?:TO|To)\s*:?\s*(?:Ship\s*To\s*:?\s*)?\n\s*([A-Za-z][A-Za-z0-9 &\-\./,]+)',
        text, re.IGNORECASE
    )
    if bill_to_match:
        customer = _sanitize_customer_candidate(bill_to_match.group(1).strip())
        # Don't return "Ship To" as customer
        if len(customer) >= 3 and customer.lower() not in ('ship to', 'ship to:'):
            return customer

    # Strategy 2: "Bill To:" with name on same line (PPE: "To: Power Products Unlimited")
    bill_to_inline = re.search(
        r'(?:Bill\s+)?To\s*:\s*([A-Za-z][A-Za-z0-9 &\-\./]+?)(?:\s+Dealer|\s+Ship|\n)',
        text, re.IGNORECASE
    )
    if bill_to_inline:
        customer = _sanitize_customer_candidate(bill_to_inline.group(1).strip())
        if len(customer) >= 3 and customer.lower() not in ('ship to',):
            return customer

    # Strategy 3: "Customer: Name" pattern (S&B remittance, FL)
    customer_match = re.search(
        r'Customer\s*:\s*([A-Za-z][A-Za-z0-9 &\-\.]+)',
        text, re.IGNORECASE
    )
    if customer_match:
        customer = _sanitize_customer_candidate(customer_match.group(1).strip())
        # Skip if it's just a customer ID like "DLPP03"
        if len(customer) >= 3 and not re.match(r'^[A-Z]{2,5}\d+$', customer):
            return customer

    # Strategy 4: "Customer ID Name" pattern
    customer_match = re.search(
        r'Customer\s+\d+\s+([A-Za-z][A-Za-z0-9 &\-\.]+)',
        text, re.IGNORECASE
    )
    if customer_match:
        customer = _sanitize_customer_candidate(customer_match.group(1).strip())
        if len(customer) >= 3:
            return customer

    # Strategy 5: "Billed To" / "Invoice To"
    billed_match = re.search(
        r'(?:Billed\s+To|Invoice\s+To)[:\s]+([A-Za-z][A-Za-z0-9 &\-\.]+)',
        text, re.IGNORECASE
    )
    if billed_match:
        customer = _sanitize_customer_candidate(billed_match.group(1).strip())
        if len(customer) >= 3:
            return customer

    return ""


# ---------------------------------------------------------------------------
# Vendor Address Detection
# ---------------------------------------------------------------------------

def extract_vendor_address(text):
    """Extract vendor address from the letterhead area.

    Looks for address lines between the vendor name and 'Bill To' section.
    """
    # S&B specific address (keep for backwards compat)
    sb_match = re.search(
        r'(15461\s+Slover\s+Avenue\s*\n\s*Fontana\s+CA\s+\d+)',
        text, re.IGNORECASE
    )
    if sb_match:
        return sb_match.group(1).strip().replace('\n', ', ')

    # Generic: Find address-like lines in the header area (before Bill To)
    bill_to_pos = re.search(r'(?:Bill|Sold)\s+To', text, re.IGNORECASE)
    header_text = text[:bill_to_pos.start()] if bill_to_pos else text[:500]

    # Look for street address pattern: NUMBER STREET_NAME
    addr_match = re.search(
        r'(\d+\s+[A-Za-z][A-Za-z0-9 \.]+(?:Ave|St|Rd|Blvd|Dr|Way|Ln|Ct|Ste|Suite|Commerce)[A-Za-z0-9 \.,]*\n\s*[A-Za-z]+[\w ,]+\d{5})',
        header_text, re.IGNORECASE
    )
    if addr_match:
        return addr_match.group(1).strip().replace('\n', ', ')

    return ''


# ---------------------------------------------------------------------------
# Line Item Extraction (Table-first, Text-fallback)
# ---------------------------------------------------------------------------

def is_non_product_row(item):
    """Check if a line item is a non-product row (shipping, core, discount, etc.)."""
    item_num = str(item.get('item_number', '')).lower().strip()
    desc = str(item.get('description', '')).lower().strip()
    combined = f"{item_num} {desc}"

    for keyword in NON_PRODUCT_KEYWORDS:
        if len(keyword) <= 2:
            if re.search(r'\b' + re.escape(keyword) + r'\b', combined):
                return True
        else:
            if keyword in combined:
                return True

    return False


def mark_freight_item(item):
    """Mark item as freight/shipping if keywords present."""
    if not item:
        return item
    item_num = str(item.get('item_number', '')).lower().strip()
    desc = str(item.get('description', '')).lower().strip()
    combined = f"{item_num} {desc}"
    if any(k in combined for k in FREIGHT_KEYWORDS):
        item['is_freight'] = True
    return item


def identify_line_item_table(table):
    """Find the header row in a table and map columns to standard fields.

    Returns:
        tuple: (header_row_index, column_map_dict) or (None, None) if not found

    column_map is like: {'item_number': 3, 'quantity': 5, 'unit_price': 7, 'amount': 8, ...}
    where values are column indices.
    """
    # Keywords that identify each column type
    ITEM_KEYWORDS = ['item', 'part', 'sku', 'product', 'item code', 'part number']
    QTY_KEYWORDS = ['qty', 'quantity', 'order qty', 'ship qty', 'invoiced qt', 'invoiced qty']
    PRICE_KEYWORDS = ['unit price', 'price each', 'rate']
    AMOUNT_KEYWORDS = ['amount', 'total', 'total price', 'ext.', 'ext', 'amount(net)']
    DESC_KEYWORDS = ['description', 'desc', 'product and description']
    UNIT_KEYWORDS = ['u/m', 'um', 'qty um', 'price um', 'unit', 'units']

    for row_idx, row in enumerate(table):
        if not row:
            continue

        # Build a list of cleaned header cell values
        headers = []
        for cell in row:
            cell_text = str(cell).strip().lower() if cell else ''
            headers.append(cell_text)

        # Check if this row looks like a header (needs at least 2 recognized columns)
        col_map = {}
        recognized = 0
        potential_unit_cols = []

        for col_idx, header in enumerate(headers):
            if not header:
                continue

            # Check each column type
            if any(kw in header for kw in ITEM_KEYWORDS):
                if 'item_number' not in col_map:
                    col_map['item_number'] = col_idx
                    recognized += 1
            elif any(kw == header or kw in header for kw in QTY_KEYWORDS):
                if 'quantity' not in col_map:
                    col_map['quantity'] = col_idx
                    recognized += 1
            elif any(kw in header for kw in PRICE_KEYWORDS):
                if 'unit_price' not in col_map:
                    col_map['unit_price'] = col_idx
                    recognized += 1
            elif any(kw in header for kw in AMOUNT_KEYWORDS):
                if 'amount' not in col_map:
                    col_map['amount'] = col_idx
                    recognized += 1
            elif any(kw in header for kw in DESC_KEYWORDS):
                if 'description' not in col_map:
                    col_map['description'] = col_idx
                    recognized += 1
            elif any(kw in header for kw in UNIT_KEYWORDS):
                if header in ('unit', 'units'):
                    potential_unit_cols.append(col_idx)
                elif 'units' not in col_map:
                    col_map['units'] = col_idx
                    recognized += 1

        # Heuristic: if we saw a bare "Unit(s)" column, decide whether it's price or units.
        # If a unit_price column was already found (e.g., "Rate" or "Unit Price"), treat
        # "Unit(s)" as UOM. Otherwise, assume "Unit" is the unit price (II-style headers).
        if potential_unit_cols:
            if 'unit_price' not in col_map:
                col_map['unit_price'] = potential_unit_cols[0]
                recognized += 1
                if len(potential_unit_cols) > 1 and 'units' not in col_map:
                    col_map['units'] = potential_unit_cols[1]
                    recognized += 1
            elif 'units' not in col_map:
                col_map['units'] = potential_unit_cols[0]
                recognized += 1

        # Need at least 2 recognized columns to consider this a line-item table
        if recognized >= 2:
            return row_idx, col_map

    return None, None


def _clean_cell(value):
    """Clean a table cell value - handle None, newlines, whitespace."""
    if value is None:
        return ''
    # Take first line if cell contains newline-separated values
    text = str(value).strip()
    if '\n' in text:
        text = text.split('\n')[0].strip()
    return text


def _clean_price(value):
    """Clean a price string - remove $, commas, whitespace."""
    if not value:
        return ''
    text = str(value).strip()
    # Take first value if newline-separated
    if '\n' in text:
        text = text.split('\n')[0].strip()
    text = text.replace('$', '').replace(',', '').strip()
    # Validate it looks like a number
    try:
        float(text)
        return text
    except (ValueError, TypeError):
        return ''


def _normalize_qty(value):
    """Normalize quantity like 1.00 -> 1, while preserving real decimals."""
    if value is None:
        return ''
    s = str(value).strip()
    if s == '':
        return ''
    s_clean = s.replace(',', '')
    try:
        num = float(s_clean)
    except (ValueError, TypeError):
        return s
    if abs(num - round(num)) < 1e-9:
        return str(int(round(num)))
    return s


def _split_cell_lines(value):
    if value is None:
        return []
    text = str(value).strip()
    if not text:
        return []
    return [line.strip() for line in text.split('\n') if line.strip()]


def _is_sb_vendor_name(name):
    """Return True if vendor name looks like S&B."""
    key = _normalize_vendor_key(name or '')
    if key:
        if key in {'sb', 'sbandb'} or 'sandb' in key:
            return True
    canonical = normalize_vendor_name(name or '')
    canonical_key = _normalize_vendor_key(canonical or '')
    if canonical_key:
        if canonical_key in {'sb', 'sbandb'} or 'sandb' in canonical_key:
            return True
    return False


def _is_sb_delivery_fee(item):
    """Match S&B delivery fee variants (Colorado/CO + Delivery Fee)."""
    item_number = str(item.get('item_number', '')).strip()
    desc = str(item.get('description', '')).strip()
    text = f"{item_number} {desc}".lower()
    return bool(re.search(r'\b(?:colorado|co)\b.*\bdelivery\s+fee\b', text))


def _expand_multiline_row(row, col_map):
    """Expand a table row that contains multiple items separated by newlines."""
    def _lines_for(key):
        idx = col_map.get(key)
        if idx is None or idx >= len(row):
            return []
        return _split_cell_lines(row[idx])

    item_lines = _lines_for('item_number')
    qty_lines = _lines_for('quantity')
    unit_lines = _lines_for('units')
    price_lines = _lines_for('unit_price')
    amount_lines = _lines_for('amount')
    desc_lines = _lines_for('description')

    item_count = max(len(item_lines), len(qty_lines), len(price_lines), len(amount_lines), 1)

    def _pick(lines, i):
        if not lines:
            return ''
        if len(lines) == 1:
            return lines[0]
        if i < len(lines):
            return lines[i]
        return ''

    if len(desc_lines) > item_count:
        desc_per_item = []
        for i in range(item_count):
            if i < item_count - 1:
                desc_per_item.append(desc_lines[i])
            else:
                desc_per_item.append(' '.join(desc_lines[i:]))
    elif len(desc_lines) == 1:
        desc_per_item = [desc_lines[0]] * item_count
    else:
        desc_per_item = [_pick(desc_lines, i) for i in range(item_count)]

    items = []
    for i in range(item_count):
        item = {
            'item_number': _clean_cell(_pick(item_lines, i)),
            'quantity': _clean_cell(_pick(qty_lines, i)),
            'unit_price': _clean_price(_pick(price_lines, i)),
            'amount': _clean_price(_pick(amount_lines, i)),
            'description': _clean_cell(desc_per_item[i]),
            'units': _clean_cell(_pick(unit_lines, i)) or 'Each',
        }
        items.append(mark_freight_item(item))
    return items


def _expand_multiline_row_sb(row, col_map):
    """S&B-specific: merge multiline item_number into one description when qty is single-line."""
    def _lines_for(key):
        idx = col_map.get(key)
        if idx is None or idx >= len(row):
            return []
        return _split_cell_lines(row[idx])

    item_lines = _lines_for('item_number')
    qty_lines = _lines_for('quantity')
    unit_lines = _lines_for('units')
    price_lines = _lines_for('unit_price')
    amount_lines = _lines_for('amount')
    desc_lines = _lines_for('description')

    def _all_alpha(lines):
        return all(not re.search(r'\d', line) for line in lines)

    if (
        len(item_lines) > 1
        and _all_alpha(item_lines)
        and len(qty_lines) <= 1
        and len(price_lines) <= 1
        and len(amount_lines) <= 1
    ):
        desc_parts = []
        if item_lines:
            desc_parts.append(' '.join(item_lines))
        if desc_lines:
            desc_parts.append(' '.join(desc_lines))
        description = ' '.join([p for p in desc_parts if p]).strip()
        item = {
            'item_number': '',
            'quantity': _clean_cell(qty_lines[0]) if qty_lines else '',
            'unit_price': _clean_price(price_lines[0]) if price_lines else '',
            'amount': _clean_price(amount_lines[0]) if amount_lines else '',
            'description': _clean_cell(description),
            'units': _clean_cell(unit_lines[0]) if unit_lines else 'Each',
        }
        return [mark_freight_item(item)]

    # If item_number has extra lines beyond qty/price/amount, merge extra item lines into
    # the corresponding description rows (S&B specific behavior).
    item_count = max(len(qty_lines), len(price_lines), len(amount_lines), 1)
    if len(item_lines) > item_count:
        # If we see a tail like "Retail", "Delivery Fee", push those onto the LAST item description.
        extra_lines = item_lines[item_count:]
        item_lines = item_lines[:item_count]
        desc_lines = desc_lines or [''] * item_count
        if len(desc_lines) < item_count:
            desc_lines = desc_lines + ([''] * (item_count - len(desc_lines)))
        tail_text = ' '.join(extra_lines).strip()
        if tail_text:
            last_idx = item_count - 1
            last_item = item_lines[last_idx] if last_idx < len(item_lines) else ''
            if last_item and not re.search(r'\d', last_item):
                item_lines[last_idx] = (last_item + ' ' + tail_text).strip()
            elif not last_item:
                item_lines[last_idx] = tail_text
            else:
                desc_lines[last_idx] = (desc_lines[last_idx] + ' ' + tail_text).strip()
            if last_idx > 0:
                candidate = {'item_number': item_lines[last_idx], 'description': ''}
                if _is_sb_delivery_fee(candidate) and desc_lines[last_idx]:
                    desc_lines[last_idx - 1] = (desc_lines[last_idx - 1] + ' ' + desc_lines[last_idx]).strip()
                    desc_lines[last_idx] = ''

        items = []
        for i in range(item_count):
            item = {
                'item_number': _clean_cell(item_lines[i]) if i < len(item_lines) else '',
                'quantity': _clean_cell(qty_lines[i]) if i < len(qty_lines) else '',
                'unit_price': _clean_price(price_lines[i]) if i < len(price_lines) else '',
                'amount': _clean_price(amount_lines[i]) if i < len(amount_lines) else '',
                'description': _clean_cell(desc_lines[i]) if i < len(desc_lines) else '',
                'units': _clean_cell(unit_lines[i]) if i < len(unit_lines) else 'Each',
            }
            items.append(mark_freight_item(item))
        return items

    return _expand_multiline_row(row, col_map)


def _row_has_multiline_values(row, col_map):
    for key in ('item_number', 'quantity', 'unit_price', 'amount', 'description', 'units'):
        idx = col_map.get(key)
        if idx is None or idx >= len(row):
            continue
        val = row[idx]
        if val is not None and '\n' in str(val):
            return True
    return False


def _find_nearby_value(row, col_idx, max_offset=2, predicate=None, exclude_cols=None):
    """Find a nearby cell value around col_idx that matches predicate."""
    if not row or col_idx is None:
        return ''
    exclude_cols = set(exclude_cols or [])
    seen = set()
    for offset in range(0, max_offset + 1):
        for j in (col_idx - offset, col_idx + offset):
            if j in seen:
                continue
            seen.add(j)
            if j < 0 or j >= len(row):
                continue
            if j in exclude_cols:
                continue
            val = row[j]
            if val is None or str(val).strip() == '':
                continue
            if predicate is None or predicate(val):
                return val
    return ''


def extract_item_from_table_row(row, col_map):
    """Extract a line item dict from a table row using the column map."""
    item = {}

    item['item_number'] = _clean_cell(row[col_map['item_number']]) if 'item_number' in col_map and col_map['item_number'] < len(row) else ''
    item['quantity'] = _clean_cell(row[col_map['quantity']]) if 'quantity' in col_map and col_map['quantity'] < len(row) else ''
    item['unit_price'] = _clean_price(row[col_map['unit_price']]) if 'unit_price' in col_map and col_map['unit_price'] < len(row) else ''
    item['amount'] = _clean_price(row[col_map['amount']]) if 'amount' in col_map and col_map['amount'] < len(row) else ''
    item['description'] = _clean_cell(row[col_map['description']]) if 'description' in col_map and col_map['description'] < len(row) else ''
    item['units'] = _clean_cell(row[col_map['units']]) if 'units' in col_map and col_map['units'] < len(row) else 'Each'

    # PD tables can be shifted; look near the mapped column for values
    if not item['quantity'] and 'quantity' in col_map:
        exclude_cols = {col_map.get('amount'), col_map.get('unit_price')}
        val = _find_nearby_value(
            row,
            col_map['quantity'],
            predicate=lambda v: re.match(r'^\d+(\.\d+)?$', str(v).strip()),
            exclude_cols=exclude_cols,
        )
        if val:
            item['quantity'] = _clean_cell(val)
    if not item['unit_price'] and 'unit_price' in col_map:
        val = _find_nearby_value(row, col_map['unit_price'], predicate=lambda v: _clean_price(v))
        if val:
            item['unit_price'] = _clean_price(val)
    if not item['amount'] and 'amount' in col_map:
        val = _find_nearby_value(row, col_map['amount'], predicate=lambda v: _clean_price(v))
        if val:
            item['amount'] = _clean_price(val)
    if (not item['units'] or item['units'] == 'Each') and 'units' in col_map:
        val = _find_nearby_value(row, col_map['units'], predicate=lambda v: bool(re.match(r'^[A-Za-z]+$', str(v).strip())))
        if val:
            item['units'] = _clean_cell(val)

    # If description is in the same column as item_number (some formats combine them)
    if not item['description'] and item['item_number'] and '\n' in str(row[col_map.get('item_number', 0)] or ''):
        parts = str(row[col_map['item_number']]).strip().split('\n')
        item['item_number'] = parts[0].strip()
        item['description'] = ' '.join(parts[1:]).strip()

    # For "Product and Description" combined columns (PD format)
    if 'description' in col_map and 'item_number' not in col_map:
        desc_val = str(row[col_map['description']] or '').strip()
        if '\n' in desc_val:
            parts = desc_val.split('\n')
            item['item_number'] = parts[0].strip()
            item['description'] = ' '.join(parts[1:]).strip()

    return mark_freight_item(item)


def extract_items_from_tables(filepath, sb_mode=False):
    """Extract line items from pdfplumber tables.

    This is the PRIMARY extraction method. pdfplumber can detect table
    structures in most invoice PDFs.
    """
    if not filepath:
        return []

    tables = extract_tables_from_pdf(filepath)

    for table in tables:
        if not table or len(table) < 2:
            continue

        header_row_idx, col_map = identify_line_item_table(table)
        if col_map is None:
            continue

        items = []
        last_qty = ''
        last_units = ''
        last_price = ''
        last_amount = ''
        last_desc = ''
        last_item_number = ''
        for row in table[header_row_idx + 1:]:
            if not row or all(c is None or str(c).strip() == '' for c in row):
                continue
            # Skip summary/total rows that snuck in
            combined_text = ' '.join(str(v) for v in row if v).lower()
            if any(word in combined_text for word in ['subtotal', 'total', 'tax', 'balance']):
                continue
            row_items = []
            if _row_has_multiline_values(row, col_map):
                if sb_mode:
                    row_items = _expand_multiline_row_sb(row, col_map)
                else:
                    row_items = _expand_multiline_row(row, col_map)
            else:
                item = extract_item_from_table_row(row, col_map)
                # S&B: merge description-only continuation rows into previous item
                if sb_mode:
                    item_number = str(item.get('item_number', '')).strip()
                    qty = str(item.get('quantity', '')).strip()
                    unit_price = str(item.get('unit_price', '')).strip()
                    amount = str(item.get('amount', '')).strip()
                    desc = str(item.get('description', '')).strip()
                    units = str(item.get('units', '')).strip()

                    has_numbers = bool(re.search(r'\d', item_number + qty + unit_price + amount))
                    if desc and not has_numbers and last_qty and (not qty) and (not unit_price) and (not amount):
                        merged_desc = (last_desc + ' ' + desc).strip() if last_desc else desc
                        if items:
                            items[-1]['description'] = merged_desc
                            last_desc = merged_desc
                            continue
                row_items = [item]

            for item in row_items:
                # Skip if no item number and no amount (probably a summary row)
                if not item.get('item_number') and not item.get('amount'):
                    continue

                # If quantity is missing, carry forward the last seen quantity (RH-style tables)
                if not item.get('quantity'):
                    if last_qty:
                        item['quantity'] = last_qty
                    elif item.get('amount'):
                        item['quantity'] = '1'
                else:
                    last_qty = item.get('quantity')
                # Track last values for S&B continuation merging
                last_units = item.get('units') or last_units
                last_price = item.get('unit_price') or last_price
                last_amount = item.get('amount') or last_amount
                last_desc = item.get('description') or last_desc
                last_item_number = item.get('item_number') or last_item_number

                # If unit price is missing but amount exists, use amount as rate (RH-style tables)
                if not item.get('unit_price') and item.get('amount') and item.get('quantity') in ('', '1'):
                    item['unit_price'] = item.get('amount')

                # Skip non-product rows
                if is_non_product_row(item):
                    continue

                # Validate: need at least an amount or item number
                if item.get('amount') or item.get('item_number'):
                    items.append(item)

        if items:
            return items

    return []


def extract_items_from_text(text):
    """Extract line items from invoice text using regex patterns.

    This is the FALLBACK method when table extraction fails.
    Handles various text layouts across different invoice formats.
    """
    items = []

    # Industrial Injection (II) format
    if re.search(r'Quantity\s+Item\s+RGA\s+Serial\s*#\s+Unit\s+Total', text, re.IGNORECASE):
        ii_items = _extract_ii_items(text)
        if ii_items:
            return ii_items

    # CNC Fabrication invoices have split headers; use specialized parser
    if re.search(r'ITEM\s+DESCRIPTION\s+QUANTITY\s+PRICE\s+EXT', text, re.IGNORECASE) and \
       re.search(r'SO\s+No\.\s+Customer\s+PO', text, re.IGNORECASE):
        cnc_items = _extract_cnc_items(text)
        if cnc_items:
            return cnc_items

    # Find the items section by looking for header rows
    header_patterns = [
        r'(Item/Description\s+.*?(?:Total\s+Price|Amount))\s*\n',  # PPE format
        r'(LINE\s+NO\.\s+ITEM\s+.*?(?:PRICE|EXT\.))\s*\n',  # CNC format
        r'(Quantity\s+Item\s+.*?(?:Unit|Total))\s*\n',  # II format
        r'(Item\s+Code\s+.*?Amount)\s*\n',  # NL format
        r'((?:Part\s+Number|Item|SKU|Product|Qty)\s+.*?(?:Amount|Total|Price|Ext\.))\s*\n',
    ]

    items_section = ""
    is_ppe = False
    for pattern in header_patterns:
        header_match = re.search(pattern, text, re.IGNORECASE)
        if header_match:
            header_end = header_match.end()
            header_text = header_match.group(1).lower() if header_match.lastindex else header_match.group(0).lower()
            is_ppe = ('item/description' in header_text) and ('total price' in header_text)

            # Find end of items section
            end_match = re.search(
                r'(?:^|\n)\s*(?:Subtotal|Sub\s*-?\s*total|Total\s+\$|Shipping\s+Cost|Tax\s+\d|I\s+HEREBY|Amount\s+Subject)',
                text[header_end:],
                re.IGNORECASE | re.MULTILINE
            )

            if end_match:
                items_section = text[header_end:header_end + end_match.start()]
            else:
                items_section = text[header_end:header_end + 2000]
            break

    if items_section.strip():
        if is_ppe:
            items = _extract_ppe_items(items_section)
        if not items:
            items = _parse_text_table_rows(items_section)

    # Fallback: look for lines with price patterns
    if not items:
        items = _extract_items_by_price_patterns(text)

    return items


def _parse_text_table_rows(items_section):
    """Parse rows from a text-based table section."""
    items = []
    lines = items_section.strip().split('\n')
    accumulated = ""
    last_item = None
    stop_patterns = [
        r'^Tracking\b', r'^Tracking\s+No', r'^Subtotal\b', r'^Total\b',
        r'^Taxes?\b', r'^Paid\b', r'^Balance\b', r'^Amount\b',
        r'^Thank\s+you', r'^Page\b', r'^Ship\b', r'^Bill\b',
        r'^\*', r'SHIPPING\s+ACT',
    ]

    for line in lines:
        line = line.strip()
        if not line:
            continue

        # Drop Ship / Freight lines (keep as line items)
        if re.search(r'^(Drop\s+Ship|Freight(?:\s+Out)?)\b', line, re.IGNORECASE):
            prices = re.findall(r'[\d,]+\.\d{2}', line)
            qty_match = re.search(r'(?:Drop\s+Ship|Freight(?:\s+Out)?)\s+\$?[\d,]+\.\d{2}\s+(\d+)', line, re.IGNORECASE)
            qty = qty_match.group(1) if qty_match else '1'
            unit_price = prices[-1].replace(',', '') if prices else ''
            desc = line
            if re.search(r'^Drop\s+Ship\b', line, re.IGNORECASE):
                drop_match = re.search(r'Drop\s+Ship\s+\$?([\d,]+\.\d{2})', line, re.IGNORECASE)
                if drop_match:
                    unit_price = drop_match.group(1).replace(',', '')
                    desc = f"Drop Ship ${unit_price}"
                else:
                    desc = "Drop Ship"
            item = {
                'item_number': 'Drop Ship' if re.search(r'Drop\s+Ship', line, re.IGNORECASE) else 'Freight',
                'quantity': qty,
                'units': 'Each',
                'description': desc,
                'unit_price': unit_price,
                'amount': unit_price,
            }
            mark_freight_item(item)
            items.append(item)
            accumulated = ""
            continue

        # Skip header-like or summary lines
        if re.match(r'^(item|sku|qty|description|subtotal|total)\s*$', line, re.IGNORECASE):
            continue

        # Description continuation line (no price at end)
        if last_item and not re.search(r'\$?\d[\d,]*\.\d{2}\s*$', line):
            if any(re.search(pat, line, re.IGNORECASE) for pat in stop_patterns):
                continue
            last_item['description'] = (last_item.get('description', '') + ' ' + line).strip()
            continue

        accumulated = (accumulated + " " + line).strip() if accumulated else line

        # Pattern 1: ... UNIT_PRICE AMOUNT (two numbers at end)
        price_match = re.search(r'^(.+?)\s+\$?([\d,]+\.?\d{2})\s+\$?([\d,]+\.?\d{2})\s*$', accumulated)

        if price_match:
            content = price_match.group(1)
            unit_price = price_match.group(2).replace(',', '')
            amount = price_match.group(3).replace(',', '')

            item = _parse_row_content(content, unit_price, amount)
            if item:
                mark_freight_item(item)
                if not is_non_product_row(item):
                    items.append(item)
                last_item = item
            accumulated = ""
            continue

        # Pattern 2: ... single AMOUNT (one number at end, for formats without unit price in text)
        single_price = re.search(r'^(.+?)\s+\$?([\d,]+\.?\d{2})\s*$', accumulated)
        if single_price and not re.search(r'\d+\.\d{2}\s+\d', accumulated):
            # Only use single-price if it doesn't look like there should be two prices
            pass  # Don't consume yet, might get unit_price on accumulation

    return items


def _extract_ppe_items(items_section):
    """Parse PPE-style item tables (Item/Description ... Total Price)."""
    items = []
    lines = [ln.strip() for ln in items_section.split('\n') if ln.strip()]
    i = 0

    item_line_re = re.compile(
        r'^(\S+)\s+(Each|EA|EACH|Unit|Piece|PC|PCS)\s+(\d+\.?\d*)\s+(\d+\.?\d*)\s+([\d,]+\.?\d{2})\s+([\d,]+\.?\d{2})$'
    )
    drop_ship_re = re.compile(
        r'^(?:DROP\s+SHIP|Drop\s+Ship)\s+(\d+\.?\d*)\s+(\d+\.?\d*)\s+([\d,]+\.?\d{2})\s+([\d,]+\.?\d{2})$'
    )

    while i < len(lines):
        line = lines[i]

        # Skip summary lines
        if re.search(r'(Subtotal|Total|Amount Subject|Amount Exempt|Total Sales Tax)', line, re.IGNORECASE):
            i += 1
            continue

        match = item_line_re.match(line)
        if match:
            sku = match.group(1)
            qty = match.group(4)
            unit_price = match.group(5).replace(',', '')
            amount = match.group(6).replace(',', '')
            unit_raw = match.group(2)
            units = 'Each' if unit_raw.lower() in ('each', 'ea') else unit_raw

            desc_parts = []
            j = i + 1
            while j < len(lines):
                next_line = lines[j]
                if item_line_re.match(next_line):
                    break
                if re.match(r'^(?:DROP\s+SHIP|Drop\s+Ship)\b', next_line):
                    break
                if re.search(r'(Subtotal|Total|Amount Subject|Amount Exempt|Total Sales Tax)', next_line, re.IGNORECASE):
                    break
                desc_parts.append(next_line)
                j += 1

            description = ' '.join(desc_parts).strip()
            item = {
                'item_number': sku,
                'quantity': qty,
                'units': units,
                'description': description,
                'unit_price': unit_price,
                'amount': amount,
            }
            mark_freight_item(item)
            if not is_non_product_row(item):
                items.append(item)

            i = j
            continue

        drop_match = drop_ship_re.match(line)
        if drop_match:
            qty = drop_match.group(2)
            unit_price = drop_match.group(3).replace(',', '')
            amount = drop_match.group(4).replace(',', '')
            desc = ''
            if i + 1 < len(lines) and re.match(r'^Drop\s+Ship\s+Fee', lines[i + 1], re.IGNORECASE):
                desc = lines[i + 1].strip()
                i += 1
            item = {
                'item_number': 'Drop Ship',
                'quantity': qty,
                'units': 'Each',
                'description': desc or 'Drop Ship',
                'unit_price': unit_price,
                'amount': amount,
            }
            mark_freight_item(item)
            items.append(item)
            i += 1
            continue

        i += 1

    return items


def _dedupe_token(token):
    """Collapse doubled characters in a token (common in CNC PDFs)."""
    if not token or len(token) < 4 or len(token) % 2 != 0:
        return token
    pairs = 0
    total_pairs = len(token) // 2
    for i in range(0, len(token), 2):
        if token[i] == token[i + 1]:
            pairs += 1
    if total_pairs > 0 and (pairs / total_pairs) >= 0.6:
        return ''.join(token[i] for i in range(0, len(token), 2))
    return token


def _dedupe_line(line):
    tokens = line.split()
    if not tokens:
        return line
    tokens = [_dedupe_token(tok) for tok in tokens]
    return ' '.join(tokens)


def _extract_cnc_items(text):
    """Parse CNC Fabrication line items from text when header is split."""
    if not text:
        return []

    lines = text.split('\n')
    header_idx = None
    header_re = re.compile(r'ITEM\s+DESCRIPTION\s+QUANTITY\s+PRICE\s+EXT\.?', re.IGNORECASE)
    for i, line in enumerate(lines):
        if header_re.search(line):
            header_idx = i
            break

    if header_idx is None:
        return []

    items = []
    stop_re = re.compile(r'^(Tracking|Subtotal|Taxes?|Total|Paid|Balance|Thank\s+you|Page)\b', re.IGNORECASE)
    skip_cont_re = re.compile(r'^(SN:|CORE\s+TRACKING#|Tracking\s+No)', re.IGNORECASE)
    cnc_non_product = ['hazmat', 'environmental', 'discount', 'handling', 'surcharge']

    i = header_idx + 1
    while i < len(lines):
        raw = lines[i].strip()
        if not raw:
            i += 1
            continue
        if stop_re.search(raw):
            break
        if re.match(r'^(LINE|NO\.?)$', raw, re.IGNORECASE):
            i += 1
            continue

        line = _dedupe_line(raw)

        match = re.match(
            r'^(\d+)\s+(\S+)\s+(.+?)\s+(\d+\.?\d*)\s+(\d+\.?\d*)\s+(\d+\.?\d*)$',
            line
        )
        if match:
            item = {
                'item_number': match.group(2),
                'quantity': match.group(4),
                'units': 'Each',
                'description': match.group(3).strip(),
                'unit_price': match.group(5),
                'amount': match.group(6),
            }

            # Handle SKU continuation on next line (e.g., CORE-6.0E- + HPOP)
            j = i + 1
            while j < len(lines):
                next_raw = lines[j].strip()
                if not next_raw:
                    j += 1
                    continue
                if stop_re.search(next_raw) or skip_cont_re.search(next_raw) or re.match(r'^(LINE|NO\.?)$', next_raw, re.IGNORECASE):
                    break
                next_line = _dedupe_line(next_raw)
                if re.match(r'^[A-Za-z0-9\-]+$', next_line) and ' ' not in next_line:
                    if item['item_number'].endswith('-') or next_line.isalpha():
                        item['item_number'] = item['item_number'] + next_line
                        i = j  # consume the continuation line
                    break
                break

            combined = f"{item['item_number']} {item['description']}".lower()
            if not any(k in combined for k in cnc_non_product):
                items.append(item)
            i += 1
            continue

        # Append short continuation lines (e.g., color) to last item description
        if items:
            if stop_re.search(line) or skip_cont_re.search(line):
                i += 1
                continue
            items[-1]['description'] = (items[-1]['description'] + ' ' + line).strip()

        i += 1

    return items


def _extract_ii_items(text):
    """Parse Industrial Injection items from text (Quantity Item RGA Serial # Unit Total)."""
    if not text:
        return []

    lines = text.split('\n')
    header_idx = None
    header_re = re.compile(r'Quantity\s+Item\s+RGA\s+Serial\s*#\s+Unit\s+Total', re.IGNORECASE)
    for i, line in enumerate(lines):
        if header_re.search(line):
            header_idx = i
            break

    if header_idx is None:
        return []

    items = []
    stop_re = re.compile(r'^(Subtotal|Total|Taxes?|Balance|I\s+HEREBY|RECEIVED|Page)\b', re.IGNORECASE)

    for i in range(header_idx + 1, len(lines)):
        raw = lines[i].strip()
        if not raw:
            continue
        if stop_re.search(raw):
            break

        # New line item starts with qty + sku
        m = re.match(r'^(\d+)\s+(\S+)\s+(.*)$', raw)
        if m:
            qty = m.group(1)
            sku = m.group(2)
            rest = m.group(3)

            prices = re.findall(r'\$?[\d,]+\.\d{2}', rest)
            unit_price = ''
            amount = ''
            if len(prices) >= 2:
                unit_price = prices[-2].replace('$', '').replace(',', '')
                amount = prices[-1].replace('$', '').replace(',', '')
            elif len(prices) == 1:
                unit_price = prices[0].replace('$', '').replace(',', '')
                amount = '' if qty == '0' else unit_price

            item = {
                'item_number': sku,
                'quantity': qty,
                'units': 'Each',
                'description': '',
                'unit_price': unit_price,
                'amount': amount,
            }
            combined = f"{sku} {rest}".lower()
            mark_freight_item(item)
            items.append(item)
            continue

        # Description continuation line
        if items:
            items[-1]['description'] = (items[-1]['description'] + ' ' + raw).strip()

    return items


def _parse_row_content(content, unit_price, amount):
    """Parse row content to extract item_number, quantity, units, description."""
    content = re.sub(r'\s+', ' ', content).strip()

    if len(content) < 3:
        return None
    price_token_re = r'\$?\d{1,3}(?:,\d{3})*\.\d{2}\b'

    # Pattern: SKU QTY [UNITS] DESCRIPTION
    match = re.match(
        r'^(\S+)\s+(\d+)\s*(?:\d+\s+)?(?:Each|EA|Piece|pc|pcs|units?)?\s+(\d+)\s*(?:Each|EA|Piece|pc|pcs|units?)?\s*(.*)',
        content, re.IGNORECASE
    )
    if match:
        # Check if this is LINE_NO ITEM_NO ... format (PD/CNC)
        pass

    # Pattern A0: SKU QTY BACKORDERED [U/M] (no description on this line)
    match = re.match(
        r'^(\S+)\s+(\d+)\s+\d+\s*(?:Each|EA|Piece|pc|pcs|units?)?$',
        content, re.IGNORECASE
    )
    if match:
        return {
            'item_number': match.group(1),
            'quantity': match.group(2),
            'units': 'Each',
            'description': '',
            'unit_price': unit_price,
            'amount': amount,
        }

    # Pattern A: SKU QTY [UNITS] DESCRIPTION (most common - S&B, FL)
    match = re.match(
        r'^(\S+)\s+(\d+)\s+(?:\d+\s+)?(?:Each|EA|Piece|pc|pcs|units?)\s+(.+)$',
        content, re.IGNORECASE
    )
    if match:
        return {
            'item_number': match.group(1),
            'quantity': match.group(2),
            'units': 'Each',
            'description': match.group(3).strip(),
            'unit_price': unit_price,
            'amount': amount,
        }

    # Pattern D: LINE_NO SKU DESCRIPTION QTY (CNC format)
    if not re.search(price_token_re, content):
        match = re.match(
            r'^(\d+)\s+(\S+)\s+(.+?)\s+(\d+\.?\d*)$',
            content
        )
        if match:
            return {
                'item_number': match.group(2),
                'quantity': match.group(4),
                'units': 'Each',
                'description': match.group(3).strip(),
                'unit_price': unit_price,
                'amount': amount,
            }

    # Pattern B: QTY ITEM_NO DESCRIPTION (II format: "1 5326058SE RA054795")
    match = re.match(
        r'^(\d+)\s+(\S+)\s+(.+)$',
        content
    )
    if match:
        qty = match.group(1)
        sku = match.group(2)
        desc = match.group(3).strip()
        # Remove RGA/Serial fields from description (II-specific)
        desc = re.sub(r'^(?:RA|RGA|SN|SER|SERIAL)\d+\s*', '', desc, flags=re.IGNORECASE).strip()
        # If description embeds price/discount tokens (e.g., T14), strip them out
        if re.search(price_token_re, desc):
            if unit_price:
                desc = re.sub(r'\$?' + re.escape(unit_price) + r'\b', '', desc)
            if amount:
                desc = re.sub(r'\$?' + re.escape(amount) + r'\b', '', desc)
            desc = re.sub(r'\s+\d+(\.\d+)?%?\s*$', '', desc).strip()
            desc = re.sub(r'\s{2,}', ' ', desc).strip()
        if desc and len(desc) >= 3:
            return {
                'item_number': sku,
                'quantity': qty,
                'units': 'Each',
                'description': desc,
                'unit_price': unit_price,
                'amount': amount,
            }

    # Pattern C: SKU QTY DESCRIPTION (no units keyword)
    match = re.match(
        r'^(\S+)\s+(\d+)\s+(.{3,})$',
        content
    )
    if match:
        return {
            'item_number': match.group(1),
            'quantity': match.group(2),
            'units': 'Each',
            'description': match.group(3).strip(),
            'unit_price': unit_price,
            'amount': amount,
        }

    # Pattern E: just DESCRIPTION (no SKU or qty parsed)
    if len(content) >= 5:
        return {
            'item_number': '',
            'quantity': '1',
            'units': 'Each',
            'description': content,
            'unit_price': unit_price,
            'amount': amount,
        }

    return None


def _extract_items_by_price_patterns(text):
    """Fallback: Extract items by finding lines with price patterns."""
    items = []

    end_match = re.search(r'(Subtotal|Sub\s*-?\s*total|Total\s+\$)', text, re.IGNORECASE)
    search_text = text[:end_match.start()] if end_match else text

    # Lines ending with two decimal numbers (unit_price amount)
    pattern = r'^(.{5,200}?)\s+\$?([\d,]+\.?\d{2})\s+\$?([\d,]+\.?\d{2})\s*$'

    for match in re.finditer(pattern, search_text, re.MULTILINE):
        content = match.group(1).strip()
        unit_price = match.group(2).replace(',', '')
        amount = match.group(3).replace(',', '')

        if re.search(r'^(item|sku|qty|description|price|amount)', content, re.IGNORECASE):
            continue

        item = _parse_row_content(content, unit_price, amount)
        if item and not is_non_product_row(item):
            items.append(item)

    return items


def extract_line_items(text, filepath=None, vendor_name=None):
    """Extract line items using table-first, text-fallback approach.

    Args:
        text: Extracted text from PDF
        filepath: Path to PDF file (for table extraction)

    Returns:
        list of dicts with item_number, quantity, units, description, unit_price, amount
    """
    # Step 1: Try pdfplumber table extraction (most reliable)
    sb_mode = _is_sb_vendor_name(vendor_name)
    items = extract_items_from_tables(filepath, sb_mode=sb_mode)

    # Step 2: Fall back to text-based extraction
    if not items:
        items = extract_items_from_text(text)

    for item in items:
        item['quantity'] = _normalize_qty(item.get('quantity'))
        if sb_mode and _is_sb_delivery_fee(item):
            item['sb_delivery_fee'] = True

    return items


# Keep the old function name as an alias for compatibility
def extract_line_items_sb(text, filepath=None):
    """Alias for backwards compatibility (S&B-specific)."""
    return extract_line_items(text, filepath=filepath, vendor_name='S&B')


# ---------------------------------------------------------------------------
# Main Parse Function
# ---------------------------------------------------------------------------

def _extract_table_fields(filepath):
    """Extract structured fields from pdfplumber tables.

    Many invoices have label/value table pairs like:
        | Date | Invoice # |
        | 1/28/2026 | 550350 |

    This function finds those tables and extracts key fields.
    """
    if not filepath:
        return {}

    tables = extract_tables_from_pdf(filepath)
    fields = {}

    for table in tables:
        if not table or len(table) < 2:
            continue

        for row_idx, row in enumerate(table):
            if not row:
                continue

            # Look for header+value pairs
            headers = [str(cell).strip().lower() if cell else '' for cell in row]
            # Check if this row has field labels and the next row has values
            if row_idx + 1 < len(table):
                values = table[row_idx + 1]
                if not values:
                    continue

                for col_idx, header in enumerate(headers):
                    if col_idx >= len(values):
                        continue
                    val = str(values[col_idx]).strip() if values[col_idx] else ''
                    if not val:
                        continue

                    # Invoice Number
                    if header in ('invoice #', 'invoice#', 'invoice number'):
                        if not fields.get('invoice_number'):
                            fields['invoice_number'] = val

                    # Date
                    if header in ('date', 'invoice date'):
                        if not fields.get('date') and re.match(r'\d', val):
                            fields['date'] = val

                    # Due Date
                    if header in ('due date',):
                        if not fields.get('due_date'):
                            fields['due_date'] = val

                    # PO Number
                    if header in ('po #', 'p.o. no.', 'p.o. number', 'po number',
                                  'purchase order number', 'customer po', 'customer po#'):
                        if not fields.get('po_number') and re.match(r'\d', val):
                            fields['po_number'] = val

                    # Terms
                    if header in ('terms', 'payment terms'):
                        if not fields.get('terms'):
                            fields['terms'] = val

                    # SO Number (for CNC-style invoices where SO is the invoice)
                    if header in ('so no.', 'so no', 'so number'):
                        if not fields.get('so_number'):
                            fields['so_number'] = val

    return fields


def _extract_label_value_pairs(text):
    """Extract fields from label-then-value line patterns.

    Handles the common pattern where labels are on one line and values on the next:
        Date Ship Via Tracking Terms
        2026-01-29 UPS UPSG 1Z8E37A90395308892 NET30
    """
    fields = {}
    lines = text.split('\n')

    def _next_matching(start_idx, pattern, max_ahead=3):
        """Find the next non-empty line matching a regex pattern."""
        for j in range(1, max_ahead + 1):
            if start_idx + j >= len(lines):
                break
            cand = lines[start_idx + j].strip()
            if not cand:
                continue
            if re.match(pattern, cand):
                return cand
        return ""

    for i, line in enumerate(lines):
        line_stripped = line.strip()

        # "Date Invoice #" -> "1/28/2026 550350" (RH, NL)
        if re.match(r'^Date\s+Invoice\s*#', line_stripped, re.IGNORECASE):
            next_line = _next_matching(i, r'^(\d{1,2}/\d{1,2}/\d{2,4}|\d{4}-\d{2}-\d{2})\s+\S+')
            match = re.match(r'^(\S+)\s+(\S+)', next_line)
            if match:
                if not fields.get('date'):
                    fields['date'] = match.group(1)
                if not fields.get('invoice_number'):
                    fields['invoice_number'] = match.group(2)

        # "Invoice Date Due Date Customer # Invoice #" -> "1/27/26 2/10/26 10525 383366-00" (PD)
        if re.match(r'^Invoice\s+Date\s+Due\s+Date', line_stripped, re.IGNORECASE):
            next_line = _next_matching(i, r'^\d{1,2}/\d{1,2}/\d{2,4}\s+\d{1,2}/\d{1,2}/\d{2,4}\s+\S+')
            parts = next_line.split()
            if len(parts) >= 4:
                if not fields.get('date'):
                    fields['date'] = parts[0]
                if not fields.get('due_date'):
                    fields['due_date'] = parts[1]
                if not fields.get('invoice_number'):
                    fields['invoice_number'] = parts[-1]

        # "PO Date PO # Placed By Page #" -> "1/27/26 0037305 ..." (PD)
        if re.match(r'^PO\s+Date\s+PO\s*#', line_stripped, re.IGNORECASE):
            next_line = _next_matching(i, r'^\d{1,2}/\d{1,2}/\d{2,4}\s+\S+')
            parts = next_line.split()
            if len(parts) >= 2:
                if not fields.get('po_number'):
                    fields['po_number'] = parts[1]

        # "P.O. No. Terms Ship Via" -> "0037362 N30 FEDEX" (RH)
        if re.match(r'^P\.?O\.?\s+No\.?\s+Terms', line_stripped, re.IGNORECASE):
            next_line = _next_matching(i, r'^\d+\s+\S+')
            parts = next_line.split()
            if len(parts) >= 2:
                if not fields.get('po_number'):
                    fields['po_number'] = parts[0]
                if not fields.get('terms'):
                    fields['terms'] = parts[1]

        # "P.O. Number Terms Rep Ship Via F.O.B." -> "0035050 Due Upon Receipt MRD..." (NL)
        if re.match(r'^P\.?O\.?\s+Number\s+Terms', line_stripped, re.IGNORECASE):
            next_line = _next_matching(i, r'^\d+\b')
            parts = next_line.split()
            if parts:
                if not fields.get('po_number'):
                    fields['po_number'] = parts[0]

        # "Date Ship Via Tracking Terms" -> "2026-01-29 UPS ..." (II)
        if re.match(r'^Date\s+Ship\s+Via', line_stripped, re.IGNORECASE):
            next_line = _next_matching(i, r'^(\d{4}-\d{2}-\d{2}|\d{1,2}/\d{1,2}/\d{2,4})\s+\S+')
            parts = next_line.split()
            if parts:
                if not fields.get('date'):
                    fields['date'] = parts[0]
                if not fields.get('terms') and len(parts) >= 4:
                    fields['terms'] = parts[-1]

        # "Purchase Order Number Order Date Sales Person Our Order" -> "36485 ..." (II)
        if re.match(r'^Purchase\s+Order\s+Number', line_stripped, re.IGNORECASE):
            next_line = _next_matching(i, r'^\d+\b')
            parts = next_line.split()
            if parts:
                if not fields.get('po_number'):
                    fields['po_number'] = parts[0]

        # "Customer PO# Ship Via Ship Date Tracking Number(s)" -> "0036970 UPS Ground..." (FL)
        if re.match(r'^Customer\s+PO#', line_stripped, re.IGNORECASE):
            next_line = _next_matching(i, r'^\d+\b')
            parts = next_line.split()
            if parts:
                if not fields.get('po_number'):
                    fields['po_number'] = parts[0]

        # "SO No. Customer PO" -> "62957 0037817" (CNC)
        if re.match(r'^SO\s+No\.?\s+Customer\s+PO', line_stripped, re.IGNORECASE):
            next_line = _next_matching(i, r'^\d+\s+\d+')
            parts = next_line.split()
            if len(parts) >= 2:
                if not fields.get('invoice_number'):
                    fields['invoice_number'] = parts[0]
                if not fields.get('po_number'):
                    fields['po_number'] = parts[1]

    return fields


def parse_invoice_text(text, filepath=None):
    """Parse extracted text into structured invoice data.

    Args:
        text: Extracted text from the PDF
        filepath: Path to the PDF file (used for table-based extraction)
    """
    data = {}

    # First, extract fields from structured tables and label/value line pairs
    table_fields = _extract_table_fields(filepath)
    line_pair_fields = _extract_label_value_pairs(text)

    # --- Invoice Number ---
    data['invoice_number'] = parse_field(text, [
        r'Invoice\s*#\s*:?\s*([A-Za-z0-9][\w\-]*\d[\w\-]*)',
        r'Invoice\s+Number\s*:?\s*([A-Za-z0-9][\w\-]*\d[\w\-]*)',
        r'^([A-Z]-\d{5,})$',                  # II: "I-424615" on its own line
    ])
    # Fallback to table/line-pair extracted values
    if not data['invoice_number']:
        data['invoice_number'] = line_pair_fields.get('invoice_number', '') or table_fields.get('invoice_number', '')
    # CNC uses SO number as invoice number
    if not data['invoice_number']:
        data['invoice_number'] = table_fields.get('so_number', '') or line_pair_fields.get('so_number', '')

    # --- Vendor ---
    data['vendor'] = detect_vendor(text)

    # --- Vendor Address ---
    data['vendor_address'] = extract_vendor_address(text)

    # --- Customer Name ---
    data['customer'] = extract_customer_name(text)

    # --- Date ---
    data['date'] = parse_field(text, [
        r'Invoice\s+Date\s*:?\s*(\d{1,2}/\d{1,2}/\d{2,4})',    # "Invoice Date: 1/26/2026"
        r'Invoice\s+Date\s*(\d{1,2}/\d{1,2}/\d{2,4})',          # "Invoice Date 1/26/2026"
        r'(?:^|\n)\s*Date\s*:\s*(\d{4}-\d{2}-\d{2})',           # "Date: 2026-01-29"
        r'(?:^|\n)\s*Date\s*:\s*(\d{1,2}/\d{1,2}/\d{2,4})',     # "Date: 01/30/26"
        r'(?m)^(?!.*(?:Due|Ship|Order|P\.O\.|P\.O|PO)\s+Date).*\bDate\s+(\d{1,2}/\d{1,2}/\d{2,4})',
    ])
    if not data['date']:
        data['date'] = line_pair_fields.get('date', '') or table_fields.get('date', '')

    # --- Due Date ---
    data['due_date'] = parse_field(text, [
        r'Due\s+Date\s*:?\s*(\d{1,2}/\d{1,2}/\d{2,4})',
        r'Due\s+Date\s*(\d{1,2}/\d{1,2}/\d{2,4})',
    ])
    if not data['due_date']:
        data['due_date'] = line_pair_fields.get('due_date', '') or table_fields.get('due_date', '')

    # --- Terms ---
    data['terms'] = parse_field(text, [
        r'Terms\s*:\s*(Net\s*\d+\w*(?:\s+Prox)?)',             # "Terms : Net 30", "Net10th Prox"
        r'Terms\s+(NET\s*\d+)',                                  # "Terms NET30"
        r'Terms\s*:\s*(N\d+)',                                   # "Terms: N30"
        r'Terms\s*:\s*(Due\s+(?:on|Upon)\s+[Rr]eceipt)',       # "Terms: Due on receipt"
        r'(?:Payment\s+)?Terms\s*:\s*(Credit\s+Card[^\n]*)',    # "Payment Terms: Credit Card..."
    ])
    if not data['terms']:
        data['terms'] = line_pair_fields.get('terms', '') or table_fields.get('terms', '')

    # --- PO Number ---
    data['po_number'] = parse_field(text, [
        r'PO\s*#\s*:?\s*(\d+)',                     # "PO #: 0037993", "PO# 0036788"
        r'P\.O\.\s+Number\s+(\d+)',                  # "P.O. Number 0038106"
    ])
    if not data['po_number']:
        data['po_number'] = line_pair_fields.get('po_number', '') or table_fields.get('po_number', '')

    # --- Tracking Number ---
    data['tracking_number'] = parse_field(text, [
        r'Tracking\s*(?:#|No\.?|Number)\s*(?:\(s\))?\s*:?\s*\n?\s*([A-Z0-9]{10,})',
        r'Tracking\s*#?\s*:?\s*([A-Z0-9]{10,})',
    ])

    # --- Shipping Method ---
    data['shipping_method'] = parse_field(text, [
        r'Ship\s+(?:Method|Via)\s*:?\s*([^\n]+)',
        r'Via\s+([^\n]+)',
    ])

    # --- Ship Date ---
    data['ship_date'] = parse_field(text, [
        r'Ship\s+Date\s*:?\s*(\d{1,2}/\d{1,2}/\d{2,4})',
        r'Ship\s+Date\s*(\d{1,2}/\d{1,2}/\d{2,4})',
        r'Shipped\s+(\d{1,2}/\d{1,2}/\d{2,4})',
    ])

    # --- Shipping Tax Code ---
    data['shipping_tax_code'] = parse_field(text, [
        r'Shipping\s+Tax\s+Code\s+(\S+)',
    ])

    # --- Shipping Tax Rate ---
    data['shipping_tax_rate'] = parse_field(text, [
        r'Shipping\s+Tax\s+Rate\s+(\d+)',
    ])

    # --- Subtotal ---
    data['subtotal'] = parse_field(text, [
        r'Subtotal\s*:?\s*\$?([\d,]+\.?\d*)',
        r'Sub\s*-?\s*total\s*:?\s*\$?([\d,]+\.?\d*)',
    ])

    # --- Shipping Cost ---
    # Supports: "Shipping Cost (FedEx...) 12.00", "Drop Ship $5.00",
    #           "Freight $0.00", "FREIGHT OUT $67.00", "FreightEB"
    data['shipping_cost'] = parse_field(text, [
        r'Shipping\s+Cost\s*\([^)]+\)\s*\$?([\d,]+\.?\d*)',   # S&B
        r'(?im)^Drop\s+Ship\s+\d+\.?\d*\s+\d+\.?\d*\s+[\d,]+\.?\d{2}\s+([\d,]+\.?\d{2})\s*$',  # PPE
        r'Drop\s+Ship\s+\$?([\d,]+\.?\d*)',                    # FL, PPE
        r'FREIGHT\s+OUT\s+\$?([\d,]+\.?\d*)',                   # II
        r'Freight\s+\$?([\d,]+\.?\d*)',                         # T14, general
    ])
    if data['shipping_cost'] and not data.get('shipping_description'):
        if re.search(r'Drop\s+Ship', text, re.IGNORECASE):
            data['shipping_description'] = 'Drop Ship'

    # --- Total ---
    data['total'] = _extract_total_amount(text)
    if not data['total']:
        data['total'] = parse_field(text, [
            r'(?:Total\s+USD|Total\s+Amount|Invoice\s+Total|Grand\s+Total|Total\s+Due|^Total)\s*:?\s*\$?([\d,]+\.?\d*)',
            r'(?:^|\n)\s*Total\s+\$?([\d,]+\.?\d*)',
            r'Amount\s+Due\s*:?\s*\$?([\d,]+\.?\d*)',
            r'Balance\s+Due\s+\$?([\d,]+\.?\d*)',
        ])

    # --- Line Items ---
    data['line_items'] = extract_line_items(text, filepath, vendor_name=data.get('vendor'))
    if data['line_items']:
        freight_items = [i for i in data['line_items'] if i.get('is_freight')]
        if freight_items and not data.get('shipping_description'):
            desc = freight_items[0].get('description') or freight_items[0].get('item_number') or 'Freight'
            data['shipping_description'] = desc

    return data


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

    # Step 3: Parse the extracted text (pass filepath for table extraction)
    cb(f"  Parsing invoice data from {filename}...")
    data = parse_invoice_text(text, filepath)

    if not data.get('vendor'):
        data['vendor'] = infer_vendor_from_filename(filename)
    data['vendor'] = normalize_vendor_name(data.get('vendor', ''))

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
