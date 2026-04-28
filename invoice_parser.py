"""Invoice parser: extracts structured data from PDF invoices.

Designed to be vendor-agnostic - works with any invoice format by using
multiple extraction strategies (table-based + text-based) and broad
pattern matching for all fields.
"""
import os
import sys
import re
import csv
import json
from html import unescape
from email.utils import parseaddr
from datetime import datetime, timedelta
from urllib.parse import parse_qs, unquote, urlparse
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
    key_to_mailing_address = {}
    key_to_default_terms = {}
    key_to_due_date_days = {}
    sender_alias_pairs = []
    if not os.path.exists(VENDORS_CSV_PATH):
        return (
            vendors,
            key_to_canonical,
            alias_names,
            key_to_mailing_address,
            key_to_default_terms,
            key_to_due_date_days,
            sender_alias_pairs,
        )
    try:
        with open(VENDORS_CSV_PATH, newline='', encoding='utf-8') as f:
            rows = list(csv.reader(f))
    except Exception:
        return [], {}, [], {}, {}, {}, []
    if not rows:
        return (
            vendors,
            key_to_canonical,
            alias_names,
            key_to_mailing_address,
            key_to_default_terms,
            key_to_due_date_days,
            sender_alias_pairs,
        )

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
        return (
            vendors,
            key_to_canonical,
            alias_names,
            key_to_mailing_address,
            key_to_default_terms,
            key_to_due_date_days,
            sender_alias_pairs,
        )

    def col(row, *names):
        for name in names:
            if name in header:
                idx = header.index(name)
                if idx < len(row):
                    return str(row[idx]).strip()
        return ''

    def set_mailing_address(key, value):
        if not key or not value:
            return
        if key not in key_to_mailing_address:
            key_to_mailing_address[key] = value

    def set_default_terms(key, value):
        if not key or not value:
            return
        if key not in key_to_default_terms:
            key_to_default_terms[key] = value

    def set_due_date_days(key, value):
        if not key or value in (None, ''):
            return
        try:
            parsed = int(str(value).strip())
        except Exception:
            return
        if key not in key_to_due_date_days:
            key_to_due_date_days[key] = parsed

    seen_sender_aliases = set()

    def add_sender_alias(alias, vendor_name):
        cleaned_alias = str(alias or '').strip().lower()
        cleaned_vendor = str(vendor_name or '').strip()
        if not cleaned_alias or not cleaned_vendor:
            return
        pair = (cleaned_alias, cleaned_vendor)
        if pair in seen_sender_aliases:
            return
        seen_sender_aliases.add(pair)
        sender_alias_pairs.append(pair)

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
        mailing_address = col(
            row,
            'mailing_address',
            'vendor_address',
            'address',
            'default_address',
        )
        default_terms = col(row, 'default_terms', 'terms')
        due_date_days = col(row, 'due_date_days')
        if key:
            set_mailing_address(key, mailing_address)
            set_default_terms(key, default_terms)
            set_due_date_days(key, due_date_days)
        aliases_val = col(row, 'aliases', 'alias', 'additional_names', 'invoice_vendor')
        sender_aliases_val = col(row, 'sender_aliases', 'sender_alias', 'email_aliases')
        if not aliases_val and 'skunexus_vendor' in header:
            aliases_val = col(row, 'skunexus_vendor')
        for alias in _split_vendor_aliases(aliases_val):
            if _normalize_vendor_key(alias) == _normalize_vendor_key(vendor):
                continue
            add_alias(alias)
            alias_key = _normalize_vendor_key(alias)
            if alias_key and alias_key not in key_to_canonical:
                key_to_canonical[alias_key] = vendor
            if alias_key:
                set_mailing_address(alias_key, mailing_address)
                set_default_terms(alias_key, default_terms)
                set_due_date_days(alias_key, due_date_days)
        for sender_alias in _split_vendor_aliases(sender_aliases_val):
            add_sender_alias(sender_alias, vendor)

    return (
        vendors,
        key_to_canonical,
        alias_names,
        key_to_mailing_address,
        key_to_default_terms,
        key_to_due_date_days,
        sender_alias_pairs,
    )


(
    VENDOR_LIST,
    VENDOR_KEY_TO_CANONICAL,
    VENDOR_ALIAS_LIST,
    VENDOR_KEY_TO_MAILING_ADDRESS,
    VENDOR_KEY_TO_DEFAULT_TERMS,
    VENDOR_KEY_TO_DUE_DATE_DAYS,
    VENDOR_SENDER_ALIAS_PAIRS,
) = _load_vendor_data()


SHARED_FORWARDER_DOMAINS = (
    '@suspension.randysww.com',
)
SHARED_FORWARDER_EMAILS = {
    'system@sent-via.netsuite.com',
}


def _find_vendor_by_address_alias(text):
    if not text or not VENDOR_ALIAS_LIST:
        return ""
    normalized_text = _normalize_vendor_key(text)
    if not normalized_text:
        return ""
    for alias in VENDOR_ALIAS_LIST:
        alias_text = str(alias or '').strip()
        if not _alias_looks_like_address(alias_text):
            continue
        alias_key = _normalize_vendor_key(alias_text)
        if alias_key and alias_key in normalized_text:
            return alias
    return ""


def _alias_looks_like_address(alias_text):
    alias_text = str(alias_text or '').strip()
    if not alias_text:
        return False
    if not re.search(r'\d', alias_text):
        return False
    if len(re.findall(r'[A-Za-z]', alias_text)) < 4:
        return False
    return len(_normalize_vendor_key(alias_text)) >= 8


def _vendor_text_aliases(vendor_name):
    canonical_vendor = normalize_vendor_name(vendor_name)
    if not canonical_vendor:
        return []
    candidates = [canonical_vendor]
    seen_keys = {_normalize_vendor_key(canonical_vendor)}
    for alias in VENDOR_ALIAS_LIST:
        alias_text = str(alias or '').strip()
        if not alias_text or _alias_looks_like_address(alias_text):
            continue
        alias_key = _normalize_vendor_key(alias_text)
        if not alias_key or alias_key in seen_keys:
            continue
        if VENDOR_KEY_TO_CANONICAL.get(alias_key) != canonical_vendor:
            continue
        seen_keys.add(alias_key)
        candidates.append(alias_text)
    return candidates


def _text_explicitly_mentions_vendor(text, vendor_name):
    if not text or not vendor_name:
        return False
    if not validate_vendor_name(vendor_name):
        return False
    return bool(_find_vendor_in_text_list(text, _vendor_text_aliases(vendor_name)))


def normalize_vendor_name(name):
    """Normalize vendor name to canonical form from vendors.csv."""
    if not name:
        return name
    key = _normalize_vendor_key(name)
    return VENDOR_KEY_TO_CANONICAL.get(key, name)


def get_vendor_default_address(name):
    """Return the configured fallback mailing address for a vendor, if any."""
    if not name:
        return ''
    candidates = [
        _normalize_vendor_key(name),
        _normalize_vendor_key(normalize_vendor_name(name)),
    ]
    for key in candidates:
        if key and key in VENDOR_KEY_TO_MAILING_ADDRESS:
            return VENDOR_KEY_TO_MAILING_ADDRESS[key]
    return ''


def get_vendor_default_terms(name):
    """Return the configured fallback terms for a vendor, if any."""
    if not name:
        return ''
    candidates = [
        _normalize_vendor_key(name),
        _normalize_vendor_key(normalize_vendor_name(name)),
    ]
    for key in candidates:
        if key and key in VENDOR_KEY_TO_DEFAULT_TERMS:
            return VENDOR_KEY_TO_DEFAULT_TERMS[key]
    return ''


def _extract_sender_email(value):
    """Normalize a sender header or plain email into a lowercase email address."""
    raw = str(value or '').strip()
    if not raw:
        return ''
    parsed = parseaddr(raw)[1].strip().lower()
    if parsed:
        return parsed
    match = re.search(
        r'[A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,}',
        raw,
        re.IGNORECASE,
    )
    return match.group(0).lower() if match else ''


def infer_vendor_from_sender(sender_email='', sender_header=''):
    """Infer a vendor from a whitelisted sender alias in vendors.csv."""
    if not VENDOR_SENDER_ALIAS_PAIRS:
        return ''
    email_value = _extract_sender_email(sender_email or sender_header)
    header_value = str(sender_header or sender_email or '').strip().lower()
    for alias, vendor in VENDOR_SENDER_ALIAS_PAIRS:
        alias_value = str(alias or '').strip().lower()
        if not alias_value:
            continue
        if alias_value.startswith('@'):
            if email_value.endswith(alias_value):
                return vendor
            continue
        if '@' in alias_value:
            if email_value == alias_value:
                return vendor
            continue
        if alias_value in header_value:
            return vendor
    return ''


def _is_shared_forwarder_sender(sender_email='', sender_header=''):
    """Return True for transport-only sender addresses shared by many vendors."""
    email_value = _extract_sender_email(sender_email or sender_header)
    header_value = str(sender_header or sender_email or '').strip().lower()
    if email_value in SHARED_FORWARDER_EMAILS:
        return True
    if any(email_value.endswith(domain) for domain in SHARED_FORWARDER_DOMAINS):
        return True
    if any(address in header_value for address in SHARED_FORWARDER_EMAILS):
        return True
    return any(domain in header_value for domain in SHARED_FORWARDER_DOMAINS)


def _find_vendor_by_sender_alias_in_text(text):
    """Find a vendor when the forwarded body contains a configured vendor email alias."""
    if not text or not VENDOR_SENDER_ALIAS_PAIRS:
        return ''

    text_value = str(text or '').lower()
    for alias, vendor in VENDOR_SENDER_ALIAS_PAIRS:
        alias_value = str(alias or '').strip().lower()
        if not alias_value:
            continue
        if alias_value.startswith('@'):
            domain = re.escape(alias_value[1:])
            if re.search(r'[a-z0-9._%+\-]+@' + domain + r'\b', text_value, re.IGNORECASE):
                return vendor
            if re.search(r'(?<![a-z0-9.\-])' + domain + r'\b', text_value, re.IGNORECASE):
                return vendor
            continue
        if '@' in alias_value:
            if re.search(r'(?<![a-z0-9._%+\-])' + re.escape(alias_value) + r'\b', text_value, re.IGNORECASE):
                return vendor
            continue
        if alias_value in text_value:
            return vendor
    return ''


def _infer_vendor_from_email_content(subject='', message_text=''):
    """Resolve a vendor from forwarded email body first, then subject as fallback."""
    for text in (str(message_text or '').strip(), str(subject or '').strip()):
        if not text:
            continue
        vendor = _find_vendor_in_text(text)
        if vendor:
            return normalize_vendor_name(vendor)
        vendor = _find_vendor_by_address_alias(text)
        if vendor:
            return normalize_vendor_name(vendor)
        vendor = _find_vendor_by_sender_alias_in_text(text)
        if vendor:
            return normalize_vendor_name(vendor)
    return ''


def _infer_vendor_from_shared_sender_content(sender_email='', sender_header='', subject='', message_text=''):
    """Resolve vendors that share a sender mailbox by reading the forwarded email content."""
    if not _is_shared_forwarder_sender(sender_email=sender_email, sender_header=sender_header):
        return ''
    return _infer_vendor_from_email_content(subject=subject, message_text=message_text)


def infer_vendor_from_email_metadata(sender_email='', sender_header='', subject='', message_text=''):
    """Infer a vendor from sender metadata plus body/subject confirmations when needed."""
    shared_sender_vendor = _infer_vendor_from_shared_sender_content(
        sender_email=sender_email,
        sender_header=sender_header,
        subject=subject,
        message_text=message_text,
    )
    if shared_sender_vendor:
        return shared_sender_vendor
    sender_email_value = _extract_sender_email(sender_email or sender_header)
    if sender_email_value in {'system@sent-via.netsuite.com'} or sender_email_value.endswith('@suspension.randysww.com'):
        return ''
    return infer_vendor_from_sender(sender_email=sender_email, sender_header=sender_header)


def infer_vendor_from_folder_marker(filepath):
    """Infer vendor from a training-folder marker file when present."""
    if not filepath:
        return ''
    folder_vendor_map = {
        'ps': 'Power Stroke Products',
        'hc': 'Hamilton Cams',
        'bch': 'Bosch',
        'bdp': 'Beans Diesel Performance',
        'all': 'Diesel Forward',
        'crl': 'Carli Suspension - $10 DS Fee',
        'ico': 'Icon Vehicle Dynamics',
        'co': 'Cognito Motorsports',
    }
    try:
        abs_path = os.path.abspath(filepath)
        path_parts = {
            part.lower()
            for part in os.path.normpath(abs_path).split(os.sep)
            if part
        }
        if 'training' not in path_parts:
            return ''
        folder = os.path.dirname(abs_path)
    except Exception:
        return ''

    for marker_name in ('vendor.txt', '.vendor', 'folder_vendor.txt'):
        marker_path = os.path.join(folder, marker_name)
        if not os.path.exists(marker_path):
            continue
        try:
            with open(marker_path, 'r', encoding='utf-8') as f:
                for raw_line in f:
                    line = str(raw_line or '').strip()
            if not line or line.startswith('#'):
                continue
            line = re.sub(r'^(?:vendor|name)\s*:\s*', '', line, flags=re.IGNORECASE)
            return normalize_vendor_name(line)
        except Exception:
            continue
    folder_name = os.path.basename(folder).strip().lower()
    mapped_vendor = folder_vendor_map.get(folder_name)
    if mapped_vendor:
        return normalize_vendor_name(mapped_vendor)
    return ''


def get_vendor_due_date_days(name):
    """Return configured due-date offset in days for a vendor, if any."""
    if not name:
        return None
    candidates = [
        _normalize_vendor_key(name),
        _normalize_vendor_key(normalize_vendor_name(name)),
    ]
    for key in candidates:
        if key and key in VENDOR_KEY_TO_DUE_DATE_DAYS:
            return VENDOR_KEY_TO_DUE_DATE_DAYS[key]
    return None


def _derive_due_date_from_bill_date(bill_date, days_after_bill_date):
    """Derive a due date from bill date, preserving the original date style."""
    bill_date = str(bill_date or '').strip()
    if not bill_date:
        return ''
    try:
        days = int(days_after_bill_date)
    except Exception:
        return ''

    if days == 0:
        return bill_date

    parsed = None
    for fmt in ('%m/%d/%Y', '%m/%d/%y', '%Y-%m-%d'):
        try:
            parsed = (datetime.strptime(bill_date, fmt), fmt)
            break
        except ValueError:
            continue

    if not parsed:
        return ''

    dt, fmt = parsed
    due_dt = dt + timedelta(days=days)
    if fmt == '%Y-%m-%d':
        return due_dt.strftime('%Y-%m-%d')
    year = due_dt.strftime('%y') if fmt == '%m/%d/%y' else due_dt.strftime('%Y')
    return f"{due_dt.month}/{due_dt.day}/{year}"


def _normalize_date_value(value):
    """Normalize vendor-specific date strings to M/D/YYYY when possible."""
    text = re.sub(r'\s+', ' ', str(value or '')).strip()
    if not text:
        return ''
    for fmt in (
        '%m/%d/%Y',
        '%m/%d/%y',
        '%Y-%m-%d',
        '%b %d, %Y',
        '%B %d, %Y',
        '%d-%b-%y',
        '%d-%b-%Y',
        '%d-%B-%y',
        '%d-%B-%Y',
    ):
        try:
            parsed = datetime.strptime(text, fmt)
            return f"{parsed.month}/{parsed.day}/{parsed.year}"
        except ValueError:
            continue
    return text


def _find_vendor_in_text_list(text, vendor_list):
    if not text or not vendor_list:
        return ""

    text_lower = text.lower()
    matches = []

    # Avoid matching known customers as vendors
    customer_phrases = [c for c in KNOWN_CUSTOMERS if len(c) >= 4]

    for vendor in vendor_list:
        v_lower = vendor.lower()
        if any(cust in v_lower for cust in customer_phrases):
            continue

        tokens = re.findall(r'[a-z0-9]+', v_lower)
        if not tokens:
            continue

        if len(tokens) == 1 and len(tokens[0]) <= 4:
            pattern = r'(?<![a-z0-9])' + re.escape(tokens[0]) + r'(?![a-z0-9])'
        else:
            pattern = r'(?<![a-z0-9])' + r'[\W_]+'.join(re.escape(tok) for tok in tokens) + r'(?![a-z0-9])'

        match = re.search(pattern, text_lower)
        if match:
            matches.append((len(vendor), match.start(), vendor))

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
    if re.match(
        r'^\s*(?:Attn|Attention|Tracking|Reference|Ship\s+Via|Customer\s+No|Sales\s+Account\s+Number)\b\s*:?',
        str(text or ''),
        re.IGNORECASE,
    ):
        return False
    if re.match(r'^\s*Printed\s+\d{1,2}/\d{1,2}/\d{2,4}\b', str(text or ''), re.IGNORECASE):
        return False
    collapsed = re.sub(r'\s+', '', str(text or '')).lower()
    if collapsed in ('reprint', 'invoice', 'page'):
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


def _is_reprint_vendor_name(name):
    """Return True when a detected vendor is just a reprint stamp/header."""
    text = str(name or '').strip()
    collapsed = re.sub(r'\s+', '', text).lower()
    return collapsed == 'reprint' or bool(
        re.match(r'^Printed\s+\d{1,2}/\d{1,2}/\d{2,4}\b', text, re.IGNORECASE)
    )


def _text_matches_pt_layout(text):
    """Return True when text matches the PT/Diesel USA invoice structure."""
    if not text:
        return False
    has_header = bool(re.search(
        r'Part\s+Number\s+Order\s+Ship\s+B\/?O\s+Description\s+Unit\s+Net(?:\s+TE)?\s+Value',
        text,
        re.IGNORECASE,
    ))
    has_account_block = bool(re.search(r'\bBr\s+Accnt\b', text, re.IGNORECASE))
    has_inv_ord = bool(re.search(r'Inv\s*#\s*\d{2}\s+\d{5,}\s+Ord#\s*\d+', text, re.IGNORECASE))
    return has_header and has_account_block and has_inv_ord


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


def extract_layout_text_from_pdf(filepath):
    """Extract text using pdfplumber's layout mode to preserve column spacing."""
    text = ""
    try:
        with pdfplumber.open(filepath) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text(layout=True)
                if page_text:
                    text += page_text + "\n"
    except Exception:
        pass
    return text.strip()


def _extract_first_page_words(filepath):
    """Return first-page words plus page width for coordinate-based vendor parsers."""
    if not filepath:
        return [], 0
    try:
        with pdfplumber.open(filepath) as pdf:
            if not pdf.pages:
                return [], 0
            page = pdf.pages[0]
            words = page.extract_words(x_tolerance=1, y_tolerance=1, keep_blank_chars=False) or []
            return words, float(getattr(page, 'width', 0) or 0)
    except Exception:
        return [], 0


def _group_words_into_lines(words, tolerance=2.5):
    """Group extracted words into visual lines using their top coordinate."""
    if not words:
        return []

    sorted_words = sorted(words, key=lambda w: (float(w.get('top', 0)), float(w.get('x0', 0))))
    lines = []
    current = []
    current_top = None

    for word in sorted_words:
        top = float(word.get('top', 0))
        if current and current_top is not None and abs(top - current_top) > tolerance:
            lines.append(sorted(current, key=lambda w: float(w.get('x0', 0))))
            current = [word]
            current_top = top
            continue
        if not current:
            current_top = top
        current.append(word)

    if current:
        lines.append(sorted(current, key=lambda w: float(w.get('x0', 0))))
    return lines


def _words_to_line_text(words):
    return ' '.join(str(word.get('text', '')).strip() for word in words if str(word.get('text', '')).strip()).strip()


def _extract_column_lines_from_words(words, min_x, max_x, min_top, max_top, tolerance=2.5):
    """Extract text lines from a coordinate-bounded column region."""
    selected = []
    for word in words or []:
        x0 = float(word.get('x0', 0))
        top = float(word.get('top', 0))
        if x0 < min_x or x0 >= max_x:
            continue
        if top <= min_top or top >= max_top:
            continue
        selected.append(word)

    lines = []
    for line_words in _group_words_into_lines(selected, tolerance=tolerance):
        text = _words_to_line_text(line_words)
        if text:
            lines.append(text)
    return lines


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


def _extract_ppe_drop_ship_quantity(text):
    """Extract PPE drop-ship quantity from the Order Qty column."""
    if not text:
        return ''

    patterns = [
        r'(?im)^Drop\s+Ship\s+(\d+\.?\d*)\s+\d+\.?\d*\s+[\d,]+\.?\d{2}(?:\s+[\d,]+\.?\d{2})?\s*$',
        r'(?im)^Drop\s+Ship\s+(\d+\.?\d*)\s+[\d,]+\.?\d{2}\s*$',
    ]
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
        if match:
            return _normalize_qty(match.group(1))
    return ''


def _extract_ppe_total_usd(text):
    """Extract PPE invoice total from footer totals, preferring explicit 'Total USD'."""
    if not text:
        return ''

    lines = text.splitlines()
    if not lines:
        return ''

    # PPE table headers use "Total Price" for item rows; this is not invoice total.
    table_header_tokens = ('item/description', 'order qty', 'invoiced qt', 'unit price', 'disc %')
    strong_tokens = ('total usd', 'invoice total', 'grand total', 'total amount', 'amount due', 'balance due')

    explicit_usd = []
    footer_candidates = []

    inline_usd_re = re.compile(
        r'(?i)\btotal\s+usd\b\s*:?\s*\$?\s*([0-9]{1,3}(?:,[0-9]{3})*(?:\.\d{1,2})?|[0-9]+(?:\.\d{1,2})?)'
    )
    label_only_usd_re = re.compile(r'(?i)\btotal\s+usd\b\s*:?\s*$')

    for idx, raw_line in enumerate(lines):
        line = raw_line.strip()
        if not line:
            continue
        lower = line.lower()

        # Fast path: explicit Total USD on same line.
        inline_match = inline_usd_re.search(line)
        if inline_match:
            normalized = _normalize_amount_string(inline_match.group(1))
            if normalized:
                explicit_usd.append((idx, normalized))
            continue

        # "Total USD" label on one line, amount on the next line.
        if label_only_usd_re.search(line):
            for j in range(idx + 1, min(len(lines), idx + 3)):
                next_line = lines[j].strip()
                if not next_line:
                    continue
                amounts = _extract_amounts_from_line(next_line)
                if amounts:
                    normalized = _normalize_amount_string(amounts[0][1])
                    if normalized:
                        explicit_usd.append((j, normalized))
                    break
            continue

        # Generic PPE footer total detection.
        label_kind = None
        if any(tok in lower for tok in strong_tokens):
            label_kind = 'strong'
        elif 'total price' in lower:
            # Ignore the PPE line-item table header.
            if any(tok in lower for tok in table_header_tokens):
                continue
            # Allow footer variants that may say "Total Price".
            label_kind = 'strong'
        elif 'total' in lower:
            label_kind = 'total'
        elif 'subtotal' in lower:
            label_kind = 'subtotal'

        if not label_kind:
            continue

        amounts = _extract_amounts_from_line(line)
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
            priority = 0
            if label_kind == 'strong':
                priority = 3
            elif label_kind == 'total':
                priority = 2
            elif label_kind == 'subtotal':
                priority = 1
            footer_candidates.append({
                'line_index': idx + lookahead,
                'value': value,
                'raw': raw,
                'priority': priority,
            })

    # Prefer explicit Total USD when present.
    if explicit_usd:
        explicit_usd.sort(key=lambda x: x[0])
        return explicit_usd[-1][1]

    if not footer_candidates:
        return ''

    # Focus on the bottom-most total-like group, then pick the largest amount.
    max_idx = max(c['line_index'] for c in footer_candidates)
    bottom_group = [c for c in footer_candidates if c['line_index'] >= (max_idx - 12)]
    if not bottom_group:
        bottom_group = footer_candidates

    bottom_group.sort(key=lambda c: (c['value'], c['priority'], c['line_index']), reverse=True)
    return _normalize_amount_string(bottom_group[0]['raw'])


def _extract_fleece_total(text):
    """Extract Fleece total from footer lines without matching ZIP codes."""
    if not text:
        return ''

    patterns = [
        r'Amount\s+Due\s*:?\s*\$?([\d,]+\.\d{2})',
        r'(?im)^Total\s+\$?([\d,]+\.\d{2})\s*$',
        r'(?im)^Subtotal\s+\$?([\d,]+\.\d{2})\s*$',
    ]
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
        if match:
            return _normalize_amount_string(match.group(1))
    return ''


def _has_fleece_stock_order_marker(text):
    """Return True when a Fleece invoice advertises the stocking-order discount."""
    if not text:
        return False
    return bool(
        re.search(
            r'5%\s+Off\s+Stocking\s+Orders\s+Over\s+\$?\s*30,?000',
            text,
            re.IGNORECASE,
        )
    )


def _extract_turn14_footer_discount_item(text):
    """Extract a footer discount from Turn 14 invoices as a synthetic line item."""
    if not text:
        return None

    for raw_line in text.splitlines():
        line = str(raw_line or '').strip()
        if not line or not re.match(r'^Discount\b', line, re.IGNORECASE):
            continue
        if '%' in line:
            continue
        if '$' not in line and '(' not in line:
            continue

        amount_match = re.search(r'(?P<amount>-?\$?[\d,]+\.\d{2}|\(\$?[\d,]+\.\d{2}\))', line)
        if not amount_match:
            continue

        amount_value = _parse_amount_value(amount_match.group('amount'))
        if amount_value is None:
            continue
        if amount_value > 0:
            amount_value = -amount_value
        if abs(amount_value) < 0.005:
            return None

        amount_text = f"{amount_value:.2f}"
        return {
            'item_number': 'DPP DISCOUNT',
            'quantity': '1',
            'units': 'Each',
            'description': '',
            'unit_price': amount_text,
            'amount': amount_text,
            'is_discount': True,
            'qb_type_override': 'Category Details',
            'qb_category_override': 'Freight and shipping costs',
            'qb_product_service_override': 'Shipping',
            'qb_sku_override': 'DPP DISCOUNT',
        }

    return None


def _is_discount_line_item(item):
    """Return True when a parsed line item represents a discount row."""
    if not item:
        return False
    item_num = str(item.get('item_number', '')).lower()
    desc = str(item.get('description', '')).lower()
    return bool(item.get('is_discount')) or ('discount' in item_num) or ('discount' in desc)


def _apply_export_overrides(item, *, row_type=None, category=None, product_service=None, sku=None):
    """Attach explicit export-field overrides to a parsed line item."""
    if not item:
        return item
    if row_type is not None:
        item['qb_type_override'] = row_type
    if category is not None:
        item['qb_category_override'] = category
    if product_service is not None:
        item['qb_product_service_override'] = product_service
    if sku is not None:
        item['qb_sku_override'] = sku
    return item


def _apply_item_style_discount_overrides(line_items):
    """Restore discount rows to the expected item-style export mapping."""
    for item in line_items or []:
        if not _is_discount_line_item(item):
            continue
        _apply_export_overrides(
            item,
            row_type='Item Details',
            category='Purchases',
            product_service='Inventory Item (Sellable Item)',
            sku='DPP Discount',
        )


def _apply_redhead_discount_overrides(line_items):
    """Restore Red Head discount rows to the expected item-style export mapping."""
    _apply_item_style_discount_overrides(line_items)


def _apply_stock_order_summary(data, description='STOCK ORDER', customer=''):
    """Collapse a stock order into a single summary row."""
    data['stock_order'] = True
    data['stock_order_description'] = description
    if customer:
        data['customer'] = customer
    data['line_items'] = []
    data['shipping_cost'] = ''
    data['shipping_description'] = ''
    data['total'] = ''
    return data


def _matches_internal_stock_customer_hint(value):
    """Return True when a parsed customer points to our company or warehouse address."""
    text = re.sub(r'\s+', ' ', str(value or '')).strip(' ,')
    if not text:
        return False
    if _OUR_COMPANY_NAMES.search(text):
        return True
    return bool(
        re.search(
            r'\b(?:6200\s+E(?:ast)?\.?\s+Main(?:\s+Avenue|\s+Ave\.?)?|E(?:ast)?\.?\s+Main(?:\s+Avenue|\s+Ave\.?)|5204\s+E(?:ast)?\.?\s+Broadway(?:\s+Avenue|\s+Ave\.?)?)\b',
            text,
            re.IGNORECASE,
        )
    )


_OUR_ADDRESS_PATTERNS = [
    re.compile(r'6200\s+E\.?\s+Main', re.IGNORECASE),
    re.compile(r'5204\s+E(?:ast)?\.?\s+Broadway', re.IGNORECASE),
    re.compile(r'spokane\s+valley', re.IGNORECASE),
    re.compile(r'\b9921(?:2|6)\b'),
]

_OUR_COMPANY_NAMES = re.compile(
    r'diesel\s+power\s+products|power\s+products\s+unlimited',
    re.IGNORECASE,
)

_TABLE_HEADER_RE = re.compile(
    r'\b(invoice\s+date|due\s+date|po\s+date|line\s+product|qty\s+part|'
    r'order\s+qty|unit\s+price|subtotal|ship\s+method|payment\s+terms|'
    r'fob\s+point|terms\s+taken)\b',
    re.IGNORECASE,
)


def _extract_ship_to_lines(text):
    """Return lines that belong to the Ship To block, handling side-by-side columns.

    For side-by-side layouts (Bill To / Ship To on same line), extracts only the
    right-hand (Ship To) portion of each line.
    """
    if not text:
        return []

    lines = text.split('\n')
    # Find the "Bill To Ship To" or "Ship To" header line
    header_idx = None
    side_by_side = False
    for i, line in enumerate(lines):
        if re.search(r'bill\s+to\s+ship\s+to|ship\s+to\s+bill\s+to', line, re.IGNORECASE):
            header_idx = i
            side_by_side = True
            break
        if re.search(r'^\s*ship\s+to\b', line, re.IGNORECASE):
            header_idx = i
            break

    if header_idx is None:
        return []

    ship_to_lines = []
    for line in lines[header_idx + 1:header_idx + 12]:
        if _TABLE_HEADER_RE.search(line):
            break
        if side_by_side:
            # Split on the first gap of 2+ spaces to separate Bill To / Ship To columns
            col_split = re.search(r'\s{2,}', line)
            if col_split:
                ship_part = line[col_split.end():].strip()
            else:
                ship_part = line.strip()
        else:
            ship_part = line.strip()
        if ship_part:
            ship_to_lines.append(ship_part)

    return ship_to_lines


_OTHER_ZIP_RE = re.compile(r'\b(?!(?:99212|99216)\b)\d{5}\b')
_OTHER_STREET_RE = re.compile(r'\d+\s+\w.*(?:St|Ave|Rd|Blvd|Dr|Ln|Way|Hwy|Street|Avenue|Road)\b', re.IGNORECASE)


def _our_address_line_clean(line):
    """True if line contains our address signal without a foreign address mixed in."""
    if not any(p.search(line) for p in _OUR_ADDRESS_PATTERNS):
        return False
    # If the line also has a non-99212 zip or a different street address, it's mixed
    if _OTHER_ZIP_RE.search(line):
        return False
    if _OTHER_STREET_RE.search(line) and not re.search(
        r'(?:6200\s+E\.?\s+Main|5204\s+E(?:ast)?\.?\s+Broadway)',
        line,
        re.IGNORECASE,
    ):
        return False
    return True


def _ship_to_block_lines(text):
    """Return only the ship-to address block, stopping when a foreign address appears.

    Walks ship-to lines and stops as soon as a non-our-address zip or street is seen,
    preventing bill-to column overflow from being included.
    """
    ship_lines = _extract_ship_to_lines(text)
    result = []
    for line in ship_lines:
        # Stop if we hit a foreign zip code — indicates another address mixed in
        if _OTHER_ZIP_RE.search(line):
            break
        result.append(line)
    return result


def _ship_to_our_address(text):
    """Return True if the ship-to block is our warehouse address."""
    return _ship_to_our_address_from_lines(_ship_to_block_lines(text))


def _will_call_customer_from_ship_to(text):
    """If ship-to is our address but has a customer name, return it (will call).

    Returns the customer name string, or '' if it's a plain stock order.
    """
    return _will_call_customer_from_lines(_ship_to_block_lines(text))


def _extract_ii_total(text):
    """Extract Industrial Injection total from explicit footer lines."""
    if not text:
        return ''

    patterns = [
        r'(?im)^Total:\s*\$?([\d,]+\.\d{2})\s*$',
        r'(?im)^Subtotal:\s*\$?([\d,]+\.\d{2})\s*$',
        r'Amount\s+Due\s*:?\s*\$?([\d,]+\.\d{2})',
    ]
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
        if match:
            return _normalize_amount_string(match.group(1))
    return ''


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

    # Strategy 3b: PT invoices can lose their readable vendor name on reprints,
    # but the invoice structure is still distinctive.
    if re.search(r'Diesel\s+USA\s+Group', text, re.IGNORECASE):
        return 'Performance Turbochargers'
    if _text_matches_pt_layout(text):
        return 'Performance Turbochargers'

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
                    r'^(diesel\s+power(?:\s+products)?|power\s+products\s+unlimited|dpp)\b(?:\s*/\s*)*',
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


def _looks_like_carrier_shipping(text):
    """Return True when a line item clearly describes carrier shipping service."""
    combined = re.sub(r'\s+', ' ', str(text or '')).strip().lower()
    if not combined:
        return False
    return bool(re.search(r'\b(?:fedex|ups|usps|ground delivery|home delivery)\b', combined))


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
            elif 'ship qty' in header or 'invoiced qt' in header or 'invoiced qty' in header:
                if 'ship_quantity' not in col_map:
                    col_map['ship_quantity'] = col_idx
                if 'quantity' not in col_map:
                    col_map['quantity'] = col_idx
                    recognized += 1
            elif 'bo qty' in header or 'backorder qty' in header:
                if 'backorder_quantity' not in col_map:
                    col_map['backorder_quantity'] = col_idx
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


def _extract_sb_shipping_cost(text):
    """Extract S&B shipping from the footer line, allowing nested carrier labels."""
    if not text:
        return ''
    patterns = [
        r'(?im)^\s*Shipping\s+Cost\b[^\n]*?\$?([\d,]+\.\d{2})\s*$',
        r'(?im)^\s*Shipping\s+\$?([\d,]+\.\d{2})\s*$',
    ]
    for pattern in patterns:
        match = re.search(pattern, text)
        if match:
            return _clean_price(match.group(1))
    return ''


def _extract_dd_shipping_cost(text):
    """Extract Dynomite Diesel shipping from a plain footer row above total."""
    if not text:
        return ''
    match = re.search(r'(?im)^\s*SHIPPING\s+\$?([\d,]+\.\d{2})\s*$', text)
    if not match:
        return ''
    return _clean_price(match.group(1))


def _extract_dd_customer_from_text(text):
    """Extract Dynomite Diesel ship-to customer from OCR text when columns collapse."""
    if not text:
        return ''

    lines = [str(line).strip() for line in str(text).splitlines() if str(line).strip()]
    for idx, line in enumerate(lines):
        if not re.search(r'\bBILL\s+TO\s+SHIP\s+TO\s+INVOICE\b', line, re.IGNORECASE):
            continue

        for candidate in lines[idx + 1: idx + 4]:
            base = re.split(r'\bDATE\b', candidate, maxsplit=1, flags=re.IGNORECASE)[0].strip(' ,')
            if not base:
                continue

            match = re.search(
                r'(?i)(?:\d+\s+.*?\b(?:Ave|Avenue|St|Street|Rd|Road|Blvd|Boulevard|Dr|Drive|Ln|Lane|Way|Hwy|Highway)\b\s+)?'
                r'([A-Z][A-Za-z\'\.-]+(?:\s+[A-Z][A-Za-z\'\.-]+){1,3})\s*$',
                base,
            )
            if not match:
                continue

            customer = _clean_ship_to_contact_name(match.group(1))
            if customer:
                return customer

    return ''


def _extract_sb_new_template_customer(text):
    """Extract the ship-to customer name from S&B's new side-by-side template."""
    if not text:
        return ''
    match = re.search(
        r'(?im)^\s*(.+?)\s+Diesel\s+Power\s+Products\s+DBA\s+Power\s+Products\s+Unlimited,\s*Inc\.?\s*505\s*$',
        text,
    )
    if not match:
        return ''
    customer = re.sub(r'\s+', ' ', match.group(1)).strip(' ,')
    if customer and customer.lower() not in KNOWN_CUSTOMERS:
        return customer
    return ''


def _clean_sb_customer_name(value):
    """Remove known S&B routing/sales suffixes from customer names."""
    text = re.sub(r'\s+', ' ', str(value or '')).strip(' ,')
    if not text:
        return ''
    text = re.sub(r'\s+Coop\s+Rasmussen\s*$', '', text, flags=re.IGNORECASE).strip(' ,')
    return text


def _is_sb_new_template(text):
    """Return True when the S&B invoice matches the newer no-letterhead template."""
    if not text:
        return False
    if re.search(r'\bShip\s+To\s+Bill\s+To\b', text, re.IGNORECASE):
        return False
    return bool(_extract_sb_new_template_customer(text))


def _is_ppe_vendor_name(name):
    """Return True if vendor name looks like Pacific Performance Engineering."""
    key = _normalize_vendor_key(name or '')
    if key and 'pacificperformanceengineering' in key:
        return True
    canonical = normalize_vendor_name(name or '')
    canonical_key = _normalize_vendor_key(canonical or '')
    if canonical_key and 'pacificperformanceengineering' in canonical_key:
        return True
    return False


def _is_fl_vendor_name(name):
    """Return True if vendor name looks like Fleece Performance Engineering."""
    key = _normalize_vendor_key(name or '')
    if key and 'fleeceperformanceengineering' in key:
        return True
    canonical = normalize_vendor_name(name or '')
    canonical_key = _normalize_vendor_key(canonical or '')
    if canonical_key and 'fleeceperformanceengineering' in canonical_key:
        return True
    return False


def _is_turn14_vendor_name(name):
    """Return True if vendor name looks like Turn 14 Distribution."""
    key = _normalize_vendor_key(name or '')
    if key and 'turn14distribution' in key:
        return True
    canonical = normalize_vendor_name(name or '')
    canonical_key = _normalize_vendor_key(canonical or '')
    if canonical_key and 'turn14distribution' in canonical_key:
        return True
    return False


def _is_redhead_vendor_name(name):
    """Return True if vendor name looks like Red-Head Steering Gears."""
    key = _normalize_vendor_key(name or '')
    if key and 'redheadsteeringgears' in key:
        return True
    canonical = normalize_vendor_name(name or '')
    canonical_key = _normalize_vendor_key(canonical or '')
    if canonical_key and 'redheadsteeringgears' in canonical_key:
        return True
    return False


def _is_ii_vendor_name(name):
    """Return True if vendor name looks like Industrial Injection."""
    key = _normalize_vendor_key(name or '')
    if key and 'industrialinjection' in key:
        return True
    canonical = normalize_vendor_name(name or '')
    canonical_key = _normalize_vendor_key(canonical or '')
    if canonical_key and 'industrialinjection' in canonical_key:
        return True
    return False


def _is_pd_vendor_name(name):
    """Return True if vendor name looks like Power Distributing."""
    key = _normalize_vendor_key(name or '')
    if key and 'powerdistributing' in key:
        return True
    canonical = normalize_vendor_name(name or '')
    canonical_key = _normalize_vendor_key(canonical or '')
    if canonical_key and 'powerdistributing' in canonical_key:
        return True
    return False


def _is_daystar_vendor_name(name):
    """Return True if vendor name looks like Daystar."""
    key = _normalize_vendor_key(name or '')
    if key and key == 'daystar':
        return True
    canonical = normalize_vendor_name(name or '')
    canonical_key = _normalize_vendor_key(canonical or '')
    return bool(canonical_key and canonical_key == 'daystar')


def _allow_global_ship_to_stock_detection(vendor_name):
    """Return True when generic ship-to stock/will-call detection should run."""
    if _is_pd_vendor_name(vendor_name):
        return False
    return True


def _is_holley_vendor_name(name):
    """Return True if vendor name looks like Holley Performance Brands."""
    key = _normalize_vendor_key(name or '')
    if key and 'holleyperformancebrands' in key:
        return True
    canonical = normalize_vendor_name(name or '')
    canonical_key = _normalize_vendor_key(canonical or '')
    return bool(canonical_key and 'holleyperformancebrands' in canonical_key)


def _is_valair_vendor_name(name):
    """Return True if vendor name looks like Valair Clutch / ValAir, Inc."""
    key = _normalize_vendor_key(name or '')
    if key and key in {
        _normalize_vendor_key('Valair Clutch'),
        _normalize_vendor_key('ValAir, Inc.'),
        _normalize_vendor_key('ValAir Inc'),
        _normalize_vendor_key('ValAir'),
    }:
        return True
    canonical = normalize_vendor_name(name or '')
    canonical_key = _normalize_vendor_key(canonical or '')
    return bool(canonical_key and canonical_key == _normalize_vendor_key('Valair Clutch'))


def _is_ats_vendor_name(name):
    """Return True if vendor name looks like ATS Diesel Performance."""
    key = _normalize_vendor_key(name or '')
    if key and 'atsdieselperformance' in key:
        return True
    canonical = normalize_vendor_name(name or '')
    canonical_key = _normalize_vendor_key(canonical or '')
    return bool(canonical_key and canonical_key == _normalize_vendor_key('ATS Diesel Performance'))


def _is_isspro_vendor_name(name):
    """Return True if vendor name looks like Isspro."""
    key = _normalize_vendor_key(name or '')
    if key and 'isspro' in key:
        return True
    canonical = normalize_vendor_name(name or '')
    canonical_key = _normalize_vendor_key(canonical or '')
    return bool(canonical_key and 'isspro' in canonical_key)


def _is_power_stroke_vendor_name(name):
    """Return True if vendor name looks like Power Stroke Products."""
    key = _normalize_vendor_key(name or '')
    if key and 'powerstrokeproducts' in key:
        return True
    canonical = normalize_vendor_name(name or '')
    canonical_key = _normalize_vendor_key(canonical or '')
    return bool(canonical_key and canonical_key == _normalize_vendor_key('Power Stroke Products'))


def _is_hamilton_vendor_name(name):
    """Return True if vendor name looks like Hamilton Cams."""
    key = _normalize_vendor_key(name or '')
    if key and 'hamiltoncams' in key:
        return True
    canonical = normalize_vendor_name(name or '')
    canonical_key = _normalize_vendor_key(canonical or '')
    return bool(canonical_key and canonical_key == _normalize_vendor_key('Hamilton Cams'))


def _is_beans_vendor_name(name):
    """Return True if vendor name looks like Beans Diesel Performance."""
    key = _normalize_vendor_key(name or '')
    if key and ('beansdieselperformance' in key or 'beanmachine' in key):
        return True
    canonical = normalize_vendor_name(name or '')
    canonical_key = _normalize_vendor_key(canonical or '')
    return bool(canonical_key and canonical_key == _normalize_vendor_key('Beans Diesel Performance'))


def _is_bosch_vendor_name(name):
    """Return True if vendor name looks like Bosch / Robert Bosch LLC."""
    key = _normalize_vendor_key(name or '')
    if key and ('bosch' in key or 'robertboschllc' in key):
        return True
    canonical = normalize_vendor_name(name or '')
    canonical_key = _normalize_vendor_key(canonical or '')
    return bool(canonical_key and canonical_key == _normalize_vendor_key('Bosch'))


def _is_diesel_forward_vendor_name(name):
    """Return True if vendor name looks like Diesel Forward / Alliant Power."""
    key = _normalize_vendor_key(name or '')
    if key and ('dieselforward' in key or 'alliantpower' in key):
        return True
    canonical = normalize_vendor_name(name or '')
    canonical_key = _normalize_vendor_key(canonical or '')
    return bool(canonical_key and canonical_key == _normalize_vendor_key('Diesel Forward'))


def _is_carli_vendor_name(name):
    """Return True if vendor name looks like Carli Suspension."""
    key = _normalize_vendor_key(name or '')
    if key and 'carlisuspension' in key:
        return True
    canonical = normalize_vendor_name(name or '')
    canonical_key = _normalize_vendor_key(canonical or '')
    return bool(canonical_key and canonical_key == _normalize_vendor_key('Carli Suspension - $10 DS Fee'))


def _is_cognito_vendor_name(name):
    """Return True if vendor name looks like Cognito Motorsports."""
    key = _normalize_vendor_key(name or '')
    if key and 'cognitomotorsports' in key:
        return True
    canonical = normalize_vendor_name(name or '')
    canonical_key = _normalize_vendor_key(canonical or '')
    return bool(canonical_key and canonical_key == _normalize_vendor_key('Cognito Motorsports'))


def _is_icon_vendor_name(name):
    """Return True if vendor name looks like Icon Vehicle Dynamics."""
    key = _normalize_vendor_key(name or '')
    if key and ('iconvehicledynamics' in key or key == 'vehicledynamics'):
        return True
    canonical = normalize_vendor_name(name or '')
    canonical_key = _normalize_vendor_key(canonical or '')
    return bool(canonical_key and canonical_key == _normalize_vendor_key('Icon Vehicle Dynamics'))


def _is_river_city_vendor_name(name):
    """Return True if vendor name looks like River City Turbo."""
    key = _normalize_vendor_key(name or '')
    if key and 'rivercityturbo' in key:
        return True
    canonical = normalize_vendor_name(name or '')
    canonical_key = _normalize_vendor_key(canonical or '')
    return bool(canonical_key and canonical_key == _normalize_vendor_key('River City Turbo'))


def _is_rock_krawler_vendor_name(name):
    """Return True if vendor name looks like Rock Krawler / Pure Performance Group."""
    key = _normalize_vendor_key(name or '')
    if key and ('rockkrawler' in key or 'pureperformance' in key):
        return True
    canonical = normalize_vendor_name(name or '')
    canonical_key = _normalize_vendor_key(canonical or '')
    return bool(canonical_key and canonical_key == _normalize_vendor_key('Rock Krawler'))


def _is_sport_truck_vendor_name(name):
    """Return True if vendor name looks like Sport Truck USA / ST USA Holding Corp."""
    key = _normalize_vendor_key(name or '')
    if key and ('sporttruckusa' in key or 'stusaholdingcorp' in key):
        return True
    canonical = normalize_vendor_name(name or '')
    canonical_key = _normalize_vendor_key(canonical or '')
    return bool(canonical_key and canonical_key == _normalize_vendor_key('Sport Truck USA - $5 DS Fee'))


def _is_mishimoto_vendor_name(name):
    """Return True if vendor name looks like Mishimoto Automotive."""
    key = _normalize_vendor_key(name or '')
    if key and 'mishimotoautomotive' in key:
        return True
    canonical = normalize_vendor_name(name or '')
    canonical_key = _normalize_vendor_key(canonical or '')
    return bool(canonical_key and 'mishimotoautomotive' in canonical_key)


def _is_pt_vendor_name(name):
    """Return True if vendor name looks like Performance Turbochargers."""
    key = _normalize_vendor_key(name or '')
    if key and 'performanceturbochargers' in key:
        return True
    canonical = normalize_vendor_name(name or '')
    canonical_key = _normalize_vendor_key(canonical or '')
    return bool(canonical_key and 'performanceturbochargers' in canonical_key)


def _is_fumoto_vendor_name(name):
    """Return True if vendor name looks like Fumoto Engineering of America."""
    key = _normalize_vendor_key(name or '')
    if key and 'fumotoengineeringofamerica' in key:
        return True
    canonical = normalize_vendor_name(name or '')
    canonical_key = _normalize_vendor_key(canonical or '')
    return bool(canonical_key and 'fumotoengineeringofamerica' in canonical_key)


def _is_diamond_eye_vendor_name(name):
    """Return True if vendor name looks like Diamond Eye Manufacturing."""
    key = _normalize_vendor_key(name or '')
    if key and 'diamondeyemanufacturing' in key:
        return True
    canonical = normalize_vendor_name(name or '')
    canonical_key = _normalize_vendor_key(canonical or '')
    return bool(canonical_key and 'diamondeyemanufacturing' in canonical_key)


def _is_poly_vendor_name(name):
    """Return True if vendor name looks like Poly Performance."""
    key = _normalize_vendor_key(name or '')
    if key and 'polyperformance' in key:
        return True
    canonical = normalize_vendor_name(name or '')
    canonical_key = _normalize_vendor_key(canonical or '')
    return bool(canonical_key and 'polyperformance' in canonical_key)


def _is_merchant_vendor_name(name):
    """Return True if vendor name looks like Merchant Automotive."""
    key = _normalize_vendor_key(name or '')
    if key and 'merchantautomotive' in key:
        return True
    canonical = normalize_vendor_name(name or '')
    canonical_key = _normalize_vendor_key(canonical or '')
    return bool(canonical_key and 'merchantautomotive' in canonical_key)


def _is_dynomite_vendor_name(name):
    """Return True if vendor name looks like Dynomite Diesel."""
    key = _normalize_vendor_key(name or '')
    if key and 'dynomitediesel' in key:
        return True
    canonical = normalize_vendor_name(name or '')
    canonical_key = _normalize_vendor_key(canonical or '')
    return bool(canonical_key and 'dynomitediesel' in canonical_key)


def _is_kc_turbos_vendor_name(name):
    """Return True if vendor name looks like KC Turbos."""
    key = _normalize_vendor_key(name or '')
    if key and 'kcturbos' in key:
        return True
    canonical = normalize_vendor_name(name or '')
    canonical_key = _normalize_vendor_key(canonical or '')
    return bool(canonical_key and 'kcturbos' in canonical_key)


def _is_serra_vendor_name(name):
    """Return True if vendor name looks like Serra Chrysler Dodge Ram Jeep of Traverse City."""
    key = _normalize_vendor_key(name or '')
    if key and 'serrachryslerdodgeramjeepoftraversecity' in key:
        return True
    canonical = normalize_vendor_name(name or '')
    canonical_key = _normalize_vendor_key(canonical or '')
    return bool(canonical_key and 'serrachryslerdodgeramjeepoftraversecity' in canonical_key)


def _is_suspensionmaxx_vendor_name(name):
    """Return True if vendor name looks like SuspensionMAXX."""
    key = _normalize_vendor_key(name or '')
    if key and 'suspensionmaxx' in key:
        return True
    canonical = normalize_vendor_name(name or '')
    canonical_key = _normalize_vendor_key(canonical or '')
    return bool(canonical_key and 'suspensionmaxx' in canonical_key)


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


def _expand_multiline_row_pd(row, col_map):
    """Power Distributing-specific handling for combined product/description cells.

    pdfplumber often extracts PD rows as:
      product cell: "SKU\\nDESCRIPTION"
      qty/price/amount cells: shifted one column left from the header.
    Treat that as one item row instead of splitting into two pseudo-items.
    """
    product_idx = col_map.get('item_number')
    if product_idx is None or product_idx >= len(row):
        return _expand_multiline_row(row, col_map)

    product_lines = _split_cell_lines(row[product_idx])
    if len(product_lines) < 2:
        return _expand_multiline_row(row, col_map)

    quantity = ''
    quantity_key = 'ship_quantity' if 'ship_quantity' in col_map else 'quantity'
    if quantity_key in col_map:
        quantity_val = _find_nearby_value(
            row,
            col_map[quantity_key],
            predicate=lambda v: re.match(r'^\d+(\.\d+)?$', str(v).strip()),
            exclude_cols={col_map.get('amount'), col_map.get('unit_price')},
        )
        quantity = _clean_cell(quantity_val)

    unit_price = ''
    if 'unit_price' in col_map:
        unit_price_val = _find_nearby_value(
            row,
            col_map['unit_price'],
            predicate=lambda v: bool(_clean_price(v)),
        )
        unit_price = _clean_price(unit_price_val)

    amount = ''
    if 'amount' in col_map:
        amount_val = _find_nearby_value(
            row,
            col_map['amount'],
            predicate=lambda v: bool(_clean_price(v)),
        )
        amount = _clean_price(amount_val)

    units = 'Each'
    if 'units' in col_map:
        units_val = _find_nearby_value(
            row,
            col_map['units'],
            predicate=lambda v: bool(re.match(r'^[A-Za-z]+$', str(v).strip())),
        )
        units = _clean_cell(units_val) or 'Each'

    item = {
        'item_number': _clean_cell(product_lines[0]),
        'quantity': quantity,
        'unit_price': unit_price,
        'amount': amount,
        'description': _clean_cell(' '.join(product_lines[1:])),
        'units': units,
    }
    return [mark_freight_item(item)]


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


def extract_item_from_table_row(row, col_map, prefer_shipped_qty=False):
    """Extract a line item dict from a table row using the column map."""
    item = {}
    quantity_key = 'ship_quantity' if prefer_shipped_qty and 'ship_quantity' in col_map else 'quantity'

    item['item_number'] = _clean_cell(row[col_map['item_number']]) if 'item_number' in col_map and col_map['item_number'] < len(row) else ''
    item['quantity'] = _clean_cell(row[col_map[quantity_key]]) if quantity_key in col_map and col_map[quantity_key] < len(row) else ''
    item['unit_price'] = _clean_price(row[col_map['unit_price']]) if 'unit_price' in col_map and col_map['unit_price'] < len(row) else ''
    item['amount'] = _clean_price(row[col_map['amount']]) if 'amount' in col_map and col_map['amount'] < len(row) else ''
    item['description'] = _clean_cell(row[col_map['description']]) if 'description' in col_map and col_map['description'] < len(row) else ''
    item['units'] = _clean_cell(row[col_map['units']]) if 'units' in col_map and col_map['units'] < len(row) else 'Each'

    # PD tables can be shifted; look near the mapped column for values
    if not item['quantity'] and quantity_key in col_map:
        exclude_cols = {col_map.get('amount'), col_map.get('unit_price')}
        val = _find_nearby_value(
            row,
            col_map[quantity_key],
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


def extract_items_from_tables(filepath, sb_mode=False, pd_mode=False):
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
                if pd_mode:
                    row_items = _expand_multiline_row_pd(row, col_map)
                elif sb_mode:
                    row_items = _expand_multiline_row_sb(row, col_map)
                else:
                    row_items = _expand_multiline_row(row, col_map)
            else:
                item = extract_item_from_table_row(row, col_map, prefer_shipped_qty=pd_mode)
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


def _customer_name_from_ship_to_lines(lines):
    """Return the first plausible customer name from ship-to lines."""
    for raw_line in lines or []:
        line = str(raw_line or '').strip()
        if not line:
            continue
        if line.lower() in (
            'us',
            'usa',
            'united states',
            'united states of america',
            'ship to',
            'bill to',
            'remit to',
        ):
            continue
        if _OUR_COMPANY_NAMES.search(line):
            continue
        if any(p.search(line) for p in _OUR_ADDRESS_PATTERNS):
            continue
        if re.match(r'^\d', line):
            continue
        if re.match(r'^(suite|ste\.?|building|bldg\.?|unit|apt)\b', line, re.IGNORECASE):
            continue
        if re.match(r'^[A-Za-z][A-Za-z\s\.\'-]+[A-Z]{2}\s+\d{5}(?:-\d{4})?$', line):
            continue
        return line
    return ''


def _clean_ship_to_contact_name(value):
    """Strip trailing vendor metadata from a ship-to contact line."""
    text = re.sub(r'\s+', ' ', str(value or '')).strip(' ,')
    if not text:
        return ''
    text = re.sub(r'\s+DATE\s+\d{1,2}/\d{1,2}/\d{2,4}\b.*$', '', text, flags=re.IGNORECASE).strip()
    text = re.sub(
        r'\s+\d{4,}\s+\d{1,2}/\d{1,2}/\d{2,4}\s+SO[\w-]+\b.*$',
        '',
        text,
        flags=re.IGNORECASE,
    ).strip()
    text = re.sub(r'\s*-\([A-Za-z]+\)', '', text).strip()
    text = re.sub(r'\s+\d{3}[-.\s]\d{3}[-.\s]\d{4}\b.*$', '', text).strip()
    text = re.sub(r'\s{2,}', ' ', text).strip(' ,')
    return text


def _ship_to_our_address_from_lines(lines):
    """Return True when ship-to lines clearly match our warehouse address."""
    block = '\n'.join(str(line or '').strip() for line in lines or [])
    matches = sum(1 for p in _OUR_ADDRESS_PATTERNS if p.search(block))
    return matches >= 2


def _will_call_customer_from_lines(lines):
    """Return the will-call customer name from ship-to lines, if present."""
    return _customer_name_from_ship_to_lines(lines)


def _extract_side_by_side_ship_to_lines(filepath, header_pattern, stop_pattern, *, right_boundary_token=None):
    """Extract the right-hand Ship To column from a side-by-side header block."""
    words, page_width = _extract_first_page_words(filepath)
    if not words:
        return []

    lines = _group_words_into_lines(words)
    header_words = None
    stop_top = None
    for line_words in lines:
        line_text = _words_to_line_text(line_words)
        if header_words is None and re.search(header_pattern, line_text, re.IGNORECASE):
            header_words = line_words
            continue
        if header_words and re.search(stop_pattern, line_text, re.IGNORECASE):
            stop_top = float(line_words[0].get('top', 0))
            break

    if not header_words:
        return []

    header_top = float(header_words[0].get('top', 0))
    ship_x = min(
        (float(word.get('x0', 0)) for word in header_words if str(word.get('text', '')).lower() == 'ship'),
        default=0,
    )
    ship_x = max(0, ship_x - 28)
    max_x = page_width or 1000
    if right_boundary_token:
        boundary_x = min(
            (
                float(word.get('x0', 0))
                for word in header_words
                if str(word.get('text', '')).lower() == right_boundary_token.lower()
            ),
            default=max_x,
        )
        if boundary_x > ship_x:
            max_x = boundary_x
    if stop_top is None:
        stop_top = max(float(word.get('top', 0)) for word in words) + 1

    return _extract_column_lines_from_words(
        words,
        min_x=ship_x,
        max_x=max_x,
        min_top=header_top,
        max_top=stop_top,
    )


def _extract_power_stroke_ship_to_lines(filepath):
    """Extract Power Stroke Products ship-to lines."""
    return _extract_side_by_side_ship_to_lines(
        filepath,
        header_pattern=r'\bBill\s+to\s+Ship\s+to\b',
        stop_pattern=r'\bInvoice\s+details\b',
    )


def _extract_redhead_ship_to_lines(filepath):
    """Extract Red-Head's right-hand Ship To column from the Bill To / Ship To block."""
    return _extract_side_by_side_ship_to_lines(
        filepath,
        header_pattern=r'\bBill\s+To\s+Ship\s+To\b',
        stop_pattern=r'\bP\.?O\.?\s+No\.?\s+Terms\s+Ship\s+Via\b|\bItem\s+Description\b',
    )


def _extract_hamilton_ship_to_lines(filepath):
    """Extract Hamilton Cams ship-to lines."""
    return _extract_side_by_side_ship_to_lines(
        filepath,
        header_pattern=r'\bBill\s+To\s+Ship\s+To\b',
        stop_pattern=r'\bTracking\b|\bQuantity\b',
    )


def _extract_daystar_ship_to_lines(filepath):
    """Extract Daystar's left-hand Ship To column without Bill To spillover."""
    return _extract_side_by_side_ship_to_lines(
        filepath,
        header_pattern=r'\bShip\s+To\s+Bill\s+To\b',
        stop_pattern=r'\bTracking#?\b|\bItem\s+Customer\b',
        right_boundary_token='bill',
    )


def _extract_diesel_forward_ship_to_lines(filepath):
    """Extract Diesel Forward ship-to lines."""
    return _extract_side_by_side_ship_to_lines(
        filepath,
        header_pattern=r'\bBill\s+To\s+Ship\s+To\b',
        stop_pattern=r'\bGST/PST\b|\bTerms\b|\bTracking\b|\bQuantity\b',
    )


def _extract_carli_ship_to_lines(filepath):
    """Extract Carli Suspension ship-to lines."""
    return _extract_side_by_side_ship_to_lines(
        filepath,
        header_pattern=r'\bBill\s+To\s*:?\s+Ship\s+To\s*:?\b',
        stop_pattern=r'\bItem\s+Number\b',
    )


def _extract_icon_cognito_ship_to_lines(filepath):
    """Extract Icon / Cognito ship-to lines."""
    return _extract_side_by_side_ship_to_lines(
        filepath,
        header_pattern=r'\bBill\s+To\s+Ship\s+To\b',
        stop_pattern=r'\bTerms\b|\bTracking\b|\bQuantity\b',
    )


def _extract_redhead_ship_to_lines(filepath):
    """Extract Red Head's right-hand Ship To column without bill-to spillover."""
    return _extract_side_by_side_ship_to_lines(
        filepath,
        header_pattern=r'\bBill\s+To\s+Ship\s+To\b',
        stop_pattern=r'\bP\.O\.\s+No\.\s+Terms\s+Ship\s+Via\b',
    )


def _extract_bosch_ship_to_lines(filepath):
    """Extract Bosch ship-to lines from the Bill To / Ship To / Remit To header."""
    return _extract_side_by_side_ship_to_lines(
        filepath,
        header_pattern=r'\bBill\s+To\s+Ship\s+To\s+Remit\s+To\b',
        stop_pattern=r'\bCarrier\b|\bIncoterm\b|\bTerms\b',
        right_boundary_token='remit',
    )


def _extract_mishimoto_ship_to_lines(filepath):
    """Extract the Ship To column from Mishimoto's side-by-side Bill To / Ship To layout."""
    words, page_width = _extract_first_page_words(filepath)
    if not words:
        return []

    lines = _group_words_into_lines(words)
    header_words = None
    stop_top = 0
    for line_words in lines:
        line_text = _words_to_line_text(line_words)
        if re.search(r'\bBill\s+To\s+Ship\s+To\b', line_text, re.IGNORECASE):
            header_words = line_words
            continue
        if header_words and re.search(r'\bTerms\s+Due\s+Date\b', line_text, re.IGNORECASE):
            stop_top = float(line_words[0].get('top', 0))
            break

    if not header_words:
        return []

    header_top = float(header_words[0].get('top', 0))
    ship_x = min(
        (float(word.get('x0', 0)) for word in header_words if str(word.get('text', '')).lower() == 'ship'),
        default=0,
    )
    total_x = min(
        (float(word.get('x0', 0)) for word in header_words if str(word.get('text', '')).lower() == 'total'),
        default=page_width or 1000,
    )
    if stop_top <= 0:
        stop_top = max(float(word.get('top', 0)) for word in words) + 1

    return _extract_column_lines_from_words(
        words,
        min_x=ship_x,
        max_x=total_x,
        min_top=header_top,
        max_top=stop_top,
    )


def _extract_pt_ship_to_lines(filepath):
    """Extract the right-hand ship-to column from PT's side-by-side address block."""
    words, _page_width = _extract_first_page_words(filepath)
    if not words:
        return []

    lines = _group_words_into_lines(words)
    charge_top = None
    stop_top = None
    for line_words in lines:
        line_text = _words_to_line_text(line_words)
        if charge_top is None and 'C H A R G E' in line_text:
            charge_top = float(line_words[0].get('top', 0))
            continue
        if charge_top is not None and 'Part Number' in line_text and 'Description' in line_text:
            stop_top = float(line_words[0].get('top', 0))
            break

    if charge_top is None or stop_top is None:
        return []

    return _extract_column_lines_from_words(
        words,
        min_x=290,
        max_x=530,
        min_top=charge_top,
        max_top=stop_top,
    )


def _extract_ma_kt_ship_to_lines(filepath):
    """Extract the Ship To column from MA/KT's Bill To | Ship To | Total layout."""
    words, page_width = _extract_first_page_words(filepath)
    if not words:
        return []

    lines = _group_words_into_lines(words)
    header_words = None
    stop_top = None
    for line_words in lines:
        line_text = _words_to_line_text(line_words)
        if re.search(r'\bBill\s+To\s+Ship\s+To\s+TOTAL\b', line_text, re.IGNORECASE):
            header_words = line_words
            continue
        if header_words and re.search(r'\bTerms\s+Due\s+Date\s+PO\s*#\b', line_text, re.IGNORECASE):
            stop_top = float(line_words[0].get('top', 0))
            break

    if not header_words:
        return []

    header_top = float(header_words[0].get('top', 0))
    ship_x = min(
        (float(word.get('x0', 0)) for word in header_words if str(word.get('text', '')).lower() == 'ship'),
        default=0,
    )
    total_x = min(
        (float(word.get('x0', 0)) for word in header_words if str(word.get('text', '')).lower() == 'total'),
        default=page_width or 1000,
    )
    if stop_top is None:
        stop_top = max(float(word.get('top', 0)) for word in words) + 1

    return _extract_column_lines_from_words(
        words,
        min_x=ship_x,
        max_x=total_x,
        min_top=header_top,
        max_top=stop_top,
    )


def _extract_dd_ship_to_lines(filepath):
    """Extract the Ship To column from DD's Bill To | Ship To | Invoice layout."""
    words, page_width = _extract_first_page_words(filepath)
    if not words:
        return []

    lines = _group_words_into_lines(words)
    header_words = None
    stop_top = None
    for line_words in lines:
        line_text = _words_to_line_text(line_words)
        if re.search(r'\bBILL\s+TO\s+SHIP\s+TO\s+INVOICE\b', line_text, re.IGNORECASE):
            header_words = line_words
            continue
        if header_words and re.search(r'\bP\.O\.\s+NUMBER\b|\bSKU\s+DESCRIPTION\b', line_text, re.IGNORECASE):
            stop_top = float(line_words[0].get('top', 0))
            break

    if not header_words:
        return []

    header_top = float(header_words[0].get('top', 0))
    ship_x = min(
        (float(word.get('x0', 0)) for word in header_words if str(word.get('text', '')).lower() == 'ship'),
        default=0,
    )
    invoice_x = min(
        (float(word.get('x0', 0)) for word in header_words if str(word.get('text', '')).lower() == 'invoice'),
        default=page_width or 1000,
    )
    if stop_top is None:
        stop_top = max(float(word.get('top', 0)) for word in words) + 1

    return _extract_column_lines_from_words(
        words,
        min_x=ship_x,
        max_x=invoice_x,
        min_top=header_top,
        max_top=stop_top,
    )


def _extract_sm_ship_to_lines(filepath):
    """Extract the Ship To column from SuspensionMAXX's two-column address block."""
    words, page_width = _extract_first_page_words(filepath)
    if not words:
        return []

    lines = _group_words_into_lines(words)
    header_words = None
    stop_top = None
    for line_words in lines:
        line_text = _words_to_line_text(line_words)
        if re.search(r'\bBill\s+to\s+Ship\s+to\b', line_text, re.IGNORECASE):
            header_words = line_words
            continue
        if header_words and re.search(r'\bShipping\s+info\s+Invoice\s+details\b', line_text, re.IGNORECASE):
            stop_top = float(line_words[0].get('top', 0))
            break

    if not header_words:
        return []

    header_top = float(header_words[0].get('top', 0))
    ship_x = min(
        (float(word.get('x0', 0)) for word in header_words if str(word.get('text', '')).lower() == 'ship'),
        default=page_width / 2 if page_width else 250,
    )
    if stop_top is None:
        stop_top = max(float(word.get('top', 0)) for word in words) + 1

    return _extract_column_lines_from_words(
        words,
        min_x=ship_x,
        max_x=page_width or 1000,
        min_top=header_top,
        max_top=stop_top,
    )


def _extract_ats_ship_to_lines(filepath):
    """Extract the ATS Ship To column from the Bill To | Ship To | PO layout."""
    words, page_width = _extract_first_page_words(filepath)
    if not words:
        return []

    lines = _group_words_into_lines(words)
    header_words = None
    stop_top = None
    for line_words in lines:
        line_text = _words_to_line_text(line_words)
        if (
            'Bill To' in line_text
            and 'Ship To' in line_text
            and 'PO #' in line_text
            and 'Due Date' in line_text
        ):
            header_words = line_words
            continue
        if header_words and 'Qty' in line_text and 'Part Number' in line_text:
            stop_top = float(line_words[0].get('top', 0))
            break

    if not header_words:
        return []

    header_top = float(header_words[0].get('top', 0))
    ship_x = min(
        (float(word.get('x0', 0)) for word in header_words if str(word.get('text', '')).lower() == 'ship'),
        default=0,
    )
    po_x = min(
        (float(word.get('x0', 0)) for word in header_words if str(word.get('text', '')).lower() == 'po'),
        default=page_width or 1000,
    )
    if stop_top is None:
        stop_top = max(float(word.get('top', 0)) for word in words) + 1

    return _extract_column_lines_from_words(
        words,
        min_x=ship_x,
        max_x=po_x,
        min_top=header_top,
        max_top=stop_top,
    )


def _extract_isspro_ship_to_lines(filepath):
    """Extract ISSPRO's right-hand ship-to address block from layout text."""
    layout_text = extract_layout_text_from_pdf(filepath)
    if not layout_text:
        return []

    lines = layout_text.splitlines()
    start_idx = None
    stop_idx = None
    for idx, raw_line in enumerate(lines):
        line = str(raw_line or '')
        if start_idx is None and line.upper().count('DIESEL POWER PRODUCTS') >= 2:
            start_idx = idx
            continue
        if start_idx is not None and 'TOTAL PRICE' in line:
            stop_idx = idx
            break

    if start_idx is None:
        return []
    if stop_idx is None:
        stop_idx = len(lines)

    ship_to_lines = []
    for raw_line in lines[start_idx:stop_idx]:
        stripped = str(raw_line or '').rstrip()
        if not stripped:
            continue
        parts = [part.strip() for part in re.split(r'\s{2,}', stripped.strip()) if part.strip()]
        if len(parts) >= 2:
            candidate = parts[-1]
        else:
            candidate = stripped.strip()
        if candidate:
            ship_to_lines.append(candidate)
    return ship_to_lines


def _extract_rock_krawler_ship_to_lines(filepath):
    """Extract the Rock Krawler Ship To column from the two-column header."""
    words, page_width = _extract_first_page_words(filepath)
    if not words:
        return []

    lines = _group_words_into_lines(words)
    header_words = None
    stop_top = None
    for line_words in lines:
        line_text = _words_to_line_text(line_words)
        if re.search(r'\bBILL\s+TO\s+SHIP\s+TO\s+INVOICE\b', line_text, re.IGNORECASE):
            header_words = line_words
            continue
        if header_words and re.search(r'\bSHIP\s+DATE\b', line_text, re.IGNORECASE):
            stop_top = float(line_words[0].get('top', 0))
            break

    if not header_words:
        return []

    header_top = float(header_words[0].get('top', 0))
    ship_x = min(
        (float(word.get('x0', 0)) for word in header_words if str(word.get('text', '')).lower() == 'ship'),
        default=0,
    )
    invoice_x = min(
        (float(word.get('x0', 0)) for word in header_words if str(word.get('text', '')).lower() == 'invoice'),
        default=page_width or 1000,
    )
    if stop_top is None:
        stop_top = max(float(word.get('top', 0)) for word in words) + 1

    return _extract_column_lines_from_words(
        words,
        min_x=ship_x,
        max_x=invoice_x,
        min_top=header_top,
        max_top=stop_top,
    )


def _extract_rock_krawler_po_number(filepath):
    """Extract the Customer PO from Rock Krawler's shipping header block."""
    words, page_width = _extract_first_page_words(filepath)
    if not words:
        return ''

    lines = _group_words_into_lines(words)
    header_words = None
    stop_top = None
    for line_words in lines:
        line_text = _words_to_line_text(line_words)
        if 'CUSTOMER PO' in line_text and 'SHIP DATE' in line_text:
            header_words = line_words
            continue
        if header_words and 'ACTIVITY' in line_text:
            stop_top = float(line_words[0].get('top', 0))
            break

    if not header_words:
        return ''

    header_top = float(header_words[0].get('top', 0))
    customer_x = min(
        (float(word.get('x0', 0)) for word in header_words if str(word.get('text', '')).lower() == 'customer'),
        default=0,
    )
    if stop_top is None:
        stop_top = max(float(word.get('top', 0)) for word in words) + 1

    po_lines = _extract_column_lines_from_words(
        words,
        min_x=customer_x,
        max_x=page_width or 1000,
        min_top=header_top,
        max_top=stop_top,
    )
    return str(po_lines[0]).strip() if po_lines else ''


def _extract_sport_truck_ship_to_lines(filepath):
    """Extract the Sport Truck Ship To column from the Bill To | Ship To layout."""
    words, page_width = _extract_first_page_words(filepath)
    if not words:
        return []

    lines = _group_words_into_lines(words)
    header_words = None
    stop_top = None
    for line_words in lines:
        line_text = _words_to_line_text(line_words)
        if re.search(r'\bBill\s+To\s+Ship\s+To\b', line_text, re.IGNORECASE):
            header_words = line_words
            continue
        if header_words and re.search(r'\bCustomerNumber\b|\bPO\s+Number\b', line_text, re.IGNORECASE):
            stop_top = float(line_words[0].get('top', 0))
            break

    if not header_words:
        return []

    header_top = float(header_words[0].get('top', 0))
    ship_x = min(
        (float(word.get('x0', 0)) for word in header_words if str(word.get('text', '')).lower() == 'ship'),
        default=0,
    )
    if stop_top is None:
        stop_top = max(float(word.get('top', 0)) for word in words) + 1

    return _extract_column_lines_from_words(
        words,
        min_x=ship_x,
        max_x=(page_width or 1000) + 1,
        min_top=header_top,
        max_top=stop_top,
    )


def _clean_signed_price_token(value):
    """Normalize a money token while preserving a leading minus sign."""
    if value is None:
        return ''
    text = str(value).strip()
    if not text:
        return ''
    text = (
        text.replace('\u2212', '-')
        .replace('−', '-')
        .replace('–', '-')
        .replace('?', '')
        .replace('$', '')
        .replace(',', '')
        .strip()
    )
    if re.fullmatch(r'-?\d+(?:\.\d{2})?', text):
        return text
    return ''


def _extract_sm_items_from_words(filepath):
    """Parse SuspensionMAXX line items from first-page coordinates."""
    words, _page_width = _extract_first_page_words(filepath)
    if not words:
        return []

    lines = _group_words_into_lines(words)
    in_table = False
    items = []
    current = None

    for line_words in lines:
        line_text = _words_to_line_text(line_words)
        if not in_table:
            if re.search(
                r'#\s+Product\s+or\s+service\s+Description\s+Qty\s+Rate\s+Amount',
                line_text,
                re.IGNORECASE,
            ):
                in_table = True
            continue

        if re.match(r'(?i)^(Total|Note to customer)\b', line_text):
            break

        index_text = _words_to_line_text(
            [word for word in line_words if float(word.get('x0', 0)) < 40]
        )
        item_number = _words_to_line_text(
            [word for word in line_words if 40 <= float(word.get('x0', 0)) < 150]
        )
        description = _words_to_line_text(
            [word for word in line_words if 200 <= float(word.get('x0', 0)) < 420]
        )
        quantity = _words_to_line_text(
            [word for word in line_words if 420 <= float(word.get('x0', 0)) < 455]
        )
        unit_price = _clean_signed_price_token(
            _words_to_line_text(
                [word for word in line_words if 470 <= float(word.get('x0', 0)) < 515]
            )
        )
        amount = _clean_signed_price_token(
            _words_to_line_text(
                [word for word in line_words if 530 <= float(word.get('x0', 0)) < 575]
            )
        )

        starts_item = bool(re.match(r'^\d+\.$', index_text)) and bool(
            item_number or description or quantity or unit_price or amount
        )
        has_numeric_tail = bool(quantity or unit_price or amount)

        if starts_item and has_numeric_tail:
            if current:
                items.append(mark_freight_item(current))
            current = {
                'item_number': item_number,
                'quantity': _normalize_qty(quantity),
                'units': 'Each',
                'description': description,
                'unit_price': unit_price,
                'amount': amount,
            }
            desc_lower = str(description).lower()
            if 'discount' in desc_lower:
                current['is_discount'] = True
            continue

        if current and description and not has_numeric_tail and not item_number:
            current['description'] = f"{current.get('description', '')} {description}".strip()

    if current:
        items.append(mark_freight_item(current))

    cleaned_items = []
    for item in items:
        if not item.get('amount') and not item.get('description'):
            continue
        cleaned_items.append(item)
    return cleaned_items


def _extract_dd_items_from_words(filepath):
    """Parse Dynomite Diesel line items from the SKU/DESCRIPTION table."""
    words, _page_width = _extract_first_page_words(filepath)
    if not words:
        return []

    lines = _group_words_into_lines(words)
    in_table = False
    items = []
    current = None

    for line_words in lines:
        line_text = _words_to_line_text(line_words)
        if not in_table:
            if re.search(r'\bSKU\s+DESCRIPTION\s+QTY\s+RATE\s+AMOUNT\b', line_text, re.IGNORECASE):
                in_table = True
            continue

        if re.search(r'(?i)\b(SUBTOTAL|TAX|TOTAL|BALANCE DUE)\b', line_text):
            break

        sku = _words_to_line_text(
            [word for word in line_words if 15 <= float(word.get('x0', 0)) < 120]
        )
        description = _words_to_line_text(
            [word for word in line_words if 120 <= float(word.get('x0', 0)) < 475]
        )
        quantity = _words_to_line_text(
            [word for word in line_words if 475 <= float(word.get('x0', 0)) < 510]
        )
        unit_price = _clean_signed_price_token(
            _words_to_line_text(
                [word for word in line_words if 510 <= float(word.get('x0', 0)) < 552]
            )
        )
        amount = _clean_signed_price_token(
            _words_to_line_text(
                [word for word in line_words if 552 <= float(word.get('x0', 0)) < 600]
            )
        )

        starts_item = bool(sku) and bool(quantity or unit_price or amount)
        if starts_item:
            if current:
                items.append(current)
            current = {
                'item_number': sku,
                'quantity': _normalize_qty(quantity),
                'units': 'Each',
                'description': description,
                'unit_price': unit_price,
                'amount': amount,
            }
            continue

        if current and description and not sku and not quantity and not unit_price and not amount:
            current['description'] = f"{current.get('description', '')} {description}".strip()

    if current:
        items.append(current)

    return [item for item in items if item.get('description') or item.get('amount')]


def _extract_poly_ship_to_lines(filepath):
    """Extract Poly's Ship To column from either Ship/Bill or Bill/Ship layouts."""
    words, page_width = _extract_first_page_words(filepath)
    if not words:
        return []

    lines = _group_words_into_lines(words)
    header_words = None
    stop_top = None

    for line_words in lines:
        line_text = _words_to_line_text(line_words)
        if header_words is None and re.search(
            r'\b(?:Bill\s+To\s+Ship\s+To|Ship\s+To\s+Bill\s+To)\b',
            line_text,
            re.IGNORECASE,
        ):
            header_words = line_words
            continue
        if header_words and re.search(
            r'(?:\bTracking\b|\bNotes?\b|Shipping\s+Method\b|Item\b.*Quantity\b|Item\b.*Description\b)',
            line_text,
            re.IGNORECASE,
        ):
            stop_top = float(line_words[0].get('top', 0))
            break

    if not header_words:
        return []

    header_top = float(header_words[0].get('top', 0))
    ship_x = min(
        (float(word.get('x0', 0)) for word in header_words if str(word.get('text', '')).lower() == 'ship'),
        default=0,
    )
    bill_x = min(
        (float(word.get('x0', 0)) for word in header_words if str(word.get('text', '')).lower() == 'bill'),
        default=page_width or 1000,
    )

    if stop_top is None:
        stop_top = max(float(word.get('top', 0)) for word in words) + 1

    if ship_x and bill_x and ship_x < bill_x:
        min_x = max(0, ship_x - 24)
        max_x = max(min_x + 40, bill_x - 18)
    else:
        min_x = max(0, ship_x - 24)
        max_x = page_width or 1000

    return _extract_column_lines_from_words(
        words,
        min_x=min_x,
        max_x=max_x,
        min_top=header_top,
        max_top=stop_top,
    )


def _extract_poly_items_from_words(filepath):
    """Parse Poly line items from the Item / Quantity / Description table."""
    words, _page_width = _extract_first_page_words(filepath)
    if not words:
        return []

    lines = _group_words_into_lines(words)
    in_table = False
    items = []
    current = None

    for line_words in lines:
        line_text = _words_to_line_text(line_words)
        if not in_table:
            if re.search(r'\bItem\s+Quantity\s+Description\s+Unit\s+Price\s+Amount\b', line_text, re.IGNORECASE):
                in_table = True
            continue

        if re.search(r'(?i)\b(Subtotal|Shipping\s+Cost|Total|Amount\s+Due)\b', line_text):
            break

        item_number = _words_to_line_text(
            [word for word in line_words if 15 <= float(word.get('x0', 0)) < 105]
        )
        quantity = _words_to_line_text(
            [word for word in line_words if 150 <= float(word.get('x0', 0)) < 180]
        )
        description = _words_to_line_text(
            [word for word in line_words if 180 <= float(word.get('x0', 0)) < 450]
        )
        unit_price = _clean_signed_price_token(
            _words_to_line_text(
                [word for word in line_words if 450 <= float(word.get('x0', 0)) < 535]
            )
        )
        amount = _clean_signed_price_token(
            _words_to_line_text(
                [word for word in line_words if 535 <= float(word.get('x0', 0)) < 600]
            )
        )

        starts_item = bool(item_number) and bool(quantity) and bool(unit_price or amount)
        if starts_item:
            if current:
                items.append(current)
            current = {
                'item_number': item_number,
                'quantity': _normalize_qty(quantity),
                'units': 'Each',
                'description': description,
                'unit_price': unit_price,
                'amount': amount,
            }
            continue

        if current and description and not item_number and not quantity and not unit_price and not amount:
            current['description'] = f"{current.get('description', '')} {description}".strip()

    if current:
        items.append(current)

    return [item for item in items if item.get('description') or item.get('amount')]


def _split_item_token_and_description(text):
    """Split a combined item-code/description string into separate pieces."""
    raw = str(text or '').strip()
    if not raw:
        return '', ''

    def _looks_like_part_number(token):
        token = str(token or '').strip()
        if not token:
            return False
        if re.search(r'\d', token):
            return True
        if re.fullmatch(r'[A-Z]{2,}[A-Z0-9]*(?:-[A-Z0-9]+)+', token):
            return True
        return False

    if re.match(r'^SKU:\s*', raw, re.IGNORECASE):
        return '', raw
    parts = raw.split(None, 1)
    if len(parts) == 1:
        return (parts[0], '') if _looks_like_part_number(parts[0]) else ('', parts[0])
    first, rest = parts[0].strip(), parts[1].strip()
    if _looks_like_part_number(first):
        return first, rest
    return '', raw


def _extract_ma_kt_items_from_words(filepath):
    """Parse Merchant Automotive / KC Turbos / Mishimoto line items from the Item table."""
    words, _page_width = _extract_first_page_words(filepath)
    if not words:
        return []

    lines = _group_words_into_lines(words)
    in_table = False
    items = []
    current = None

    for line_words in lines:
        line_text = _words_to_line_text(line_words)
        if not in_table:
            if re.search(r'\b(?:QTY|Quantity)\s+Item\s+Rate\s+Amount\b', line_text, re.IGNORECASE):
                in_table = True
            continue

        if re.match(r'(?i)^(Subtotal|Shipping\s+Cost|Tax\s+Total|Total)\b', line_text):
            break

        qty = _words_to_line_text(
            [word for word in line_words if 45 <= float(word.get('x0', 0)) < 95]
        )
        item_text = _words_to_line_text(
            [word for word in line_words if 100 <= float(word.get('x0', 0)) < 430]
        )
        unit_price = _clean_signed_price_token(
            _words_to_line_text(
                [word for word in line_words if 430 <= float(word.get('x0', 0)) < 520]
            )
        )
        amount = _clean_signed_price_token(
            _words_to_line_text(
                [word for word in line_words if 520 <= float(word.get('x0', 0)) < 590]
            )
        )

        starts_item = bool(qty) and bool(item_text) and bool(unit_price or amount)
        if starts_item:
            if current:
                if _looks_like_carrier_shipping(
                    f"{current.get('item_number', '')} {current.get('description', '')}"
                ):
                    current['is_freight'] = True
                    current['quantity'] = ''
                current = mark_freight_item(current)
                if current.get('is_freight') and not re.search(r'\d', current.get('item_number', '')):
                    current['item_number'] = ''
                items.append(current)
            sku, desc = _split_item_token_and_description(item_text)
            current = {
                'item_number': sku,
                'quantity': _normalize_qty(qty),
                'units': 'Each',
                'description': desc,
                'unit_price': unit_price,
                'amount': amount,
            }
            if _looks_like_carrier_shipping(f"{item_text} {desc}"):
                current['is_freight'] = True
                current['quantity'] = ''
            continue

        if current and item_text and not qty and not unit_price and not amount:
            if re.match(r'^SKU:\s*', item_text, re.IGNORECASE):
                _, sku_text = _split_item_token_and_description(item_text.replace('SKU:', '', 1).strip())
                if not current.get('item_number'):
                    current['item_number'] = sku_text or item_text.split(':', 1)[-1].strip()
                continue
            current['description'] = f"{current.get('description', '')} {item_text}".strip()
            if _looks_like_carrier_shipping(current.get('description', '')):
                current['is_freight'] = True
                current['quantity'] = ''

    if current:
        if _looks_like_carrier_shipping(
            f"{current.get('item_number', '')} {current.get('description', '')}"
        ):
            current['is_freight'] = True
            current['quantity'] = ''
        current = mark_freight_item(current)
        if current.get('is_freight') and not re.search(r'\d', current.get('item_number', '')):
            current['item_number'] = ''
        items.append(current)

    return [item for item in items if item.get('description') or item.get('amount')]


def _extract_serra_items_from_words(filepath):
    """Parse Serra's parts table so the part number lands in SKU."""
    words, _page_width = _extract_first_page_words(filepath)
    if not words:
        return []

    lines = _group_words_into_lines(words)
    in_table = False
    items = []
    current = None

    for line_words in lines:
        line_text = _words_to_line_text(line_words)
        if not in_table:
            if re.search(r'\bREQ\s+SH\s+ORD\s+BIN\s+PART\s+NUMBER\s+DESCRIPTION\b', line_text, re.IGNORECASE):
                in_table = True
            continue

        if re.search(r'(?i)\b(Parts Sale|Total Parts Sales|Net Total Parts|Total Invoice)\b', line_text):
            break

        shipped_qty = _words_to_line_text(
            [word for word in line_words if 55 <= float(word.get('x0', 0)) < 80]
        )
        part_number = _words_to_line_text(
            [word for word in line_words if 160 <= float(word.get('x0', 0)) < 230]
        )
        description = _words_to_line_text(
            [word for word in line_words if 265 <= float(word.get('x0', 0)) < 455]
        )
        unit_price = _clean_signed_price_token(
            _words_to_line_text(
                [word for word in line_words if 510 <= float(word.get('x0', 0)) < 545]
            )
        )
        amount = _clean_signed_price_token(
            _words_to_line_text(
                [word for word in line_words if 555 <= float(word.get('x0', 0)) < 590]
            )
        )

        starts_item = bool(part_number) and bool(shipped_qty) and bool(unit_price or amount)
        if starts_item:
            if current:
                items.append(current)
            current = {
                'item_number': part_number,
                'quantity': _normalize_qty(shipped_qty),
                'units': 'Each',
                'description': description,
                'unit_price': unit_price,
                'amount': amount,
            }
            continue

        if current and description and not part_number and not shipped_qty and not unit_price and not amount:
            current['description'] = f"{current.get('description', '')} {description}".strip()

    if current:
        items.append(current)

    return [item for item in items if item.get('description') or item.get('amount')]


def _extract_holley_customer_from_layout(filepath):
    """Extract Holley ship-to customer while stripping the vertical SHIP TO letters."""
    layout_text = extract_layout_text_from_pdf(filepath)
    if not layout_text:
        return ''

    lines = layout_text.splitlines()
    start_idx = None
    for idx, line in enumerate(lines):
        if 'Customer :' in line:
            start_idx = idx + 1
            break
    if start_idx is None:
        return ''

    for line in lines[start_idx:start_idx + 8]:
        if 'ALL SALES OUTRIGHT' in line:
            break
        parts = [part.strip() for part in re.split(r'\s{2,}', line.strip()) if part.strip()]
        if len(parts) < 2:
            continue
        candidate = re.sub(r'^[A-Z]\s+', '', parts[-1]).strip()
        candidate = re.sub(r'\s+', ' ', candidate).strip(' ,')
        if candidate and not re.match(r'^\d', candidate):
            if not any(p.search(candidate) for p in _OUR_ADDRESS_PATTERNS):
                return candidate
    return ''


def _extract_river_city_items_from_layout(filepath):
    """Parse River City Turbo item rows from layout text."""
    layout_text = extract_layout_text_from_pdf(filepath)
    if not layout_text:
        return []

    lines = layout_text.splitlines()
    header_idx = None
    for idx, line in enumerate(lines):
        if all(token in line for token in ('QTY', 'PART#', 'DESCRIPTION', 'PRICE', 'EXT.')):
            header_idx = idx
            break
    if header_idx is None:
        return []

    items = []
    current_item = None
    row_re = re.compile(
        r'^\s*(-?\d+\.\d{2})([A-Z0-9-]+)\s+(.+?)\s+(-?[\d,]+\.\d{2})\s+(-?[\d,]+\.\d{2})\s*$'
    )

    for raw_line in lines[header_idx + 1:]:
        stripped = raw_line.strip()
        if not stripped:
            continue
        if re.search(r'^(?:SUBTOTAL|SALES\s+TAX|GARRETT|FREIGHT|RC\s+SERIES|TOTAL)\b', stripped, re.IGNORECASE):
            break

        row_match = row_re.match(raw_line)
        if row_match:
            current_item = {
                'item_number': row_match.group(2),
                'quantity': _normalize_qty(row_match.group(1)),
                'units': 'Each',
                'description': re.sub(r'\s+', ' ', row_match.group(3)).strip(' ,'),
                'unit_price': _clean_price(row_match.group(4)),
                'amount': _clean_price(row_match.group(5)),
            }
            items.append(current_item)
            continue

        if current_item:
            current_item['description'] = (
                f"{current_item.get('description', '')} {re.sub(r'\\s+', ' ', stripped)}"
            ).strip(' ,')

    return items


def _extract_holley_items_from_layout(filepath):
    """Parse Holley item rows from layout text, including drop ship and freight rows."""
    layout_text = extract_layout_text_from_pdf(filepath)
    if not layout_text:
        return []

    lines = layout_text.splitlines()
    header_idx = None
    for idx, line in enumerate(lines):
        if 'Ln' in line and 'Item' in line and 'Extended Price' in line:
            header_idx = idx
            break
    if header_idx is None:
        return []

    items = []
    current_item = None
    item_row_re = re.compile(
        r'^\s*(\d+)\s+(\S+)\s+([A-Za-z/]+)\s+(\d+(?:\.\d+)?)\s+(\d+(?:\.\d+)?)\s+(\d+(?:\.\d+)?)\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s*$'
    )

    for raw_line in lines[header_idx + 1:]:
        line = raw_line.strip()
        if not line:
            continue
        if line.startswith('Carrier(s)'):
            break

        match = item_row_re.match(line)
        if match:
            shipped_qty = match.group(5)
            current_item = {
                'item_number': match.group(2),
                'quantity': shipped_qty if shipped_qty not in ('', '0') else match.group(4),
                'units': 'Each',
                'description': '',
                'unit_price': _clean_price(match.group(7)),
                'amount': _clean_price(match.group(8)),
            }
            items.append(current_item)
            continue

        if current_item and re.search(r'D\s*escription\s*:', line, re.IGNORECASE):
            desc = re.sub(r'^D\s*escription\s*:\s*', '', line, flags=re.IGNORECASE).strip()
            desc = re.sub(r'\bCI:\s*.*$', '', desc, flags=re.IGNORECASE).strip()
            desc = re.sub(r'\bOrigin:\s*.*$', '', desc, flags=re.IGNORECASE).strip()
            current_item['description'] = re.sub(r'\s+', ' ', desc).strip(' ,')

    for footer_label, item_number in (
        ('Drop Ship & Other', 'Drop Ship'),
        ('Freight', 'Freight'),
    ):
        match = re.search(
            rf'(?m)^\s*{re.escape(footer_label)}\s+([\d,]+\.\d{{2}})\s*$',
            layout_text,
        )
        if not match:
            continue
        item = {
            'item_number': item_number,
            'quantity': '1',
            'units': 'Each',
            'description': footer_label,
            'unit_price': _clean_price(match.group(1)),
            'amount': _clean_price(match.group(1)),
        }
        items.append(mark_freight_item(item))

    return items


def _extract_valair_items_from_layout(filepath):
    """Parse Valair item rows from layout text, skipping tracking pseudo-items."""
    layout_text = extract_layout_text_from_pdf(filepath)
    if not layout_text:
        return []

    lines = layout_text.splitlines()
    header_idx = None
    for idx, line in enumerate(lines):
        if all(token in line for token in ('Qty', 'Item', 'Description', 'Rate', 'Amount')):
            header_idx = idx
            break
    if header_idx is None:
        return []

    items = []
    current_item = None
    product_row_re = re.compile(
        r'^\s*(\d+(?:\.\d+)?)\s+(\S+)\s+(.+?)\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s*$'
    )
    tracking_row_re = re.compile(
        r'^\s*(?:\d+(?:\.\d+)?)?\s*Tracking\s+#\s+([A-Z0-9]{10,})\s+[\d,]+\.\d{2}\s+[\d,]+\.\d{2}\s*$',
        re.IGNORECASE,
    )
    freight_row_re = re.compile(
        r'^\s*(Freight\s+Charges)\s+\1\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s*$',
        re.IGNORECASE,
    )

    for raw_line in lines[header_idx + 1:]:
        stripped = raw_line.strip()
        if not stripped:
            continue
        if 'Thank you for your business.' in stripped or re.search(r'\bSubtotal\b', stripped, re.IGNORECASE):
            break

        tracking_match = tracking_row_re.match(raw_line)
        if tracking_match:
            current_item = None
            continue

        freight_match = freight_row_re.match(raw_line)
        if freight_match:
            current_item = None
            rate = _clean_price(freight_match.group(2))
            amount = _clean_price(freight_match.group(3))
            if rate not in ('', '0', '0.00') or amount not in ('', '0', '0.00'):
                item = {
                    'item_number': 'Freight Charges',
                    'quantity': '1',
                    'units': 'Each',
                    'description': 'Freight Charges',
                    'unit_price': rate,
                    'amount': amount or rate,
                }
                items.append(mark_freight_item(item))
            continue

        product_match = product_row_re.match(raw_line)
        if product_match:
            current_item = {
                'item_number': product_match.group(2),
                'quantity': product_match.group(1),
                'units': 'Each',
                'description': re.sub(r'\s+', ' ', product_match.group(3)).strip(' ,'),
                'unit_price': _clean_price(product_match.group(4)),
                'amount': _clean_price(product_match.group(5)),
            }
            items.append(current_item)
            continue

        if current_item:
            current_item['description'] = (
                f"{current_item.get('description', '')} {re.sub(r'\\s+', ' ', stripped)}"
            ).strip(' ,')

    return items


def _extract_valair_shipping_from_layout(filepath):
    """Extract Valair freight from the line-item area, including zero-dollar freight rows."""
    layout_text = extract_layout_text_from_pdf(filepath)
    if not layout_text:
        return '', ''

    freight_match = re.search(
        r'(?im)^\s*(Freight\s+Charges)\s+\1\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s*$',
        layout_text,
    )
    if not freight_match:
        return '', ''

    rate = _clean_price(freight_match.group(2))
    amount = _clean_price(freight_match.group(3))
    shipping_cost = amount or rate
    if shipping_cost == '':
        return '', ''

    return shipping_cost, 'Freight Charges'


def _extract_ats_items_from_layout(filepath):
    """Parse ATS item rows from layout text."""
    layout_text = extract_layout_text_from_pdf(filepath)
    if not layout_text:
        return []

    lines = layout_text.splitlines()
    header_idx = None
    for idx, line in enumerate(lines):
        if all(token in line for token in ('Qty', 'Part Number', 'Item Description', 'Unit Price', 'Amount')):
            header_idx = idx
            break
    if header_idx is None:
        return []

    items = []
    current_item = None
    row_re = re.compile(
        r'^\s*(\d+(?:\.\d+)?)\s+(\S+)\s+(.+?)\s+\$?([\d,]+\.\d{2})\s+\$?([\d,]+\.\d{2})\s*$'
    )

    for raw_line in lines[header_idx + 1:]:
        stripped = raw_line.strip()
        if not stripped:
            continue
        if stripped.startswith('Return Policy') or re.search(r'\bSubtotal\b', stripped, re.IGNORECASE):
            break

        row_match = row_re.match(stripped)
        if row_match:
            current_item = {
                'item_number': row_match.group(2),
                'quantity': row_match.group(1),
                'units': 'Each',
                'description': re.sub(r'\s+', ' ', row_match.group(3)).strip(' ,'),
                'unit_price': _clean_price(row_match.group(4)),
                'amount': _clean_price(row_match.group(5)),
            }
            items.append(current_item)
            continue

        if current_item:
            current_item['description'] = (
                f"{current_item.get('description', '')} {re.sub(r'\\s+', ' ', stripped)}"
            ).strip(' ,')

    return items


def _extract_isspro_items_from_layout(filepath):
    """Parse Isspro product rows from layout text, including footer discount."""
    layout_text = extract_layout_text_from_pdf(filepath)
    if not layout_text:
        return []

    lines = layout_text.splitlines()
    header_idx = None
    for idx, line in enumerate(lines):
        if 'TOTAL PRICE' in line:
            header_idx = idx
            break
    if header_idx is None:
        return []

    items = []
    current_item = None
    row_re = re.compile(
        r'^\s*\d{4}\s+(\S+)\s+(\d+(?:\.\d+)?)\s+(\d+(?:\.\d+)?)\s+(\d+(?:\.\d+)?)\s+([\d,]+\.\d{2,5})\s+([\d,]+\.\d{2})\s*$'
    )

    for raw_line in lines[header_idx + 1:]:
        stripped = raw_line.strip()
        if not stripped:
            continue
        if stripped.startswith('REMITTANCE ADDRESS') or re.search(r'\bSUBTOTAL\b', stripped, re.IGNORECASE):
            break

        row_match = row_re.match(stripped)
        if row_match:
            ordered_qty = row_match.group(2)
            shipped_qty = row_match.group(3)
            current_item = {
                'item_number': row_match.group(1),
                'quantity': shipped_qty if shipped_qty not in ('', '0', '0.0') else ordered_qty,
                'units': 'Each',
                'description': '',
                'unit_price': _clean_price(row_match.group(5)),
                'amount': _clean_price(row_match.group(6)),
            }
            items.append(current_item)
            continue

        if current_item:
            current_item['description'] = (
                f"{current_item.get('description', '')} {re.sub(r'\\s+', ' ', stripped)}"
            ).strip(' ,')

    discount_match = re.search(r'LESS\s+DISCOUNT\s+([\d,]+\.\d{2})', layout_text, re.IGNORECASE)
    if discount_match:
        discount_amount = _clean_price(discount_match.group(1))
        if discount_amount and discount_amount not in ('0', '0.00'):
            amount_text = f"-{discount_amount}" if not discount_amount.startswith('-') else discount_amount
            items.append({
                'item_number': 'DPP DISCOUNT',
                'quantity': '1',
                'units': 'Each',
                'description': 'LESS DISCOUNT',
                'unit_price': amount_text,
                'amount': amount_text,
                'is_discount': True,
            })

    return items


def _extract_rock_krawler_items_from_layout(filepath):
    """Parse Rock Krawler activity rows from layout text."""
    layout_text = extract_layout_text_from_pdf(filepath)
    if not layout_text:
        return []

    lines = layout_text.splitlines()
    header_idx = None
    for idx, line in enumerate(lines):
        if all(token in line for token in ('ACTIVITY', 'QTY', 'RATE', 'AMOUNT')):
            header_idx = idx
            break
    if header_idx is None:
        return []

    items = []
    current_item = None
    row_re = re.compile(
        r'^\s*([A-Z]{2,}[0-9A-Z-]+)\s+(\d+(?:\.\d+)?)\s+([\d,]+\.\d{1,3})\s+([\d,]+\.\d{2})\s*$'
    )

    for raw_line in lines[header_idx + 1:]:
        stripped = raw_line.strip()
        if not stripped:
            continue
        if re.search(r'\bSUBTOTAL\b', stripped, re.IGNORECASE):
            break

        row_match = row_re.match(stripped)
        if row_match:
            current_item = {
                'item_number': row_match.group(1),
                'quantity': row_match.group(2),
                'units': 'Each',
                'description': '',
                'unit_price': _clean_price(row_match.group(3)),
                'amount': _clean_price(row_match.group(4)),
            }
            items.append(current_item)
            continue

        if current_item and not re.match(
            r'^(?:Shipped On:|Total Shipment Weight:|Pack\b|Tracking\b|SHIPPING\b|TOTAL\b)',
            stripped,
            re.IGNORECASE,
        ):
            current_item['description'] = (
                f"{current_item.get('description', '')} {re.sub(r'\\s+', ' ', stripped)}"
            ).strip(' ,')

    return items


def _extract_sport_truck_items_from_layout(filepath):
    """Parse Sport Truck line items from layout text."""
    layout_text = extract_layout_text_from_pdf(filepath)
    if not layout_text:
        return []

    lines = layout_text.splitlines()
    header_idx = None
    for idx, line in enumerate(lines):
        if all(token in line for token in ('Item / Part Number', 'Unit Price', 'Surcharge')):
            header_idx = idx
            break
    if header_idx is None:
        return []

    items = []
    current_part = ''
    current_item = None
    qty_row_re = re.compile(
        r'^\s*(\d+)\s+(\d+(?:\.\d+)?)\s+(\d+(?:\.\d+)?)\s+(\d+(?:\.\d+)?)\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s*$'
    )

    for raw_line in lines[header_idx + 1:]:
        stripped = raw_line.strip()
        if not stripped:
            continue
        if stripped.startswith('Comments:') or re.search(r'\bSUB-TOTAL\b', stripped, re.IGNORECASE):
            break

        if re.fullmatch(r'[A-Z0-9-]+', stripped) and re.search(r'[A-Z]', stripped):
            current_part = stripped
            current_item = None
            continue

        qty_match = qty_row_re.match(stripped)
        if qty_match and current_part:
            ordered_qty = qty_match.group(2)
            shipped_qty = qty_match.group(3)
            current_item = {
                'item_number': current_part,
                'quantity': shipped_qty if shipped_qty not in ('', '0', '0.0') else ordered_qty,
                'units': 'Each',
                'description': '',
                'unit_price': _clean_price(qty_match.group(5)),
                'amount': _clean_price(qty_match.group(7)),
            }
            items.append(current_item)
            continue

        if current_item:
            current_item['description'] = (
                f"{current_item.get('description', '')} {re.sub(r'\\s+', ' ', stripped)}"
            ).strip(' ,')

    return items


def _extract_power_stroke_items_from_layout(filepath):
    """Parse Power Stroke Products item rows from layout text."""
    layout_text = extract_layout_text_from_pdf(filepath)
    if not layout_text:
        return []

    lines = layout_text.splitlines()
    header_idx = None
    for idx, line in enumerate(lines):
        if 'Product/service' in line and 'Amount' in line:
            header_idx = idx
            break
    if header_idx is None:
        return []

    items = []
    current_item = None
    row_re = re.compile(
        r'^\s*\d+\.\s+(\S+)\s+(.+?)\s+(\d+(?:\.\d+)?)\s+\$?([\d,]+\.\d{2})\s+\$?([\d,]+\.\d{2})\s*$'
    )

    for raw_line in lines[header_idx + 1:]:
        stripped = raw_line.strip()
        if not stripped:
            continue
        if re.match(r'^(?:Total|Ways\s+to\s+pay|View\s+and\s+pay)\b', stripped, re.IGNORECASE):
            break

        row_match = row_re.match(stripped)
        if row_match:
            current_item = {
                'item_number': row_match.group(1),
                'quantity': _normalize_qty(row_match.group(3)),
                'units': 'Each',
                'description': re.sub(r'\s+', ' ', row_match.group(2)).strip(' ,'),
                'unit_price': _clean_price(row_match.group(4)),
                'amount': _clean_price(row_match.group(5)),
            }
            items.append(current_item)
            continue

        if re.match(r'^\s*\d+\.\s+', stripped):
            current_item = None
            continue

        if current_item:
            current_item['description'] = (
                f"{current_item.get('description', '')} {re.sub(r'\\s+', ' ', stripped)}"
            ).strip(' ,')

    return items


def _extract_bosch_items_from_layout(filepath):
    """Parse Bosch item rows from layout text."""
    layout_text = extract_layout_text_from_pdf(filepath)
    if not layout_text:
        return []

    lines = layout_text.splitlines()
    header_idx = None
    for idx, line in enumerate(lines):
        if 'Item No.' in line and 'Description' in line and 'Extended' in line:
            header_idx = idx
            break
    if header_idx is None:
        return []

    items = []
    for raw_line in lines[header_idx + 1:]:
        stripped = raw_line.strip()
        if not stripped:
            continue
        if re.search(r'^(?:Product\s+Total|Warranty|Market\s+Support|Advertising\s+Discount|Subtotal|Total\s+in\s+currency)\b', stripped, re.IGNORECASE):
            break
        if not re.match(r'^\d+/\d+\b', stripped):
            continue

        tokens = stripped.split()
        if len(tokens) < 9:
            continue

        gross = _clean_price(tokens[-1])
        amount = _clean_price(tokens[-2])
        unit_price = _clean_price(tokens[-3])
        unit = tokens[-4]
        if not (gross and amount and unit_price and re.fullmatch(r'[A-Z]{2,}', unit)):
            continue

        qty_end = len(tokens) - 5
        qty_values = []
        while qty_end >= 3 and re.fullmatch(r'\d+(?:\.\d+)?', tokens[qty_end]) and len(qty_values) < 3:
            qty_values.insert(0, tokens[qty_end])
            qty_end -= 1
        if not qty_values:
            continue

        item_number = tokens[1]
        desc_start = 2
        if len(tokens) > 3 and re.fullmatch(r'[A-Z0-9]{8,}', tokens[2], re.IGNORECASE):
            desc_start = 3
        description = ' '.join(tokens[desc_start:qty_end + 1]).strip()
        if not description:
            continue

        items.append({
            'item_number': item_number,
            'quantity': _normalize_qty(qty_values[-1]),
            'units': unit,
            'description': re.sub(r'\s+', ' ', description).strip(' ,'),
            'unit_price': unit_price,
            'amount': amount,
        })

    return items


def _extract_diesel_forward_items_from_layout(filepath):
    """Parse Diesel Forward item rows from layout text."""
    layout_text = extract_layout_text_from_pdf(filepath)
    if not layout_text:
        return []

    lines = layout_text.splitlines()
    header_idx = None
    for idx, line in enumerate(lines):
        if all(token in line for token in ('Item', 'Description', 'Unit', 'Quantity', 'Unit Price', 'Total Price')):
            header_idx = idx
            break
    if header_idx is None:
        return []

    items = []
    current_item = None
    row_re = re.compile(
        r'^\s*(\S+)\s+(.+?)\s+([A-Za-z]+)\s+(\d+(?:\.\d+)?)\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s*$'
    )

    for raw_line in lines[header_idx + 1:]:
        stripped = raw_line.strip()
        if not stripped:
            continue
        if re.search(r'^(?:Subtotal:|Invoice\s+Discount:|Tax:|Total\s+USD:|US\s+customers\s+remit)', stripped, re.IGNORECASE):
            break

        row_match = row_re.match(stripped)
        if row_match:
            current_item = {
                'item_number': row_match.group(1),
                'quantity': _normalize_qty(row_match.group(4)),
                'units': row_match.group(3),
                'description': re.sub(r'\s+', ' ', row_match.group(2)).strip(' ,'),
                'unit_price': _clean_price(row_match.group(5)),
                'amount': _clean_price(row_match.group(6)),
            }
            items.append(current_item)
            continue

        if current_item:
            current_item['description'] = (
                f"{current_item.get('description', '')} {re.sub(r'\\s+', ' ', stripped)}"
            ).strip(' ,')

    return items


def _extract_beans_items_from_layout(filepath):
    """Parse Beans Diesel Performance item rows from layout text."""
    layout_text = extract_layout_text_from_pdf(filepath)
    if not layout_text:
        return []

    lines = layout_text.splitlines()
    header_idx = None
    for idx, line in enumerate(lines):
        if all(token in line for token in ('Item', 'Description', 'Invoiced', 'Rate', 'Amount')):
            header_idx = idx
            break
    if header_idx is None:
        return []

    items = []
    current_item = None
    row_re = re.compile(
        r'^\s*(\S+)\s+(.+?)\s+'
        r'(\d+(?:\.\d+)?)\s+'          # ordered qty
        r'(\d+(?:\.\d+)?)\s+'          # prev invoiced
        r'(\d+(?:\.\d+)?)\s+'          # backordered
        r'(\d+(?:\.\d+)?)\s+'          # invoiced qty
        r'(?:([A-Za-z]+)\s+)?'         # optional U/M
        r'(-?[\d,]+\.\d{2})\s+'        # rate
        r'(-?[\d,]+\.\d{2})T?\s*$',    # amount (trailing T in source)
        re.IGNORECASE,
    )

    for raw_line in lines[header_idx + 1:]:
        stripped = raw_line.strip()
        if not stripped:
            continue
        if re.search(r'^(?:Your\s+Tracking\s+Number|Subtotal|Sales\s+Tax|Total|Payments/Credits|Balance\s+Due)\b', stripped, re.IGNORECASE):
            break

        row_match = row_re.match(stripped)
        if row_match:
            current_item = {
                'item_number': row_match.group(1),
                'quantity': _normalize_qty(row_match.group(6)),
                'units': row_match.group(7) or 'Each',
                'description': re.sub(r'\s+', ' ', row_match.group(2)).strip(' ,'),
                'unit_price': _clean_price(row_match.group(8)),
                'amount': _clean_price(row_match.group(9)),
            }
            items.append(current_item)
            continue

        if current_item:
            current_item['description'] = (
                f"{current_item.get('description', '')} {re.sub(r'\\s+', ' ', stripped)}"
            ).strip(' ,')

    return items


def _extract_carli_items_from_layout(filepath):
    """Parse Carli item rows from layout text, including drop-ship fee lines."""
    layout_text = extract_layout_text_from_pdf(filepath)
    if not layout_text:
        return []

    lines = layout_text.splitlines()
    header_idx = None
    for idx, line in enumerate(lines):
        if all(token in line for token in ('Item Number', 'Description', 'Quantity', 'Price', 'Extension')):
            header_idx = idx
            break
    if header_idx is None:
        return []

    items = []
    current_item = None
    meta_row_re = re.compile(
        r'^\s*\d{5,}(?:\s+\d{1,2}/\d{1,2}/\d{4})?\s+[A-Za-z0-9-]+\s+\d{1,2}/\d{1,2}/\d{4}\b',
        re.IGNORECASE,
    )

    def _prefix_to_item_and_desc(prefix):
        prefix = re.sub(r'\s+', ' ', prefix).strip(' ,')
        if not prefix:
            return '', ''
        words = prefix.split()
        if len(words) >= 4 and words[:2] == ['DROP', 'SHIP']:
            return 'DROP SHIP', ' '.join(words[2:]).strip()
        return words[0], ' '.join(words[1:]).strip()

    def _normalize_carli_drop_ship_item(item):
        if not item:
            return item
        combined = ' '.join(
            part for part in (
                str(item.get('item_number', '')).strip(),
                str(item.get('description', '')).strip(),
            )
            if part
        )
        if not re.search(r'\bdrop\s+ship\b', combined, re.IGNORECASE):
            return item
        item['description'] = 'Drop Ship'
        return _apply_export_overrides(
            item,
            row_type='Category Details',
            category='Purchases',
            product_service='Drop Ship',
        )

    def _build_carli_item(item_number, quantity, units, description, unit_price, amount):
        item = {
            'item_number': item_number,
            'quantity': quantity,
            'units': units,
            'description': description,
            'unit_price': unit_price,
            'amount': amount,
        }
        item = mark_freight_item(item)
        return _normalize_carli_drop_ship_item(item)

    for raw_line in lines[header_idx + 1:]:
        stripped = raw_line.strip()
        if not stripped:
            continue
        if meta_row_re.match(stripped):
            continue
        if re.search(r'^(?:Pack\s+Slip\s+#|All\s+Prices\s+Are\s+Shown|Subtotal:|Threshold\s+Disc:|Tax:|Freight:|Thank\s+You|Total:)', stripped, re.IGNORECASE):
            if current_item and stripped.startswith('Pack Slip #'):
                current_item = current_item
            elif re.search(r'^(?:All\s+Prices\s+Are\s+Shown|Subtotal:|Threshold\s+Disc:|Tax:|Freight:|Thank\s+You|Total:)', stripped, re.IGNORECASE):
                break
            continue

        tokens = stripped.split()
        if len(tokens) >= 5:
            amount = _clean_price(tokens[-1])
            unit = tokens[-2]
            unit_price = _clean_price(tokens[-3])
            quantity = _normalize_qty(tokens[-4]) if re.fullmatch(r'\d+(?:\.\d+)?', tokens[-4]) else ''
            prefix = ' '.join(tokens[:-4])
            if amount and unit_price and quantity and re.fullmatch(r'[A-Za-z]+', unit):
                item_number, description = _prefix_to_item_and_desc(prefix)
                current_item = _build_carli_item(
                    item_number=item_number,
                    quantity=quantity,
                    units=unit,
                    description=description,
                    unit_price=unit_price,
                    amount=amount,
                )
                items.append(current_item)
                continue
            if (
                amount
                and unit_price
                and re.fullmatch(r'[A-Za-z]+', unit)
                and re.fullmatch(r'\d+(?:\.\d+)?', tokens[0])
                and len(tokens) >= 6
            ):
                quantity = _normalize_qty(tokens[0])
                item_tokens = tokens[1:-3]
                if item_tokens:
                    if len(item_tokens) >= 2 and item_tokens[:2] == ['DROP', 'SHIP']:
                        item_number = 'DROP SHIP'
                        description = ' '.join(item_tokens[2:]).strip()
                    else:
                        item_number = item_tokens[0]
                        description = ' '.join(item_tokens[1:]).strip()
                    current_item = _build_carli_item(
                        item_number=item_number,
                        quantity=quantity,
                        units=unit,
                        description=description,
                        unit_price=unit_price,
                        amount=amount,
                    )
                    items.append(current_item)
                    continue

        if current_item:
            current_item['description'] = (
                f"{current_item.get('description', '')} {re.sub(r'\\s+', ' ', stripped)}"
            ).strip(' ,')

    return items


def _extract_icon_cognito_items_from_layout(filepath):
    """Parse Icon / Cognito item rows from layout text."""
    layout_text = extract_layout_text_from_pdf(filepath)
    if not layout_text:
        return []

    lines = layout_text.splitlines()
    header_idx = None
    for idx, line in enumerate(lines):
        if re.search(r'^\s*Quantity\s+Item\s+Number\s+Description\s+Price\s+Extension\b', line, re.IGNORECASE):
            header_idx = idx
            break
    if header_idx is None:
        return []

    items = []
    current_item = None
    row_re = re.compile(
        r'^\s*(\d+(?:\.\d+)?)\s+(\S+)\s+(.+?)\s+([\d,]+\.\d{2})\s+([A-Za-z]+)\s+([\d,]+\.\d{2})\s*$'
    )
    meta_row_re = re.compile(r'^\s*\S+\s+(\d{4,})\s+\d{1,2}/\d{1,2}/\d{4}\s*$')

    for raw_line in lines[header_idx + 1:]:
        stripped = raw_line.strip()
        if not stripped:
            continue
        if re.search(r'^(?:All\s+Prices\s+Are\s+Shown|Subtotal:|Tax:|Freight:|Total:|Thank\s+You|Payments\s+Applied:|Balance\s+Due:)', stripped, re.IGNORECASE):
            break
        if meta_row_re.match(stripped):
            continue

        row_match = row_re.match(stripped)
        if row_match:
            current_item = {
                'item_number': row_match.group(2),
                'quantity': _normalize_qty(row_match.group(1)),
                'units': row_match.group(5),
                'description': re.sub(r'\s+', ' ', row_match.group(3)).strip(' ,'),
                'unit_price': _clean_price(row_match.group(4)),
                'amount': _clean_price(row_match.group(6)),
            }
            items.append(current_item)
            continue

        if current_item:
            current_item['description'] = (
                f"{current_item.get('description', '')} {re.sub(r'\\s+', ' ', stripped)}"
            ).strip(' ,')

    return items


def _extract_pt_items_from_words(filepath):
    """Parse Performance Turbochargers rows using first-page coordinates."""
    words, _page_width = _extract_first_page_words(filepath)
    if not words:
        return []

    lines = _group_words_into_lines(words)
    header_words = None
    thank_you_top = None
    order_x = desc_x = net_x = None

    for line_words in lines:
        line_text = _words_to_line_text(line_words)
        if header_words is None and 'Part Number' in line_text and 'Value' in line_text:
            header_words = line_words
            for word in line_words:
                text = str(word.get('text', ''))
                if text == 'Order':
                    order_x = float(word.get('x0', 0))
                elif text == 'Description':
                    desc_x = float(word.get('x0', 0))
                elif text == 'Net':
                    net_x = float(word.get('x0', 0))
            continue
        if header_words and 'Thank you for choosing Diesel USA Group' in line_text:
            thank_you_top = float(line_words[0].get('top', 0))
            break

    if header_words is None or order_x is None or desc_x is None or net_x is None:
        return []

    items = []
    header_top = float(header_words[0].get('top', 0))
    for line_words in lines:
        top = float(line_words[0].get('top', 0))
        if top <= header_top:
            continue
        if thank_you_top is not None and top >= thank_you_top:
            break

        line_text = _words_to_line_text(line_words)
        if not line_text:
            continue
        if re.match(r'^(?:\*send all invoices|COUNTRY OF ORIGIN|MOUNTING KIT|CORES MUST)\b', line_text, re.IGNORECASE):
            continue
        if re.match(r'^[A-Z0-9]{12,}$', line_text):
            continue

        price_words = [
            word for word in line_words
            if re.fullmatch(r'\d[\d,]*\.\d{2}', str(word.get('text', '')).strip())
        ]
        if len(price_words) < 2:
            continue

        if re.fullmatch(r'\d+\s+[\d,]+\.\d{2}\s+[\d,]+\.\d{2}', line_text):
            continue

        amount = _clean_price(price_words[-1].get('text', ''))
        unit_price = _clean_price(price_words[-2].get('text', ''))
        first_price_x = float(price_words[-2].get('x0', net_x))

        sku_words = [word for word in line_words if float(word.get('x0', 0)) < order_x - 5]
        qty_words = [
            word for word in line_words
            if order_x <= float(word.get('x0', 0)) < desc_x - 5
            and re.fullmatch(r'\d+', str(word.get('text', '')).strip())
        ]
        desc_words = [
            word for word in line_words
            if desc_x <= float(word.get('x0', 0)) < first_price_x - 1
        ]

        if sku_words:
            sku = _words_to_line_text(sku_words)
            qty = str(qty_words[1].get('text', '')).strip() if len(qty_words) >= 2 else (
                str(qty_words[0].get('text', '')).strip() if qty_words else '1'
            )
            description = _words_to_line_text(desc_words)
            description = re.sub(r'\s+[A-Z]?\d[\d,]*\.\d{2}$', '', description).strip()
        else:
            sku = ''
            qty = '1'
            description = _words_to_line_text(
                [word for word in line_words if float(word.get('x0', 0)) < first_price_x - 1]
            ) or line_text
            description = re.sub(r'\s+\d[\d,]*\.\d{2}$', '', description).strip()

        item = {
            'item_number': sku,
            'quantity': qty or '1',
            'units': 'Each',
            'description': description.strip(),
            'unit_price': unit_price,
            'amount': amount,
        }
        if item.get('item_number') or item.get('description'):
            items.append(mark_freight_item(item))

    return items


def _extract_fumoto_ship_to_name(filepath):
    """Extract the first ship-to line from Fumoto's dedicated SHIP TO table."""
    lines = _extract_fumoto_ship_to_lines(filepath)
    return lines[0].strip() if lines else ''


def _extract_fumoto_ship_to_lines(filepath):
    """Extract Fumoto's dedicated Ship To table lines."""
    for table in extract_tables_from_pdf(filepath):
        if not table:
            continue
        header = str(table[0][0] or '').strip().lower() if table[0] and table[0][0] is not None else ''
        if header != 'ship to':
            continue
        if len(table) < 2 or not table[1]:
            return []
        ship_block = str(table[1][0] or '').strip()
        return [line.strip() for line in ship_block.splitlines() if line.strip()]
    return []


def _extract_fumoto_items_from_tables(filepath):
    """Parse Fumoto's multiline line-item table."""
    tables = extract_tables_from_pdf(filepath)
    for table in tables:
        if not table or len(table) < 2:
            continue
        header = [str(cell or '').strip().lower() for cell in table[0]]
        if header[:6] != ['date', 'activity', 'description', 'qty', 'rate', 'amount']:
            continue

        items = []
        for row in table[1:]:
            if not row:
                continue
            activity_lines = _split_cell_lines(row[1] if len(row) > 1 else '')
            desc_lines = _split_cell_lines(row[2] if len(row) > 2 else '')
            qty_lines = _split_cell_lines(row[3] if len(row) > 3 else '')
            rate_lines = _split_cell_lines(row[4] if len(row) > 4 else '')
            amount_lines = _split_cell_lines(row[5] if len(row) > 5 else '')
            line_count = max(len(activity_lines), len(qty_lines), len(rate_lines), len(amount_lines), 1)

            for idx in range(line_count):
                activity = activity_lines[idx] if idx < len(activity_lines) else ''
                qty = qty_lines[idx] if idx < len(qty_lines) else ''
                rate = rate_lines[idx] if idx < len(rate_lines) else ''
                amount = amount_lines[idx] if idx < len(amount_lines) else ''

                if 'shipping fees' in activity.lower():
                    description = desc_lines[0] if desc_lines else activity
                    item = {
                        'item_number': 'Shipping Fees',
                        'quantity': qty or '1',
                        'units': 'Each',
                        'description': description,
                        'unit_price': _clean_price(rate),
                        'amount': _clean_price(amount),
                    }
                    items.append(mark_freight_item(item))
                    continue

                description = ''
                if len(desc_lines) == 1 and len(activity_lines) == 1:
                    description = desc_lines[0]
                elif len(desc_lines) == line_count and idx < len(desc_lines):
                    description = desc_lines[idx]

                item = {
                    'item_number': activity,
                    'quantity': qty or '1',
                    'units': 'Each',
                    'description': description or activity,
                    'unit_price': _clean_price(rate),
                    'amount': _clean_price(amount),
                }
                items.append(mark_freight_item(item))

        if items:
            return items

    return []


def _extract_diamond_eye_handling_fee(layout_text):
    """Parse Diamond Eye's HDL handling fee row as a drop-ship purchase line."""
    if not layout_text:
        return None

    match = re.search(
        r'(?m)^\s*(\d+)\s+HDL\s+HANDLING\s+FEE\s+\$?([\d,]+\.\d{2})\s+\$?([\d,]+\.\d{2})\s*$',
        layout_text,
        re.IGNORECASE,
    )
    if not match:
        return None

    return {
        'item_number': 'HDL',
        'quantity': match.group(1),
        'units': 'Each',
        'description': 'HANDLING FEE',
        'unit_price': _clean_price(match.group(2)),
        'amount': _clean_price(match.group(3)),
        'qb_category_override': 'Purchases',
        'qb_product_service_override': 'Drop Ship',
        'qb_sku_override': 'HDL',
    }


def _sum_item_amounts(items):
    total = 0.0
    found = False
    for item in items or []:
        amount = _clean_price(item.get('amount', ''))
        if not amount:
            continue
        try:
            total += float(amount)
            found = True
        except Exception:
            continue
    if not found:
        return ''
    return f"{total:.2f}"


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
            # PPE invoices may leave "Invoiced Qt" blank; use "Order Qty" for export quantity.
            qty = match.group(3)
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
            desc = ''
            if i + 1 < len(lines) and re.match(r'^Drop\s+Ship\s+Fee', lines[i + 1], re.IGNORECASE):
                desc = lines[i + 1].strip()
                i += 1
            item = {
                'item_number': 'Drop Ship',
                'quantity': drop_match.group(1),
                'units': 'Each',
                'description': desc or 'Drop Ship',
                'unit_price': drop_match.group(3).replace(',', ''),
                'amount': drop_match.group(4).replace(',', ''),
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
        r'^([A-Za-z0-9][A-Za-z0-9\-\/\.]*\s+CORE)\s+(\d+)\s+\d+\s*(?:Each|EA|Piece|pc|pcs|units?)?$',
        content, re.IGNORECASE
    )
    if match:
        return {
            'item_number': match.group(1).strip(),
            'quantity': match.group(2),
            'units': 'Each',
            'description': '',
            'unit_price': unit_price,
            'amount': amount,
        }

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
    if _is_river_city_vendor_name(vendor_name):
        river_city_items = _extract_river_city_items_from_layout(filepath)
        if river_city_items:
            return river_city_items

    if _is_power_stroke_vendor_name(vendor_name):
        power_stroke_items = _extract_power_stroke_items_from_layout(filepath)
        if power_stroke_items:
            return power_stroke_items

    if _is_bosch_vendor_name(vendor_name):
        bosch_items = _extract_bosch_items_from_layout(filepath)
        if bosch_items:
            return bosch_items

    if _is_diesel_forward_vendor_name(vendor_name):
        diesel_forward_items = _extract_diesel_forward_items_from_layout(filepath)
        if diesel_forward_items:
            return diesel_forward_items

    if _is_beans_vendor_name(vendor_name):
        beans_items = _extract_beans_items_from_layout(filepath)
        if beans_items:
            return beans_items

    if _is_carli_vendor_name(vendor_name):
        carli_items = _extract_carli_items_from_layout(filepath)
        if carli_items:
            return carli_items

    if _is_icon_vendor_name(vendor_name) or _is_cognito_vendor_name(vendor_name):
        icon_cognito_items = _extract_icon_cognito_items_from_layout(filepath)
        if icon_cognito_items:
            return icon_cognito_items
        reprint_items = _extract_carli_items_from_layout(filepath)
        if reprint_items:
            return reprint_items

    if _is_holley_vendor_name(vendor_name):
        holley_items = _extract_holley_items_from_layout(filepath)
        if holley_items:
            return holley_items

    if _is_ats_vendor_name(vendor_name):
        ats_items = _extract_ats_items_from_layout(filepath)
        if ats_items:
            return ats_items

    if _is_isspro_vendor_name(vendor_name):
        isspro_items = _extract_isspro_items_from_layout(filepath)
        if isspro_items:
            return isspro_items

    if _is_rock_krawler_vendor_name(vendor_name):
        rock_krawler_items = _extract_rock_krawler_items_from_layout(filepath)
        if rock_krawler_items:
            return rock_krawler_items

    if _is_sport_truck_vendor_name(vendor_name):
        sport_truck_items = _extract_sport_truck_items_from_layout(filepath)
        if sport_truck_items:
            return sport_truck_items

    if _is_valair_vendor_name(vendor_name):
        valair_items = _extract_valair_items_from_layout(filepath)
        if valair_items:
            return valair_items

    if _is_pt_vendor_name(vendor_name):
        pt_items = _extract_pt_items_from_words(filepath)
        if pt_items:
            return pt_items

    if _is_fumoto_vendor_name(vendor_name):
        fumoto_items = _extract_fumoto_items_from_tables(filepath)
        if fumoto_items:
            return fumoto_items

    if _is_dynomite_vendor_name(vendor_name):
        dd_items = _extract_dd_items_from_words(filepath)
        if dd_items:
            return dd_items

    if _is_poly_vendor_name(vendor_name):
        poly_items = _extract_poly_items_from_words(filepath)
        if poly_items:
            return poly_items

    if (
        _is_merchant_vendor_name(vendor_name)
        or _is_kc_turbos_vendor_name(vendor_name)
        or _is_mishimoto_vendor_name(vendor_name)
    ):
        ma_kt_items = _extract_ma_kt_items_from_words(filepath)
        if ma_kt_items:
            return ma_kt_items

    if _is_serra_vendor_name(vendor_name):
        serra_items = _extract_serra_items_from_words(filepath)
        if serra_items:
            return serra_items

    if _is_suspensionmaxx_vendor_name(vendor_name):
        sm_items = _extract_sm_items_from_words(filepath)
        if sm_items:
            return sm_items

    # Step 1: Try pdfplumber table extraction (most reliable)
    sb_mode = _is_sb_vendor_name(vendor_name)
    pd_mode = _is_pd_vendor_name(vendor_name)
    items = extract_items_from_tables(filepath, sb_mode=sb_mode, pd_mode=pd_mode)

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
                                  'purchase order', 'purchase order number',
                                  'customer po', 'customer po#'):
                        if not fields.get('po_number') and re.match(r'\d', val):
                            fields['po_number'] = val

                    # Terms
                    if header in ('terms', 'payment terms'):
                        if not fields.get('terms'):
                            fields['terms'] = val

                    # Total
                    if header in ('total due', 'invoice total', 'amount due'):
                        if not fields.get('total'):
                            fields['total'] = _clean_price(val) or val

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


def _apply_vendor_specific_overrides(data, text, filepath=None):
    """Apply vendor-specific fixes without changing the generic parser paths."""
    vendor_name = normalize_vendor_name(data.get('vendor', ''))
    if vendor_name:
        data['vendor'] = vendor_name

    layout_text = ''

    def _get_layout_text():
        nonlocal layout_text
        if not layout_text:
            layout_text = extract_layout_text_from_pdf(filepath)
        return layout_text

    if _is_sb_vendor_name(vendor_name):
        default_address = get_vendor_default_address(vendor_name)
        default_terms = get_vendor_default_terms(vendor_name)
        if default_terms and not str(data.get('terms') or '').strip():
            data['terms'] = default_terms

        sb_shipping_cost = _extract_sb_shipping_cost(text)
        if sb_shipping_cost:
            data['shipping_cost'] = sb_shipping_cost
            data['shipping_description'] = 'Shipping'

        if _is_sb_new_template(text):
            customer = _extract_sb_new_template_customer(text)
            if customer:
                data['customer'] = _clean_sb_customer_name(customer)
            if default_address:
                data['vendor_address'] = default_address
        elif not str(data.get('vendor_address') or '').strip() and default_address:
            data['vendor_address'] = default_address

        cleaned_customer = _clean_sb_customer_name(data.get('customer', ''))
        if cleaned_customer:
            data['customer'] = cleaned_customer

    elif _is_river_city_vendor_name(vendor_name):
        default_address = get_vendor_default_address(vendor_name)
        if default_address:
            data['vendor_address'] = default_address

        default_terms = get_vendor_default_terms(vendor_name)
        if default_terms:
            data['terms'] = default_terms

        layout = _get_layout_text()
        if not data.get('invoice_number'):
            invoice_match = re.search(r'(?im)^\s*INVOICE\s+(\d+)\s*$', layout)
            if invoice_match:
                data['invoice_number'] = invoice_match.group(1)
        if not data.get('date'):
            date_match = re.search(r'(?im)^\s*DATE\s+([A-Za-z]{3,9}\s+\d{1,2},\s+\d{4})\s*$', layout)
            if date_match:
                data['date'] = _normalize_date_value(date_match.group(1))
        if not data.get('po_number'):
            po_match = re.search(r'(?im)^\s*\w+\s+(\d{4,})\s+Net\s+30\s+Days\s*$', layout)
            if po_match:
                data['po_number'] = po_match.group(1)
        total_match = re.search(r'(?im)^\s*TOTAL\s+(-?[\d,]+\.\d{2})\s*$', layout)
        if total_match:
            data['total'] = _clean_signed_price_token(total_match.group(1))
        if data.get('date'):
            data['due_date'] = ''
        for item in (data.get('line_items') or []):
            if str(item.get('unit_price') or '').strip():
                continue
            qty_text = _normalize_qty(item.get('quantity'))
            amount_text = _clean_price(item.get('amount'))
            if not qty_text or not amount_text:
                continue
            try:
                qty_value = float(str(qty_text).replace(',', ''))
                amount_value = float(str(amount_text).replace(',', ''))
            except (TypeError, ValueError):
                continue
            if abs(qty_value) < 1e-9:
                continue
            item['unit_price'] = f"{abs(amount_value / qty_value):.2f}"
        has_freight_item = any(item.get('is_freight') for item in (data.get('line_items') or []))
        if not has_freight_item and not str(data.get('shipping_cost') or '').strip():
            data['suppress_zero_shipping_row'] = True

    elif _is_ats_vendor_name(vendor_name):
        default_address = get_vendor_default_address(vendor_name)
        if default_address:
            data['vendor_address'] = default_address

        default_terms = get_vendor_default_terms(vendor_name)
        if default_terms:
            data['terms'] = default_terms

        flat_text = re.sub(r'\s+', ' ', text)
        invoice_match = re.search(r'#?(INVC\d{4,})\b', flat_text, re.IGNORECASE)
        if invoice_match:
            data['invoice_number'] = invoice_match.group(1).upper()

        date_match = re.search(r'(\d{1,2}/\d{1,2}/\d{4})\s+#INVC\d{4,}\b', flat_text, re.IGNORECASE)
        if date_match:
            data['date'] = _normalize_date_value(date_match.group(1))

        ship_to_lines = _extract_ats_ship_to_lines(filepath)
        customer = _clean_ship_to_contact_name(_customer_name_from_ship_to_lines(ship_to_lines))
        if customer:
            data['customer'] = customer

        if data.get('date'):
            data['due_date'] = ''
        has_freight_item = any(item.get('is_freight') for item in (data.get('line_items') or []))
        if not has_freight_item and not str(data.get('shipping_cost') or '').strip():
            data['suppress_zero_shipping_row'] = True

    elif _is_isspro_vendor_name(vendor_name):
        default_address = get_vendor_default_address(vendor_name)
        if default_address:
            data['vendor_address'] = default_address

        default_terms = get_vendor_default_terms(vendor_name)
        if default_terms:
            data['terms'] = default_terms

        layout = _get_layout_text()
        flat_text = re.sub(r'\s+', ' ', text)
        if not data.get('invoice_number'):
            invoice_match = re.search(r'\bINVOICE\s+(\d{5,})\b', flat_text, re.IGNORECASE)
            if invoice_match:
                data['invoice_number'] = invoice_match.group(1)
        if not data.get('po_number'):
            po_match = re.search(r'CUSTOMER\s+PURCHASE\s+ORDER\s+NBR\s+DATE\s+SHIPPED\s+(\d{4,})\s+\d{2}/\d{2}/\d{2}', flat_text, re.IGNORECASE)
            if po_match:
                data['po_number'] = po_match.group(1)
        if not data.get('date'):
            date_match = re.search(r'\b\d{7}\s+(\d{2}/\d{2}/\d{2})\b', flat_text)
            if date_match:
                data['date'] = _normalize_date_value(date_match.group(1))
        ship_to_lines = _extract_isspro_ship_to_lines(filepath)
        customer = _clean_ship_to_contact_name(_customer_name_from_ship_to_lines(ship_to_lines))
        if customer:
            data['customer'] = customer
        if data.get('date'):
            data['due_date'] = ''
        total_match = re.search(r'(?m)^\s*TOTAL\s+([\d,]+\.\d{2})\s*$', layout, re.IGNORECASE)
        if total_match:
            data['total'] = _clean_price(total_match.group(1))
        has_freight_item = any(item.get('is_freight') for item in (data.get('line_items') or []))
        if not has_freight_item and not str(data.get('shipping_cost') or '').strip():
            data['suppress_zero_shipping_row'] = True

    elif _is_power_stroke_vendor_name(vendor_name):
        default_address = get_vendor_default_address(vendor_name)
        if default_address:
            data['vendor_address'] = default_address

        default_terms = get_vendor_default_terms(vendor_name)
        if default_terms:
            data['terms'] = default_terms

        layout = _get_layout_text()
        if not data.get('invoice_number'):
            invoice_match = re.search(r'Invoice\s+no\.\s*:\s*([A-Za-z0-9-]+)', layout, re.IGNORECASE)
            if invoice_match:
                data['invoice_number'] = invoice_match.group(1)
        if not data.get('po_number'):
            po_match = re.search(r'Purchase\s+Order\s*:\s*(\d+)', layout, re.IGNORECASE)
            if po_match:
                data['po_number'] = po_match.group(1)
        if not data.get('date'):
            date_match = re.search(r'Invoice\s+date\s*:\s*(\d{2}/\d{2}/\d{4})', layout, re.IGNORECASE)
            if date_match:
                data['date'] = _normalize_date_value(date_match.group(1))

        ship_to_lines = _extract_power_stroke_ship_to_lines(filepath)
        customer = _clean_ship_to_contact_name(_customer_name_from_ship_to_lines(ship_to_lines))
        if customer:
            data['customer'] = customer

        total_match = re.search(r'(?m)^\s*Total\s+\$?([\d,]+\.\d{2})\s*$', layout, re.IGNORECASE)
        if total_match:
            data['total'] = _clean_price(total_match.group(1))

    elif _is_hamilton_vendor_name(vendor_name):
        default_address = get_vendor_default_address(vendor_name)
        if default_address:
            data['vendor_address'] = default_address

        default_terms = get_vendor_default_terms(vendor_name)
        if default_terms:
            data['terms'] = default_terms

        ship_to_lines = _extract_hamilton_ship_to_lines(filepath)
        customer = _clean_ship_to_contact_name(_customer_name_from_ship_to_lines(ship_to_lines))
        if customer:
            data['customer'] = customer

        layout = _get_layout_text()
        total_match = re.search(r'Total\s+\$?([\d,]+\.\d{2})', layout, re.IGNORECASE)
        if total_match:
            data['total'] = _clean_price(total_match.group(1))

    elif _is_beans_vendor_name(vendor_name):
        default_address = get_vendor_default_address(vendor_name)
        if default_address:
            data['vendor_address'] = default_address

        default_terms = get_vendor_default_terms(vendor_name)
        if default_terms:
            data['terms'] = default_terms

    elif _is_bosch_vendor_name(vendor_name):
        default_address = get_vendor_default_address(vendor_name)
        if default_address:
            data['vendor_address'] = default_address

        default_terms = get_vendor_default_terms(vendor_name)
        if default_terms:
            data['terms'] = default_terms

        layout = _get_layout_text()
        flat_text = re.sub(r'\s+', ' ', text)
        if not data.get('invoice_number'):
            invoice_match = re.search(r'(?m)^\s*\d+\s+\d{2}/\d{2}/\d{4}\s+(\d{6,})\s*$', layout)
            if invoice_match:
                data['invoice_number'] = invoice_match.group(1)
        if not data.get('po_number'):
            po_match = re.search(
                r'Customer\s+Reference\s+PO\s+Date\s+.*?\b(\d{4,})\b\s+\d{2}/\d{2}/\d{4}',
                flat_text,
                re.IGNORECASE,
            )
            if po_match:
                data['po_number'] = po_match.group(1)

        ship_to_lines = _extract_bosch_ship_to_lines(filepath)
        customer = _clean_ship_to_contact_name(_customer_name_from_ship_to_lines(ship_to_lines))
        if ship_to_lines:
            data['customer'] = customer

        total_match = re.search(r'Total\s+in\s+currency\s+USD\s+([\d,]+\.\d{2})', layout, re.IGNORECASE)
        if total_match:
            data['total'] = _clean_price(total_match.group(1))

    elif _is_diesel_forward_vendor_name(vendor_name):
        default_address = get_vendor_default_address(vendor_name)
        if default_address:
            data['vendor_address'] = default_address

        default_terms = get_vendor_default_terms(vendor_name)
        if default_terms:
            data['terms'] = default_terms

        layout = _get_layout_text()
        if not data.get('invoice_number'):
            invoice_match = re.search(r'Invoice\s+Number\s*:\s*([A-Za-z0-9-]+)', layout, re.IGNORECASE)
            if invoice_match:
                data['invoice_number'] = invoice_match.group(1)
        if not data.get('po_number'):
            po_match = re.search(r'P\.O\.\s+Number\s+(\d+)', layout, re.IGNORECASE)
            if po_match:
                data['po_number'] = po_match.group(1)
        if not data.get('date'):
            date_match = re.search(r'Invoice\s+Date\s*:\s*(\d{1,2}/\d{1,2}/\d{4})', layout, re.IGNORECASE)
            if date_match:
                data['date'] = _normalize_date_value(date_match.group(1))

        ship_to_lines = _extract_diesel_forward_ship_to_lines(filepath)
        customer = _clean_ship_to_contact_name(_customer_name_from_ship_to_lines(ship_to_lines))
        if customer:
            data['customer'] = customer

        total_amount = _extract_ppe_total_usd(layout) or _extract_ppe_total_usd(text)
        if total_amount:
            data['total'] = total_amount

    elif _is_carli_vendor_name(vendor_name):
        default_address = get_vendor_default_address(vendor_name)
        if default_address:
            data['vendor_address'] = default_address

        default_terms = get_vendor_default_terms(vendor_name)
        if default_terms:
            data['terms'] = default_terms

        layout = _get_layout_text()
        if not data.get('invoice_number'):
            invoice_match = re.search(r'Invoice\s+No\.\s*:\s*([A-Za-z0-9-]+)', layout, re.IGNORECASE)
            if invoice_match:
                data['invoice_number'] = invoice_match.group(1)
        if not data.get('invoice_number') or not data.get('date'):
            header_match = re.search(
                r'\bINVOICE\b\s*[\r\n]+\s*([A-Za-z0-9-]+)\s+(\d{1,2}/\d{1,2}/\d{4})\b',
                layout,
                re.IGNORECASE,
            )
            if header_match:
                if not data.get('invoice_number'):
                    data['invoice_number'] = header_match.group(1)
                if not data.get('date'):
                    data['date'] = _normalize_date_value(header_match.group(2))
        if not data.get('date'):
            date_match = re.search(r'Invoice\s+Date\s*:\s*(\d{1,2}/\d{1,2}/\d{4})', layout, re.IGNORECASE)
            if date_match:
                data['date'] = _normalize_date_value(date_match.group(1))
        if not data.get('po_number'):
            po_match = re.search(
                r'Pack\s+Slip\s+#(?:\s+Ship\s+Date)?\s+PO\s+Number\s+Order\s+Date.*?\n'
                r'\s*\S+(?:\s+\d{1,2}/\d{1,2}/\d{4})?\s+(\d{4,})\b',
                layout,
                re.IGNORECASE | re.DOTALL,
            )
            if po_match:
                data['po_number'] = po_match.group(1)

        ship_to_lines = _extract_carli_ship_to_lines(filepath)
        customer = _clean_ship_to_contact_name(_customer_name_from_ship_to_lines(ship_to_lines))
        if customer:
            data['customer'] = customer

    elif _is_icon_vendor_name(vendor_name) or _is_cognito_vendor_name(vendor_name):
        default_address = get_vendor_default_address(vendor_name)
        if default_address:
            data['vendor_address'] = default_address

        default_terms = get_vendor_default_terms(vendor_name)
        if default_terms:
            data['terms'] = default_terms

        layout = _get_layout_text()
        if not data.get('invoice_number') or not data.get('date'):
            header_match = re.search(
                r'(?m)^\s*([A-Za-z0-9-]+)\s+(\d{1,2}/\d{1,2}/\d{4})\s*$',
                layout,
            )
            if header_match:
                if not data.get('invoice_number'):
                    data['invoice_number'] = header_match.group(1)
                if not data.get('date'):
                    data['date'] = _normalize_date_value(header_match.group(2))
        if not data.get('invoice_number'):
            invoice_match = re.search(r'Invoice\s+No\.\s*:\s*([A-Za-z0-9-]+)', layout, re.IGNORECASE)
            if invoice_match:
                data['invoice_number'] = invoice_match.group(1)
        if not data.get('date'):
            date_match = re.search(r'Invoice\s+Date\s*:\s*(\d{1,2}/\d{1,2}/\d{4})', layout, re.IGNORECASE)
            if date_match:
                data['date'] = _normalize_date_value(date_match.group(1))
        if not data.get('po_number'):
            po_match = re.search(
                r'Pack\s+Slip\s+#\s+PO\s+Number\s+Order\s+Date.*?\n\s*\S+\s+(\d{4,})\s+\d{1,2}/\d{1,2}/\d{4}\b',
                layout,
                re.IGNORECASE | re.DOTALL,
            )
            if po_match:
                data['po_number'] = po_match.group(1)
        if not data.get('po_number'):
            po_match = re.search(
                r'Pack\s+Slip\s+#\s+Ship\s+Date\s+PO\s+Number.*?\n\s*\S+\s+\d{1,2}/\d{1,2}/\d{4}\s+([A-Za-z0-9-]+)\b',
                layout,
                re.IGNORECASE | re.DOTALL,
            )
            if po_match:
                data['po_number'] = po_match.group(1)

        ship_to_lines = _extract_icon_cognito_ship_to_lines(filepath)
        customer = _clean_ship_to_contact_name(_customer_name_from_ship_to_lines(ship_to_lines))
        if customer:
            data['customer'] = customer
        freight_match = re.search(r'(?m)^\s*Freight:\s*([\d,]+\.\d{2})\s*$', layout, re.IGNORECASE)
        if freight_match:
            shipping_cost = _clean_price(freight_match.group(1))
            if shipping_cost != '':
                data['shipping_cost'] = shipping_cost
                if not data.get('shipping_description'):
                    data['shipping_description'] = 'Freight'

    elif _is_rock_krawler_vendor_name(vendor_name):
        default_address = get_vendor_default_address(vendor_name)
        if default_address:
            data['vendor_address'] = default_address

        default_terms = get_vendor_default_terms(vendor_name)
        if default_terms:
            data['terms'] = default_terms

        layout = _get_layout_text()
        rk_po_number = _extract_rock_krawler_po_number(filepath)
        if rk_po_number:
            data['po_number'] = rk_po_number
        ship_to_lines = _extract_rock_krawler_ship_to_lines(filepath)
        customer = _clean_ship_to_contact_name(_customer_name_from_ship_to_lines(ship_to_lines))
        if customer:
            data['customer'] = customer
        shipping_match = re.search(r'\bSHIPPING\s+([\d,]+\.\d{2})\b', layout, re.IGNORECASE)
        if shipping_match:
            data['shipping_cost'] = _clean_price(shipping_match.group(1))
            data['shipping_description'] = 'Freight'
        has_freight_item = any(item.get('is_freight') for item in (data.get('line_items') or []))
        if not has_freight_item and not str(data.get('shipping_cost') or '').strip():
            data['suppress_zero_shipping_row'] = True

    elif _is_sport_truck_vendor_name(vendor_name):
        default_address = get_vendor_default_address(vendor_name)
        if default_address:
            data['vendor_address'] = default_address

        default_terms = get_vendor_default_terms(vendor_name)
        if default_terms:
            data['terms'] = default_terms

        layout = _get_layout_text()
        flat_text = re.sub(r'\s+', ' ', text)
        if not data.get('invoice_number'):
            invoice_match = re.search(r'P\.O\.\s*Box\s+\d+\s+(\d{6,})\b', flat_text, re.IGNORECASE)
            if invoice_match:
                data['invoice_number'] = invoice_match.group(1)
        ship_to_lines = _extract_sport_truck_ship_to_lines(filepath)
        customer = _clean_ship_to_contact_name(_customer_name_from_ship_to_lines(ship_to_lines))
        if customer:
            data['customer'] = customer
        if data.get('date'):
            data['date'] = _normalize_date_value(data.get('date'))
        if data.get('date'):
            data['due_date'] = ''
        shipping_match = re.search(
            r'\bFREIGHT(?:\s+FOB:ORIGIN:)?\s+([\d,]+\.\d{2})\b',
            layout,
            re.IGNORECASE,
        )
        if shipping_match:
            data['shipping_cost'] = _clean_price(shipping_match.group(1))
            data['shipping_description'] = 'Freight'
        has_freight_item = any(item.get('is_freight') for item in (data.get('line_items') or []))
        if not has_freight_item and not str(data.get('shipping_cost') or '').strip():
            data['suppress_zero_shipping_row'] = True

    elif _is_mishimoto_vendor_name(vendor_name):
        layout = _get_layout_text()
        header_text = layout.split('Bill To', 1)[0] if 'Bill To' in layout else layout[:500]
        invoice_match = re.search(r'\bINV\d{5,}\b', layout, re.IGNORECASE)
        if invoice_match:
            data['invoice_number'] = invoice_match.group(0)

        header_dates = re.findall(r'\b\d{2}/\d{2}/\d{4}\b', header_text)
        if header_dates:
            data['date'] = header_dates[-1]

        terms_match = re.search(
            r'(?m)^\s*(Net\s+15)\s+(\d{2}/\d{2}/\d{4})\s+Sales\s+Order\s+(\d+)\b',
            layout,
        )
        if terms_match:
            data['terms'] = terms_match.group(1)
            data['due_date'] = terms_match.group(2)
            data['po_number'] = terms_match.group(3)

        total_match = re.search(r'Amount\s+Due\s+\$?([\d,]+\.\d{2})', layout, re.IGNORECASE)
        if not total_match:
            total_match = re.search(r'(?m)^\s*Total\s+\$?([\d,]+\.\d{2})\s*$', layout, re.IGNORECASE)
        if total_match:
            data['total'] = _clean_price(total_match.group(1))

        shipping_match = re.search(r'Shipping\s+Cost\s+\$?([\d,]+\.\d{2})', layout, re.IGNORECASE)
        if shipping_match:
            data['shipping_cost'] = _clean_price(shipping_match.group(1))
            data['shipping_description'] = 'Drop Ship'

        ship_to_lines = _extract_mishimoto_ship_to_lines(filepath)
        customer = _customer_name_from_ship_to_lines(ship_to_lines)
        if customer:
            data['customer'] = customer

    elif _is_holley_vendor_name(vendor_name):
        holley_address = get_vendor_default_address(vendor_name)
        if holley_address:
            data['vendor_address'] = holley_address

        customer = _extract_holley_customer_from_layout(filepath)
        if customer:
            data['customer'] = customer

        if not data.get('po_number'):
            layout = _get_layout_text()
            po_match = re.search(r'(?m)^\s*C?\d[\w-]*\s+(\d{4,})\s+\S+', layout)
            if po_match:
                data['po_number'] = po_match.group(1)

    elif _is_valair_vendor_name(vendor_name):
        default_terms = get_vendor_default_terms(vendor_name)
        current_terms = str(data.get('terms') or '').strip()
        if default_terms and (not current_terms or not re.search(r'\d', current_terms)):
            data['terms'] = default_terms
        if not data.get('tracking_number'):
            tracking_match = re.search(
                r'(?im)^\s*(?:\d+(?:\.\d+)?)?\s*Tracking\s+#\s+([A-Z0-9]{10,})\b',
                text,
            )
            if tracking_match:
                data['tracking_number'] = tracking_match.group(1)
        freight_items = [item for item in (data.get('line_items') or []) if item.get('is_freight')]
        shipping_cost = str(data.get('shipping_cost') or '').strip()
        shipping_desc = str(data.get('shipping_description') or '').strip()

        if not shipping_cost and freight_items:
            first_freight_item = freight_items[0]
            shipping_cost = _clean_price(
                first_freight_item.get('amount') or first_freight_item.get('unit_price') or ''
            )
            if shipping_cost:
                data['shipping_cost'] = shipping_cost
            if not shipping_desc:
                data['shipping_description'] = (
                    first_freight_item.get('description')
                    or first_freight_item.get('item_number')
                    or 'Freight'
                )
                shipping_desc = str(data.get('shipping_description') or '').strip()

        if not shipping_cost:
            layout_shipping_cost, layout_shipping_desc = _extract_valair_shipping_from_layout(filepath)
            if layout_shipping_cost:
                data['shipping_cost'] = layout_shipping_cost
                shipping_cost = layout_shipping_cost
            if layout_shipping_desc and not shipping_desc:
                data['shipping_description'] = layout_shipping_desc
                shipping_desc = layout_shipping_desc

        if not shipping_cost:
            data['shipping_cost'] = '0'
        if not shipping_desc:
            data['shipping_description'] = 'Freight'
        data['suppress_zero_shipping_row'] = False

    elif _is_pt_vendor_name(vendor_name):
        layout = _get_layout_text()
        header_text = layout.split('Part Number', 1)[0] if 'Part Number' in layout else layout

        invoice_match = re.search(r'Inv\s*#\s*([0-9][0-9 ]*\d)', layout, re.IGNORECASE)
        if invoice_match:
            data['invoice_number'] = re.sub(r'\s+', ' ', invoice_match.group(1)).strip()

        po_match = re.search(r'P\/O\s*#\s*(\d+)', layout, re.IGNORECASE)
        if po_match:
            data['po_number'] = po_match.group(1)

        header_dates = re.findall(r'\b\d{1,2}/\d{1,2}/\d{4}\b', header_text)
        if header_dates:
            data['date'] = header_dates[-1]

        ship_to_lines = _extract_pt_ship_to_lines(filepath)
        customer = _customer_name_from_ship_to_lines(ship_to_lines)
        if customer:
            data['customer'] = customer

        if not data.get('total') and data.get('line_items'):
            data['total'] = _sum_item_amounts(data.get('line_items'))

    elif _is_merchant_vendor_name(vendor_name):
        layout = _get_layout_text()
        header_text = layout.split('Bill To', 1)[0] if 'Bill To' in layout else layout[:500]

        invoice_match = re.search(r'#\s*(INV\d+)\b', layout, re.IGNORECASE)
        if invoice_match:
            data['invoice_number'] = invoice_match.group(1)

        if not data.get('date'):
            header_dates = re.findall(r'\b\d{1,2}/\d{1,2}/\d{4}\b', header_text)
            if header_dates:
                data['date'] = header_dates[-1]

        po_match = re.search(
            r'Terms\s+Due\s+Date\s+PO\s*#.*?\n\s*(?:\d{1,2}/\d{1,2}/\d{4}\s+)?(\d{4,})\b',
            layout,
            re.IGNORECASE | re.DOTALL,
        )
        if po_match:
            data['po_number'] = po_match.group(1)

        ship_to_lines = _extract_ma_kt_ship_to_lines(filepath)
        customer = _customer_name_from_ship_to_lines(ship_to_lines)
        if customer:
            data['customer'] = customer

        total_match = re.search(r'(?m)^\s*Total\s+\$?([\d,]+\.\d{2})\s*$', layout, re.IGNORECASE)
        if not total_match:
            total_match = re.search(r'Amount\s+Due\s+\$?([\d,]+\.\d{2})', layout, re.IGNORECASE)
        if total_match:
            data['total'] = _clean_price(total_match.group(1))

    elif _is_dynomite_vendor_name(vendor_name):
        layout = _get_layout_text()

        invoice_match = re.search(r'INVOICE\s*#\s*([0-9]+/[0-9]+)', layout, re.IGNORECASE)
        if invoice_match:
            data['invoice_number'] = invoice_match.group(1)

        terms_match = re.search(
            r'TERMS\s+([0-9]+%\s+[0-9]+\s+Net\s+[0-9]+)',
            layout,
            re.IGNORECASE,
        )
        if terms_match:
            data['terms'] = re.sub(r'\s+', ' ', terms_match.group(1)).strip()

        ship_to_lines = _extract_dd_ship_to_lines(filepath)
        customer = _customer_name_from_ship_to_lines(ship_to_lines)
        if not customer:
            customer = _extract_dd_customer_from_text(text)
        if customer:
            data['customer'] = customer

        if not data.get('shipping_cost'):
            dd_shipping_cost = _extract_dd_shipping_cost(text) or _extract_dd_shipping_cost(layout)
            if dd_shipping_cost:
                data['shipping_cost'] = dd_shipping_cost
                if not data.get('shipping_description'):
                    data['shipping_description'] = 'Shipping'

    elif _is_kc_turbos_vendor_name(vendor_name):
        layout = _get_layout_text()
        header_text = layout.split('Bill To', 1)[0] if 'Bill To' in layout else layout[:500]

        if not data.get('invoice_number'):
            invoice_match = re.search(r'#\s*(INV\d+)\b', layout, re.IGNORECASE)
            if invoice_match:
                data['invoice_number'] = invoice_match.group(1)

        if not data.get('date'):
            header_dates = re.findall(r'\b\d{1,2}/\d{1,2}/\d{4}\b', header_text)
            if header_dates:
                data['date'] = header_dates[-1]

        po_match = re.search(
            r'Terms\s+Due\s+Date\s+PO\s*#.*?\n\s*(?:Net\s+30\s+)?(?:\d{1,2}/\d{1,2}/\d{2,4}\s+)?(\d{4,})\b',
            layout,
            re.IGNORECASE | re.DOTALL,
        )
        if po_match:
            data['po_number'] = po_match.group(1)

        ship_to_lines = _extract_ma_kt_ship_to_lines(filepath)
        customer = _customer_name_from_ship_to_lines(ship_to_lines)
        if customer:
            data['customer'] = customer

    elif _is_daystar_vendor_name(vendor_name):
        ship_to_lines = _extract_daystar_ship_to_lines(filepath)
        customer = _customer_name_from_ship_to_lines(ship_to_lines)
        if customer:
            data['customer'] = customer

    elif _is_poly_vendor_name(vendor_name):
        ship_to_lines = _extract_poly_ship_to_lines(filepath)
        customer = _customer_name_from_ship_to_lines(ship_to_lines)
        if customer:
            customer = re.sub(r'\s+1Z[0-9A-Z]{16}\b.*$', '', customer, flags=re.IGNORECASE).strip()
            customer = re.sub(r'\s+[A-HJ-NPR-Z0-9]{17}\b$', '', customer).strip()
            data['customer'] = customer

    elif _is_serra_vendor_name(vendor_name):
        default_address = get_vendor_default_address(vendor_name)
        if default_address:
            data['vendor_address'] = default_address
        else:
            data['vendor_address'] = ''

        date_match = re.search(r'Ship\s+Date\s+(\d{2}/\d{2}/\d{4})', text, re.IGNORECASE)
        if not date_match:
            date_match = re.search(r'Printed\s+(\d{2}/\d{2}/\d{4})', text, re.IGNORECASE)
        if date_match:
            data['date'] = date_match.group(1)

    elif _is_suspensionmaxx_vendor_name(vendor_name):
        layout = _get_layout_text()

        invoice_match = re.search(
            r'Invoice\s+no\.\s*:\s*([0-9][0-9-]*(?:\s+DS)?)',
            layout,
            re.IGNORECASE,
        )
        if invoice_match:
            data['invoice_number'] = re.sub(r'\s+', ' ', invoice_match.group(1)).strip()

        po_match = re.search(r'P\.O\.\s*#:\s*(\d+)', layout, re.IGNORECASE)
        if po_match:
            data['po_number'] = po_match.group(1)

        date_match = re.search(r'Invoice\s+date:\s*(\d{2}/\d{2}/\d{4})', layout, re.IGNORECASE)
        if date_match:
            data['date'] = date_match.group(1)

        total_match = re.search(r'(?m)^\s*Total\s+\$?([\d,]+\.\d{2})\s*$', layout, re.IGNORECASE)
        if total_match:
            data['total'] = _clean_price(total_match.group(1))

        ship_to_lines = _extract_sm_ship_to_lines(filepath)
        customer = _customer_name_from_ship_to_lines(ship_to_lines)
        if customer:
            data['customer'] = customer

    elif _is_fumoto_vendor_name(vendor_name):
        layout = _get_layout_text()
        customer = _extract_fumoto_ship_to_name(filepath)
        if customer:
            data['customer'] = customer

        po_match = re.search(r'CUSTOMER\s+PO\s+(\d+)', layout, re.IGNORECASE)
        if po_match:
            data['po_number'] = po_match.group(1)

        total_match = re.search(
            r'(?m)^\s*Ticket-\d+\s+\d{2}/\d{2}/\d{4}\s+\$?([\d,]+\.\d{2})\b',
            layout,
        )
        if total_match:
            data['total'] = _clean_price(total_match.group(1))

    elif _is_diamond_eye_vendor_name(vendor_name):
        layout = _get_layout_text()
        has_handling_fee = any(
            str(item.get('item_number', '')).strip().upper() == 'HDL'
            or 'handling fee' in str(item.get('description', '')).lower()
            for item in (data.get('line_items') or [])
        )
        if not has_handling_fee:
            handling_item = _extract_diamond_eye_handling_fee(layout)
            if handling_item:
                data.setdefault('line_items', []).append(handling_item)

    return data


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
        r'Terms\s*:?\s*(Due\s+(?:on|Upon)\s+[Rr]eceipt)',      # "Terms: Due on receipt"
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
    if _is_sb_vendor_name(data.get('vendor', '')):
        sb_shipping_cost = _extract_sb_shipping_cost(text)
        if sb_shipping_cost:
            data['shipping_cost'] = sb_shipping_cost
            if not data.get('shipping_description'):
                data['shipping_description'] = 'Shipping'
    _shipping_patterns = [
        (r'Shipping\s+Cost\s*\([^)]+\)\s*\$?([\d,]+\.?\d*)', 'Shipping'),   # S&B
        (r'Drop\s+Ship\s+\$?([\d,]+\.?\d*)', 'Drop Ship'),                    # FL, PPE
        (r'FREIGHT\s+OUT\s+\$?([\d,]+\.?\d*)', 'Freight'),                    # II
        (r'Freight\s+\$?([\d,]+\.?\d*)', 'Freight'),                          # T14, general
    ]
    if _is_ppe_vendor_name(data.get('vendor', '')):
        _shipping_patterns[1:1] = [
            (r'(?im)^Drop\s+Ship\s+\d+\.?\d*\s+\d+\.?\d*\s+([\d,]+\.?\d{2})(?:\s+[\d,]+\.?\d{2})?\s*$', 'Drop Ship'),  # PPE full row
            (r'(?im)^Drop\s+Ship\s+\d+\.?\d*\s+([\d,]+\.?\d{2})\s*$', 'Drop Ship'),  # PPE abbreviated row
        ]
    if not data.get('shipping_cost'):
        for _pat, _desc in _shipping_patterns:
            _m = re.search(_pat, text, re.IGNORECASE | re.MULTILINE)
            if _m:
                data['shipping_cost'] = _m.group(1).strip()
                if not data.get('shipping_description'):
                    data['shipping_description'] = _desc
                break
    if not data.get('shipping_cost'):
        data['shipping_cost'] = ''
    if _is_ppe_vendor_name(data.get('vendor', '')):
        data['shipping_quantity'] = _extract_ppe_drop_ship_quantity(text)

    # --- Total ---
    if _is_ppe_vendor_name(data.get('vendor', '')):
        data['total'] = _extract_ppe_total_usd(text)
    elif _is_fl_vendor_name(data.get('vendor', '')):
        data['total'] = _extract_fleece_total(text)
    elif _is_ii_vendor_name(data.get('vendor', '')):
        data['total'] = _extract_ii_total(text)
    else:
        data['total'] = _extract_total_amount(text)
    if not data['total']:
        data['total'] = _extract_total_amount(text)
    if not data['total']:
        data['total'] = parse_field(text, [
            r'(?:Total\s+USD|Total\s+Amount|Invoice\s+Total|Grand\s+Total|Total\s+Due|^Total)\s*:?\s*\$?([\d,]+\.?\d*)',
            r'(?:^|\n)\s*Total\s+\$?([\d,]+\.?\d*)',
            r'Amount\s+Due\s*:?\s*\$?([\d,]+\.?\d*)',
            r'Balance\s+Due\s+\$?([\d,]+\.?\d*)',
        ])
    if not data['total']:
        data['total'] = table_fields.get('total', '')

    data = _refresh_vendor_dependent_fields(data, text, filepath)

    return data


def _refresh_vendor_dependent_fields(data, text, filepath=None):
    """Re-run vendor-specific item extraction and overrides after vendor resolution changes."""
    data['line_items'] = extract_line_items(text, filepath, vendor_name=data.get('vendor'))
    if (
        _is_redhead_vendor_name(data.get('vendor', ''))
        or _is_isspro_vendor_name(data.get('vendor', ''))
        or _is_suspensionmaxx_vendor_name(data.get('vendor', ''))
    ):
        _apply_item_style_discount_overrides(data['line_items'])
    if _is_turn14_vendor_name(data.get('vendor', '')):
        has_discount_item = any(
            bool(item.get('is_discount')) or 'discount' in (
                f"{item.get('item_number', '')} {item.get('description', '')}"
            ).lower()
            for item in data['line_items']
        )
        if not has_discount_item:
            footer_discount_item = _extract_turn14_footer_discount_item(text)
            if footer_discount_item:
                data['line_items'].append(footer_discount_item)
    if data.get('line_items'):
        freight_items = [i for i in data['line_items'] if i.get('is_freight')]
        if freight_items and not data.get('shipping_description'):
            desc = freight_items[0].get('description') or freight_items[0].get('item_number') or 'Freight'
            data['shipping_description'] = desc
    return _apply_vendor_specific_overrides(data, text, filepath)


def _refresh_invoice_for_resolved_vendor(data, text, filepath, vendor_name):
    """Re-run vendor-dependent extraction after metadata resolves the vendor."""
    resolved_vendor = normalize_vendor_name(vendor_name)
    if resolved_vendor:
        data['vendor'] = resolved_vendor

    data['line_items'] = extract_line_items(text, filepath, vendor_name=resolved_vendor)
    if data.get('line_items'):
        freight_items = [i for i in data['line_items'] if i.get('is_freight')]
        if freight_items and not data.get('shipping_description'):
            desc = freight_items[0].get('description') or freight_items[0].get('item_number') or 'Freight'
            data['shipping_description'] = desc

    return _apply_vendor_specific_overrides(data, text, filepath)


def _normalize_email_body_text(text):
    value = str(text or '')
    value = (
        value
        .replace('\u202f', ' ')
        .replace('\xa0', ' ')
        .replace('â€¯', ' ')
        .replace('Â', '')
        .replace('Ã—', '×')
    )
    value = re.sub(r'\r\n?', '\n', value)
    value = re.sub(r'[ \t]+', ' ', value)
    value = re.sub(r'\n{3,}', '\n\n', value)
    return value.strip()


def _parse_email_body_date(value):
    text = re.sub(r'\s+', ' ', str(value or '')).strip()
    if not text:
        return ''
    match = re.search(
        r'\b([A-Z][a-z]{2,8})\s+(\d{1,2}),\s*(\d{4})\b',
        text,
    )
    if not match:
        return ''
    for fmt in ('%B %d %Y', '%b %d %Y'):
        try:
            parsed = datetime.strptime(
                f"{match.group(1)} {match.group(2)} {match.group(3)}",
                fmt,
            )
            return f"{parsed.month}/{parsed.day}/{parsed.year}"
        except ValueError:
            continue
    return ''


def _extract_block_between_labels(text, start_label, end_labels):
    lines = [line.strip() for line in str(text or '').splitlines()]
    start_idx = None
    for idx, line in enumerate(lines):
        if re.fullmatch(start_label, line, re.IGNORECASE):
            start_idx = idx + 1
            break
    if start_idx is None:
        return []

    result = []
    for line in lines[start_idx:]:
        if any(re.fullmatch(label, line, re.IGNORECASE) for label in end_labels):
            break
        if line:
            result.append(line)
    return result


def _extract_sb_body_order_url(text):
    """Return the Shopify/S&B order URL from a body-invoice email."""
    body = str(text or '')

    def clean_url(value):
        url = unescape(str(value or '')).strip().strip('\'"<>).,')
        if not url:
            return ''
        parsed = urlparse(url)
        if parsed.netloc.lower().endswith('google.com') and parsed.path == '/url':
            query_url = parse_qs(parsed.query).get('q')
            if query_url:
                url = query_url[0]
        return unquote(url).strip().strip('\'"<>).,')

    view_match = re.search(
        r'(?is)\bView\s+your\s+order\b\s*<?(https?://[^>\s]+)',
        body,
    )
    if view_match:
        return clean_url(view_match.group(1))
    signed_match = re.search(
        r'https://(?:www\.)?(?:sbfilters\.com|shopify\.com)/[^\s<>]+/orders/[A-Za-z0-9]+/[^\s<>]*authenticate\?[^>\s]+',
        body,
        re.IGNORECASE,
    )
    if signed_match:
        return clean_url(signed_match.group(0))
    order_url_match = re.search(
        r'https://shopify\.com/\d+/account/orders/[A-Za-z0-9]+(?:/[A-Za-z0-9_/-]+)?(?:\?[^>\s]+)?',
        body,
        re.IGNORECASE,
    )
    return clean_url(order_url_match.group(0)) if order_url_match else ''


def _extract_sb_body_order_items(text):
    summary_match = re.search(
        r'(?is)\bOrder\s+summary\b(.+?)(?:\bSubtotal\b|\bCustomer\s+information\b|\bTotal\s+due\b)',
        text,
    )
    search_text = summary_match.group(1) if summary_match else text
    lines = [line.strip() for line in search_text.splitlines() if line.strip()]
    items = []
    seen = set()
    quantity_line_re = re.compile(r'(?i)^(.+?)\s*[x×]\s*(\d+)\s*$')
    price_re = re.compile(r'(?:\$([\d,]+(?:\.\d{2})?)\b|\b([\d,]+\.\d{2})\b)')

    for idx, line in enumerate(lines):
        if re.search(r'(?i)\b(Subtotal|Shipping|Taxes?|Total paid|Total due)\b', line):
            continue
        match = re.search(r'(?i)(.*?)\s*[x×]\s*(\d+)\s+\$?([\d,]+\.?\d{2})\b', line)
        variant = ''
        if match:
            description = match.group(1).strip(' -')
            quantity = match.group(2)
            description = f"{description} x {quantity}"
            amount = _clean_price(match.group(3))
        else:
            price_match = price_re.search(line)
            if not price_match:
                continue
            quantity = '1'
            amount = _clean_price(price_match.group(1) or price_match.group(2))
            description = price_re.sub('', line).strip(' -')
            if idx > 0:
                prev_match = re.search(r'(?i)^(.+?)\s*[x×]\s*(\d+)\s*$', lines[idx - 1].strip())
                if prev_match:
                    description = prev_match.group(1).strip(' -')
                    quantity = prev_match.group(2)
                    description = f"{description} x {quantity}"
                    variant = price_re.sub('', line).strip(' -')
                elif idx > 1:
                    prev_prev_match = quantity_line_re.search(lines[idx - 2].strip())
                    if prev_prev_match:
                        description = prev_prev_match.group(1).strip(' -')
                        quantity = prev_prev_match.group(2)
                        description = f"{description} x {quantity}"
                        variant = lines[idx - 1].strip(' -')

        if not description and idx > 0:
            description = lines[idx - 1].strip()
        if not description:
            continue

        if not variant and idx + 1 < len(lines):
            next_line = lines[idx + 1].strip()
            if (
                next_line
                and not price_re.search(next_line)
                and not re.search(r'(?i)\b(Subtotal|Shipping|Taxes?|Total paid|Total due)\b', next_line)
            ):
                variant = next_line
        full_description = f"{description} - {variant}" if variant else description
        key = (full_description.lower(), quantity, amount)
        if key in seen:
            continue
        seen.add(key)

        unit_price = amount
        try:
            qty_num = float(quantity)
            amount_num = float(amount)
            if qty_num:
                unit_price = f"{amount_num / qty_num:.2f}"
        except (TypeError, ValueError):
            pass

        items.append({
            'item_number': '',
            'quantity': _normalize_qty(quantity),
            'units': 'Each',
            'description': full_description,
            'unit_price': unit_price,
            'amount': amount,
        })

    return items


def _extract_sb_body_order_items(text):
    summary_match = re.search(
        r'(?is)\bOrder\s+summary\b(.+?)(?:\bSubtotal\b|\bCustomer\s+information\b|\bTotal\s+due\b)',
        text,
    )
    search_text = summary_match.group(1) if summary_match else text
    lines = [line.strip() for line in search_text.splitlines() if line.strip()]
    items = []
    seen = set()
    price_re = re.compile(r'(?:\$([\d,]+(?:\.\d{2})?)\b|\b([\d,]+\.\d{2})\b)')
    quantity_re = re.compile(r'(?i)\s*[x×]\s*(\d+)\s*$')
    boundary_re = re.compile(
        r'(?i)\b(Subtotal|Shipping|Taxes?|Total paid|Total due|Order summary|Customer information)\b'
    )

    def is_noise(line):
        return (
            not line
            or boundary_re.search(line)
            or re.match(r'https?://', line, re.IGNORECASE)
            or re.search(r'(?i)\b(View your order|Visit our store|Order Status AI Agent)\b', line)
        )

    def build_item(desc_lines, price_line):
        price_match = price_re.search(price_line)
        if not price_match:
            return None
        amount = _clean_price(price_match.group(1) or price_match.group(2))
        current_text = price_re.sub('', price_line).strip(' -')
        parts = [part for part in desc_lines if part and not is_noise(part)]
        if current_text:
            parts.append(current_text)
        if not parts:
            return None

        quantity = '1'
        qty_idx = None
        qty_prefix = ''
        for part_idx in range(len(parts) - 1, -1, -1):
            qty_match = quantity_re.search(parts[part_idx])
            if qty_match:
                quantity = qty_match.group(1)
                qty_idx = part_idx
                qty_prefix = quantity_re.sub('', parts[part_idx]).strip(' -')
                break

        if qty_idx is None:
            description = ' '.join(parts).strip(' -')
        else:
            product_parts = parts[:qty_idx]
            if qty_prefix:
                product_parts.append(qty_prefix)
            description = ' '.join(product_parts).strip(' -')
            variant = ' '.join(parts[qty_idx + 1:]).strip(' -')
            description = f"{description} x {quantity}".strip()
            if variant:
                description = f"{description} - {variant}"
        description = re.sub(r'\s+', ' ', description).strip(' -')
        if not description:
            return None

        unit_price = amount
        try:
            qty_num = float(quantity)
            amount_num = float(amount)
            if qty_num:
                unit_price = f"{amount_num / qty_num:.2f}"
        except (TypeError, ValueError):
            pass

        return {
            'item_number': '',
            'quantity': _normalize_qty(quantity),
            'units': 'Each',
            'description': description,
            'unit_price': unit_price,
            'amount': amount,
        }

    desc_buffer = []
    for line in lines:
        if is_noise(line):
            if desc_buffer and items:
                variant = ' '.join(desc_buffer).strip(' -')
                if variant and not price_re.search(variant) and not quantity_re.search(variant):
                    items[-1]['description'] = f"{items[-1]['description']} - {variant}"
                desc_buffer = []
            continue
        if price_re.search(line):
            item = build_item(desc_buffer, line)
            desc_buffer = []
            if not item:
                continue
            key = (item['description'].lower(), item['quantity'], item['amount'])
            if key in seen:
                continue
            seen.add(key)
            items.append(item)
            continue
        desc_buffer.append(line)

    if desc_buffer and items:
        variant = ' '.join(desc_buffer).strip(' -')
        if variant and not price_re.search(variant) and not quantity_re.search(variant):
            items[-1]['description'] = f"{items[-1]['description']} - {variant}"

    return items


def parse_email_invoice(filepath, status_callback=None):
    """Parse a saved no-attachment email body invoice source."""
    cb = status_callback or (lambda msg, tag=None: None)
    filename = os.path.basename(filepath)
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            payload = json.load(f)
    except Exception as e:
        cb(f"  Could not read email invoice source {filename}: {e}", "error")
        return None

    parser = str(payload.get('parser') or '').strip()
    if parser != 'sb_shopify_order':
        cb(f"  Unsupported email invoice parser for {filename}: {parser}", "error")
        return None

    text = _normalize_email_body_text(payload.get('message_text', ''))
    subject = str(payload.get('subject') or '')
    combined = f"{subject}\n{text}"
    order_url = str(payload.get('order_url') or '').strip() or _extract_sb_body_order_url(text)
    cb(f"  Parsing S&B body invoice data from {filename}...")

    order_match = re.search(r'(?i)\bOrder\s*#\s*(\d+)', combined)
    po_match = re.search(r'(?i)\bPO\s+Number\s*#?\s*(\d+)', combined)
    date_match = re.search(r'(?im)^\s*Date:\s*(.+)$', text)
    due_match = re.search(r'(?i)\bTotal\s+due\s+([A-Z][a-z]+\s+\d{1,2},\s*\d{4})', text)
    subtotal_match = re.search(r'(?i)\bSubtotal\b\s*\$?([\d,]+\.?\d{2})', text)
    shipping_match = re.search(r'(?i)\bShipping\b\s*\$?([\d,]+\.?\d{2})', text)
    total_match = re.search(
        r'(?i)\bTotal\s+due\s+[A-Z][a-z]+\s+\d{1,2},\s*\d{4}\s*\$?([\d,]+\.?\d{2})',
        text,
    )
    shipping_method_match = re.search(
        r'(?is)\bShipping\s+method\b\s*([^\n]+)',
        text,
    )

    shipping_address = _extract_block_between_labels(
        text,
        r'Shipping address',
        [r'Billing address', r'Location', r'Payment', r'Shipping method'],
    )
    customer = shipping_address[0] if shipping_address else ''
    line_items = _extract_sb_body_order_items(text)

    data = {
        'invoice_number': order_match.group(1) if order_match else '',
        'vendor': 'S&B Filters',
        'vendor_address': get_vendor_default_address('S&B Filters'),
        'customer': customer,
        'date': _parse_email_body_date(date_match.group(1) if date_match else ''),
        'due_date': _parse_email_body_date(due_match.group(1) if due_match else ''),
        'terms': 'Net 30',
        'po_number': po_match.group(1) if po_match else '',
        'tracking_number': '',
        'shipping_method': shipping_method_match.group(1).strip() if shipping_method_match else '',
        'ship_date': '',
        'shipping_tax_code': '',
        'shipping_tax_rate': '',
        'subtotal': _clean_price(subtotal_match.group(1)) if subtotal_match else '',
        'shipping_cost': _clean_price(shipping_match.group(1)) if shipping_match else '',
        'shipping_description': 'Shipping',
        'total': _clean_price(total_match.group(1)) if total_match else '',
        'line_items': line_items,
        'source_url': order_url,
        'source_file': filename,
        'raw_text': text,
    }

    has_invoice_signal = bool(data['invoice_number'] or data['po_number'] or line_items)
    if not has_invoice_signal:
        cb(f"  Not an invoice (no order no, PO, or line items): {filename}", "warning")
        return {'not_an_invoice': True}

    if not data['total'] and data['subtotal']:
        try:
            total = float(data['subtotal']) + float(data['shipping_cost'] or 0)
            data['total'] = f"{total:.2f}"
        except (TypeError, ValueError):
            pass

    filled = sum(
        1 for key, value in data.items()
        if key not in {'source_file', 'raw_text', 'line_items'} and value
    )
    cb(f"  Extracted {filled} fields + {len(line_items)} line item(s) from {filename}", "success")
    return data


def parse_invoice(
    filepath,
    status_callback=None,
    sender_email='',
    sender_header='',
    sender_subject='',
    sender_message_text='',
):
    """Parse a single invoice file and return structured data.

    Args:
        filepath: Path to the invoice file (PDF)
        status_callback: Optional function(msg, tag) for status updates

    Returns:
        dict with invoice data, or None if parsing failed
    """
    cb = status_callback or (lambda msg, tag=None: None)
    filename = os.path.basename(filepath)
    page_count = 1

    # Best effort page count (used for vendor-specific stock-order handling).
    try:
        with pdfplumber.open(filepath) as pdf:
            page_count = max(1, len(pdf.pages))
    except Exception:
        page_count = 1

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

    folder_vendor = infer_vendor_from_folder_marker(filepath)
    if folder_vendor:
        normalized_folder_vendor = normalize_vendor_name(folder_vendor)
        current_vendor = normalize_vendor_name(data.get('vendor', ''))
        if current_vendor != normalized_folder_vendor:
            cb(
                f"  Vendor inferred from training folder marker before validation: "
                f"{current_vendor or '(blank)'} -> {normalized_folder_vendor}"
            )
            data['vendor'] = normalized_folder_vendor
            data = _refresh_vendor_dependent_fields(data, text, filepath)

    parsed_vendor = normalize_vendor_name(data.get('vendor', ''))
    if (_is_reprint_vendor_name(parsed_vendor) or not parsed_vendor) and _text_matches_pt_layout(text):
        cb("  PT layout detected despite missing readable vendor text; applying PT parser.")
        data['vendor'] = 'Performance Turbochargers'
        parsed_vendor = normalize_vendor_name(data.get('vendor', ''))
        data = _refresh_vendor_dependent_fields(data, text, filepath)

    sender_vendor = infer_vendor_from_email_metadata(
        sender_email=sender_email,
        sender_header=sender_header,
        subject=sender_subject,
        message_text=sender_message_text,
    )
    if sender_vendor and not folder_vendor:
        normalized_sender_vendor = normalize_vendor_name(sender_vendor)
        current_vendor = normalize_vendor_name(data.get('vendor', ''))
        has_invoice_signal = bool(str(data.get('invoice_number', '')).strip())
        if not has_invoice_signal:
            has_invoice_signal = bool(str(data.get('po_number', '')).strip())
        if not has_invoice_signal:
            has_invoice_signal = any(
                str(item.get('amount', '')).strip()
                for item in (data.get('line_items') or [])
            )
        sender_ref = _extract_sender_email(sender_email) or str(sender_header or '').strip()
        sender_requires_refresh = False
        invoice_confirms_current_vendor = _text_explicitly_mentions_vendor(text, current_vendor)
        invoice_confirms_sender_vendor = _text_explicitly_mentions_vendor(text, normalized_sender_vendor)
        if (
            current_vendor
            and current_vendor != normalized_sender_vendor
            and invoice_confirms_current_vendor
            and not invoice_confirms_sender_vendor
        ):
            cb(
                f"  Preserving invoice-detected vendor {current_vendor}; "
                f"sender {sender_ref} mapped to {normalized_sender_vendor} "
                f"but the invoice text explicitly names {current_vendor}."
            )
        elif current_vendor != normalized_sender_vendor:
            if current_vendor:
                cb(
                    f"  Vendor overridden from sender {sender_ref}: "
                    f"{current_vendor} -> {normalized_sender_vendor}",
                    "warning"
                )
            else:
                cb(
                    f"  Vendor resolved from sender {sender_ref}: {normalized_sender_vendor}"
                )
            data['vendor'] = normalized_sender_vendor
            sender_requires_refresh = True
        elif not has_invoice_signal:
            cb(
                f"  Reapplying sender-confirmed vendor parser for {normalized_sender_vendor} "
                "before validation."
            )
            sender_requires_refresh = True
        if sender_requires_refresh:
            data = _refresh_vendor_dependent_fields(data, text, filepath)
            parsed_vendor = normalize_vendor_name(data.get('vendor', ''))

    # Step 4: Pre-validate — if no bill number, no PO number, and no line items,
    # this is not an invoice (e.g. return forms, flyers, packing slips).
    has_bill_no = bool(str(data.get('invoice_number', '')).strip())
    has_po = bool(str(data.get('po_number', '')).strip())
    has_line_items = any(
        str(item.get('amount', '')).strip()
        for item in (data.get('line_items') or [])
    )
    if not has_bill_no and not has_po and not has_line_items:
        cb(f"  Not an invoice (no bill no, PO, or line items): {filename}", "warning")
        return {'not_an_invoice': True}

    if folder_vendor:
        normalized_folder_vendor = normalize_vendor_name(folder_vendor)
        current_vendor = normalize_vendor_name(data.get('vendor', ''))
        if current_vendor != normalized_folder_vendor:
            cb(
                f"  Vendor overridden from training marker: "
                f"{current_vendor or '(blank)'} -> {normalized_folder_vendor}"
            )
            data['vendor'] = normalized_folder_vendor

    if not data.get('vendor'):
        data['vendor'] = infer_vendor_from_filename(filename)
    data['vendor'] = normalize_vendor_name(data.get('vendor', ''))
    resolved_vendor = normalize_vendor_name(data.get('vendor', ''))
    vendor_changed = parsed_vendor != resolved_vendor
    if vendor_changed and (
        _is_sb_vendor_name(resolved_vendor)
        or _is_mishimoto_vendor_name(resolved_vendor)
        or _is_holley_vendor_name(resolved_vendor)
        or _is_pt_vendor_name(resolved_vendor)
        or _is_fumoto_vendor_name(resolved_vendor)
        or _is_dynomite_vendor_name(resolved_vendor)
        or _is_poly_vendor_name(resolved_vendor)
        or _is_merchant_vendor_name(resolved_vendor)
        or _is_kc_turbos_vendor_name(resolved_vendor)
        or _is_serra_vendor_name(resolved_vendor)
        or _is_suspensionmaxx_vendor_name(resolved_vendor)
    ):
        data['line_items'] = extract_line_items(text, filepath, vendor_name=resolved_vendor)
        if data['line_items']:
            freight_items = [i for i in data['line_items'] if i.get('is_freight')]
            if freight_items:
                desc = freight_items[0].get('description') or freight_items[0].get('item_number') or 'Freight'
                data['shipping_description'] = desc

    data = _apply_vendor_specific_overrides(data, text, filepath)
    if not str(data.get('vendor_address', '')).strip():
        data['vendor_address'] = get_vendor_default_address(data.get('vendor', ''))
    if not str(data.get('terms', '')).strip():
        data['terms'] = get_vendor_default_terms(data.get('vendor', ''))
    if not str(data.get('due_date', '')).strip():
        due_date_days = get_vendor_due_date_days(data.get('vendor', ''))
        if due_date_days is not None:
            data['due_date'] = _derive_due_date_from_bill_date(
                data.get('date', ''),
                due_date_days,
            )
    data['page_count'] = page_count

    # PPE-specific rule: multi-page invoices are treated as stock orders.
    if page_count > 1 and _is_ppe_vendor_name(data.get('vendor', '')):
        _apply_stock_order_summary(
            data,
            description='STOCK ORDER',
            customer='Power Products Unlimited',
        )
        bill_no = str(data.get('invoice_number') or '').strip() or 'N/A'
        po_number = str(data.get('po_number') or '').strip() or 'N/A'
        cb(
            "  PPE multi-page invoice detected; outputting STOCK ORDER summary row "
            f"(Bill No: {bill_no}, Memo/PO: {po_number}; line items and totals suppressed).",
            "warning"
        )
    elif _is_fl_vendor_name(data.get('vendor', '')) and _has_fleece_stock_order_marker(text):
        stock_customer = 'Diesel Power Products' if re.search(
            r'diesel\s+power\s+products',
            text,
            re.IGNORECASE,
        ) else str(data.get('customer') or '').strip()
        _apply_stock_order_summary(
            data,
            description='STOCK ORDER',
            customer=stock_customer,
        )
        bill_no = str(data.get('invoice_number') or '').strip() or 'N/A'
        po_number = str(data.get('po_number') or '').strip() or 'N/A'
        cb(
            "  Fleece stocking-order invoice detected from footer note; outputting "
            "STOCK ORDER summary row "
            f"(Bill No: {bill_no}, Memo/PO: {po_number}; line items and totals suppressed).",
            "warning"
        )
    elif _is_holley_vendor_name(data.get('vendor', '')) and _matches_internal_stock_customer_hint(
        data.get('customer', '')
    ):
        _apply_stock_order_summary(
            data,
            description='STOCK ORDER',
            customer='Diesel Power Products',
        )
        bill_no = str(data.get('invoice_number') or '').strip() or 'N/A'
        po_number = str(data.get('po_number') or '').strip() or 'N/A'
        cb(
            "  Holley stock-order invoice detected from internal ship-to customer; "
            "outputting STOCK ORDER summary row "
            f"(Bill No: {bill_no}, Memo/PO: {po_number}; line items and totals suppressed).",
            "warning"
        )
    elif _is_fl_vendor_name(data.get('vendor', '')) and _matches_internal_stock_customer_hint(
        data.get('customer', '')
    ):
        _apply_stock_order_summary(
            data,
            description='STOCK ORDER',
            customer='Diesel Power Products',
        )
        bill_no = str(data.get('invoice_number') or '').strip() or 'N/A'
        po_number = str(data.get('po_number') or '').strip() or 'N/A'
        cb(
            "  Fleece stock-order invoice detected from internal ship-to customer; "
            "outputting STOCK ORDER summary row "
            f"(Bill No: {bill_no}, Memo/PO: {po_number}; line items and totals suppressed).",
            "warning"
        )

    # Ship-to address stock order / will call detection (any vendor)
    # If not already flagged as a stock order and ship-to is our warehouse address:
    #   - plain stock order if ship-to name is us or blank
    #   - will call if a different customer name appears in the ship-to block
    ship_to_lines = None
    if _is_mishimoto_vendor_name(data.get('vendor', '')):
        ship_to_lines = _extract_mishimoto_ship_to_lines(filepath)
    elif _is_power_stroke_vendor_name(data.get('vendor', '')):
        ship_to_lines = _extract_power_stroke_ship_to_lines(filepath)
    elif _is_redhead_vendor_name(data.get('vendor', '')):
        ship_to_lines = _extract_redhead_ship_to_lines(filepath)
    elif _is_hamilton_vendor_name(data.get('vendor', '')):
        ship_to_lines = _extract_hamilton_ship_to_lines(filepath)
    elif _is_daystar_vendor_name(data.get('vendor', '')):
        ship_to_lines = _extract_daystar_ship_to_lines(filepath)
    elif _is_bosch_vendor_name(data.get('vendor', '')):
        ship_to_lines = _extract_bosch_ship_to_lines(filepath)
    elif _is_diesel_forward_vendor_name(data.get('vendor', '')):
        ship_to_lines = _extract_diesel_forward_ship_to_lines(filepath)
    elif _is_carli_vendor_name(data.get('vendor', '')):
        ship_to_lines = _extract_carli_ship_to_lines(filepath)
    elif _is_icon_vendor_name(data.get('vendor', '')) or _is_cognito_vendor_name(data.get('vendor', '')):
        ship_to_lines = _extract_icon_cognito_ship_to_lines(filepath)
    elif _is_redhead_vendor_name(data.get('vendor', '')):
        ship_to_lines = _extract_redhead_ship_to_lines(filepath)
    elif _is_ats_vendor_name(data.get('vendor', '')):
        ship_to_lines = _extract_ats_ship_to_lines(filepath)
    elif _is_isspro_vendor_name(data.get('vendor', '')):
        ship_to_lines = _extract_isspro_ship_to_lines(filepath)
    elif _is_rock_krawler_vendor_name(data.get('vendor', '')):
        ship_to_lines = _extract_rock_krawler_ship_to_lines(filepath)
    elif _is_sport_truck_vendor_name(data.get('vendor', '')):
        ship_to_lines = _extract_sport_truck_ship_to_lines(filepath)
    elif _is_pt_vendor_name(data.get('vendor', '')):
        ship_to_lines = _extract_pt_ship_to_lines(filepath)
    elif _is_merchant_vendor_name(data.get('vendor', '')) or _is_kc_turbos_vendor_name(data.get('vendor', '')):
        ship_to_lines = _extract_ma_kt_ship_to_lines(filepath)
    elif _is_dynomite_vendor_name(data.get('vendor', '')):
        ship_to_lines = _extract_dd_ship_to_lines(filepath)
    elif _is_poly_vendor_name(data.get('vendor', '')):
        ship_to_lines = _extract_poly_ship_to_lines(filepath)
    elif _is_suspensionmaxx_vendor_name(data.get('vendor', '')):
        ship_to_lines = _extract_sm_ship_to_lines(filepath)
    elif _is_fumoto_vendor_name(data.get('vendor', '')):
        ship_to_lines = _extract_fumoto_ship_to_lines(filepath)

    ship_to_is_ours = (
        _ship_to_our_address_from_lines(ship_to_lines)
        if ship_to_lines is not None
        else _ship_to_our_address(text)
    )

    if (
        not data.get('stock_order')
        and ship_to_is_ours
        and _allow_global_ship_to_stock_detection(data.get('vendor', ''))
    ):
        will_call_customer = (
            _will_call_customer_from_lines(ship_to_lines)
            if ship_to_lines is not None
            else _will_call_customer_from_ship_to(text)
        )
        bill_no = str(data.get('invoice_number') or '').strip() or 'N/A'
        po_number = str(data.get('po_number') or '').strip() or 'N/A'
        if will_call_customer:
            _apply_stock_order_summary(
                data,
                description='WILL CALL',
                customer=will_call_customer,
            )
            cb(
                f"  Will call detected (ship-to is our address, customer: {will_call_customer}); "
                f"outputting WILL CALL summary row (Bill No: {bill_no}, Memo/PO: {po_number}).",
                "warning"
            )
        else:
            _apply_stock_order_summary(
                data,
                description='STOCK ORDER',
                customer='Diesel Power Products',
            )
            cb(
                "  Stock order detected (ship-to is our warehouse address); "
                f"outputting STOCK ORDER summary row (Bill No: {bill_no}, Memo/PO: {po_number}).",
                "warning"
            )

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
