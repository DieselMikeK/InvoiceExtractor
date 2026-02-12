# -*- coding: utf-8 -*-
"""SkuNexus API client for PO validation."""
import re
import requests


class SkuNexusClient:
    """Client for interacting with SkuNexus GraphQL API."""

    BASE_URL = "https://dpp.skunexus.com"

    def __init__(self, email, password):
        """Initialize the client with credentials."""
        self.email = email
        self.password = password
        self.session = requests.Session()
        self.session.headers.update({
            "Accept": "application/json",
            "Content-Type": "application/json",
            "X-Requested-With": "XMLHttpRequest",
        })
        self.logged_in = False

    def login(self):
        """Authenticate with SkuNexus and establish session."""
        try:
            resp = self.session.post(
                f"{self.BASE_URL}/api/users/login",
                json={"email": self.email, "password": self.password},
                timeout=30
            )
            if resp.status_code == 200:
                data = resp.json()
                if data.get('success'):
                    self.logged_in = True
                    return True, "Successfully logged in to SkuNexus"
                return False, data.get('message', 'Login failed')
            return False, f"Login failed with status {resp.status_code}"
        except requests.exceptions.Timeout:
            return False, "Login timed out"
        except Exception as e:
            return False, f"Login error: {str(e)}"

    def _query(self, query, timeout=60):
        """Execute a GraphQL query."""
        if not self.logged_in:
            return None, "Not logged in"

        try:
            resp = self.session.post(
                f"{self.BASE_URL}/api/query",
                json={"query": query},
                timeout=timeout
            )
            if resp.status_code == 200:
                data = resp.json()
                if 'errors' in data:
                    return None, data['errors'][0].get('message', 'GraphQL error')
                return data.get('data'), None
            return None, f"Query failed with status {resp.status_code}"
        except requests.exceptions.Timeout:
            return None, "Query timed out"
        except Exception as e:
            return None, f"Query error: {str(e)}"

    def search_po_candidates(self, po_number):
        """Search for PO candidates by number (full-text search).

        Args:
            po_number: The PO number to search for (e.g., "0036788")

        Returns:
            tuple: (list of candidate rows, error message or None)
        """
        po_number = _clean_po_number(po_number)
        if not po_number:
            return [], "PO number is empty"

        query = f"""
        query V1Queries {{
          purchaseOrder {{
            grid(
              filter: {{fulltext_search: "%{po_number}%"}}
              limit: {{size: 10, page: 1}}
            ) {{
              totalSize
              rows {{
                id
                label
                vendor {{ name id }}
                total_price
                items_count
                items_sum
                created_at
              }}
            }}
          }}
        }}
        """

        data, error = self._query(query)
        if error:
            return [], error

        rows = data.get('purchaseOrder', {}).get('grid', {}).get('rows', [])
        return rows, None

    def search_po(self, po_number):
        """Search for a PO by number with strict label matching.

        Args:
            po_number: The PO number to search for (e.g., "0036788")

        Returns:
            tuple: (po_data dict or None, error message or None)
        """
        po_number = _clean_po_number(po_number)
        rows, error = self.search_po_candidates(po_number)
        if error:
            return None, error

        # Find exact match for the PO number
        for row in rows:
            if row.get('label') == po_number:
                return row, None

        # Fallback: match by numeric normalization (handles leading zeros like 0037307 vs 37307)
        target_norm = _normalize_po(po_number)
        if target_norm:
            for row in rows:
                label = row.get('label', '')
                if _normalize_po(label) == target_norm:
                    return row, None

        return None, f"PO {po_number} not found"

    def get_po_details(self, po_id):
        """Get detailed PO information including line items.

        Args:
            po_id: The UUID of the PO

        Returns:
            tuple: (details dict or None, error message or None)
        """
        query = f"""
        query V1Queries {{
          purchaseOrder {{
            details(id: "{po_id}") {{
              id
              label
              total_price
              vendor {{ name id }}
              sourceAddress {{
                company
                street1
                street2
                city
                region
                postcode
                country
              }}
              lineItems(sort: {{}}, limit: {{size: 100, page: 1}}) {{
                totalSize
                rows {{
                  id
                  product {{ id name sku }}
                  quantity
                  price
                  total_price
                }}
              }}
            }}
          }}
        }}
        """

        data, error = self._query(query)
        if error:
            return None, error

        details = data.get('purchaseOrder', {}).get('details')
        return details, None

    def get_best_po_with_line_items(self, po_number, invoice_vendor='', invoice_skus=None, vendor_aliases=None):
        """Search for PO and get best-matching details including line items.

        This method tries strict PO label matching first. If multiple candidates
        are returned by full-text search, it uses vendor and SKU hints to select
        the best match instead of blindly picking the first row.

        Args:
            po_number: The PO number to search for
            invoice_vendor: Vendor name from the invoice (optional)
            invoice_skus: List of invoice SKUs (optional)
            vendor_aliases: Optional map of canonical vendor -> list of alias names

        Returns:
            tuple: (full_po_data dict or None, error message or None)
        """
        po_number = _clean_po_number(po_number)
        rows, error = self.search_po_candidates(po_number)
        if error:
            return None, error

        if not rows:
            return None, f"PO {po_number} not found"

        # 1) Prefer exact label match
        for row in rows:
            if row.get('label') == po_number:
                details, error = self.get_po_details(row['id'])
                if error:
                    return None, error
                return details, None

        # 2) Prefer normalized label match (leading zeros)
        target_norm = _normalize_po(po_number)
        if target_norm:
            for row in rows:
                label = row.get('label', '')
                if _normalize_po(label) == target_norm:
                    details, error = self.get_po_details(row['id'])
                    if error:
                        return None, error
                    return details, None

        # 3) Vendor hint (if it produces a single candidate)
        if invoice_vendor:
            alias_candidates = _get_vendor_aliases(invoice_vendor, vendor_aliases)
            vendor_matches = [
                row for row in rows
                if _vendors_match(invoice_vendor, row.get('vendor', {}).get('name', ''))
            ]
            if len(vendor_matches) == 1:
                details, error = self.get_po_details(vendor_matches[0]['id'])
                if error:
                    return None, error
                return details, None
            if len(vendor_matches) == 0 and alias_candidates:
                alias_matches = [
                    row for row in rows
                    if any(
                        _vendors_match(alias, row.get('vendor', {}).get('name', ''))
                        for alias in alias_candidates
                    )
                ]
                if len(alias_matches) == 1:
                    details, error = self.get_po_details(alias_matches[0]['id'])
                    if error:
                        return None, error
                    return details, None

        # 4) SKU hint across candidates (fetch details and score)
        invoice_skus = [s for s in (invoice_skus or []) if str(s).strip()]
        if invoice_skus:
            invoice_sku_norms = []
            seen = set()
            for sku in invoice_skus:
                norm = _normalize_sku(sku, invoice_vendor)
                if norm and norm not in seen:
                    seen.add(norm)
                    invoice_sku_norms.append(norm)

            best_details = None
            best_score = 0
            last_error = None

            for row in rows:
                details, error = self.get_po_details(row['id'])
                if error:
                    last_error = error
                    continue

                sn_vendor = details.get('vendor', {}).get('name', '')
                vendor_match = _vendors_match(invoice_vendor, sn_vendor)
                if not vendor_match:
                    for alias in _get_vendor_aliases(invoice_vendor, vendor_aliases):
                        if _vendors_match(alias, sn_vendor):
                            vendor_match = True
                            break

                sn_line_items = details.get('lineItems', {}).get('rows', [])
                sku_matches = 0
                for item in sn_line_items:
                    product = item.get('product', {})
                    sn_sku = product.get('sku', '')
                    sn_norm = _normalize_sku(sn_sku, invoice_vendor)
                    for inv_norm in invoice_sku_norms:
                        if inv_norm == sn_norm or inv_norm in sn_norm or sn_norm in inv_norm:
                            sku_matches += 1
                            break

                score = (sku_matches * 100) + (10 if vendor_match else 0)
                if score > best_score:
                    best_score = score
                    best_details = details

            if best_details and best_score > 0:
                return best_details, None

            if last_error:
                return None, last_error

        return None, f"PO {po_number} not found"

    def get_po_with_line_items(self, po_number):
        """Backwards-compatible wrapper (strict matching)."""
        return self.get_best_po_with_line_items(po_number)


def _clean_po_number(po_number):
    """Normalize PO number input for search."""
    po_number = str(po_number or '').strip()
    if po_number.upper().startswith('PO'):
        po_number = po_number[2:]
    return po_number.strip()


def _normalize_po(value):
    digits = ''.join(ch for ch in str(value) if ch.isdigit())
    if digits == '':
        return ''
    return digits.lstrip('0') or '0'


def _normalize_vendor_key(name):
    if not name:
        return ''
    s = name.lower().strip()
    s = s.replace('&', 'and')
    s = re.sub(r'[^a-z0-9]+', '', s)
    return s


def _vendors_match(invoice_vendor, skunexus_vendor):
    if not invoice_vendor or not skunexus_vendor:
        return False
    invoice_key = _normalize_vendor_key(invoice_vendor)
    sn_key = _normalize_vendor_key(skunexus_vendor)
    if not invoice_key or not sn_key:
        return False
    return invoice_key in sn_key or sn_key in invoice_key


def _normalize_product_service(value):
    s = str(value or '').strip().lower()
    s = re.sub(r'[^a-z0-9]+', '', s)
    return s


def _is_non_sku_product_service(value):
    key = _normalize_product_service(value)
    if not key:
        return False
    return key in {
        'core',
        'ere',
        'dppdiscount',
        'discount',
        'dropship',
        'shipping',
        'freight',
    }


def _get_vendor_aliases(name, vendor_aliases):
    if not name or not vendor_aliases:
        return []
    key = _normalize_vendor_key(name)
    aliases = vendor_aliases.get(key, [])
    if not isinstance(aliases, list):
        return []
    return [a for a in aliases if a]


def _normalize_sku(sku, vendor_name=''):
    """Normalize SKU for comparison. Vendor-specific tweaks allowed."""
    s = str(sku or '').strip().lower()
    # Remove all non-alphanumeric characters (spaces, dashes, underscores)
    s = re.sub(r'[^a-z0-9]+', '', s)

    vendor_key = _normalize_vendor_key(vendor_name)
    # No Limit often prepends vendor letters (NL / EZ / EZL) in SkuNexus or invoice
    if 'nolimit' in vendor_key:
        # Drop leading letters up to first digit
        s = re.sub(r'^[a-z]+', '', s)

    return s


def validate_po_row(skunexus_data, invoice_row, vendor_aliases=None):
    """Compare a single invoice row against SkuNexus PO data.

    Args:
        skunexus_data: Dict with PO details from SkuNexus (including lineItems)
        invoice_row: Dict with invoice data from spreadsheet

    Returns:
        tuple: (is_valid bool, list of failed field names)
    """
    failed_fields = []

    # Extract data from invoice row
    invoice_sku = str(invoice_row.get('sku', '')).strip()
    if not invoice_sku:
        invoice_sku = str(invoice_row.get('product_service', '')).strip()
    invoice_qty = invoice_row.get('qty', '')
    invoice_price = invoice_row.get('rate', '')  # Unit price
    invoice_amount = invoice_row.get('amount', '')
    invoice_vendor = str(invoice_row.get('vendor', '')).strip()
    invoice_description = str(invoice_row.get('description', '')).strip()

    # Skip shipping rows - they don't have SKUs to validate
    category = str(invoice_row.get('category', '')).strip()
    if category in ('Freight/Shipping', 'Freight and shipping costs'):
        return True, []

    # Skip non-SKU product/service labels (Core, E.R.E., Discount, Shipping, etc.)
    if _is_non_sku_product_service(invoice_sku):
        return True, []

    # Skip rows without product/service (continuation rows)
    if not invoice_sku:
        return True, []

    # Get SkuNexus data
    sn_vendor = skunexus_data.get('vendor', {}).get('name', '')
    sn_line_items = skunexus_data.get('lineItems', {}).get('rows', [])

    # Validate vendor (only on first row with vendor)
    if invoice_vendor:
        candidates = [invoice_vendor] + _get_vendor_aliases(invoice_vendor, vendor_aliases)
        vendor_ok = False
        for candidate in candidates:
            if not candidate:
                continue
            invoice_vendor_lower = candidate.lower().replace('&', 'and')
            sn_vendor_lower = sn_vendor.lower().replace('&', 'and')
            if invoice_vendor_lower in sn_vendor_lower or sn_vendor_lower in invoice_vendor_lower:
                vendor_ok = True
                break
            if 's&b' in candidate.lower() or 's & b' in candidate.lower():
                if 's&b' in sn_vendor.lower() or 's & b' in sn_vendor.lower():
                    vendor_ok = True
                    break
        if not vendor_ok:
            failed_fields.append('Vendor')

    # Find matching line item by SKU
    matching_item = None
    for item in sn_line_items:
        product = item.get('product', {})
        sn_sku = product.get('sku', '')

        # Normalize SKUs for comparison (remove separators; vendor-specific tweaks)
        invoice_sku_norm = _normalize_sku(invoice_sku, invoice_vendor)
        sn_sku_norm = _normalize_sku(sn_sku, invoice_vendor)

        if invoice_sku_norm == sn_sku_norm or invoice_sku_norm in sn_sku_norm or sn_sku_norm in invoice_sku_norm:
            matching_item = item
            break

    if not matching_item:
        failed_fields.append('SKU (not found)')
        return False, failed_fields

    # Validate quantity
    try:
        invoice_qty_num = float(str(invoice_qty).replace(',', ''))
        sn_qty_num = float(matching_item.get('quantity', 0))
        if abs(invoice_qty_num - sn_qty_num) > 0.01:
            failed_fields.append(f'Qty (invoice:{invoice_qty_num} vs SKN:{sn_qty_num})')
    except (ValueError, TypeError):
        if invoice_qty:
            failed_fields.append('Qty (parse error)')

    # Validate unit price
    try:
        invoice_price_num = float(str(invoice_price).replace(',', '').replace('$', ''))
        sn_price_num = float(matching_item.get('price', 0))
        # Allow small difference (rounding)
        if abs(invoice_price_num - sn_price_num) > 0.02:
            failed_fields.append(f'Price (invoice:{invoice_price_num} vs SKN:{sn_price_num})')
    except (ValueError, TypeError):
        if invoice_price:
            failed_fields.append('Price (parse error)')

    # Validate line total/amount
    try:
        invoice_amount_num = float(str(invoice_amount).replace(',', '').replace('$', ''))
        sn_total_num = float(str(matching_item.get('total_price', 0)).replace(',', ''))
        if abs(invoice_amount_num - sn_total_num) > 0.02:
            failed_fields.append(f'Amount (invoice:{invoice_amount_num} vs SKN:{sn_total_num})')
    except (ValueError, TypeError):
        if invoice_amount:
            failed_fields.append('Amount (parse error)')

    is_valid = len(failed_fields) == 0
    return is_valid, failed_fields
