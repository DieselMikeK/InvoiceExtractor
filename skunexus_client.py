# -*- coding: utf-8 -*-
"""SkuNexus API client for PO validation."""
from difflib import SequenceMatcher
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
              relatedOrder {{
                id
                label
              }}
              allRelatedOrders {{
                id
                label
                is_master
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

    def get_order_grouped_items(self, order_id):
        """Fetch grouped order items used by the related-order UI.

        Args:
            order_id: Related order UUID

        Returns:
            tuple: (order_details dict or None, error message or None)
        """
        query = f"""
        query V1Queries {{
          order {{
            details(id: "{order_id}") {{
              id
              label
              groupedDecisionItems {{
                qty
                relatedProduct {{
                  sku
                  customValues {{
                    custom_field_id
                    value
                  }}
                }}
                decisionItems {{
                  decidedItems {{
                    decisions {{
                      qty
                      relatedPurchaseOrder {{
                        label
                      }}
                    }}
                  }}
                }}
              }}
            }}
          }}
        }}
        """

        data, error = self._query(query)
        if error:
            return None, error

        details = data.get('order', {}).get('details')
        if not details:
            return None, "Related order not found"

        return details, None

    def get_po_margin(self, po_details, po_number=None):
        """Calculate PO margin using related-order item prices.

        Formula:
            margin = (related_item_price_sum - po_unit_sum) / related_item_price_sum

        Notes:
        - `po_unit_sum` uses PO line item unit prices (not line totals).
        - `related_item_price_sum` uses related order grouped item custom field
          `price` for rows mapped to this PO.
        """
        if not po_details:
            return None, "Missing PO details"

        po_label = str(po_details.get('label', '')).strip()
        target_po = _clean_po_number(po_number or po_label)
        target_norm = _normalize_po(target_po or po_label)
        if not target_norm:
            return None, "PO number is empty"

        # Sum PO "Unit" prices (per row) as requested.
        po_unit_sum = 0.0
        for line in po_details.get('lineItems', {}).get('rows', []) or []:
            price_val = _to_float(line.get('price'))
            if price_val is not None:
                po_unit_sum += price_val

        related_orders = po_details.get('allRelatedOrders') or []
        if not related_orders:
            single_related = po_details.get('relatedOrder')
            if single_related and single_related.get('id'):
                related_orders = [single_related]

        if not related_orders:
            return None, "No related order found"

        related_item_sum = 0.0
        counted_rows = 0

        for related in related_orders:
            order_id = related.get('id')
            if not order_id:
                continue

            order_details, error = self.get_order_grouped_items(order_id)
            if error:
                continue

            grouped_items = order_details.get('groupedDecisionItems') or []
            for grouped in grouped_items:
                if not _group_item_maps_to_po(grouped, target_norm):
                    continue

                item_price = _extract_custom_value(
                    grouped.get('relatedProduct', {}).get('customValues', []),
                    'price'
                )
                item_price_num = _to_float(item_price)
                if item_price_num is None:
                    continue

                related_item_sum += item_price_num
                counted_rows += 1

        if related_item_sum <= 0:
            return None, "No related order item prices found"

        margin = (related_item_sum - po_unit_sum) / related_item_sum
        return margin, None


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


def _to_float(value):
    if value is None:
        return None
    try:
        text = str(value).strip().replace(',', '').replace('$', '')
        if text == '':
            return None
        return float(text)
    except (ValueError, TypeError):
        return None


def _looks_like_line_amount(invoice_amount, invoice_qty, invoice_price):
    """Return True when an amount appears to be a per-line total (qty * rate).

    New exports store invoice-level total in the `amount` field (header: Total Amount)
    on the first line only. In that format, `amount` should not be validated as a
    line total against SkuNexus line item `total_price`.
    """
    amount_num = _to_float(invoice_amount)
    qty_num = _to_float(invoice_qty)
    price_num = _to_float(invoice_price)
    if amount_num is None or qty_num is None or price_num is None:
        return False

    expected = qty_num * price_num
    return abs(amount_num - expected) <= 0.05


def _extract_custom_value(custom_values, field_id):
    for cv in custom_values or []:
        if str(cv.get('custom_field_id', '')).strip().lower() == str(field_id).lower():
            return cv.get('value')
    return None


def _group_item_maps_to_po(grouped_item, target_po_norm):
    decision_items = grouped_item.get('decisionItems') or []
    for decision_item in decision_items:
        decided_items = decision_item.get('decidedItems') or []
        for decided in decided_items:
            decisions = decided.get('decisions') or []
            for decision in decisions:
                related_po_label = str(
                    (decision.get('relatedPurchaseOrder') or {}).get('label', '')
                ).strip()
                if _normalize_po(related_po_label) == target_po_norm:
                    return True
    return False


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


def _normalize_description(value):
    text = str(value or '').lower()
    text = re.sub(r'[^a-z0-9]+', ' ', text)
    return re.sub(r'\s+', ' ', text).strip()


def _description_tokens(value):
    text = _normalize_description(value)
    if not text:
        return set()

    stop_words = {
        'a',
        'an',
        'and',
        'for',
        'of',
        'or',
        's',
        'sb',
        'the',
        'with',
        'x',
    }
    tokens = set()
    for token in text.split():
        if token in stop_words:
            continue
        tokens.add(token)
        if len(token) == 4 and token.isdigit() and token.startswith(('19', '20')):
            tokens.add(token[2:])
    return tokens


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


def _description_similarity(left, right):
    left_norm = _normalize_description(left)
    right_norm = _normalize_description(right)
    if not left_norm or not right_norm:
        return 0.0

    sequence_score = SequenceMatcher(None, left_norm, right_norm).ratio()
    left_tokens = _description_tokens(left)
    right_tokens = _description_tokens(right)
    if not left_tokens or not right_tokens:
        return sequence_score

    overlap = left_tokens & right_tokens
    dice_score = (2 * len(overlap)) / (len(left_tokens) + len(right_tokens))
    invoice_coverage = len(overlap) / len(left_tokens)
    token_score = max(dice_score, invoice_coverage)
    score = max(sequence_score, token_score)

    left_has_hot = 'hot' in left_tokens
    left_has_cold = 'cold' in left_tokens
    right_has_hot = 'hot' in right_tokens
    right_has_cold = 'cold' in right_tokens
    if (left_has_hot and right_has_cold and not right_has_hot) or (
        left_has_cold and right_has_hot and not right_has_cold
    ):
        score -= 0.20

    return max(0.0, min(1.0, score))


def _prices_close(left, right):
    if left is None or right is None:
        return False
    if abs(left - right) <= 0.02:
        return True
    largest = max(abs(left), abs(right), 1.0)
    return (abs(left - right) / largest) <= 0.01


def match_invoice_row_to_po_line(skunexus_data, invoice_row, used_line_item_ids=None):
    """Find the best SkuNexus PO line for an invoice row without relying on SKU.

    This is used for S&B body invoices where the email has item name/price but no
    SKU. Product description/name is the primary signal; price and quantity only
    help rank otherwise similar candidates because either can legitimately differ
    from the vendor invoice.
    """
    if not skunexus_data or not invoice_row:
        return None

    used_line_item_ids = set(used_line_item_ids or [])
    invoice_qty = _to_float(invoice_row.get('qty'))
    invoice_price = _to_float(invoice_row.get('rate'))
    invoice_amount = _to_float(invoice_row.get('amount'))
    invoice_description = str(invoice_row.get('description') or '').strip()

    candidates = []
    for item in skunexus_data.get('lineItems', {}).get('rows', []) or []:
        item_id = str(item.get('id') or '').strip()
        if item_id and item_id in used_line_item_ids:
            continue

        product = item.get('product') or {}
        sn_sku = str(product.get('sku') or '').strip()
        sn_description = str(product.get('name') or '').strip()
        if not sn_sku:
            continue

        similarity = _description_similarity(invoice_description, sn_description)
        score = similarity * 100
        reasons = []
        sn_price = _to_float(item.get('price'))
        sn_qty = _to_float(item.get('quantity'))
        sn_total = _to_float(item.get('total_price'))

        if invoice_price is not None and sn_price is not None and _prices_close(invoice_price, sn_price):
            score += 8
            reasons.append('price')

        if invoice_qty is not None and sn_qty is not None and abs(invoice_qty - sn_qty) <= 0.01:
            score += 3
            reasons.append('qty')

        if (
            invoice_amount is not None
            and invoice_qty is not None
            and invoice_price is not None
            and abs(invoice_amount - (invoice_qty * invoice_price)) <= 0.05
            and sn_total is not None
            and abs(invoice_amount - sn_total) <= 0.05
        ):
            score += 4
            reasons.append('amount')

        if similarity >= 0.55:
            reasons.append('description')

        candidates.append((score, similarity, item, reasons))

    if not candidates:
        return None

    candidates.sort(key=lambda value: (value[0], value[1]), reverse=True)
    best_score, best_similarity, best_item, best_reasons = candidates[0]
    second_score = candidates[1][0] if len(candidates) > 1 else -1.0
    second_similarity = candidates[1][1] if len(candidates) > 1 else -1.0

    # SKU inference is intentionally description-led. Validation still runs
    # afterward and can fail price or quantity, but those failures should not
    # prevent the SKU from carrying into the sheet.
    if best_similarity < 0.50:
        return None
    if len(candidates) > 1 and (best_score - second_score) < 6 and (best_similarity - second_similarity) < 0.08:
        return None
    return best_item


def infer_invoice_row_sku_from_po(skunexus_data, invoice_row, used_line_item_ids=None):
    """Infer an invoice row SKU from SkuNexus PO line details."""
    item = match_invoice_row_to_po_line(skunexus_data, invoice_row, used_line_item_ids)
    if not item:
        return '', ''
    product = item.get('product') or {}
    sku = str(product.get('sku') or '').strip()
    vendor = str((skunexus_data.get('vendor') or {}).get('name') or '').strip()
    if _vendors_match(vendor, 'S&B Filters') or sku.upper().startswith('SB-'):
        sku = re.sub(r'(?i)^SB-', '', sku).strip()
    return sku, str(item.get('id') or '').strip()


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

    # Validate line total/amount only when the invoice value looks like a true
    # per-line amount (qty * rate). This avoids false failures when the column
    # stores invoice-level "Total Amount" on the first row.
    if _looks_like_line_amount(invoice_amount, invoice_qty, invoice_price):
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
