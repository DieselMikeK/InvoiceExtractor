# -*- coding: utf-8 -*-
"""SkuNexus API client for PO validation."""
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

    def search_po(self, po_number):
        """Search for a PO by number.

        Args:
            po_number: The PO number to search for (e.g., "0036788")

        Returns:
            tuple: (po_data dict or None, error message or None)
        """
        # Clean up PO number - remove "PO" prefix if present
        po_number = str(po_number).strip()
        if po_number.upper().startswith('PO'):
            po_number = po_number[2:]

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
            return None, error

        rows = data.get('purchaseOrder', {}).get('grid', {}).get('rows', [])

        # Find exact match for the PO number
        for row in rows:
            if row.get('label') == po_number:
                return row, None

        # If no exact match, return first result if any
        if rows:
            return rows[0], None

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

    def get_po_with_line_items(self, po_number):
        """Search for PO and get full details including line items.

        Args:
            po_number: The PO number to search for

        Returns:
            tuple: (full_po_data dict or None, error message or None)
        """
        # First search for the PO to get its ID
        po_data, error = self.search_po(po_number)
        if error:
            return None, error

        if not po_data:
            return None, f"PO {po_number} not found"

        # Get full details with line items
        details, error = self.get_po_details(po_data['id'])
        if error:
            return None, error

        return details, None


def validate_po_row(skunexus_data, invoice_row):
    """Compare a single invoice row against SkuNexus PO data.

    Args:
        skunexus_data: Dict with PO details from SkuNexus (including lineItems)
        invoice_row: Dict with invoice data from spreadsheet

    Returns:
        tuple: (is_valid bool, list of failed field names)
    """
    failed_fields = []

    # Extract data from invoice row
    invoice_sku = str(invoice_row.get('product_service', '')).strip()
    invoice_qty = invoice_row.get('qty', '')
    invoice_price = invoice_row.get('rate', '')  # Unit price
    invoice_amount = invoice_row.get('amount', '')
    invoice_vendor = str(invoice_row.get('vendor', '')).strip()
    invoice_description = str(invoice_row.get('description', '')).strip()

    # Skip shipping rows - they don't have SKUs to validate
    if invoice_row.get('category') == 'Freight/Shipping':
        return True, []

    # Skip rows without product/service (continuation rows)
    if not invoice_sku:
        return True, []

    # Get SkuNexus data
    sn_vendor = skunexus_data.get('vendor', {}).get('name', '')
    sn_line_items = skunexus_data.get('lineItems', {}).get('rows', [])

    # Validate vendor (only on first row with vendor)
    if invoice_vendor:
        # Normalize vendor names for comparison
        invoice_vendor_lower = invoice_vendor.lower().replace('&', 'and')
        sn_vendor_lower = sn_vendor.lower().replace('&', 'and')

        # Check if vendor names match (partial match is OK)
        if invoice_vendor_lower not in sn_vendor_lower and sn_vendor_lower not in invoice_vendor_lower:
            # Try matching key parts
            if 's&b' in invoice_vendor.lower() or 's & b' in invoice_vendor.lower():
                if 's&b' not in sn_vendor.lower() and 's & b' not in sn_vendor.lower():
                    failed_fields.append('Vendor')
            else:
                failed_fields.append('Vendor')

    # Find matching line item by SKU
    matching_item = None
    for item in sn_line_items:
        product = item.get('product', {})
        sn_sku = product.get('sku', '')

        # Normalize SKUs for comparison (remove common prefixes, lowercase)
        invoice_sku_norm = invoice_sku.lower().replace('-', '').replace('_', '')
        sn_sku_norm = sn_sku.lower().replace('-', '').replace('_', '')

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
