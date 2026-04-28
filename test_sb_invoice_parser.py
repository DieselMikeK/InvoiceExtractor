import os
import json
import tempfile
import unittest
from unittest import mock

from invoice_parser import (
    _extract_sb_shopify_line_items_from_json,
    _extract_sb_shopify_order_page_data,
    _parse_sb_shopify_order_url,
    infer_vendor_from_sender,
    parse_email_invoice,
    parse_invoice,
)


SB_BODY = """---------- Forwarded message ---------
From: S&B <store+69841617189@t.shopifyemail.com>
Date: Mon, Apr 27, 2026 at 2:30 PM
Subject: Order #743234 Confirmed
To: <ap@dieselpowerproducts.com>

ORDER #743234
PO NUMBER #0064464
Hi Power Products. Thank you for your purchase!
Order summary
Cold Air Intake for 2006-2007 Chevy / GMC Duramax LLY-LBZ 6.6L x 1 $253.93
Dry Extendable
Subtotal $253.93
Shipping $12.00
Taxes $0.00
Total paid today $0.00 USD
Total due May 27, 2026 $265.93 USD
Customer information
Shipping address
Donald Ortiz
5041 Brighton Hills Dr NE
Rio Rancho NM 87144
Billing address
Josh Ulrich
Diesel Power Products DBA Power Products Unlimited, Inc. 505
5204 East Broadway Avenue
Spokane Valley WA 99212
Payment
Net 30: Due May 27, 2026
Shipping method
Ground
If you have any questions, reply to this email or contact us at
customerservice@sbfilters.com
"""


class SBInvoiceParserTests(unittest.TestCase):
    def test_sb_sender_aliases_match_sbfilters_domain(self):
        cases = [
            ('erika@sbfilters.com', ''),
            ('', 'Matt <matt@sbfilters.com>'),
            ('orders@sbfilters.com', 'S&B Filters <orders@sbfilters.com>'),
        ]

        for sender_email, sender_header in cases:
            with self.subTest(sender_email=sender_email, sender_header=sender_header):
                self.assertEqual(
                    infer_vendor_from_sender(sender_email=sender_email, sender_header=sender_header),
                    'S&B Filters',
                )

    def test_kc_sender_aliases_match_kcturbos_domain(self):
        cases = [
            ('invoicing@kcturbos.com', ''),
            ('orders@kcturbos.com', ''),
            ('', 'KC Turbos <billing@kcturbos.com>'),
        ]

        for sender_email, sender_header in cases:
            with self.subTest(sender_email=sender_email, sender_header=sender_header):
                self.assertEqual(
                    infer_vendor_from_sender(sender_email=sender_email, sender_header=sender_header),
                    'KC Turbos',
                )

    def test_sb_shopify_body_invoice_parser(self):
        payload = {
            'type': 'email_body_invoice',
            'parser': 'sb_shopify_order',
            'subject': 'Order #743234 Confirmed',
            'message_text': SB_BODY,
        }

        with tempfile.TemporaryDirectory() as tmpdir:
            path = os.path.join(tmpdir, 'SB_Order_743234.email.json')
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(payload, f)

            invoice_data = parse_email_invoice(path)

        self.assertEqual(invoice_data['vendor'], 'S&B Filters')
        self.assertEqual(invoice_data['invoice_number'], '743234')
        self.assertEqual(invoice_data['po_number'], '0064464')
        self.assertEqual(invoice_data['date'], '4/27/2026')
        self.assertEqual(invoice_data['due_date'], '5/27/2026')
        self.assertEqual(invoice_data['customer'], 'Donald Ortiz')
        self.assertEqual(invoice_data['shipping_method'], 'Ground')
        self.assertEqual(invoice_data['subtotal'], '253.93')
        self.assertEqual(invoice_data['shipping_cost'], '12.00')
        self.assertEqual(invoice_data['total'], '265.93')
        self.assertEqual(len(invoice_data['line_items']), 1)
        self.assertEqual(invoice_data['line_items'][0]['quantity'], '1')
        self.assertEqual(invoice_data['line_items'][0]['unit_price'], '253.93')
        self.assertIn('Cold Air Intake', invoice_data['line_items'][0]['description'])
        self.assertIn('Dry Extendable', invoice_data['line_items'][0]['description'])

    def test_sb_shopify_body_invoice_uses_product_line_before_option_price(self):
        body = """Order summary
S&B Intake Replacement Filter x 1
Cotton Cleanable $46.23
Subtotal $46.23
Shipping $0.00
Taxes $0.00
Total due May 27, 2026 $46.23 USD
Customer information
Shipping address
Test Customer
1 Main St
Billing address
Diesel Power Products
Payment
Net 30: Due May 27, 2026
Shipping method
Ground
"""
        payload = {
            'type': 'email_body_invoice',
            'parser': 'sb_shopify_order',
            'subject': 'Order #743235 Confirmed',
            'message_text': body,
        }

        with tempfile.TemporaryDirectory() as tmpdir:
            path = os.path.join(tmpdir, 'SB_Order_743235.email.json')
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(payload, f)

            invoice_data = parse_email_invoice(path)

        self.assertEqual(len(invoice_data['line_items']), 1)
        item = invoice_data['line_items'][0]
        self.assertEqual(item['quantity'], '1')
        self.assertEqual(item['unit_price'], '46.23')
        self.assertEqual(item['description'], 'S&B Intake Replacement Filter x 1 - Cotton Cleanable')

    def test_sb_shopify_body_invoice_does_not_parse_model_year_as_price(self):
        body = """Order summary
Cold Air Intake for 2018-2021 Ford F-150 Powerstroke 3.0L x 1
Cotton Cleanable

$253.93

Subtotal
$253.93
Shipping
$12.00
Taxes
$0.00
Total due May 28, 2026
$265.93 USD
Customer information
Shipping address
Omar Calderon
1300 Marden Rd
Billing address
Diesel Power Products
Payment
Net 30: Due May 28, 2026
Shipping method
Ground
"""
        payload = {
            'type': 'email_body_invoice',
            'parser': 'sb_shopify_order',
            'subject': 'Order #743624 Confirmed',
            'message_text': body,
        }

        with tempfile.TemporaryDirectory() as tmpdir:
            path = os.path.join(tmpdir, 'SB_Order_743624.email.json')
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(payload, f)

            invoice_data = parse_email_invoice(path)

        self.assertEqual(len(invoice_data['line_items']), 1)
        item = invoice_data['line_items'][0]
        self.assertEqual(item['quantity'], '1')
        self.assertEqual(item['unit_price'], '253.93')
        self.assertEqual(item['description'], (
            'Cold Air Intake for 2018-2021 Ford F-150 Powerstroke 3.0L '
            'x 1 - Cotton Cleanable'
        ))

    def test_sb_shopify_order_page_text_extracts_skus(self):
        page_text = """Order #743636
PO #0064810
1
Hot Side Intercooler Pipe for 2016-2026 Ford Powerstroke 6.7L
SKU: 83-2004
$166.83
1
Cold Side Intercooler Pipe for 2017-2026 Ford Super Duty, 6.7L Powerstroke
SKU: 83-1001
$200.33
Subtotal · 2 items
$367.16
Shipping
$19.50
Due May 28
USD $386.66
"""
        data = _extract_sb_shopify_order_page_data(page_text)

        self.assertEqual(data['invoice_number'], '743636')
        self.assertEqual(data['po_number'], '0064810')
        self.assertEqual(data['shipping_cost'], '19.50')
        self.assertEqual(data['total'], '386.66')
        self.assertEqual(len(data['line_items']), 2)
        self.assertEqual(data['line_items'][0]['item_number'], '83-2004')
        self.assertEqual(data['line_items'][0]['unit_price'], '166.83')
        self.assertIn('Hot Side Intercooler Pipe', data['line_items'][0]['description'])
        self.assertEqual(data['line_items'][1]['item_number'], '83-1001')

    def test_sb_shopify_order_url_uses_rendered_body_text(self):
        page_text = """Order #743636
PO #0064810
1
Hot Side Intercooler Pipe for 2016-2026 Ford Powerstroke 6.7L
SKU: 83-2004
$166.83
Subtotal Â· 1 item
$166.83
Shipping
$19.50
Due May 28
USD $186.33
"""

        with mock.patch(
            'invoice_parser._fetch_rendered_order_context',
            return_value={'text': page_text, 'line_items': []},
        ):
            data = _parse_sb_shopify_order_url('https://sbfilters.com/order/authenticate?key=abc')

        self.assertEqual(data['line_items'][0]['item_number'], '83-2004')
        self.assertEqual(data['total'], '186.33')

    def test_sb_shopify_line_items_parse_from_graphql_payload(self):
        payload = {
            'data': {
                'order': {
                    'lineItemContainers': [
                        {
                            'lineItems': {
                                'nodes': [
                                    {
                                        'lineItem': {
                                            'sku': '83-2004',
                                            'title': 'Hot Side Intercooler Pipe',
                                        }
                                    },
                                    {
                                        'lineItem': {
                                            'sku': '83-2004',
                                            'title': 'Hot Side Intercooler Pipe',
                                            'quantity': 1,
                                            'price': {'amount': '166.83'},
                                            'currentTotalPrice': {'amount': '166.83'},
                                        }
                                    },
                                    {
                                        'lineItem': {
                                            'sku': '83-1001',
                                            'title': 'Cold Side Intercooler Pipe',
                                            'quantity': 1,
                                            'price': {'amount': '200.33'},
                                            'currentTotalPrice': {'amount': '200.33'},
                                        }
                                    },
                                ]
                            }
                        }
                    ]
                }
            }
        }

        items = _extract_sb_shopify_line_items_from_json(payload)

        self.assertEqual(len(items), 2)
        self.assertEqual(items[0]['item_number'], '83-2004')
        self.assertEqual(items[0]['unit_price'], '166.83')
        self.assertEqual(items[1]['item_number'], '83-1001')

    def test_sb_body_invoice_prefers_shopify_order_page_items_and_keeps_url(self):
        order_url = 'https://shopify.com/69841617189/account/orders/token?locale=en-US'
        payload = {
            'type': 'email_body_invoice',
            'parser': 'sb_shopify_order',
            'subject': 'Order #743636 Confirmed',
            'order_url': order_url,
            'message_text': """Order #743636
PO number #0064810
Order summary
Fallback Item x 1
Cotton Cleanable
$1.00
Subtotal $1.00
Shipping $0.00
Taxes $0.00
Total due May 28, 2026 $1.00 USD
Customer information
Shipping address
Test Customer
Billing address
Diesel Power Products
Payment
Net 30: Due May 28, 2026
Shipping method
Ground
""",
        }
        page_data = {
            'line_items': [
                {
                    'item_number': '83-2004',
                    'quantity': '1',
                    'units': 'Each',
                    'description': 'Hot Side Intercooler Pipe',
                    'unit_price': '166.83',
                    'amount': '166.83',
                }
            ],
            'subtotal': '166.83',
            'shipping_cost': '19.50',
            'total': '186.33',
        }

        with tempfile.TemporaryDirectory() as tmpdir:
            path = os.path.join(tmpdir, 'SB_Order_743636.email.json')
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(payload, f)

            with mock.patch('invoice_parser._parse_sb_shopify_order_url', return_value=page_data):
                invoice_data = parse_email_invoice(path)

        self.assertEqual(invoice_data['source_url'], order_url)
        self.assertEqual(invoice_data['line_items'][0]['item_number'], '83-2004')
        self.assertEqual(invoice_data['line_items'][0]['unit_price'], '166.83')
        self.assertEqual(invoice_data['shipping_cost'], '19.50')

    def test_sb_shipping_cost_handles_nested_parentheses(self):
        cases = {
            'S&B I464016.pdf': '42.00',
            'S&B I464035.pdf': '55.00',
            'S&B I464039.pdf': '36.00',
        }
        training_dir = os.path.join(os.path.dirname(__file__), 'training', 'SB')

        for filename, expected_shipping in cases.items():
            with self.subTest(filename=filename):
                invoice_data = parse_invoice(os.path.join(training_dir, filename))
                self.assertEqual(invoice_data['shipping_cost'], expected_shipping)
                self.assertEqual(invoice_data['shipping_description'], 'Shipping')

    def test_sb_total_tax_line_does_not_break_shipping_or_total(self):
        pdf_path = os.path.join(
            os.path.dirname(__file__),
            'training',
            'SB',
            'S&B I464317.pdf',
        )

        invoice_data = parse_invoice(pdf_path)

        self.assertEqual(invoice_data['subtotal'], '46.23')
        self.assertEqual(invoice_data['shipping_cost'], '84.00')
        self.assertEqual(invoice_data['shipping_description'], 'Shipping')
        self.assertEqual(invoice_data['total'], '130.88')
    def test_sb_old_template_strips_coop_rasmussen_from_customer(self):
        pdf_path = os.path.join(
            os.path.dirname(__file__),
            'training',
            'SB',
            'S&B I464016.pdf',
        )

        invoice_data = parse_invoice(pdf_path)

        self.assertEqual(invoice_data['customer'], 'Keon Evans')

    def test_sb_new_template_uses_customer_left_column_and_default_address(self):
        cases = {
            'I468610.pdf': {
                'customer': 'Neil West',
                'shipping_cost': '12.00',
                'total': '265.93',
            },
            'I468613.pdf': {
                'customer': 'Bob Bivans',
                'shipping_cost': '19.50',
                'total': '293.50',
            },
            'I468616.pdf': {
                'customer': 'Hunter McMasters',
                'shipping_cost': '12.00',
                'total': '265.93',
            },
            'I468620.pdf': {
                'customer': 'Brandon Martinez',
                'shipping_cost': '12.00',
                'total': '265.93',
            },
        }
        training_dir = os.path.join(os.path.dirname(__file__), 'training', 'SB')

        for filename, expected in cases.items():
            with self.subTest(filename=filename):
                invoice_data = parse_invoice(os.path.join(training_dir, filename))
                self.assertEqual(invoice_data['vendor'], 'S&B Filters')
                self.assertEqual(invoice_data['customer'], expected['customer'])
                self.assertEqual(invoice_data['vendor_address'], '15461 Slover Avenue, Fontana CA 92337')
                self.assertEqual(invoice_data['shipping_cost'], expected['shipping_cost'])
                self.assertEqual(invoice_data['shipping_description'], 'Shipping')
                self.assertEqual(invoice_data['total'], expected['total'])

if __name__ == '__main__':
    unittest.main()
