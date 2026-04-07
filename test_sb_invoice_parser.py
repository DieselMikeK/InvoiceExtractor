import os
import unittest

from invoice_parser import infer_vendor_from_sender, parse_invoice


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
