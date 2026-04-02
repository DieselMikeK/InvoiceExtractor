import os
import unittest

from invoice_parser import parse_invoice


class SBInvoiceParserTests(unittest.TestCase):
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


if __name__ == '__main__':
    unittest.main()
