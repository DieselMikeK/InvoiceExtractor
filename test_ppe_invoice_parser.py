import os
import unittest

from invoice_parser import parse_invoice


APP_DIR = os.path.dirname(os.path.abspath(__file__))


class PPEInvoiceParserTests(unittest.TestCase):
    def test_ppe_letterhead_wins_over_cognito_ship_to_address_alias(self):
        invoice_path = os.path.join(
            APP_DIR,
            'ringping_attachments',
            'request-37',
            'Sales Invoice P-INV270264.pdf',
        )

        data = parse_invoice(invoice_path)

        self.assertEqual(data.get('vendor'), 'Pacific Performance Engineering - $5 DS Fee')
        self.assertEqual(data.get('invoice_number'), 'P-INV270264')
        self.assertEqual(data.get('po_number'), '0065929')
        self.assertEqual(data.get('customer'), 'Ray Trevino')
        self.assertEqual(data.get('terms'), 'Due on receipt')
        self.assertEqual(data.get('total'), '364.99')


if __name__ == '__main__':
    unittest.main()
