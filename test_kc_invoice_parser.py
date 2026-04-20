import unittest
from unittest import mock

from invoice_parser import parse_invoice_text


class KCTurbosInvoiceParserTests(unittest.TestCase):
    def test_kc_turbos_po_number_accepts_two_digit_due_date_year(self):
        raw_text = (
            'KC Turbos\n'
            'Bill To\n'
            'Acme Diesel\n'
            'Total 125.00\n'
        )
        layout_text = (
            'KC Turbos\n'
            'Terms Due Date PO # Invoice #\n'
            'Net 30 3/15/26 0037993 INV12345\n'
        )

        with mock.patch('invoice_parser.extract_layout_text_from_pdf', return_value=layout_text):
            invoice_data = parse_invoice_text(raw_text, filepath='kc.pdf')

        self.assertEqual(invoice_data['vendor'], 'KC Turbos')
        self.assertEqual(invoice_data['po_number'], '0037993')


if __name__ == '__main__':
    unittest.main()
