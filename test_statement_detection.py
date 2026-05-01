import unittest

from invoice_parser import _is_statement_document, parse_invoice


class StatementDetectionTests(unittest.TestCase):
    def test_power_distributing_statement_is_not_invoice(self):
        data = parse_invoice(r'..\Invoices\invoices_5-1\Statement_10525.pdf')

        self.assertTrue(data.get('not_an_invoice'))

    def test_statement_table_text_is_not_invoice(self):
        text = (
            'Statement\n'
            'Power Distributing\n'
            'Customer Customer # Statement Date Page #\n'
            'POWER PRODUCTS UNLIMITED, INC. 10525 5/1/26 1 of 7\n'
            'Total Net Amount Due (USD) Amount Paid\n'
            'Please Return This Portion With Your Remittance\n'
            'Invoice Date Due Date Type Status Invoice # Customer PO # Invoice $s (IN) Credit $s (MC)\n'
            '4/1/26 5/10/26 IN DUE 386421-00 0055759 69.96\n'
        )

        self.assertTrue(_is_statement_document(text, 'Statement_10525.pdf'))

    def test_regular_invoice_wording_is_not_statement(self):
        text = (
            'Invoice\n'
            'Invoice Date Due Date Invoice # Customer PO #\n'
            '5/1/26 5/1/26 65734 0065760\n'
            'Item Details Purchases Inventory Item 1 511.96\n'
        )

        self.assertFalse(_is_statement_document(text, 'Invoice_65734.pdf'))


if __name__ == '__main__':
    unittest.main()
