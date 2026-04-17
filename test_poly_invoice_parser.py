import unittest
from unittest import mock

from invoice_parser import (
    _extract_poly_ship_to_lines,
    _ship_to_our_address,
    _will_call_customer_from_ship_to,
    parse_invoice,
)


def _word(text, x0, top):
    return {'text': text, 'x0': x0, 'top': top}


class PolyInvoiceParserTests(unittest.TestCase):
    def test_extract_poly_ship_to_lines_handles_ship_to_bill_to_layout(self):
        words = [
            _word('Ship', 60, 100),
            _word('To', 84, 100),
            _word('Bill', 300, 100),
            _word('To', 326, 100),
            _word('Charles', 60, 118),
            _word('Olson', 104, 118),
            _word('Diesel', 300, 118),
            _word('Power', 344, 118),
            _word('Products', 388, 118),
            _word('88570', 60, 130),
            _word('Young', 102, 130),
            _word('Drive', 146, 130),
            _word('5204', 300, 130),
            _word('E', 334, 130),
            _word('BROADWAY', 348, 130),
            _word('AVE', 418, 130),
            _word('Valentine', 60, 142),
            _word('NE', 118, 142),
            _word('69201', 140, 142),
            _word('SPOKANE', 300, 142),
            _word('VALLEY', 356, 142),
            _word('WA', 406, 142),
            _word('99212', 426, 142),
            _word('Shipping', 60, 154),
            _word('Method', 118, 154),
            _word('Ground', 164, 154),
            _word('Item', 60, 176),
            _word('Customer', 116, 176),
            _word('Quantity', 198, 176),
            _word('Units', 248, 176),
            _word('Description', 300, 176),
        ]

        with mock.patch('invoice_parser._extract_first_page_words', return_value=(words, 612)):
            ship_to_lines = _extract_poly_ship_to_lines('poly.pdf')

        self.assertEqual(
            ship_to_lines,
            ['Charles Olson', '88570 Young Drive', 'Valentine NE 69201'],
        )

    def test_parse_invoice_does_not_false_positive_poly_as_will_call(self):
        raw_text = (
            'Date 4/16/2026\n'
            'Invoice # IN471736\n'
            'Terms Net 30\n'
            'Due Date 5/16/2026\n'
            'PO # 0060786\n'
            'Ship To Bill To\n'
            'Charles Olson Diesel Power Products\n'
            '5204 E BROADWAY AVE SPOKANE VALLEY WA 99212\n'
            'Shipping Method Ground CUST UPS ACCT\n'
            '88570 Young Drive\n'
            'Valentine NE 69201\n'
            'Item Customer P... Quantity Units Description Unit Price Amount\n'
            'KC09101BK 1 Each SPACER,FRONT,2\",94-10 25/3500 $46.23 $46.23\n'
            'Subtotal $46.23\n'
            'Total $46.23\n'
        )
        parsed_data = {
            'invoice_number': 'IN471736',
            'vendor': 'Poly Performance',
            'vendor_address': '',
            'customer': 'Charles Olson',
            'date': '4/16/2026',
            'due_date': '5/16/2026',
            'terms': 'Net 30',
            'po_number': '0060786',
            'tracking_number': '',
            'shipping_method': 'Ground CUST UPS ACCT',
            'ship_date': '4/16/2026',
            'shipping_tax_code': '',
            'shipping_tax_rate': '',
            'subtotal': '46.23',
            'shipping_cost': '',
            'shipping_description': '',
            'total': '46.23',
            'line_items': [
                {
                    'item_number': 'KC09101BK',
                    'quantity': '1',
                    'units': 'Each',
                    'description': 'SPACER,FRONT,2",94-10 25/3500',
                    'unit_price': '46.23',
                    'amount': '46.23',
                }
            ],
        }

        self.assertTrue(_ship_to_our_address(raw_text))
        self.assertEqual(
            _will_call_customer_from_ship_to(raw_text),
            'Shipping Method Ground CUST UPS ACCT',
        )

        mock_pdf = mock.MagicMock()
        mock_pdf.pages = [object()]
        mock_context = mock.MagicMock()
        mock_context.__enter__.return_value = mock_pdf
        mock_context.__exit__.return_value = False

        with mock.patch('invoice_parser.pdfplumber.open', return_value=mock_context), \
             mock.patch('invoice_parser.extract_text_from_pdf', return_value=raw_text), \
             mock.patch('invoice_parser.parse_invoice_text', return_value=parsed_data.copy()), \
             mock.patch(
                 'invoice_parser._extract_poly_ship_to_lines',
                 return_value=['Charles Olson', '88570 Young Drive', 'Valentine NE 69201'],
             ):
            data = parse_invoice('C:\\temp\\poly-performance.pdf', lambda *args, **kwargs: None)

        self.assertFalse(data.get('stock_order'))
        self.assertEqual(data.get('customer'), 'Charles Olson')
        self.assertEqual(data.get('total'), '46.23')
        self.assertEqual(len(data.get('line_items') or []), 1)
        self.assertEqual(data['line_items'][0].get('item_number'), 'KC09101BK')


if __name__ == '__main__':
    unittest.main()
