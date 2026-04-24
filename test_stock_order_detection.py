import unittest
from unittest import mock

from invoice_parser import (
    _extract_label_value_pairs,
    _extract_redhead_ship_to_lines,
    _matches_internal_stock_customer_hint,
    _ship_to_our_address_from_lines,
    parse_invoice,
)


def _mock_pdf_context():
    mock_pdf = mock.MagicMock()
    mock_pdf.pages = [object()]
    mock_context = mock.MagicMock()
    mock_context.__enter__.return_value = mock_pdf
    mock_context.__exit__.return_value = False
    return mock_context


def _word(text, x0, top):
    return {'text': text, 'x0': x0, 'top': top}


class StockOrderDetectionTests(unittest.TestCase):
    def test_matches_company_name(self):
        self.assertTrue(_matches_internal_stock_customer_hint('Diesel Power Products'))

    def test_matches_warehouse_address_fragment(self):
        self.assertTrue(_matches_internal_stock_customer_hint('E. Main Ave.'))
        self.assertTrue(_matches_internal_stock_customer_hint('6200 East Main Avenue'))
        self.assertTrue(_matches_internal_stock_customer_hint('5204 E Broadway Ave'))

    def test_ignores_regular_customer_names(self):
        self.assertFalse(_matches_internal_stock_customer_hint('Thomas Shelton'))
        self.assertFalse(_matches_internal_stock_customer_hint('Lane Ricketson'))

    def test_redhead_ship_to_uses_right_hand_column(self):
        words = [
            _word('Bill', 75, 100),
            _word('To', 90, 100),
            _word('Ship', 340, 100),
            _word('To', 362, 100),
            _word('DIESEL', 66, 120),
            _word('POWER', 99, 120),
            _word('PRODUCTS', 133, 120),
            _word('ROBERT', 332, 120),
            _word('FRANCIS', 369, 120),
            _word('5204', 66, 132),
            _word('E', 87, 132),
            _word('BROADWAY', 94, 132),
            _word('AVE', 149, 132),
            _word('6566', 332, 132),
            _word('COUNTY', 352, 132),
            _word('HIGHWAY', 392, 132),
            _word('SPOKANE,', 66, 144),
            _word('WA', 112, 144),
            _word('99212-0904', 131, 144),
            _word('280', 332, 144),
            _word('E', 348, 144),
            _word('DEFUNIAK', 332, 156),
            _word('SPRINGS,', 380, 156),
            _word('FL', 421, 156),
            _word('32435', 434, 156),
            _word('P.O.', 358, 190),
            _word('No.', 379, 190),
            _word('Terms', 448, 190),
            _word('Ship', 523, 190),
            _word('Via', 544, 190),
        ]

        with mock.patch('invoice_parser._extract_first_page_words', return_value=(words, 612)):
            lines = _extract_redhead_ship_to_lines('invoice.pdf')

        self.assertEqual(
            lines,
            ['ROBERT FRANCIS', '6566 COUNTY HIGHWAY', '280 E', 'DEFUNIAK SPRINGS, FL 32435'],
        )
        self.assertFalse(_ship_to_our_address_from_lines(lines))

    def test_extracts_redhead_po_no_for_skunexus_lookup(self):
        fields = _extract_label_value_pairs(
            'P.O. No. Terms Ship Via\n'
            '0063208 N30 FEDEX\n'
        )

        self.assertEqual(fields.get('po_number'), '0063208')
        self.assertEqual(fields.get('terms'), 'N30')

    def test_redhead_drop_ship_keeps_line_items_for_po_validation(self):
        invoice_text = (
            'Red-Head Steering Gears, Inc. Invoice\n'
            'Date Invoice #\n'
            '4/23/2026 564994\n'
            'Bill To Ship To\n'
            'DIESEL POWER PRODUCTS ROBERT FRANCIS\n'
            '5204 E BROADWAY AVE 6566 COUNTY HIGHWAY\n'
            'SPOKANE, WA 99212-0904 280 E\n'
            'DEFUNIAK SPRINGS, FL 32435\n'
            'P.O. No. Terms Ship Via\n'
            '0063208 N30 FEDEX\n'
        )
        parsed_invoice = {
            'invoice_number': '564994',
            'vendor': 'Red-Head Steering Gears',
            'vendor_address': '',
            'customer': 'ROBERT FRANCIS',
            'date': '4/23/2026',
            'due_date': '',
            'terms': 'N30',
            'po_number': '0063208',
            'tracking_number': '',
            'shipping_method': 'FEDEX',
            'ship_date': '',
            'shipping_tax_code': '',
            'shipping_tax_rate': '',
            'subtotal': '',
            'shipping_cost': '',
            'shipping_description': '',
            'total': '697.50',
            'line_items': [
                {
                    'item_number': '2879',
                    'description': '03 - 08 FULL SIZE DODGE PICK-UP',
                    'quantity': '1',
                    'amount': '385.00',
                }
            ],
        }

        with mock.patch('invoice_parser.pdfplumber.open', return_value=_mock_pdf_context()), \
             mock.patch('invoice_parser.extract_text_from_pdf', return_value=invoice_text), \
             mock.patch('invoice_parser.parse_invoice_text', return_value=parsed_invoice), \
             mock.patch(
                 'invoice_parser._extract_redhead_ship_to_lines',
                 return_value=['ROBERT FRANCIS', '6566 COUNTY HIGHWAY', '280 E', 'DEFUNIAK SPRINGS, FL 32435'],
             ) as redhead_ship_to_mock:
            data = parse_invoice('C:\\temp\\Inv_564994_from_RedHead.pdf')

        redhead_ship_to_mock.assert_called_once()
        self.assertFalse(data.get('stock_order'))
        self.assertEqual(data.get('po_number'), '0063208')
        self.assertEqual(data.get('customer'), 'ROBERT FRANCIS')
        self.assertEqual(len(data.get('line_items') or []), 1)


if __name__ == '__main__':
    unittest.main()
