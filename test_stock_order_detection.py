import unittest
from unittest import mock

from invoice_parser import (
    _extract_carli_ship_to_lines,
    _extract_label_value_pairs,
    _extract_no_limit_ship_to_lines,
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
    def test_carli_ship_to_accepts_colon_header_and_detects_stock_pdf(self):
        path = (
            'training\\CRL\\'
            'INVOICE 122780 04_27_26 16_27_06 569.PDF'
        )
        lines = _extract_carli_ship_to_lines(path)

        self.assertEqual(
            lines,
            [
                'POWER PRODUCTS UNLIMITED, INC',
                '6200 E. Main Ave.',
                'Building 1 Suite A',
                'SPOKANE VALLEY WA99212',
                'UNITED STATES OF AMERICA',
            ],
        )
        self.assertTrue(_ship_to_our_address_from_lines(lines))

    def test_carli_stock_pdf_collapses_to_stock_order(self):
        data = parse_invoice(
            'training\\CRL\\'
            'INVOICE 122780 04_27_26 16_27_06 569.PDF'
        )

        self.assertEqual(data.get('vendor'), 'Carli Suspension - $10 DS Fee')
        self.assertEqual(data.get('invoice_number'), '122780')
        self.assertEqual(data.get('po_number'), '0060623')
        self.assertTrue(data.get('stock_order'))
        self.assertEqual(data.get('stock_order_description'), 'STOCK ORDER')
        self.assertEqual(data.get('customer'), 'Diesel Power Products')

    def test_carli_regular_drop_ship_pdfs_do_not_become_stock_orders(self):
        for path, customer in [
            (
                'training\\CRL\\INVOICE 120871 03_05_26 14_22_22 516.PDF',
                'Dominic Monego',
            ),
            (
                'training\\CRL\\INVOICE 120872 03_05_26 14_24_14 5.PDF',
                'Cade Grant',
            ),
        ]:
            with self.subTest(path=path):
                lines = _extract_carli_ship_to_lines(path)
                data = parse_invoice(path)

                self.assertIn(customer, lines)
                self.assertFalse(_ship_to_our_address_from_lines(lines))
                self.assertFalse(data.get('stock_order'))
                self.assertEqual(data.get('customer'), customer)

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

    def test_no_limit_drop_ship_does_not_mix_bill_to_into_will_call(self):
        invoice_text = (
            'in woainal Invoice\n'
            '5317 Bonsai Avenue 4/28/2026 78092\n'
            'Moorpark, CA 93021\n'
            'Power Products Unlimited/Diesel Power Tristen Wood\n'
            'Products 8854 Webster Rd\n'
            '5204 East Broadway Avenue Camden-on-Gauley, WV\n'
            'Spokane Valley, WA 99212\n'
            '0064998 Due Upon Receipt uo 4/28/2026\n'
            'EZL 100EEQOAA3 1 | EZ-Lynk Auto Agent 3 OBD Scan Tool 439.20 439.20\n'
        )
        parsed_invoice = {
            'invoice_number': '78092',
            'vendor': '',
            'vendor_address': '',
            'customer': 'Products 8854 Webster Rd',
            'date': '4/28/2026',
            'due_date': '',
            'terms': '',
            'po_number': '0064998',
            'line_items': [
                {
                    'item_number': 'EZL 100EEQOAA3',
                    'description': 'EZ-Lynk Auto Agent 3 OBD Scan Tool',
                    'quantity': '1',
                    'unit_price': '439.20',
                    'amount': '439.20',
                }
            ],
        }

        self.assertEqual(
            _extract_no_limit_ship_to_lines(invoice_text),
            ['Tristen Wood', '8854 Webster Rd', 'Camden-on-Gauley, WV'],
        )

        with mock.patch('invoice_parser.pdfplumber.open', return_value=_mock_pdf_context()), \
             mock.patch('invoice_parser.extract_text_from_pdf', return_value=invoice_text), \
             mock.patch('invoice_parser.parse_invoice_text', return_value=parsed_invoice):
            data = parse_invoice('C:\\temp\\training\\NL\\Invoice 78092.pdf')

        self.assertEqual(data.get('vendor'), 'No Limit Fabrication')
        self.assertFalse(data.get('stock_order'))
        self.assertEqual(data.get('customer'), 'Tristen Wood')
        self.assertEqual(len(data.get('line_items') or []), 1)

    def test_industrial_injection_stockorder_marker_collapses_to_stock_order(self):
        invoice_text = (
            'INVOICE\n'
            'I-431260\n'
            'INDUSTRIAL INJECTION SERVICE, INC. Date: 2026-05-01\n'
            'Bill To:Power Products Unlimited Ship To:Power Products Unlimited\n'
            '5204 E. Broadway Ave. 6200 E. Main Ave.\n'
            'Spokane, WA 99212 US Building 1 Suite A\n'
            'Spokane, WA 99212 US\n'
            'Date Ship Via Tracking Terms\n'
            '2026-05-01 EXTERNAL STOCKORDER OD 78081071454 NET30\n'
            'Purchase Order Number Order Date Sales Person Our Order\n'
            '63292 2026-04-27 18 S-ORD357409\n'
        )
        parsed_invoice = {
            'invoice_number': 'I-431260',
            'vendor': 'Industrial Injection',
            'vendor_address': '',
            'customer': 'Power Products Unlimited',
            'date': '2026-05-01',
            'due_date': '',
            'terms': 'NET30',
            'po_number': '63292',
            'tracking_number': '78081071454',
            'shipping_method': 'EXTERNAL STOCKORDER OD',
            'ship_date': '',
            'shipping_tax_code': '',
            'shipping_tax_rate': '',
            'subtotal': '',
            'shipping_cost': '',
            'shipping_description': '',
            'total': '32588.29',
            'line_items': [
                {
                    'item_number': '1464650366',
                    'description': '1989-1993 Cummins 12 Valve Governor Spring',
                    'quantity': '1',
                    'unit_price': '16.42',
                    'amount': '16.42',
                }
            ],
        }

        with mock.patch('invoice_parser.pdfplumber.open', return_value=_mock_pdf_context()), \
             mock.patch('invoice_parser.extract_text_from_pdf', return_value=invoice_text), \
             mock.patch('invoice_parser.parse_invoice_text', return_value=parsed_invoice):
            data = parse_invoice('C:\\temp\\I-431260.pdf')

        self.assertEqual(data.get('po_number'), '63292')
        self.assertTrue(data.get('stock_order'))
        self.assertEqual(data.get('stock_order_description'), 'STOCK ORDER')
        self.assertEqual(data.get('customer'), 'Diesel Power Products')
        self.assertEqual(data.get('line_items'), [])
        self.assertEqual(data.get('total'), '')


if __name__ == '__main__':
    unittest.main()
