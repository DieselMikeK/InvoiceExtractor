import unittest
from unittest import mock

from invoice_parser import parse_invoice


def _mock_pdf_context():
    mock_pdf = mock.MagicMock()
    mock_pdf.pages = [object()]
    mock_context = mock.MagicMock()
    mock_context.__enter__.return_value = mock_pdf
    mock_context.__exit__.return_value = False
    return mock_context


def _base_parsed_invoice(vendor):
    return {
        'invoice_number': 'I471734',
        'vendor': vendor,
        'vendor_address': '',
        'customer': 'Heather Bilek',
        'date': '4/23/2026',
        'due_date': '5/23/2026',
        'terms': 'Net 30',
        'po_number': '0061612',
        'tracking_number': '380753820778',
        'shipping_method': 'FedEx Ground',
        'ship_date': '4/23/2026',
        'shipping_tax_code': '',
        'shipping_tax_rate': '',
        'subtotal': '100.49',
        'shipping_cost': '0.00',
        'shipping_description': '',
        'total': '100.49',
        'line_items': [
            {
                'item_number': 'KF04060BK',
                'description': '08-16 F250 BDY MNT STEE KIT IN',
                'amount': '100.49',
            }
        ],
    }


class DaystarVendorResolutionTests(unittest.TestCase):
    def test_preserves_invoice_vendor_when_sender_alias_points_to_sb(self):
        invoice_text = (
            'Invoice # I471734\n'
            'PO # 0061612\n'
            'Please detach and return with your payment\n'
            'Daystar\n'
            'Invoice # I471734\n'
            '15461 Slover Avenue\n'
            'Fontana CA 92337\n'
        )

        with mock.patch('invoice_parser.pdfplumber.open', return_value=_mock_pdf_context()), \
             mock.patch('invoice_parser.extract_text_from_pdf', return_value=invoice_text), \
             mock.patch('invoice_parser.parse_invoice_text', return_value=_base_parsed_invoice('Daystar')), \
             mock.patch('invoice_parser._refresh_vendor_dependent_fields') as refresh_mock:
            data = parse_invoice(
                'C:\\temp\\I471734.pdf',
                sender_email='erika@sbfilters.com',
                sender_header='S&B Filters <erika@sbfilters.com>',
            )

        self.assertEqual(data.get('vendor'), 'Daystar')
        self.assertEqual(data.get('vendor_address'), '15461 Slover Avenue, Fontana CA 92337')
        self.assertEqual(data.get('invoice_number'), 'I471734')
        refresh_mock.assert_not_called()

    def test_sender_alias_can_correct_shared_address_vendor_when_invoice_names_sb(self):
        invoice_text = (
            'Invoice # I471734\n'
            'PO # 0061612\n'
            'S&B Filters\n'
            '15461 Slover Avenue\n'
            'Fontana CA 92337\n'
        )
        refreshed_parse = _base_parsed_invoice('S&B Filters')
        refreshed_parse['vendor_address'] = '15461 Slover Avenue, Fontana CA 92337'

        with mock.patch('invoice_parser.pdfplumber.open', return_value=_mock_pdf_context()), \
             mock.patch('invoice_parser.extract_text_from_pdf', return_value=invoice_text), \
             mock.patch('invoice_parser.parse_invoice_text', return_value=_base_parsed_invoice('Daystar')), \
             mock.patch(
                 'invoice_parser._refresh_vendor_dependent_fields',
                 return_value=refreshed_parse,
             ) as refresh_mock:
            data = parse_invoice(
                'C:\\temp\\I471734.pdf',
                sender_email='erika@sbfilters.com',
                sender_header='S&B Filters <erika@sbfilters.com>',
            )

        self.assertEqual(data.get('vendor'), 'S&B Filters')
        self.assertEqual(data.get('vendor_address'), '15461 Slover Avenue, Fontana CA 92337')
        refresh_mock.assert_called_once()


if __name__ == '__main__':
    unittest.main()
