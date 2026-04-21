import os
import unittest
from unittest import mock

from invoice_parser import infer_vendor_from_email_metadata, infer_vendor_from_sender, parse_invoice


APP_DIR = os.path.dirname(os.path.abspath(__file__))
TRAINING_DIR = os.path.join(APP_DIR, 'training')


class NewVendorFirstPassTests(unittest.TestCase):
    def test_sender_aliases_cover_carli_hamilton_and_icon(self):
        cases = [
            ('sales@carlisuspension.com', '', 'Carli Suspension - $10 DS Fee'),
            ('hamiltoncamsales@gmail.com', '', 'Hamilton Cams - $20 Dropship Fee'),
            ('', 'ICON Vehicle Dynamics <orders@iconvehicledynamics.com>', 'Icon Vehicle Dynamics'),
        ]

        for sender_email, sender_header, expected_vendor in cases:
            with self.subTest(sender_email=sender_email, sender_header=sender_header):
                self.assertEqual(
                    infer_vendor_from_sender(sender_email=sender_email, sender_header=sender_header),
                    expected_vendor,
                )

    def test_carli_shared_sender_domain_requires_subject_confirmation(self):
        carli_subject = (
            'Fwd: ***DO NOT REPLY***Carli Suspension Invoice(s) Attached '
            '( Invoice# 122303 0002000266 )'
        )

        self.assertEqual(
            infer_vendor_from_sender(sender_email='noreply@suspension.randysww.com', sender_header=''),
            '',
        )
        self.assertEqual(
            infer_vendor_from_email_metadata(
                sender_email='noreply@suspension.randysww.com',
                sender_header='',
                subject=carli_subject,
            ),
            'Carli Suspension - $10 DS Fee',
        )
        self.assertEqual(
            infer_vendor_from_email_metadata(
                sender_email='noreply@suspension.randysww.com',
                sender_header='',
                subject='Fwd: ***DO NOT REPLY***Shared vendor invoice attached',
                message_text='',
            ),
            '',
        )

    def test_daystar_shared_sender_domain_uses_email_body_address(self):
        daystar_body = (
            'Daystar\n'
            '15461 Slover Ave\n'
            'Fontana CA 92337\n'
        )

        self.assertEqual(
            infer_vendor_from_email_metadata(
                sender_email='noreply@suspension.randysww.com',
                sender_header='',
                subject='Fwd: invoice attached',
                message_text=daystar_body,
            ),
            'Daystar',
        )

    def test_carli_shared_sender_domain_uses_full_email_body(self):
        carli_body = (
            'Thanks for your order.\n'
            'The invoice is attached.\n'
            'Sincerely,\n'
            'Carli\n'
        )

        self.assertEqual(
            infer_vendor_from_email_metadata(
                sender_email='noreply@suspension.randysww.com',
                sender_header='',
                subject='Fwd: invoice attached',
                message_text=carli_body,
            ),
            'Carli Suspension - $10 DS Fee',
        )

    def test_kc_turbos_sender_header_matches_embedded_whitelisted_address(self):
        self.assertEqual(
            infer_vendor_from_sender(
                sender_email='system@sent-via.netsuite.com',
                sender_header='"KC Turbos Invoicing (invoicing@kcturbos.com)" <system@sent-via.netsuite.com>',
            ),
            'KC Turbos',
        )

    def test_carli_shared_sender_reapplies_vendor_parser_before_prevalidation(self):
        mock_pdf = mock.MagicMock()
        mock_pdf.pages = [object()]
        mock_context = mock.MagicMock()
        mock_context.__enter__.return_value = mock_pdf
        mock_context.__exit__.return_value = False
        generic_parse = {
            'invoice_number': '',
            'vendor': 'Tracking: 1Z351FW60374214427',
            'vendor_address': '',
            'customer': '',
            'date': '',
            'due_date': '',
            'terms': '',
            'po_number': '',
            'tracking_number': '',
            'shipping_method': '',
            'ship_date': '',
            'shipping_tax_code': '',
            'shipping_tax_rate': '',
            'subtotal': '',
            'shipping_cost': '',
            'shipping_description': '',
            'total': '',
            'line_items': [],
        }
        refreshed_parse = {
            **generic_parse,
            'vendor': 'Carli Suspension - $10 DS Fee',
            'invoice_number': '120872',
            'po_number': '0045969',
            'customer': 'Cade Grant',
            'terms': '1% 10 Net 30',
            'line_items': [
                {'item_number': 'CS-CA-MS14-94', 'amount': '565.50'},
                {'item_number': 'DROP SHIP', 'amount': '20.00'},
            ],
        }

        with mock.patch('invoice_parser.pdfplumber.open', return_value=mock_context), \
             mock.patch(
                 'invoice_parser.extract_text_from_pdf',
                 return_value='Carli layout text with enough extracted characters to skip OCR fallback.',
             ), \
             mock.patch('invoice_parser.parse_invoice_text', return_value=generic_parse), \
             mock.patch('invoice_parser._refresh_vendor_dependent_fields', return_value=refreshed_parse) as refresh_mock:
            data = parse_invoice(
                'C:\\temp\\INVOICE 120872 03_05_26 14_24_14 5.PDF',
                sender_email='noreply@suspension.randysww.com',
                sender_subject='Fwd: invoice attached',
                sender_message_text=(
                    'Please find attached Invoice# 120872 for your PO# 0045969\n'
                    'Sincerely,\n'
                    'Carli Suspension\n'
                ),
            )

        self.assertFalse(data.get('not_an_invoice'))
        self.assertEqual(data.get('vendor'), 'Carli Suspension - $10 DS Fee')
        self.assertEqual(data.get('invoice_number'), '120872')
        self.assertEqual(data.get('po_number'), '0045969')
        self.assertEqual(data.get('customer'), 'Cade Grant')
        self.assertEqual(data.get('terms'), '1% 10 Net 30')
        self.assertEqual(len(data.get('line_items') or []), 2)
        refresh_mock.assert_called_once()

    def test_carli_shared_sender_parses_quantity_first_layout(self):
        mock_pdf = mock.MagicMock()
        mock_pdf.pages = [object()]
        mock_context = mock.MagicMock()
        mock_context.__enter__.return_value = mock_pdf
        mock_context.__exit__.return_value = False
        carli_layout = (
            'INVOICE\n'
            '122225 4/14/2026\n'
            '596Crane St, Lake Elsinore, CA 92530\n'
            'ph:888-992-2754\n'
            'Bill To: Ship To:\n'
            'POWER PRODUCTS UNLIMITED, INC John Lilienthal\n'
            '5204E BROADWAY AVE 101S east street\n'
            'SPOKANE VALLEY WA UNITED STATES OF Eckley CO 80727UNITED STATES OF AMERICA\n'
            'AMERICA\n'
            'Terms: 1% 10 NET 30 Payment Due: 5/14/2026\n'
            'Tracking: 1Z351FW60375328820 Shipped Via: UPS THIRD PARTY\n'
            'Quantity Item Number Description Price Extension\n'
            'Pack Slip # PO Number Order Date\n'
            '1.00 CS-DBMM-0359 DODGE BILLET MOTOR MOUNT, 2003-2007 5.9L DIESEL 255.45 EACH 255.45\n'
            '621227 0057931 4/14/2026\n'
            '1.00 DROP SHIP DROP SHIP FEE 20.00 EACH 20.00\n'
            '621227 0057931 4/14/2026\n'
            'All Prices Are Shown in United States Dollar\n'
            'Subtotal: 275.45\n'
            'Tax: 0.00\n'
            'Freight: 0.00\n'
            'Total: 275.45\n'
            'Thank You\n'
            'Payments Applied: 0.00\n'
            'Balance Due: 275.45\n'
        )

        with mock.patch('invoice_parser.pdfplumber.open', return_value=mock_context), \
             mock.patch('invoice_parser.extract_text_from_pdf', return_value=carli_layout), \
             mock.patch('invoice_parser.extract_layout_text_from_pdf', return_value=carli_layout), \
             mock.patch('invoice_parser._extract_carli_ship_to_lines', return_value=[]):
            data = parse_invoice(
                'C:\\temp\\INVOICE-4-20-2026 04_20_26 09_02_16 880.PDF',
                sender_email='noreply@suspension.randysww.com',
                sender_header='<noreply@suspension.randysww.com>',
                sender_subject='***DO NOT REPLY***Carli Suspension Invoice(s) Attached ( Invoice# 122225 0002000266 )',
                sender_message_text=(
                    'Please find attached Invoice# 122225 for your PO# 0057931\n'
                    'Sincerely,\n'
                    'Carli Suspension\n'
                ),
            )

        self.assertFalse(data.get('not_an_invoice'))
        self.assertEqual(data.get('vendor'), 'Carli Suspension - $10 DS Fee')
        self.assertEqual(data.get('invoice_number'), '122225')
        self.assertEqual(data.get('po_number'), '0057931')
        self.assertEqual(data.get('terms'), '1% 10 Net 30')
        self.assertEqual(len(data.get('line_items') or []), 2)
        self.assertEqual(data['line_items'][0].get('item_number'), 'CS-DBMM-0359')
        self.assertEqual(data['line_items'][1].get('description'), 'Drop Ship')

    def test_power_stroke_products_credit_card_and_will_call(self):
        stock_path = os.path.join(TRAINING_DIR, 'PS', 'Invoice_10513_from_PowerStroke_Products.pdf')
        retail_path = os.path.join(TRAINING_DIR, 'PS', 'Invoice_10488_from_PowerStroke_Products.pdf')

        stock_data = parse_invoice(stock_path)
        retail_data = parse_invoice(retail_path)

        self.assertEqual(stock_data.get('vendor'), 'Power Stroke Products')
        self.assertEqual(stock_data.get('invoice_number'), '10513')
        self.assertEqual(stock_data.get('po_number'), '0048459')
        self.assertEqual(stock_data.get('terms'), 'Credit Card')
        self.assertTrue(stock_data.get('stock_order'))
        self.assertEqual(stock_data.get('stock_order_description'), 'WILL CALL')
        self.assertEqual(stock_data.get('customer'), 'Davie Stommes')

        self.assertEqual(retail_data.get('vendor'), 'Power Stroke Products')
        self.assertEqual(retail_data.get('invoice_number'), '10488')
        self.assertEqual(retail_data.get('customer'), 'Dana Dickson')
        self.assertEqual(retail_data.get('terms'), 'Credit Card')
        self.assertEqual(len(retail_data.get('line_items') or []), 1)
        self.assertEqual(retail_data['line_items'][0].get('item_number'), 'PP-12vLHDVS')

    def test_hamilton_cams_uses_configured_vendor_identity(self):
        invoice_path = os.path.join(TRAINING_DIR, 'HC', 'Inv_203231_from_Hamilton_Cams_26140.pdf')

        data = parse_invoice(invoice_path)

        self.assertEqual(data.get('vendor'), 'Hamilton Cams - $20 Dropship Fee')
        self.assertEqual(data.get('vendor_address'), '2881 CR 207, Lampasas, TX 76550')
        self.assertEqual(data.get('terms'), 'Credit Card')
        self.assertEqual(data.get('customer'), 'Trevor Morris')
        self.assertEqual(data.get('total'), '158.26')

    def test_beans_diesel_performance_normalizes_vendor_and_terms(self):
        invoice_path = os.path.join(TRAINING_DIR, 'BDP', 'Inv_13264_from_BDP_Bean_Machine_10716.pdf')

        data = parse_invoice(invoice_path)

        self.assertEqual(data.get('vendor'), 'Beans Diesel Performance')
        self.assertEqual(data.get('vendor_address'), '210 Rollin Coal Lane, Woodbury, TN 37190')
        self.assertEqual(data.get('terms'), 'Credit Card')
        self.assertEqual(data.get('invoice_number'), '13264')
        self.assertEqual(data.get('po_number'), '0042161')
        self.assertEqual(len(data.get('line_items') or []), 2)
        self.assertEqual(data['line_items'][0].get('item_number'), '288013')
        self.assertEqual(data['line_items'][0].get('description'), '-12 Fitting')
        self.assertEqual(data['line_items'][0].get('amount'), '13.50')

    def test_beans_diesel_performance_keeps_288013_as_its_own_line(self):
        invoice_path = os.path.join(TRAINING_DIR, 'BDP', 'Inv_13265_from_BDP_Bean_Machine_10716.pdf')

        data = parse_invoice(invoice_path)

        self.assertEqual(data.get('invoice_number'), '13265')
        self.assertEqual(len(data.get('line_items') or []), 3)
        self.assertEqual(data['line_items'][0].get('item_number'), '288012')
        self.assertEqual(data['line_items'][0].get('description'), '-12 O-Ring Boss Plug')
        self.assertEqual(data['line_items'][1].get('item_number'), '288013')
        self.assertEqual(data['line_items'][1].get('description'), '-12 Fitting')
        self.assertEqual(data['line_items'][1].get('amount'), '6.75')

    def test_bosch_uses_configured_terms_and_stock_order_detection(self):
        invoice_path = os.path.join(TRAINING_DIR, 'BCH', '903983345 03.03.2026.PDF')

        data = parse_invoice(invoice_path)

        self.assertEqual(data.get('vendor'), 'Bosch')
        self.assertEqual(data.get('invoice_number'), '903983345')
        self.assertEqual(data.get('po_number'), '0044578')
        self.assertEqual(data.get('terms'), 'Net 10th Prox.')
        self.assertEqual(data.get('vendor_address'), 'P.O. Box 7410506, Chicago, IL 60674-0506')
        self.assertTrue(data.get('stock_order'))
        self.assertEqual(data.get('stock_order_description'), 'STOCK ORDER')
        self.assertEqual(data.get('customer'), 'Diesel Power Products')

    def test_diesel_forward_uses_alliant_layout(self):
        invoice_path = os.path.join(TRAINING_DIR, 'ALL', 'Sales Invoice SI-940658.pdf')

        data = parse_invoice(invoice_path)

        self.assertEqual(data.get('vendor'), 'Diesel Forward')
        self.assertEqual(data.get('invoice_number'), 'SI-940658')
        self.assertEqual(data.get('po_number'), '0045605')
        self.assertEqual(data.get('terms'), 'Credit Card')
        self.assertEqual(data.get('customer'), 'Scott Cesaroni')
        self.assertEqual(data.get('total'), '67.47')
        self.assertEqual(data['line_items'][0].get('item_number'), '3945268')

    def test_carli_extracts_drop_ship_fee_line(self):
        invoice_path = os.path.join(TRAINING_DIR, 'CRL', 'INVOICE 120871 03_05_26 14_22_22 516.PDF')

        data = parse_invoice(invoice_path)

        self.assertEqual(data.get('vendor'), 'Carli Suspension - $10 DS Fee')
        self.assertEqual(data.get('invoice_number'), '120871')
        self.assertEqual(data.get('po_number'), '0045952')
        self.assertEqual(data.get('terms'), '1% 10 Net 30')
        self.assertEqual(data.get('customer'), 'Dominic Monego')
        self.assertEqual(len(data.get('line_items') or []), 2)
        self.assertEqual(data['line_items'][1].get('description'), 'DROP SHIP FEE')
        self.assertEqual(data['line_items'][1].get('amount'), '20.00')

    def test_icon_vehicle_dynamics_extracts_freight(self):
        invoice_path = os.path.join(TRAINING_DIR, 'ICO', 'INVOICE 317009 03_13_26 10_05_48 314.PDF')

        data = parse_invoice(invoice_path)

        self.assertEqual(data.get('vendor'), 'Icon Vehicle Dynamics')
        self.assertEqual(data.get('invoice_number'), '317009')
        self.assertEqual(data.get('po_number'), '0041464')
        self.assertEqual(data.get('terms'), '2% 20 Net 30')
        self.assertEqual(data.get('customer'), 'Brody Kirkham')
        self.assertEqual(data.get('shipping_cost'), '24.66')
        self.assertEqual(data.get('shipping_description'), 'Freight')
        self.assertEqual(len(data.get('line_items') or []), 2)

    def test_icon_reprint_layout_still_reads_primary_description(self):
        invoice_path = os.path.join(TRAINING_DIR, 'ICO', 'INVOICE 318632 03_05_26 14_31_36 914.PDF')

        data = parse_invoice(invoice_path)

        self.assertEqual(data.get('vendor'), 'Icon Vehicle Dynamics')
        self.assertEqual(data.get('invoice_number'), '318632')
        self.assertEqual(data.get('po_number'), '0026074-1')
        self.assertEqual(data['line_items'][0].get('item_number'), '66516')
        self.assertEqual(data['line_items'][0].get('description'), '17-25 FSD REAR 0-2" 2.0 VS IR')

    def test_cognito_motorsports_extracts_key_fields(self):
        invoice_path = os.path.join(TRAINING_DIR, 'CO', 'INVOICE-3-4-2026-Cognito Motorsports 03_04_26 17_28_24 341.PDF')

        data = parse_invoice(invoice_path)

        self.assertEqual(data.get('vendor'), 'Cognito Motorsports')
        self.assertEqual(data.get('invoice_number'), '7176-CMS')
        self.assertEqual(data.get('po_number'), '0042913')
        self.assertEqual(data.get('terms'), 'Net 15')
        self.assertEqual(data.get('customer'), 'Richie Shawver')
        self.assertEqual(data['line_items'][0].get('item_number'), '210-91023')

    def test_cognito_freight_line_item_does_not_override_zero_footer_freight(self):
        invoice_path = os.path.join(TRAINING_DIR, 'CO', 'INVOICE-3-3-2026 03_03_26 17_10_51 356.PDF')

        data = parse_invoice(invoice_path)

        self.assertEqual(data.get('vendor'), 'Cognito Motorsports')
        self.assertEqual(data.get('invoice_number'), '7042-CMS')
        self.assertEqual(data.get('shipping_cost'), '0.00')
        self.assertEqual(data.get('shipping_description'), 'Freight')
        self.assertEqual(len(data.get('line_items') or []), 2)
        self.assertEqual(data['line_items'][1].get('item_number'), 'Freight')
        self.assertEqual(data['line_items'][1].get('amount'), '21.28')


if __name__ == '__main__':
    unittest.main()
