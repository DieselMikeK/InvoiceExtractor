import os
import unittest

from invoice_parser import infer_vendor_from_sender, parse_invoice


APP_DIR = os.path.dirname(os.path.abspath(__file__))
TRAINING_DIR = os.path.join(APP_DIR, 'training')


class NewVendorFirstPassTests(unittest.TestCase):
    def test_sender_aliases_cover_carli_hamilton_and_icon(self):
        cases = [
            ('sales@carlisuspension.com', '', 'Carli Suspension - $10 DS Fee'),
            ('noreply@suspension.randysww.com', '', 'Carli Suspension - $10 DS Fee'),
            ('hamiltoncamsales@gmail.com', '', 'Hamilton Cams - $20 Dropship Fee'),
            ('', 'ICON Vehicle Dynamics <orders@iconvehicledynamics.com>', 'Icon Vehicle Dynamics'),
        ]

        for sender_email, sender_header, expected_vendor in cases:
            with self.subTest(sender_email=sender_email, sender_header=sender_header):
                self.assertEqual(
                    infer_vendor_from_sender(sender_email=sender_email, sender_header=sender_header),
                    expected_vendor,
                )

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
