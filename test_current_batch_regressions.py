import json
import os
import unittest

from invoice_parser import (
    _text_explicitly_mentions_vendor,
    infer_vendor_from_email_metadata,
    parse_invoice,
    validate_vendor_name,
)


APP_DIR = os.path.dirname(os.path.abspath(__file__))
REPO_ROOT = os.path.dirname(APP_DIR)
CURRENT_BATCH_DIR = os.path.join(REPO_ROOT, 'Invoices', 'invoices_4-24')


def _load_sidecar(pdf_path):
    with open(pdf_path + '.sender.json', encoding='utf-8') as f:
        return json.load(f)


class CurrentBatchRegressionTests(unittest.TestCase):
    def test_shared_forwarder_icon_body_resolves_icon_vendor(self):
        body = (
            '---------- Forwarded message ---------\n'
            'From: <noreply@suspension.randysww.com>\n'
            'Subject: ***DO NOT REPLY***ICON Vehicle Dynamics Invoice(s) Attached\n'
            'Please reply to mandy@iconvehicledynamics.com.\n'
        )

        self.assertEqual(
            infer_vendor_from_email_metadata(
                sender_email='noreply@suspension.randysww.com',
                subject='Invoice attached',
                message_text=body,
            ),
            'Icon Vehicle Dynamics',
        )

    def test_printed_timestamp_is_not_a_vendor_name(self):
        self.assertFalse(validate_vendor_name('Printed 04/23/2026 12:42:05 PM'))

    def test_label_like_lines_are_not_vendor_names(self):
        self.assertFalse(validate_vendor_name('Tracking: 1Z351FW60375659526'))
        self.assertFalse(validate_vendor_name('Attn: Josh'))
        self.assertFalse(
            _text_explicitly_mentions_vendor(
                'Tracking: 1Z351FW60375659526',
                'Tracking: 1Z351FW60375659526',
            )
        )

    @unittest.skipUnless(
        os.path.exists(os.path.join(CURRENT_BATCH_DIR, 'INVOICE 122605 04_23_26 09_42_55 940.PDF')),
        'current 4-24 invoice batch not available',
    )
    def test_carli_shared_forwarder_is_applied_before_not_invoice_check(self):
        pdf_path = os.path.join(CURRENT_BATCH_DIR, 'INVOICE 122605 04_23_26 09_42_55 940.PDF')
        meta = _load_sidecar(pdf_path)

        data = parse_invoice(
            pdf_path,
            sender_email=meta.get('sender_email', ''),
            sender_header=meta.get('sender_header', ''),
            sender_subject=meta.get('subject', ''),
            sender_message_text=meta.get('message_text', ''),
        )

        self.assertFalse(data.get('not_an_invoice'))
        self.assertEqual(data.get('vendor'), 'Carli Suspension - $10 DS Fee')
        self.assertEqual(data.get('invoice_number'), '122605')
        self.assertEqual(data.get('po_number'), '0060773')
        self.assertGreaterEqual(len(data.get('line_items') or []), 1)

    @unittest.skipUnless(
        os.path.exists(os.path.join(CURRENT_BATCH_DIR, 'INVOICE 322615 04_22_26 09_58_46 591.PDF')),
        'current 4-24 invoice batch not available',
    )
    def test_icon_shared_forwarder_is_applied_before_not_invoice_check(self):
        pdf_path = os.path.join(CURRENT_BATCH_DIR, 'INVOICE 322615 04_22_26 09_58_46 591.PDF')
        meta = _load_sidecar(pdf_path)

        data = parse_invoice(
            pdf_path,
            sender_email=meta.get('sender_email', ''),
            sender_header=meta.get('sender_header', ''),
            sender_subject=meta.get('subject', ''),
            sender_message_text=meta.get('message_text', ''),
        )

        self.assertFalse(data.get('not_an_invoice'))
        self.assertEqual(data.get('vendor'), 'Icon Vehicle Dynamics')
        self.assertEqual(data.get('invoice_number'), '322615')
        self.assertEqual(data.get('po_number'), '0062028')
        self.assertGreaterEqual(len(data.get('line_items') or []), 1)

    @unittest.skipUnless(
        os.path.exists(
            os.path.join(
                CURRENT_BATCH_DIR,
                'Inv_564985_from_RedHead_Steering_Gears_Inc._20410003_77848.pdf',
            )
        ),
        'current 4-24 invoice batch not available',
    )
    def test_redhead_customer_ship_to_is_not_marked_stock_order(self):
        pdf_path = os.path.join(
            CURRENT_BATCH_DIR,
            'Inv_564985_from_RedHead_Steering_Gears_Inc._20410003_77848.pdf',
        )

        data = parse_invoice(pdf_path)

        self.assertFalse(data.get('stock_order'))
        self.assertEqual(data.get('customer'), 'YOSVANI DIAZ')

    @unittest.skipUnless(
        os.path.exists(os.path.join(REPO_ROOT, 'Invoices', 'invoices_4-27_2', 'INVOICE 122685 04_24_26 09_25_27 899.PDF')),
        'current 4-27_2 Carli invoice batch not available',
    )
    def test_carli_tracking_line_does_not_block_sender_vendor(self):
        pdf_path = os.path.join(REPO_ROOT, 'Invoices', 'invoices_4-27_2', 'INVOICE 122685 04_24_26 09_25_27 899.PDF')
        meta = _load_sidecar(pdf_path)

        data = parse_invoice(
            pdf_path,
            sender_email=meta.get('sender_email', ''),
            sender_header=meta.get('sender_header', ''),
            sender_subject=meta.get('subject', ''),
            sender_message_text=meta.get('message_text', ''),
        )

        self.assertFalse(data.get('not_an_invoice'))
        self.assertEqual(data.get('vendor'), 'Carli Suspension - $10 DS Fee')
        self.assertEqual(data.get('invoice_number'), '122685')
        self.assertEqual(data.get('po_number'), '0063442')
        self.assertGreaterEqual(len(data.get('line_items') or []), 1)

    @unittest.skipUnless(
        os.path.exists(os.path.join(REPO_ROOT, 'Invoices', 'invoices_4-27', '2146721-Customer-Copy.pdf')),
        'current 4-27 Serra invoice batch not available',
    )
    def test_serra_attn_line_does_not_block_sender_vendor(self):
        pdf_path = os.path.join(REPO_ROOT, 'Invoices', 'invoices_4-27', '2146721-Customer-Copy.pdf')
        meta = _load_sidecar(pdf_path)

        data = parse_invoice(
            pdf_path,
            sender_email=meta.get('sender_email', ''),
            sender_header=meta.get('sender_header', ''),
            sender_subject=meta.get('subject', ''),
            sender_message_text=meta.get('message_text', ''),
        )

        self.assertEqual(data.get('vendor'), 'Serra Chrysler Dodge Ram Jeep of Traverse City')
        self.assertEqual(data.get('invoice_number'), '2146721')
        self.assertEqual(data.get('customer'), 'DANIEL MULLENBACH')

    @unittest.skipUnless(
        os.path.exists(os.path.join(REPO_ROOT, 'Invoices', 'invoices_4-27', 'Invoice_I472020_1777072161408.pdf')),
        'current 4-27 Daystar invoice batch not available',
    )
    def test_daystar_uses_left_ship_to_for_stock_detection(self):
        pdf_path = os.path.join(REPO_ROOT, 'Invoices', 'invoices_4-27', 'Invoice_I472020_1777072161408.pdf')

        data = parse_invoice(pdf_path)

        self.assertEqual(data.get('vendor'), 'Daystar')
        self.assertFalse(data.get('stock_order'))
        self.assertEqual(data.get('customer'), 'Jon Kineshanko')

    @unittest.skipUnless(
        os.path.exists(
            os.path.join(
                REPO_ROOT,
                'Invoices',
                'invoices_4-27',
                'Invoice_Ticket63907_from_FUMOTO_ENGINEERING_OF_AMERICA_INC.pdf',
            )
        ),
        'current 4-27 Fumoto invoice batch not available',
    )
    def test_fumoto_uses_ship_to_table_for_stock_detection(self):
        pdf_path = os.path.join(
            REPO_ROOT,
            'Invoices',
            'invoices_4-27',
            'Invoice_Ticket63907_from_FUMOTO_ENGINEERING_OF_AMERICA_INC.pdf',
        )

        data = parse_invoice(pdf_path)

        self.assertEqual(data.get('vendor'), 'Fumoto Engineering of America')
        self.assertFalse(data.get('stock_order'))
        self.assertEqual(data.get('customer'), 'Carl A. Rossi')
        self.assertEqual(data.get('invoice_number'), 'Ticket-63907')
        self.assertEqual(data.get('po_number'), '0063646')


if __name__ == '__main__':
    unittest.main()
