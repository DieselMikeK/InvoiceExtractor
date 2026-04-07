import os
import tempfile
import unittest

from invoice_extractor_gui import (
    _is_diamond_eye_zero_shipping_batch_row,
    _load_sender_sidecar,
    _lookup_sender_metadata_entry,
    _save_sender_sidecar,
)


class DiamondEyeBatchExportFilterTests(unittest.TestCase):
    def test_drops_zero_shipping_row_for_diamond_eye(self):
        row = {
            'vendor': 'Diamond Eye Manufacturing - $3.00 DS Fee',
            'category': 'Freight and shipping costs',
            'product_service': 'Shipping',
            'sku': '',
            'rate': '0',
        }

        self.assertTrue(_is_diamond_eye_zero_shipping_batch_row(row))

    def test_keeps_positive_shipping_row_for_diamond_eye(self):
        row = {
            'vendor': 'Diamond Eye Manufacturing',
            'category': 'Freight and shipping costs',
            'product_service': 'Shipping',
            'sku': '',
            'rate': '18.75',
        }

        self.assertFalse(_is_diamond_eye_zero_shipping_batch_row(row))

    def test_keeps_zero_shipping_row_for_other_vendors(self):
        row = {
            'vendor': 'ATS Diesel',
            'category': 'Freight and shipping costs',
            'product_service': 'Shipping',
            'sku': '',
            'rate': '0',
        }

        self.assertFalse(_is_diamond_eye_zero_shipping_batch_row(row))

class SenderMetadataLookupTests(unittest.TestCase):
    def test_falls_back_to_filename_when_source_path_misses(self):
        entries = {
            'oldroot/Invoices/Invoice_123.pdf': {
                'source_file': 'oldroot/Invoices/Invoice_123.pdf',
                'filename': 'Invoice_123.pdf',
                'sender_email': 'invoicing@kcturbos.com',
            }
        }

        entry = _lookup_sender_metadata_entry(
            entries,
            'newroot/Invoices/Invoice_123.pdf',
            'Invoice_123.pdf',
        )

        self.assertEqual(entry.get('sender_email'), 'invoicing@kcturbos.com')

    def test_sender_sidecar_round_trip(self):
        entry = {
            'source_file': 'Invoices/Invoice_123.pdf',
            'filename': 'Invoice_123.pdf',
            'sender_email': 'invoicing@kcturbos.com',
            'sender_header': 'KC Turbos <invoicing@kcturbos.com>',
            'subject': 'Invoice attached',
            'message_id': 'abc123',
            'downloaded_at': '2026-04-07 14:00:00',
        }

        with tempfile.TemporaryDirectory() as tmpdir:
            pdf_path = os.path.join(tmpdir, 'Invoice_123.pdf')
            with open(pdf_path, 'wb') as f:
                f.write(b'%PDF-1.4')
            _save_sender_sidecar(pdf_path, entry)
            loaded = _load_sender_sidecar(pdf_path)

        self.assertEqual(loaded.get('sender_email'), 'invoicing@kcturbos.com')
        self.assertEqual(loaded.get('message_id'), 'abc123')


if __name__ == '__main__':
    unittest.main()
