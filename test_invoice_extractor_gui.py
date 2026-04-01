import unittest

from invoice_extractor_gui import _is_diamond_eye_zero_shipping_batch_row


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


if __name__ == '__main__':
    unittest.main()
