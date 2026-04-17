import os
import unittest

from invoice_parser import (
    _is_hamilton_vendor_name,
    infer_vendor_from_folder_marker,
    normalize_vendor_name,
)


class HamiltonInvoiceParserTests(unittest.TestCase):
    def test_normalize_vendor_name_maps_old_and_new_hamilton_names_to_new_canonical_name(self):
        self.assertEqual(normalize_vendor_name('Hamilton Cams'), 'Hamilton Cams')
        self.assertEqual(
            normalize_vendor_name('Hamilton Cams - $20 Dropship Fee'),
            'Hamilton Cams',
        )

    def test_folder_marker_infers_hamilton_cams_canonical_name(self):
        invoice_path = os.path.join('C:\\repo', 'training', 'HC', 'invoice.pdf')

        self.assertEqual(infer_vendor_from_folder_marker(invoice_path), 'Hamilton Cams')

    def test_hamilton_vendor_match_accepts_legacy_name(self):
        self.assertTrue(_is_hamilton_vendor_name('Hamilton Cams - $20 Dropship Fee'))


if __name__ == '__main__':
    unittest.main()
