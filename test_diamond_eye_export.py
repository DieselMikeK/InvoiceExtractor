import os
import tempfile
import unittest

from spreadsheet_writer import read_spreadsheet_rows, write_invoice_to_spreadsheet


class DiamondEyeExportTests(unittest.TestCase):
    def test_diamond_eye_rows_export_as_category_details(self):
        invoice_data = {
            'invoice_number': '26528',
            'vendor': 'Diamond Eye Manufacturing - $3.00 DS Fee',
            'date': '2/6/2026',
            'po_number': '0039420',
            'total': '107.21',
            'line_items': [
                {
                    'item_number': '445053',
                    'quantity': '1',
                    'unit_price': '11.13',
                    'amount': '11.13',
                    'description': 'CLAMP, 5", WELD-ON HANGER; ALUMINIZED',
                    'units': 'Each',
                },
                {
                    'item_number': '222011',
                    'quantity': '1',
                    'unit_price': '88.58',
                    'amount': '88.58',
                    'description': 'TAILPIPE',
                    'units': 'Each',
                },
                {
                    'item_number': 'HDL',
                    'quantity': '1',
                    'unit_price': '7.50',
                    'amount': '7.50',
                    'description': 'HANDLING FEE',
                    'units': 'Each',
                    'qb_category_override': 'Purchases',
                    'qb_product_service_override': 'Drop Ship',
                    'qb_sku_override': 'HDL',
                },
            ],
            'shipping_cost': '',
            'shipping_description': 'Shipping',
        }

        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = os.path.join(tmpdir, 'diamond_eye.xlsx')
            write_invoice_to_spreadsheet(output_path, invoice_data)
            rows = read_spreadsheet_rows(output_path)

        sku_rows = {
            str(row.get('sku', '')).strip(): row
            for row in rows
            if str(row.get('sku', '')).strip()
        }
        shipping_row = next(
            row for row in rows
            if str(row.get('description', '')).strip() == 'Shipping'
        )

        for sku in ('445053', '222011', 'HDL'):
            row = sku_rows[sku]
            self.assertEqual(row['type'], 'Category Details')
            self.assertEqual(row['category'], 'Purchases')
            self.assertEqual(row['product_service'], '')
            self.assertEqual(row['qty'], '')
            self.assertEqual(row['rate'], '')
            self.assertNotEqual(str(row['amount']).strip(), '')

        self.assertEqual(shipping_row['type'], 'Category Details')
        self.assertEqual(shipping_row['category'], 'Purchases')
        self.assertEqual(shipping_row['product_service'], '')
        self.assertEqual(shipping_row['qty'], '')
        self.assertEqual(shipping_row['rate'], '')
        self.assertEqual(str(shipping_row['amount']).strip(), '0')


if __name__ == '__main__':
    unittest.main()
