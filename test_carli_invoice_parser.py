import unittest
from unittest.mock import patch

from invoice_parser import _extract_carli_items_from_layout


class CarliInvoiceParserTests(unittest.TestCase):
    def test_carli_drop_ship_fee_exports_as_category_detail_with_drop_ship_description(self):
        layout_text = """Invoice No.: INV-1001
Invoice Date: 4/16/2026
Item Number Description Quantity Price Unit Extension
CS-KIT Front Coilovers 1 1000.00 Each 1000.00
DROP SHIP FEE 1 10.00 Each 10.00
Subtotal: 1010.00
"""

        with patch('invoice_parser.extract_layout_text_from_pdf', return_value=layout_text):
            line_items = _extract_carli_items_from_layout('carli.pdf')

        self.assertEqual(len(line_items), 2)

        drop_ship_item = line_items[1]
        self.assertTrue(drop_ship_item.get('is_freight'))
        self.assertEqual(drop_ship_item.get('description'), 'Drop Ship')
        self.assertEqual(drop_ship_item.get('qb_type_override'), 'Category Details')
        self.assertEqual(drop_ship_item.get('qb_category_override'), 'Purchases')
        self.assertEqual(drop_ship_item.get('qb_product_service_override'), 'Drop Ship')


if __name__ == '__main__':
    unittest.main()
