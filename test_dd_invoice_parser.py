import os
import unittest

from invoice_parser import parse_invoice


APP_DIR = os.path.dirname(os.path.abspath(__file__))
TRAINING_DD_DIR = os.path.join(APP_DIR, 'training', 'DD')


class DynomiteDieselParserTests(unittest.TestCase):
    def test_12632_1_extracts_footer_shipping_charge(self):
        invoice_path = os.path.join(TRAINING_DD_DIR, '12632.1.pdf')

        data = parse_invoice(invoice_path, lambda *args, **kwargs: None)

        self.assertEqual(data.get('vendor'), 'Dynomite Diesel')
        self.assertEqual(data.get('customer'), 'Rudy Carrillo')
        self.assertEqual(data.get('shipping_cost'), '28.11')
        self.assertEqual(data.get('shipping_description'), 'Shipping')
        self.assertEqual(data.get('total'), '543.11')


if __name__ == '__main__':
    unittest.main()
