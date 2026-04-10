import unittest

from invoice_parser import _matches_internal_stock_customer_hint


class StockOrderDetectionTests(unittest.TestCase):
    def test_matches_company_name(self):
        self.assertTrue(_matches_internal_stock_customer_hint('Diesel Power Products'))

    def test_matches_warehouse_address_fragment(self):
        self.assertTrue(_matches_internal_stock_customer_hint('E. Main Ave.'))
        self.assertTrue(_matches_internal_stock_customer_hint('6200 East Main Avenue'))

    def test_ignores_regular_customer_names(self):
        self.assertFalse(_matches_internal_stock_customer_hint('Thomas Shelton'))
        self.assertFalse(_matches_internal_stock_customer_hint('Lane Ricketson'))


if __name__ == '__main__':
    unittest.main()
