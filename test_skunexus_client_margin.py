import unittest
from unittest.mock import patch

from skunexus_client import SkuNexusClient, infer_invoice_row_sku_from_po


def _mapped_group(price, cost=None, po_label='0055056'):
    custom_values = [{'custom_field_id': 'price', 'value': price}]
    if cost is not None:
        custom_values.append({'custom_field_id': 'cost', 'value': cost})
    return {
        'relatedProduct': {
            'customValues': custom_values,
        },
        'decisionItems': [
            {
                'decidedItems': [
                    {
                        'decisions': [
                            {
                                'qty': 1,
                                'relatedPurchaseOrder': {'label': po_label},
                            }
                        ]
                    }
                ]
            }
        ],
    }


class PoMarginTests(unittest.TestCase):
    def setUp(self):
        self.client = SkuNexusClient('', '')

    def test_industrial_injection_still_uses_po_line_price_when_related_cost_exists(self):
        po_details = {
            'label': '0055056',
            'vendor': {'name': 'Industrial Injection'},
            'lineItems': {'rows': [{'price': 1376.40}]},
            'allRelatedOrders': [{'id': 'order-1', 'label': '880443'}],
        }
        order_details = {
            'groupedDecisionItems': [
                _mapped_group(1696.48, 1302.97),
            ]
        }

        with patch.object(self.client, 'get_order_grouped_items', return_value=(order_details, None)):
            margin, error = self.client.get_po_margin(po_details, '0055056')

        self.assertIsNone(error)
        self.assertAlmostEqual(margin, 0.188673, places=4)

    def test_non_industrial_injection_uses_po_line_price(self):
        po_details = {
            'label': '0055056',
            'vendor': {'name': 'Some Other Vendor'},
            'lineItems': {'rows': [{'price': 1376.40}]},
            'allRelatedOrders': [{'id': 'order-1', 'label': '880443'}],
        }
        order_details = {
            'groupedDecisionItems': [
                _mapped_group(1696.48, 1302.97),
            ]
        }

        with patch.object(self.client, 'get_order_grouped_items', return_value=(order_details, None)):
            margin, error = self.client.get_po_margin(po_details, '0055056')

        self.assertIsNone(error)
        self.assertAlmostEqual(margin, 0.188673, places=4)

    def test_industrial_injection_uses_po_line_price_without_related_cost(self):
        po_details = {
            'label': '0055056',
            'vendor': {'name': 'Industrial Injection'},
            'lineItems': {'rows': [{'price': 1376.40}]},
            'allRelatedOrders': [{'id': 'order-1', 'label': '880443'}],
        }
        order_details = {
            'groupedDecisionItems': [
                _mapped_group(1696.48, None),
            ]
        }

        with patch.object(self.client, 'get_order_grouped_items', return_value=(order_details, None)):
            margin, error = self.client.get_po_margin(po_details, '0055056')

        self.assertIsNone(error)
        self.assertAlmostEqual(margin, 0.188673, places=4)

    def test_infers_missing_sb_sku_by_price_and_description(self):
        po_details = {
            'vendor': {'name': 'S&B Filters'},
            'lineItems': {
                'rows': [
                    {
                        'id': 'line-1',
                        'product': {
                            'sku': '83-2004',
                            'name': 'Hot Side Intercooler Pipe for 2016-2026 Ford Powerstroke 6.7L',
                        },
                        'quantity': 1,
                        'price': '166.83',
                        'total_price': '166.83',
                    },
                    {
                        'id': 'line-2',
                        'product': {
                            'sku': 'SB-83-1001',
                            'name': 'Cold Side Intercooler Pipe for 2017-2026 Ford Super Duty, 6.7L Powerstroke',
                        },
                        'quantity': 1,
                        'price': '200.33',
                        'total_price': '200.33',
                    },
                ]
            }
        }
        row = {
            'description': 'Cold Side Intercooler Pipe for 2017-2026 Ford Super Duty, 6.7L Powerstroke',
            'qty': '1',
            'rate': '200.33',
            'amount': '200.33',
        }

        sku, line_id = infer_invoice_row_sku_from_po(po_details, row)

        self.assertEqual(sku, '83-1001')
        self.assertEqual(line_id, 'line-2')

    def test_infers_missing_sb_sku_respects_used_line_items(self):
        po_details = {
            'lineItems': {
                'rows': [
                    {
                        'id': 'line-1',
                        'product': {'sku': '83-2004', 'name': 'Hot Side Intercooler Pipe'},
                        'quantity': 1,
                        'price': '166.83',
                        'total_price': '166.83',
                    }
                ]
            }
        }
        row = {
            'description': 'Hot Side Intercooler Pipe',
            'qty': '1',
            'rate': '166.83',
            'amount': '166.83',
        }

        sku, line_id = infer_invoice_row_sku_from_po(po_details, row, {'line-1'})

        self.assertEqual(sku, '')
        self.assertEqual(line_id, '')


if __name__ == '__main__':
    unittest.main()
