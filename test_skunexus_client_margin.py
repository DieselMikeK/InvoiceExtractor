import unittest
from unittest.mock import patch

from skunexus_client import SkuNexusClient


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


if __name__ == '__main__':
    unittest.main()
