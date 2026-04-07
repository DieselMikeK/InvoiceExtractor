import os
import unittest

from invoice_parser import detect_vendor, extract_text_from_pdf, infer_vendor_from_sender, parse_invoice


class SerraInvoiceParserTests(unittest.TestCase):
    def test_serra_does_not_self_identify_from_generic_layout(self):
        pdf_path = os.path.join(
            os.path.dirname(__file__),
            'training',
            'ST',
            '2144553-Customer-Copy.pdf',
        )

        text = extract_text_from_pdf(pdf_path)

        self.assertNotEqual(detect_vendor(text), 'Serra Chrysler Dodge Ram Jeep of Traverse City')

    def test_serra_sender_alias_matches_serratc_domain(self):
        self.assertEqual(
            infer_vendor_from_sender(sender_email='parts@serratc.com'),
            'Serra Chrysler Dodge Ram Jeep of Traverse City',
        )

    def test_serra_parse_uses_credit_card_terms_and_drops_bad_vendor_address(self):
        pdf_path = os.path.join(
            os.path.dirname(__file__),
            'training',
            'ST',
            '2144553-Customer-Copy.pdf',
        )

        invoice_data = parse_invoice(pdf_path)

        self.assertEqual(invoice_data['vendor'], 'Serra Chrysler Dodge Ram Jeep of Traverse City')
        self.assertEqual(invoice_data['terms'], 'Credit Card')
        self.assertEqual(invoice_data['vendor_address'], '')


if __name__ == '__main__':
    unittest.main()
