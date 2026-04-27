import json
import os
import tempfile
import unittest
import base64

from gmail_client import GmailClient, _message_matches_time_filter


SB_BODY = """---------- Forwarded message ---------
From: S&B <store+69841617189@t.shopifyemail.com>
Date: Mon, Apr 27, 2026 at 2:30 PM
Subject: Order #743234 Confirmed
To: <ap@dieselpowerproducts.com>

ORDER #743234
PO NUMBER #0064464
Hi Power Products. Thank you for your purchase!
Order summary
Cold Air Intake for 2006-2007 Chevy / GMC Duramax LLY-LBZ 6.6L x 1 $253.93
Dry Extendable
Subtotal $253.93
Shipping $12.00
Taxes $0.00
Total paid today $0.00 USD
Total due May 27, 2026 $265.93 USD
Customer information
Shipping address
Donald Ortiz
5041 Brighton Hills Dr NE
Rio Rancho NM 87144
Billing address
Josh Ulrich
Diesel Power Products DBA Power Products Unlimited, Inc. 505
5204 East Broadway Avenue
Spokane Valley WA 99212
Payment
Net 30: Due May 27, 2026
Shipping method
Ground
If you have any questions, reply to this email or contact us at
customerservice@sbfilters.com
"""


class GmailClientDownloadTests(unittest.TestCase):
    def test_message_time_filter_allows_open_ended_ranges(self):
        message = {'internalDate': '1712520000000'}

        self.assertTrue(
            _message_matches_time_filter(message, {'start_ts': 1712510000, 'end_ts': None})
        )
        self.assertTrue(
            _message_matches_time_filter(message, {'start_ts': None, 'end_ts': 1712523600})
        )
        self.assertFalse(
            _message_matches_time_filter(message, {'start_ts': 1712523600, 'end_ts': None})
        )

    def test_labels_only_after_all_attachments_download(self):
        stop_state = {'stop': False}

        with tempfile.TemporaryDirectory() as tmpdir:
            client = GmailClient(
                tmpdir,
                data_dir=tmpdir,
                invoices_dir=tmpdir,
                should_stop=lambda: stop_state['stop'],
            )
            client.processed_label_id = 'label-1'

            labeled = []

            client.fetch_all_message_ids = lambda query=None: [{'id': 'msg-1'}]
            client.get_message_details = lambda msg_id: {
                'payload': {
                    'headers': [
                        {'name': 'Subject', 'value': 'Test'},
                        {'name': 'From', 'value': 'KC Turbos <invoicing@kcturbos.com>'},
                    ],
                    'parts': [
                        {'filename': 'a.pdf', 'body': {'attachmentId': 'att-1'}},
                        {'filename': 'b.pdf', 'body': {'attachmentId': 'att-2'}},
                    ],
                }
            }
            client.find_attachments_in_parts = lambda parts, msg_id: [
                {'filename': 'a.pdf', 'attachment_id': 'att-1', 'msg_id': msg_id},
                {'filename': 'b.pdf', 'attachment_id': 'att-2', 'msg_id': msg_id},
            ]

            def fake_download(msg_id, attachment_id, filename):
                if attachment_id == 'att-1':
                    stop_state['stop'] = True
                return filename

            client.download_attachment = fake_download
            client._add_label_to_message = lambda msg_id, label_id: labeled.append((msg_id, label_id))

            downloaded, total_emails, new_emails = client.fetch_and_download_new_attachments()

        self.assertEqual(total_emails, 1)
        self.assertEqual(new_emails, 1)
        self.assertEqual(len(downloaded), 1)
        self.assertEqual(labeled, [])

    def test_download_error_does_not_label_message(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            client = GmailClient(tmpdir, data_dir=tmpdir, invoices_dir=tmpdir)
            client.processed_label_id = 'label-1'

            labeled = []

            client.fetch_all_message_ids = lambda query=None: [{'id': 'msg-1'}]
            client.get_message_details = lambda msg_id: {
                'payload': {
                    'headers': [
                        {'name': 'Subject', 'value': 'Test'},
                        {'name': 'From', 'value': 'KC Turbos <invoicing@kcturbos.com>'},
                    ],
                    'parts': [
                        {'filename': 'a.pdf', 'body': {'attachmentId': 'att-1'}},
                    ],
                }
            }
            client.find_attachments_in_parts = lambda parts, msg_id: [
                {'filename': 'a.pdf', 'attachment_id': 'att-1', 'msg_id': msg_id},
            ]
            client.download_attachment = lambda msg_id, attachment_id, filename: (_ for _ in ()).throw(
                RuntimeError('download failed')
            )
            client._add_label_to_message = lambda msg_id, label_id: labeled.append((msg_id, label_id))

            downloaded, total_emails, new_emails = client.fetch_and_download_new_attachments()

        self.assertEqual(downloaded, [])
        self.assertEqual(total_emails, 1)
        self.assertEqual(new_emails, 1)
        self.assertEqual(labeled, [])

    def test_forwarded_message_sender_overrides_outer_from_header(self):
        forwarded_body = (
            "---------- Forwarded message ---------\n"
            "From: KC Turbos Invoicing (invoicing@kcturbos.com) "
            "<system@sent-via.netsuite.com>\n"
            "Subject: KC Turbos Invoice(s) Attached\n"
            "Date: Tue, Apr 7, 2026 at 9:00 AM\n"
            "To: Accounts Payable <ap@dieselpowerproducts.com>\n"
        )
        encoded_body = base64.urlsafe_b64encode(
            forwarded_body.encode('utf-8')
        ).decode('ascii').rstrip('=')

        with tempfile.TemporaryDirectory() as tmpdir:
            client = GmailClient(tmpdir, data_dir=tmpdir, invoices_dir=tmpdir)
            client.processed_label_id = 'label-1'

            labeled = []

            client.fetch_all_message_ids = lambda query=None: [{'id': 'msg-1'}]
            client.get_message_details = lambda msg_id: {
                'snippet': 'Forwarded message from KC Turbos Invoicing',
                'payload': {
                    'headers': [
                        {
                            'name': 'Subject',
                            'value': 'Fwd: KC invoice',
                        },
                        {
                            'name': 'From',
                            'value': 'Accounts Payable <ap@dieselpowerproducts.com>',
                        },
                    ],
                    'parts': [
                        {
                            'mimeType': 'text/plain',
                            'body': {'data': encoded_body},
                        },
                        {
                            'filename': 'kc.pdf',
                            'body': {'attachmentId': 'att-1'},
                        },
                    ],
                }
            }
            client.find_attachments_in_parts = lambda parts, msg_id: [
                {'filename': 'kc.pdf', 'attachment_id': 'att-1', 'msg_id': msg_id},
            ]
            client.download_attachment = lambda msg_id, attachment_id, filename: filename
            client._add_label_to_message = lambda msg_id, label_id: labeled.append((msg_id, label_id))

            downloaded, total_emails, new_emails = client.fetch_and_download_new_attachments()

        self.assertEqual(total_emails, 1)
        self.assertEqual(new_emails, 1)
        self.assertEqual(labeled, [('msg-1', 'label-1')])
        self.assertEqual(len(downloaded), 1)
        self.assertEqual(downloaded[0]['sender_email'], 'invoicing@kcturbos.com')
        self.assertIn('KC Turbos Invoicing', downloaded[0]['sender_header'])
        self.assertEqual(downloaded[0]['subject'], 'KC Turbos Invoice(s) Attached')
        self.assertIn('KC Turbos Invoice(s) Attached', downloaded[0]['message_text'])
        self.assertIn('Forwarded message', downloaded[0]['message_text'])

    def test_message_time_filter_skips_messages_outside_requested_window(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            client = GmailClient(tmpdir, data_dir=tmpdir, invoices_dir=tmpdir)
            client.processed_label_id = 'label-1'

            labeled = []

            client.fetch_all_message_ids = lambda query=None: [{'id': 'msg-1'}]
            client.get_message_details = lambda msg_id: {
                'internalDate': '1712520000000',
                'payload': {
                    'headers': [
                        {'name': 'Subject', 'value': 'Test'},
                        {'name': 'From', 'value': 'KC Turbos <invoicing@kcturbos.com>'},
                    ],
                    'parts': [
                        {'filename': 'a.pdf', 'body': {'attachmentId': 'att-1'}},
                    ],
                }
            }
            client.find_attachments_in_parts = lambda parts, msg_id: [
                {'filename': 'a.pdf', 'attachment_id': 'att-1', 'msg_id': msg_id},
            ]
            client.download_attachment = lambda msg_id, attachment_id, filename: filename
            client._add_label_to_message = lambda msg_id, label_id: labeled.append((msg_id, label_id))

            downloaded, total_emails, new_emails = client.fetch_and_download_new_attachments(
                query='after:1712510000 before:1712515000',
                message_time_filter={'start_ts': 1712510000, 'end_ts': 1712515000},
            )

        self.assertEqual(downloaded, [])
        self.assertEqual(total_emails, 1)
        self.assertEqual(new_emails, 1)
        self.assertEqual(labeled, [])

    def test_no_attachment_sb_body_invoice_is_saved_and_labeled(self):
        encoded_body = base64.urlsafe_b64encode(
            SB_BODY.encode('utf-8')
        ).decode('ascii').rstrip('=')

        with tempfile.TemporaryDirectory() as tmpdir:
            client = GmailClient(tmpdir, data_dir=tmpdir, invoices_dir=tmpdir)
            client.processed_label_id = 'label-1'

            labeled = []

            client.fetch_all_message_ids = lambda query=None: [{'id': 'msg-sb'}]
            client.get_message_details = lambda msg_id: {
                'payload': {
                    'headers': [
                        {'name': 'Subject', 'value': 'Fwd: Order #743234 Confirmed'},
                        {'name': 'From', 'value': 'Accounts Payable <ap@dieselpowerproducts.com>'},
                    ],
                    'parts': [
                        {
                            'mimeType': 'text/plain',
                            'body': {'data': encoded_body},
                        },
                    ],
                }
            }
            client._add_label_to_message = lambda msg_id, label_id: labeled.append((msg_id, label_id))

            downloaded, total_emails, new_emails = client.fetch_and_download_new_attachments()

            saved_path = os.path.join(tmpdir, downloaded[0]['filename'])
            with open(saved_path, 'r', encoding='utf-8') as f:
                saved_payload = json.load(f)

        self.assertEqual(total_emails, 1)
        self.assertEqual(new_emails, 1)
        self.assertEqual(labeled, [('msg-sb', 'label-1')])
        self.assertEqual(len(downloaded), 1)
        self.assertEqual(downloaded[0]['filename'], 'SB_Order_743234.email.json')
        self.assertTrue(downloaded[0]['email_body_invoice'])
        self.assertEqual(saved_payload['parser'], 'sb_shopify_order')
        self.assertIn('customerservice@sbfilters.com', saved_payload['message_text'])

    def test_sb_body_text_with_attachment_is_not_saved_as_body_invoice(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            client = GmailClient(tmpdir, data_dir=tmpdir, invoices_dir=tmpdir)
            client.processed_label_id = 'label-1'

            client.fetch_all_message_ids = lambda query=None: [{'id': 'msg-sb'}]
            client.get_message_details = lambda msg_id: {
                'payload': {
                    'headers': [
                        {'name': 'Subject', 'value': 'Fwd: Order #743234 Confirmed'},
                        {'name': 'From', 'value': 'Accounts Payable <ap@dieselpowerproducts.com>'},
                    ],
                    'parts': [
                        {'filename': 'daystar.pdf', 'body': {'attachmentId': 'att-1'}},
                    ],
                }
            }
            client.find_attachments_in_parts = lambda parts, msg_id: [
                {'filename': 'daystar.pdf', 'attachment_id': 'att-1', 'msg_id': msg_id},
            ]
            client.download_attachment = lambda msg_id, attachment_id, filename: filename
            client._add_label_to_message = lambda msg_id, label_id: None

            downloaded, _, _ = client.fetch_and_download_new_attachments()

        self.assertEqual(len(downloaded), 1)
        self.assertEqual(downloaded[0]['filename'], 'daystar.pdf')
        self.assertNotIn('email_body_invoice', downloaded[0])


if __name__ == '__main__':
    unittest.main()
