import os
import tempfile
import unittest
import base64

from gmail_client import GmailClient


class GmailClientDownloadTests(unittest.TestCase):
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


if __name__ == '__main__':
    unittest.main()
