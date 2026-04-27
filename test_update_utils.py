import base64
import json
import unittest
from unittest.mock import patch

from update_utils import (
    DEFAULT_UPDATE_MANIFEST_URL,
    decode_release_manifest_payload,
    fetch_release_manifest,
)


class DecodeReleaseManifestPayloadTests(unittest.TestCase):
    def test_decodes_github_contents_api_payload(self):
        manifest = {
            'version': '1.2.27',
            'download_url': 'https://example.com/InvoiceExtractor.exe',
            'sha256': 'a' * 64,
            'notes': 'Test release',
            'files': [
                {
                    'relative_path': 'InvoiceExtractor.exe',
                    'download_url': 'https://example.com/InvoiceExtractor.exe',
                    'sha256': 'a' * 64,
                }
            ],
        }
        payload = {
            'encoding': 'base64',
            'content': base64.b64encode(
                json.dumps(manifest).encode('utf-8')
            ).decode('ascii'),
        }

        decoded = decode_release_manifest_payload(payload)

        self.assertEqual(decoded['version'], '1.2.27')
        self.assertEqual(decoded['download_url'], manifest['download_url'])


class FakeHeaders:
    def get_content_charset(self):
        return 'utf-8'


class FakeResponse:
    def __init__(self, payload):
        self.payload = payload.encode('utf-8')
        self.headers = FakeHeaders()

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        return False

    def read(self):
        return self.payload


class FetchReleaseManifestTests(unittest.TestCase):
    def test_uses_api_manifest_by_default(self):
        manifest = {
            'version': '1.2.27',
            'download_url': 'https://example.com/InvoiceExtractor.exe',
            'sha256': 'b' * 64,
            'notes': 'API manifest',
            'files': [
                {
                    'relative_path': 'InvoiceExtractor.exe',
                    'download_url': 'https://example.com/InvoiceExtractor.exe',
                    'sha256': 'b' * 64,
                }
            ],
        }
        payload = {
            'encoding': 'base64',
            'content': base64.b64encode(
                json.dumps(manifest).encode('utf-8')
            ).decode('ascii'),
        }

        with patch('update_utils.get_update_manifest_url', return_value=DEFAULT_UPDATE_MANIFEST_URL):
            with patch(
                'update_utils.open_url_with_tls_fallback',
                return_value=FakeResponse(json.dumps(payload)),
            ) as mocked_open:
                parsed = fetch_release_manifest(timeout=1)

        request = mocked_open.call_args[0][0]
        self.assertIn('api.github.com/repos/', request.full_url)
        self.assertEqual(request.headers.get('Accept'), 'application/vnd.github+json')
        self.assertEqual(parsed['version'], '1.2.27')
        self.assertEqual(parsed['download_url'], manifest['download_url'])


if __name__ == '__main__':
    unittest.main()
