"""Slim googleapiclient discovery cache to APIs used by this app.

By default, PyInstaller bundles every static discovery document shipped with
google-api-python-client, which adds significant size. InvoiceExtractor only
calls Gmail v1.
"""

from PyInstaller.utils.hooks import collect_data_files

# Keep only Gmail discovery document required by build('gmail', 'v1').
datas = collect_data_files(
    'googleapiclient',
    includes=['discovery_cache/documents/gmail.v1.json'],
)
