"""Shopify Admin API client with OAuth support for desktop validation flows."""

import json
import os
import re
import secrets
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
import webbrowser
from http.server import BaseHTTPRequestHandler, HTTPServer

try:
    from core_detection import is_core_candidate
except ImportError:
    from app.core_detection import is_core_candidate


DEFAULT_SCOPES = ['read_orders']
DEFAULT_API_VERSION = '2025-10'
CALLBACK_HOST = '127.0.0.1'
CALLBACK_PORT = 8765
CALLBACK_PATH = '/shopify/callback'
OAUTH_TIMEOUT_SECONDS = 300


def _normalize_shop_domain(value):
    shop = str(value or '').strip()
    if not shop:
        return ''
    shop = re.sub(r'^https?://', '', shop, flags=re.IGNORECASE)
    shop = shop.split('/')[0].strip().lower()
    return shop


def _normalize_po_digits(value):
    digits = ''.join(ch for ch in str(value or '') if ch.isdigit())
    if digits == '':
        return ''
    return digits.lstrip('0') or '0'


def _to_float(value):
    if value is None:
        return None
    try:
        text = str(value).strip().replace(',', '').replace('$', '')
        if text == '':
            return None
        return float(text)
    except (TypeError, ValueError):
        return None


def _text_contains_po(text, po_norm):
    if not text or not po_norm:
        return False
    digit_groups = re.findall(r'\d+', str(text))
    for group in digit_groups:
        if _normalize_po_digits(group) == po_norm:
            return True
    return False


def _is_core_line_item(sku, name):
    return is_core_candidate('', sku, name)


class ShopifyClient:
    def __init__(
        self,
        shop,
        client_id,
        client_secret,
        scopes=None,
        api_version=DEFAULT_API_VERSION,
        token_file=None,
        status_callback=None,
        auth_mode='auto',
        callback_host=CALLBACK_HOST,
        redirect_uri=None,
        callback_bind_host=None,
        callback_bind_port=None,
        callback_port=CALLBACK_PORT,
        callback_path=CALLBACK_PATH,
        oauth_timeout=OAUTH_TIMEOUT_SECONDS
    ):
        self.shop = _normalize_shop_domain(shop)
        self.client_id = str(client_id or '').strip()
        self.client_secret = str(client_secret or '').strip()
        self.scopes = self._normalize_scopes(scopes)
        self.api_version = str(api_version or DEFAULT_API_VERSION).strip()
        self.token_file = token_file
        self.status_callback = status_callback or (lambda msg, tag=None: None)
        mode = str(auth_mode or 'auto').strip().lower()
        self.auth_mode = mode if mode in {'auto', 'oauth', 'client_credentials'} else 'auto'
        self.callback_host = str(callback_host or CALLBACK_HOST).strip() or CALLBACK_HOST
        self.callback_port = int(callback_port)
        self.callback_path = (
            callback_path if str(callback_path).startswith('/') else f"/{callback_path}"
        )

        raw_redirect_uri = str(redirect_uri or '').strip()
        self._explicit_redirect_uri = False
        if raw_redirect_uri:
            try:
                parsed = urllib.parse.urlparse(raw_redirect_uri)
                if parsed.scheme and parsed.netloc:
                    self.callback_host = parsed.hostname or self.callback_host
                    if parsed.port:
                        self.callback_port = int(parsed.port)
                    redirect_path = str(parsed.path or '').strip() or self.callback_path
                    self.callback_path = (
                        redirect_path if redirect_path.startswith('/') else f"/{redirect_path}"
                    )
                    self.callback_url = (
                        f"{parsed.scheme}://{parsed.netloc}{self.callback_path}"
                    )
                    self._explicit_redirect_uri = True
                else:
                    self.callback_url = (
                        f"http://{self.callback_host}:{self.callback_port}{self.callback_path}"
                    )
            except Exception:
                self.callback_url = (
                    f"http://{self.callback_host}:{self.callback_port}{self.callback_path}"
                )
        else:
            self.callback_url = (
                f"http://{self.callback_host}:{self.callback_port}{self.callback_path}"
            )

        self.callback_bind_host = str(
            callback_bind_host or self.callback_host
        ).strip() or self.callback_host
        self.callback_bind_port = int(callback_bind_port or self.callback_port)
        self.oauth_timeout = int(oauth_timeout)
        self.access_token = ''

    def _normalize_scopes(self, scopes):
        if scopes is None:
            return list(DEFAULT_SCOPES)
        if isinstance(scopes, str):
            parts = re.split(r'[,\s]+', scopes)
            vals = [p.strip() for p in parts if p.strip()]
            return vals or list(DEFAULT_SCOPES)
        vals = [str(s).strip() for s in scopes if str(s).strip()]
        return vals or list(DEFAULT_SCOPES)

    def _status(self, message, tag='info'):
        self.status_callback(message, tag)

    def _request_json(self, url, method='GET', data=None, headers=None, timeout=30):
        req_headers = dict(headers or {})
        body = None
        if data is not None:
            body = json.dumps(data).encode('utf-8')
            req_headers.setdefault('Content-Type', 'application/json')
        request = urllib.request.Request(url, data=body, headers=req_headers, method=method)
        try:
            with urllib.request.urlopen(request, timeout=timeout) as response:
                payload = response.read().decode('utf-8')
        except urllib.error.HTTPError as e:
            err_body = ''
            try:
                err_body = e.read().decode('utf-8', errors='replace')
            except Exception:
                err_body = ''
            detail = f"HTTP {e.code}"
            if err_body:
                detail = f"{detail} - {err_body}"
            return None, detail
        except Exception as e:
            return None, str(e)

        try:
            return json.loads(payload), None
        except json.JSONDecodeError:
            return None, "Invalid JSON response from Shopify"

    def _request_form(self, url, form_data, headers=None, timeout=30):
        req_headers = dict(headers or {})
        req_headers.setdefault('Content-Type', 'application/x-www-form-urlencoded')
        encoded = urllib.parse.urlencode(form_data or {}).encode('utf-8')
        request = urllib.request.Request(url, data=encoded, headers=req_headers, method='POST')
        try:
            with urllib.request.urlopen(request, timeout=timeout) as response:
                payload = response.read().decode('utf-8')
        except urllib.error.HTTPError as e:
            err_body = ''
            try:
                err_body = e.read().decode('utf-8', errors='replace')
            except Exception:
                err_body = ''
            detail = f"HTTP {e.code}"
            if err_body:
                detail = f"{detail} - {err_body}"
            return None, detail
        except Exception as e:
            return None, str(e)

        try:
            return json.loads(payload), None
        except json.JSONDecodeError:
            return None, "Invalid JSON response from Shopify"

    def _load_token(self):
        if not self.token_file or not os.path.exists(self.token_file):
            return None
        try:
            with open(self.token_file, 'r', encoding='utf-8-sig') as f:
                data = json.load(f)
        except Exception:
            return None
        if not isinstance(data, dict):
            return None
        if _normalize_shop_domain(data.get('shop')) != self.shop:
            return None
        token = str(data.get('access_token', '')).strip()
        if not token:
            return None
        return data

    def _save_token(self, token_data):
        if not self.token_file:
            return
        token_dir = os.path.dirname(self.token_file)
        if token_dir:
            os.makedirs(token_dir, exist_ok=True)
        with open(self.token_file, 'w', encoding='utf-8') as f:
            json.dump(token_data, f, indent=2)

    def _format_graphql_errors(self, errors):
        if not errors:
            return ''
        messages = []
        for err in errors:
            if isinstance(err, dict):
                msg = str(err.get('message', '')).strip()
                if msg:
                    messages.append(msg)
            else:
                text = str(err).strip()
                if text:
                    messages.append(text)
        return '; '.join(messages)

    def graphql(self, query, variables=None):
        if not self.access_token:
            return None, "Shopify access token is not available"
        url = f"https://{self.shop}/admin/api/{self.api_version}/graphql.json"
        payload = {'query': query, 'variables': variables or {}}
        headers = {
            'X-Shopify-Access-Token': self.access_token,
            'Content-Type': 'application/json',
        }
        data, error = self._request_json(url, method='POST', data=payload, headers=headers, timeout=45)
        if error:
            return None, error

        gql_errors = data.get('errors')
        if gql_errors:
            return None, self._format_graphql_errors(gql_errors) or "GraphQL request failed"
        return data.get('data') or {}, None

    def test_connection(self):
        data, error = self.graphql("query { shop { name myshopifyDomain } }")
        if error:
            return False, error
        shop_data = (data or {}).get('shop') or {}
        shop_name = shop_data.get('name') or shop_data.get('myshopifyDomain') or self.shop
        return True, shop_name

    def _exchange_code_for_token(self, code):
        url = f"https://{self.shop}/admin/oauth/access_token"
        payload = {
            'client_id': self.client_id,
            'client_secret': self.client_secret,
            'code': code,
        }
        data, error = self._request_json(url, method='POST', data=payload, timeout=45)
        if error:
            return None, error
        access_token = str((data or {}).get('access_token', '')).strip()
        if not access_token:
            return None, "Shopify OAuth response missing access_token"
        return data, None

    def _exchange_client_credentials_for_token(self):
        url = f"https://{self.shop}/admin/oauth/access_token"
        payload = {
            'grant_type': 'client_credentials',
            'client_id': self.client_id,
            'client_secret': self.client_secret,
        }
        data, error = self._request_form(url, payload, timeout=45)
        if error:
            return None, error
        access_token = str((data or {}).get('access_token', '')).strip()
        if not access_token:
            return None, "Client credentials response missing access_token"
        return data, None

    def _run_oauth_flow_once(self, callback_host, timeout_seconds):
        if self._explicit_redirect_uri:
            callback_url = self.callback_url
        else:
            callback_url = f"http://{callback_host}:{self.callback_port}{self.callback_path}"
        state = secrets.token_urlsafe(24)
        params = {
            'client_id': self.client_id,
            'scope': ','.join(self.scopes),
            'redirect_uri': callback_url,
            'state': state,
        }
        auth_url = f"https://{self.shop}/admin/oauth/authorize?{urllib.parse.urlencode(params)}"
        oauth_result = {}
        oauth_event = threading.Event()
        callback_path = self.callback_path

        class CallbackHandler(BaseHTTPRequestHandler):
            def do_GET(self):
                parsed = urllib.parse.urlparse(self.path)
                if parsed.path != callback_path:
                    self.send_response(404)
                    self.end_headers()
                    self.wfile.write(b'Not Found')
                    return

                query = urllib.parse.parse_qs(parsed.query)
                oauth_result['code'] = (query.get('code') or [''])[0]
                oauth_result['state'] = (query.get('state') or [''])[0]
                oauth_result['error'] = (query.get('error') or [''])[0]
                oauth_result['error_description'] = (query.get('error_description') or [''])[0]
                oauth_event.set()

                self.send_response(200)
                self.send_header('Content-Type', 'text/html; charset=utf-8')
                self.end_headers()
                self.wfile.write(
                    b"<html><body><h3>Shopify authorization complete.</h3>"
                    b"You can close this window and return to Invoice Extractor.</body></html>"
                )

            def log_message(self, fmt, *args):
                return

        try:
            server = HTTPServer((self.callback_bind_host, self.callback_bind_port), CallbackHandler)
        except OSError as e:
            return False, (
                f"Could not start local callback server on "
                f"{self.callback_bind_host}:{self.callback_bind_port}: {e}"
            )

        try:
            threading.Thread(target=server.handle_request, daemon=True).start()
            self._status(
                f"Shopify callback URL: {callback_url}. "
                "If OAuth shows redirect/app URL host mismatch, update Shopify app settings "
                "so App URL host and redirect host match exactly.",
                "info"
            )
            self._status("Opening browser for Shopify authentication...", "info")
            browser_opened = webbrowser.open(auth_url)
            if not browser_opened:
                self._status(
                    f"Could not open browser automatically. Open this URL manually: {auth_url}",
                    "warning"
                )
            deadline = time.time() + timeout_seconds
            while time.time() < deadline:
                if oauth_event.wait(timeout=1):
                    break
            if not oauth_event.is_set():
                return False, (
                    f"Timed out waiting for Shopify OAuth callback at {callback_url}. "
                    "If browser showed a redirect/app URL host mismatch, set Shopify App URL host "
                    "to match callback_host and add this exact callback URL to allowed redirects."
                )

            error_val = str(oauth_result.get('error', '')).strip()
            if error_val:
                desc = str(oauth_result.get('error_description', '')).strip()
                return False, f"{error_val}: {desc}" if desc else error_val

            returned_state = str(oauth_result.get('state', '')).strip()
            if returned_state != state:
                return False, "OAuth state mismatch; authorization was rejected"

            code = str(oauth_result.get('code', '')).strip()
            if not code:
                return False, "OAuth callback did not include an authorization code"

            token_payload, error = self._exchange_code_for_token(code)
            if error:
                return False, f"Token exchange failed: {error}"

            self.access_token = str(token_payload.get('access_token', '')).strip()
            self._save_token({
                'shop': self.shop,
                'access_token': self.access_token,
                'scope': token_payload.get('scope', ''),
                'created_at': int(time.time()),
                'api_version': self.api_version,
            })
            return True, None
        finally:
            try:
                server.server_close()
            except Exception:
                pass

    def _run_oauth_flow(self):
        callback_hosts = [self.callback_host]
        if (not self._explicit_redirect_uri) and self.callback_host in {'localhost', '127.0.0.1'}:
            alt_host = '127.0.0.1' if self.callback_host == 'localhost' else 'localhost'
            callback_hosts.append(alt_host)

        if len(callback_hosts) > 1:
            per_host_timeout = max(45, int(self.oauth_timeout / len(callback_hosts)))
        else:
            per_host_timeout = self.oauth_timeout

        last_error = None
        for idx, host in enumerate(callback_hosts, start=1):
            if len(callback_hosts) > 1:
                self._status(
                    f"Starting Shopify OAuth attempt {idx}/{len(callback_hosts)} "
                    f"using callback host '{host}'...",
                    "info"
                )
            ok, message = self._run_oauth_flow_once(host, per_host_timeout)
            if ok:
                return True, message
            last_error = message
            if idx < len(callback_hosts):
                self._status(
                    f"OAuth attempt with callback host '{host}' failed: {message}. "
                    f"Retrying with '{callback_hosts[idx]}'...",
                    "warning"
                )

        return False, last_error or "Shopify OAuth failed"

    def authenticate(self):
        if not self.shop:
            return False, "Shopify config missing 'shop' (store domain)"
        if not self.client_id or not self.client_secret:
            return False, "Shopify config missing client_id/client_secret"

        token_data = self._load_token()
        if token_data:
            self.access_token = str(token_data.get('access_token', '')).strip()
            ok, message = self.test_connection()
            if ok:
                return True, message
            self._status(f"Cached Shopify token invalid, re-auth required: {message}", "warning")
            self.access_token = ''

        cc_error = None
        if self.auth_mode in {'auto', 'client_credentials'}:
            token_payload, cc_error = self._exchange_client_credentials_for_token()
            if token_payload:
                self.access_token = str(token_payload.get('access_token', '')).strip()
                self._save_token({
                    'shop': self.shop,
                    'access_token': self.access_token,
                    'scope': token_payload.get('scope', ''),
                    'created_at': int(time.time()),
                    'api_version': self.api_version,
                    'auth_mode': 'client_credentials',
                    'expires_in': token_payload.get('expires_in'),
                })
                ok, message = self.test_connection()
                if ok:
                    return True, message
                cc_error = message
                self.access_token = ''
            if self.auth_mode == 'client_credentials':
                return False, (
                    f"Client credentials auth failed: {cc_error or 'unknown error'}. "
                    "Set auth_mode to 'auto' or 'oauth' to allow browser sign-in."
                )

        if self.auth_mode in {'auto', 'oauth'}:
            ok, message = self._run_oauth_flow()
            if not ok:
                if cc_error:
                    return False, f"Client credentials failed ({cc_error}); OAuth failed ({message})"
                return False, message

            ok, message = self.test_connection()
            if not ok:
                return False, message
            return True, message

        return False, cc_error or "Shopify authentication failed"

    def fetch_orders_for_query(self, query_text, max_pages=4, page_size=50):
        gql = """
        query Orders($query: String!, $cursor: String, $pageSize: Int!) {
          orders(first: $pageSize, query: $query, after: $cursor, sortKey: CREATED_AT, reverse: true) {
            pageInfo {
              hasNextPage
              endCursor
            }
            edges {
              node {
                id
                name
                note
                tags
                customAttributes {
                  key
                  value
                }
                lineItems(first: 100) {
                  edges {
                    node {
                      sku
                      name
                      quantity
                      originalUnitPriceSet {
                        shopMoney {
                          amount
                          currencyCode
                        }
                      }
                    }
                  }
                }
              }
            }
          }
        }
        """
        orders = []
        cursor = None
        page_count = 0
        while page_count < max_pages:
            page_count += 1
            variables = {
                'query': str(query_text or '').strip(),
                'cursor': cursor,
                'pageSize': int(page_size),
            }
            data, error = self.graphql(gql, variables)
            if error:
                return None, error
            orders_data = (data or {}).get('orders') or {}
            edges = orders_data.get('edges') or []
            for edge in edges:
                node = edge.get('node') if isinstance(edge, dict) else None
                if node:
                    orders.append(node)
            page_info = orders_data.get('pageInfo') or {}
            if not page_info.get('hasNextPage'):
                break
            cursor = page_info.get('endCursor')
            if not cursor:
                break
        return orders, None

    def _order_contains_po(self, order, po_norm):
        if not po_norm:
            return False
        text_parts = []
        for key in ('name', 'note'):
            text_parts.append(order.get(key, ''))
        tags = order.get('tags')
        if tags:
            text_parts.append(' '.join(tags if isinstance(tags, list) else [str(tags)]))
        attrs = order.get('customAttributes') or []
        for item in attrs:
            if not isinstance(item, dict):
                continue
            text_parts.append(item.get('key', ''))
            text_parts.append(item.get('value', ''))

        for value in text_parts:
            if _text_contains_po(value, po_norm):
                return True
        return False

    def _order_name_matches_number(self, order, target_norm):
        if not target_norm:
            return False
        name = str((order or {}).get('name', '')).strip()
        return _text_contains_po(name, target_norm)

    def find_orders_for_po(self, po_number):
        po_text = str(po_number or '').strip()
        if po_text.upper().startswith('PO'):
            po_text = po_text[2:].strip()
        if not po_text:
            return [], None

        queries = [
            f"po_number:{po_text}",
            f"\"{po_text}\"",
            po_text,
        ]
        seen_ids = set()
        combined = []
        last_error = None

        for query_text in queries:
            orders, error = self.fetch_orders_for_query(query_text)
            if error:
                last_error = error
                continue
            if not orders:
                continue
            for order in orders:
                order_id = str(order.get('id', '')).strip()
                if order_id and order_id in seen_ids:
                    continue
                if order_id:
                    seen_ids.add(order_id)
                combined.append(order)
            if combined:
                break

        if not combined:
            return [], last_error

        po_norm = _normalize_po_digits(po_text)
        strict_matches = [o for o in combined if self._order_contains_po(o, po_norm)]
        return strict_matches or combined, None

    def find_orders_for_order_number(self, order_number):
        order_text = str(order_number or '').strip()
        if order_text.startswith('#'):
            order_text = order_text[1:].strip()
        if not order_text:
            return [], None

        queries = [
            f"name:{order_text}",
            f"name:#{order_text}",
            f"\"{order_text}\"",
            order_text,
        ]
        seen_ids = set()
        combined = []
        last_error = None

        for query_text in queries:
            orders, error = self.fetch_orders_for_query(query_text)
            if error:
                last_error = error
                continue
            if not orders:
                continue
            for order in orders:
                order_id = str(order.get('id', '')).strip()
                if order_id and order_id in seen_ids:
                    continue
                if order_id:
                    seen_ids.add(order_id)
                combined.append(order)
            if combined:
                break

        if not combined:
            return [], last_error

        target_norm = _normalize_po_digits(order_text)
        name_matches = [o for o in combined if self._order_name_matches_number(o, target_norm)]
        if name_matches:
            return name_matches, None

        strict_matches = [o for o in combined if self._order_contains_po(o, target_norm)]
        return strict_matches or combined, None

    def extract_core_amounts_from_orders(self, orders):
        values = []
        for order in orders or []:
            line_items = ((order.get('lineItems') or {}).get('edges') or [])
            for edge in line_items:
                line = edge.get('node') if isinstance(edge, dict) else None
                if not line:
                    continue
                sku = line.get('sku', '')
                name = line.get('name', '')
                if not _is_core_line_item(sku, name):
                    continue
                shop_money = ((line.get('originalUnitPriceSet') or {}).get('shopMoney') or {})
                amount = _to_float(shop_money.get('amount'))
                if amount is None:
                    continue
                values.append(round(amount, 2))
        return values

    def get_po_core_amounts(self, po_number):
        orders, error = self.find_orders_for_po(po_number)
        if error:
            return None, error
        core_amounts = self.extract_core_amounts_from_orders(orders)
        return {
            'orders': orders,
            'core_amounts': core_amounts,
        }, None

    def get_order_number_core_amounts(self, order_number):
        orders, error = self.find_orders_for_order_number(order_number)
        if error:
            return None, error
        core_amounts = self.extract_core_amounts_from_orders(orders)
        return {
            'orders': orders,
            'core_amounts': core_amounts,
        }, None
