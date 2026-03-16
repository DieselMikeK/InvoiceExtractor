import re


_CORE_TOKEN_RE = re.compile(r'(?<![a-z0-9])core(?![a-z0-9])', re.IGNORECASE)
_CORE_DESCRIPTION_MARKERS = (
    'core charge',
    'conditional core',
    'refundable core',
    'rebuildable core',
    'core credit',
    'core must',
    'must be returned',
    'returnable core',
    'core deposit',
    'core return',
)


def _normalize_text(value):
    return re.sub(r'\s+', ' ', str(value or '').strip()).lower()


def has_core_sku_marker(value):
    sku = _normalize_text(value)
    if not sku:
        return False
    if sku == 'core':
        return True
    if sku.startswith('core ') or sku.startswith('core-') or sku.startswith('core/'):
        return True
    if sku.endswith(' core') or sku.endswith('-core') or sku.endswith('/core'):
        return True
    return bool(_CORE_TOKEN_RE.search(sku))


def has_core_description_marker(value):
    desc = _normalize_text(value)
    if not desc:
        return False

    match = _CORE_TOKEN_RE.search(desc)
    if not match:
        return False

    if desc.startswith('core'):
        return True

    first_core_idx = match.start()
    prefix = desc[:first_core_idx]
    if any(ch in prefix for ch in '.:;'):
        return False
    if first_core_idx > 40:
        return False

    return any(marker in desc for marker in _CORE_DESCRIPTION_MARKERS)


def is_core_candidate(product_service='', sku='', description=''):
    if str(product_service or '').strip().lower() == 'core':
        return True
    if has_core_sku_marker(sku):
        return True
    return has_core_description_marker(description)
