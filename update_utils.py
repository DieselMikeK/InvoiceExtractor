"""Helpers for app versioning and release-manifest based updates."""

import hashlib
import json
import os
import shutil
import sys
import tempfile
import urllib.request


APP_NAME = "Invoice Extractor"
MAIN_EXECUTABLE_NAME = "InvoiceExtractor.exe"
UPDATER_EXECUTABLE_NAME = "InvoiceExtractorUpdater.exe"
VERSION_FILENAME = "VERSION"
DEFAULT_UPDATE_MANIFEST_URL = (
    "https://raw.githubusercontent.com/DieselMikeK/InvoiceExtractor/main/docs/release.json"
)


def get_source_dir():
    """Return the folder containing the application source files."""
    return os.path.dirname(os.path.abspath(__file__))


def get_base_dir():
    """Return the app's runtime directory."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return get_source_dir()


def get_resource_path(relative_path):
    """Resolve a bundled resource path for source and PyInstaller builds."""
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(get_source_dir(), relative_path)


def normalize_version(value):
    """Normalize a version string for comparisons and display."""
    return str(value or "").strip().lstrip("vV")


def parse_version_tuple(value):
    """Convert dotted version strings into comparable tuples."""
    version = normalize_version(value)
    if not version:
        return (0,)
    parts = []
    for token in version.split("."):
        token = token.strip()
        if token.isdigit():
            parts.append(int(token))
            continue
        digits = "".join(ch for ch in token if ch.isdigit())
        parts.append(int(digits) if digits else 0)
    return tuple(parts or [0])


def load_app_version():
    """Read the current application version from the bundled VERSION file."""
    candidates = [
        get_resource_path(VERSION_FILENAME),
        os.path.join(get_base_dir(), VERSION_FILENAME),
        os.path.join(get_source_dir(), VERSION_FILENAME),
    ]
    seen = set()
    for path in candidates:
        normalized = os.path.abspath(path)
        if normalized in seen or not os.path.exists(normalized):
            continue
        seen.add(normalized)
        try:
            with open(normalized, "r", encoding="utf-8") as f:
                version = normalize_version(f.read())
            if version:
                return version
        except OSError:
            continue
    return "0.0.0"


def load_update_config(required_dir):
    """Load optional update configuration overrides from App/required."""
    if not required_dir:
        return {}
    path = os.path.join(required_dir, "update_config.json")
    if not os.path.exists(path):
        return {}
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, dict):
            return data
    except (OSError, ValueError, TypeError):
        return {}
    return {}


def get_update_manifest_url(required_dir=None):
    """Return the remote release manifest URL for update checks."""
    env_url = str(os.environ.get("INVOICE_EXTRACTOR_UPDATE_MANIFEST_URL") or "").strip()
    if env_url:
        return env_url
    config = load_update_config(required_dir)
    config_url = str(config.get("manifest_url") or "").strip()
    if config_url:
        return config_url
    return DEFAULT_UPDATE_MANIFEST_URL


def normalize_release_manifest(data, source_url=""):
    """Normalize a release manifest payload into the fields the app expects."""
    if not isinstance(data, dict):
        raise ValueError("Release manifest must be a JSON object.")

    version = normalize_version(data.get("version") or data.get("tag_name"))
    if not version:
        raise ValueError("Release manifest is missing a version.")

    download_url = str(data.get("download_url") or "").strip()
    sha256 = str(data.get("sha256") or "").strip().lower()
    notes = str(data.get("notes") or data.get("body") or "").strip()
    published_at = str(data.get("published_at") or "").strip()

    if sha256:
        sha256 = "".join(ch for ch in sha256 if ch in "0123456789abcdef")
        if len(sha256) != 64:
            raise ValueError("Release manifest sha256 must be a 64-character hex string.")

    return {
        "version": version,
        "download_url": download_url,
        "sha256": sha256,
        "notes": notes,
        "published_at": published_at,
        "source_url": source_url,
    }


def fetch_release_manifest(required_dir=None, timeout=5):
    """Fetch and parse the remote release manifest."""
    url = get_update_manifest_url(required_dir)
    req = urllib.request.Request(
        url,
        headers={"User-Agent": f"{APP_NAME.replace(' ', '')}/UpdateCheck"},
    )
    with urllib.request.urlopen(req, timeout=timeout) as response:
        charset = response.headers.get_content_charset() or "utf-8"
        payload = response.read().decode(charset)
    data = json.loads(payload)
    return normalize_release_manifest(data, source_url=url)


def compute_file_sha256(path):
    """Return the SHA-256 hash of a file."""
    digest = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def find_updater_source_path():
    """Locate the updater executable bundled with or beside the app."""
    candidates = [
        get_resource_path(os.path.join("update", UPDATER_EXECUTABLE_NAME)),
        os.path.join(get_base_dir(), UPDATER_EXECUTABLE_NAME),
        os.path.join(get_base_dir(), "update", UPDATER_EXECUTABLE_NAME),
        os.path.join(get_source_dir(), "update", UPDATER_EXECUTABLE_NAME),
        os.path.join(get_source_dir(), "dist", UPDATER_EXECUTABLE_NAME),
    ]
    seen = set()
    for path in candidates:
        normalized = os.path.abspath(path)
        if normalized in seen:
            continue
        seen.add(normalized)
        if os.path.exists(normalized):
            return normalized
    raise FileNotFoundError(
        f"{UPDATER_EXECUTABLE_NAME} not found. Build it before shipping updates."
    )


def stage_updater_executable(current_version):
    """Copy the updater helper to a stable temp location and return that path."""
    source_path = find_updater_source_path()
    staged_dir = os.path.join(tempfile.gettempdir(), "InvoiceExtractorUpdater")
    os.makedirs(staged_dir, exist_ok=True)

    version_tag = normalize_version(current_version) or "dev"
    staged_path = os.path.join(
        staged_dir,
        f"{version_tag}-{UPDATER_EXECUTABLE_NAME}",
    )

    needs_copy = True
    if os.path.exists(staged_path):
        try:
            needs_copy = (
                os.path.getsize(staged_path) != os.path.getsize(source_path)
                or int(os.path.getmtime(staged_path)) != int(os.path.getmtime(source_path))
            )
        except OSError:
            needs_copy = True

    if needs_copy:
        shutil.copy2(source_path, staged_path)

    return staged_path
