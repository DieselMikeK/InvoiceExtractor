"""Helpers for app versioning and release-manifest based updates."""

import base64
import hashlib
import json
import os
import shutil
import ssl
import sys
import tempfile
import urllib.error
import urllib.request

try:
    import certifi
except Exception:  # pragma: no cover - optional fallback
    certifi = None


APP_NAME = "Invoice Extractor"
MAIN_EXECUTABLE_NAME = "InvoiceExtractor.exe"
UPDATER_EXECUTABLE_NAME = "InvoiceExtractorUpdater.exe"
UPDATER_RELEASE_RELATIVE_PATH = f"update/{UPDATER_EXECUTABLE_NAME}"
VERSION_FILENAME = "VERSION"
PRIMARY_RELEASE_RELATIVE_PATH = MAIN_EXECUTABLE_NAME
DEFAULT_UPDATE_MANIFEST_URL = (
    "https://api.github.com/repos/DieselMikeK/InvoiceExtractor/contents/docs/release.json?ref=main"
)
DEFAULT_UPDATE_MANIFEST_FALLBACK_URLS = [
    "https://raw.githubusercontent.com/DieselMikeK/InvoiceExtractor/main/docs/release.json",
]
CERTIFI_CA_BUNDLE = certifi.where() if certifi else ""


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


def normalize_release_relative_path(value):
    """Normalize a manifest file path and reject traversal outside the install root."""
    raw_value = str(value or "").strip().replace("\\", "/")
    parts = []
    for token in raw_value.split("/"):
        token = token.strip()
        if not token or token == ".":
            continue
        if token == "..":
            raise ValueError("Release manifest file path cannot contain '..'.")
        parts.append(token)
    if not parts:
        raise ValueError("Release manifest file path is missing.")
    return "/".join(parts)


def normalize_sha256(value):
    """Normalize and validate a SHA-256 string."""
    sha256 = str(value or "").strip().lower()
    if sha256:
        sha256 = "".join(ch for ch in sha256 if ch in "0123456789abcdef")
        if len(sha256) != 64:
            raise ValueError("Release manifest sha256 must be a 64-character hex string.")
    return sha256


def normalize_release_file(entry):
    """Normalize a single file entry from the release manifest."""
    if not isinstance(entry, dict):
        raise ValueError("Release manifest file entries must be JSON objects.")

    relative_path = normalize_release_relative_path(
        entry.get("relative_path") or entry.get("path")
    )
    download_url = str(entry.get("download_url") or "").strip()
    sha256 = normalize_sha256(entry.get("sha256"))

    return {
        "relative_path": relative_path,
        "download_url": download_url,
        "sha256": sha256,
    }


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


def _is_cert_verification_error(exc):
    """Return True when an exception indicates TLS certificate verification failed."""
    if isinstance(exc, ssl.SSLCertVerificationError):
        return True
    if isinstance(exc, ssl.SSLError) and "CERTIFICATE_VERIFY_FAILED" in str(exc):
        return True
    if isinstance(exc, urllib.error.URLError):
        reason = getattr(exc, "reason", None)
        if isinstance(reason, ssl.SSLCertVerificationError):
            return True
        if isinstance(reason, ssl.SSLError) and "CERTIFICATE_VERIFY_FAILED" in str(reason):
            return True
        if "CERTIFICATE_VERIFY_FAILED" in str(exc):
            return True
    return "CERTIFICATE_VERIFY_FAILED" in str(exc)


def open_url_with_tls_fallback(request, timeout=5):
    """Open a URL, retrying with certifi's CA bundle when default TLS trust fails."""
    try:
        return urllib.request.urlopen(request, timeout=timeout)
    except Exception as exc:
        if not CERTIFI_CA_BUNDLE or not _is_cert_verification_error(exc):
            raise

    ssl_context = ssl.create_default_context(cafile=CERTIFI_CA_BUNDLE)
    return urllib.request.urlopen(request, timeout=timeout, context=ssl_context)


def normalize_release_manifest(data, source_url=""):
    """Normalize a release manifest payload into the fields the app expects."""
    if not isinstance(data, dict):
        raise ValueError("Release manifest must be a JSON object.")

    version = normalize_version(data.get("version") or data.get("tag_name"))
    if not version:
        raise ValueError("Release manifest is missing a version.")

    download_url = str(data.get("download_url") or "").strip()
    sha256 = normalize_sha256(data.get("sha256"))
    notes = str(data.get("notes") or data.get("body") or "").strip()
    published_at = str(data.get("published_at") or "").strip()
    files = []

    raw_files = data.get("files")
    if raw_files is not None:
        if not isinstance(raw_files, list):
            raise ValueError("Release manifest files must be a JSON array.")
        seen_paths = set()
        for entry in raw_files:
            normalized_entry = normalize_release_file(entry)
            key = normalized_entry["relative_path"].lower()
            if key in seen_paths:
                raise ValueError(
                    f"Release manifest contains duplicate file entry '{normalized_entry['relative_path']}'."
                )
            seen_paths.add(key)
            files.append(normalized_entry)

    if not files and download_url:
        files.append(
            {
                "relative_path": PRIMARY_RELEASE_RELATIVE_PATH,
                "download_url": download_url,
                "sha256": sha256,
            }
        )

    primary_file = next(
        (
            entry
            for entry in files
            if entry["relative_path"].lower() == PRIMARY_RELEASE_RELATIVE_PATH.lower()
        ),
        None,
    )
    if primary_file:
        if not download_url:
            download_url = primary_file["download_url"]
        if not sha256:
            sha256 = primary_file["sha256"]

    return {
        "version": version,
        "download_url": download_url,
        "sha256": sha256,
        "notes": notes,
        "published_at": published_at,
        "files": files,
        "source_url": source_url,
    }


def decode_release_manifest_payload(data):
    """Decode supported manifest transport payloads into the underlying manifest JSON."""
    if not isinstance(data, dict):
        raise ValueError("Release manifest must be a JSON object.")

    if "version" in data or "tag_name" in data:
        return data

    raw_content = data.get("content")
    if raw_content is None:
        return data

    encoding = str(data.get("encoding") or "").strip().lower()
    if encoding and encoding != "base64":
        raise ValueError(f"Unsupported release manifest content encoding '{encoding}'.")

    try:
        decoded = base64.b64decode(str(raw_content))
        nested_data = json.loads(decoded.decode("utf-8"))
    except Exception as exc:
        raise ValueError("Release manifest content payload could not be decoded.") from exc

    if not isinstance(nested_data, dict):
        raise ValueError("Decoded release manifest content must be a JSON object.")
    return nested_data


def find_release_file(manifest, relative_path):
    """Return a normalized manifest file entry by relative path, if present."""
    normalized_manifest = normalize_release_manifest(manifest)
    target_key = normalize_release_relative_path(relative_path).lower()
    for entry in normalized_manifest.get("files") or []:
        if entry["relative_path"].lower() == target_key:
            return entry
    return None


def fetch_release_manifest(required_dir=None, timeout=5):
    """Fetch and parse the remote release manifest."""
    configured_url = get_update_manifest_url(required_dir)
    urls = [configured_url]
    if configured_url == DEFAULT_UPDATE_MANIFEST_URL:
        for fallback_url in DEFAULT_UPDATE_MANIFEST_FALLBACK_URLS:
            if fallback_url not in urls:
                urls.append(fallback_url)

    last_error = None
    for url in urls:
        headers = {"User-Agent": f"{APP_NAME.replace(' ', '')}/UpdateCheck"}
        if "api.github.com/" in url:
            headers["Accept"] = "application/vnd.github+json"
        req = urllib.request.Request(url, headers=headers)
        try:
            with open_url_with_tls_fallback(req, timeout=timeout) as response:
                charset = response.headers.get_content_charset() or "utf-8"
                payload = response.read().decode(charset)
            data = decode_release_manifest_payload(json.loads(payload))
            return normalize_release_manifest(data, source_url=url)
        except Exception as exc:
            last_error = exc

    if last_error is not None:
        raise last_error
    raise ValueError("No release manifest URLs are configured.")


def compute_file_sha256(path):
    """Return the SHA-256 hash of a file."""
    digest = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def download_release_file_to_path(entry, destination_path, timeout=60):
    """Download a release file entry to a specific path and verify its hash."""
    normalized_entry = normalize_release_file(entry)
    download_url = str(normalized_entry.get("download_url") or "").strip()
    if not download_url:
        raise ValueError(
            f"Release file '{normalized_entry['relative_path']}' is missing a download URL."
        )

    expected_hash = str(normalized_entry.get("sha256") or "").strip().lower()
    destination_path = os.path.abspath(destination_path)
    os.makedirs(os.path.dirname(destination_path), exist_ok=True)

    if expected_hash and os.path.exists(destination_path):
        try:
            if compute_file_sha256(destination_path).lower() == expected_hash:
                return destination_path
        except OSError:
            pass

    temp_path = destination_path + ".download"
    req = urllib.request.Request(
        download_url,
        headers={"User-Agent": f"{APP_NAME.replace(' ', '')}/UpdaterBootstrap"},
    )

    try:
        with open_url_with_tls_fallback(req, timeout=timeout) as response:
            with open(temp_path, "wb") as f:
                shutil.copyfileobj(response, f)

        if expected_hash:
            actual_hash = compute_file_sha256(temp_path).lower()
            if actual_hash != expected_hash:
                raise ValueError(
                    f"Release file '{normalized_entry['relative_path']}' hash does not match."
                )

        os.replace(temp_path, destination_path)
        return destination_path
    finally:
        if os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except OSError:
                pass


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


def stage_updater_executable(current_version, manifest=None):
    """Stage the updater helper from the release manifest when available, else local bundle."""
    staged_dir = os.path.join(tempfile.gettempdir(), "InvoiceExtractorUpdater")
    os.makedirs(staged_dir, exist_ok=True)

    version_tag = normalize_version(current_version) or "dev"
    staged_path = os.path.join(
        staged_dir,
        f"{version_tag}-{UPDATER_EXECUTABLE_NAME}",
    )

    manifest_entry = None
    if manifest:
        try:
            manifest_entry = find_release_file(manifest, UPDATER_RELEASE_RELATIVE_PATH)
        except Exception:
            manifest_entry = None

    if manifest_entry and str(manifest_entry.get("download_url") or "").strip():
        return download_release_file_to_path(manifest_entry, staged_path)

    source_path = find_updater_source_path()
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


def stage_release_manifest(manifest, current_version=""):
    """Write a normalized release manifest to a stable temp path for the updater helper."""
    normalized_manifest = normalize_release_manifest(manifest)
    staged_dir = os.path.join(tempfile.gettempdir(), "InvoiceExtractorUpdater")
    os.makedirs(staged_dir, exist_ok=True)

    version_tag = normalize_version(current_version or normalized_manifest.get("version")) or "dev"
    staged_path = os.path.join(staged_dir, f"{version_tag}-release-manifest.json")

    with open(staged_path, "w", encoding="utf-8") as f:
        json.dump(normalized_manifest, f, indent=2)

    return staged_path
