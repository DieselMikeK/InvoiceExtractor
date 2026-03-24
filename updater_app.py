"""Standalone updater helper for swapping in a new InvoiceExtractor.exe."""

import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
import tkinter as tk
from tkinter import ttk

from update_utils import (
    APP_NAME,
    MAIN_EXECUTABLE_NAME,
    compute_file_sha256,
    get_resource_path,
    normalize_release_manifest,
    open_url_with_tls_fallback,
)

try:
    import ctypes
except Exception:  # pragma: no cover - Windows-only behavior is best effort
    ctypes = None

import urllib.request


DOWNLOAD_CHUNK_SIZE = 1024 * 256
PROCESS_WAIT_TIMEOUT_SECONDS = 45
FILE_UNLOCK_TIMEOUT_SECONDS = 45


def parse_args(argv=None):
    parser = argparse.ArgumentParser(description="Invoice Extractor updater")
    parser.add_argument("--current-exe", required=True, help="Path to the installed InvoiceExtractor.exe")
    parser.add_argument("--manifest-file", default="", help="Path to a release manifest file")
    parser.add_argument("--download-url", required=True, help="URL of the replacement executable")
    parser.add_argument("--target-version", required=True, help="Version being installed")
    parser.add_argument("--source-version", default="", help="Version currently installed")
    parser.add_argument("--sha256", default="", help="Expected SHA-256 hash for the downloaded executable")
    parser.add_argument("--wait-pid", type=int, default=0, help="PID of the app process that should exit first")
    return parser.parse_args(argv)


def wait_for_process_exit(pid, timeout_seconds):
    """Wait for a process to exit on Windows. Returns True if it exited in time."""
    if not pid or os.name != "nt" or ctypes is None:
        return True

    SYNCHRONIZE = 0x00100000
    WAIT_OBJECT_0 = 0x00000000
    WAIT_TIMEOUT = 0x00000102

    kernel32 = ctypes.windll.kernel32
    process_handle = kernel32.OpenProcess(SYNCHRONIZE, False, pid)
    if not process_handle:
        return True

    try:
        result = kernel32.WaitForSingleObject(process_handle, int(timeout_seconds * 1000))
        return result == WAIT_OBJECT_0 or result != WAIT_TIMEOUT
    finally:
        kernel32.CloseHandle(process_handle)


class UpdaterWindow:
    def __init__(self, args):
        self.args = args
        self.target_exe = os.path.abspath(args.current_exe)
        self.install_root = os.path.dirname(self.target_exe)
        self.target_dir = self.install_root
        self.staging_dir = tempfile.mkdtemp(prefix="InvoiceExtractorUpdate-")
        self.release_files = self._load_release_files()

        self.root = tk.Tk()
        self.root.title(f"{APP_NAME} Updater")
        self.root.resizable(False, False)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.configure(bg="#171717")

        try:
            icon_path = get_resource_path("logo.ico")
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
        except Exception:
            pass

        self.status_var = tk.StringVar(value="Preparing update...")
        version_text = f"v{args.source_version or '?'} -> v{args.target_version}"
        self.detail_var = tk.StringVar(value=version_text)
        self.can_close = False

        self._build_ui()

    def _build_ui(self):
        width, height = 480, 210
        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()
        self.root.geometry(f"{width}x{height}+{(screen_w-width)//2}+{(screen_h-height)//2}")

        container = tk.Frame(self.root, bg="#171717", padx=24, pady=20)
        container.pack(fill=tk.BOTH, expand=True)

        tk.Label(
            container,
            text=f"{APP_NAME} Update",
            font=("Segoe UI", 14, "bold"),
            bg="#171717",
            fg="#f3f3f3",
        ).pack(anchor="w")

        tk.Label(
            container,
            textvariable=self.detail_var,
            font=("Segoe UI", 9),
            bg="#171717",
            fg="#a8a8a8",
        ).pack(anchor="w", pady=(4, 16))

        style = ttk.Style()
        style.theme_use("default")
        style.configure(
            "Updater.Horizontal.TProgressbar",
            troughcolor="#303030",
            background="#2ea043",
            bordercolor="#171717",
            lightcolor="#2ea043",
            darkcolor="#2ea043",
        )

        self.progress = ttk.Progressbar(
            container,
            orient="horizontal",
            mode="determinate",
            maximum=100,
            length=428,
            style="Updater.Horizontal.TProgressbar",
        )
        self.progress.pack(anchor="w")

        tk.Label(
            container,
            textvariable=self.status_var,
            font=("Segoe UI", 10),
            bg="#171717",
            fg="#d6d6d6",
            wraplength=430,
            justify="left",
        ).pack(anchor="w", pady=(16, 0))

        self.close_button = ttk.Button(container, text="Close", command=self.root.destroy)
        self.close_button.pack(anchor="e", pady=(18, 0))
        self.close_button.configure(state=tk.DISABLED)

    def _on_close(self):
        if self.can_close:
            self.root.destroy()

    def set_status(self, text):
        self.root.after(0, lambda: self.status_var.set(text))

    def set_progress(self, value):
        clamped = max(0, min(100, int(value)))
        self.root.after(0, lambda: self.progress.configure(value=clamped))

    def allow_close(self):
        self.can_close = True
        self.root.after(0, lambda: self.close_button.configure(state=tk.NORMAL))

    def run(self):
        self.root.after(100, self._start_update_thread)
        self.root.mainloop()

    def _start_update_thread(self):
        import threading

        threading.Thread(target=self._perform_update, daemon=True).start()

    def _perform_update(self):
        try:
            self._download_release_files()
            self._wait_for_app_exit()
            self._install_release_files()
            self._finish_success()
        except Exception as exc:
            self._cleanup_partial_files()
            self.set_status(f"Update failed: {exc}")
            self.allow_close()

    def _resolve_target_path(self, relative_path):
        normalized = str(relative_path or "").replace("\\", "/").strip()
        parts = [part for part in normalized.split("/") if part and part != "."]
        if not parts or any(part == ".." for part in parts):
            raise RuntimeError(f"Invalid release target path '{relative_path}'.")
        target_path = os.path.abspath(os.path.join(self.install_root, *parts))
        if os.path.commonpath([self.install_root, target_path]) != self.install_root:
            raise RuntimeError(f"Release target path '{relative_path}' is outside the install folder.")
        return target_path

    def _build_release_entry(self, relative_path, download_url, sha256):
        target_path = self._resolve_target_path(relative_path)
        return {
            "relative_path": str(relative_path).replace("\\", "/"),
            "download_url": str(download_url or "").strip(),
            "sha256": str(sha256 or "").strip().lower(),
            "target_path": target_path,
            "staged_path": os.path.join(
                self.staging_dir,
                *str(relative_path).replace("\\", "/").split("/"),
            ),
            "backup_path": target_path + ".bak",
            "is_main_exe": os.path.abspath(target_path) == self.target_exe,
        }

    def _load_release_files(self):
        manifest_path = str(self.args.manifest_file or "").strip()
        release_files = []

        if manifest_path:
            with open(manifest_path, "r", encoding="utf-8") as f:
                manifest = normalize_release_manifest(json.load(f), source_url=manifest_path)
            release_files.extend(
                self._build_release_entry(
                    entry["relative_path"],
                    entry.get("download_url"),
                    entry.get("sha256"),
                )
                for entry in manifest.get("files") or []
            )

        if not any(entry["is_main_exe"] for entry in release_files):
            release_files.append(
                self._build_release_entry(
                    MAIN_EXECUTABLE_NAME,
                    self.args.download_url,
                    self.args.sha256,
                )
            )

        return release_files

    def _download_one_file(self, entry, progress_start, progress_span):
        display_name = entry["relative_path"].replace("/", "\\")
        self.set_status(f"Downloading {display_name}...")
        req = urllib.request.Request(
            entry["download_url"],
            headers={"User-Agent": "InvoiceExtractorUpdater/1.0"},
        )
        os.makedirs(os.path.dirname(entry["staged_path"]), exist_ok=True)
        with open_url_with_tls_fallback(req, timeout=60) as response:
            total_bytes = int(response.headers.get("Content-Length") or 0)
            downloaded_bytes = 0
            with open(entry["staged_path"], "wb") as f:
                while True:
                    chunk = response.read(DOWNLOAD_CHUNK_SIZE)
                    if not chunk:
                        break
                    f.write(chunk)
                    downloaded_bytes += len(chunk)
                    if total_bytes > 0:
                        percent = min(
                            progress_start + progress_span,
                            progress_start + int(downloaded_bytes * progress_span / total_bytes),
                        )
                        self.set_progress(percent)
                        self.set_status(
                            f"Downloading {display_name}... "
                            f"{downloaded_bytes / 1_048_576:.1f} / {total_bytes / 1_048_576:.1f} MB"
                        )
                    else:
                        self.set_status(
                            f"Downloading {display_name}... {downloaded_bytes / 1_048_576:.1f} MB"
                        )

        expected_hash = str(entry.get("sha256") or "").strip().lower()
        if expected_hash:
            self.set_status(f"Verifying {display_name}...")
            actual_hash = compute_file_sha256(entry["staged_path"])
            if actual_hash.lower() != expected_hash:
                raise RuntimeError(f"{display_name} hash does not match the release manifest")

    def _download_release_files(self):
        file_count = max(1, len(self.release_files))
        progress_span = max(1, 85 // file_count)
        for index, entry in enumerate(self.release_files):
            start = index * progress_span
            self._download_one_file(entry, start, progress_span)
        self.set_progress(88)

    def _wait_for_app_exit(self):
        self.set_status("Waiting for Invoice Extractor to close...")
        wait_for_process_exit(self.args.wait_pid, PROCESS_WAIT_TIMEOUT_SECONDS)

    def _install_one_file(self, entry):
        if not os.path.exists(entry["staged_path"]):
            raise RuntimeError(f"Downloaded file was not found for {entry['relative_path']}")

        os.makedirs(os.path.dirname(entry["target_path"]), exist_ok=True)
        if os.path.exists(entry["backup_path"]):
            os.remove(entry["backup_path"])

        had_existing_file = os.path.exists(entry["target_path"])
        deadline = time.time() + FILE_UNLOCK_TIMEOUT_SECONDS
        while True:
            try:
                if had_existing_file:
                    os.replace(entry["target_path"], entry["backup_path"])
                os.replace(entry["staged_path"], entry["target_path"])
                return had_existing_file
            except PermissionError:
                if not entry["is_main_exe"] or time.time() >= deadline:
                    raise RuntimeError(
                        f"{entry['relative_path'].replace('/', chr(92))} is still locked after waiting"
                    )
                time.sleep(0.5)
            except Exception:
                if had_existing_file and os.path.exists(entry["backup_path"]) and not os.path.exists(entry["target_path"]):
                    os.replace(entry["backup_path"], entry["target_path"])
                raise

    def _restore_file(self, entry, had_existing_file):
        if os.path.exists(entry["target_path"]):
            try:
                os.remove(entry["target_path"])
            except OSError:
                pass
        if had_existing_file and os.path.exists(entry["backup_path"]):
            try:
                os.replace(entry["backup_path"], entry["target_path"])
            except OSError:
                pass

    def _install_release_files(self):
        ordered_files = [
            entry for entry in self.release_files if not entry["is_main_exe"]
        ] + [
            entry for entry in self.release_files if entry["is_main_exe"]
        ]
        applied_files = []

        try:
            for index, entry in enumerate(ordered_files, start=1):
                display_name = entry["relative_path"].replace("/", "\\")
                self.set_status(f"Installing {display_name}...")
                self.set_progress(88 + min(11, index * 10 // max(1, len(ordered_files))))
                had_existing_file = self._install_one_file(entry)
                applied_files.append((entry, had_existing_file))
        except Exception:
            for entry, had_existing_file in reversed(applied_files):
                self._restore_file(entry, had_existing_file)
            raise
        finally:
            for entry, _had_existing_file in applied_files:
                if os.path.exists(entry["backup_path"]):
                    try:
                        os.remove(entry["backup_path"])
                    except OSError:
                        pass

        self._cleanup_partial_files()
        self.set_progress(100)

    def _finish_success(self):
        self.set_status("Update installed. Reopening app...")
        time.sleep(1.0)
        try:
            subprocess.Popen([self.target_exe], cwd=self.target_dir)
        except Exception:
            self.set_status("Update installed. Please reopen Invoice Extractor manually.")
            self.allow_close()
            return
        self.can_close = True
        self.root.after(0, self.root.destroy)

    def _cleanup_partial_files(self):
        if self.staging_dir and os.path.exists(self.staging_dir):
            shutil.rmtree(self.staging_dir, ignore_errors=True)


def main(argv=None):
    args = parse_args(argv)
    app = UpdaterWindow(args)
    app.run()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
