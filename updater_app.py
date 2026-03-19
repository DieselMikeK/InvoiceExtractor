"""Standalone updater helper for swapping in a new InvoiceExtractor.exe."""

import argparse
import os
import subprocess
import sys
import time
import tkinter as tk
from tkinter import ttk

from update_utils import APP_NAME, compute_file_sha256, get_resource_path

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
        self.target_dir = os.path.dirname(self.target_exe)
        self.download_path = self.target_exe + ".download"
        self.backup_path = self.target_exe + ".bak"

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
            self._download_new_exe()
            self._wait_for_app_exit()
            self._install_download()
            self._finish_success()
        except Exception as exc:
            self._cleanup_partial_files()
            self.set_status(f"Update failed: {exc}")
            self.allow_close()

    def _download_new_exe(self):
        self.set_status("Downloading update...")
        req = urllib.request.Request(
            self.args.download_url,
            headers={"User-Agent": "InvoiceExtractorUpdater/1.0"},
        )
        with urllib.request.urlopen(req, timeout=60) as response:
            total_bytes = int(response.headers.get("Content-Length") or 0)
            downloaded_bytes = 0
            with open(self.download_path, "wb") as f:
                while True:
                    chunk = response.read(DOWNLOAD_CHUNK_SIZE)
                    if not chunk:
                        break
                    f.write(chunk)
                    downloaded_bytes += len(chunk)
                    if total_bytes > 0:
                        percent = min(85, int(downloaded_bytes * 85 / total_bytes))
                        self.set_progress(percent)
                        self.set_status(
                            "Downloading update... "
                            f"{downloaded_bytes / 1_048_576:.1f} / {total_bytes / 1_048_576:.1f} MB"
                        )
                    else:
                        self.set_status(
                            f"Downloading update... {downloaded_bytes / 1_048_576:.1f} MB"
                        )

        expected_hash = (self.args.sha256 or "").strip().lower()
        if expected_hash:
            self.set_status("Verifying download...")
            actual_hash = compute_file_sha256(self.download_path)
            if actual_hash.lower() != expected_hash:
                raise RuntimeError("downloaded file hash does not match release manifest")
        self.set_progress(88)

    def _wait_for_app_exit(self):
        self.set_status("Waiting for Invoice Extractor to close...")
        wait_for_process_exit(self.args.wait_pid, PROCESS_WAIT_TIMEOUT_SECONDS)

    def _install_download(self):
        if not os.path.exists(self.download_path):
            raise RuntimeError("downloaded update was not found")

        if os.path.exists(self.backup_path):
            os.remove(self.backup_path)

        deadline = time.time() + FILE_UNLOCK_TIMEOUT_SECONDS
        while True:
            try:
                self.set_status("Installing update...")
                self.set_progress(94)
                if os.path.exists(self.target_exe):
                    os.replace(self.target_exe, self.backup_path)
                os.replace(self.download_path, self.target_exe)
                break
            except PermissionError:
                if time.time() >= deadline:
                    raise RuntimeError("existing app is still locked after waiting")
                time.sleep(0.5)
            except Exception:
                if os.path.exists(self.backup_path) and not os.path.exists(self.target_exe):
                    os.replace(self.backup_path, self.target_exe)
                raise

        if os.path.exists(self.backup_path):
            try:
                os.remove(self.backup_path)
            except OSError:
                pass
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
        for path in (self.download_path,):
            if os.path.exists(path):
                try:
                    os.remove(path)
                except OSError:
                    pass


def main(argv=None):
    args = parse_args(argv)
    app = UpdaterWindow(args)
    app.run()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
