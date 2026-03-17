"""
Invoice Extractor - Updater
Downloads the latest pre-built exe from GitHub Releases and swaps it in.
Launched by the running exe via sys.executable — no separate Python needed.
"""

import os
import sys
import time
import shutil
import urllib.request
import urllib.error
import tkinter as tk
from tkinter import ttk
import threading


def main():
    if len(sys.argv) < 5:
        print("Usage: updater.py <old_version> <new_version> <download_url> <exe_path>")
        sys.exit(1)

    old_version = sys.argv[1]
    new_version = sys.argv[2]
    download_url = sys.argv[3]
    exe_path = sys.argv[4]

    app = UpdaterApp(old_version, new_version, download_url, exe_path)
    app.run()


class UpdaterApp:
    def __init__(self, old_version, new_version, download_url, exe_path):
        self.old_version = old_version
        self.new_version = new_version
        self.download_url = download_url
        self.exe_path = exe_path

        self.root = tk.Tk()
        self.root.title("Invoice Extractor — Updating")
        self.root.resizable(False, False)
        self.root.protocol("WM_DELETE_WINDOW", lambda: None)  # block close

        # Center window
        w, h = 460, 180
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        self.root.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")
        self.root.configure(bg="#1e1e1e")

        try:
            self.root.iconbitmap(self._find_icon())
        except Exception:
            pass

        self._build_ui()

    def _find_icon(self):
        # updater.py is in app/update/, icon is in app/
        here = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(here, '..', 'logo.ico')

    def _build_ui(self):
        bg = "#1e1e1e"
        fg = "#ffffff"
        accent = "#2ea043"

        tk.Label(
            self.root,
            text="Invoice Extractor — Update",
            font=("Segoe UI", 13, "bold"),
            bg=bg, fg=fg
        ).pack(pady=(20, 4))

        tk.Label(
            self.root,
            text=f"v{self.old_version}  →  v{self.new_version}",
            font=("Segoe UI", 10),
            bg=bg, fg="#aaaaaa"
        ).pack(pady=(0, 14))

        bar_frame = tk.Frame(self.root, bg=bg)
        bar_frame.pack(fill=tk.X, padx=40)

        style = ttk.Style()
        style.theme_use("default")
        style.configure(
            "Green.Horizontal.TProgressbar",
            troughcolor="#333333",
            background=accent,
            bordercolor="#1e1e1e",
            lightcolor=accent,
            darkcolor=accent,
        )

        self.progress = ttk.Progressbar(
            bar_frame,
            style="Green.Horizontal.TProgressbar",
            orient="horizontal",
            length=380,
            mode="determinate"
        )
        self.progress.pack()

        self.status_var = tk.StringVar(value="Starting download...")
        tk.Label(
            self.root,
            textvariable=self.status_var,
            font=("Segoe UI", 9),
            bg=bg, fg="#888888"
        ).pack(pady=(10, 0))

    def run(self):
        threading.Thread(target=self._update_thread, daemon=True).start()
        self.root.mainloop()

    def _set_status(self, text):
        self.root.after(0, lambda: self.status_var.set(text))

    def _set_progress(self, pct):
        self.root.after(0, lambda: self.progress.configure(value=pct))

    def _update_thread(self):
        tmp_path = self.exe_path + ".new"
        try:
            # Download
            self._set_status("Downloading update...")
            self._download(self.download_url, tmp_path)

            # Wait for original exe to be unlocked (it already exited, but give OS a moment)
            self._set_status("Preparing to install...")
            self._set_progress(95)
            time.sleep(1.5)

            # Swap
            self._set_status("Installing...")
            backup = self.exe_path + ".bak"
            if os.path.exists(backup):
                os.remove(backup)
            if os.path.exists(self.exe_path):
                os.rename(self.exe_path, backup)
            os.rename(tmp_path, self.exe_path)
            if os.path.exists(backup):
                os.remove(backup)

            self._set_progress(100)
            self._set_status("Update complete! Relaunching...")
            time.sleep(1.0)

        except Exception as e:
            # Clean up partial download
            if os.path.exists(tmp_path):
                try:
                    os.remove(tmp_path)
                except Exception:
                    pass
            self._set_status(f"Update failed: {e}")
            time.sleep(4)
            self.root.after(0, self.root.destroy)
            return

        # Relaunch
        import subprocess
        subprocess.Popen([self.exe_path])
        self.root.after(0, self.root.destroy)

    def _download(self, url, dest):
        req = urllib.request.Request(url, headers={"User-Agent": "InvoiceExtractor-Updater"})
        with urllib.request.urlopen(req, timeout=60) as response:
            total = int(response.headers.get("Content-Length", 0))
            downloaded = 0
            chunk = 65536
            with open(dest, "wb") as f:
                while True:
                    block = response.read(chunk)
                    if not block:
                        break
                    f.write(block)
                    downloaded += len(block)
                    if total > 0:
                        pct = min(int(downloaded / total * 90), 90)
                        self._set_progress(pct)
                        mb_done = downloaded / 1_048_576
                        mb_total = total / 1_048_576
                        self._set_status(f"Downloading...  {mb_done:.1f} / {mb_total:.1f} MB")
                    else:
                        self._set_status(f"Downloading...  {downloaded / 1_048_576:.1f} MB")


if __name__ == "__main__":
    main()
