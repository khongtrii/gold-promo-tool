"""Standalone updater UI for Gold Promo Tool."""

from __future__ import annotations

import argparse
import ctypes
import os
import subprocess
import sys
import tempfile
import threading
import time
from pathlib import Path
from tkinter import Tk, messagebox, ttk

from src.update.check_update import UpdateInfo, check_update
from src.update.download import download_update


class UpdaterApp:
    def __init__(
        self,
        root: Tk,
        app_path: Path,
        app_pid: int,
        current_version: str,
        check_only: bool = False,
    ) -> None:
        self.root = root
        self.app_path = app_path.resolve()
        self.app_pid = app_pid
        self.current_version = current_version
        self.check_only = check_only
        self.info: UpdateInfo | None = None
        root.title("Gold Promo Tool Updater")
        root.resizable(False, False)
        root.protocol("WM_DELETE_WINDOW", lambda: None)
        self.label = ttk.Label(root, text="Đang kiểm tra bản cập nhật…", width=54)
        self.label.pack(padx=24, pady=(20, 10))
        self.progress = ttk.Progressbar(root, length=410, mode="indeterminate")
        self.progress.pack(padx=24, pady=(0, 20))
        self.progress.start(12)
        threading.Thread(target=self._check, daemon=True).start()

    def _check(self) -> None:
        try:
            info = check_update(self.current_version)
        except Exception as error:
            self.root.after(0, self._check_failed, str(error))
            return
        self.root.after(0, self._show_result, info)

    def _check_failed(self, detail: str) -> None:
        self.progress.stop()
        messagebox.showerror("Không thể kiểm tra cập nhật", f"Không thể kết nối máy chủ cập nhật.\n\n{detail}", parent=self.root)
        self.root.destroy()

    def _show_result(self, info: UpdateInfo) -> None:
        self.progress.stop()
        self.info = info
        if not info.available:
            messagebox.showinfo("Cập nhật", f"Bạn đang dùng phiên bản mới nhất ({info.current_version}).", parent=self.root)
            self.root.destroy()
            return
        if self.check_only:
            messagebox.showinfo(
                "Có bản cập nhật mới",
                f"Phiên bản {info.latest_version} đã sẵn sàng (hiện tại: {info.current_version}).\n\n"
                "Chế độ chạy source chỉ kiểm tra, không tự thay thế file.",
                parent=self.root,
            )
            self.root.destroy()
            return
        if not messagebox.askyesno(
            "Có bản cập nhật mới",
            f"Phiên bản {info.latest_version} đã sẵn sàng (hiện tại: {info.current_version}).\n\nTải và cài đặt ngay?",
            parent=self.root,
        ):
            self.root.destroy()
            return
        self.label.config(text=f"Đang tải phiên bản {info.latest_version}…")
        self.progress.config(mode="determinate", maximum=100, value=0)
        threading.Thread(target=self._download, daemon=True).start()

    def _download(self) -> None:
        assert self.info is not None and self.info.download_url
        destination = Path(tempfile.gettempdir()) / f"gold-promo-tool-{self.info.latest_version}.download"
        try:
            download_update(self.info.download_url, destination, self._on_progress)
            self.root.after(0, self._install, destination)
        except Exception as error:
            destination.unlink(missing_ok=True)
            self.root.after(0, self._download_failed, str(error))

    def _on_progress(self, downloaded: int, total: int) -> None:
        percent = downloaded * 100 / total if total else 0
        text = f"Đang tải… {downloaded / 1024 / 1024:.1f} MB"
        if total:
            text += f" / {total / 1024 / 1024:.1f} MB"
        self.root.after(0, self._set_progress, percent, text, bool(total))

    def _set_progress(self, percent: float, text: str, determinate: bool) -> None:
        self.label.config(text=text)
        if determinate:
            self.progress["value"] = percent

    def _download_failed(self, detail: str) -> None:
        messagebox.showerror("Tải cập nhật thất bại", detail, parent=self.root)
        self.root.destroy()

    def _install(self, downloaded: Path) -> None:
        self.label.config(text="Đang cài đặt bản cập nhật…")
        self.progress["value"] = 100
        self.root.update_idletasks()
        try:
            self._stop_application()
            backup = self.app_path.with_suffix(self.app_path.suffix + ".old")
            backup.unlink(missing_ok=True)
            self._replace_with_retry(self.app_path, backup)
            try:
                self._replace_with_retry(downloaded, self.app_path)
            except Exception:
                self._replace_with_retry(backup, self.app_path)
                raise
            subprocess.Popen([str(self.app_path)], cwd=str(self.app_path.parent))
        except Exception as error:
            messagebox.showerror("Cài đặt thất bại", str(error), parent=self.root)
        finally:
            self.root.destroy()

    def _stop_application(self) -> None:
        if self.app_pid <= 0 or self.app_pid == os.getpid():
            return
        # Open the process handle before killing it so we can reliably wait
        # until Windows releases the executable file. Do not use /T here:
        # updater is a child of the desktop app and would kill itself.
        handle = ctypes.windll.kernel32.OpenProcess(0x00100000, False, self.app_pid)
        subprocess.run(
            ["taskkill", "/PID", str(self.app_pid), "/F"],
            check=False,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        if handle:
            ctypes.windll.kernel32.WaitForSingleObject(handle, 15_000)
            ctypes.windll.kernel32.CloseHandle(handle)

    @staticmethod
    def _replace_with_retry(source: Path, destination: Path) -> None:
        """Retry briefly while Windows releases executable file handles."""
        last_error: OSError | None = None
        for _ in range(30):
            try:
                os.replace(source, destination)
                return
            except PermissionError as error:
                last_error = error
                time.sleep(0.2)
        if last_error is not None:
            raise last_error


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--app-path", required=True, type=Path)
    parser.add_argument("--app-pid", required=True, type=int)
    parser.add_argument("--current-version", required=True)
    parser.add_argument("--check-only", action="store_true")
    args = parser.parse_args()
    root = Tk()
    UpdaterApp(root, args.app_path, args.app_pid, args.current_version, args.check_only)
    root.mainloop()


if __name__ == "__main__":
    main()
