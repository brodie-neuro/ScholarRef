from __future__ import annotations

import ctypes
import ctypes.wintypes
import subprocess
import sys
import time
from pathlib import Path

from PIL import Image


ROOT = Path(__file__).resolve().parents[1]
ASSETS = ROOT / "docs" / "assets"
ASSETS.mkdir(parents=True, exist_ok=True)

USER32 = ctypes.windll.user32
GDI32 = ctypes.windll.gdi32

PW_RENDERFULLCONTENT = 0x00000002
SRCCOPY = 0x00CC0020


def _find_window(exact_title: str, timeout: float = 20.0) -> int:
    deadline = time.time() + timeout
    hwnd = 0
    while time.time() < deadline:
        hwnd = USER32.FindWindowW(None, exact_title)
        if hwnd and USER32.IsWindowVisible(hwnd):
            return hwnd
        time.sleep(0.5)
    raise RuntimeError(f"Window '{exact_title}' was not found.")


def _capture_window(hwnd: int, output_path: Path) -> None:
    rect = ctypes.wintypes.RECT()
    if not USER32.GetWindowRect(hwnd, ctypes.byref(rect)):
        raise RuntimeError("Could not get window rectangle.")

    width = rect.right - rect.left
    height = rect.bottom - rect.top
    if width <= 0 or height <= 0:
        raise RuntimeError("Window rectangle is empty.")

    hwnd_dc = USER32.GetWindowDC(hwnd)
    if not hwnd_dc:
        raise RuntimeError("Could not get window device context.")

    mem_dc = GDI32.CreateCompatibleDC(hwnd_dc)
    bitmap = GDI32.CreateCompatibleBitmap(hwnd_dc, width, height)
    old_obj = GDI32.SelectObject(mem_dc, bitmap)

    try:
        rendered = USER32.PrintWindow(hwnd, mem_dc, PW_RENDERFULLCONTENT)
        if not rendered:
            GDI32.BitBlt(mem_dc, 0, 0, width, height, hwnd_dc, 0, 0, SRCCOPY)

        class BITMAPINFOHEADER(ctypes.Structure):
            _fields_ = [
                ("biSize", ctypes.c_uint32),
                ("biWidth", ctypes.c_long),
                ("biHeight", ctypes.c_long),
                ("biPlanes", ctypes.c_ushort),
                ("biBitCount", ctypes.c_ushort),
                ("biCompression", ctypes.c_uint32),
                ("biSizeImage", ctypes.c_uint32),
                ("biXPelsPerMeter", ctypes.c_long),
                ("biYPelsPerMeter", ctypes.c_long),
                ("biClrUsed", ctypes.c_uint32),
                ("biClrImportant", ctypes.c_uint32),
            ]

        class BITMAPINFO(ctypes.Structure):
            _fields_ = [
                ("bmiHeader", BITMAPINFOHEADER),
                ("bmiColors", ctypes.c_uint32 * 3),
            ]

        bmi = BITMAPINFO()
        bmi.bmiHeader.biSize = ctypes.sizeof(BITMAPINFOHEADER)
        bmi.bmiHeader.biWidth = width
        bmi.bmiHeader.biHeight = -height
        bmi.bmiHeader.biPlanes = 1
        bmi.bmiHeader.biBitCount = 32
        bmi.bmiHeader.biCompression = 0

        buffer_len = width * height * 4
        buffer = ctypes.create_string_buffer(buffer_len)
        bits = GDI32.GetDIBits(
            mem_dc,
            bitmap,
            0,
            height,
            buffer,
            ctypes.byref(bmi),
            0,
        )
        if bits != height:
            raise RuntimeError("Could not read bitmap pixels from window.")

        image = Image.frombuffer("RGBA", (width, height), buffer, "raw", "BGRA", 0, 1)
        image.save(output_path)
    finally:
        GDI32.SelectObject(mem_dc, old_obj)
        GDI32.DeleteObject(bitmap)
        GDI32.DeleteDC(mem_dc)
        USER32.ReleaseDC(hwnd, hwnd_dc)


def _terminate(process: subprocess.Popen[str]) -> None:
    process.terminate()
    try:
        process.wait(timeout=10)
    except subprocess.TimeoutExpired:
        process.kill()
        process.wait(timeout=5)


def main() -> int:
    gui_process = subprocess.Popen([sys.executable, "scholarref_gui.py"], cwd=ROOT)
    try:
        hwnd = _find_window("ScholarRef")
        time.sleep(2.5)
        _capture_window(hwnd, ASSETS / "gui-window.png")
    finally:
        _terminate(gui_process)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
