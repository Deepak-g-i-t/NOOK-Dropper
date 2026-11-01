"""
Universal Notch Drop Target & OLE Bridge for Nook Drop Changer.
Provides seamless Drag & Drop support from Web Browsers (Chrome, Microsoft Edge,
Firefox), desktop applications, document editors, and Windows File Explorer directly
into Nook Notch.

Key Capabilities:
1. Browser text selection (plain text, formatted text, HTML).
2. Browser images (JPEG, PNG, WebP, GIF, SVG, data URIs, DIB bitmaps).
3. Browser hyperlinks and address bar URLs.
4. Windows File Explorer local files (single or multiple).
5. Zero-Clash Architecture: Window is 100% invisible and click-through (HTTRANSPARENT)
   when not dragging; only activates as an OLE IDropTarget during active mouse drag.
"""

import os
import re
import sys
import time
import base64
import struct
import tempfile
import threading
import urllib.request
import urllib.parse
from io import BytesIO
from pathlib import Path
from PIL import Image

import ctypes
from ctypes import wintypes
import win32gui
import win32con
import win32api
import win32clipboard
import pythoncom

from diagnostic_logger import DiagnosticLogger

user32 = ctypes.windll.user32

# Registered Clipboard & OLE Formats
CF_HTML = win32clipboard.RegisterClipboardFormat("HTML Format")
CF_URL = win32clipboard.RegisterClipboardFormat("UniformResourceLocatorW")
CF_URL_A = win32clipboard.RegisterClipboardFormat("UniformResourceLocator")
CF_FILEDESCW = win32clipboard.RegisterClipboardFormat("FileGroupDescriptorW")
CF_FILECONTENTS = win32clipboard.RegisterClipboardFormat("FileContents")


# ---------------------------------------------------------------------------
# Cache Directory for Web Assets
# ---------------------------------------------------------------------------

def get_cache_dir():
    cache = Path(tempfile.gettempdir()) / "nook_drop_cache"
    cache.mkdir(parents=True, exist_ok=True)
    return str(cache)


# ---------------------------------------------------------------------------
# Mouse Drag State Detection
# ---------------------------------------------------------------------------

def is_mouse_dragging():
    """Returns True if the left or right mouse button is currently held down."""
    l_down = bool(user32.GetAsyncKeyState(0x01) & 0x8000)
    r_down = bool(user32.GetAsyncKeyState(0x02) & 0x8000)
    return l_down or r_down


# ---------------------------------------------------------------------------
# DIB Bitmap Conversion
# ---------------------------------------------------------------------------

def dib_bytes_to_image(dib_data):
    """Converts CF_DIB / CF_DIBV5 raw memory bytes into a PIL Image."""
    header_size = struct.unpack('<I', dib_data[:4])[0]
    bpp = struct.unpack('<H', dib_data[14:16])[0]
    compression = struct.unpack('<I', dib_data[16:20])[0] if len(dib_data) >= 20 else 0

    palette_len = 0
    if bpp <= 8:
        clr_used = struct.unpack('<I', dib_data[32:36])[0] if len(dib_data) >= 36 else 0
        palette_len = (clr_used or (1 << bpp)) * 4
    elif compression == 3:  # BI_BITFIELDS
        palette_len = 12

    offset = 14 + header_size + palette_len
    file_header = struct.pack('<2sIHHI', b'BM', 14 + len(dib_data), 0, 0, offset)
    return Image.open(BytesIO(file_header + dib_data))


# ---------------------------------------------------------------------------
# Web Resource Downloader
# ---------------------------------------------------------------------------

def download_web_resource(url, cache_dir, timeout=8):
    """
    Downloads an image or file from a URL to the local cache directory.
    Validates URL safety and returns the local file path.
    """
    parsed = urllib.parse.urlparse(url)
    if parsed.scheme not in ('http', 'https'):
        return None

    path_lower = parsed.path.lower()
    ext = os.path.splitext(path_lower)[1]
    if not ext or len(ext) > 6:
        ext = ".png"

    raw_stem = Path(parsed.path).stem or "web_asset"
    clean_stem = re.sub(r'[^a-zA-Z0-9_-]', '_', raw_stem)[:30]
    filename = f"{clean_stem}_{int(time.time()*1000)}{ext}"
    out_path = os.path.join(cache_dir, filename)

    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
        "Accept": "*/*"
    }
    req = urllib.request.Request(url, headers=headers)
    with urllib.request.urlopen(req, timeout=timeout) as resp, open(out_path, "wb") as f:
        f.write(resp.read())

    return out_path


# ---------------------------------------------------------------------------
# Format Inspection Helper
# ---------------------------------------------------------------------------

def inspect_available_formats(dataObject):
    """Inspects and returns a list of human-readable format names in dataObject."""
    formats = []
    try:
        enum = dataObject.EnumFormatEtc(pythoncom.DATADIR_GET)
        while True:
            fetc = enum.Next(1)
            if not fetc:
                break
            cf = fetc[0][0]
            try:
                name = win32clipboard.GetClipboardFormatName(cf)
            except Exception:
                standard = {
                    win32con.CF_TEXT: "CF_TEXT",
                    win32con.CF_BITMAP: "CF_BITMAP",
                    win32con.CF_DIB: "CF_DIB",
                    win32con.CF_DIBV5: "CF_DIBV5",
                    win32con.CF_UNICODETEXT: "CF_UNICODETEXT",
                    win32con.CF_HDROP: "CF_HDROP",
                    win32con.CF_LOCALE: "CF_LOCALE",
                }
                name = standard.get(cf, f"CF_{cf}")
            formats.append(name)
    except Exception:
        pass
    return formats


# ---------------------------------------------------------------------------
# Universal Payload Parser
# ---------------------------------------------------------------------------

def parse_dropped_data_object(dataObject, diag=None):
    """
    Universally extracts content from a COM IDataObject dragged from:
    - Google Chrome, Microsoft Edge, Mozilla Firefox
    - Windows File Explorer
    - Microsoft Office, Notepad, Electron apps
    Returns a normalized dictionary describing the input and its metadata.
    """
    cache_dir = get_cache_dir()

    # 1. Check for CF_HDROP (Local files dragged from Explorer / Desktop)
    try:
        fetc = (win32con.CF_HDROP, None, pythoncom.DVASPECT_CONTENT, -1, pythoncom.TYMED_HGLOBAL)
        stg = dataObject.GetData(fetc)
        if stg and stg.data:
            p_files = ctypes.c_uint32.from_buffer_copy(stg.data, 0).value
            is_wide = ctypes.c_int.from_buffer_copy(stg.data, 16).value
            raw_data = stg.data[p_files:]
            enc = "utf-16le" if is_wide else "latin1"
            paths = [p for p in raw_data.decode(enc, errors="ignore").split("\x00") if p and os.path.exists(p)]
            if paths:
                return {
                    "kind": "files",
                    "paths": paths,
                    "source": "File Explorer",
                    "count": len(paths)
                }
    except Exception:
        pass

    # 2. Check for CF_HTML (Dragged web images, links, or rich text from Chrome/Edge/Firefox)
    html_raw = ""
    try:
        fetc = (CF_HTML, None, pythoncom.DVASPECT_CONTENT, -1, pythoncom.TYMED_HGLOBAL)
        stg = dataObject.GetData(fetc)
        if stg and stg.data:
            html_raw = stg.data.decode("utf-8", errors="ignore")
    except Exception:
        pass

    # Check if HTML contains an <img> tag (Web Image Drag)
    if html_raw:
        img_match = re.search(r'<img[^>]+src=["\']([^"\']+)["\']', html_raw, re.IGNORECASE)
        if img_match:
            img_src = img_match.group(1).strip()
            # Case 2a: Base64 data URI image
            if img_src.startswith("data:image/"):
                try:
                    header, b64_data = img_src.split(",", 1)
                    img_bytes = base64.b64decode(b64_data)
                    ext = ".png"
                    if "jpeg" in header or "jpg" in header: ext = ".jpg"
                    elif "webp" in header: ext = ".webp"
                    cached_file = os.path.join(cache_dir, f"browser_image_{int(time.time()*1000)}{ext}")
                    with open(cached_file, "wb") as f:
                        f.write(img_bytes)
                    # Verify with PIL
                    with Image.open(cached_file) as im:
                        width, height = im.size
                        im_fmt = im.format
                    return {
                        "kind": "image",
                        "path": cached_file,
                        "source": "Browser Data Image",
                        "name": f"web_image_{width}x{height}",
                        "dimensions": (width, height),
                        "format": im_fmt
                    }
                except Exception as e:
                    if diag: diag.log("IMG_DATA_URI_ERROR", str(e), level="WARN")

            # Case 2b: Web image URL
            elif img_src.startswith(('http://', 'https://')):
                try:
                    cached_file = download_web_resource(img_src, cache_dir)
                    if cached_file and os.path.exists(cached_file):
                        with Image.open(cached_file) as im:
                            width, height = im.size
                            im_fmt = im.format
                        clean_name = re.sub(r'[^a-zA-Z0-9_-]', '_', Path(urllib.parse.urlparse(img_src).path).stem or "browser_image")[:25]
                        return {
                            "kind": "image",
                            "path": cached_file,
                            "url": img_src,
                            "source": "Browser Web Image",
                            "name": clean_name,
                            "dimensions": (width, height),
                            "format": im_fmt
                        }
                except Exception as e:
                    if diag: diag.log("IMG_DOWNLOAD_ERROR", str(e), level="WARN")
                    return {
                        "kind": "image_url",
                        "data": img_src,
                        "source": "Browser Image URL",
                        "name": "browser_image_url"
                    }

    # 3. Check for UniformResourceLocatorW (URL dragged from address bar, links, or bookmarks)
    url_text = ""
    for url_fmt in (CF_URL, CF_URL_A):
        try:
            fetc = (url_fmt, None, pythoncom.DVASPECT_CONTENT, -1, pythoncom.TYMED_HGLOBAL)
            stg = dataObject.GetData(fetc)
            if stg and stg.data:
                if url_fmt == CF_URL:
                    url_text = stg.data.decode("utf-16le", errors="ignore").rstrip("\x00").strip()
                else:
                    url_text = stg.data.decode("latin1", errors="ignore").rstrip("\x00").strip()
                if url_text:
                    break
        except Exception:
            pass

    if url_text and url_text.startswith(('http://', 'https://')):
        path_lower = urllib.parse.urlparse(url_text).path.lower()
        is_img_ext = any(path_lower.endswith(e) for e in ('.png', '.jpg', '.jpeg', '.webp', '.gif', '.bmp', '.svg'))
        is_doc_ext = any(path_lower.endswith(e) for e in ('.pdf', '.docx', '.xlsx', '.csv', '.txt'))

        if is_img_ext or is_doc_ext:
            try:
                cached_file = download_web_resource(url_text, cache_dir)
                if cached_file and os.path.exists(cached_file):
                    if is_img_ext:
                        with Image.open(cached_file) as im:
                            w, h = im.size
                            im_fmt = im.format
                        clean_name = Path(cached_file).stem
                        return {
                            "kind": "image",
                            "path": cached_file,
                            "url": url_text,
                            "source": "Browser Image URL",
                            "name": clean_name,
                            "dimensions": (w, h),
                            "format": im_fmt
                        }
                    else:
                        return {
                            "kind": "text_file" if is_doc_ext and path_lower.endswith('.txt') else "file",
                            "path": cached_file,
                            "url": url_text,
                            "source": "Browser Web Document",
                            "name": Path(cached_file).stem
                        }
            except Exception:
                pass

        # Webpage link
        clean_name = re.sub(r'[^a-zA-Z0-9_-]', '_', urllib.parse.urlparse(url_text).netloc)[:25] or "web_link"
        return {
            "kind": "url",
            "data": url_text,
            "source": "Browser Hyperlink",
            "name": clean_name,
            "url": url_text
        }

    # 4. Check for CF_DIB / CF_DIBV5 (In-memory bitmap image dragged from browser)
    for dib_fmt in (win32con.CF_DIBV5, win32con.CF_DIB):
        try:
            fetc = (dib_fmt, None, pythoncom.DVASPECT_CONTENT, -1, pythoncom.TYMED_HGLOBAL)
            stg = dataObject.GetData(fetc)
            if stg and stg.data:
                img = dib_bytes_to_image(stg.data)
                cached_file = os.path.join(cache_dir, f"browser_bitmap_{int(time.time()*1000)}.png")
                img.save(cached_file, "PNG")
                return {
                    "kind": "image",
                    "path": cached_file,
                    "source": "Browser Image Bitmap (DIB)",
                    "name": f"web_image_{img.width}x{img.height}",
                    "dimensions": img.size,
                    "format": "PNG"
                }
        except Exception:
            pass

    # 5. Check for CF_UNICODETEXT (Selected text from Chrome, Edge, Firefox, or any text editor)
    try:
        fetc = (win32con.CF_UNICODETEXT, None, pythoncom.DVASPECT_CONTENT, -1, pythoncom.TYMED_HGLOBAL)
        stg = dataObject.GetData(fetc)
        if stg and stg.data:
            text = stg.data.decode("utf-16le", errors="ignore").rstrip("\x00").strip()
            if text:
                if text.startswith(('http://', 'https://')):
                    return {
                        "kind": "url",
                        "data": text,
                        "source": "Browser Link",
                        "name": "web_link",
                        "url": text
                    }
                first_ln = text.splitlines()[0][:30].strip()
                clean_name = re.sub(r'[^a-zA-Z0-9_-]', '_', first_ln)[:20] or "browser_text"
                return {
                    "kind": "text",
                    "data": text,
                    "source": "Browser Selected Text",
                    "name": clean_name
                }
    except Exception:
        pass

    # 6. Check for HTML snippet text (if unicode text was empty but HTML is available)
    if html_raw:
        clean_text = re.sub(r'<[^>]+>', ' ', html_raw)
        clean_text = re.sub(r'\s+', ' ', clean_text).strip()
        if clean_text:
            first_ln = clean_text[:30].strip()
            clean_name = re.sub(r'[^a-zA-Z0-9_-]', '_', first_ln)[:20] or "html_snippet"
            return {
                "kind": "text",
                "data": clean_text,
                "source": "Browser HTML Content",
                "name": clean_name
            }

    return None


# ---------------------------------------------------------------------------
# COM IDropTarget Implementation
# ---------------------------------------------------------------------------

class NotchDropTargetCOM:
    _com_interfaces_ = [pythoncom.IID_IDropTarget]

    def __init__(self, on_drop_cb, on_hover_cb=None, diag=None):
        self.on_drop_cb = on_drop_cb
        self.on_hover_cb = on_hover_cb
        self.diag = diag or DiagnosticLogger()
        self.is_hovered = False

    def DragEnter(self, dataObject, keyState, pt, effect):
        self.is_hovered = True
        formats = inspect_available_formats(dataObject)
        self.diag.log_drag_enter(formats)

        if self.on_hover_cb:
            self.on_hover_cb(True, formats)
        return win32con.DROPEFFECT_COPY

    def DragOver(self, keyState, pt, effect):
        return win32con.DROPEFFECT_COPY

    def DragLeave(self):
        self.is_hovered = False
        self.diag.log_drag_leave()
        if self.on_hover_cb:
            self.on_hover_cb(False, [])
        return None

    def Drop(self, dataObject, keyState, pt, effect):
        self.is_hovered = False
        formats = inspect_available_formats(dataObject)
        if self.on_hover_cb:
            self.on_hover_cb(False, [])

        try:
            parsed = parse_dropped_data_object(dataObject, self.diag)
            self.diag.log_drop(formats, parsed_payload=parsed)
            if parsed and self.on_drop_cb:
                self.on_drop_cb(parsed)
        except Exception as e:
            self.diag.log_drop(formats, error=e)

        return win32con.DROPEFFECT_COPY


# ---------------------------------------------------------------------------
# Smart Notch Drop Target Window
# ---------------------------------------------------------------------------

class NotchDropZoneWindow:
    """
    Invisible, zero-clash drop zone positioned exactly over the notch pill.
    When idle: 100% invisible (alpha=1) and click-through (HTTRANSPARENT).
    During drag: Detects mouse drag, activates OLE IDropTarget, smoothly opens
    Nook in mid-air, and delivers browser text, images, URLs, and files to the plugin.
    """

    def __init__(self, on_item_received, on_open_request=None, app_logger=None):
        self.on_item_received = on_item_received
        self.on_open_request = on_open_request
        self.diag = DiagnosticLogger(app_logger=app_logger)
        self.hwnd = None
        self.hovered = False
        self.hover_formats = []
        self._thread = None
        self.width = 460
        self.height = 56

    def start(self):
        self._thread = threading.Thread(target=self._run_window, daemon=True)
        self._thread.start()

    def _on_hover_change(self, is_hovered, formats):
        self.hovered = is_hovered
        self.hover_formats = formats

        if self.hwnd:
            if is_hovered:
                # Smoothly open Nook notch so user sees the page slide down in mid-air!
                if self.on_open_request:
                    self.on_open_request()
                # Show subtle glowing glass indicator
                win32gui.SetLayeredWindowAttributes(self.hwnd, 0, 220, win32con.LWA_ALPHA)
                win32gui.InvalidateRect(self.hwnd, None, True)
            else:
                # Restore to 100% invisible
                win32gui.SetLayeredWindowAttributes(self.hwnd, 0, 1, win32con.LWA_ALPHA)
                win32gui.InvalidateRect(self.hwnd, None, True)

    def _on_drop_delivered(self, parsed_item):
        if self.hwnd:
            win32gui.SetLayeredWindowAttributes(self.hwnd, 0, 1, win32con.LWA_ALPHA)
            win32gui.InvalidateRect(self.hwnd, None, True)

        if self.on_open_request:
            self.on_open_request()
        if self.on_item_received:
            self.on_item_received(parsed_item)

    def _run_window(self):
        pythoncom.OleInitialize()

        screen_w = win32api.GetSystemMetrics(0)
        x = (screen_w - self.width) // 2
        y = 0

        def wnd_proc(hwnd, msg, wparam, lparam):
            if msg == win32con.WM_NCHITTEST:
                # Only accept hit-testing when a mouse drag is actively in progress!
                # When normal mouse interaction occurs (clicks, hovers), return HTTRANSPARENT
                # so everything passes through to Nook without interception!
                if is_mouse_dragging():
                    return win32con.HTCLIENT
                return win32con.HTTRANSPARENT

            elif msg == win32con.WM_PAINT:
                hdc, ps = win32gui.BeginPaint(hwnd)
                rect = win32gui.GetClientRect(hwnd)

                if self.hovered:
                    # Sleek glowing halo indicator during active drag-hover
                    # Glass dark pill with accent border
                    brush = win32gui.CreateSolidBrush(win32api.RGB(15, 23, 42))  # Slate dark
                    win32gui.FillRect(hdc, rect, brush)
                    win32gui.DeleteObject(brush)

                    # Accent border (#3D8BFF)
                    pen = win32gui.CreatePen(win32con.PS_SOLID, 2, win32api.RGB(61, 139, 255))
                    old_pen = win32gui.SelectObject(hdc, pen)
                    null_brush = win32gui.GetStockObject(win32con.NULL_BRUSH)
                    old_brush = win32gui.SelectObject(hdc, null_brush)
                    win32gui.RoundRect(hdc, 2, 2, rect[2] - 2, rect[3] - 2, 16, 16)
                    win32gui.SelectObject(hdc, old_pen)
                    win32gui.SelectObject(hdc, old_brush)
                    win32gui.DeleteObject(pen)

                    # Centered drop indicator text
                    win32gui.SetBkMode(hdc, win32con.TRANSPARENT)
                    win32gui.SetTextColor(hdc, win32api.RGB(255, 255, 255))
                    hint_text = "Release to Convert (Text, Image, Link, File)"
                    win32gui.DrawText(
                        hdc,
                        hint_text,
                        -1,
                        (0, 0, rect[2], rect[3]),
                        win32con.DT_CENTER | win32con.DT_VCENTER | win32con.DT_SINGLELINE
                    )

                win32gui.EndPaint(hwnd, ps)
                return 0

            elif msg == win32con.WM_DESTROY:
                win32gui.PostQuitMessage(0)
                return 0

            return win32gui.DefWindowProc(hwnd, msg, wparam, lparam)

        wc = win32gui.WNDCLASS()
        wc.lpfnWndProc = wnd_proc
        wc.lpszClassName = "NookDropZoneSmart"
        wc.hCursor = win32gui.LoadCursor(0, win32con.IDC_ARROW)

        try:
            atom = win32gui.RegisterClass(wc)
        except Exception:
            atom = "NookDropZoneSmart"

        self.hwnd = win32gui.CreateWindowEx(
            win32con.WS_EX_TOPMOST | win32con.WS_EX_TOOLWINDOW | win32con.WS_EX_LAYERED | win32con.WS_EX_NOACTIVATE,
            atom,
            "Nook Smart Drop Target",
            win32con.WS_POPUP,
            x, y, self.width, self.height,
            0, 0, 0, None
        )

        # Set 100% invisible by default (alpha=1)
        win32gui.SetLayeredWindowAttributes(self.hwnd, 0, 1, win32con.LWA_ALPHA)
        win32gui.ShowWindow(self.hwnd, win32con.SW_SHOWNOACTIVATE)

        # Register OLE IDropTarget
        drop_target = NotchDropTargetCOM(
            on_drop_cb=self._on_drop_delivered,
            on_hover_cb=self._on_hover_change,
            diag=self.diag
        )
        com_target = pythoncom.WrapObject(drop_target, pythoncom.IID_IDropTarget)
        pythoncom.RegisterDragDrop(self.hwnd, com_target)
        self.diag.log("WINDOW_INITIALIZED", {"hwnd": self.hwnd, "rect": (x, y, self.width, self.height)})

        # Message loop
        win32gui.PumpMessages()

        # Cleanup on exit
        try:
            pythoncom.RevokeDragDrop(self.hwnd)
        except Exception:
            pass
        pythoncom.CoUninitialize()
