"""
Nook Drop Changer Plugin
Universal drag & drop converter for the notch.
Completely clean tile layout with ZERO icon overlaps across small (1x1), wide (2x1), and large (2x2).
Supports seamless drag-to-notch auto-open, shelf drop bridge, folder relocation, and pre-conversion renaming.
"""

import os
import sys
import time
import threading
import ctypes
from ctypes import wintypes
from datetime import datetime
from pathlib import Path

import nook
import converter_engine as ce
import notch_drop_target as ndt

app = nook.Plugin()


def get_initial_tile_size():
    """Reads configured tile size from Nook's config.toml at startup."""
    cfg_path = os.path.expandvars(r"%APPDATA%\Nook\config.toml")
    if os.path.exists(cfg_path):
        try:
            with open(cfg_path, "r", encoding="utf-8") as f:
                for line in f:
                    if '"plugin:changer"' in line and '=' in line:
                        val = line.split('=')[1].strip().strip('"\'')
                        if val in ("small", "wide", "large"):
                            return val
        except Exception:
            pass
    return "wide"


# Persistent configuration & history
config = app.load("config", {})
history = app.load("history", [])

# Output directory (relocation)
default_output = ce.get_default_output_dir()
current_output_dir = config.get("output_dir", default_output)
if not os.path.exists(current_output_dir):
    try:
        os.makedirs(current_output_dir, exist_ok=True)
    except Exception:
        current_output_dir = default_output

# Runtime state
current_tile_size = get_initial_tile_size()  # "small", "wide", or "large"
current_input = None
selected_format_idx = 0
last_result = None
manual_text_buffer = ""
status_banner = ""
folder_preset_idx = 0  # 0: Downloads, 1: Desktop, 2: Documents, 3: Source Folder, 4: Custom

# Shelf watcher tracking
shelf_path = os.path.expandvars(r"%APPDATA%\Nook\shelf.txt")
last_shelf_line_count = 0
if os.path.exists(shelf_path):
    try:
        with open(shelf_path, "r", encoding="utf-8", errors="ignore") as f:
            last_shelf_line_count = len([ln.strip() for ln in f if ln.strip()])
    except Exception:
        pass


def save_state():
    config["output_dir"] = current_output_dir
    app.save("config", config)
    app.save("history", history[:20])


def get_input_icon(kind):
    if kind == "image":
        return "file-image"
    elif kind == "spreadsheet":
        return "file-xls"
    elif kind == "pdf":
        return "file-pdf"
    elif kind == "word":
        return "file-doc"
    elif kind in ("text", "text_file", "html"):
        return "file-text"
    elif kind in ("url", "image_url"):
        return "globe"
    return "file"


def get_current_format_options():
    if not current_input:
        return []
    return ce.get_available_target_formats(current_input)


# ---------------------------------------------------------------------------
# Zero-Overlap Tile Layouts (Small, Wide, and Large)
# ---------------------------------------------------------------------------

def render_tile():
    """
    Renders tile with absolutely ZERO interior buttons.
    In Nook, clicking any tile naturally opens the page.
    Omitting interior buttons prevents any possible icon-over-text collision.
    """
    formats = get_current_format_options()
    global selected_format_idx
    if formats and selected_format_idx >= len(formats):
        selected_format_idx = 0

    target_fmt = formats[selected_format_idx]["id"] if formats else "pdf"

    # Determine icon and subtitle text based on state
    if current_input:
        tile_icon = get_input_icon(current_input.get("kind"))
        icon_color = "accent"
        staged_name = current_input.get("custom_name") or current_input.get("name") or "File"
        subtitle = f"Ready -> {target_fmt.upper()}"
        short_label = target_fmt.upper()
    elif last_result and os.path.exists(last_result.get("path", "")):
        tile_icon = "check-circle"
        icon_color = "green"
        subtitle = last_result.get("name", "Converted!")
        short_label = "Done"
    else:
        tile_icon = "swap"
        icon_color = "accent"
        subtitle = "Drop files or paste"
        short_label = "Drop"

    # STYLE 1: SMALL (1x1 square, ~70x70)
    if current_tile_size == "small":
        return nook.column([
            nook.icon(tile_icon, size=28, color=icon_color, tooltip="Drop Changer (Click to open)"),
            nook.text(short_label, size="small", bold=True, color=icon_color),
        ], align="center", justify="center", gap=2)

    # STYLE 2: WIDE (2x1 rectangle, ~150x70)
    # Icon on the left + full unobstructed text column on the right. Zero buttons!
    elif current_tile_size == "wide":
        return nook.row([
            nook.icon(tile_icon, size=28, color=icon_color, tooltip="Drop Changer (Click to open)"),
            nook.column([
                nook.text("Drop Changer", size="small", color="dim", bold=True),
                nook.text(subtitle, size="body", wrap=False),
            ], gap=0),
        ], gap=10, align="center")

    # STYLE 3: LARGE (2x2 square, ~150x150)
    # Balanced, centered widget design that fills the square beautifully
    else:
        return nook.column([
            nook.icon(tile_icon, size=40, color=icon_color, tooltip="Drop Changer (Click to open)"),
            nook.text("Drop Changer", size="body", bold=True),
            nook.text(subtitle, size="small", color="dim", wrap=False),
        ], align="center", justify="center", gap=4)


# ---------------------------------------------------------------------------
# Main Draw (Tile & Expanded Glass Page)
# ---------------------------------------------------------------------------

def draw():
    # 1. Update Tile
    app.tile(render_tile())

    # 2. Update Expanded Page
    formats = get_current_format_options()
    global selected_format_idx
    if formats and selected_format_idx >= len(formats):
        selected_format_idx = 0

    target_fmt = formats[selected_format_idx]["id"] if formats else "pdf"
    target_label = formats[selected_format_idx]["label"] if formats else "PDF"

    page_children = []

    # Top Quick Toolbar
    page_children.append(nook.row([
        nook.button("btn_paste", label="Paste Clipboard", icon="clipboard-text", primary=True),
        nook.button("btn_browse_folder", label="Output Folder", icon="folder-open"),
        nook.spacer(),
        nook.button("btn_clear_all", icon="trash", tooltip="Clear / Reset"),
    ], align="center"))

    # Staged File Card or Drop Zone
    if not current_input:
        page_children.append(nook.card([
            nook.column([
                nook.row([
                    nook.icon("arrow-down-to-bracket", size=24, color="accent"),
                    nook.column([
                        nook.text("Seamless Drag & Drop", bold=True),
                        nook.text("Drag any file towards the Notch — it opens automatically!", size="small", color="dim"),
                    ], gap=1),
                ], align="center"),
            ], gap=4)
        ], padding=12))

        # Direct text input box
        page_children.append(nook.text_input(
            "manual_input",
            value=manual_text_buffer,
            hint="Type or paste text here and press Enter to convert..."
        ))

    else:
        # Input Staged Card
        inp_kind = current_input.get("kind", "file")
        inp_preview = current_input.get("preview", current_input.get("name", "Input"))
        inp_source = current_input.get("source", "Source")

        page_children.append(nook.card([
            nook.row([
                nook.icon(get_input_icon(inp_kind), size=24, color="accent"),
                nook.column([
                    nook.text(inp_preview, bold=True),
                    nook.text(f"{inp_source} • Ready to convert", size="small", color="dim"),
                ], gap=1),
                nook.spacer(),
                nook.button("btn_dismiss_input", icon="x", tooltip="Remove this input"),
            ], align="center")
        ], padding=10))

        # Target Format Selector Chips
        chip_options = [f["label"] for f in formats]
        page_children.append(nook.column([
            nook.text("Target Format:", size="small", color="dim", bold=True),
            nook.chips("fmt_chips", chip_options, selected=selected_format_idx),
        ], gap=4))

        # Renaming Option Before Conversion
        current_name_val = current_input.get("custom_name") or current_input.get("name") or "document"
        page_children.append(nook.card([
            nook.column([
                nook.row([
                    nook.icon("pencil-simple", size=18, color="accent"),
                    nook.text("Output File Name:", size="small", color="dim", bold=True),
                    nook.spacer(),
                    nook.text(f".{target_fmt}", size="small", color="accent", bold=True),
                ], align="center"),
                nook.text_input("input_rename", value=current_name_val, hint="Enter filename (without extension)..."),
                nook.row([
                    nook.button("btn_name_orig", label="Original Name", icon="arrow-counter-clockwise"),
                    nook.button("btn_name_date", label="+ Today's Date", icon="calendar"),
                ], align="center"),
            ], gap=6)
        ], padding=10))

        # Convert Action Button
        page_children.append(nook.button(
            "btn_convert_now",
            label=f"Convert to {target_label} Now",
            icon="lightning",
            primary=True,
            width=0
        ))

    # Destination Folder Section (Relocation)
    short_folder = current_output_dir
    if len(short_folder) > 38:
        short_folder = "..." + short_folder[-35:]

    page_children.append(nook.card([
        nook.column([
            nook.row([
                nook.icon("folder", size=18, color="accent"),
                nook.column([
                    nook.text("Destination Folder", size="small", color="dim", bold=True),
                    nook.text(short_folder, size="small", color="text", wrap=False),
                ], gap=0),
                nook.spacer(),
                nook.button("btn_browse_folder", label="Change", icon="folder-open"),
                nook.button("btn_open_folder", icon="arrow-square-out", tooltip="Open Folder in Explorer"),
            ], align="center"),
            nook.chips("folder_preset_chips", ["Downloads", "Desktop", "Documents", "Source Folder"], selected=folder_preset_idx),
        ], gap=6)
    ], padding=10))

    # Last Result Banner (if available)
    if last_result and os.path.exists(last_result.get("path", "")):
        page_children.append(nook.card([
            nook.row([
                nook.icon("check-circle", size=24, color="green"),
                nook.column([
                    nook.text("Conversion Successful!", bold=True, color="green"),
                    nook.text(f"{last_result['name']} ({last_result['size_str']})", size="body"),
                    nook.text(f"Saved: {last_result['path']}", size="small", color="dim", wrap=False),
                ], gap=1),
                nook.spacer(),
                nook.button("btn_open_last", label="Open", icon="arrow-square-out", primary=True),
                nook.button("btn_copy_path", icon="copy", tooltip="Copy File Path"),
            ], align="center")
        ], padding=10))

    # Recent Conversions History List
    if history:
        page_children.append(nook.row([
            nook.text("Recent Conversions", size="small", color="dim", bold=True),
            nook.spacer(),
            nook.text(f"{len(history)} files", size="small", color="faint"),
        ], align="center"))

        history_items = []
        for h in history[:10]:
            history_items.append(nook.item(
                id=h["path"],
                title=h["name"],
                subtitle=f"{h.get('fmt', '').upper()} • {h.get('size', '')} • {h.get('time', '')}",
                icon=h.get("icon", "file-arrow-up"),
                trailing="Open"
            ))
        page_children.append(nook.list_view(history_items))

    # Page Header Actions
    page_actions = [
        {"id": "btn_paste", "icon": "clipboard-text", "tooltip": "Paste Clipboard"},
        {"id": "btn_open_folder", "icon": "folder-open", "tooltip": "Open Output Folder"},
    ]

    app.page(nook.column(page_children, gap=8, scroll=True), actions=page_actions)


# ---------------------------------------------------------------------------
# Conversion Action
# ---------------------------------------------------------------------------

def perform_conversion():
    global current_input, last_result, status_banner
    if not current_input:
        return

    formats = get_current_format_options()
    if not formats:
        return

    target_fmt = formats[selected_format_idx]["id"]
    custom_name = current_input.get("custom_name") or current_input.get("name") or "converted"

    output_dir = current_output_dir
    if folder_preset_idx == 3 and current_input.get("path"):
        output_dir = os.path.dirname(current_input["path"])

    res = ce.execute_conversion(
        current_input,
        target_fmt,
        output_dir=output_dir,
        custom_name=custom_name
    )

    if res["success"]:
        out_path = res["output_path"]
        last_result = {
            "path": out_path,
            "name": res["name"],
            "size_str": res["size_str"],
            "time": datetime.now().strftime("%H:%M:%S"),
            "fmt": target_fmt,
            "icon": formats[selected_format_idx].get("icon", "file")
        }
        history.insert(0, {
            "path": out_path,
            "name": res["name"],
            "size": res["size_str"],
            "time": datetime.now().strftime("%b %d %H:%M"),
            "fmt": target_fmt,
            "icon": formats[selected_format_idx].get("icon", "file")
        })
        save_state()
        current_input = None
        status_banner = f"Converted to {res['name']}"
        app.badge_glyph("check-circle")
    else:
        status_banner = f"Error: {res.get('error')}"
        app.log(f"Conversion error: {res.get('error')}", "error")

    draw()


def load_clipboard():
    global current_input, selected_format_idx, status_banner
    clip = ce.read_windows_clipboard()
    if clip:
        if clip["kind"] == "files" and clip.get("data"):
            first_path = clip["data"][0]
            file_info = ce.inspect_file(first_path)
            if file_info:
                current_input = file_info
                current_input["custom_name"] = file_info.get("name")
        else:
            current_input = clip
            current_input["custom_name"] = clip.get("name")

        selected_format_idx = 0
        status_banner = "Clipboard loaded"
        app.badge_glyph("lightning")
        draw()


def stage_dropped_file(file_path):
    global current_input, selected_format_idx
    if not os.path.exists(file_path):
        return
    file_info = ce.inspect_file(file_path)
    if file_info:
        current_input = file_info
        current_input["custom_name"] = file_info.get("name")
        selected_format_idx = 0
        app.active(True)
        app.badge_glyph("lightning")
        app.show()
        draw()


# ---------------------------------------------------------------------------
# Background Drag Detector (Auto-opens Notch when dragging towards top screen)
# ---------------------------------------------------------------------------

_last_drag_open = 0.0

def start_drag_listener():
    """
    Detects when a user is dragging a file towards the top notch.
    Automatically slides open Drop Changer's glass page in mid-air so the file
    can be dropped directly onto the panel without clicking first!
    """
    def _drag_loop():
        global _last_drag_open
        user32 = ctypes.windll.user32

        class POINT(ctypes.Structure):
            _fields_ = [('x', ctypes.c_long), ('y', ctypes.c_long)]

        pt = POINT()
        screen_w = user32.GetSystemMetrics(0)

        while True:
            time.sleep(0.08)
            try:
                # Check left mouse button held down (drag in progress)
                lbutton = user32.GetAsyncKeyState(0x01)
                if lbutton & 0x8000:
                    user32.GetCursorPos(ctypes.byref(pt))
                    # Notch is at top center of screen (y <= 50, within center 50%)
                    if pt.y <= 50 and (screen_w * 0.25 <= pt.x <= screen_w * 0.75):
                        now = time.monotonic()
                        if now - _last_drag_open > 2.0:
                            _last_drag_open = now
                            app.show()
            except Exception:
                pass

    threading.Thread(target=_drag_loop, daemon=True).start()


# ---------------------------------------------------------------------------
# Shelf Drops Watcher
# ---------------------------------------------------------------------------

@app.every(0.4)
def check_shelf_drops():
    """Detects files dropped on Nook's shelf/notch and opens Drop Changer."""
    global last_shelf_line_count
    if not os.path.exists(shelf_path):
        return

    try:
        with open(shelf_path, "r", encoding="utf-8", errors="ignore") as f:
            lines = [ln.strip() for ln in f if ln.strip()]

        current_count = len(lines)
        if current_count > last_shelf_line_count:
            last_shelf_line_count = current_count
            newest_file = lines[-1]
            if os.path.exists(newest_file):
                stage_dropped_file(newest_file)
        elif current_count != last_shelf_line_count:
            last_shelf_line_count = current_count
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Event Handlers
# ---------------------------------------------------------------------------

@app.on("tile_paste")
@app.on("btn_paste")
def on_paste(e):
    load_clipboard()


@app.on("tile_convert")
@app.on("btn_convert_now")
def on_convert(e):
    perform_conversion()


@app.on("btn_open_last")
def on_open_last(e):
    if last_result and os.path.exists(last_result.get("path", "")):
        app.open(last_result["path"])


@app.on("btn_copy_path")
def on_copy_path(e):
    if last_result and os.path.exists(last_result.get("path", "")):
        app.copy(last_result["path"])


@app.on("btn_open_folder")
def on_open_folder(e):
    app.open(current_output_dir)


def on_folder_picked(new_dir):
    global current_output_dir, folder_preset_idx
    current_output_dir = new_dir
    folder_preset_idx = 4  # custom
    save_state()
    draw()


@app.on("btn_browse_folder")
def on_browse_folder(e):
    ce.open_folder_picker_async(current_output_dir, on_folder_picked)


@app.on("folder_preset_chips")
def on_folder_presets(e):
    global current_output_dir, folder_preset_idx
    idx = e.get("value", {}).get("index", 0)
    folder_preset_idx = idx

    user_home = Path.home()
    if idx == 0:
        current_output_dir = ce.get_default_output_dir()
    elif idx == 1:
        current_output_dir = str(user_home / "Desktop")
    elif idx == 2:
        current_output_dir = str(user_home / "Documents")
    elif idx == 3:
        if current_input and current_input.get("path"):
            current_output_dir = os.path.dirname(current_input["path"])
        else:
            current_output_dir = ce.get_default_output_dir()

    save_state()
    draw()


@app.on("input_rename")
def on_input_rename(e):
    global current_input
    if current_input:
        val = e.get("value", "").strip()
        if val:
            current_input["custom_name"] = val
            draw()


@app.on("btn_name_orig")
def on_name_orig(e):
    global current_input
    if current_input:
        orig = current_input.get("name", "converted")
        current_input["custom_name"] = orig
        draw()


@app.on("btn_name_date")
def on_name_date(e):
    global current_input
    if current_input:
        base = current_input.get("custom_name") or current_input.get("name") or "converted"
        today = datetime.now().strftime("%Y%m%d")
        current_input["custom_name"] = f"{base}_{today}"
        draw()


@app.on("fmt_chips")
def on_fmt_chips(e):
    global selected_format_idx
    val = e.get("value", {})
    if isinstance(val, dict) and "index" in val:
        selected_format_idx = val["index"]
    draw()


@app.on("btn_clear_all")
def on_clear_all(e):
    global current_input, manual_text_buffer, last_result
    current_input = None
    manual_text_buffer = ""
    app.clear_badge()
    app.active(False)
    draw()


@app.on("btn_dismiss_input")
def on_dismiss_input(e):
    global current_input
    current_input = None
    app.active(False)
    draw()


@app.on("manual_input")
def on_manual_input(e):
    global current_input, manual_text_buffer, selected_format_idx
    if e.get("kind") == "submit":
        text_val = e.get("value", "").strip()
        if text_val:
            first_ln = text_val.splitlines()[0][:30]
            clean_name = ce.safe_filename(first_ln, 20) or "text_note"
            current_input = {
                "kind": "text",
                "data": text_val,
                "preview": f'"{first_ln}..." ({len(text_val)} chars)',
                "source": "Manual Input",
                "name": clean_name,
                "custom_name": clean_name
            }
            manual_text_buffer = ""
            selected_format_idx = 0
            draw()


@app.on_message("files")
def on_files(msg):
    paths = msg.get("paths", [])
    if paths:
        stage_dropped_file(paths[0])


@app.on_message("layout")
def on_layout(msg):
    global current_tile_size
    if msg.get("target") == "tile":
        size = msg.get("size")
        if size in ("small", "wide", "large"):
            current_tile_size = size
            draw()


@app.on_message("tile_click")
def on_tile_click(msg):
    app.show()


@app.on_message("hotkey")
def on_hotkey(msg):
    app.show()


@app.on_message("event")
def on_generic_event(e):
    event_id = e.get("id", "")
    if event_id and os.path.exists(event_id):
        app.open(event_id)


# ---------------------------------------------------------------------------
# Universal Notch Drop Target Integration (Browser & All Platforms)
# ---------------------------------------------------------------------------

def on_universal_item_received(item):
    """
    Handles universal drag-and-drop items from web browsers (Chrome, Edge, Firefox),
    desktop apps, or Windows File Explorer.
    """
    global current_input, selected_format_idx
    kind = item.get("kind")
    source = item.get("source", "Drop Target")

    if kind in ("file", "image"):
        file_path = item.get("path")
        if file_path and os.path.exists(file_path):
            file_info = ce.inspect_file(file_path)
            if file_info:
                current_input = file_info
                current_input["source"] = source
                current_input["custom_name"] = item.get("name") or file_info.get("name")
                if "dimensions" in item:
                    current_input["dimensions"] = item["dimensions"]

    elif kind == "files":
        paths = item.get("paths", [])
        if paths and os.path.exists(paths[0]):
            stage_dropped_file(paths[0])
            return

    elif kind == "text":
        text_val = item.get("data", "").strip()
        first_ln = text_val.splitlines()[0][:30] if text_val else "text_note"
        clean_name = item.get("name") or ce.safe_filename(first_ln, 20) or "browser_text"
        current_input = {
            "kind": "text",
            "data": text_val,
            "preview": f'"{first_ln}..." ({len(text_val)} chars)',
            "source": source,
            "name": clean_name,
            "custom_name": clean_name
        }

    elif kind == "html":
        html_val = item.get("data", "").strip()
        clean_text = re.sub(r'<[^>]+>', ' ', html_val)
        first_ln = clean_text[:30].strip() or "html_snippet"
        clean_name = item.get("name") or "html_snippet"
        current_input = {
            "kind": "html",
            "data": html_val,
            "preview": f'HTML Content ({len(html_val)} chars)',
            "source": source,
            "name": clean_name,
            "custom_name": clean_name
        }

    elif kind in ("url", "image_url"):
        url_val = item.get("data") or item.get("url", "")
        clean_name = item.get("name") or "web_link"
        current_input = {
            "kind": "url",
            "data": url_val,
            "url": url_val,
            "preview": f'Web Link: {url_val[:45]}...',
            "source": source,
            "name": clean_name,
            "custom_name": clean_name
        }

    if current_input:
        selected_format_idx = 0
        app.active(True)
        app.badge_glyph("lightning")
        app.show()
        draw()


try:
    drop_zone_win = ndt.NotchDropZoneWindow(
        on_item_received=on_universal_item_received,
        on_open_request=lambda: app.show(),
        app_logger=lambda msg, lvl: app.log(msg, lvl)
    )
    drop_zone_win.start()
except Exception as e:
    app.log(f"Universal drop target init error: {e}", "warn")

# Start background drag listener
start_drag_listener()

# Initial draw & run
draw()
app.run()
