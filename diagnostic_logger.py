"""
Diagnostic Logger for Nook Drop Changer.
Provides structured diagnostic logging for OLE Drag & Drop events,
format inspection, payload details, and conversion lifecycle.
Writes to changer_diagnostic.log and forwards warnings/errors to Nook.
"""

import os
import sys
import time
import tempfile
from datetime import datetime
from pathlib import Path

class DiagnosticLogger:
    def __init__(self, log_path=None, app_logger=None):
        if log_path is None:
            # Store outside plugin folder so Nook's live watcher doesn't trigger endless reloads
            self.log_path = os.path.join(tempfile.gettempdir(), "nook_changer_diagnostic.log")
        else:
            self.log_path = log_path
        self.app_logger = app_logger
        self._init_log_file()

    def _init_log_file(self):
        try:
            with open(self.log_path, "a", encoding="utf-8") as f:
                f.write(f"\n--- Diagnostic Session Started at {datetime.now().isoformat()} ---\n")
        except Exception:
            pass

    def log(self, event, details=None, level="INFO"):
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
        detail_str = ""
        if details:
            if isinstance(details, dict):
                parts = []
                for k, v in details.items():
                    # Mask long or sensitive text
                    if k in ("text", "data") and isinstance(v, str):
                        v_str = f"'{v[:40]}...' (len={len(v)})" if len(v) > 40 else repr(v)
                    else:
                        v_str = repr(v)
                    parts.append(f"{k}={v_str}")
                detail_str = " | " + ", ".join(parts)
            else:
                detail_str = f" | {details}"

        msg = f"[{now_str}] [{level:<5}] [{event}] {detail_str}\n"

        # Write to log file
        try:
            with open(self.log_path, "a", encoding="utf-8") as f:
                f.write(msg)
        except Exception:
            pass

        # Write to stderr / Nook host log if app_logger available
        if self.app_logger:
            try:
                self.app_logger(f"[DropChanger] [{event}] {detail_str}", "info" if level == "INFO" else "warn")
            except Exception:
                pass

    def log_drag_enter(self, formats):
        self.log("DRAG_ENTER", {"available_formats": formats, "format_count": len(formats)})

    def log_drag_over(self):
        pass  # keep log clean, do not flood on every drag_over tick

    def log_drag_leave(self):
        self.log("DRAG_LEAVE", "Drag exited notch area")

    def log_drop(self, formats, parsed_payload=None, error=None):
        details = {
            "available_formats": formats,
            "format_count": len(formats)
        }
        if error:
            details["error"] = str(error)
            self.log("DROP_FAILED", details, level="ERROR")
        elif parsed_payload:
            details["kind"] = parsed_payload.get("kind")
            details["source"] = parsed_payload.get("source")
            details["name"] = parsed_payload.get("name")
            if "path" in parsed_payload:
                details["file_path"] = parsed_payload["path"]
                try:
                    details["file_size_bytes"] = os.path.getsize(parsed_payload["path"])
                except Exception:
                    pass
            if "url" in parsed_payload:
                details["url"] = parsed_payload["url"]
            if "data" in parsed_payload and isinstance(parsed_payload["data"], str):
                details["text_length"] = len(parsed_payload["data"])
            self.log("DROP_SUCCESS", details, level="INFO")
        else:
            self.log("DROP_EMPTY", details, level="WARN")

    def log_conversion(self, input_kind, target_fmt, success, output_path=None, error=None):
        details = {
            "input_kind": input_kind,
            "target_format": target_fmt,
            "success": success
        }
        if output_path:
            details["output_path"] = output_path
        if error:
            details["error"] = str(error)
        self.log("CONVERSION", details, level="INFO" if success else "ERROR")
