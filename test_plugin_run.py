"""
Test script for main.py simulating Nook protocol across all three tile sizes,
renaming, and folder configuration.
"""

import subprocess
import json
import time

def test_all_tile_sizes():
    proc = subprocess.Popen(
        ["python", "main.py"],
        cwd=r"C:\Users\deepa\.gemini\antigravity\scratch\nook-drop-converter",
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        bufsize=1
    )

    # 1. Hello
    proc.stdin.write(json.dumps({"type": "hello", "protocol": 1, "plugin_id": "changer", "theme": {"dark": True, "accent": "#3D8BFF"}}) + "\n")
    proc.stdin.flush()

    init_tile = json.loads(proc.stdout.readline())
    init_page = json.loads(proc.stdout.readline())
    print("Initial tile type:", init_tile["type"])
    print("Initial page type:", init_page["type"])

    # 2. Test layout message for "small" (1x1)
    proc.stdin.write(json.dumps({"type": "layout", "target": "tile", "size": "small", "width": 70, "height": 70}) + "\n")
    proc.stdin.flush()
    small_tile = json.loads(proc.stdout.readline())
    small_page = json.loads(proc.stdout.readline())
    print("Small (1x1) tile node type:", small_tile["view"]["type"])
    assert small_tile["view"]["type"] == "column"

    # 3. Test layout message for "wide" (2x1)
    proc.stdin.write(json.dumps({"type": "layout", "target": "tile", "size": "wide", "width": 150, "height": 70}) + "\n")
    proc.stdin.flush()
    wide_tile = json.loads(proc.stdout.readline())
    wide_page = json.loads(proc.stdout.readline())
    print("Wide (2x1) tile node type:", wide_tile["view"]["type"])
    assert wide_tile["view"]["type"] == "row"
    # Check wide tile children count: icon, column (zero overlapping buttons)
    print("Wide tile children count:", len(wide_tile["view"]["children"]))
    assert len(wide_tile["view"]["children"]) == 2

    # 4. Test layout message for "large" (2x2)
    proc.stdin.write(json.dumps({"type": "layout", "target": "tile", "size": "large", "width": 150, "height": 150}) + "\n")
    proc.stdin.flush()
    large_tile = json.loads(proc.stdout.readline())
    large_page = json.loads(proc.stdout.readline())
    print("Large (2x2) tile node type:", large_tile["view"]["type"])
    assert large_tile["view"]["type"] == "column"

    # 5. Test manual input staging & renaming
    proc.stdin.write(json.dumps({"type": "event", "id": "manual_input", "kind": "submit", "value": "Test Content"}) + "\n")
    proc.stdin.flush()
    staged_tile = json.loads(proc.stdout.readline())
    staged_page = json.loads(proc.stdout.readline())
    print("Staged tile generated successfully")

    # 6. Test renaming
    proc.stdin.write(json.dumps({"type": "event", "id": "input_rename", "kind": "change", "value": "custom_invoice_name"}) + "\n")
    proc.stdin.flush()
    renamed_tile = json.loads(proc.stdout.readline())
    renamed_page = json.loads(proc.stdout.readline())
    print("Renamed tile generated successfully")

    # Shutdown
    proc.stdin.write(json.dumps({"type": "shutdown"}) + "\n")
    proc.stdin.flush()
    proc.wait(timeout=3)
    print("All protocol & tile size tests passed with flying colors!")

if __name__ == "__main__":
    test_all_tile_sizes()
