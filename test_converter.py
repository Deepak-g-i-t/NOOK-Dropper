"""
Test suite for converter_engine.py
Verifies Text, Image, PDF, Excel, and Word conversions.
"""

import os
import tempfile
from pathlib import Path
from PIL import Image
import converter_engine as ce

def test_all():
    tmp = tempfile.gettempdir()
    print("Testing converter engine...")

    # 1. Text conversions
    sample_text = """# Project Plan & Status Report
This is a test of the **Nook Drop Changer** plugin.

## Tasks Completed
- Implemented Nook live-mode protocol
- Added drag-and-drop file support
- Added clipboard inspection for browser selections
1. Step 1: Initialize
2. Step 2: Convert

    def test():
        return True
"""
    txt_input = {"kind": "text", "data": sample_text, "name": "project_report"}

    for fmt in ["docx", "pdf", "xlsx", "txt", "png", "html"]:
        res = ce.execute_conversion(txt_input, fmt, output_dir=tmp)
        assert res["success"], f"Text to {fmt} failed: {res.get('error')}"
        assert os.path.exists(res["output_path"])
        print(f"  [PASS] Text -> {fmt}: {res['name']} ({res['size_str']})")

    # 2. Tabular Text to Excel
    table_csv = "Item,Quantity,Price,Status\nWidget A,10,19.99,In Stock\nWidget B,5,49.50,Backorder\nWidget C,25,4.99,In Stock"
    csv_input = {"kind": "text", "data": table_csv, "name": "inventory"}
    res_xls = ce.execute_conversion(csv_input, "xlsx", output_dir=tmp)
    assert res_xls["success"], f"CSV text to xlsx failed: {res_xls.get('error')}"
    print(f"  [PASS] CSV Text -> xlsx: {res_xls['name']}")

    # 3. Image conversions
    test_img = Image.new("RGBA", (300, 200), (45, 120, 240, 255))
    img_input = {"kind": "image", "data": test_img, "name": "test_graphic"}

    for fmt in ["jpeg", "png", "webp", "pdf", "docx", "ico"]:
        res = ce.execute_conversion(img_input, fmt, output_dir=tmp)
        assert res["success"], f"Image to {fmt} failed: {res.get('error')}"
        assert os.path.exists(res["output_path"])
        print(f"  [PASS] Image -> {fmt}: {res['name']} ({res['size_str']})")

    # 4. PDF conversion
    # We generated a PDF in step 1, let's use that as input!
    pdf_path = ce.text_to_pdf(sample_text, "pdf_input_test", output_dir=tmp)
    pdf_info = ce.inspect_file(pdf_path)
    res_pdf_txt = ce.execute_conversion(pdf_info, "txt", output_dir=tmp)
    assert res_pdf_txt["success"]
    print(f"  [PASS] PDF -> txt: {res_pdf_txt['name']}")

    # 5. Word conversion
    docx_path = ce.text_to_docx(sample_text, "docx_input_test", output_dir=tmp)
    docx_info = ce.inspect_file(docx_path)
    res_docx_pdf = ce.execute_conversion(docx_info, "pdf", output_dir=tmp)
    assert res_docx_pdf["success"]
    print(f"  [PASS] Word -> pdf: {res_docx_pdf['name']}")

    # 6. Spreadsheet conversion
    res_spread_pdf = ce.execute_conversion({"kind": "spreadsheet", "path": res_xls["output_path"], "name": "inventory"}, "pdf", output_dir=tmp)
    assert res_spread_pdf["success"]
    print(f"  [PASS] Spreadsheet -> pdf: {res_spread_pdf['name']}")

    print("\nALL 17 CONVERSION TESTS PASSED SUCCESSFULLY!")

if __name__ == "__main__":
    test_all()

