# Drop Changer — Nook Notch Plugin

A universal drag-and-drop & clipboard converter built for the [Nook Notch](https://nooknotch.xyz).

Drop files, paste selected text from your web browser, or grab images directly from your clipboard to instantly convert them into Word (`.docx`), PDF (`.pdf`), Excel (`.xlsx`), JPEG, PNG, TXT, HTML, and more.

---

## ⚡ What's New & Fixed

### 1. Zero Overlap — Optimized for All 3 Tile Sizes
- **Small (1x1 square)**: Clean vertical layout featuring a centered Phosphor icon and status indicator. Clicking anywhere on the tile opens the full converter panel.
- **Wide (2x1 rectangle)**: Horizontal row with left icon, two-line text column (title + subtitle), and at most one compact right-aligned action button. Guaranteed zero icon-over-text clipping!
- **Large (2x2 square)**: Full-featured card layout taking complete advantage of the 2x2 grid with top toolbar, centered hero icon and details, and full-width bottom conversion buttons.

### 2. Seamless Universal Drag & Drop (Browser & System Level)
- **Web Browsers (Chrome, Edge, Firefox, Brave)**:
  - **Web Images**: Drag any image (JPEG, WebP, PNG, GIF) directly from a webpage or Google Images to the top notch bar. It is automatically captured, cached, and staged for conversion.
  - **Browser Text**: Select text on any web page or PDF viewer and drag it directly to the notch bar.
  - **Links & URLs**: Drag links from the browser address bar or webpage hyperlinks. Direct image/document URLs are downloaded automatically.
- **System Level (Windows File Explorer / Desktop)**:
  - Drag local files (PDFs, Word documents, spreadsheets, images, code files) directly onto the notch.
- **Top Notch Drop Zone**: A dedicated, semi-transparent frosted drop bar appears at the top center notch:
  `⇄ Drop anything here to convert (Images, Links, Text, Files)`
  Hovering over it turns it accent blue, releasing drops it right into Drop Changer!

### 3. File Naming & Renaming Before Conversion
- Before clicking Convert, type any custom filename directly into the **"Output File Name"** field.
- Quick presets:
  - **Original Name**: Keeps the source file's original stem name.
  - **+ Today's Date**: Appends today's date (e.g. `report_20261101`).
- Clean filenames: No unwanted timestamp clutter unless requested (e.g. `document.pdf` instead of `document_20261101_184807.pdf`).

### 4. Destination Folder Relocation
- Store converted files wherever you need.
- **One-click presets**:
  - `Downloads` (Default `%USERPROFILE%\Downloads\NookConverted\`)
  - `Desktop`
  - `Documents`
  - `Source Folder` (saves alongside the original dropped file)
- **Folder Browser Dialog**: Click **"Change"** to open the standard Windows Folder Browser and select any folder on your PC.
- Preferences persist across app launches.

---

## 📊 Conversion Capabilities

| Input Type | Target Formats | Features & Styling |
| :--- | :--- | :--- |
| **Text / Browser Selection / Markdown** | **Word (`.docx`)**<br>**PDF (`.pdf`)**<br>**Excel (`.xlsx`)**<br>**TXT (`.txt`)**<br>**Image Card (`.png`)**<br>**HTML (`.html`)** | • Styled Word headings, lists & typography<br>• Professional PDF layout with margins<br>• Smart Excel table parser (CSV, TSV, JSON, tables)<br>• Notion/developer-style dark social card image<br>• Responsive HTML page with modern CSS |
| **Images** (PNG, JPG, WEBP, BMP, etc.) | **JPEG (`.jpg`)**<br>**PNG (`.png`)**<br>**WEBP (`.webp`)**<br>**PDF (`.pdf`)**<br>**Word (`.docx`)**<br>**Icon (`.ico`)** | • Quality-optimized JPEG with RGB composite<br>• Lossless alpha-preserving PNG<br>• High-efficiency WEBP<br>• Page-centered, proportional Image to PDF<br>• Word document with embedded graphic & caption<br>• Multi-size Windows application icons |
| **Spreadsheets** (`.xlsx`, `.csv`, `.tsv`) | **Excel (`.xlsx`)**<br>**PDF Report (`.pdf`)**<br>**Word Table (`.docx`)**<br>**CSV (`.csv`)**<br>**TXT (`.txt`)**<br>**HTML Table (`.html`)** | • Visual PDF table report with alternating colored rows<br>• Word table styled with headers<br>• Quick format cross-conversion |
| **PDF Documents** (`.pdf`) | **TXT (`.txt`)**<br>**Word (`.docx`)** | • Extracts all text page by page<br>• Editable Word document creation |
| **Word Documents** (`.docx`) | **PDF (`.pdf`)**<br>**TXT (`.txt`)**<br>**HTML (`.html`)** | • Clean PDF export<br>• Extracted text & HTML formatting |

---

## 🚀 How to Use

1. **Drop Any File**: Drag any file from File Explorer or Desktop onto the Nook notch.
2. **Or Paste Clipboard**: Copy text or an image in Chrome, Edge, or Snipping Tool, then click **"Paste Clipboard"** on the tile or press `Ctrl+Alt+D`.
3. **Pick Format & Rename**:
   - Choose your format chip (`Word`, `PDF`, `Excel`, `JPEG`, etc.).
   - (Optional) Edit the filename in the **"Output File Name"** box.
   - (Optional) Choose or change the save folder.
4. **Convert**: Click **"Convert Now"**.
5. **Open**: Click **"Open"** to launch the converted file immediately in your default Windows viewer.
