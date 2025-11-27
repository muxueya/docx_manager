# DOCX Manager

A small Flask web app to scan folders of Word documents (.docx), review links, check track changes, and run find/replace (single-file or bulk). Designed for non-technical users on Windows/macOS.

## Prerequisites
- Python 3.8+ installed
- Word documents accessible on your machine
- Tkinter (usually included with Python) for the folder picker; on macOS the app uses the native folder dialog.

## Setup
1) Open a terminal and go to the project folder:
   ```
   cd /path/to/docx_manager
   ```
2) (Optional) Create/activate a virtual environment.
3) Install requirements (pinned for consistent link-replace behavior):
   ```
   pip install -r requirements.txt
   ```
   To verify, run:
   ```
   pip show python-docx Flask openpyxl
   ```

## Run the app
Start the Flask server:
```
python app.py
```
You should see `Running on http://127.0.0.1:5000/`.

Open a browser to `http://127.0.0.1:5000/` (or `localhost:5000`). Keep the terminal window open while using the app.

## Using the web UI 
1) **Select Local Folder**
   - Click **Browse…** (opens a native folder picker) or type the full folder path that contains your `.docx` files.
   - Click **Scan Directory**. The left column will show the file tree (only `.docx` files are listed, temp/lock files are skipped).

2) **Bulk Tools (default tab)**
   - **Load Links (all files):** scans every `.docx` in the folder, listing all links found. Use **Export Links CSV** to download a spreadsheet of links.
   - **Dependencies:** shows how files reference each other (links to other files vs. other files linking here). Sort by outgoing/incoming and export the dependency table via **Export Dependencies CSV**.
   - **Bulk Find / Replace (text):** enter text to find across all files. Optionally enter replacement text and click **Replace All** (overwrites files) or **Find All** (no changes). Results list only files with matches, each filename links to its location (file://).
   - **Bulk Links Find / Replace:** targets hyperlinks. Enter the text/URL to find, optional replacement, and choose the target:
     - **Both:** look in link text and URL
     - **Link Text Only:** change visible link text
     - **Link URL Only:** change the actual hyperlink target
     Use **Find All** to report matches only, **Replace All** to overwrite files.

3) **Per-file analysis (Links & Info / Find & Replace tabs)**
   - In the left file tree, click a document. The per-file tabs unlock:
     - **Links & Info:** shows hyperlinks in the selected document and whether Track Changes is on.
     - **Find & Replace:** search (and optionally replace) within just that document. Shows snippets of matches.

## Tips and notes
- **Back up important documents** before running bulk replace. The app writes changes in place.
- **Saved copies:** if “Save copies to Desktop” is checked, originals are copied to `~/Desktop/bulk_found` (or `<selected-folder>/bulk_found` if Desktop is unavailable) before replacements are written.
- **Paths:** use absolute paths (e.g., `C:\Users\name\Docs` on Windows, `/Users/name/Docs` on macOS).
- **Permissions:** if a file can’t be read or written, you’ll see an error in the UI.
- **CSV exports:** open in Excel/Sheets; columns: File, Text, URL, Type, Error (links) or File, LinksToOtherFiles, OtherFilesLinkingHere (dependencies).


## Troubleshooting
- If `127.0.0.1` returns 403, make sure you include the port: `http://127.0.0.1:5000/`.
- If Tkinter is missing, install a Python build that includes it or rely on manual path entry.
- On macOS, the folder picker uses `osascript`; if it fails, enter the path manually.
- If the app doesn’t start, check the terminal for Python errors (missing packages, syntax issues). Reinstall requirements if needed.
- If bulk link replace doesn’t work on another machine, ensure `pip show python-docx` reports version `0.8.11` (matches `requirements.txt`); reinstall with `pip install -r requirements.txt` if not.

## Stopping the app
Press `Ctrl+C` in the terminal running `python app.py` to stop the server.
