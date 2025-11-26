import os
import sys
import subprocess
import sqlite3
import io
from flask import Flask, render_template, request, jsonify, send_file
from docx import Document

# Import our custom modules
from modules.file_scanner import scan_folder_structure, list_docx_files
from modules.docx_processor import (
    is_track_changes_on,
    get_links,
    process_find_replace,
    collect_links_for_files,
    process_find_replace_bulk,
    process_links_find_replace_bulk,
)

# XLSX export
try:
    from openpyxl import Workbook
except Exception:
    Workbook = None

# Optional import for native folder picker
try:
    import tkinter as tk
    from tkinter import filedialog
except Exception:
    tk = None

app = Flask(__name__)


def _to_rel_target(normalized, root_path):
    """Normalize a target to a rel path under root_path when possible."""
    if not normalized:
        return None
    val = str(normalized).replace('\\', '/')
    # Absolute path -> rel
    if os.path.isabs(val):
        try:
            rel = os.path.relpath(val, root_path)
            if not rel.startswith('..'):
                return rel.replace('\\', '/')
        except Exception:
            pass
    # Try joining relative to root
    try:
        candidate = os.path.normpath(os.path.join(root_path, val))
        rel = os.path.relpath(candidate, root_path)
        if not rel.startswith('..'):
            return rel.replace('\\', '/')
    except Exception:
        pass
    return val


def build_dependencies(root_path, files, link_data):
    """
    Build a simple dependency graph using an in-memory SQLite DB.
    Returns a list of {path, rel_path, outgoing_files, incoming_files, outgoing_details, incoming_details}
    """
    doc_map = {}
    id_to_doc = {}
    base_to_ids = {}

    # Build doc maps
    for idx, path in enumerate(files, start=1):
        rel = os.path.relpath(path, root_path).replace('\\', '/')
        base = os.path.splitext(os.path.basename(rel))[0].lower()
        info = {'id': idx, 'rel': rel, 'path': path, 'base': base}
        doc_map[path] = info
        id_to_doc[idx] = info
        base_to_ids.setdefault(base, set()).add(idx)

    # Aggregators
    outgoing_targets = {info['id']: set() for info in id_to_doc.values()}
    incoming_sources = {info['id']: set() for info in id_to_doc.values()}
    outgoing_details = {info['id']: [] for info in id_to_doc.values()}
    incoming_details = {info['id']: [] for info in id_to_doc.values()}

    # Helper to match a link to target doc IDs
    def match_targets(link_obj):
        matches = set()
        # normalized path
        normalized = link_obj.get('normalized') or link_obj.get('url')
        normalized_rel = _to_rel_target(normalized, root_path)
        if normalized_rel:
            for info in id_to_doc.values():
                if normalized_rel == info['rel']:
                    matches.add(info['id'])
        # text match to base name (case-insensitive)
        text = (link_obj.get('text') or '').strip().lower()
        if text and text in base_to_ids:
            matches.update(base_to_ids[text])
        return matches

    # Walk links per file
    for item in link_data:
        doc_info = doc_map.get(item.get('path'))
        if not doc_info:
            continue
        src_id = doc_info['id']
        for link in item.get('links', []):
            if link.get('type') not in ('internal', 'document'):
                continue
            targets = match_targets(link)
            # Record outgoing
            for tgt_id in targets:
                if tgt_id == src_id:
                    continue
                outgoing_targets[src_id].add(tgt_id)
                outgoing_details[src_id].append({
                    'text': link.get('text'),
                    'href': link.get('raw') or link.get('url'),
                    'target': id_to_doc[tgt_id]['rel']
                })
                incoming_sources[tgt_id].add(src_id)
                incoming_details[tgt_id].append({
                    'from': id_to_doc[src_id]['rel'],
                    'text': link.get('text'),
                    'href': link.get('raw') or link.get('url')
                })

    deps = []
    for info in id_to_doc.values():
        deps.append({
            'path': info['path'],
            'rel_path': info['rel'],
            'outgoing_files': len(outgoing_targets[info['id']]),
            'incoming_files': len(incoming_sources[info['id']]),
            'outgoing_details': outgoing_details[info['id']],
            'incoming_details': incoming_details[info['id']]
        })

    return deps

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/pick_folder', methods=['POST'])
def pick_folder():
    try:
        if sys.platform == 'darwin':
            # Use macOS native dialog via AppleScript to avoid Tk main-thread requirements
            script = 'POSIX path of (choose folder with prompt "Select a folder")'
            completed = subprocess.run(
                ['osascript', '-e', script],
                capture_output=True,
                text=True
            )
            if completed.returncode != 0:
                err_msg = completed.stderr.strip() or completed.stdout.strip() or 'No folder selected'
                return jsonify({'error': err_msg}), 400
            path = completed.stdout.strip()
            if not path:
                return jsonify({'error': 'No folder selected'}), 400
            return jsonify({'path': path})

        # Fallback to Tk for Windows/Linux where it works reliably inside Flask
        if tk is None:
            return jsonify({'error': 'Folder picker not available (tkinter missing or unsupported).'}), 500

        root = None
        try:
            root = tk.Tk()
            root.withdraw()
            root.attributes('-topmost', True)
            root.update()
            selected_path = filedialog.askdirectory(title='Select a folder')
        finally:
            if root is not None:
                root.destroy()

        if not selected_path:
            return jsonify({'error': 'No folder selected'}), 400

        return jsonify({'path': selected_path})

    except Exception as exc:
        return jsonify({'error': f'Folder picker failed: {exc}'}), 500

@app.route('/bulk_links', methods=['POST'])
def bulk_links():
    data = request.get_json(silent=True) or {}
    root_path = data.get('path')
    if not root_path or not os.path.exists(root_path):
        return jsonify({'error': 'Path does not exist'}), 400

    files = list_docx_files(root_path)
    link_data = collect_links_for_files(files, base_dir=root_path)
    total_links = sum(len(item.get('links', [])) for item in link_data)
    dependencies = build_dependencies(root_path, files, link_data)
    return jsonify({'files': link_data, 'total_links': total_links, 'dependencies': dependencies})


@app.route('/export_links_xlsx', methods=['POST'])
def export_links_xlsx():
    """Generate an .xlsx file in-memory and return it as a download.

    Accepts JSON POST with either:
      - { path: '<folder>' }  -> will collect links for all docx files under the folder
      - { rows: [[file,text,url,type,error], ...] } -> use given rows directly
    """
    if Workbook is None:
        return jsonify({'error': 'openpyxl not installed on server'}), 500

    data = request.get_json(silent=True) or {}
    rows = []

    if data.get('path'):
        root_path = data['path']
        if not os.path.exists(root_path):
            return jsonify({'error': 'Path does not exist'}), 400
        files = list_docx_files(root_path)
        link_data = collect_links_for_files(files, base_dir=root_path)
        for item in link_data:
            p = item.get('path')
            for link in item.get('links', []):
                rows.append([p, link.get('text'), link.get('url'), link.get('type'), ''])
            if item.get('error'):
                rows.append([p, '', '', '', item.get('error')])
    else:
        # Accept rows directly: list-of-lists
        rows = data.get('rows', []) or []

    # Build workbook
    wb = Workbook()
    ws = wb.active
    ws.title = 'Links'
    header = ['File', 'Text', 'URL', 'Type', 'Error']
    ws.append(header)
    for r in rows:
        # ensure the row has same length
        ws.append([r[i] if i < len(r) else '' for i in range(len(header))])

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)

    # send_file - try using modern parameter name then fallback
    try:
        return send_file(
            bio,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name='links.xlsx'
        )
    except TypeError:
        # Older Flask versions use attachment_filename
        return send_file(
            bio,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            attachment_filename='links.xlsx'
        )


@app.route('/bulk_find_replace', methods=['POST'])
def bulk_find_replace():
    data = request.get_json(silent=True) or {}
    root_path = data.get('path')
    find_text = data.get('find_text')
    replace_text = data.get('replace_text')
    save_copies = data.get('save_copies', True)

    if not root_path or not os.path.exists(root_path):
        return jsonify({'error': 'Path does not exist'}), 400
    if not find_text:
        return jsonify({'error': 'No find_text provided', 'total_matches': 0, 'files': []}), 400

    files = list_docx_files(root_path)
    # If requested, inform the processor of a save root so it can save copies before overwriting
    if save_copies:
        # Prefer saving to the user's Desktop if it exists, under a 'bulk_found' folder.
        try:
            desktop = os.path.join(os.path.expanduser('~'), 'Desktop')
            if os.path.exists(desktop):
                save_root = os.path.join(desktop, 'bulk_found')
            else:
                save_root = os.path.join(root_path, 'bulk_found')
        except Exception:
            save_root = os.path.join(root_path, 'bulk_found')

        # attach temporary metadata to the function for per-file path computation
        setattr(process_find_replace_bulk, '_save_root', save_root)
        setattr(process_find_replace_bulk, '_base_dir', root_path)
    else:
        setattr(process_find_replace_bulk, '_save_root', None)
        setattr(process_find_replace_bulk, '_base_dir', None)

    result = process_find_replace_bulk(files, find_text, replace_text)

    # If we saved copies, include the save_root in the response so the frontend can display it
    if save_copies:
        try:
            result['save_root'] = save_root
        except Exception:
            pass

    # cleanup temporary attributes
    try:
        delattr(process_find_replace_bulk, '_save_root')
        delattr(process_find_replace_bulk, '_base_dir')
    except Exception:
        pass
    return jsonify(result)


@app.route('/scan', methods=['POST'])
def scan():
    data = request.get_json(silent=True) or {}
    path = data.get('path')
    if not path or not os.path.exists(path):
        return jsonify({'error': 'Path does not exist'}), 400
    
    structure = scan_folder_structure(path)
    return jsonify({'structure': structure})


@app.route('/bulk_links_find_replace', methods=['POST'])
def bulk_links_find_replace():
    data = request.get_json(silent=True) or {}
    root_path = data.get('path')
    find_text = data.get('find_text')
    replace_text = data.get('replace_text')  # Can be None
    target = data.get('target', 'both')  # 'name' | 'url' | 'both'
    save_copies = data.get('save_copies', True)

    if not root_path or not os.path.exists(root_path):
        return jsonify({'error': 'Path does not exist'}), 400
    if not find_text:
        return jsonify({'error': 'No find_text provided', 'total_matches': 0, 'files': []}), 400

    files = list_docx_files(root_path)
    # If requested, inform the processor of a save root so it can save copies before overwriting
    if save_copies:
        try:
            desktop = os.path.join(os.path.expanduser('~'), 'Desktop')
            if os.path.exists(desktop):
                save_root = os.path.join(desktop, 'bulk_found')
            else:
                save_root = os.path.join(root_path, 'bulk_found')
        except Exception:
            save_root = os.path.join(root_path, 'bulk_found')

        setattr(process_links_find_replace_bulk, '_save_root', save_root)
        setattr(process_links_find_replace_bulk, '_base_dir', root_path)
    else:
        setattr(process_links_find_replace_bulk, '_save_root', None)
        setattr(process_links_find_replace_bulk, '_base_dir', None)

    # Avoid scanning saved-copy folder if it lives under the scanned root_path.
    # This prevents nested 'bulk_found/bulk_found/...' when running multiple searches.
    try:
        if save_copies and save_root:
            norm_save = os.path.normcase(os.path.abspath(save_root))
            norm_root = os.path.normcase(os.path.abspath(root_path))
            if norm_save.startswith(norm_root):
                # filter out any files that live under the save_root
                files = [f for f in files if not os.path.normcase(os.path.abspath(f)).startswith(norm_save + os.sep)]
    except Exception:
        pass

    result = process_links_find_replace_bulk(files, find_text, replace_text, target=target)

    if save_copies:
        try:
            result['save_root'] = save_root
        except Exception:
            pass

    try:
        delattr(process_links_find_replace_bulk, '_save_root')
        delattr(process_links_find_replace_bulk, '_base_dir')
    except Exception:
        pass
    return jsonify(result)

@app.route('/analyze_file', methods=['POST'])
def analyze_file():
    data = request.get_json(silent=True) or {}
    file_path = data.get('path')
    if not file_path or not os.path.exists(file_path):
        return jsonify({'error': 'File not found'}), 404

    try:
        doc = Document(file_path)
        
        # Check functionalities
        tracked = is_track_changes_on(doc)
        links = get_links(doc, doc_path=file_path, base_dir=os.path.dirname(file_path))
        
        return jsonify({
            'tracked_changes': tracked,
            'links': links,
            'path': file_path
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/find_replace', methods=['POST'])
def find_replace():
    data = request.get_json(silent=True) or {}
    file_path = data.get('path')
    find_text = data.get('find_text')
    replace_text = data.get('replace_text')  # Can be None

    # Basic validations
    if not file_path or not os.path.exists(file_path):
        return jsonify({'error': 'File not found'}), 404

    if not find_text:
        return jsonify({'matches': 0, 'snippets': [], 'status': 'No find text provided'})

    try:
        result = process_find_replace(file_path, find_text, replace_text)
        return jsonify(result)

    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', '5000'))
    app.run(debug=True, port=port)
