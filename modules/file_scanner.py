import os

def scan_folder_structure(root_path):
    """
    Recursive function to build a file tree.
    Returns a dictionary structure compatible with the frontend.
    """
    tree = {
        'name': os.path.basename(root_path), 
        'type': 'folder', 
        'path': root_path, 
        'children': []
    }
    try:
        # scandir is faster than listdir for large directories
        with os.scandir(root_path) as entries:
            for entry in entries:
                if entry.is_dir():
                    # Recursively scan subfolders
                    tree['children'].append(scan_folder_structure(entry.path))
                elif entry.is_file() and entry.name.endswith('.docx') and not entry.name.startswith('~$'):
                    # Only add .docx files that are not temporary/lock files (~$)
                    tree['children'].append({
                        'name': entry.name, 
                        'type': 'file', 
                        'path': entry.path
                    })
    except PermissionError:
        pass # Skip folders the system prevents us from reading
    return tree


def list_docx_files(root_path):
    """
    Returns a flat list of .docx file paths under root_path (skips temp files).
    """
    docx_files = []
    try:
        with os.scandir(root_path) as entries:
            for entry in entries:
                if entry.is_dir():
                    docx_files.extend(list_docx_files(entry.path))
                elif entry.is_file() and entry.name.endswith('.docx') and not entry.name.startswith('~$'):
                    docx_files.append(entry.path)
    except PermissionError:
        pass
    return docx_files
