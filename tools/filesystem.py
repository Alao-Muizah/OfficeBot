import os 
import shutil 
from datetime import datetime 
from send2trash import send2trash

def _resolve_path(path:str, workdir:str) -> str:

    if os.path.isabs(path):
        return path 
    return os.path.join(workdir, path)

# ====================== Fetch file ====================== #

def fetch_file(path:str, workdir:str) -> dict:

    full_path = _resolve_path(path, workdir)

    if not os.path.exists(full_path):
        return {
            "status": "error",
            "message": f"File not found: {full_path}"
        }
    
    if not os.path.isfile(full_path):
        return {
                    "status": "ok",
                    "message": f"'{path}' is a directory not a file"
                }
    
    size_bytes = os.path.getsize(full_path)

    if size_bytes < 1024:
        size_str = f"{size_bytes}B"
    elif size_bytes < 1024 ** 2 :
        size_str = f"{size_bytes / 1024:.1f} KB"
    else:
        size_str = f"{size_bytes / (1024 ** 2):.1f} MB"

    modified_time = datetime.fromtimestamp(os.path.getmtime(full_path))
    modified_str = modified_time.strftime("%Y-%m-%d %H:%m")

    extension = os.path.splitext(full_path)[1].lower()

    return {
        "status": "error",
        "name": os.path.basename(full_path),
        "full_path": full_path,
        "extension": extension,
        "size": size_str,
        "last_modified": modified_str
    }


# ====================== List files ====================== #

def list_files(directory:str, workdir:str) -> dict:
    
    if not directory or directory.strip() in (".", "", "workdir"):
        target_dir = workdir
    else:
        target_dir = _resolve_path(directory, workdir)

    if not os.path.exists(target_dir):
        return {"status": "error", 
                "message": f"Directory not found: {target_dir}"}
    
    if not os.path.isdir(target_dir):
        return {"status": "error", 
                "message": f"'{directory}' is a file, not a directory"}
    
    allowed_extension = {".docx", ".xlsx", ".csv", ".pdf"}

    files = []

    for name in os.listdir(target_dir):
        full_path = os.path.join(target_dir, name)

        if name.startswith(".") or not os.path.isfile(full_path):
            continue 

        extension = os.path.splitext(name)[1].lower()
        # if extension not in allowed_extension:
        #     continue
        

        size_byte = os.path.getsize(full_path)

        # size_byte = os.path.getsize(full_path)
        size_str = ""

        if size_byte < 1024:
            size_str = f"{size_byte} B"
        elif size_byte < (1024 ** 2):
            size_str = f"{size_byte / 1024:.1f} KB"
        else:
            size_str = f"{size_byte / (1024 ** 2):.1f} MB"

        files.append({
            "name": name,
            "extension": extension,
            "size": size_str
        })

    # For alphabetical sorting
    files.sort(key=lambda f: f['name'])

    return {
        "status": "ok",
        "directory": target_dir,
        "file_count": len(files),
        "files": files 
    }
    

# ====================== Copying file ====================== #

def copy_file(source_path:str, destination_path:str, workdir:str) -> dict:

    full_source = _resolve_path(source_path, workdir)
    full_dest = _resolve_path(destination_path, workdir)

    if not os.path.exists(full_source):
        return {
            "status": "error",
            "message": f"Source file not found '{full_source}'"
        }
    
    dest_ext = os.path.splitext(full_dest)[1]

    if not dest_ext:
        os.makedirs(full_dest, exist_ok=True)
    else:
        os.makedirs(os.path.dirname(full_dest), exist_ok=True)

    shutil.copy2(full_source, full_dest)

    if os.path.isdir(full_dest):
        final_path = os.path.join(full_dest, os.path.basename(full_source))
    else:
        final_path = full_dest
    
    return {
        "status": "ok",
        "message": f"Copied '{source_path}' to '{destination_path}'",
        "destination": final_path
    }


# ====================== Delete files ====================== #

def delete_file(path: str, workdir: str) -> dict:
    full_path = path if os.path.isabs(path) else os.path.join(workdir, path)

    if not os.path.exists(full_path):
        return {"status": "error", "message": f"File not found: {full_path}"}

    try:
        send2trash(full_path)
        return {"status": "ok", "message": f"Moved {full_path} to Recycle Bin"}
    except Exception as e:
        return {"status": "error", "message": f"Could not delete {full_path}: {e}"}


# ====================== Renaming files ====================== #

def rename_file(path: str, new_name: str, workdir: str) -> dict:

    # if absolute path use it directly, else join with workdir
    if os.path.isabs(path):
        full_path = path
    else:
        full_path = os.path.join(workdir, path)
    if os.path.isabs(path):
        full_path = path
    else:
        full_path = os.path.join(workdir, path)

    if not os.path.exists(full_path):
        return {"status": "error", "message": f"File not found: {full_path}"}

    if not os.path.isfile(full_path):
        return {"status": "error", "message": f"'{full_path}' is not a file"}

    new_path = os.path.join(os.path.dirname(full_path), new_name)

    if os.path.exists(new_path):
        return {"status": "error", "message": f"A file named '{new_name}' already exists"}

    os.rename(full_path, new_path)

    return {
        "status": "ok",
        "message": f"Renamed '{os.path.basename(full_path)}' to '{new_name}'",
        "path": new_path
    }
