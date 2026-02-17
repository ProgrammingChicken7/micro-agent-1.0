import os
import shutil
from .base import get_workspace_path

def create_file(file_path, content=""):
    """Create a new file with content"""
    workspace = get_workspace_path()
    if not workspace:
        return {"error": "Workspace not configured"}

    full_path = os.path.join(workspace, file_path)
    if not full_path.startswith(workspace):
        return {"error": "File path must be within workspace"}

    try:
        os.makedirs(os.path.dirname(full_path), exist_ok=True)
        with open(full_path, 'w', encoding='utf-8') as f:
            f.write(content)
        return {"success": True, "file_path": full_path}
    except Exception as e:
        return {"error": str(e)}

def read_file(file_path):
    """Read content from a file"""
    workspace = get_workspace_path()
    if not workspace:
        return {"error": "Workspace not configured"}

    full_path = os.path.join(workspace, file_path)
    if not full_path.startswith(workspace):
        return {"error": "File path must be within workspace"}

    try:
        with open(full_path, 'r', encoding='utf-8') as f:
            content = f.read()
        return {"success": True, "content": content, "file_path": full_path}
    except Exception as e:
        return {"error": str(e)}

def update_file(file_path, content):
    """Update an existing file with new content"""
    workspace = get_workspace_path()
    if not workspace:
        return {"error": "Workspace not configured"}

    full_path = os.path.join(workspace, file_path)
    if not full_path.startswith(workspace):
        return {"error": "File path must be within workspace"}

    if not os.path.exists(full_path):
        return {"error": "File does not exist"}

    try:
        with open(full_path, 'w', encoding='utf-8') as f:
            f.write(content)
        return {"success": True, "file_path": full_path}
    except Exception as e:
        return {"error": str(e)}

def delete_file(file_path):
    """Delete a file"""
    workspace = get_workspace_path()
    if not workspace:
        return {"error": "Workspace not configured"}

    full_path = os.path.join(workspace, file_path)
    if not full_path.startswith(workspace):
        return {"error": "File path must be within workspace"}

    try:
        os.remove(full_path)
        return {"success": True, "file_path": full_path}
    except Exception as e:
        return {"error": str(e)}

def create_directory(dir_path):
    """Create a directory"""
    workspace = get_workspace_path()
    if not workspace:
        return {"error": "Workspace not configured"}

    full_path = os.path.join(workspace, dir_path)
    if not full_path.startswith(workspace):
        return {"error": "Directory path must be within workspace"}

    try:
        os.makedirs(full_path, exist_ok=True)
        return {"success": True, "dir_path": full_path}
    except Exception as e:
        return {"error": str(e)}

def list_directory(dir_path="."):
    """List directory contents"""
    workspace = get_workspace_path()
    if not workspace:
        return {"error": "Workspace not configured"}

    full_path = os.path.join(workspace, dir_path)
    if not full_path.startswith(workspace):
        return {"error": "Directory path must be within workspace"}

    try:
        items = []
        for item in os.listdir(full_path):
            item_path = os.path.join(full_path, item)
            items.append({
                "name": item,
                "path": os.path.join(dir_path, item),
                "is_directory": os.path.isdir(item_path),
                "size": os.path.getsize(item_path) if os.path.isfile(item_path) else 0
            })
        return {"success": True, "items": items, "dir_path": full_path}
    except Exception as e:
        return {"error": str(e)}

def rename_file(old_path, new_path):
    """Rename a file or directory"""
    workspace = get_workspace_path()
    if not workspace:
        return {"error": "Workspace not configured"}

    old_full_path = os.path.join(workspace, old_path)
    new_full_path = os.path.join(workspace, new_path)

    if not old_full_path.startswith(workspace) or not new_full_path.startswith(workspace):
        return {"error": "Paths must be within workspace"}

    try:
        os.rename(old_full_path, new_full_path)
        return {"success": True, "old_path": old_full_path, "new_path": new_full_path}
    except Exception as e:
        return {"error": str(e)}

def copy_file(src_path, dst_path):
    """Copy a file or directory"""
    workspace = get_workspace_path()
    if not workspace:
        return {"error": "Workspace not configured"}

    src_full_path = os.path.join(workspace, src_path)
    dst_full_path = os.path.join(workspace, dst_path)

    if not src_full_path.startswith(workspace) or not dst_full_path.startswith(workspace):
        return {"error": "Paths must be within workspace"}

    try:
        if os.path.isdir(src_full_path):
            shutil.copytree(src_full_path, dst_full_path)
        else:
            shutil.copy2(src_full_path, dst_full_path)
        return {"success": True, "src_path": src_full_path, "dst_path": dst_full_path}
    except Exception as e:
        return {"error": str(e)}

def move_file(src_path, dst_path):
    """Move a file or directory"""
    workspace = get_workspace_path()
    if not workspace:
        return {"error": "Workspace not configured"}

    src_full_path = os.path.join(workspace, src_path)
    dst_full_path = os.path.join(workspace, dst_path)

    if not src_full_path.startswith(workspace) or not dst_full_path.startswith(workspace):
        return {"error": "Paths must be within workspace"}

    try:
        shutil.move(src_full_path, dst_full_path)
        return {"success": True, "src_path": src_full_path, "dst_path": dst_full_path}
    except Exception as e:
        return {"error": str(e)}

def delete_directory(dir_path):
    """Delete a directory"""
    workspace = get_workspace_path()
    if not workspace:
        return {"error": "Workspace not configured"}

    full_path = os.path.join(workspace, dir_path)
    if not full_path.startswith(workspace):
        return {"error": "Directory path must be within workspace"}

    try:
        shutil.rmtree(full_path)
        return {"success": True, "dir_path": full_path}
    except Exception as e:
        return {"error": str(e)}
