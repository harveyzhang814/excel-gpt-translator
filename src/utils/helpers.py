from pathlib import Path
import os

def ensure_directory(path: str) -> Path:
    """Ensure a directory exists, create if it doesn't."""
    path = Path(path)
    path.mkdir(parents=True, exist_ok=True)
    return path

def is_valid_excel_file(file_path: str) -> bool:
    """Check if the file is a valid Excel file."""
    if not os.path.exists(file_path):
        return False
    
    ext = os.path.splitext(file_path)[1].lower()
    return ext in ['.xlsx', '.xls']

def format_progress(completed: int, total: int) -> str:
    """Format progress as a percentage string."""
    if total == 0:
        return "0%"
    return f"{(completed / total) * 100:.1f}%" 