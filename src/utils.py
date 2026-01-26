"""
Utility Functions
Helper functions for file handling, validation, and formatting
"""

import os
from pathlib import Path
from typing import Optional, Tuple
from datetime import datetime


def validate_excel_file(file_path: str) -> Tuple[bool, str]:
    """
    Validate that the file is a valid Excel file
    Returns: (is_valid, message)
    """
    if not file_path:
        return False, "File path is empty"

    if not os.path.exists(file_path):
        return False, f"File not found: {file_path}"

    # Check file extension
    valid_extensions = {'.xlsx', '.xls', '.xlsm'}
    file_ext = Path(file_path).suffix.lower()

    if file_ext not in valid_extensions:
        return False, f"Invalid file format. Supported: {', '.join(valid_extensions)}"

    # Check if file is readable
    if not os.access(file_path, os.R_OK):
        return False, "File is not readable"

    return True, "File is valid"


def get_output_path(input_file: str, output_dir: Optional[str] = None) -> str:
    """
    Generate output markdown file path from input Excel file
    """
    if output_dir is None:
        output_dir = "./output"

    # Create output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)

    # Get base filename without extension
    base_name = Path(input_file).stem
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    # Create output filename
    output_filename = f"{base_name}_documentation_{timestamp}.md"
    output_path = os.path.join(output_dir, output_filename)

    return output_path


def format_file_size(size_bytes: int) -> str:
    """Format byte size to human readable format"""
    for unit in ['B', 'KB', 'MB', 'GB']:
        if size_bytes < 1024:
            return f"{size_bytes:.2f} {unit}"
        size_bytes /= 1024
    return f"{size_bytes:.2f} TB"


def truncate_string(s: str, max_length: int = 50, suffix: str = "...") -> str:
    """Truncate string to max length"""
    if len(s) <= max_length:
        return s
    return s[:max_length - len(suffix)] + suffix


def clean_column_name(name: str) -> str:
    """Clean and normalize column names"""
    # Remove special characters except spaces, underscores, hyphens
    cleaned = "".join(c if c.isalnum() or c in " _-" else "" for c in name)
    # Remove extra spaces
    cleaned = " ".join(cleaned.split())
    return cleaned


def sanitize_markdown(text: str) -> str:
    """Sanitize text for markdown output"""
    # Escape special markdown characters
    special_chars = ['\\', '`', '*', '_', '{', '}', '[', ']', '(', ')', '#', '+', '-', '.', '!', '|']
    for char in special_chars:
        text = text.replace(char, f"\\{char}")
    return text


def create_markdown_table(headers: list, rows: list) -> str:
    """
    Create a markdown table from headers and rows
    """
    if not headers or not rows:
        return ""

    # Create header row
    header_row = "| " + " | ".join(str(h) for h in headers) + " |\n"

    # Create separator row
    separator_row = "|" + "|".join(["---"] * len(headers)) + "|\n"

    # Create data rows
    data_rows = ""
    for row in rows:
        data_rows += "| " + " | ".join(str(cell) for cell in row) + " |\n"

    return header_row + separator_row + data_rows


def extract_column_names_from_file(file_path: str) -> list:
    """
    Extract column names from Excel file
    """
    import pandas as pd

    try:
        df = pd.read_excel(file_path, nrows=0)
        return list(df.columns)
    except Exception as e:
        raise ValueError(f"Error reading Excel file: {str(e)}")


class ProgressTracker:
    """Track progress of long-running operations"""

    def __init__(self, total_steps: int):
        self.total_steps = total_steps
        self.current_step = 0
        self.messages = []

    def update(self, message: str = ""):
        """Update progress"""
        self.current_step += 1
        if message:
            self.messages.append(message)

    def get_progress(self) -> float:
        """Get progress percentage"""
        if self.total_steps == 0:
            return 0
        return (self.current_step / self.total_steps) * 100

    def get_status(self) -> dict:
        """Get current status"""
        return {
            "current_step": self.current_step,
            "total_steps": self.total_steps,
            "progress_percent": self.get_progress(),
            "messages": self.messages,
        }
