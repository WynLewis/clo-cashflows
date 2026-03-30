"""
File picker utility for selecting Excel files in the notebook.

Uses tkinter's native file dialog on Windows/Mac/Linux.
"""

from __future__ import annotations

from pathlib import Path
from typing import Optional


def pick_file(
    title: str = "Select a file",
    filetypes: list[tuple[str, str]] | None = None,
    initial_dir: str | Path | None = None,
) -> Optional[Path]:
    """Open a native file browser dialog and return the selected path.

    Args:
        title: Dialog window title.
        filetypes: List of (label, pattern) tuples, e.g. [("Excel files", "*.xlsx")].
        initial_dir: Starting directory for the dialog.

    Returns:
        Path to the selected file, or None if the user cancelled.
    """
    import tkinter as tk
    from tkinter import filedialog

    if filetypes is None:
        filetypes = [
            ("Excel files", "*.xlsx *.xls *.xlsm"),
            ("All files", "*.*"),
        ]

    # Create a hidden root window (required by tkinter).
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)  # Bring dialog to front.

    kwargs = {"title": title, "filetypes": filetypes}
    if initial_dir:
        kwargs["initialdir"] = str(initial_dir)

    filepath = filedialog.askopenfilename(**kwargs)

    root.destroy()

    if filepath:
        result = Path(filepath)
        print(f"  Selected: {result}")
        return result
    else:
        print("  No file selected.")
        return None


def pick_excel_file(
    title: str = "Select an Excel file",
    initial_dir: str | Path | None = None,
) -> Optional[Path]:
    """Shortcut for picking an Excel file."""
    return pick_file(
        title=title,
        filetypes=[("Excel files", "*.xlsx *.xls *.xlsm"), ("All files", "*.*")],
        initial_dir=initial_dir,
    )
