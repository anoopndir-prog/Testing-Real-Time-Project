"""Bundled-resource path resolution, kept dependency-free (no Tkinter, no
Flask) so both the desktop views and the web app can import it without
pulling in UI toolkit code they don't need.
"""

from __future__ import annotations

import sys
from pathlib import Path


def resource_path(relative_path: Path) -> Path:
    if hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS) / relative_path
    return Path(__file__).resolve().parents[1] / relative_path
