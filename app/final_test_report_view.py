"""Final Test Report screen.

Lets the user pick the report date and attach the two source documents
(Project Specification, Test Inspection sheet) that will be merged into the
fixed Final Test Report PDF format. The merge/generation logic itself is
added in a later step; this screen only captures the inputs.
"""

from __future__ import annotations

import datetime as dt
from pathlib import Path
from typing import Callable
import tkinter as tk
from tkinter import filedialog, messagebox

from app.ui_common import (
    BG_DARKEST,
    BG_SHINY,
    BOX_BORDER,
    TEXT_MAIN,
    TEXT_SUB,
    RoundedButton,
    add_date_field,
)

try:
    from tkinterdnd2 import DND_FILES  # type: ignore
except Exception:  # pragma: no cover
    DND_FILES = None

PROJECT_SPEC_DEFAULT_MAIN = "Attach Project Specification here"
PROJECT_SPEC_DEFAULT_SUB = (
    "Drag & drop the Project Specification file here or click to browse"
    if DND_FILES is not None
    else "Click to browse the Project Specification file"
)
TEST_INSPECTION_DEFAULT_MAIN = "Attach Test Inspection sheet here"
TEST_INSPECTION_DEFAULT_SUB = (
    "Drag & drop the scanned Test Inspection sheet (PDF) here or click to browse"
    if DND_FILES is not None
    else "Click to browse the scanned Test Inspection sheet (PDF)"
)


class AttachSlot:
    """One pin-shaped attachment box used inside the Final Test Report screen."""

    def __init__(
        self,
        parent: tk.Widget,
        root: tk.Misc,
        default_main_text: str,
        default_sub_text: str,
        allowed_suffixes: set[str],
        invalid_message: str,
        width: int = 720,
        height: int = 150,
    ) -> None:
        self.root = root
        self.default_main_text = default_main_text
        self.default_sub_text = default_sub_text
        self.allowed_suffixes = allowed_suffixes
        self.invalid_message = invalid_message
        self.file_path: Path | None = None

        self.icon_var = tk.StringVar(value="📌")
        self.main_text_var = tk.StringVar(value=default_main_text)
        self.sub_text_var = tk.StringVar(value=default_sub_text)

        self.box = tk.Frame(
            parent,
            bg=BG_SHINY,
            width=width,
            height=height,
            highlightbackground=BOX_BORDER,
            highlightthickness=2,
            cursor="hand2",
        )
        self.box.pack_propagate(False)

        shine_line = tk.Frame(self.box, bg="#2f2f2f", height=2)
        shine_line.pack(fill="x", side="top")

        self.icon_label = tk.Label(
            self.box,
            textvariable=self.icon_var,
            bg=BG_SHINY,
            fg="#e6e6e6",
            font=("Segoe UI Emoji", 34),
        )
        self.icon_label.pack(pady=(16, 4))

        self.main_label = tk.Label(
            self.box,
            textvariable=self.main_text_var,
            bg=BG_SHINY,
            fg=TEXT_MAIN,
            font=("Segoe UI", 15, "bold"),
            wraplength=650,
            justify="center",
        )
        self.main_label.pack()

        self.sub_label = tk.Label(
            self.box,
            textvariable=self.sub_text_var,
            bg=BG_SHINY,
            fg=TEXT_SUB,
            font=("Segoe UI", 10),
            wraplength=650,
            justify="center",
        )
        self.sub_label.pack(pady=(8, 0))

        for widget in (self.box, self.icon_label, self.main_label, self.sub_label):
            widget.bind("<Button-1>", self._browse)

        self._setup_drag_drop()

    def _setup_drag_drop(self) -> None:
        if DND_FILES is None:
            return

        for widget in (self.box, self.icon_label, self.main_label, self.sub_label):
            try:
                widget.drop_target_register(DND_FILES)
                widget.dnd_bind("<<Drop>>", self._on_drop)
            except Exception:
                continue

    def _on_drop(self, event) -> None:
        try:
            dropped_items = self.root.tk.splitlist(event.data)
        except Exception:
            dropped_items = [event.data]

        if not dropped_items:
            return

        candidate = Path(str(dropped_items[0])).expanduser()
        if self._set_if_valid(candidate):
            return
        messagebox.showwarning("Invalid File", self.invalid_message)

    def _browse(self, _event=None) -> None:
        downloads = Path.home() / "Downloads"
        filetype_patterns = " ".join(
            f"*{suffix}" for suffix in sorted(self.allowed_suffixes))
        selected = filedialog.askopenfilename(
            title="Select File",
            initialdir=str(downloads) if downloads.exists() else None,
            filetypes=[("Allowed Files", filetype_patterns), ("All Files", "*.*")],
        )
        if not selected:
            return

        selected_path = Path(selected)
        if not self._set_if_valid(selected_path):
            messagebox.showwarning("Invalid File", self.invalid_message)

    def _set_if_valid(self, file_path: Path) -> bool:
        if not file_path.exists() or file_path.suffix.lower() not in self.allowed_suffixes:
            return False

        self.file_path = file_path
        self.icon_var.set("📎")
        self.main_text_var.set(file_path.name)
        self.sub_text_var.set("Attached successfully")
        return True

    def reset(self) -> None:
        self.file_path = None
        self.icon_var.set("📌")
        self.main_text_var.set(self.default_main_text)
        self.sub_text_var.set(self.default_sub_text)


class FinalTestReportView:
    """Builds the Final Test Report screen inside a given container frame."""

    def __init__(self, container: tk.Widget, root: tk.Misc, on_back: Callable[[], None]) -> None:
        self.container = container
        self.root = root
        self.on_back = on_back

        self.report_date_var = tk.StringVar(
            value=dt.date.today().strftime("%d/%m/%Y"))

        self._build_ui()

    def _build_ui(self) -> None:
        outer = self.container

        content = tk.Frame(outer, bg=BG_DARKEST)
        content.pack(fill="both", expand=True)

        center = tk.Frame(content, bg=BG_DARKEST)
        center.place(relx=0.5, rely=0.52, anchor="center")

        title = tk.Label(
            center,
            text="Final Test Report",
            bg=BG_DARKEST,
            fg="#e8e8e8",
            font=("Segoe UI", 22, "bold"),
        )
        title.pack(pady=(0, 18))

        date_row = tk.Frame(center, bg=BG_DARKEST)
        date_row.pack(pady=(0, 16))
        add_date_field(
            date_row, "Date", self.report_date_var, 0, self.root, allow_clear=False, row=0)

        self.project_spec_slot = AttachSlot(
            center,
            self.root,
            PROJECT_SPEC_DEFAULT_MAIN,
            PROJECT_SPEC_DEFAULT_SUB,
            allowed_suffixes={".pdf", ".docx"},
            invalid_message="Please attach a valid Project Specification file (.pdf or .docx).",
        )
        self.project_spec_slot.box.pack(pady=(0, 16))

        self.test_inspection_slot = AttachSlot(
            center,
            self.root,
            TEST_INSPECTION_DEFAULT_MAIN,
            TEST_INSPECTION_DEFAULT_SUB,
            allowed_suffixes={".pdf"},
            invalid_message="Please attach a valid scanned Test Inspection sheet (.pdf).",
        )
        self.test_inspection_slot.box.pack()

        button_area = tk.Frame(center, bg=BG_DARKEST)
        button_area.pack(pady=(28, 0))

        self.back_btn = RoundedButton(
            button_area, "Back", self.on_back,
            width=200, height=48,
        )
        self.back_btn.grid(row=0, column=0, padx=8)

        self.reset_btn = RoundedButton(
            button_area, "Reset", self._reset,
            width=200, height=48,
        )
        self.reset_btn.grid(row=0, column=1, padx=8)

        self.exit_btn = RoundedButton(
            button_area, "Exit", self.root.destroy,
            width=200, height=48,
        )
        self.exit_btn.grid(row=0, column=2, padx=8)

    def _reset(self) -> None:
        self.report_date_var.set(dt.date.today().strftime("%d/%m/%Y"))
        self.project_spec_slot.reset()
        self.test_inspection_slot.reset()
