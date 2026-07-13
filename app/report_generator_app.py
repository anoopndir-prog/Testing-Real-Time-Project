#!/usr/bin/env python3
"""Global Testing India desktop app.

Main menu offers two tools:
- Project Specification: attach the Excel request sheet and generate the
  Word/PDF Test Project Specification report (pre-existing tool, unchanged).
- Final Test Report: attach the Project Specification and the scanned Test
  Inspection sheet to build the fixed-format Final Test Report.
"""

from __future__ import annotations
from app.final_test_report_view import FinalTestReportView
from app.project_specification_view import ProjectSpecificationView
from app.ui_common import APP_TITLE, BG_DARKEST, RoundedButton

import sys
from pathlib import Path
import tkinter as tk

# Allow running as `python3 app/report_generator_app.py` from project root.
PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))


try:
    from tkinterdnd2 import TkinterDnD  # type: ignore
except Exception:  # pragma: no cover
    TkinterDnD = None


class AppShell:
    """Owns the root window and switches between the main menu and screens."""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("980x680")
        self.root.minsize(900, 620)
        self.root.configure(bg=BG_DARKEST)

        self.current_frame: tk.Frame | None = None
        self.show_main_menu()

    def _new_container(self) -> tk.Frame:
        if self.current_frame is not None:
            self.current_frame.destroy()

        frame = tk.Frame(self.root, bg=BG_DARKEST)
        frame.pack(fill="both", expand=True)
        self.current_frame = frame
        return frame

    def show_main_menu(self) -> None:
        container = self._new_container()

        center = tk.Frame(container, bg=BG_DARKEST)
        center.place(relx=0.5, rely=0.5, anchor="center")

        tk.Label(
            center,
            text="Global Testing India",
            bg=BG_DARKEST,
            fg="#e8e8e8",
            font=("Segoe UI", 30, "bold"),
        ).pack(pady=(0, 6))

        tk.Label(
            center,
            text="Select a tool to continue",
            bg=BG_DARKEST,
            fg="#a3a3a3",
            font=("Segoe UI", 12),
        ).pack(pady=(0, 40))

        button_area = tk.Frame(center, bg=BG_DARKEST)
        button_area.pack()

        RoundedButton(
            button_area,
            "Project Specification",
            self.show_project_specification,
            width=420,
            height=64,
        ).pack(pady=(0, 18))

        RoundedButton(
            button_area,
            "Final Test Report",
            self.show_final_test_report,
            width=420,
            height=64,
        ).pack(pady=(0, 18))

        RoundedButton(
            button_area,
            "Exit",
            self.root.destroy,
            width=420,
            height=64,
        ).pack()

    def show_project_specification(self) -> None:
        container = self._new_container()
        ProjectSpecificationView(
            container, self.root, on_back=self.show_main_menu)

    def show_final_test_report(self) -> None:
        container = self._new_container()
        FinalTestReportView(
            container, self.root, on_back=self.show_main_menu)


def main() -> None:
    root = TkinterDnD.Tk() if TkinterDnD is not None else tk.Tk()
    AppShell(root)
    root.mainloop()


if __name__ == "__main__":
    main()
