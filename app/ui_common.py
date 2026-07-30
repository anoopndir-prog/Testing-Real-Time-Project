"""Shared UI building blocks used across the app's screens.

Kept separate so the main menu, Project Specification screen, and Final Test
Report screen can all reuse the same look (colors, rounded buttons, date
picker) without duplicating widget code.
"""

from __future__ import annotations

import calendar as calendar_lib
import datetime as dt
import tkinter as tk
from tkinter import ttk

from app.paths import resource_path

try:
    from tkcalendar import Calendar  # type: ignore
except Exception:  # pragma: no cover
    Calendar = None

APP_TITLE = "Global Testing India"

BG_DARKEST = "#050505"
BG_SHINY = "#0c0c0c"
BOX_BORDER = "#2a2a2a"
TEXT_MAIN = "#d9d9d9"
TEXT_SUB = "#a3a3a3"
ACCENT_GREEN = "#2eea6f"
ACCENT_GREEN_HOVER = "#48f17f"
BUTTON_TEXT = "#020202"
ENTRY_BG = "#101010"
ENTRY_FG = "#f0f0f0"
ENTRY_BORDER = "#2a2a2a"


class RoundedButton(tk.Canvas):
    """Rounded button drawn on canvas (for curved look in Tkinter)."""

    def __init__(
        self,
        master,
        text: str,
        command,
        width: int = 420,
        height: int = 54,
        radius: int = 22,
        bg_color: str = ACCENT_GREEN,
        hover_color: str = ACCENT_GREEN_HOVER,
        text_color: str = BUTTON_TEXT,
    ) -> None:
        super().__init__(
            master,
            width=width,
            height=height,
            bg=master.cget("bg"),
            highlightthickness=0,
            bd=0,
            cursor="hand2",
        )
        self._text = text
        self._command = command
        self._width = width
        self._height = height
        self._radius = radius
        self._bg_color = bg_color
        self._hover_color = hover_color
        self._text_color = text_color
        self._current_color = bg_color

        self._draw()
        self.bind("<Button-1>", self._on_click)
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)

    def _rounded_rect(self, x1: int, y1: int, x2: int, y2: int, radius: int, fill: str) -> None:
        points = [
            x1 + radius, y1,
            x2 - radius, y1,
            x2, y1,
            x2, y1 + radius,
            x2, y2 - radius,
            x2, y2,
            x2 - radius, y2,
            x1 + radius, y2,
            x1, y2,
            x1, y2 - radius,
            x1, y1 + radius,
            x1, y1,
        ]
        self.create_polygon(points, smooth=True, fill=fill, outline=fill)

    def _draw(self) -> None:
        self.delete("all")
        self._rounded_rect(2, 2, self._width - 2, self._height -
                           2, self._radius, self._current_color)
        self.create_text(
            self._width // 2,
            self._height // 2,
            text=self._text,
            fill=self._text_color,
            font=("Segoe UI", 12, "bold"),
        )

    def _on_click(self, _event) -> None:
        self._command()

    def _on_enter(self, _event) -> None:
        self._current_color = self._hover_color
        self._draw()

    def _on_leave(self, _event) -> None:
        self._current_color = self._bg_color
        self._draw()


def validate_digits(proposed: str) -> bool:
    return proposed == "" or proposed.isdigit()


def add_text_field(
    parent: tk.Widget,
    label: str,
    var: tk.StringVar,
    column: int,
    root: tk.Misc,
    width: int = 14,
    numeric_only: bool = False,
    row: int = 0,
) -> tk.Entry:
    field = tk.Frame(parent, bg=BG_DARKEST)
    field.grid(row=row, column=column, padx=8, pady=(
        0, 8) if row == 0 else (0, 0), sticky="w")

    tk.Label(
        field,
        text=label,
        bg=BG_DARKEST,
        fg="#d7d7d7",
        font=("Segoe UI", 10, "bold"),
    ).pack(anchor="w", pady=(0, 5))

    entry = tk.Entry(
        field,
        textvariable=var,
        width=width,
        bg="#101010",
        fg="#f0f0f0",
        insertbackground="#f0f0f0",
        relief="flat",
        highlightthickness=1,
        highlightbackground="#2a2a2a",
        highlightcolor="#2a2a2a",
        font=("Segoe UI", 10),
    )
    if numeric_only:
        validator = root.register(validate_digits)
        entry.configure(validate="key", validatecommand=(validator, "%P"))
    entry.pack(anchor="w", ipady=4)
    return entry


def add_dropdown_field(
    parent: tk.Widget,
    label: str,
    var: tk.StringVar,
    column: int,
    values: list[str],
    width: int = 14,
    row: int = 0,
) -> ttk.Combobox:
    field = tk.Frame(parent, bg=BG_DARKEST)
    field.grid(row=row, column=column, padx=8, pady=(
        0, 8) if row == 0 else (0, 0), sticky="w")

    tk.Label(
        field,
        text=label,
        bg=BG_DARKEST,
        fg="#d7d7d7",
        font=("Segoe UI", 10, "bold"),
    ).pack(anchor="w", pady=(0, 5))

    combo = ttk.Combobox(field, textvariable=var,
                         values=values, width=width, state="readonly")
    combo.pack(anchor="w", ipady=2)
    return combo


def add_date_field(
    parent: tk.Widget,
    label: str,
    var: tk.StringVar,
    column: int,
    root: tk.Misc,
    allow_clear: bool,
    row: int = 0,
) -> tk.Entry:
    field = tk.Frame(parent, bg=BG_DARKEST)
    field.grid(row=row, column=column, padx=8, pady=(
        0, 8) if row == 0 else (0, 0), sticky="w")

    tk.Label(
        field,
        text=label,
        bg=BG_DARKEST,
        fg="#d7d7d7",
        font=("Segoe UI", 10, "bold"),
    ).pack(anchor="w", pady=(0, 5))

    date_entry = tk.Entry(
        field,
        textvariable=var,
        width=12,
        state="readonly",
        readonlybackground="#101010",
        fg="#f0f0f0",
        relief="flat",
        highlightthickness=1,
        highlightbackground="#2a2a2a",
        highlightcolor="#2a2a2a",
        font=("Segoe UI", 10),
        cursor="hand2",
    )
    date_entry.pack(anchor="w", ipady=4)
    date_entry.bind(
        "<Button-1>", lambda _e: open_calendar(root, var, allow_clear=allow_clear))
    return date_entry


def open_calendar(root: tk.Misc, target_var: tk.StringVar, allow_clear: bool) -> None:
    if Calendar is None:
        open_native_calendar(root, target_var, allow_clear)
        return

    popup = tk.Toplevel(root)
    popup.title("Select Date")
    popup.configure(bg="#111111")
    popup.resizable(False, False)
    popup.transient(root)
    popup.grab_set()

    selected = target_var.get().strip()
    today = dt.date.today()
    try:
        current_date = dt.datetime.strptime(
            selected, "%d/%m/%Y").date() if selected else today
    except ValueError:
        current_date = today

    cal = Calendar(
        popup,
        selectmode="day",
        date_pattern="dd/mm/yyyy",
        year=current_date.year,
        month=current_date.month,
        day=current_date.day,
        background="#1a1a1a",
        foreground="#f0f0f0",
        headersbackground="#161616",
        headersforeground="#f0f0f0",
        normalbackground="#101010",
        normalforeground="#f0f0f0",
        weekendbackground="#101010",
        weekendforeground="#f0f0f0",
        selectbackground="#2eea6f",
        selectforeground="#050505",
    )
    cal.pack(padx=10, pady=10)

    btn_row = tk.Frame(popup, bg="#111111")
    btn_row.pack(fill="x", padx=10, pady=(0, 10))

    def apply_date() -> None:
        target_var.set(cal.get_date())
        popup.destroy()

    tk.Button(
        btn_row,
        text="Select",
        command=apply_date,
        bg=ACCENT_GREEN,
        fg=BUTTON_TEXT,
        relief="flat",
        font=("Segoe UI", 10, "bold"),
        padx=12,
        pady=4,
        cursor="hand2",
    ).pack(side="left")

    if allow_clear:
        tk.Button(
            btn_row,
            text="Clear",
            command=lambda: (target_var.set(""), popup.destroy()),
            bg="#1b1b1b",
            fg="#eaeaea",
            relief="flat",
            font=("Segoe UI", 10),
            padx=12,
            pady=4,
            cursor="hand2",
        ).pack(side="left", padx=(8, 0))


def open_native_calendar(root: tk.Misc, target_var: tk.StringVar, allow_clear: bool) -> None:
    popup = tk.Toplevel(root)
    popup.title("Select Date")
    popup.configure(bg="#111111")
    popup.resizable(False, False)
    popup.transient(root)
    popup.grab_set()

    selected = target_var.get().strip()
    today = dt.date.today()
    try:
        current_date = dt.datetime.strptime(
            selected, "%d/%m/%Y").date() if selected else today
    except ValueError:
        current_date = today

    shown_year = current_date.year
    shown_month = current_date.month

    header = tk.Frame(popup, bg="#111111")
    header.pack(fill="x", padx=10, pady=(10, 6))

    title_var = tk.StringVar()

    def change_month(delta: int) -> None:
        nonlocal shown_year, shown_month
        month_index = shown_month - 1 + delta
        shown_year += month_index // 12
        shown_month = month_index % 12 + 1
        draw_days()

    tk.Button(
        header,
        text="<",
        command=lambda: change_month(-1),
        bg="#1b1b1b",
        fg="#eaeaea",
        relief="flat",
        width=3,
        cursor="hand2",
    ).pack(side="left")

    tk.Label(
        header,
        textvariable=title_var,
        bg="#111111",
        fg="#f0f0f0",
        font=("Segoe UI", 10, "bold"),
        width=18,
    ).pack(side="left", padx=8)

    tk.Button(
        header,
        text=">",
        command=lambda: change_month(1),
        bg="#1b1b1b",
        fg="#eaeaea",
        relief="flat",
        width=3,
        cursor="hand2",
    ).pack(side="left")

    day_frame = tk.Frame(popup, bg="#111111")
    day_frame.pack(padx=10, pady=(0, 10))

    for column, day_name in enumerate(("Mo", "Tu", "We", "Th", "Fr", "Sa", "Su")):
        tk.Label(
            day_frame,
            text=day_name,
            bg="#111111",
            fg="#a3a3a3",
            font=("Segoe UI", 9, "bold"),
            width=4,
        ).grid(row=0, column=column, padx=1, pady=1)

    def choose_date(day: int) -> None:
        target_var.set(dt.date(shown_year, shown_month,
                       day).strftime("%d/%m/%Y"))
        popup.destroy()

    def draw_days() -> None:
        title_var.set(f"{calendar_lib.month_name[shown_month]} {shown_year}")
        for child in day_frame.grid_slaves():
            row = int(child.grid_info().get("row", 0))
            if row > 0:
                child.destroy()

        month_days = calendar_lib.monthcalendar(shown_year, shown_month)
        for row_index, week in enumerate(month_days, start=1):
            for column, day in enumerate(week):
                if day == 0:
                    tk.Label(
                        day_frame,
                        text="",
                        bg="#111111",
                        width=4,
                    ).grid(row=row_index, column=column, padx=1, pady=1)
                    continue

                is_selected = (
                    shown_year == current_date.year
                    and shown_month == current_date.month
                    and day == current_date.day
                )
                tk.Button(
                    day_frame,
                    text=str(day),
                    command=lambda selected_day=day: choose_date(selected_day),
                    bg=ACCENT_GREEN if is_selected else "#101010",
                    fg=BUTTON_TEXT if is_selected else "#f0f0f0",
                    activebackground=ACCENT_GREEN_HOVER,
                    activeforeground=BUTTON_TEXT,
                    relief="flat",
                    width=4,
                    cursor="hand2",
                ).grid(row=row_index, column=column, padx=1, pady=1)

    draw_days()

    if allow_clear:
        tk.Button(
            popup,
            text="Clear",
            command=lambda: (target_var.set(""), popup.destroy()),
            bg="#1b1b1b",
            fg="#eaeaea",
            relief="flat",
            font=("Segoe UI", 10),
            padx=12,
            pady=4,
            cursor="hand2",
        ).pack(anchor="w", padx=10, pady=(0, 10))
