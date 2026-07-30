"""Modal Tkinter dialog: shows a seal photo with an auto-suggested crop
box (see app/seal_photo.py) that the user can drag into place before it's
accepted. Used by the Final Test Report screen's seal-photo attach slot.
"""

from __future__ import annotations

import tkinter as tk
from tkinter import messagebox

from PIL import Image, ImageTk

from app.seal_photo import crop_to_box, encode_jpeg, normalize_for_report, suggest_crop_box

CANVAS_MAX_W = 720
CANVAS_MAX_H = 480
HANDLE_HALF = 6
HANDLE_HIT_RADIUS = 10


class SealPhotoCropDialog:
    """Blocks until the user confirms or cancels. `result_jpeg_bytes` holds
    the final cropped photo (JPEG bytes), or None if cancelled."""

    def __init__(self, root: tk.Misc, image: Image.Image) -> None:
        self.result_jpeg_bytes: bytes | None = None
        self.image = normalize_for_report(image)
        self._build_scaled_preview()
        self.box = list(suggest_crop_box(self.image))
        self._active_handle: str | None = None
        self._drag_offset = (0.0, 0.0)

        self.top = tk.Toplevel(root)
        self.top.title("Adjust seal photo crop")
        self.top.configure(bg="#0c0c0c")
        self.top.transient(root)
        self.top.grab_set()
        self.top.resizable(False, False)
        self.top.protocol("WM_DELETE_WINDOW", self._cancel)

        tk.Label(
            self.top,
            text="Drag the green box's corners to fit it tightly around the seal, then confirm.",
            bg="#0c0c0c",
            fg="#d9d9d9",
            font=("Segoe UI", 10),
        ).pack(padx=12, pady=(12, 6))

        self.canvas = tk.Canvas(
            self.top,
            width=self.preview.width,
            height=self.preview.height,
            highlightthickness=1,
            highlightbackground="#2a2a2a",
            bg="#050505",
        )
        self.canvas.pack(padx=12, pady=6)

        self._tk_image = ImageTk.PhotoImage(self.preview)
        self.canvas.create_image(0, 0, anchor="nw", image=self._tk_image)

        self.canvas.bind("<ButtonPress-1>", self._on_press)
        self.canvas.bind("<B1-Motion>", self._on_drag)
        self.canvas.bind("<ButtonRelease-1>", self._on_release)

        button_row = tk.Frame(self.top, bg="#0c0c0c")
        button_row.pack(pady=(6, 12))
        tk.Button(button_row, text="Re-detect", command=self._redetect).grid(row=0, column=0, padx=4)
        tk.Button(button_row, text="Use full photo", command=self._use_full).grid(row=0, column=1, padx=4)
        tk.Button(button_row, text="Cancel", command=self._cancel).grid(row=0, column=2, padx=4)
        tk.Button(button_row, text="Confirm", command=self._confirm).grid(row=0, column=3, padx=4)

        self._draw_box()
        self.top.wait_window()

    def _build_scaled_preview(self) -> None:
        img_w, img_h = self.image.size
        self.scale = max(0.05, min(CANVAS_MAX_W / img_w, CANVAS_MAX_H / img_h))
        preview_w = max(1, int(img_w * self.scale))
        preview_h = max(1, int(img_h * self.scale))
        self.preview = self.image.resize((preview_w, preview_h), Image.LANCZOS)

    def _to_canvas(self, box: list[float]) -> list[float]:
        return [coord * self.scale for coord in box]

    def _to_image(self, box: list[float]) -> list[float]:
        return [coord / self.scale for coord in box]

    @staticmethod
    def _handle_positions(l: float, t: float, r: float, b: float) -> dict[str, tuple[float, float]]:
        return {"nw": (l, t), "ne": (r, t), "sw": (l, b), "se": (r, b)}

    def _draw_box(self) -> None:
        self.canvas.delete("cropbox")
        l, t, r, b = self._to_canvas(self.box)
        self.canvas.create_rectangle(l, t, r, b, outline="#2eea6f", width=2, tags="cropbox")
        for hx, hy in self._handle_positions(l, t, r, b).values():
            self.canvas.create_rectangle(
                hx - HANDLE_HALF, hy - HANDLE_HALF, hx + HANDLE_HALF, hy + HANDLE_HALF,
                fill="#2eea6f", outline="#0c0c0c", tags="cropbox",
            )

    def _on_press(self, event: tk.Event) -> None:
        l, t, r, b = self._to_canvas(self.box)
        for name, (hx, hy) in self._handle_positions(l, t, r, b).items():
            if abs(event.x - hx) <= HANDLE_HIT_RADIUS and abs(event.y - hy) <= HANDLE_HIT_RADIUS:
                self._active_handle = name
                return
        if l <= event.x <= r and t <= event.y <= b:
            self._active_handle = "move"
            self._drag_offset = (event.x - l, event.y - t)
        else:
            self._active_handle = None

    def _on_drag(self, event: tk.Event) -> None:
        if self._active_handle is None:
            return
        preview_w, preview_h = self.preview.size
        x = min(max(event.x, 0), preview_w)
        y = min(max(event.y, 0), preview_h)
        l, t, r, b = self._to_canvas(self.box)

        if self._active_handle == "move":
            width, height = r - l, b - t
            l = min(max(x - self._drag_offset[0], 0), preview_w - width)
            t = min(max(y - self._drag_offset[1], 0), preview_h - height)
            r, b = l + width, t + height
        elif self._active_handle == "nw":
            l, t = min(x, r - HANDLE_HIT_RADIUS), min(y, b - HANDLE_HIT_RADIUS)
        elif self._active_handle == "ne":
            r, t = max(x, l + HANDLE_HIT_RADIUS), min(y, b - HANDLE_HIT_RADIUS)
        elif self._active_handle == "sw":
            l, b = min(x, r - HANDLE_HIT_RADIUS), max(y, t + HANDLE_HIT_RADIUS)
        elif self._active_handle == "se":
            r, b = max(x, l + HANDLE_HIT_RADIUS), max(y, t + HANDLE_HIT_RADIUS)

        self.box = self._to_image([l, t, r, b])
        self._draw_box()

    def _on_release(self, _event: tk.Event) -> None:
        self._active_handle = None

    def _redetect(self) -> None:
        self.box = list(suggest_crop_box(self.image))
        self._draw_box()

    def _use_full(self) -> None:
        img_w, img_h = self.image.size
        self.box = [0, 0, img_w, img_h]
        self._draw_box()

    def _confirm(self) -> None:
        left, top, right, bottom = (int(round(v)) for v in self.box)
        if right - left < 10 or bottom - top < 10:
            messagebox.showwarning(
                "Crop too small", "Please select a larger crop area around the seal.")
            return
        cropped = crop_to_box(self.image, (left, top, right, bottom))
        self.result_jpeg_bytes = encode_jpeg(cropped)
        self.top.destroy()

    def _cancel(self) -> None:
        self.result_jpeg_bytes = None
        self.top.destroy()
