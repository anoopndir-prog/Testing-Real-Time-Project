"""Final Test Report screen.

Lets the user pick the report date and attach the two source documents
(Project Specification, Test Inspection sheet) that will be merged into the
fixed Final Test Report PDF format. The merge/generation logic itself is
added in a later step; this screen only captures the inputs.
"""

from __future__ import annotations

import datetime as dt
import threading
from pathlib import Path
from typing import Callable
import tkinter as tk
from tkinter import filedialog, messagebox

from PIL import Image

from app.final_test_report_generator import (
    DistributionEntry,
    FinalReportInputs,
    SamplePhotoEntry,
    detect_sample_labels,
    generate_final_report,
)
from app.seal_photo_crop_dialog import SealPhotoCropDialog
from app.ui_common import (
    BG_DARKEST,
    BG_SHINY,
    BOX_BORDER,
    TEXT_MAIN,
    TEXT_SUB,
    RoundedButton,
    add_date_field,
    add_dropdown_field,
    add_text_field,
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
    "Drag & drop the Test Inspection & Execution sheet here or click to browse"
    if DND_FILES is not None
    else "Click to browse the Test Inspection & Execution sheet"
)
MONITORING_SHEET_DEFAULT_MAIN = "Attach the Monitoring Sheet"
MONITORING_SHEET_DEFAULT_SUB = (
    "Drag & drop one or more Monitoring sheets here or click to browse"
    if DND_FILES is not None
    else "Click to browse one or more Monitoring sheets"
)

REVIEWER_ROLES = [
    "Lead engineer Testing - Global Testing India",
    "Engineer Testing - Global Testing India",
]

APPROVER_ROLES = [
    "Manager - Global Testing India",
    "Lead engineer Testing - Global Testing India",
    "Engineer Testing - Global Testing India",
]


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
        allow_multiple: bool = False,
        width: int = 720,
        height: int = 150,
        on_change: Callable[[], None] | None = None,
    ) -> None:
        self.root = root
        self.default_main_text = default_main_text
        self.default_sub_text = default_sub_text
        self.allowed_suffixes = allowed_suffixes
        self.invalid_message = invalid_message
        self.allow_multiple = allow_multiple
        self.on_change = on_change
        self.file_path: Path | None = None
        self.file_paths: list[Path] = []

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

        candidates = [Path(str(item)).expanduser() for item in dropped_items]
        if self._set_if_valid(candidates):
            return
        messagebox.showwarning("Invalid File", self.invalid_message)

    def _browse(self, _event=None) -> None:
        downloads = Path.home() / "Downloads"
        filetype_patterns = " ".join(
            f"*{suffix}" for suffix in sorted(self.allowed_suffixes))
        dialog_kwargs = {
            "title": "Select Files" if self.allow_multiple else "Select File",
            "initialdir": str(downloads) if downloads.exists() else None,
            "filetypes": [("Allowed Files", filetype_patterns), ("All Files", "*.*")],
        }
        selected = (
            filedialog.askopenfilenames(**dialog_kwargs)
            if self.allow_multiple
            else filedialog.askopenfilename(**dialog_kwargs)
        )
        if not selected:
            return

        selected_paths = [Path(path) for path in selected] if self.allow_multiple else [Path(selected)]
        if not self._set_if_valid(selected_paths):
            messagebox.showwarning("Invalid File", self.invalid_message)

    def _set_if_valid(self, file_paths: list[Path]) -> bool:
        if not self.allow_multiple and len(file_paths) != 1:
            return False
        valid_paths = [
            path for path in file_paths
            if path.exists() and path.suffix.lower() in self.allowed_suffixes
        ]
        if len(valid_paths) != len(file_paths):
            return False

        self.file_paths = valid_paths
        self.file_path = valid_paths[0] if valid_paths else None
        self.icon_var.set("📎")
        if self.allow_multiple and len(valid_paths) > 1:
            self.main_text_var.set(f"{len(valid_paths)} Monitoring sheets attached")
            self.sub_text_var.set(", ".join(path.name for path in valid_paths[:3]) + ("..." if len(valid_paths) > 3 else ""))
        else:
            self.main_text_var.set(valid_paths[0].name)
            self.sub_text_var.set("Attached successfully")
        if self.on_change:
            self.on_change()
        return True

    def reset(self) -> None:
        self.file_path = None
        self.file_paths = []
        self.icon_var.set("📌")
        self.main_text_var.set(self.default_main_text)
        self.sub_text_var.set(self.default_sub_text)
        if self.on_change:
            self.on_change()


class SealPhotoSlot:
    """Attach box for the optional seal identification photo (section 3.2).
    Picking a file immediately opens the crop dialog, so whatever ends up
    stored here is already tightly cropped around the seal."""

    DEFAULT_MAIN = "Attach seal photo here (optional)"
    DEFAULT_SUB = "Click to browse a photo of the tested seal"

    def __init__(
        self,
        parent: tk.Widget,
        root: tk.Misc,
        width: int = 720,
        height: int = 120,
        main_text: str | None = None,
        sub_text: str | None = None,
        dialog_title: str = "Select seal photo",
    ) -> None:
        self.root = root
        self.photo_bytes: bytes | None = None
        self.default_main_text = main_text or self.DEFAULT_MAIN
        self.default_sub_text = sub_text or self.DEFAULT_SUB
        self.dialog_title = dialog_title

        self.icon_var = tk.StringVar(value="📷")
        self.main_text_var = tk.StringVar(value=self.default_main_text)
        self.sub_text_var = tk.StringVar(value=self.default_sub_text)

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
            self.box, textvariable=self.icon_var, bg=BG_SHINY, fg="#e6e6e6",
            font=("Segoe UI Emoji", 30),
        )
        self.icon_label.pack(pady=(12, 4))

        self.main_label = tk.Label(
            self.box, textvariable=self.main_text_var, bg=BG_SHINY, fg=TEXT_MAIN,
            font=("Segoe UI", 14, "bold"), wraplength=650, justify="center",
        )
        self.main_label.pack()

        self.sub_label = tk.Label(
            self.box, textvariable=self.sub_text_var, bg=BG_SHINY, fg=TEXT_SUB,
            font=("Segoe UI", 10), wraplength=650, justify="center",
        )
        self.sub_label.pack(pady=(6, 0))

        for widget in (self.box, self.icon_label, self.main_label, self.sub_label):
            widget.bind("<Button-1>", self._browse)

    def _browse(self, _event=None) -> None:
        downloads = Path.home() / "Downloads"
        selected = filedialog.askopenfilename(
            title=self.dialog_title,
            initialdir=str(downloads) if downloads.exists() else None,
            filetypes=[
                ("Image files", "*.jpg *.jpeg *.png *.bmp *.tif *.tiff"),
                ("All Files", "*.*"),
            ],
        )
        if not selected:
            return
        try:
            image = Image.open(selected)
            image.load()
        except Exception:
            messagebox.showwarning("Invalid Photo", "Could not open that file as an image.")
            return

        dialog = SealPhotoCropDialog(self.root, image)
        if dialog.result_jpeg_bytes is None:
            return

        self.photo_bytes = dialog.result_jpeg_bytes
        self.icon_var.set("✅")
        self.main_text_var.set(Path(selected).name)
        self.sub_text_var.set("Photo attached and cropped — click to change")

    def reset(self) -> None:
        self.photo_bytes = None
        self.icon_var.set("📷")
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
        self.project_leader_var = tk.StringVar()
        self.reviewer_var = tk.StringVar()
        self.approved_by_var = tk.StringVar()
        self.reviewer_role_var = tk.StringVar(value=REVIEWER_ROLES[0])
        self.approved_by_role_var = tk.StringVar(value=APPROVER_ROLES[0])
        self.distribution_rows: list[tuple[tk.StringVar,
                                           tk.StringVar, tk.StringVar]] = []
        self.sample_photo_slots: dict[str, dict[str, object]] = {}
        self._detect_seq = 0

        self._build_ui()

    def _build_ui(self) -> None:
        outer = self.container

        content = tk.Frame(outer, bg=BG_DARKEST)
        content.pack(fill="both", expand=True)

        center = tk.Frame(content, bg=BG_DARKEST)
        center.place(relx=0.5, rely=0.5, anchor="center")

        title = tk.Label(
            center,
            text="Final Test Report",
            bg=BG_DARKEST,
            fg="#e8e8e8",
            font=("Segoe UI", 22, "bold"),
        )
        title.pack(pady=(0, 18))

        date_row = tk.Frame(center, bg=BG_DARKEST)
        date_row.pack(pady=(0, 10))
        add_date_field(
            date_row, "Date", self.report_date_var, 0, self.root, allow_clear=False, row=0)
        add_text_field(
            date_row, "Project Leader", self.project_leader_var, 1, self.root, width=20, row=0)
        add_text_field(
            date_row, "Reviewer", self.reviewer_var, 2, self.root, width=20, row=0)
        add_text_field(
            date_row, "Approved by", self.approved_by_var, 3, self.root, width=20, row=0)

        role_row = tk.Frame(center, bg=BG_DARKEST)
        role_row.pack(pady=(0, 10))
        add_dropdown_field(
            role_row, "Reviewer Role", self.reviewer_role_var, 0, REVIEWER_ROLES, width=38)
        add_dropdown_field(
            role_row, "Approved By Role", self.approved_by_role_var, 1, APPROVER_ROLES, width=38)

        self.project_spec_slot = AttachSlot(
            center,
            self.root,
            PROJECT_SPEC_DEFAULT_MAIN,
            PROJECT_SPEC_DEFAULT_SUB,
            allowed_suffixes={".pdf", ".docx"},
            invalid_message="Please attach a valid Project Specification file (.pdf or .docx).",
            height=120,
            on_change=self._on_detect_inputs_changed,
        )
        self.project_spec_slot.box.pack(pady=(0, 10))

        attachment_row = tk.Frame(center, bg=BG_DARKEST)
        attachment_row.pack(pady=(0, 10))

        self.test_inspection_slot = AttachSlot(
            attachment_row,
            self.root,
            TEST_INSPECTION_DEFAULT_MAIN,
            TEST_INSPECTION_DEFAULT_SUB,
            allowed_suffixes={".pdf", ".docx", ".xlsx", ".xlsm"},
            invalid_message="Please attach a valid Test Inspection & Execution sheet (.pdf, .docx, .xlsx, .xlsm).",
            width=355,
            height=120,
            on_change=self._on_detect_inputs_changed,
        )
        self.test_inspection_slot.box.grid(row=0, column=0, padx=(0, 10))

        self.monitoring_sheet_slot = AttachSlot(
            attachment_row,
            self.root,
            MONITORING_SHEET_DEFAULT_MAIN,
            MONITORING_SHEET_DEFAULT_SUB,
            allowed_suffixes={".pdf", ".docx", ".xlsx", ".xlsm"},
            invalid_message="Please attach valid Monitoring sheet files (.pdf, .docx, .xlsx, .xlsm).",
            allow_multiple=True,
            width=355,
            height=120,
            on_change=self._on_detect_inputs_changed,
        )
        self.monitoring_sheet_slot.box.grid(row=0, column=1, padx=(0, 0))

        self.seal_photo_slot = SealPhotoSlot(center, self.root, height=110)
        self.seal_photo_slot.box.pack(pady=(0, 10))

        sample_photos_header = tk.Frame(center, bg=BG_DARKEST)
        sample_photos_header.pack(fill="x", pady=(4, 0))
        tk.Label(
            sample_photos_header,
            text="Before / After Test Photos (optional)",
            bg=BG_DARKEST,
            fg="#d7d7d7",
            font=("Segoe UI", 11, "bold"),
        ).pack(side="left", padx=(4, 0))

        self.sample_photos_status_var = tk.StringVar(
            value="Attach the Project Specification and Test Inspection sheet above to detect samples.")
        tk.Label(
            center,
            textvariable=self.sample_photos_status_var,
            bg=BG_DARKEST,
            fg=TEXT_SUB,
            font=("Segoe UI", 10),
            wraplength=720,
            justify="center",
        ).pack(pady=(2, 6))

        self.sample_photos_container = tk.Frame(center, bg=BG_DARKEST)
        self.sample_photos_container.pack(fill="x", pady=(0, 10))

        self._build_distribution_list(center)

        button_area = tk.Frame(center, bg=BG_DARKEST)
        button_area.pack(pady=(18, 0))

        self.word_btn = RoundedButton(
            button_area, "Generate Word", lambda: self._generate("word"),
            width=170, height=48,
        )
        self.word_btn.grid(row=0, column=0, padx=6)

        self.pdf_btn = RoundedButton(
            button_area, "Generate PDF", lambda: self._generate("pdf"),
            width=170, height=48,
        )
        self.pdf_btn.grid(row=0, column=1, padx=6)

        self.back_btn = RoundedButton(
            button_area, "Back", self.on_back,
            width=140, height=48,
        )
        self.back_btn.grid(row=0, column=2, padx=6)

        self.reset_btn = RoundedButton(
            button_area, "Reset", self._reset,
            width=140, height=48,
        )
        self.reset_btn.grid(row=0, column=3, padx=6)

        self.exit_btn = RoundedButton(
            button_area, "Exit", self.root.destroy,
            width=140, height=48,
        )
        self.exit_btn.grid(row=0, column=4, padx=6)

    def _build_distribution_list(self, parent: tk.Widget) -> None:
        frame = tk.Frame(parent, bg=BG_DARKEST)
        frame.pack(fill="x", pady=(0, 0))

        header = tk.Frame(frame, bg=BG_DARKEST)
        header.pack(fill="x")
        tk.Label(
            header,
            text="Distribution List",
            bg=BG_DARKEST,
            fg="#d7d7d7",
            font=("Segoe UI", 11, "bold"),
        ).pack(side="left", padx=(4, 12))

        RoundedButton(
            header,
            "Add New",
            self._add_distribution_row,
            width=120,
            height=34,
            radius=16,
        ).pack(side="left")

        grid = tk.Frame(frame, bg=BG_DARKEST)
        grid.pack(fill="x", pady=(8, 0))
        self.distribution_grid = grid

        for col, label in enumerate(("Name", "Organization", "Location")):
            tk.Label(
                grid,
                text=label,
                bg=BG_DARKEST,
                fg=TEXT_SUB,
                font=("Segoe UI", 9, "bold"),
            ).grid(row=0, column=col, padx=4, sticky="w")

        self._add_distribution_row()

    def _add_distribution_row(self) -> None:
        row_idx = len(self.distribution_rows) + 1
        vars_row = (tk.StringVar(), tk.StringVar(), tk.StringVar())
        widths = (26, 34, 24)
        for col, var in enumerate(vars_row):
            entry = tk.Entry(
                self.distribution_grid,
                textvariable=var,
                width=widths[col],
                bg="#101010",
                fg="#f0f0f0",
                insertbackground="#f0f0f0",
                relief="flat",
                highlightthickness=1,
                highlightbackground="#2a2a2a",
                highlightcolor="#2a2a2a",
                font=("Segoe UI", 10),
            )
            entry.grid(row=row_idx, column=col, padx=4,
                       pady=(0, 4), ipady=3, sticky="w")
        self.distribution_rows.append(vars_row)

    def _on_detect_inputs_changed(self) -> None:
        """Re-runs sample detection whenever the Project Spec, Test
        Inspection sheet, or Monitoring sheets change, so the Before/After
        photo slots always match what the report will actually generate.
        Any photos already attached to the old slots are discarded, since
        there's no way to know they still correspond to the right sample
        once the source files change."""
        self._detect_seq += 1
        seq = self._detect_seq
        self._clear_sample_photo_slots()

        project_spec = self.project_spec_slot.file_path
        inspection = self.test_inspection_slot.file_path
        if project_spec is None or inspection is None:
            self.sample_photos_status_var.set(
                "Attach the Project Specification and Test Inspection sheet above to detect samples.")
            return

        monitoring_paths = list(self.monitoring_sheet_slot.file_paths)
        self.sample_photos_status_var.set("Detecting samples…")

        def worker() -> None:
            try:
                labels = detect_sample_labels(project_spec, inspection, monitoring_paths)
                error = None
            except Exception as exc:
                labels = []
                error = str(exc)
            self.root.after(0, lambda: self._on_detect_complete(seq, labels, error))

        threading.Thread(target=worker, daemon=True).start()

    def _on_detect_complete(self, seq: int, labels: list[str], error: str | None) -> None:
        if seq != self._detect_seq:
            return  # Inputs changed again while this run was still parsing.
        if error:
            self.sample_photos_status_var.set(f"Could not detect samples: {error}")
            return
        if not labels:
            self.sample_photos_status_var.set("No samples detected.")
            return
        plural = "s" if len(labels) != 1 else ""
        self.sample_photos_status_var.set(
            f"Detected {len(labels)} sample{plural} — attach Before/After photos below (optional).")
        self._build_sample_photo_slots(labels)

    def _clear_sample_photo_slots(self) -> None:
        for widget in self.sample_photos_container.winfo_children():
            widget.destroy()
        self.sample_photo_slots = {}

    def _build_sample_photo_slots(self, labels: list[str]) -> None:
        for label in labels:
            row = tk.Frame(self.sample_photos_container, bg=BG_DARKEST)
            row.pack(pady=(0, 8))
            tk.Label(
                row, text=label, bg=BG_DARKEST, fg=TEXT_MAIN, font=("Segoe UI", 11, "bold"),
            ).pack(anchor="w", padx=4, pady=(0, 4))

            pair_frame = tk.Frame(row, bg=BG_DARKEST)
            pair_frame.pack()

            before_col = tk.Frame(pair_frame, bg=BG_DARKEST)
            before_col.grid(row=0, column=0, padx=(0, 10))

            before_slot = SealPhotoSlot(
                before_col, self.root, width=345, height=100,
                main_text="Attach Before Test photo",
                sub_text=f"For {label} — click to browse",
                dialog_title=f"Select Before Test photo — {label}",
            )
            before_slot.box.pack()

            tk.Label(
                before_col, text="What does this show? (optional, e.g. Shaft, Bore)",
                bg=BG_DARKEST, fg=TEXT_SUB, font=("Segoe UI", 8),
            ).pack(anchor="w", pady=(4, 1))
            before_caption_var = tk.StringVar()
            tk.Entry(
                before_col, textvariable=before_caption_var, bg="#101010", fg="#f0f0f0",
                insertbackground="#f0f0f0", relief="flat", highlightthickness=1,
                highlightbackground="#2a2a2a", highlightcolor="#2a2a2a", font=("Segoe UI", 9),
            ).pack(fill="x", ipady=3)

            after_slot = SealPhotoSlot(
                pair_frame, self.root, width=345, height=100,
                main_text="Attach After Test photo",
                sub_text=f"For {label} — click to browse",
                dialog_title=f"Select After Test photo — {label}",
            )
            after_slot.box.grid(row=0, column=1, sticky="n")

            self.sample_photo_slots[label] = {
                "before": before_slot,
                "before_caption": before_caption_var,
                "after": after_slot,
            }

    def _reset(self) -> None:
        self.report_date_var.set(dt.date.today().strftime("%d/%m/%Y"))
        self.project_leader_var.set("")
        self.reviewer_var.set("")
        self.approved_by_var.set("")
        self.reviewer_role_var.set(REVIEWER_ROLES[0])
        self.approved_by_role_var.set(APPROVER_ROLES[0])
        self.project_spec_slot.reset()
        self.test_inspection_slot.reset()
        self.monitoring_sheet_slot.reset()
        self.seal_photo_slot.reset()
        self._detect_seq += 1
        self._clear_sample_photo_slots()
        self.sample_photos_status_var.set(
            "Attach the Project Specification and Test Inspection sheet above to detect samples.")
        for widgets in self.distribution_grid.grid_slaves():
            row = int(widgets.grid_info().get("row", 0))
            if row > 0:
                widgets.destroy()
        self.distribution_rows.clear()
        self._add_distribution_row()

    def _generate(self, output_format: str) -> None:
        if self.project_spec_slot.file_path is None:
            messagebox.showwarning(
                "Missing File", "Please attach the Project Specification file.")
            return
        if self.test_inspection_slot.file_path is None:
            messagebox.showwarning(
                "Missing File", "Please attach the Test Inspection & Execution sheet.")
            return

        distribution = [
            DistributionEntry(name.get().strip(),
                              org.get().strip(), location.get().strip())
            for name, org, location in self.distribution_rows
        ]
        sample_photos = {}
        for label, slots in self.sample_photo_slots.items():
            before_bytes = slots["before"].photo_bytes
            after_bytes = slots["after"].photo_bytes
            if before_bytes or after_bytes:
                sample_photos[label] = SamplePhotoEntry(
                    before_photo=before_bytes,
                    before_caption=slots["before_caption"].get().strip(),
                    after_photo=after_bytes,
                )
        inputs = FinalReportInputs(
            project_spec_path=self.project_spec_slot.file_path,
            inspection_sheet_path=self.test_inspection_slot.file_path,
            monitoring_sheet_paths=self.monitoring_sheet_slot.file_paths,
            report_date=self.report_date_var.get(),
            project_leader=self.project_leader_var.get(),
            reviewer=self.reviewer_var.get(),
            approved_by=self.approved_by_var.get(),
            reviewer_role=self.reviewer_role_var.get(),
            approved_by_role=self.approved_by_role_var.get(),
            distribution=distribution,
            seal_photo=self.seal_photo_slot.photo_bytes,
            sample_photos=sample_photos,
        )

        try:
            output_path = generate_final_report(inputs, output_format)
        except Exception as exc:
            messagebox.showerror("Generation Failed", str(exc))
            return

        messagebox.showinfo("Report Generated", f"Saved to:\n{output_path}")
