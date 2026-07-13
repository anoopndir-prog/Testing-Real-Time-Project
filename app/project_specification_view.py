"""Project Specification screen: attach the SKF Excel request sheet and
generate the Word/PDF Test Project Specification report.

This is the pre-existing tool, unchanged in appearance and behavior, now
built inside a container frame so it can be shown/hidden from the main menu
shell instead of owning the root window outright.
"""

from __future__ import annotations

import sys
import tempfile
import time
import datetime as dt
from pathlib import Path
from typing import Callable
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from tools.excel_to_word_converter import convert

from app.ui_common import (
    BG_DARKEST,
    BG_SHINY,
    BOX_BORDER,
    TEXT_MAIN,
    TEXT_SUB,
    RoundedButton,
    resource_path,
    add_text_field,
    add_dropdown_field,
    add_date_field,
)

try:
    from tkinterdnd2 import DND_FILES  # type: ignore
except Exception:  # pragma: no cover
    DND_FILES = None


TEMPLATE_RELATIVE_PATH = Path(
    "assets") / "Project Specification - Template.docx"
DECISION_RULE_SOURCE_RELATIVE_PATH = Path(
    "assets") / "Project Specification - Decision Rule Source.docx"


def _convert_docx_to_pdf(docx_path: Path, pdf_path: Path) -> None:
    errors: list[str] = []
    platform: str = sys.platform  # str annotation prevents static literal narrowing

    if platform.startswith("win"):
        try:
            import win32com.client  # type: ignore

            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            document = word.Documents.Open(str(docx_path.resolve()))
            try:
                document.SaveAs(str(pdf_path.resolve()), FileFormat=17)
            finally:
                document.Close(False)
                word.Quit()
            return
        except Exception as exc:
            errors.append(f"MS Word automation failed: {exc}")

        try:
            from docx2pdf import convert as docx2pdf_convert

            docx2pdf_convert(str(docx_path.resolve()), str(pdf_path.resolve()))
            if pdf_path.exists():
                return
            errors.append("docx2pdf completed without creating output PDF")
        except Exception as exc:
            errors.append(f"docx2pdf failed: {exc}")

        raise RuntimeError(" | ".join(errors))

    elif platform == "darwin":
        # Try Microsoft Word via docx2pdf
        try:
            from docx2pdf import convert as docx2pdf_convert

            docx2pdf_convert(str(docx_path.resolve()), str(pdf_path.resolve()))
            if pdf_path.exists():
                return
            errors.append("docx2pdf completed without creating output PDF")
        except Exception as exc:
            errors.append(f"docx2pdf (Word) failed: {exc}")

        # Fallback: Apple Pages via AppleScript (built into macOS)
        try:
            import subprocess

            def _applescript_escape(text: str) -> str:
                return text.replace("\\", "\\\\").replace('"', '\\"')

            script = f'''
tell application "Pages"
    set theDoc to open POSIX file "{_applescript_escape(str(docx_path.resolve()))}"
    delay 2
    export theDoc to POSIX file "{_applescript_escape(str(pdf_path.resolve()))}" as PDF
    close theDoc saving no
end tell
'''
            result = subprocess.run(
                ["osascript", "-e", script],
                capture_output=True, text=True, timeout=60,
            )
            if pdf_path.exists():
                return
            errors.append(f"Pages export failed: {result.stderr.strip()}")
        except Exception as exc:
            errors.append(f"Apple Pages automation failed: {exc}")

        # Fallback: LibreOffice if installed
        try:
            import subprocess

            lo_candidates = [
                "/Applications/LibreOffice.app/Contents/MacOS/soffice",
                "libreoffice",
                "soffice",
            ]
            lo_bin = next((c for c in lo_candidates if Path(c).exists()), None)
            if lo_bin:
                result = subprocess.run(
                    [lo_bin, "--headless", "--convert-to", "pdf",
                     "--outdir", str(pdf_path.parent), str(docx_path.resolve())],
                    capture_output=True, text=True, timeout=120,
                )
                converted = pdf_path.parent / (docx_path.stem + ".pdf")
                if converted.exists():
                    if converted != pdf_path:
                        converted.rename(pdf_path)
                    return
                errors.append(f"LibreOffice conversion failed: {result.stderr.strip()}")
            else:
                errors.append("LibreOffice not found")
        except Exception as exc:
            errors.append(f"LibreOffice failed: {exc}")

        raise RuntimeError(
            "PDF conversion failed. Install Microsoft Word or LibreOffice for reliable PDF export.\n\n"
            + " | ".join(errors)
        )

    else:
        try:
            from docx2pdf import convert as docx2pdf_convert

            docx2pdf_convert(str(docx_path.resolve()), str(pdf_path.resolve()))
            if pdf_path.exists():
                return
            errors.append("docx2pdf completed without creating output PDF")
        except Exception as exc:
            errors.append(f"docx2pdf failed: {exc}")

        raise RuntimeError(" | ".join(errors))


def _write_instruction_pdf(output_path: Path) -> None:
    """Create a simple PDF file with usage instructions."""

    def esc(text: str) -> str:
        return text.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")

    title = "SKF Report Generator - Work Instruction"
    lines = [
        "1. Open the software.",
        "2. Attach Excel using drag-and-drop or click the center box.",
        "3. Ensure the file is the SKF request template (.xlsm/.xlsx).",
        "4. Click Generate Report in Word to create editable .docx output.",
        "5. Click Generate Report in PDF to create PDF output.",
        "6. Output files are saved to your Downloads folder.",
        "7. Use File > Reset to clear selected Excel and start over.",
        "8. Use File > Exit or Exit button to close the software.",
        "",
        "Important:",
        "- The tool keeps fixed template content unchanged (monitoring/disclaimer/tolerance line).",
        "- Word template must remain in assets folder unless manually selected.",
    ]

    commands = [
        "BT",
        "/F1 18 Tf",
        "50 790 Td",
        f"({esc(title)}) Tj",
        "ET",
        "BT",
        "/F1 11 Tf",
        "50 760 Td",
        "14 TL",
    ]
    for line in lines:
        commands.append(f"({esc(line)}) Tj")
        commands.append("T*")
    commands.append("ET")

    stream = "\n".join(commands).encode("latin-1", "replace")

    objects = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>",
        b"<< /Length " + str(len(stream)).encode("ascii") +
        b" >>\nstream\n" + stream + b"\nendstream",
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
    ]

    pdf = bytearray(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
    offsets = [0]

    for index, obj in enumerate(objects, start=1):
        offsets.append(len(pdf))
        pdf.extend(f"{index} 0 obj\n".encode("ascii"))
        pdf.extend(obj)
        pdf.extend(b"\nendobj\n")

    xref_start = len(pdf)
    pdf.extend(f"xref\n0 {len(objects) + 1}\n".encode("ascii"))
    pdf.extend(b"0000000000 65535 f \n")
    for off in offsets[1:]:
        pdf.extend(f"{off:010d} 00000 n \n".encode("ascii"))

    pdf.extend(
        (
            f"trailer\n<< /Size {len(objects) + 1} /Root 1 0 R >>\n"
            f"startxref\n{xref_start}\n%%EOF\n"
        ).encode("ascii")
    )

    output_path.write_bytes(pdf)


class ProjectSpecificationView:
    """Builds the Project Specification screen inside a given container frame."""

    def __init__(self, container: tk.Widget, root: tk.Misc, on_back: Callable[[], None]) -> None:
        self.container = container
        self.root = root
        self.on_back = on_back

        self.excel_path: Path | None = None
        self.template_path = resource_path(TEMPLATE_RELATIVE_PATH)
        self.decision_rule_source_path = resource_path(
            DECISION_RULE_SOURCE_RELATIVE_PATH)

        today = dt.date.today().strftime("%d/%m/%Y")
        self.report_date_var = tk.StringVar(value=today)
        self.revision_var = tk.StringVar(value="0")
        self.revision_date_var = tk.StringVar(value="")
        self.project_no_var = tk.StringVar(value="")
        self.project_leader_var = tk.StringVar(value="")
        self.tooling_lead_time_var = tk.StringVar(value="Available")

        self.icon_var = tk.StringVar(value="📌")
        self.main_text_var = tk.StringVar(value="Attach your excel")
        self.sub_text_var = tk.StringVar(
            value="Drag & drop your Excel here or click to browse"
            if DND_FILES is not None
            else "Click to browse your Excel file"
        )

        self._build_ui()

    def _build_ui(self) -> None:
        outer = self.container

        self._build_menu(outer)

        content = tk.Frame(outer, bg=BG_DARKEST)
        content.pack(fill="both", expand=True)

        center = tk.Frame(content, bg=BG_DARKEST)
        center.place(relx=0.5, rely=0.52, anchor="center")

        title = tk.Label(
            center,
            text="SKF Test Report Generator",
            bg=BG_DARKEST,
            fg="#e8e8e8",
            font=("Segoe UI", 22, "bold"),
        )
        title.pack(pady=(0, 22))

        meta_row = tk.Frame(center, bg=BG_DARKEST)
        meta_row.pack(fill="x", pady=(0, 16))

        add_date_field(
            meta_row, "Date", self.report_date_var, 0, self.root, allow_clear=False, row=0)
        add_text_field(
            meta_row, "Revision", self.revision_var, 1, self.root, width=8, numeric_only=True, row=0)
        add_date_field(meta_row, "Revision Date",
                       self.revision_date_var, 2, self.root, allow_clear=True, row=0)
        add_text_field(meta_row, "Project No.", self.project_no_var,
                       3, self.root, width=18, numeric_only=False, row=0)
        add_dropdown_field(
            meta_row,
            "Tooling Lead Time",
            self.tooling_lead_time_var,
            0,
            values=["Available"] + [f"{i} Week" if i ==
                                    1 else f"{i} Weeks" for i in range(1, 11)],
            width=13,
            row=1,
        )
        add_text_field(meta_row, "Project Leader",
                       self.project_leader_var, 1, self.root, width=22, numeric_only=False, row=1)

        self.attach_box = tk.Frame(
            center,
            bg=BG_SHINY,
            width=720,
            height=300,
            highlightbackground=BOX_BORDER,
            highlightthickness=2,
            cursor="hand2",
        )
        self.attach_box.pack()
        self.attach_box.pack_propagate(False)

        shine_line = tk.Frame(self.attach_box, bg="#2f2f2f", height=2)
        shine_line.pack(fill="x", side="top")

        self.icon_label = tk.Label(
            self.attach_box,
            textvariable=self.icon_var,
            bg=BG_SHINY,
            fg="#e6e6e6",
            font=("Segoe UI Emoji", 58),
        )
        self.icon_label.pack(pady=(40, 8))

        self.main_label = tk.Label(
            self.attach_box,
            textvariable=self.main_text_var,
            bg=BG_SHINY,
            fg=TEXT_MAIN,
            font=("Segoe UI", 20, "bold"),
            wraplength=650,
            justify="center",
        )
        self.main_label.pack()

        self.sub_label = tk.Label(
            self.attach_box,
            textvariable=self.sub_text_var,
            bg=BG_SHINY,
            fg=TEXT_SUB,
            font=("Segoe UI", 11),
            wraplength=650,
            justify="center",
        )
        self.sub_label.pack(pady=(12, 0))

        for widget in (self.attach_box, self.icon_label, self.main_label, self.sub_label):
            widget.bind("<Button-1>", self._attach_excel)

        self._setup_drag_drop()

        button_area = tk.Frame(center, bg=BG_DARKEST)
        button_area.pack(pady=(28, 0))

        self.word_btn = RoundedButton(
            button_area, "Generate Report in Word", self._generate_word)
        self.word_btn.pack(pady=(0, 14))

        self.pdf_btn = RoundedButton(
            button_area, "Generate Report in PDF", self._generate_pdf)
        self.pdf_btn.pack(pady=(0, 14))

        self.exit_btn = RoundedButton(button_area, "Exit", self.root.destroy)
        self.exit_btn.pack()

    def _build_menu(self, parent: tk.Widget) -> None:
        menu_bar = tk.Frame(parent, bg="#0f0f0f", height=42,
                            highlightbackground="#1d1d1d", highlightthickness=1)
        menu_bar.pack(fill="x", side="top")
        menu_bar.pack_propagate(False)

        menu_left = tk.Frame(menu_bar, bg="#0f0f0f")
        menu_left.pack(side="left", padx=(12, 0))

        file_btn = tk.Menubutton(
            menu_left,
            text="File",
            bg="#0f0f0f",
            fg="#e3e3e3",
            activebackground="#161616",
            activeforeground="#ffffff",
            relief="flat",
            font=("Segoe UI", 11, "bold"),
            padx=10,
            pady=7,
            cursor="hand2",
        )
        file_btn.pack(side="left", pady=2)
        file_menu = tk.Menu(file_btn, tearoff=0, bg="#151515",
                            fg="#ebebeb", activebackground="#1f1f1f")
        file_menu.add_command(label="Reset", command=self._reset_selection)
        file_menu.add_command(
            label="Back to Main Menu", command=self.on_back)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.destroy)
        file_btn.configure(menu=file_menu)

        help_btn = tk.Menubutton(
            menu_left,
            text="Help",
            bg="#0f0f0f",
            fg="#e3e3e3",
            activebackground="#161616",
            activeforeground="#ffffff",
            relief="flat",
            font=("Segoe UI", 11, "bold"),
            padx=10,
            pady=7,
            cursor="hand2",
        )
        help_btn.pack(side="left", pady=2)
        help_menu = tk.Menu(help_btn, tearoff=0, bg="#151515",
                            fg="#ebebeb", activebackground="#1f1f1f")
        help_menu.add_command(
            label="Download Work Instruction (PDF)", command=self._download_work_instruction)
        help_btn.configure(menu=help_menu)

    def _setup_drag_drop(self) -> None:
        if DND_FILES is None:
            return

        for widget in (self.attach_box, self.icon_label, self.main_label, self.sub_label):
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
        if self._set_excel_if_valid(candidate):
            return
        messagebox.showwarning(
            "Invalid File", "Please drop a valid Excel file (.xlsm or .xlsx).")

    def _attach_excel(self, _event=None) -> None:
        downloads = Path.home() / "Downloads"
        selected = filedialog.askopenfilename(
            title="Select Excel Sheet",
            initialdir=str(downloads) if downloads.exists() else None,
            filetypes=[("Excel Files", "*.xlsm *.xlsx"), ("All Files", "*.*")],
        )
        if not selected:
            return

        selected_path = Path(selected)
        if not self._set_excel_if_valid(selected_path):
            messagebox.showwarning(
                "Invalid File", "Please select a valid Excel file (.xlsm or .xlsx).")

    def _set_excel_if_valid(self, file_path: Path) -> bool:
        if not file_path.exists() or file_path.suffix.lower() not in {".xlsm", ".xlsx"}:
            return False

        self.excel_path = file_path
        self.icon_var.set("📊")
        self.main_text_var.set(file_path.name)
        self.sub_text_var.set("Excel attached successfully")
        return True

    def _reset_selection(self) -> None:
        self.excel_path = None
        self.icon_var.set("📌")
        self.main_text_var.set("Attach your excel")
        self.sub_text_var.set(
            "Drag & drop your Excel here or click to browse"
            if DND_FILES is not None
            else "Click to browse your Excel file"
        )
        self.report_date_var.set(dt.date.today().strftime("%d/%m/%Y"))
        self.revision_var.set("0")
        self.revision_date_var.set("")
        self.project_no_var.set("")
        self.project_leader_var.set("")
        self.tooling_lead_time_var.set("Available")

    def _ensure_inputs(self) -> bool:
        if self.excel_path is None:
            messagebox.showwarning(
                "Attach Excel", "Please attach an Excel file first.")
            return False

        if not self.project_no_var.get().strip():
            messagebox.showwarning(
                "Project No.", "Please enter Project No. (e.g., TR26-0002-BTS).")
            return False

        if not self.project_leader_var.get().strip():
            messagebox.showwarning(
                "Project Leader", "Please enter Project Leader name.")
            return False

        if not self.revision_var.get().strip():
            self.revision_var.set("0")

        if not self.template_path.exists():
            messagebox.showwarning(
                "Template Missing",
                "Word template is missing. Please select the template .docx file.",
            )
            selected = filedialog.askopenfilename(
                title="Select Word Template",
                filetypes=[("Word Document", "*.docx"), ("All Files", "*.*")],
            )
            if not selected:
                return False
            self.template_path = Path(selected)

        return True

    def _generate_word(self) -> None:
        if not self._ensure_inputs():
            return

        out_path = self._default_download_path(".docx")
        try:
            with tempfile.TemporaryDirectory() as tmp:
                convert(
                    self.excel_path,
                    self.template_path,
                    out_path,
                    Path(tmp),
                    report_date=self.report_date_var.get().strip() or None,
                    revision_no=self.revision_var.get().strip() or None,
                    revision_date=self.revision_date_var.get().strip() or None,
                    project_no=self.project_no_var.get().strip() or None,
                    project_leader=self.project_leader_var.get().strip() or None,
                    tooling_lead_time=self.tooling_lead_time_var.get().strip() or None,
                    decision_rule_source_path=self.decision_rule_source_path if self.decision_rule_source_path.exists() else None,
                )
            messagebox.showinfo(
                "Success", f"Word report generated and downloaded to:\n{out_path}")
        except Exception as exc:
            messagebox.showerror("Generation Failed",
                                 f"Could not generate Word report.\n\n{exc}")

    def _generate_pdf(self) -> None:
        if not self._ensure_inputs():
            return

        out_pdf_path = self._default_download_path(".pdf")
        try:
            with tempfile.TemporaryDirectory() as tmp:
                temp_docx = Path(tmp) / "temp_report.docx"
                convert(
                    self.excel_path,
                    self.template_path,
                    temp_docx,
                    Path(tmp) / "assets",
                    report_date=self.report_date_var.get().strip() or None,
                    revision_no=self.revision_var.get().strip() or None,
                    revision_date=self.revision_date_var.get().strip() or None,
                    project_no=self.project_no_var.get().strip() or None,
                    project_leader=self.project_leader_var.get().strip() or None,
                    tooling_lead_time=self.tooling_lead_time_var.get().strip() or None,
                    decision_rule_source_path=self.decision_rule_source_path if self.decision_rule_source_path.exists() else None,
                )
                _convert_docx_to_pdf(temp_docx, out_pdf_path)
            messagebox.showinfo(
                "Success", f"PDF report generated and downloaded to:\n{out_pdf_path}")
        except Exception as exc:
            messagebox.showerror(
                "Generation Failed",
                "Could not generate PDF report.\n"
                "For Windows build, install Microsoft Word for reliable PDF conversion.\n\n"
                f"{exc}",
            )

    def _download_work_instruction(self) -> None:
        output_path = self._default_download_named(
            "SKF Report Generator - Work Instruction.pdf")
        try:
            _write_instruction_pdf(output_path)
            messagebox.showinfo(
                "Downloaded", f"Work instruction downloaded to:\n{output_path}")
        except Exception as exc:
            messagebox.showerror(
                "Download Failed", f"Could not download work instruction PDF.\n\n{exc}")

    def _default_download_path(self, extension: str) -> Path:
        return self._default_download_named(f"{self.excel_path.stem} - Generated Report{extension}")

    def _default_download_named(self, filename: str) -> Path:
        downloads = Path.home() / "Downloads"
        downloads.mkdir(parents=True, exist_ok=True)

        target = downloads / filename
        if not target.exists():
            return target

        base = Path(filename).stem
        suffix = Path(filename).suffix

        for index in range(2, 1000):
            candidate = downloads / f"{base} ({index}){suffix}"
            if not candidate.exists():
                return candidate

        return downloads / f"{base} - {int(time.time())}{suffix}"
