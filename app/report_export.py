"""Docx -> PDF conversion and the Work Instruction PDF builder.

Shared by the Tkinter desktop views and the Flask web app. Deliberately
free of any UI toolkit import (no Tkinter, no Flask) so the web server
process never has to pull in Tkinter just to reach this code.
"""

from __future__ import annotations

import sys
from pathlib import Path


def convert_docx_to_pdf(docx_path: Path, pdf_path: Path) -> None:
    errors: list[str] = []
    platform: str = sys.platform  # str annotation prevents static literal narrowing

    if platform.startswith("win"):
        try:
            import pythoncom  # type: ignore
            import win32com.client  # type: ignore

            # COM apartments are per-thread. The desktop app gets this for
            # free on Tkinter's single main thread; a web request handled
            # on a waitress worker thread has not called this, and without
            # it DispatchEx intermittently raises "CoInitialize has not
            # been called" under concurrent/multi-threaded use.
            pythoncom.CoInitialize()
            try:
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
            finally:
                pythoncom.CoUninitialize()
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


def write_instruction_pdf(output_path: Path) -> None:
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
