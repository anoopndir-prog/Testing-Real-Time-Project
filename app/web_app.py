"""Browser (HTML) version of the SKF Report Generator.

Serves the same two tools as the Tkinter desktop app
(`app/report_generator_app.py`) - Project Specification and Final Test
Report - as server-rendered forms, so it can be shared over the office LAN:
one instance runs (bound to 0.0.0.0), coworkers reach it from their own
browsers. Files are uploaded rather than picked from local disk, and
generated reports are returned as a browser download instead of being
written into the *server's* Downloads folder.

Run directly for local/dev use: `python3 app/web_app.py`. The packaged
.exe/.app builds point at this same module and always serve through
waitress (never Flask's own dev server).
"""

from __future__ import annotations

import datetime as dt
import io
import os
import shutil
import socket
import sys
import tempfile
import threading
import time
import webbrowser
from pathlib import Path

from flask import Flask, jsonify, render_template, request, send_file
from PIL import Image
from werkzeug.utils import secure_filename

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from tools.excel_to_word_converter import convert

from app.final_test_report_generator import (
    DistributionEntry,
    FinalReportInputs,
    SamplePhotoEntry,
    detect_sample_labels,
    generate_final_report,
)
from app.paths import resource_path
from app.report_export import convert_docx_to_pdf, write_instruction_pdf
from app import seal_photo

DEFAULT_PORT = 8877
MAX_CONTENT_LENGTH = 200 * 1024 * 1024  # scanned/handwritten monitoring sheets can be large

REVIEWER_ROLES = [
    "Lead engineer Testing - Global Testing India",
    "Engineer Testing - Global Testing India",
]
APPROVER_ROLES = [
    "Manager - Global Testing India",
    "Lead engineer Testing - Global Testing India",
    "Engineer Testing - Global Testing India",
]
TOOLING_LEAD_TIME_OPTIONS = ["Available"] + [
    f"{i} Week" if i == 1 else f"{i} Weeks" for i in range(1, 11)
]

PROJECT_SPEC_TEMPLATE_PATH = resource_path(Path("assets") / "Project Specification - Template.docx")
DECISION_RULE_SOURCE_PATH = resource_path(Path("assets") / "Project Specification - Decision Rule Source.docx")

PROJECT_SPEC_ALLOWED_SUFFIXES = {".xlsx", ".xlsm"}
DOC_ALLOWED_SUFFIXES = {".pdf", ".docx"}
SHEET_ALLOWED_SUFFIXES = {".pdf", ".docx", ".xlsx", ".xlsm"}

# Word/Pages/LibreOffice automation is not safe under concurrent calls, and
# with the app shared over LAN, concurrent requests from different
# coworkers are a real scenario (they weren't for the single-user desktop
# app). Serializing PDF generation is a simple, sufficient guard for a
# small office tool - not worth a job queue.
_PDF_LOCK = threading.Lock()

app = Flask(
    __name__,
    template_folder=str(resource_path(Path("app") / "templates")),
    static_folder=str(resource_path(Path("app") / "static")),
)
app.config["MAX_CONTENT_LENGTH"] = MAX_CONTENT_LENGTH


def _format_date_for_report(value: str) -> str:
    """HTML `<input type=date>` submits `YYYY-MM-DD`; the report backends
    expect `DD/MM/YYYY` (what the desktop app's calendar widgets have
    always produced). Passes anything else through unchanged."""
    value = (value or "").strip()
    if not value:
        return ""
    try:
        return dt.datetime.strptime(value, "%Y-%m-%d").strftime("%d/%m/%Y")
    except ValueError:
        return value


def _save_upload(file_storage, dest_dir: Path, prefix: str = "") -> Path:
    name = secure_filename(file_storage.filename) or "upload"
    if prefix:
        name = f"{prefix}_{name}"
    dest = dest_dir / name
    file_storage.save(dest)
    return dest


def _cropped_photo_bytes(req, file_field: str, box_field_prefix: str) -> bytes | None:
    """Reads an optional photo upload (`file_field`) plus its
    `{box_field_prefix}_left/top/right/bottom` hidden fields (filled in by
    the crop-overlay JS) and returns the final cropped JPEG bytes. Falls
    back to the full photo if the box fields are missing (JS didn't run),
    and to None if no photo was attached at all or it can't be read as an
    image. Shared by the section 3.2 seal photo and Section 5's per-sample
    Before/After photos - only the field names differ between them."""
    photo = req.files.get(file_field)
    if not photo or not photo.filename:
        return None
    try:
        image = seal_photo.normalize_for_report(Image.open(photo.stream))
    except Exception:
        return None

    raw_box = [req.form.get(f"{box_field_prefix}_{key}", "") for key in ("left", "top", "right", "bottom")]
    try:
        box = tuple(int(float(value)) for value in raw_box) if all(raw_box) else (0, 0, *image.size)
    except ValueError:
        box = (0, 0, *image.size)

    cropped = seal_photo.crop_to_box(image, box)
    return seal_photo.encode_jpeg(cropped)


def _collect_sample_photos(req) -> dict[str, SamplePhotoEntry]:
    """Reads the dynamically-added Section 5 Before/After photo slots the
    detect-samples JS renders (`samplephoto_{idx}_label`, `..._before`,
    `..._after`, `..._before_caption`, plus their crop-box hidden fields)
    into the {label: SamplePhotoEntry} shape `_set_photo_placeholders`
    expects. The After-Test caption isn't read here - the generator
    fills it in itself from that sample's own remarks."""
    photos: dict[str, SamplePhotoEntry] = {}
    idx = 0
    while True:
        label = req.form.get(f"samplephoto_{idx}_label")
        if label is None:
            break
        before = _cropped_photo_bytes(req, f"samplephoto_{idx}_before", f"samplephoto_{idx}_before")
        after = _cropped_photo_bytes(req, f"samplephoto_{idx}_after", f"samplephoto_{idx}_after")
        caption = req.form.get(f"samplephoto_{idx}_before_caption", "").strip()
        if before or after:
            photos[label.strip()] = SamplePhotoEntry(
                before_photo=before, before_caption=caption, after_photo=after
            )
        idx += 1
    return photos


def _has_suffix(file_storage, allowed: set[str]) -> bool:
    if not file_storage or not file_storage.filename:
        return False
    return Path(file_storage.filename).suffix.lower() in allowed


def _mimetype_for(filename: str) -> str:
    if filename.lower().endswith(".pdf"):
        return "application/pdf"
    return "application/vnd.openxmlformats-officedocument.wordprocessingml.document"


def _lan_ip() -> str:
    """The outbound-facing interface IP, for the "share this address with
    your team" banner. Not `socket.gethostbyname(socket.gethostname())` -
    that's unreliable on machines with VPN/Docker Desktop virtual adapters,
    common on corporate laptops. Opening a UDP socket to an external
    address doesn't actually send a packet; it just makes the OS resolve
    which local interface would be used, which is what we want."""
    probe = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        probe.connect(("8.8.8.8", 80))
        return probe.getsockname()[0]
    except OSError:
        return "127.0.0.1"
    finally:
        probe.close()


@app.get("/")
def index():
    port = app.config.get("SERVER_PORT", DEFAULT_PORT)
    return render_template("index.html", lan_url=f"http://{_lan_ip()}:{port}/")


def _project_specification_page(error: str | None = None):
    return render_template(
        "project_specification.html",
        tooling_lead_time_options=TOOLING_LEAD_TIME_OPTIONS,
        today=dt.date.today().strftime("%Y-%m-%d"),
        error=error,
    )


@app.get("/project-specification")
def project_specification_form():
    return _project_specification_page()


@app.post("/project-specification/generate")
def project_specification_generate():
    excel_file = request.files.get("excel_file")
    if not _has_suffix(excel_file, PROJECT_SPEC_ALLOWED_SUFFIXES):
        return _project_specification_page("Please attach a valid Excel file (.xlsm or .xlsx).")

    project_no = request.form.get("project_no", "").strip()
    project_leader = request.form.get("project_leader", "").strip()
    if not project_no:
        return _project_specification_page("Please enter Project No. (e.g., TR26-0002-BTS).")
    if not project_leader:
        return _project_specification_page("Please enter Project Leader name.")

    action = "pdf" if request.form.get("action") == "pdf" else "word"
    report_date = _format_date_for_report(request.form.get("report_date", ""))
    revision_no = request.form.get("revision", "").strip() or "0"
    revision_date = _format_date_for_report(request.form.get("revision_date", ""))
    tooling_lead_time = request.form.get("tooling_lead_time", "Available").strip()
    decision_rule_path = DECISION_RULE_SOURCE_PATH if DECISION_RULE_SOURCE_PATH.exists() else None

    tmp_dir = Path(tempfile.mkdtemp(prefix="skf_projspec_"))
    try:
        excel_path = _save_upload(excel_file, tmp_dir)
        if action == "pdf":
            temp_docx = tmp_dir / "temp_report.docx"
            convert(
                excel_path, PROJECT_SPEC_TEMPLATE_PATH, temp_docx, tmp_dir / "assets",
                report_date=report_date or None,
                revision_no=revision_no or None,
                revision_date=revision_date or None,
                project_no=project_no or None,
                project_leader=project_leader or None,
                tooling_lead_time=tooling_lead_time or None,
                decision_rule_source_path=decision_rule_path,
            )
            out_path = tmp_dir / f"{excel_path.stem} - Generated Report.pdf"
            with _PDF_LOCK:
                convert_docx_to_pdf(temp_docx, out_path)
        else:
            out_path = tmp_dir / f"{excel_path.stem} - Generated Report.docx"
            convert(
                excel_path, PROJECT_SPEC_TEMPLATE_PATH, out_path, tmp_dir,
                report_date=report_date or None,
                revision_no=revision_no or None,
                revision_date=revision_date or None,
                project_no=project_no or None,
                project_leader=project_leader or None,
                tooling_lead_time=tooling_lead_time or None,
                decision_rule_source_path=decision_rule_path,
            )
        data = out_path.read_bytes()
        filename = out_path.name
    except Exception as exc:
        return _project_specification_page(f"Could not generate report.\n\n{exc}")
    finally:
        # mkdtemp() + rmtree(ignore_errors=True) rather than
        # TemporaryDirectory()'s context manager: a transient Windows
        # delete failure (AV scanning the just-written file, a lingering
        # Word/COM handle) must not turn an already-successful generation
        # (bytes already read above) into a 500 for the user.
        shutil.rmtree(tmp_dir, ignore_errors=True)

    return send_file(io.BytesIO(data), as_attachment=True, download_name=filename, mimetype=_mimetype_for(filename))


def _final_test_report_page(error: str | None = None):
    return render_template(
        "final_test_report.html",
        reviewer_roles=REVIEWER_ROLES,
        approver_roles=APPROVER_ROLES,
        today=dt.date.today().strftime("%Y-%m-%d"),
        error=error,
    )


@app.get("/final-test-report")
def final_test_report_form():
    return _final_test_report_page()


@app.post("/final-test-report/seal-photo-suggest")
def final_test_report_seal_photo_suggest():
    """Called client-side right after a seal photo is picked: runs the
    same background-subtraction heuristic used by the desktop app's crop
    dialog and returns a suggested box (original-image pixel coordinates)
    for the browser to draw as a draggable overlay."""
    photo = request.files.get("photo")
    if not photo or not photo.filename:
        return jsonify({"error": "No photo uploaded."}), 400
    try:
        image = seal_photo.normalize_for_report(Image.open(photo.stream))
    except Exception:
        return jsonify({"error": "Could not read that image."}), 400
    box = seal_photo.suggest_crop_box(image)
    width, height = image.size
    return jsonify({"width": width, "height": height, "box": list(box)})


@app.post("/final-test-report/detect-samples")
def final_test_report_detect_samples():
    """Called client-side once both the Project Spec and Test Inspection
    sheet are attached: parses just enough of them to know the sample
    count/labels (same source as Section 2 Conclusions, Section 4.3 Date
    the test was performed, and Section 5's photo table), so the browser
    can render the right number of Before/After photo slots ahead of a
    full report generation."""
    project_spec_file = request.files.get("project_spec")
    test_inspection_file = request.files.get("test_inspection")
    monitoring_files = [f for f in request.files.getlist("monitoring_sheets") if f and f.filename]

    if not _has_suffix(project_spec_file, DOC_ALLOWED_SUFFIXES):
        return jsonify({"error": "Attach a valid Project Specification file (.pdf or .docx) first."}), 400
    if not _has_suffix(test_inspection_file, SHEET_ALLOWED_SUFFIXES):
        return jsonify({"error": "Attach a valid Test Inspection & Execution sheet first."}), 400

    tmp_dir = Path(tempfile.mkdtemp(prefix="skf_detectsamples_"))
    try:
        project_spec_path = _save_upload(project_spec_file, tmp_dir, prefix="projectspec")
        inspection_path = _save_upload(test_inspection_file, tmp_dir, prefix="inspection")
        monitoring_paths = [
            _save_upload(monitoring_file, tmp_dir, prefix=f"monitoring{idx:02d}")
            for idx, monitoring_file in enumerate(monitoring_files, start=1)
        ]
        labels = detect_sample_labels(project_spec_path, inspection_path, monitoring_paths)
    except Exception as exc:
        return jsonify({"error": f"Could not detect samples.\n\n{exc}"}), 400
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)

    return jsonify({"labels": labels})


@app.post("/final-test-report/generate")
def final_test_report_generate():
    project_spec_file = request.files.get("project_spec")
    test_inspection_file = request.files.get("test_inspection")
    monitoring_files = [f for f in request.files.getlist("monitoring_sheets") if f and f.filename]

    if not _has_suffix(project_spec_file, DOC_ALLOWED_SUFFIXES):
        return _final_test_report_page("Please attach a valid Project Specification file (.pdf or .docx).")
    if not _has_suffix(test_inspection_file, SHEET_ALLOWED_SUFFIXES):
        return _final_test_report_page(
            "Please attach a valid Test Inspection & Execution sheet (.pdf, .docx, .xlsx, .xlsm)."
        )
    for monitoring_file in monitoring_files:
        if Path(monitoring_file.filename).suffix.lower() not in SHEET_ALLOWED_SUFFIXES:
            return _final_test_report_page(
                "Please attach valid Monitoring sheet files (.pdf, .docx, .xlsx, .xlsm)."
            )

    distribution = [
        DistributionEntry(name.strip(), org.strip(), location.strip())
        for name, org, location in zip(
            request.form.getlist("distribution_name"),
            request.form.getlist("distribution_org"),
            request.form.getlist("distribution_location"),
        )
    ]

    action = "pdf" if request.form.get("action") == "pdf" else "word"
    tmp_dir = Path(tempfile.mkdtemp(prefix="skf_finalreport_"))
    try:
        project_spec_path = _save_upload(project_spec_file, tmp_dir, prefix="projectspec")
        inspection_path = _save_upload(test_inspection_file, tmp_dir, prefix="inspection")
        monitoring_paths = [
            _save_upload(monitoring_file, tmp_dir, prefix=f"monitoring{idx:02d}")
            for idx, monitoring_file in enumerate(monitoring_files, start=1)
        ]
        seal_photo_bytes = _cropped_photo_bytes(request, "seal_photo", "seal_photo")
        sample_photos = _collect_sample_photos(request)

        inputs = FinalReportInputs(
            project_spec_path=project_spec_path,
            inspection_sheet_path=inspection_path,
            monitoring_sheet_paths=monitoring_paths,
            report_date=_format_date_for_report(request.form.get("report_date", "")),
            project_leader=request.form.get("project_leader", "").strip(),
            reviewer=request.form.get("reviewer", "").strip(),
            approved_by=request.form.get("approved_by", "").strip(),
            reviewer_role=request.form.get("reviewer_role", REVIEWER_ROLES[0]),
            approved_by_role=request.form.get("approved_by_role", APPROVER_ROLES[0]),
            distribution=distribution,
            seal_photo=seal_photo_bytes,
            sample_photos=sample_photos,
        )

        if action == "pdf":
            with _PDF_LOCK:
                out_path = generate_final_report(inputs, "pdf", output_dir=tmp_dir)
        else:
            out_path = generate_final_report(inputs, "word", output_dir=tmp_dir)
        data = out_path.read_bytes()
        filename = out_path.name
    except Exception as exc:
        return _final_test_report_page(f"Could not generate report.\n\n{exc}")
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)

    return send_file(io.BytesIO(data), as_attachment=True, download_name=filename, mimetype=_mimetype_for(filename))


@app.get("/work-instruction")
def work_instruction():
    tmp_dir = Path(tempfile.mkdtemp(prefix="skf_instructions_"))
    try:
        out_path = tmp_dir / "SKF Report Generator - Work Instruction.pdf"
        write_instruction_pdf(out_path)
        data = out_path.read_bytes()
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)
    return send_file(
        io.BytesIO(data),
        as_attachment=True,
        download_name="SKF Report Generator - Work Instruction.pdf",
        mimetype="application/pdf",
    )


def _open_browser_when_ready(port: int) -> None:
    def _open() -> None:
        time.sleep(1.0)
        try:
            webbrowser.open(f"http://127.0.0.1:{port}/")
        except Exception:
            pass

    threading.Thread(target=_open, daemon=True).start()


def main() -> None:
    try:
        port = int(os.environ.get("PORT", DEFAULT_PORT))
    except ValueError:
        port = DEFAULT_PORT
    app.config["SERVER_PORT"] = port

    _open_browser_when_ready(port)

    # Always waitress, never Flask's own dev server (app.run()) - the dev
    # server isn't meant for more than one client at a time and must not
    # ship in the packaged build.
    from waitress import serve

    serve(app, host="0.0.0.0", port=port)


if __name__ == "__main__":
    main()
