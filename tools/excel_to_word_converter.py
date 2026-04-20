#!/usr/bin/env python3
"""Convert a filled SKF Excel test request into a project specification Word document."""

from __future__ import annotations

import argparse
import datetime as dt
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, Optional, Tuple

import openpyxl
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches
from openpyxl.utils import column_index_from_string, get_column_letter, range_boundaries
from PIL import Image, ImageDraw, ImageFont


DEFAULT_FONT_PATHS = [
    "/System/Library/Fonts/Supplemental/Arial.ttf",
    "/System/Library/Fonts/Supplemental/Calibri.ttf",
]


@dataclass
class ExtractedData:
    requester: str
    customer: str
    application: str
    purpose: str
    num_samples: str
    part_number: str
    test_type: str

    shaft_diameter: str
    shaft_tolerance: str
    shaft_unit: str
    shaft_material: str
    shaft_ra: str
    shaft_hardness: str

    bore_diameter: str
    bore_tolerance: str
    bore_unit: str
    bore_material: str
    bore_ra: str

    dro_value: str
    dro_tolerance: str
    stbm_value: str
    stbm_tolerance: str
    seal_cock_value: str
    seal_cock_tolerance: str
    recip_value: str
    recip_tolerance: str

    oil_type: str
    pre_lube_required: str
    setup_notes: str

    contamination_type: str
    contamination_mix_ratio: str
    contamination_amount: str
    contamination_recip_freq: str
    contamination_stbm_orientation: str

    oil_change_interval: str
    acceptance_duration: str
    acceptance_failure: str


class TemplateUpdateError(RuntimeError):
    pass


def _load_font(size: int) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    for font_path in DEFAULT_FONT_PATHS:
        candidate = Path(font_path)
        if candidate.exists():
            return ImageFont.truetype(str(candidate), size)
    return ImageFont.load_default()


def _fmt_num(value: Any, decimals: int = 3) -> str:
    if value is None or value == "":
        return ""
    if isinstance(value, str):
        text = value.strip()
        return text
    if isinstance(value, (int, float)):
        return f"{float(value):.{decimals}f}"
    return str(value)


def _fmt_general(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, dt.datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, float):
        text = f"{value:.6f}".rstrip("0").rstrip(".")
        return text if text else "0"
    return str(value).strip()


def _cell(ws: openpyxl.worksheet.worksheet.Worksheet, ref: str, default: str = "") -> str:
    value = ws[ref].value
    if value is None:
        return default
    return _fmt_general(value)


def _normalize_mix_ratio(value: str) -> str:
    text = (value or "").strip()
    if not text:
        return text
    parts = text.split(":")
    if len(parts) == 3:
        h, m, s = parts
        if s in {"00", "0"}:
            try:
                return f"{int(h)}:{int(m)}"
            except ValueError:
                return text
    return text


def _normalize_ra(value: str) -> str:
    text = (value or "").strip()
    if not text:
        return ""
    lowered = text.lower()
    if lowered.startswith("ra"):
        text = text[2:].strip()
    text = text.replace("-", "~")
    text = " ".join(text.split())
    return f"{text} µm Ra" if text else ""


def _is_contamination_test(test_type: str, purpose: str) -> bool:
    t = (test_type or "").lower()
    p = (purpose or "").lower()
    triggers = ["mud", "slurry", "dry dust", "contamination"]
    return any(token in t for token in triggers) or any(token in p for token in triggers)


def extract_excel_data(excel_path: Path) -> ExtractedData:
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws1 = wb["Page 1"]
    ws2 = wb["Page 2"]

    test_type = _cell(ws1, "K10")
    purpose = _cell(ws1, "C11")

    return ExtractedData(
        requester=_cell(ws1, "C4"),
        customer=_cell(ws1, "G4"),
        application=_cell(ws1, "C10"),
        purpose=purpose,
        num_samples=_cell(ws1, "L6"),
        part_number=_cell(ws1, "C6"),
        test_type=test_type,
        shaft_diameter=_fmt_num(ws1["C18"].value, 3),
        shaft_tolerance=_fmt_num(ws1["F18"].value, 3),
        shaft_unit=_cell(ws1, "D18", "mm"),
        shaft_material=_cell(ws1, "H18"),
        shaft_ra=_cell(ws1, "K18"),
        shaft_hardness=_cell(ws1, "L18"),
        bore_diameter=_fmt_num(ws1["C20"].value, 3),
        bore_tolerance=_fmt_num(ws1["F20"].value, 3),
        bore_unit=_cell(ws1, "D20", "mm"),
        bore_material=_cell(ws1, "H20"),
        bore_ra=_cell(ws1, "K20"),
        dro_value=_fmt_num(ws1["D50"].value, 3),
        dro_tolerance=_fmt_num(ws1["G50"].value, 3),
        stbm_value=_fmt_num(ws1["D52"].value, 3),
        stbm_tolerance=_fmt_num(ws1["G52"].value, 3),
        seal_cock_value=_fmt_num(ws1["D54"].value, 3),
        seal_cock_tolerance=_fmt_num(ws1["G54"].value, 3),
        recip_value=_fmt_num(ws1["D56"].value, 3),
        recip_tolerance=_fmt_num(ws1["G56"].value, 3),
        oil_type=_cell(ws1, "B48"),
        pre_lube_required=_cell(ws1, "J48"),
        setup_notes=_cell(ws1, "B61"),
        contamination_type=_cell(ws1, "J50"),
        contamination_mix_ratio=_normalize_mix_ratio(_cell(ws1, "J52")),
        contamination_amount=_cell(ws1, "J54"),
        contamination_recip_freq=_cell(ws1, "J56"),
        contamination_stbm_orientation=_cell(ws1, "J58"),
        oil_change_interval=_cell(ws2, "E4"),
        acceptance_duration=_cell(ws2, "G5"),
        acceptance_failure=_cell(ws2, "G6"),
    )


def _safe_text(value: Any) -> str:
    text = _fmt_general(value)
    return text.replace("\n", " ").strip()


def _iter_nonempty_rows(
    ws: openpyxl.worksheet.worksheet.Worksheet, min_row: int, max_row: int, min_col: int, max_col: int
) -> Iterable[int]:
    for row in range(min_row, max_row + 1):
        has_any = False
        for col in range(min_col, max_col + 1):
            if _safe_text(ws.cell(row, col).value):
                has_any = True
                break
        if has_any:
            yield row


def render_sheet_range_to_image(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    range_ref: str,
    out_path: Path,
    end_at_last_content: bool = False,
) -> None:
    min_col, min_row, max_col, max_row = range_boundaries(range_ref)

    if end_at_last_content:
        rows = list(_iter_nonempty_rows(ws, min_row, max_row, min_col, max_col))
        if rows:
            max_row = max(rows)

    default_col_width = 8.43
    default_row_height = 15.0

    col_widths = []
    for col in range(min_col, max_col + 1):
        width = ws.column_dimensions[get_column_letter(col)].width
        width = width if width else default_col_width
        col_widths.append(max(54, int(width * 7 + 10)))

    row_heights = []
    for row in range(min_row, max_row + 1):
        height = ws.row_dimensions[row].height
        height = height if height else default_row_height
        row_heights.append(max(24, int(height * 96 / 72)))

    margin_x = 18
    margin_y = 18

    total_width = sum(col_widths) + margin_x * 2
    total_height = sum(row_heights) + margin_y * 2

    img = Image.new("RGB", (total_width, total_height), color="white")
    draw = ImageDraw.Draw(img)
    font = _load_font(12)

    merged_top_left: Dict[Tuple[int, int], Tuple[int, int]] = {}
    merged_covered: Dict[Tuple[int, int], Tuple[int, int]] = {}
    for merged in ws.merged_cells.ranges:
        m_min_col, m_min_row, m_max_col, m_max_row = range_boundaries(str(merged))
        if m_max_col < min_col or m_min_col > max_col or m_max_row < min_row or m_min_row > max_row:
            continue
        top_left = (m_min_row, m_min_col)
        merged_top_left[top_left] = (m_max_row, m_max_col)
        for rr in range(m_min_row, m_max_row + 1):
            for cc in range(m_min_col, m_max_col + 1):
                if (rr, cc) != top_left:
                    merged_covered[(rr, cc)] = top_left

    x_offsets = [margin_x]
    for w in col_widths:
        x_offsets.append(x_offsets[-1] + w)

    y_offsets = [margin_y]
    for h in row_heights:
        y_offsets.append(y_offsets[-1] + h)

    header_fill = (235, 240, 245)
    grid_color = (120, 120, 120)
    text_color = (20, 20, 20)

    for r_idx, row in enumerate(range(min_row, max_row + 1)):
        for c_idx, col in enumerate(range(min_col, max_col + 1)):
            if (row, col) in merged_covered:
                continue

            x0 = x_offsets[c_idx]
            y0 = y_offsets[r_idx]
            x1 = x_offsets[c_idx + 1]
            y1 = y_offsets[r_idx + 1]

            if (row, col) in merged_top_left:
                m_max_row, m_max_col = merged_top_left[(row, col)]
                m_max_row = min(m_max_row, max_row)
                m_max_col = min(m_max_col, max_col)
                m_r_idx = m_max_row - min_row + 1
                m_c_idx = m_max_col - min_col + 1
                x1 = x_offsets[m_c_idx]
                y1 = y_offsets[m_r_idx]

            if row <= min_row + 1:
                draw.rectangle([x0, y0, x1, y1], fill=header_fill)

            draw.rectangle([x0, y0, x1, y1], outline=grid_color, width=1)

            value = _safe_text(ws.cell(row, col).value)
            if value:
                max_text_width = max(20, x1 - x0 - 8)
                words = value.split(" ")
                lines = []
                current = ""
                for word in words:
                    test = f"{current} {word}".strip()
                    bbox = draw.textbbox((0, 0), test, font=font)
                    if bbox[2] - bbox[0] <= max_text_width:
                        current = test
                    else:
                        if current:
                            lines.append(current)
                        current = word
                if current:
                    lines.append(current)

                for line_idx, line in enumerate(lines[:5]):
                    ty = y0 + 4 + line_idx * 14
                    if ty + 12 < y1:
                        draw.text((x0 + 4, ty), line, fill=text_color, font=font)

    if total_width > 1600:
        ratio = 1600 / total_width
        img = img.resize((1600, int(total_height * ratio)), Image.Resampling.LANCZOS)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    img.save(out_path)


def _set_paragraph_text(paragraph, new_text: str) -> None:
    if not paragraph.runs:
        paragraph.add_run(new_text)
        return
    paragraph.runs[0].text = new_text
    for run in paragraph.runs[1:]:
        run.text = ""


def _find_paragraph_by_prefix(doc: Document, prefix: str, occurrence: int = 1):
    count = 0
    for p in doc.paragraphs:
        if p.text.strip().startswith(prefix):
            count += 1
            if count == occurrence:
                return p
    raise TemplateUpdateError(f"Could not find paragraph starting with: {prefix!r}")


def _replace_image_after_anchor(doc: Document, anchor_text: str, image_path: Path, width_inches: float = 6.5) -> None:
    anchor_index: Optional[int] = None
    for idx, para in enumerate(doc.paragraphs):
        if anchor_text in para.text:
            anchor_index = idx
            break

    if anchor_index is None:
        raise TemplateUpdateError(f"Could not locate anchor text: {anchor_text!r}")

    target = None
    for idx in range(anchor_index + 1, len(doc.paragraphs)):
        para = doc.paragraphs[idx]
        has_drawing = any("drawing" in run._element.xml for run in para.runs)
        if has_drawing:
            target = para
            break

    if target is None:
        target = doc.paragraphs[anchor_index + 1]

    for run in list(target.runs):
        target._p.remove(run._element)

    run = target.add_run()
    run.add_picture(str(image_path), width=Inches(width_inches))
    target.alignment = WD_ALIGN_PARAGRAPH.CENTER


def _replace_acceptance_line(doc: Document, acceptance_text: str) -> None:
    anchor_index: Optional[int] = None
    for idx, para in enumerate(doc.paragraphs):
        if para.text.strip().startswith("Acceptance Criteria"):
            anchor_index = idx
            break
    if anchor_index is None:
        raise TemplateUpdateError("Could not find 'Acceptance Criteria' section")

    for idx in range(anchor_index + 1, len(doc.paragraphs)):
        para = doc.paragraphs[idx]
        if para.text.strip():
            _set_paragraph_text(para, acceptance_text)
            return

    raise TemplateUpdateError("Could not find acceptance criteria value line")


def _fill_first_empty_between(doc: Document, start_prefix: str, end_prefix: str, text: str) -> None:
    start_index: Optional[int] = None
    end_index: Optional[int] = None

    for idx, para in enumerate(doc.paragraphs):
        stripped = para.text.strip()
        if start_index is None and stripped.startswith(start_prefix):
            start_index = idx
        if stripped.startswith(end_prefix):
            end_index = idx
            if start_index is not None and end_index > start_index:
                break

    if start_index is None or end_index is None or end_index <= start_index:
        return

    for idx in range(start_index + 1, end_index):
        para = doc.paragraphs[idx]
        if not para.text.strip():
            _set_paragraph_text(para, text)
            return


def update_word_template(
    template_path: Path,
    output_path: Path,
    data: ExtractedData,
    pre_post_image_path: Path,
    duty_cycle_image_path: Path,
) -> None:
    doc = Document(template_path)

    part_compact = data.part_number.replace(" ", "")
    title = f"{part_compact} Seal  {data.purpose} with Oil fill".strip()

    if len(doc.tables) >= 2:
        title_table = doc.tables[1]
        title_table.rows[0].cells[1].text = title
        title_table.rows[0].cells[2].text = title
        title_table.rows[0].cells[3].text = title
        title_table.rows[1].cells[1].text = data.requester

    objective = (
        f"To Conduct the {data.purpose} with Oil fill on {data.num_samples} "
        f"samples of {part_compact} shaft seal."
    )
    _set_paragraph_text(_find_paragraph_by_prefix(doc, "To Conduct the"), objective)

    _set_paragraph_text(_find_paragraph_by_prefix(doc, "Shaft:"), f"Shaft: {data.shaft_material}")
    _set_paragraph_text(
        _find_paragraph_by_prefix(doc, "Diameter:", occurrence=1),
        f"Diameter: {data.shaft_diameter} ± {data.shaft_tolerance} {data.shaft_unit}",
    )
    _set_paragraph_text(
        _find_paragraph_by_prefix(doc, "Surface Finish:", occurrence=1),
        f"Surface Finish: {_normalize_ra(data.shaft_ra)}",
    )
    _set_paragraph_text(
        _find_paragraph_by_prefix(doc, "Hardness:"),
        f"Hardness: {data.shaft_hardness.replace('HRc', 'HRC')} Min.",
    )
    _set_paragraph_text(
        _find_paragraph_by_prefix(doc, "DRO:"),
        f"DRO: {data.dro_value} ± {data.dro_tolerance} {data.shaft_unit} TIR",
    )

    _set_paragraph_text(_find_paragraph_by_prefix(doc, "Bore:"), f"Bore: {data.bore_material}")
    _set_paragraph_text(
        _find_paragraph_by_prefix(doc, "Diameter:", occurrence=2),
        f"Diameter: {data.bore_diameter} ± {data.bore_tolerance}{data.bore_unit}",
    )
    _set_paragraph_text(
        _find_paragraph_by_prefix(doc, "STBM:"),
        f"STBM: {data.stbm_value} ± {data.stbm_tolerance} {data.bore_unit} TIR",
    )
    _set_paragraph_text(
        _find_paragraph_by_prefix(doc, "Surface Finish:", occurrence=2),
        f"Surface Finish: {_normalize_ra(data.bore_ra)}",
    )
    _set_paragraph_text(
        _find_paragraph_by_prefix(doc, "STBM Orientation:"),
        f"STBM Orientation: {data.contamination_stbm_orientation}",
    )
    _set_paragraph_text(
        _find_paragraph_by_prefix(doc, "Seals cock:"),
        f"Seals cock: {data.seal_cock_value} ± {data.seal_cock_tolerance} {data.shaft_unit}",
    )
    _fill_first_empty_between(
        doc,
        "Seals cock:",
        "Fluid:",
        f"Reciprocation: {data.recip_value} ± {data.recip_tolerance} {data.shaft_unit}",
    )

    customer_caps = data.customer.upper() if data.customer else ""
    fluid_type = f"{data.oil_type} {customer_caps}".strip()
    _set_paragraph_text(_find_paragraph_by_prefix(doc, "Type –"), f"Type – {fluid_type}")
    _set_paragraph_text(
        _find_paragraph_by_prefix(doc, "Oil Change Interval:"),
        f"Oil Change Interval: {data.oil_change_interval}",
    )

    pre_lube_value = data.setup_notes or data.pre_lube_required
    _set_paragraph_text(_find_paragraph_by_prefix(doc, "Pre-lube:"), f"Pre-lube: {pre_lube_value}")

    if _is_contamination_test(data.test_type, data.purpose):
        _set_paragraph_text(_find_paragraph_by_prefix(doc, "Type:", occurrence=1), f"Type: {data.contamination_type}")
        _set_paragraph_text(
            _find_paragraph_by_prefix(doc, "Mix Ratio:"),
            f"Mix Ratio: {data.contamination_mix_ratio}",
        )
        _set_paragraph_text(_find_paragraph_by_prefix(doc, "Amount:"), f"Amount: {data.contamination_amount}")
        _fill_first_empty_between(
            doc,
            "Amount:",
            "Test Procedure:",
            f"Recip. Frequency: {data.contamination_recip_freq}",
        )

    duration_clean = data.acceptance_duration.strip()
    _set_paragraph_text(_find_paragraph_by_prefix(doc, "The total test duration is"), f"The total test duration is {duration_clean}")
    _replace_acceptance_line(doc, data.acceptance_failure.strip())

    _replace_image_after_anchor(
        doc,
        "Pre and Post-test measurements of Seals are as follows",
        pre_post_image_path,
        width_inches=6.4,
    )
    _replace_image_after_anchor(doc, "Test cycle would be as follows:", duty_cycle_image_path, width_inches=7.0)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(output_path)


def convert(excel_path: Path, template_docx_path: Path, output_docx_path: Path, temp_dir: Path) -> None:
    data = extract_excel_data(excel_path)

    wb = openpyxl.load_workbook(excel_path, data_only=True)
    page1 = wb["Page 1"]
    page2 = wb["Page 2"]

    pre_post_img = temp_dir / "pre_post_measurements.png"
    duty_cycle_img = temp_dir / "duty_cycle.png"

    render_sheet_range_to_image(page1, "A30:L44", pre_post_img, end_at_last_content=False)
    render_sheet_range_to_image(page2, "A8:O33", duty_cycle_img, end_at_last_content=True)

    update_word_template(template_docx_path, output_docx_path, data, pre_post_img, duty_cycle_img)


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Convert SKF request Excel sheet into Word project specification based on a template.",
    )
    parser.add_argument("--excel", required=True, type=Path, help="Path to source .xlsm file")
    parser.add_argument("--template", required=True, type=Path, help="Path to template .docx file")
    parser.add_argument("--output", required=True, type=Path, help="Path for generated .docx")
    parser.add_argument(
        "--temp-dir",
        type=Path,
        default=Path("output") / "generated_assets",
        help="Folder for intermediate generated images",
    )
    return parser.parse_args()


def main() -> None:
    args = _parse_args()
    convert(args.excel, args.template, args.output, args.temp_dir)
    print(f"Generated Word document: {args.output}")


if __name__ == "__main__":
    main()
