#!/usr/bin/env python3
"""Generate a human-readable PDF report from an HDsEMG JSON protocol file.

Usage:
    python generate_report.py <path/to/protocol.json>
"""

import argparse
import json
import sys
from datetime import datetime
from pathlib import Path

try:
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_LEFT
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import mm
    from reportlab.platypus import (
        HRFlowable,
        Image as RLImage,
        KeepTogether,
        Paragraph,
        SimpleDocTemplate,
        Spacer,
        Table,
        TableStyle,
    )
except ImportError:
    print("Error: reportlab is required. Install with: pip install reportlab")
    sys.exit(1)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def is_empty(value) -> bool:
    if value is None:
        return True
    if isinstance(value, str) and value.strip() == "":
        return True
    if isinstance(value, (list, dict)) and len(value) == 0:
        return True
    return False


def fmt_datetime(dt_str: str) -> str:
    try:
        dt = datetime.fromisoformat(dt_str)
        return dt.strftime("%d.%m.%Y  %H:%M:%S")
    except Exception:
        return dt_str


def fmt_value(value, field_type: str = "") -> str:
    if isinstance(value, float):
        # Round to 2 decimal places, strip trailing zeros
        formatted = f"{value:.2f}".rstrip("0").rstrip(".")
        return formatted
    return str(value)


def clean_label(label: str) -> str:
    return " ".join(label.split())  # collapse newlines and extra spaces


def find_images(json_path: Path) -> list:
    """Return all JPEG files in the same directory as the JSON, sorted by name."""
    exts = ("*.jpeg", "*.jpg", "*.JPEG", "*.JPG")
    imgs = []
    for ext in exts:
        imgs.extend(json_path.parent.glob(ext))
    return sorted(set(imgs))


def sized_image(path: Path, width_mm: float) -> "RLImage":
    """Return a reportlab Image scaled to width_mm, preserving aspect ratio."""
    img = RLImage(str(path))
    aspect = img.imageHeight / img.imageWidth
    img.drawWidth  = width_mm * mm
    img.drawHeight = width_mm * mm * aspect
    return img


def image_elements(image_paths: list, s, usable_mm: float = 166.0) -> list:
    """Build PDF elements for a list of image paths (up to 2 per row)."""
    if not image_paths:
        return []

    elements = [Paragraph("Electrode Positioning", s["meas_h"])]

    if len(image_paths) == 1:
        elements.append(sized_image(image_paths[0], min(120.0, usable_mm * 0.65)))
        elements.append(Paragraph(image_paths[0].name, s["field_label"]))
    else:
        col_w = (usable_mm - 4.0) / 2  # 2 columns with a small gap
        for i in range(0, len(image_paths), 2):
            pair = image_paths[i : i + 2]
            if len(pair) == 2:
                row_imgs = [sized_image(p, col_w) for p in pair]
                row_caps = [Paragraph(p.name, s["field_label"]) for p in pair]
                t_img = Table(
                    [[row_imgs[0],  row_imgs[1]]],
                    colWidths=[col_w * mm, col_w * mm],
                )
                t_img.setStyle(TableStyle([
                    ("ALIGN",   (0, 0), (-1, -1), "CENTER"),
                    ("VALIGN",  (0, 0), (-1, -1), "TOP"),
                    ("LEFTPADDING",  (0, 0), (-1, -1), 2),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 2),
                ]))
                t_cap = Table(
                    [[row_caps[0], row_caps[1]]],
                    colWidths=[col_w * mm, col_w * mm],
                )
                t_cap.setStyle(TableStyle([
                    ("ALIGN",  (0, 0), (-1, -1), "CENTER"),
                    ("TOPPADDING",    (0, 0), (-1, -1), 1),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                ]))
                elements += [t_img, t_cap]
            else:
                elements.append(sized_image(pair[0], col_w))
                elements.append(Paragraph(pair[0].name, s["field_label"]))

    return elements


# ---------------------------------------------------------------------------
# Style definitions
# ---------------------------------------------------------------------------

def make_styles():
    base = getSampleStyleSheet()

    DARK_BLUE = colors.HexColor("#1a2744")
    MID_BLUE  = colors.HexColor("#2c4a8c")
    GREY      = colors.HexColor("#555555")
    LIGHT_GREY = colors.HexColor("#888888")
    NOTE_BG   = colors.HexColor("#f5f7fa")

    title = ParagraphStyle(
        "RTitle", parent=base["Title"],
        fontSize=20, spaceAfter=2 * mm,
        textColor=DARK_BLUE, leading=24,
    )
    subtitle = ParagraphStyle(
        "RSubtitle", parent=base["Normal"],
        fontSize=10, spaceAfter=4 * mm,
        textColor=GREY,
    )
    section_h = ParagraphStyle(
        "RSectionH", parent=base["Heading1"],
        fontSize=13, spaceBefore=8 * mm, spaceAfter=3 * mm,
        textColor=DARK_BLUE,
    )
    step_h = ParagraphStyle(
        "RStepH", parent=base["Heading2"],
        fontSize=11, spaceBefore=6 * mm, spaceAfter=1 * mm,
        textColor=MID_BLUE,
    )
    step_desc = ParagraphStyle(
        "RStepDesc", parent=base["Normal"],
        fontSize=8.5, spaceAfter=2 * mm,
        textColor=GREY, leading=12,
    )
    timing = ParagraphStyle(
        "RTiming", parent=base["Normal"],
        fontSize=8, spaceAfter=3 * mm,
        textColor=LIGHT_GREY,
    )
    meas_h = ParagraphStyle(
        "RMeasH", parent=base["Heading3"],
        fontSize=9.5, spaceBefore=4 * mm, spaceAfter=1 * mm,
        textColor=DARK_BLUE,
    )
    attempt_label = ParagraphStyle(
        "RAttempt", parent=base["Normal"],
        fontSize=8.5, spaceBefore=2 * mm, spaceAfter=1 * mm,
        textColor=GREY, fontName="Helvetica-Bold",
    )
    field_label = ParagraphStyle(
        "RFieldLabel", parent=base["Normal"],
        fontSize=8.5, textColor=GREY, leading=12,
    )
    field_value = ParagraphStyle(
        "RFieldValue", parent=base["Normal"],
        fontSize=8.5, textColor=colors.black, leading=12,
    )
    meta_key = ParagraphStyle(
        "RMetaKey", parent=base["Normal"],
        fontSize=9, textColor=GREY,
        fontName="Helvetica-Bold", leading=13,
    )
    meta_val = ParagraphStyle(
        "RMetaVal", parent=base["Normal"],
        fontSize=9, textColor=colors.black, leading=13,
    )
    notes = ParagraphStyle(
        "RNotes", parent=base["Normal"],
        fontSize=8.5, spaceBefore=2 * mm, spaceAfter=2 * mm,
        textColor=colors.HexColor("#333333"),
        backColor=NOTE_BG,
        leftIndent=4 * mm, rightIndent=4 * mm,
        borderPad=3,
        leading=12,
    )
    footer = ParagraphStyle(
        "RFooter", parent=base["Normal"],
        fontSize=7, textColor=LIGHT_GREY,
        alignment=TA_CENTER,
    )

    return {
        "title": title, "subtitle": subtitle,
        "section_h": section_h, "step_h": step_h,
        "step_desc": step_desc, "timing": timing,
        "meas_h": meas_h, "attempt_label": attempt_label,
        "field_label": field_label, "field_value": field_value,
        "meta_key": meta_key, "meta_val": meta_val,
        "notes": notes, "footer": footer,
    }


# ---------------------------------------------------------------------------
# Field table builder
# ---------------------------------------------------------------------------

def field_table(rows, s, indent_left=0):
    """Build a two-column label/value table from a list of (label, value) pairs."""
    table_data = [
        [Paragraph(r[0], s["field_label"]), Paragraph(r[1], s["field_value"])]
        for r in rows
    ]
    col_w = [75 * mm - indent_left, None]
    t = Table(table_data, colWidths=col_w)
    t.setStyle(TableStyle([
        ("VALIGN",        (0, 0), (-1, -1), "TOP"),
        ("ROWBACKGROUNDS",(0, 0), (-1, -1),
         [colors.white, colors.HexColor("#f7f9fc")]),
        ("TOPPADDING",    (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("LEFTPADDING",   (0, 0), (-1, -1), 3 + indent_left),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 3),
    ]))
    return t


# ---------------------------------------------------------------------------
# Core builder
# ---------------------------------------------------------------------------

def build_pdf(json_path: Path) -> Path:
    with open(json_path, encoding="utf-8") as f:
        data = json.load(f)

    output_path = json_path.with_suffix(".pdf")

    doc = SimpleDocTemplate(
        str(output_path),
        pagesize=A4,
        rightMargin=22 * mm, leftMargin=22 * mm,
        topMargin=20 * mm, bottomMargin=20 * mm,
    )

    s = make_styles()
    story = []

    images = find_images(json_path)

    declaration = data.get("declaration", {})
    metadata    = data.get("metadata",    {})
    session     = data.get("session",     {})

    # ---- Title ----
    title_text = declaration.get("title") or "HDsEMG Session Report"
    story.append(Paragraph(title_text, s["title"]))

    description = declaration.get("description", "")
    if not is_empty(description):
        story.append(Paragraph(description, s["subtitle"]))

    story.append(HRFlowable(
        width="100%", thickness=1.5,
        color=colors.HexColor("#1a2744"), spaceAfter=5 * mm,
    ))

    # ---- Session overview table ----
    meta_rows = []
    FIELD_MAP = [
        ("pid",           "Participant ID"),
        ("mess_tag",      "Measurement Date"),
        ("session_type",  "Session Type"),
        ("randomization", "Randomization"),
        ("doms_score",    "DOMS Score"),
    ]
    for key, label in FIELD_MAP:
        v = metadata.get(key)
        if not is_empty(v):
            meta_rows.append((label, str(v)))

    SESSION_MAP = [
        ("started_at",          "Started",        fmt_datetime),
        ("ended_at",            "Ended",          fmt_datetime),
        ("duration_formatted",  "Total Duration", lambda x: x),
    ]
    for key, label, fn in SESSION_MAP:
        v = session.get(key)
        if not is_empty(v):
            meta_rows.append((label, fn(v)))

    if meta_rows:
        t = Table(
            [[Paragraph(r[0], s["meta_key"]), Paragraph(r[1], s["meta_val"])]
             for r in meta_rows],
            colWidths=[52 * mm, None],
        )
        t.setStyle(TableStyle([
            ("VALIGN",        (0, 0), (-1, -1), "TOP"),
            ("ROWBACKGROUNDS",(0, 0), (-1, -1),
             [colors.HexColor("#eef2fb"), colors.white]),
            ("TOPPADDING",    (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ("LEFTPADDING",   (0, 0), (-1, -1), 5),
            ("RIGHTPADDING",  (0, 0), (-1, -1), 5),
        ]))
        story.append(t)

    # General notes
    general_notes = metadata.get("notes_general", "")
    if not is_empty(general_notes):
        story.append(Spacer(1, 3 * mm))
        story.append(Paragraph(f"<b>Notes:</b> {general_notes}", s["notes"]))

    # ---- Steps ----
    story.append(Paragraph("Session Steps", s["section_h"]))
    story.append(HRFlowable(
        width="100%", thickness=0.5,
        color=colors.HexColor("#cccccc"), spaceAfter=2 * mm,
    ))

    for step in data.get("steps", []):
        step_elems = []

        num   = step.get("step_number", "")
        title = step.get("title", f"Step {num}")
        desc  = step.get("description", "")

        step_elems.append(Paragraph(f"Step {num}:  {title}", s["step_h"]))

        if not is_empty(desc):
            step_elems.append(Paragraph(desc, s["step_desc"]))

        # Timing line
        timing_parts = []
        if not is_empty(step.get("started_at")):
            timing_parts.append(f"Started: {fmt_datetime(step['started_at'])}")
        if not is_empty(step.get("completed_at")):
            timing_parts.append(f"Completed: {fmt_datetime(step['completed_at'])}")
        if not is_empty(step.get("duration_formatted")):
            dur = step["duration_formatted"]
            exp = step.get("expected_duration_formatted", "")
            timing_parts.append(
                f"Duration: {dur} (expected: {exp})" if not is_empty(exp) else f"Duration: {dur}"
            )
        if timing_parts:
            step_elems.append(Paragraph("  ·  ".join(timing_parts), s["timing"]))

        # Simple fields
        field_rows = []
        for fid, field in step.get("fields", {}).items():
            v = field.get("value")
            if is_empty(v):
                continue
            label = clean_label(field.get("label", fid))
            field_rows.append((label, fmt_value(v, field.get("type", ""))))

        if field_rows:
            step_elems.append(field_table(field_rows, s))

        # Repeated measurements
        for meas_id, meas in step.get("repeated_measurements", {}).items():
            meas_label = meas.get("label", meas_id)
            attempts   = meas.get("attempts", [])

            # Collect only non-empty attempts
            non_empty_attempts = []
            for attempt in attempts:
                afields = attempt.get("fields", {})
                rows    = []
                anotes  = ""
                otb     = attempt.get("otbiolab_file")

                for fid, field in afields.items():
                    if fid == "notes":
                        anotes = field.get("value", "")
                        continue
                    v = field.get("value")
                    if not is_empty(v):
                        rows.append((
                            clean_label(field.get("label", fid)),
                            fmt_value(v, field.get("type", "")),
                        ))

                has_content = rows or not is_empty(anotes) or not is_empty(otb)
                if has_content:
                    non_empty_attempts.append((attempt, rows, anotes, otb))

            if not non_empty_attempts:
                continue

            step_elems.append(Paragraph(meas_label, s["meas_h"]))

            for attempt, rows, anotes, otb in non_empty_attempts:
                a_num = attempt.get("attempt_number", "")
                step_elems.append(Paragraph(f"Attempt {a_num}", s["attempt_label"]))

                if rows:
                    step_elems.append(field_table(rows, s, indent_left=4))

                if not is_empty(otb):
                    step_elems.append(
                        Paragraph(f"File: {Path(otb).name}", s["field_label"])
                    )

                if not is_empty(anotes):
                    step_elems.append(
                        Paragraph(f"<b>Notes:</b> {anotes}", s["notes"])
                    )

        # Step-level notes
        step_notes = step.get("notes", "")
        if not is_empty(step_notes):
            step_elems.append(Paragraph(f"<b>Notes:</b> {step_notes}", s["notes"]))

        # Electrode positioning photos — injected into step 1 only
        if step.get("step_number") == 1 and images:
            step_elems.append(Spacer(1, 3 * mm))
            step_elems.extend(image_elements(images, s))

        step_elems.append(Spacer(1, 2 * mm))
        step_elems.append(HRFlowable(
            width="100%", thickness=0.3,
            color=colors.HexColor("#dddddd"),
        ))

        # Keep step header + first content together to avoid orphaned headers
        if len(step_elems) > 2:
            story.append(KeepTogether(step_elems[:4]))
            story.extend(step_elems[4:])
        else:
            story.extend(step_elems)

    # ---- Footer ----
    story.append(Spacer(1, 8 * mm))
    story.append(HRFlowable(
        width="100%", thickness=0.5,
        color=colors.HexColor("#aaaaaa"), spaceAfter=2 * mm,
    ))
    gen_time = datetime.now().strftime("%d.%m.%Y %H:%M")
    story.append(Paragraph(
        f"Generated: {gen_time}  ·  Source: {json_path.name}"
        f"  ·  Protocol v{data.get('protocol_version', 'N/A')}",
        s["footer"],
    ))

    doc.build(story)
    return output_path


# ---------------------------------------------------------------------------
# Batch processing
# ---------------------------------------------------------------------------

def batch_process(root: Path) -> None:
    """Find all *.json files inside any protokolle/ folder under root and generate PDFs."""
    json_files = sorted(root.glob("**/protokolle/*.json"))

    if not json_files:
        print(f"No JSON files found under protokolle/ directories in: {root}")
        return

    ok = 0
    failed = 0

    print(f"Found {len(json_files)} file(s) to process.\n")

    for json_path in json_files:
        rel = json_path.relative_to(root)
        print(f"  [{ok + failed + 1}/{len(json_files)}]  {rel}")
        try:
            output = build_pdf(json_path)
            print(f"           -> {output.relative_to(root)}")
            ok += 1
        except Exception as exc:
            print(f"           !! ERROR: {exc}")
            failed += 1

    print(f"\nDone — {ok} succeeded, {failed} failed.")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description=(
            "Generate PDF report(s) from HDsEMG JSON protocol file(s).\n\n"
            "Single file:  python generate_report.py path/to/protocol.json\n"
            "Batch mode:   python generate_report.py path/to/root_folder\n"
            "  (processes every *.json inside any protokolle/ subdirectory)"
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "path",
        help="Path to a JSON file (single) or a root directory (batch)",
    )
    args = parser.parse_args()

    target = Path(args.path).resolve()

    if not target.exists():
        print(f"Error: path not found: {target}")
        sys.exit(1)

    if target.is_dir():
        print(f"Batch mode — scanning: {target}\n")
        batch_process(target)
    else:
        print(f"Generating PDF report for: {target.name} ...")
        output = build_pdf(target)
        print(f"Report saved to:           {output}")


if __name__ == "__main__":
    main()
