from __future__ import annotations

from pathlib import Path
from typing import Any

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Inches, Pt

from utils import clean_line, ensure_directory, join_slide_numbers, write_text


PHASE_SECTION_KEYS = [
    "opening_warm_up_launch",
    "mini_lesson_modeling_concept_development",
    "guided_practice_collaborative_learning",
    "independent_practice_application_stations",
    "closure_exit_ticket_assessment",
]


def ensure_template_docx(template_path: Path, config: dict[str, Any]) -> None:
    if template_path.exists():
        return

    ensure_directory(template_path.parent)
    document = Document()
    section = document.sections[0]
    section.top_margin = Inches(0.65)
    section.bottom_margin = Inches(0.65)
    section.left_margin = Inches(0.75)
    section.right_margin = Inches(0.75)
    _apply_base_styles(document)
    if document.paragraphs:
        document.paragraphs[0].text = ""
    document.save(str(template_path))


def render_markdown(lesson_plan: dict[str, Any], template_path: Path, output_path: Path) -> None:
    del template_path
    lines = [
        f"# {lesson_plan['source_deck']['deck_title'] or 'Lesson Plan Package'}",
        "",
        f"- Date: {lesson_plan['date']}",
        f"- Source deck: {lesson_plan['source_deck']['source_filename']}",
        f"- Selected sessions: {', '.join(str(number) for number in lesson_plan['source_deck']['selected_session_numbers'])}",
        "",
    ]
    for session in lesson_plan.get("sessions", []):
        lines.extend(render_session_markdown(session))
        lines.append("")
    write_text(output_path, "\n".join(lines).strip() + "\n")


def render_session_markdown(session: dict[str, Any]) -> list[str]:
    lines = [
        f"## {session['session_label']}: {session['lesson_information']['lesson_title']}",
        "",
        "### 1. Lesson Information",
        f"- Date: {session['lesson_information']['date']}",
        f"- Course/grade: {session['lesson_information']['course_or_grade']}",
        f"- Estimated duration: {session['lesson_information']['estimated_duration_minutes']} minutes",
        "",
        "### 2. Standards and Learning Targets",
    ]
    standards = session["standards_and_learning_targets"]["standards"]
    if standards:
        lines.extend(f"- Standard: {item}" for item in standards)
    else:
        lines.append(f"- {session['standards_and_learning_targets']['standards_status']}")
    lines.extend(f"- Learning target: {item}" for item in session["standards_and_learning_targets"]["learning_targets"])
    lines.extend(
        [
            "",
            "### 3. Lesson Objective and Student Success Criteria",
            f"- Objective: {session['lesson_objective_and_student_success_criteria']['lesson_objective']}",
        ]
    )
    lines.extend(
        f"- Success criteria: {item}"
        for item in session["lesson_objective_and_student_success_criteria"]["student_success_criteria"]
    )
    lines.extend(["", "### 4. Materials and Preparation"])
    lines.extend(f"- Material: {item}" for item in session["materials_and_preparation"]["materials"])
    lines.extend(f"- Preparation: {item}" for item in session["materials_and_preparation"]["preparation_notes"])
    for key in PHASE_SECTION_KEYS:
        phase = session[key]
        lines.extend(["", f"### {phase['section_title']}"])
        lines.extend(f"- Focus task: {item}" for item in phase["focus_tasks"])
        lines.extend(f"- Teacher move: {item}" for item in phase["teacher_actions"])
        lines.extend(f"- Student action: {item}" for item in phase["student_actions"])
        lines.extend(f"- Evidence: {item}" for item in phase["evidence_of_learning"])
    lines.extend(
        [
            "",
            "### 10. Differentiation, SPED/ESOL Supports, and Teacher Notes",
            f"- Implementation note: {session['differentiation_sped_esol_supports_and_teacher_notes']['implementation_note']}",
        ]
    )
    lines.extend(
        f"- SPED: {item['student']}: {', '.join(item['supports'])}"
        for item in session["differentiation_sped_esol_supports_and_teacher_notes"]["sped"]
    )
    lines.extend(
        f"- ESOL: {item}"
        for item in session["differentiation_sped_esol_supports_and_teacher_notes"]["esol"]
    )
    lines.extend(
        f"- Teacher note: {item}"
        for item in session["differentiation_sped_esol_supports_and_teacher_notes"]["teacher_notes"]
    )
    return lines


def render_docx(
    lesson_plan: dict[str, Any],
    template_path: Path,
    output_path: Path,
    config: dict[str, Any],
) -> list[Path]:
    del config
    output_dir = output_path.parent
    ensure_directory(output_dir)
    rendered_paths: list[Path] = []

    for session in lesson_plan.get("sessions", []):
        document = Document(str(template_path))
        _apply_base_styles(document)
        _render_session_doc(document, session)
        session_output = output_dir / session["output_filename"]
        document.save(str(session_output))
        rendered_paths.append(session_output)

    return rendered_paths


def _render_session_doc(document: Document, session: dict[str, Any]) -> None:
    info = session["lesson_information"]
    title = document.add_paragraph(style="Title")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.add_run(info["lesson_title"])

    subtitle = document.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle.add_run(
        f"{session['session_label']} | {info['date']} | {info['course_or_grade']} | "
        f"{info['estimated_duration_minutes']} minutes"
    ).bold = True

    _add_numbered_heading(document, 1, "Lesson Information")
    lesson_table = document.add_table(rows=5, cols=2)
    lesson_table.style = "Table Grid"
    lesson_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    _fill_table_row(lesson_table.rows[0].cells, "Date", info["date"])
    _fill_table_row(lesson_table.rows[1].cells, "Lesson title", info["lesson_title"])
    _fill_table_row(lesson_table.rows[2].cells, "Session number", str(info["session_number"]))
    _fill_table_row(lesson_table.rows[3].cells, "Course/grade", info["course_or_grade"])
    _fill_table_row(lesson_table.rows[4].cells, "Estimated duration", f"{info['estimated_duration_minutes']} minutes")

    _add_numbered_heading(document, 2, "Standards and Learning Targets")
    standards = session["standards_and_learning_targets"]
    _add_label_paragraph(document, "Standards")
    if standards["standards"]:
        _add_bullets(document, standards["standards"])
    else:
        _add_bullets(document, [standards["standards_status"]])
    _add_label_paragraph(document, "Learning Targets")
    _add_bullets(document, standards["learning_targets"])
    if standards["i_can_statements"]:
        _add_label_paragraph(document, "I Can Statements")
        _add_bullets(document, standards["i_can_statements"])

    _add_numbered_heading(document, 3, "Lesson Objective and Student Success Criteria")
    objective = session["lesson_objective_and_student_success_criteria"]
    _add_label_paragraph(document, "Lesson Objective")
    _add_paragraph(document, objective["lesson_objective"])
    _add_label_paragraph(document, "Student Success Criteria")
    _add_bullets(document, objective["student_success_criteria"])

    _add_numbered_heading(document, 4, "Materials and Preparation")
    materials = session["materials_and_preparation"]
    _add_label_paragraph(document, "Materials")
    _add_bullets(document, materials["materials"])
    _add_label_paragraph(document, "Preparation")
    _add_bullets(document, materials["preparation_notes"])
    if materials["required_visuals"]:
        _add_label_paragraph(document, "Required Visuals/Diagrams")
        _add_bullets(document, materials["required_visuals"])

    for index, key in enumerate(PHASE_SECTION_KEYS, start=5):
        _render_phase_section(document, index, session[key])

    _add_numbered_heading(document, 10, "Differentiation, SPED/ESOL Supports, and Teacher Notes")
    supports = session["differentiation_sped_esol_supports_and_teacher_notes"]
    _add_paragraph(document, supports["implementation_note"])
    if supports["sped"]:
        _add_label_paragraph(document, "SPED Supports")
        sped_table = document.add_table(rows=1, cols=2)
        sped_table.style = "Table Grid"
        sped_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        _write_header_row(sped_table.rows[0].cells, ["Student", "Supports"])
        for item in supports["sped"]:
            row = sped_table.add_row().cells
            row[0].text = item["student"]
            row[1].text = ", ".join(item["supports"])
    if supports["esol"]:
        _add_label_paragraph(document, "ESOL Supports")
        _add_bullets(document, supports["esol"])
    if supports["teacher_notes"]:
        _add_label_paragraph(document, "Teacher Notes")
        _add_bullets(document, supports["teacher_notes"])
    if supports["precision_monitoring"]:
        _add_label_paragraph(document, "Precision Monitoring")
        _add_bullets(document, supports["precision_monitoring"])


def _render_phase_section(document: Document, index: int, phase: dict[str, Any]) -> None:
    _add_numbered_heading(document, index, phase["section_title"])
    _add_label_paragraph(document, "Estimated Time")
    _add_paragraph(document, f"{phase['time_minutes']} minutes")
    _add_label_paragraph(document, "Focus Tasks")
    _add_bullets(document, phase["focus_tasks"])
    _add_label_paragraph(document, "Teacher Actions")
    _add_bullets(document, phase["teacher_actions"])
    _add_label_paragraph(document, "Student Actions")
    _add_bullets(document, phase["student_actions"])
    _add_label_paragraph(document, "Evidence of Learning to Watch For")
    _add_bullets(document, phase["evidence_of_learning"])
    if phase["misconceptions_to_monitor"]:
        _add_label_paragraph(document, "Misconceptions / Quick Corrections")
        _add_bullets(document, phase["misconceptions_to_monitor"])
    if phase["embedded_supports"]:
        _add_label_paragraph(document, "Embedded Supports")
        _add_bullets(document, phase["embedded_supports"])
    _add_label_paragraph(document, "Source Slides")
    _add_paragraph(document, join_slide_numbers(phase["source_slide_numbers"]))
    _add_label_paragraph(document, "Source Excerpt")
    _add_paragraph(document, phase["source_excerpt"])


def _apply_base_styles(document: Document) -> None:
    normal_style = document.styles["Normal"]
    normal_style.font.name = "Calibri"
    normal_style._element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
    normal_style.font.size = Pt(11)
    normal_style.paragraph_format.line_spacing = 1.15
    normal_style.paragraph_format.space_after = Pt(4)

    for style_name, size in (("Title", 18), ("Heading 1", 13), ("Heading 2", 11)):
        style = document.styles[style_name]
        style.font.name = "Calibri"
        style._element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
        style.font.size = Pt(size)
        style.font.bold = True


def _add_numbered_heading(document: Document, number: int, text: str) -> None:
    _add_heading(document, f"{number}. {text}", level=1)


def _add_heading(document: Document, text: str, level: int = 1) -> None:
    paragraph = document.add_paragraph(style=f"Heading {level}")
    paragraph.paragraph_format.space_before = Pt(8)
    paragraph.paragraph_format.space_after = Pt(4)
    paragraph.add_run(text)


def _add_label_paragraph(document: Document, text: str) -> None:
    paragraph = document.add_paragraph(style="Normal")
    run = paragraph.add_run(text)
    run.bold = True


def _add_paragraph(document: Document, text: str) -> None:
    paragraph = document.add_paragraph(style="Normal")
    paragraph.add_run(clean_line(text))


def _add_bullets(document: Document, items: list[str]) -> None:
    for item in items:
        paragraph = document.add_paragraph(style="List Bullet")
        paragraph.paragraph_format.line_spacing = 1.15
        paragraph.paragraph_format.space_after = Pt(2)
        paragraph.add_run(clean_line(item))


def _fill_table_row(cells: list[Any], label: str, value: str) -> None:
    cells[0].text = label
    cells[1].text = clean_line(value)
    for paragraph in cells[0].paragraphs:
        for run in paragraph.runs:
            run.bold = True


def _write_header_row(cells: list[Any], labels: list[str]) -> None:
    for index, label in enumerate(labels):
        cells[index].text = label
        for paragraph in cells[index].paragraphs:
            for run in paragraph.runs:
                run.bold = True
