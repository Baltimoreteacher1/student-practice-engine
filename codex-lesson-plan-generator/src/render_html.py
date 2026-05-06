from pathlib import Path
from typing import Any
import html

from render_docx import _build_compact_session_view

def _escape(text: str) -> str:
    return html.escape(str(text))

def _list_to_html_bullets(items: list[str]) -> str:
    if not items:
        return ""
    # Filter out empty items
    valid_items = [i for i in items if i.strip()]
    if not valid_items:
        return ""
    
    html_items = [f"<li>{_escape(item)}</li>" for item in valid_items]
    return f"<ul class='bullet-list'>{''.join(html_items)}</ul>"

def _render_session_html(session: dict[str, Any], config: dict[str, Any]) -> str:
    view = _build_compact_session_view(session, config)
    
    # Build procedure rows
    procedure_html = ""
    for row in view["procedure_rows"]:
        procedure_html += f"""
        <tr>
            <td class="phase-col">{_escape(row['phase_time']).replace(chr(10), '<br>')}</td>
            <td>{_list_to_html_bullets(row['teacher_moves'])}</td>
            <td>{_list_to_html_bullets(row['student_moves'])}</td>
            <td>{_list_to_html_bullets(row['sped_supports'])}</td>
            <td>{_list_to_html_bullets(row['wida_scaffolds'])}</td>
            <td>{_list_to_html_bullets(row['formative_check'])}</td>
        </tr>
        """
        
    # Build standards rows
    standards_headers_html = "".join(f"<th>{_escape(h)}</th>" for h in view["standards_headers"])
    standards_cells_html = "".join(f"<td>{_escape(c)}</td>" for c in view["standards_cells"])
    
    # Build vocabulary rows
    vocab_html = ""
    for row in view["vocabulary_rows"]:
        vocab_html += f"""
        <tr>
            <td>{_escape(row['tier'])}</td>
            <td class="fw-bold">{_escape(row['term'])}</td>
            <td>{_escape(row['spanish'])}</td>
            <td>{_escape(row['definition'])}</td>
            <td>{_escape(row['morphology'])}</td>
            <td>{_escape(row['cognate'])}</td>
            <td>{_escape(row['cross_disciplinary'])}</td>
        </tr>
        """
        
    iep_line_html = f"<div class='iep-line'>{_escape(view['iep_students_line'])}</div>" if view['iep_students_line'] else ""
    
    small_group_html = ""
    sg_section = session.get("small_group_instruction", {})
    if isinstance(sg_section, dict) and sg_section:
        sg_rows = ""
        from render_docx import SMALL_GROUP_RENDER_FIELDS
        for label, key in SMALL_GROUP_RENDER_FIELDS:
            val = str(sg_section.get(key, ""))
            sg_rows += f"<tr><td class='fw-bold' style='width: 25%;'>{_escape(label)}</td><td>{_escape(val)}</td></tr>"
        
        small_group_html = f"""
        <div class="section-title">Small Group Instruction</div>
        <table class="data-table mb-4">
            <tbody>{sg_rows}</tbody>
        </table>
        """
        
    accommodations_html = ""
    if view["accommodations_matrix_rows"]:
        acc_rows = ""
        for row in view["accommodations_matrix_rows"]:
            acc_rows += f"<tr><td class='fw-bold'>{_escape(row['student'])}</td><td>{_escape(row['supports'])}</td></tr>"
        accommodations_html = f"""
        <div class="section-title">IEP Accommodations Matrix</div>
        <table class="data-table">
            <thead>
                <tr>
                    <th>Student</th>
                    <th>Active Accommodations This Lesson</th>
                </tr>
            </thead>
            <tbody>{acc_rows}</tbody>
        </table>
        """

    return f"""
    <div class="session-container">
        <div class="title-line">{_escape(view['title_line'])}</div>
        {iep_line_html}
        
        <div class="section-title">Do Now</div>
        <div class="content-box">{_escape(view['do_now'])}</div>
        
        <div class="section-title">Standards</div>
        <table class="data-table mb-4">
            <thead><tr>{standards_headers_html}</tr></thead>
            <tbody><tr>{standards_cells_html}</tr></tbody>
        </table>
        
        <div class="section-title">Objectives</div>
        <div class="content-box mb-4">
            <div><span class="fw-bold">Content Objective:</span> {_escape(view['content_objective'])}</div>
            <div class="checkboxes">{_escape(view['checkbox_line']).replace('☐', '<span class="box"></span>')}</div>
            <div class="mt-3"><span class="fw-bold">Language Objective:</span> {_escape(view['language_objective'])}</div>
            <div class="checkboxes">{_escape(view['checkbox_line']).replace('☐', '<span class="box"></span>')}</div>
        </div>
        
        <div class="section-title">Lesson Procedures</div>
        <table class="data-table procedure-table mb-4">
            <thead>
                <tr>
                    <th class="phase-col">Phase (Time)</th>
                    <th>Teacher Moves</th>
                    <th>Student Moves</th>
                    <th>SPED Supports</th>
                    <th>ESOL/WIDA Scaffolds</th>
                    <th>Formative Check</th>
                </tr>
            </thead>
            <tbody>{procedure_html}</tbody>
        </table>
        
        {small_group_html}
        
        <div class="section-title">Vocabulary</div>
        <table class="data-table mb-4">
            <thead>
                <tr>
                    <th>Tier</th>
                    <th>Term</th>
                    <th>Spanish</th>
                    <th>Student-Friendly Definition</th>
                    <th>Morphology</th>
                    <th>EN-ES Cognate</th>
                    <th>Cross-Disciplinary</th>
                </tr>
            </thead>
            <tbody>{vocab_html}</tbody>
        </table>
        
        {accommodations_html}
    </div>
    """

def render_html(lesson_plan: dict[str, Any], output_path: Path, config: dict[str, Any]) -> None:
    sessions_html = []
    for session in lesson_plan.get("sessions", []):
        sessions_html.append(_render_session_html(session, config))
    page_title = _escape(config.get("teacher_name", "Teacher"))
        
    full_html = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Lesson Plan - {page_title}</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
    <style>
        :root {{
            --navy: #17324D;
            --teal: #1FA6A2;
            --amber: #F2C15B;
            --bg: #F7F4EC;
            --border: #D1D5DB;
            --text-dark: #1F2937;
            --text-light: #4B5563;
        }}
        body {{
            font-family: 'Inter', sans-serif;
            background-color: var(--bg);
            color: var(--text-dark);
            line-height: 1.5;
            margin: 0;
            padding: 40px;
        }}
        .container {{
            max-width: 1000px;
            margin: 0 auto;
            background: #fff;
            padding: 40px;
            border-radius: 8px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.05);
        }}
        .session-container {{
            margin-bottom: 60px;
            page-break-after: always;
        }}
        .session-container:last-child {{
            page-break-after: auto;
        }}
        .title-line {{
            font-size: 24px;
            font-weight: 700;
            color: var(--navy);
            border-bottom: 3px solid var(--teal);
            padding-bottom: 12px;
            margin-bottom: 8px;
        }}
        .iep-line {{
            font-size: 13px;
            color: var(--text-light);
            font-style: italic;
            margin-bottom: 24px;
        }}
        .section-title {{
            font-size: 18px;
            font-weight: 700;
            color: var(--navy);
            margin-top: 32px;
            margin-bottom: 12px;
            text-transform: uppercase;
            letter-spacing: 0.05em;
        }}
        .content-box {{
            background: var(--bg);
            border-left: 4px solid var(--teal);
            padding: 16px;
            border-radius: 0 8px 8px 0;
        }}
        .data-table {{
            width: 100%;
            border-collapse: collapse;
            font-size: 13px;
        }}
        .data-table th, .data-table td {{
            border: 1px solid var(--border);
            padding: 10px 12px;
            vertical-align: top;
        }}
        .data-table th {{
            background-color: var(--navy);
            color: white;
            font-weight: 600;
            text-align: left;
        }}
        .data-table tr:nth-child(even) {{
            background-color: #f9fafb;
        }}
        .procedure-table td {{
            font-size: 12px;
        }}
        .phase-col {{
            font-weight: 600;
            width: 12%;
            background-color: #f3f4f6;
        }}
        .bullet-list {{
            margin: 0;
            padding-left: 16px;
        }}
        .bullet-list li {{
            margin-bottom: 6px;
        }}
        .bullet-list li:last-child {{
            margin-bottom: 0;
        }}
        .fw-bold {{ font-weight: 600; }}
        .mb-4 {{ margin-bottom: 24px; }}
        .mt-3 {{ margin-top: 16px; }}
        .checkboxes {{
            color: var(--text-light);
            font-size: 13px;
            margin-top: 8px;
            display: flex;
            align-items: center;
            gap: 12px;
        }}
        .box {{
            display: inline-block;
            width: 14px;
            height: 14px;
            border: 1.5px solid var(--text-light);
            border-radius: 2px;
            margin-right: 6px;
            vertical-align: middle;
        }}
        @media print {{
            body {{
                background-color: #fff;
                padding: 0;
            }}
            .container {{
                box-shadow: none;
                padding: 0;
                max-width: 100%;
            }}
            .content-box {{
                border-left-color: var(--teal) !important;
                background-color: var(--bg) !important;
                -webkit-print-color-adjust: exact;
                print-color-adjust: exact;
            }}
            .data-table th {{
                background-color: var(--navy) !important;
                color: white !important;
                -webkit-print-color-adjust: exact;
                print-color-adjust: exact;
            }}
            .phase-col, .data-table tr:nth-child(even) {{
                background-color: #f9fafb !important;
                -webkit-print-color-adjust: exact;
                print-color-adjust: exact;
            }}
            .title-line {{
                border-bottom-color: var(--teal) !important;
            }}
            .session-container {{
                page-break-after: always;
            }}
        }}
    </style>
</head>
<body>
    <div class="container">
        {''.join(sessions_html)}
    </div>
</body>
</html>"""
    
    # Ensure directory exists
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(full_html)
