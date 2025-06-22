from flask import Flask, request, send_file, jsonify
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.oxml.shared import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_LINE_SPACING
import tempfile
import os
import uuid

app = Flask(__name__)
DOWNLOAD_FOLDER = "/tmp"

@app.route("/download/<filename>", methods=["GET"])
def download_file(filename):
    filepath = os.path.join(DOWNLOAD_FOLDER, filename)
    if os.path.exists(filepath):
        return send_file(
            filepath,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    return {"error": "File not found."}, 404

# This function generates the formatted SOP document using user-provided section data.
def generate_sop_doc(data):
    doc = Document("template.docx")
    doc_section = doc.sections[0]
    doc_section.top_margin = Inches(1)
    doc_section.bottom_margin = Inches(1)
    doc_section.left_margin = Inches(1)
    doc_section.right_margin = Inches(1)
    doc_section.footer_distance = Inches(0.5)
    doc_section.different_first_page_header_footer = True

    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(11)

    def hr():
        para = doc.add_paragraph()
        para.style = doc.styles['Normal']
        para.paragraph_format.space_before = Pt(0)
        para.paragraph_format.space_after = Pt(0)
        para.paragraph_format.line_spacing = 1.0
        p = para._p
        pPr = p.get_or_add_pPr()
        pBdr = OxmlElement('w:pBdr')
        bottom = OxmlElement('w:bottom')
        bottom.set(qn('w:val'), 'single')
        bottom.set(qn('w:sz'), '6')
        bottom.set(qn('w:space'), '0')
        bottom.set(qn('w:color'), 'auto')
        pBdr.append(bottom)
        pPr.append(pBdr)

    def add_paragraph(text, bold=False, size=11, spacing=1.5, indent=None):
        if not text.strip():
            return None
        para = doc.add_paragraph()
        para.paragraph_format.space_before = Pt(0)
        para.paragraph_format.space_after = Pt(0)
        para.paragraph_format.line_spacing = spacing
        if indent is not None:
            para.paragraph_format.left_indent = Inches(indent)
        run = para.add_run(text)
        run.bold = bold
        run.font.size = Pt(size)
        run.font.color.rgb = RGBColor(0, 0, 0)
        return para

    def add_table(section_data):
        table = doc.add_table(rows=1, cols=3)
        table.autofit = False
        widths = [Inches(1.5), Inches(2.0), Inches(2.5)]
        headers = ["Date", "Revised By", "Description"]

        hdr_cells = table.rows[0].cells
        for i, cell in enumerate(hdr_cells):
            run = cell.paragraphs[0].add_run(headers[i])
            run.bold = True
            run.underline = True
            cell.paragraphs[0].paragraph_format.space_after = Pt(0)
            table.columns[i].width = widths[i]

        for row in section_data.get("content", []):
            cells = table.add_row().cells
            parts = row.get("text", "|||").split("|||")
            for i in range(min(3, len(parts))):
                cells[i].text = parts[i].strip()
                cells[i].paragraphs[0].paragraph_format.space_after = Pt(0)

    sop_title = data.get('title', 'Generated SOP')
    sop_id = data.get('sop_id', 'SOP-ID TBD') if not data.get('sop_id') else data.get('sop_id')
    prepared_by = data.get('prepared_by', 'Name')
    approved_by = data.get('approved_by', 'Approver')
    revision_date = data.get('revision_date', 'Date')

    add_paragraph(f"SOP Title: {sop_title}", bold=True, size=18, spacing=1.0)
    add_paragraph(f"SOP ID: {sop_id}", spacing=1.5)
    add_paragraph(f"Prepared By: {prepared_by}", spacing=1.0)
    add_paragraph(f"Approved By: {approved_by}", spacing=1.0)
    add_paragraph(f"Revision Date: {revision_date}", spacing=1.0)
    hr()

    import re

    label_to_indent = {
        r"^1\.$": 0,
        r"^A\.$": 1,
        r"^1\..$": 2,
        r"^a\.$": 3,
        r"^1\...$": 4
    }

    sections_data = data.get("sections", [])
    for i, sec_data in enumerate(sections_data):
        heading = sec_data.get("heading", "")
        if heading:
            add_paragraph(heading, bold=True)

        if sec_data.get("type") == "table":
            add_table(sec_data)
        else:
            last_type = None
            bullets_seen = []
            for idx, item in enumerate(sec_data.get("content", [])):
                if not item.get("text", "").strip():
                    continue
                label = ""
                if item.get("type") == "labelled" and ":" in item.get("text", ""):
                    label, _, _ = item["text"].partition(":")
                    label = label.strip().replace("*", "").rstrip(".")
                text = item.get("text", "")
                t = item.get("type", "text")

                if last_type and last_type != t and label not in ["1.", "A.", "a."]:
                    add_paragraph("", spacing=1.5)
                last_type = t

                if t == "bullet":
                    bullets_seen.append(idx)
                    is_scope_bullet = any("Scope" in sec_data["content"][j]["text"] for j in range(idx) if sec_data["content"][j]["type"] == "labelled")
                    is_role_bullet = any("Role" in sec_data["content"][j]["text"] for j in range(idx) if sec_data["content"][j]["type"] == "labelled")
                    spacing = 1.0 if (is_scope_bullet or is_role_bullet) else 1.5

                    para = doc.add_paragraph()
                    para.paragraph_format.space_before = Pt(0)
                    para.paragraph_format.space_after = Pt(0)
                    para.paragraph_format.line_spacing = spacing
                    run = para.add_run("\u2022 ")
                    run.bold = True
                    run.font.size = Pt(11)
                    run.font.color.rgb = RGBColor(0, 0, 0)
                    para.add_run(text).font.color.rgb = RGBColor(0, 0, 0)
                    para.paragraph_format.left_indent = Inches(0.25)
                elif t == "sub_bullet":
                    para = doc.add_paragraph()
                    run = para.add_run("\u2022 ")
                    run.bold = True
                    run.font.size = Pt(11)
                    run.font.color.rgb = RGBColor(0, 0, 0)
                    para.add_run(text).font.color.rgb = RGBColor(0, 0, 0)
                    para.paragraph_format.left_indent = Inches(0.75)
                elif t == "labelled" and ":" in text:
                    label, _, value = text.partition(":")
                    label = label.strip().replace("*", "")
                    value = value.strip()
                    para = doc.add_paragraph()
                    para.paragraph_format.space_before = Pt(0)
                    para.paragraph_format.space_after = Pt(0)
                    para.paragraph_format.line_spacing = 1.0 if label in ["Scope", "Role", "Output", "Interaction", "Objectives", "Process Owner", "Process Owners"] else 1.5
                    run1 = para.add_run(f"{label}: ")
                    run1.bold = True
                    run1.font.size = Pt(11)
                    run1.font.color.rgb = RGBColor(0, 0, 0)
                    run2 = para.add_run(value)
                    run2.font.size = Pt(11)
                    run2.font.color.rgb = RGBColor(0, 0, 0)
                else:
                    add_paragraph(text)

        if i < len(sections_data) - 1:
            hr()

    footer = doc.sections[0].footer
    footer_para = footer.add_paragraph()
    footer_para.alignment = 1
    run = footer_para.add_run(f"{sop_title} [{sop_id}]")
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor(0, 0, 0)
    para2 = doc.sections[0].footer.add_paragraph()
    para2.alignment = 1
    run2 = para2.add_run(f"Revision Date: {revision_date}")
    run2.font.size = Pt(10)
    run2.font.color.rgb = RGBColor(0, 0, 0)

    filename = f"sop_{uuid.uuid4().hex}.docx"
    filepath = os.path.join(DOWNLOAD_FOLDER, filename)
    doc.save(filepath)
    return filename

@app.route("/generate", methods=["POST"])
def generate():
    try:
        data = request.json
        filename = generate_sop_doc(data)
        link = f"https://sop-flask-api.onrender.com/download/{filename}"
        return jsonify({"download_link": link})
    except Exception as e:
        return {"error": str(e)}, 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
