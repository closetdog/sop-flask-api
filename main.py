from flask import Flask, request, send_file, jsonify
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.oxml.shared import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_LINE_SPACING
from docx.oxml.ns import qn as ns_qn
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

def generate_sop_doc(data):
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)
    section.footer_distance = Inches(0.5)

    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(11)

    def hr():
        para = doc.add_paragraph()
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
        para = doc.add_paragraph()
        run = para.add_run(text)
        run.bold = bold
        run.font.size = Pt(size)
        run.font.color.rgb = RGBColor(0, 0, 0)
        para.paragraph_format.line_spacing = spacing
        if indent is not None:
            para.paragraph_format.left_indent = Inches(indent)
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
            table.columns[i].width = widths[i]

        for row in section_data.get("content", []):
            cells = table.add_row().cells
            parts = row.get("text", "|||").split("|||")
            for i in range(min(3, len(parts))):
                cells[i].text = parts[i].strip()

    sop_title = data.get('title', 'Generated SOP')
    sop_id = data.get('sop_id', 'SOP-000')
    prepared_by = data.get('prepared_by', 'Name')
    approved_by = data.get('approved_by', 'Approver')
    revision_date = data.get('revision_date', 'Date')

    add_paragraph(f"SOP Title: {sop_title}", bold=True, size=18)
    add_paragraph(f"SOP ID: {sop_id}")
    add_paragraph(f"Prepared By: {prepared_by}", spacing=1.5)
    add_paragraph(f"Approved By: {approved_by}", spacing=1.5)
    add_paragraph(f"Revision Date: {revision_date}", spacing=1.5)
    hr()

    sections_data = data.get("sections", [])
    for i, section in enumerate(sections_data):
        heading = section.get("heading", "")
        if heading:
            add_paragraph(heading, bold=True)

        if section.get("type") == "table":
            add_table(section)
        else:
            for item in section.get("content", []):
                text = item.get("text", "")
                t = item.get("type", "text")
                indent_level = item.get("indent_level", 0)
                indent_inches = 0.5 + 0.25 * indent_level if indent_level > 0 else None

                if t == "bullet":
                    add_paragraph(f"\u2022 {text}", bold=True, indent=indent_inches)
                elif t == "sub_bullet":
                    add_paragraph(f"\u2022 {text}", bold=True, indent=indent_inches)
                elif t == "labelled":
                    label, _, value = text.partition(":")
                    para = doc.add_paragraph()
                    run1 = para.add_run(f"{label.strip()}: ")
                    run1.bold = True
                    run1.font.size = Pt(11)
                    run1.font.color.rgb = RGBColor(0, 0, 0)
                    run2 = para.add_run(value.strip())
                    run2.font.size = Pt(11)
                    run2.font.color.rgb = RGBColor(0, 0, 0)
                else:
                    add_paragraph(text)

        if i < len(sections_data) - 1:
            hr()

    # Footer (Page 2+ only)
    footer = section.footer
    footer_para = footer.paragraphs[0]
    run = footer_para.add_run(f"{sop_title}\n{sop_id}")
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor(0, 0, 0)
    footer_para.alignment = 0

    right_footer = footer.add_paragraph()
    run = right_footer.add_run(f"Revision Date: {revision_date}")
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor(0, 0, 0)
    right_footer.alignment = 2

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
