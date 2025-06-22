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

    def add_bullet(text, spacing=1.0):
        para = doc.add_paragraph(style='List Bullet')
        para.paragraph_format.left_indent = Inches(0.25)
        para.paragraph_format.space_before = Pt(0)
        para.paragraph_format.space_after = Pt(0)
        para.paragraph_format.line_spacing = spacing
        run = para.runs[0] if para.runs else para.add_run()
        run.text = text
        run.font.size = Pt(11)
        run.font.color.rgb = RGBColor(0, 0, 0)
        return para

    sop_title = data.get('title', 'Generated SOP')
    sop_id = data.get('sop_id', 'SOP-000')
    prepared_by = data.get('prepared_by', 'Name')
    approved_by = data.get('approved_by', 'Approver')
    revision_date = data.get('revision_date', 'Date')

    add_paragraph(f"SOP Title: {sop_title}", bold=True, size=18, spacing=1.0)
    add_paragraph(f"SOP ID: {sop_id}", spacing=1.5)
    add_paragraph(f"Prepared By: {prepared_by}", spacing=1.0)
    add_paragraph(f"Approved By: {approved_by}", spacing=1.0)
    add_paragraph(f"Revision Date: {revision_date}", spacing=1.0)
    hr()

    sections = data.get("sections", [])
    for section in sections:
        heading = section.get("heading", "")
        if heading:
            add_paragraph(heading, bold=True, spacing=1.5)

        content = section.get("content", [])
        bullets = []
        current_label = ""
        for idx, item in enumerate(content):
            text = item.get("text", "")
            type_ = item.get("type", "text")

            if type_ == "labelled":
                if bullets:
                    for i, b in enumerate(bullets):
                        spacing = 1.5 if i == len(bullets) - 1 and current_label in ["Objectives:", "Process Owners:"] else 1.0
                        add_bullet(b, spacing=spacing)
                    bullets.clear()
                current_label = text.strip()
                spacing = 1.0
                add_paragraph(current_label, bold=True, spacing=spacing)

            elif type_ == "bullet":
                bullets.append(text)

        if bullets:
            for i, b in enumerate(bullets):
                spacing = 1.5 if i == len(bullets) - 1 and current_label in ["Objectives:", "Process Owners:"] else 1.0
                add_bullet(b, spacing=spacing)
            bullets.clear()

        hr()

    # Footer
    footer = doc.sections[0].footer
    para = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    para.alignment = 1
    run1 = para.add_run(f"{sop_title} [{sop_id if sop_id else 'SOP-ID TBD'}]")
    run1.font.size = Pt(10)
    run1.font.color.rgb = RGBColor(0, 0, 0)
    para.add_run("\n")
    run2 = para.add_run(f"Revision Date: {revision_date}")
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
