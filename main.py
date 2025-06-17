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

def generate_sop_doc(data):
    doc = Document()
    sections = doc.sections[0]
    sections.top_margin = Inches(1)
    sections.bottom_margin = Inches(1)
    sections.left_margin = Inches(1)
    sections.right_margin = Inches(1)
    doc.styles['Normal'].font.name = 'Calibri'
    doc.styles['Normal'].font.size = Pt(11)

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
        para.paragraph_format.space_after = Pt(12)

    def add_paragraph(text, bold=False, size=11, spacing=1.5):
        para = doc.add_paragraph()
        run = para.add_run(text)
        run.bold = bold
        run.font.size = Pt(size)
        run.font.color.rgb = RGBColor(0, 0, 0)
        para.paragraph_format.line_spacing = spacing
        return para

    sop_title = data.get('title', 'Generated SOP')
    sop_id = data.get('sop_id', 'SOP-000')
    prepared_by = data.get('prepared_by', 'Name')
    approved_by = data.get('approved_by', 'Approver')
    revision_date = data.get('revision_date', 'Date')

    add_paragraph(f"SOP Title: {sop_title}", bold=True, size=18)
    add_paragraph(f"SOP ID: {sop_id}", bold=False)
    add_paragraph(f"Prepared By: {prepared_by}", bold=False)
    add_paragraph(f"Approved By: {approved_by}", bold=False)
    add_paragraph(f"Revision Date: {revision_date}", bold=False)
    hr()

    sections_data = data.get("sections", [])
    for i, section in enumerate(sections_data):
        add_paragraph(section["heading"], bold=True)
        for item in section.get("content", []):
            text = item.get("text", "")
            t = item.get("type", "text")
            if t == "bullet":
                para = doc.add_paragraph(style='List Bullet')
                para.add_run(text).font.color.rgb = RGBColor(0, 0, 0)
            elif t == "sub_bullet":
                para = doc.add_paragraph(style='List Bullet 2')
                para.add_run(text).font.color.rgb = RGBColor(0, 0, 0)
            else:
                add_paragraph(text)
        if i < len(sections_data) - 1:
            hr()

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
