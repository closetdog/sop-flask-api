import re
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
    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(11)

    def format_paragraph_with_prefix(text, indent):
        match = re.match(r"^([A-Za-z0-9ivxIVX]+[.)])\s+(.*)", text)
        para = doc.add_paragraph()
        if match:
            prefix = match.group(1)
            remainder = match.group(2)
            run1 = para.add_run(prefix)
            run1.bold = True
            run1.font.size = Pt(11)
            run1.font.color.rgb = RGBColor(0, 0, 0)
            run2 = para.add_run(f" {remainder}")
            run2.font.size = Pt(11)
            run2.font.color.rgb = RGBColor(0, 0, 0)
        else:
            run = para.add_run(text)
            run.font.size = Pt(11)
            run.font.color.rgb = RGBColor(0, 0, 0)
        para.paragraph_format.left_indent = Inches(indent)
        para.paragraph_format.space_before = Pt(0)
        para.paragraph_format.space_after = Pt(0)
        para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE

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
        para.paragraph_format.space_before = Pt(0)
        para.paragraph_format.space_after = Pt(0)
        para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE

    def paragraph(text, bold=False, size=11):
        para = doc.add_paragraph()
        run = para.add_run(text)
        if bold:
            run.bold = True
        run.font.size = Pt(size)
        run.font.color.rgb = RGBColor(0, 0, 0)
        para.paragraph_format.space_before = Pt(0)
        para.paragraph_format.space_after = Pt(0)
        para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE

    def heading_paragraph(text, size=14):
        para = doc.add_paragraph()
        run = para.add_run(text)
        run.bold = True
        run.font.size = Pt(size)
        run.font.color.rgb = RGBColor(0, 0, 0)
        para.paragraph_format.space_before = Pt(0)
        para.paragraph_format.space_after = Pt(6)
        para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE

    sections = data.get("sections", [])
    for i, section in enumerate(sections):
        heading_paragraph(section["heading"])
        for item in section.get("content", []):
            t = item["type"]
            text = item["text"]
            indent = item.get("indent", 0.5 if t == "bullet" else 0.45 if t == "dash" else 0.9)

            if t in ["bullet", "dash", "sub_bullet"]:
                format_paragraph_with_prefix(text, indent)
            elif t == "text":
                paragraph(text, bold=item.get("bold", False))

        if i < len(sections) - 1:
            hr()

    filename = f"sop_{uuid.uuid4().hex}.docx"
    filepath = os.path.join(DOWNLOAD_FOLDER, filename)
    doc.save(filepath)
    return filename

@app.route("/generate", methods=["POST"])
def generate():
    print("🔵 Received request to /generate")
    try:
        data = request.json
        print("🟢 JSON received:", data)
        filename = generate_sop_doc(data)
        print("✅ Document saved as:", filename)
        link = f"https://sop-flask-api.onrender.com/download/{filename}"
        return jsonify({"download_link": link})
    except Exception as e:
        print("❌ Error:", e)
        return {"error": str(e)}, 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
