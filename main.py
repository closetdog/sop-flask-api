from flask import Flask, request, send_file
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.oxml.shared import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_LINE_SPACING
import tempfile

app = Flask(__name__)

def generate_sop_doc(data):
    doc = Document()
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
        para.paragraph_format.space_before = Pt(0)
        para.paragraph_format.space_after = Pt(0)
        para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE

    def paragraph(text, bold=False, size=None):
        para = doc.add_paragraph(style="No Spacing")
        run = para.add_run(text)
        if bold:
            run.bold = True
        if size:
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

    def bullet(text, indent=0.5):
        para = doc.add_paragraph(style='List Bullet')
        para.add_run(text)
        para.paragraph_format.left_indent = Inches(indent)
        para.paragraph_format.space_before = Pt(0)
        para.paragraph_format.space_after = Pt(0)
        para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE

    def dash(text, indent=0.45):
        para = doc.add_paragraph()
        para.add_run(f"\u2013 {text}")
        para.paragraph_format.left_indent = Inches(indent)
        para.paragraph_format.space_before = Pt(0)
        para.paragraph_format.space_after = Pt(0)
        para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE

    def sub_bullet(text, indent=0.90):
        para = doc.add_paragraph(style='List Bullet')
        para.add_run(text)
        para.paragraph_format.left_indent = Inches(indent)
        para.paragraph_format.space_before = Pt(0)
        para.paragraph_format.space_after = Pt(0)
        para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE

    heading_paragraph(f"SOP Title: {data.get('title', 'Generated SOP')}", size=18)
    paragraph(f"Prepared By: {data.get('prepared_by', 'Name')}")
    paragraph(f"Approved By: {data.get('approved_by', 'Approver')}")
    paragraph(f"Revision Date: {data.get('revision_date', 'Date')}")
    hr()

    for section in data.get("sections", []):
        heading_paragraph(section["heading"])
        for item in section.get("content", []):
            t = item["type"]
            if t == "text":
                paragraph(item["text"], bold=item.get("bold", False))
            elif t == "bullet":
                bullet(item["text"], indent=item.get("indent", 0.5))
            elif t == "dash":
                dash(item["text"], indent=item.get("indent", 0.45))
            elif t == "sub_bullet":
                sub_bullet(item["text"], indent=item.get("indent", 0.90))
        hr()

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    doc.save(tmp.name)
    return tmp.name

@app.route("/generate", methods=["POST"])
def generate():
    print("\U0001f535 Received request to /generate")
    try:
        data = request.json
        print("\U0001f7e2 JSON received:", data)
        path = generate_sop_doc(data)
        print("\u2705 Document generated:", path)
        return send_file(path, as_attachment=True)
    except Exception as e:
        print("\u274c Error:", e)
        return {"error": str(e)}, 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
