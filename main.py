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

    def bullet(text, indent=0.5):
        para = doc.add_paragraph(style='List Bullet')
        para.add_run(text)
        para.paragraph_format.left_indent = Inches(indent)
        para.paragraph_format.space_before = Pt(0)
        para.paragraph_format.space_after = Pt(0)
        para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE

    def dash(text, indent=0.45):
        para = doc.add_paragraph()
        para.add_run(f"– {text}")
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

    sections = data.get("sections", [])
    for i, section in enumerate(sections):
        heading_paragraph(section["heading"])
        skip_indices = set()
        for idx, item in enumerate(section.get("content", [])):
            if idx in skip_indices:
                continue
            if not item.get("text"):
                continue
            t = item["type"]
            text = item["text"]
            if t == "text" and text.lower().startswith("process owner:"):
                label, _, rest = text.partition(":")
                para = doc.add_paragraph()
                run1 = para.add_run(f"{label}:")
                run1.bold = True
                run1.font.size = Pt(11)
                run1.font.color.rgb = RGBColor(0, 0, 0)
                if "," in rest:
                    for item in rest.split(","):
                        para = doc.add_paragraph(style='List Bullet')
                        para.paragraph_format.left_indent = Inches(0.5)
                        run = para.add_run(item.strip())
                        run.font.size = Pt(11)
                        run.font.color.rgb = RGBColor(0, 0, 0)
                        para.paragraph_format.space_before = Pt(0)
                        para.paragraph_format.space_after = Pt(0)
                        para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
                    para.paragraph_format.space_after = Pt(6)
                else:
                    if "," in rest:
                para.paragraph_format.space_after = Pt(6)
                for part in rest.split(","):
                    bpara = doc.add_paragraph(style='List Bullet')
                    bpara.paragraph_format.left_indent = Inches(0.5)
                    brun = bpara.add_run(part.strip())
                    brun.font.size = Pt(11)
                    brun.font.color.rgb = RGBColor(0, 0, 0)
                    bpara.paragraph_format.space_before = Pt(0)
                    bpara.paragraph_format.space_after = Pt(0)
                    bpara.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
                else:
                    run2 = para.add_run(f" {rest.strip()}")
                    run2.font.size = Pt(11)
                    run2.font.color.rgb = RGBColor(0, 0, 0)
                para.paragraph_format.space_before = Pt(0)
                para.paragraph_format.space_after = Pt(6)
                para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
            elif t == "text" and text.lower().startswith("roles:"):
                label, _, rest = text.partition(":")
                para = doc.add_paragraph()
                run1 = para.add_run(f"{label}:")
                run1.bold = True
                run1.font.size = Pt(11)
                run1.font.color.rgb = RGBColor(0, 0, 0)
                para.paragraph_format.space_before = Pt(0)
                para.paragraph_format.space_after = Pt(0)
                para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
                # capture following items that are plain roles
                role_index = idx + 1
                while role_index < len(section["content"]):
                    role_item = section["content"][role_index]
                    role_text = role_item.get("text", "").strip()
                    if not role_text or any(role_text.lower().startswith(x) for x in ["process owner:", "objective:", "scope:", "roles:"]):
                        break
                    para = doc.add_paragraph(style='List Bullet')
                    para.paragraph_format.left_indent = Inches(0.5)
                    run = para.add_run(role_text)
                    run.font.size = Pt(11)
                    run.font.color.rgb = RGBColor(0, 0, 0)
                    para.paragraph_format.space_before = Pt(0)
                    para.paragraph_format.space_after = Pt(0)
                    para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
                    skip_indices.add(role_index)
                    role_index += 1
                idx = role_index - 1

            elif t == "text" and any(text.lower().startswith(x) for x in ["objective:", "scope:", "inputs:", "outputs:"]):
                label, _, rest = text.partition(":")
                para = doc.add_paragraph()
                run1 = para.add_run(f"{label}:")
                run1.bold = True
                run1.font.size = Pt(11)
                run1.font.color.rgb = RGBColor(0, 0, 0)
                run2 = para.add_run(f" {rest.strip()}")
                run2.font.size = Pt(11)
                run2.font.color.rgb = RGBColor(0, 0, 0)
                para.paragraph_format.space_before = Pt(0)
                para.paragraph_format.space_after = Pt(6) if label.lower() in ["objective", "inputs"] else Pt(0)
                para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
            elif t == "text":
                paragraph(text, bold=item.get("bold", False))
            elif t == "bullet":
                bullet(item["text"], indent=item.get("indent", 0.5))
            elif t == "dash":
                dash(item["text"], indent=item.get("indent", 0.45))
            elif t == "sub_bullet":
                sub_bullet(item["text"], indent=item.get("indent", 0.90))
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
