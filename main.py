from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from docx import Document
from docx.shared import Pt, Inches, RGBColor, Twips
from docx.oxml.shared import qn, OxmlElement
from docx.oxml.ns import nsmap
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import os
import uuid

app = Flask(__name__)
CORS(app)
DOWNLOAD_FOLDER = "/tmp"


def create_numbering_definitions(doc):
    """
    Create multi-level numbering definitions in the document.
    This adds the abstractNum and num elements needed for multi-level lists.
    """
    # Get or create numbering part
    numbering_part = doc.part.numbering_part
    if numbering_part is None:
        # Create numbering part if it doesn't exist
        from docx.parts.numbering import NumberingPart
        from docx.opc.constants import CONTENT_TYPE as CT
        numbering_part = NumberingPart.new()
        doc.part.relate_to(numbering_part, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering')
    
    numbering_xml = numbering_part._element
    
    # Check if our abstractNum already exists
    existing = numbering_xml.findall('.//w:abstractNum[@w:abstractNumId="100"]', 
                                      namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
    if existing:
        return  # Already defined
    
    # Create abstractNum for multi-level steps: 1. → a. → i. → 1. → a.
    abstract_num = OxmlElement('w:abstractNum')
    abstract_num.set(qn('w:abstractNumId'), '100')
    
    # Define 5 levels
    level_configs = [
        {'ilvl': '0', 'numFmt': 'decimal', 'lvlText': '%1.', 'left': 720, 'hanging': 360},
        {'ilvl': '1', 'numFmt': 'lowerLetter', 'lvlText': '%2.', 'left': 1440, 'hanging': 360},
        {'ilvl': '2', 'numFmt': 'lowerRoman', 'lvlText': '%3.', 'left': 2160, 'hanging': 180},
        {'ilvl': '3', 'numFmt': 'decimal', 'lvlText': '%4.', 'left': 2880, 'hanging': 360},
        {'ilvl': '4', 'numFmt': 'lowerLetter', 'lvlText': '%5.', 'left': 3600, 'hanging': 360},
    ]
    
    for cfg in level_configs:
        lvl = OxmlElement('w:lvl')
        lvl.set(qn('w:ilvl'), cfg['ilvl'])
        
        start = OxmlElement('w:start')
        start.set(qn('w:val'), '1')
        lvl.append(start)
        
        numFmt = OxmlElement('w:numFmt')
        numFmt.set(qn('w:val'), cfg['numFmt'])
        lvl.append(numFmt)
        
        lvlText = OxmlElement('w:lvlText')
        lvlText.set(qn('w:val'), cfg['lvlText'])
        lvl.append(lvlText)
        
        lvlJc = OxmlElement('w:lvlJc')
        lvlJc.set(qn('w:val'), 'left')
        lvl.append(lvlJc)
        
        pPr = OxmlElement('w:pPr')
        ind = OxmlElement('w:ind')
        ind.set(qn('w:left'), str(cfg['left']))
        ind.set(qn('w:hanging'), str(cfg['hanging']))
        pPr.append(ind)
        lvl.append(pPr)
        
        abstract_num.append(lvl)
    
    # Insert abstractNum at the beginning of numbering
    first_num = numbering_xml.find('.//w:num', 
                                    namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
    if first_num is not None:
        first_num.addprevious(abstract_num)
    else:
        numbering_xml.append(abstract_num)
    
    # Create num element that references this abstractNum
    num = OxmlElement('w:num')
    num.set(qn('w:numId'), '100')
    abstract_num_id = OxmlElement('w:abstractNumId')
    abstract_num_id.set(qn('w:val'), '100')
    num.append(abstract_num_id)
    numbering_xml.append(num)


def add_horizontal_rule(doc):
    """Add a horizontal rule as an empty paragraph with a bottom border."""
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after = Pt(0)
    
    # Add bottom border via XML
    pPr = para._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')
    bottom.set(qn('w:space'), '0')
    bottom.set(qn('w:color'), 'auto')
    pBdr.append(bottom)
    pPr.append(pBdr)
    return para


def add_empty_paragraph(doc):
    """Add a blank paragraph for visual spacing."""
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after = Pt(0)
    return para


def add_text_paragraph(doc, text, bold=False, size=11, color=RGBColor(0, 0, 0)):
    """Add a simple text paragraph."""
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after = Pt(0)
    run = para.add_run(text)
    run.bold = bold
    run.font.size = Pt(size)
    run.font.color.rgb = color
    run.font.name = 'Calibri'
    return para


def add_labelled_paragraph(doc, label, value):
    """Add a labelled paragraph with bold label and normal value."""
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after = Pt(0)
    
    # Bold label
    run1 = para.add_run(f"{label}: ")
    run1.bold = True
    run1.font.size = Pt(11)
    run1.font.color.rgb = RGBColor(0, 0, 0)
    run1.font.name = 'Calibri'
    
    # Normal value
    if value:
        run2 = para.add_run(value)
        run2.font.size = Pt(11)
        run2.font.color.rgb = RGBColor(0, 0, 0)
        run2.font.name = 'Calibri'
    
    return para


def add_label_only(doc, label):
    """Add just a label (for when bullets follow on next lines)."""
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after = Pt(0)
    run = para.add_run(f"{label}:")
    run.bold = True
    run.font.size = Pt(11)
    run.font.color.rgb = RGBColor(0, 0, 0)
    run.font.name = 'Calibri'
    return para


def add_bullet(doc, text, indent_level=0):
    """Add a bullet point with proper indentation."""
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after = Pt(0)
    
    # Set first-line indent based on level
    indent = 0.25 + (indent_level * 0.5)  # 0.25" base, +0.5" per level
    para.paragraph_format.first_line_indent = Inches(indent)
    
    # Add bullet character and text
    run = para.add_run("• " + text)
    run.font.size = Pt(11)
    run.font.color.rgb = RGBColor(0, 0, 0)
    run.font.name = 'Calibri'
    return para


def add_numbered_step(doc, text, level):
    """
    Add a numbered step using Word's multi-level numbering.
    Level 1 = 1., Level 2 = a., Level 3 = i., Level 4 = 1., Level 5 = a.
    """
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after = Pt(0)
    
    # Apply numbering via XML
    pPr = para._p.get_or_add_pPr()
    numPr = OxmlElement('w:numPr')
    
    ilvl = OxmlElement('w:ilvl')
    ilvl.set(qn('w:val'), str(level - 1))  # Convert 1-5 to 0-4
    numPr.append(ilvl)
    
    numId = OxmlElement('w:numId')
    numId.set(qn('w:val'), '100')  # Reference our numbering definition
    numPr.append(numId)
    
    pPr.insert(0, numPr)
    
    # Add text
    run = para.add_run(text)
    run.font.size = Pt(11)
    run.font.color.rgb = RGBColor(0, 0, 0)
    run.font.name = 'Calibri'
    return para


def add_note(doc, text, indent_inches=2.5):
    """Add an italic note/callout paragraph."""
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after = Pt(0)
    para.paragraph_format.left_indent = Inches(indent_inches)
    
    run = para.add_run(text)
    run.italic = True
    run.font.size = Pt(11)
    run.font.color.rgb = RGBColor(0, 0, 0)
    run.font.name = 'Calibri'
    return para


def add_revision_table(doc, rows_data):
    """
    Add the Revision History table.
    rows_data is a list of dicts with 'date', 'revised_by', 'description' keys,
    or a list of strings in "date|||revised_by|||description" format.
    """
    table = doc.add_table(rows=1, cols=3)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    
    # Set column widths (proportional: ~27%, ~20%, ~53%)
    widths = [Inches(1.82), Inches(1.31), Inches(3.51)]
    
    # Header row
    headers = ["Date", "Revised By", "Description"]
    header_cells = table.rows[0].cells
    
    for i, (cell, header, width) in enumerate(zip(header_cells, headers, widths)):
        cell.width = width
        para = cell.paragraphs[0]
        para.paragraph_format.space_after = Pt(0)
        run = para.add_run(header)
        run.bold = True
        run.font.size = Pt(11)
        run.font.color.rgb = RGBColor(0, 0, 0)
        run.font.name = 'Calibri'
        
        # Add bottom border to header cell
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        tcBorders = OxmlElement('w:tcBorders')
        bottom = OxmlElement('w:bottom')
        bottom.set(qn('w:val'), 'single')
        bottom.set(qn('w:sz'), '12')
        bottom.set(qn('w:space'), '0')
        bottom.set(qn('w:color'), 'auto')
        tcBorders.append(bottom)
        tcPr.append(tcBorders)
    
    # Data rows
    for row_data in rows_data:
        if isinstance(row_data, dict):
            values = [
                row_data.get('date', ''),
                row_data.get('revised_by', ''),
                row_data.get('description', '')
            ]
        elif isinstance(row_data, str):
            parts = row_data.split('|||')
            values = [p.strip() for p in parts] + [''] * (3 - len(parts))
        else:
            values = ['', '', '']
        
        row = table.add_row()
        for i, (cell, value, width) in enumerate(zip(row.cells, values[:3], widths)):
            cell.width = width
            para = cell.paragraphs[0]
            para.paragraph_format.space_after = Pt(0)
            run = para.add_run(value)
            run.font.size = Pt(11)
            run.font.color.rgb = RGBColor(0, 0, 0)
            run.font.name = 'Calibri'
            
            # Add top border to first data row cells
            if row_data == rows_data[0]:
                tc = cell._tc
                tcPr = tc.get_or_add_tcPr()
                tcBorders = OxmlElement('w:tcBorders')
                top = OxmlElement('w:top')
                top.set(qn('w:val'), 'single')
                top.set(qn('w:sz'), '12')
                top.set(qn('w:space'), '0')
                top.set(qn('w:color'), 'auto')
                tcBorders.append(top)
                tcPr.append(tcBorders)
    
    # Remove default table borders
    tbl = table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else OxmlElement('w:tblPr')
    tblBorders = OxmlElement('w:tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'none')
        border.set(qn('w:sz'), '0')
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), 'auto')
        tblBorders.append(border)
    tblPr.append(tblBorders)
    
    return table


def setup_footer(doc, sop_title, sop_id, revision_date):
    """Set up the footer with SOP info. First page is blank, subsequent pages have content."""
    section = doc.sections[0]
    section.different_first_page_header_footer = True
    
    # Default footer (pages 2+)
    footer = section.footer
    footer.is_linked_to_previous = False
    
    # Clear existing content
    for para in footer.paragraphs:
        p = para._element
        p.getparent().remove(p)
    
    # Add blank line
    para1 = footer.add_paragraph()
    para1.paragraph_format.space_after = Pt(0)
    
    # Add SOP title line
    para2 = footer.add_paragraph()
    para2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para2.paragraph_format.space_after = Pt(0)
    run2 = para2.add_run(f"{sop_title} [{sop_id}]")
    run2.font.size = Pt(10)
    run2.font.color.rgb = RGBColor(0, 0, 0)
    run2.font.name = 'Calibri'
    
    # Add revision date line
    para3 = footer.add_paragraph()
    para3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para3.paragraph_format.space_after = Pt(0)
    run3 = para3.add_run(f"Revision Date: {revision_date}")
    run3.font.size = Pt(10)
    run3.font.color.rgb = RGBColor(0, 0, 0)
    run3.font.name = 'Calibri'
    
    # First page footer (blank)
    first_footer = section.first_page_footer
    first_footer.is_linked_to_previous = False
    for para in first_footer.paragraphs:
        p = para._element
        p.getparent().remove(p)
    first_footer.add_paragraph()


def generate_sop_doc(data):
    """
    Generate an SOP document from the provided data.
    
    Expected data structure:
    {
        "title": "SOP Title",
        "sop_id": "SOP-001",
        "prepared_by": "Name",
        "approved_by": "Approver",
        "revision_date": "Month D, YYYY",
        "sections": [
            {
                "heading": "Section 2: Purpose and Scope",
                "content": [
                    {"type": "labelled", "text": "Objective: Description here"},
                    {"type": "bullet", "text": "Bullet item"},
                    {"type": "step", "level": 1, "text": "First step"},
                    {"type": "step", "level": 2, "text": "Sub-step"},
                    {"type": "note", "text": "Note text here"},
                    {"type": "spacer"}
                ]
            },
            {
                "heading": "Section 8: Revision History",
                "type": "table",
                "content": [
                    {"text": "Jan 1, 2025|||John Doe|||Initial draft"}
                ]
            }
        ]
    }
    """
    # Create document from template or blank
    try:
        doc = Document("template.docx")
        # Clear any placeholder content
        for para in doc.paragraphs[:]:
            p = para._element
            p.getparent().remove(p)
    except:
        doc = Document()
    
    # Set up page margins
    section = doc.sections[0]
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)
    section.footer_distance = Inches(0.5)
    section.header_distance = Inches(0.5)
    
    # Set default font
    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(11)
    
    # Create numbering definitions
    create_numbering_definitions(doc)
    
    # Extract header data
    sop_title = data.get('title', 'Generated SOP')
    sop_id = data.get('sop_id', '') or 'SOP-ID TBD'
    prepared_by = data.get('prepared_by', '')
    approved_by = data.get('approved_by', '')
    revision_date = data.get('revision_date', '')
    
    # === HEADER BLOCK ===
    # Title (18pt bold)
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after = Pt(0)
    run = para.add_run(f"SOP Title: {sop_title}")
    run.bold = True
    run.font.size = Pt(18)
    run.font.color.rgb = RGBColor(0, 0, 0)
    run.font.name = 'Calibri'
    
    # SOP ID
    add_text_paragraph(doc, f"SOP ID: {sop_id}")
    
    # Blank line
    add_empty_paragraph(doc)
    
    # Prepared By, Approved By, Revision Date
    add_text_paragraph(doc, f"Prepared By: {prepared_by}")
    add_text_paragraph(doc, f"Approved By: {approved_by}")
    add_text_paragraph(doc, f"Revision Date: {revision_date}")
    
    # Horizontal rule
    add_horizontal_rule(doc)
    
    # === SECTIONS ===
    sections_data = data.get('sections', [])
    
    for sec_idx, sec_data in enumerate(sections_data):
        heading = sec_data.get('heading', '')
        
        # Section heading
        if heading:
            add_text_paragraph(doc, heading, bold=True)
            add_empty_paragraph(doc)
        
        # Handle table type (Revision History)
        if sec_data.get('type') == 'table':
            content = sec_data.get('content', [])
            rows = []
            for item in content:
                if isinstance(item, dict):
                    rows.append(item.get('text', '') or item)
                else:
                    rows.append(item)
            add_revision_table(doc, rows)
        else:
            # Process content items
            content = sec_data.get('content', [])
            i = 0
            while i < len(content):
                item = content[i]
                if not isinstance(item, dict):
                    i += 1
                    continue
                
                item_type = item.get('type', 'text')
                text = item.get('text', '').strip()
                
                if not text and item_type != 'spacer':
                    i += 1
                    continue
                
                if item_type == 'heading':
                    add_text_paragraph(doc, text, bold=True)
                    add_empty_paragraph(doc)
                
                elif item_type == 'labelled':
                    # Check if bullets follow this label
                    bullets_follow = False
                    for j in range(i + 1, len(content)):
                        next_item = content[j]
                        if isinstance(next_item, dict):
                            next_type = next_item.get('type', '')
                            if next_type in ('bullet', 'sub_bullet'):
                                bullets_follow = True
                                break
                            elif next_type not in ('spacer',):
                                break
                    
                    if ':' in text:
                        label, _, value = text.partition(':')
                        label = label.strip()
                        value = value.strip()
                        
                        if bullets_follow or not value:
                            # Label on its own line, bullets follow
                            add_label_only(doc, label)
                        else:
                            # Inline label: value
                            add_labelled_paragraph(doc, label, value)
                    else:
                        add_text_paragraph(doc, text)
                
                elif item_type == 'bullet':
                    add_bullet(doc, text, indent_level=0)
                
                elif item_type == 'sub_bullet':
                    add_bullet(doc, text, indent_level=1)
                
                elif item_type == 'step':
                    level = item.get('level', 1)
                    level = max(1, min(5, level))  # Clamp to 1-5
                    add_numbered_step(doc, text, level)
                
                elif item_type == 'note':
                    add_empty_paragraph(doc)
                    add_note(doc, text)
                    add_empty_paragraph(doc)
                
                elif item_type == 'spacer':
                    add_empty_paragraph(doc)
                
                else:
                    # Default: plain text
                    add_text_paragraph(doc, text)
                
                i += 1
        
        # Add horizontal rule between sections (not after last)
        if sec_idx < len(sections_data) - 1:
            add_horizontal_rule(doc)
    
    # === FOOTER ===
    setup_footer(doc, sop_title, sop_id, revision_date)
    
    # Save document
    filename = f"sop_{uuid.uuid4().hex}.docx"
    filepath = os.path.join(DOWNLOAD_FOLDER, filename)
    doc.save(filepath)
    return filename


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


@app.route("/generate", methods=["POST"])
def generate():
    try:
        data = request.json
        filename = generate_sop_doc(data)
        link = f"https://sop-flask-api.onrender.com/download/{filename}"
        return jsonify({"download_link": link})
    except Exception as e:
        import traceback
        return {"error": str(e), "trace": traceback.format_exc()}, 500


@app.route("/health", methods=["GET"])
def health():
    return {"status": "ok"}


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
