"""
Nipple Crime SOP Generator
Creates a formatted DOCX SOP with dual-logo header.
Usage: python scripts/create_sop.py
"""

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os


def set_cell_border(cell, **kwargs):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for edge in ('top', 'left', 'bottom', 'right'):
        element = OxmlElement(f'w:{edge}')
        for key, val in kwargs.get(edge, {}).items():
            element.set(qn(f'w:{key}'), str(val))
        tcBorders.append(element)
    tcPr.append(tcBorders)


def set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)


def add_run(para, text, bold=False, size=None, color=None, italic=False):
    run = para.add_run(text)
    run.bold = bold
    run.italic = italic
    if size:
        run.font.size = Pt(size)
    if color:
        run.font.color.rgb = RGBColor(*color)
    return run


def create_sop(
    output_path,
    sop_number,
    sop_title,
    department,
    version,
    effective_date,
    last_updated,
    sections,           # list of (heading_level, heading_text, body_lines)
    nc_logo_path="Images/NC logo.png",
    bm_logo_path="Images/BM logo.jpg",
):
    doc = Document()

    # --- Page margins ---
    page_section = doc.sections[0]
    page_section.top_margin = Inches(0.5)
    page_section.bottom_margin = Inches(0.75)
    page_section.left_margin = Inches(1)
    page_section.right_margin = Inches(1)

    # =========================================================
    # HEADER: 2-column logo row (NC left | BM right)
    # =========================================================
    logo_table = doc.add_table(rows=1, cols=2)
    logo_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    logo_table.columns[0].width = Inches(3.25)
    logo_table.columns[1].width = Inches(3.25)

    # Left cell — NC logo
    left_cell = logo_table.cell(0, 0)
    left_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    lp = left_cell.paragraphs[0]
    lp.alignment = WD_ALIGN_PARAGRAPH.LEFT
    if nc_logo_path and os.path.exists(nc_logo_path):
        lp.add_run().add_picture(nc_logo_path, height=Inches(1.0))
    else:
        add_run(lp, "NIPPLE CRIME", bold=True, size=14)

    # Right cell — BM logo
    right_cell = logo_table.cell(0, 1)
    right_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    rp = right_cell.paragraphs[0]
    rp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    if bm_logo_path and os.path.exists(bm_logo_path):
        rp.add_run().add_picture(bm_logo_path, height=Inches(1.0))
    else:
        add_run(rp, "[Burning Man Logo]", bold=True, size=14)

    # Remove table borders
    for row in logo_table.rows:
        for cell in row.cells:
            set_cell_border(cell,
                top={"val": "none"}, left={"val": "none"},
                bottom={"val": "none"}, right={"val": "none"})

    doc.add_paragraph()  # spacer

    # =========================================================
    # DIVIDER LINE
    # =========================================================
    div = doc.add_paragraph()
    pPr = div._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '12')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), '000000')
    pBdr.append(bottom)
    pPr.append(pBdr)

    # =========================================================
    # TITLE BLOCK
    # =========================================================
    label_para = doc.add_paragraph()
    label_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_run(label_para, "STANDARD OPERATING PROCEDURE", bold=True, size=11)

    title_para = doc.add_paragraph()
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_run(title_para, sop_title, bold=True, size=15)

    doc.add_paragraph()

    # =========================================================
    # METADATA TABLE
    # Fields: SOP Number | Department | Version | Effective Date
    # Last Updated spans full width
    # =========================================================
    meta_table = doc.add_table(rows=2, cols=4)
    meta_table.style = 'Table Grid'

    meta_rows = [
        [("SOP Number", sop_number), ("Department", department)],
        [("Version", version),       ("Effective Date", effective_date)],
    ]

    for r, row_data in enumerate(meta_rows):
        row = meta_table.rows[r]
        for c, (lbl, val) in enumerate(row_data):
            label_cell = row.cells[c * 2]
            value_cell = row.cells[c * 2 + 1]
            set_cell_bg(label_cell, "D9D9D9")
            lp = label_cell.paragraphs[0]
            lp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            add_run(lp, lbl, bold=True, size=9)
            vp = value_cell.paragraphs[0]
            add_run(vp, val, size=9)

    # Last updated — full-width merged row
    lu_row = meta_table.add_row()
    merged = lu_row.cells[0].merge(lu_row.cells[1]).merge(lu_row.cells[2]).merge(lu_row.cells[3])
    set_cell_bg(merged, "D9D9D9")
    lup = merged.paragraphs[0]
    lup.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_run(lup, f"Last Updated: {last_updated}", bold=True, size=9)

    doc.add_paragraph()

    # =========================================================
    # BODY SECTIONS
    # =========================================================
    for (level, heading, lines) in sections:
        h = doc.add_heading(heading, level=level)
        h.runs[0].font.size = Pt({1: 13, 2: 11, 3: 10}.get(level, 10))
        h.runs[0].font.color.rgb = RGBColor(0, 0, 0)

        for line in lines:
            if line.startswith("- "):
                p = doc.add_paragraph(line[2:], style='List Bullet')
            elif len(line) > 2 and line[0].isdigit() and line[1] in '.':
                p = doc.add_paragraph(line[line.index(".")+1:].strip(), style='List Number')
            elif len(line) > 3 and line[:2].isdigit() and line[2] == '.':
                p = doc.add_paragraph(line[line.index(".")+1:].strip(), style='List Number')
            elif line == "":
                p = doc.add_paragraph()
                continue
            else:
                p = doc.add_paragraph(line)
            if p.runs:
                p.runs[0].font.size = Pt(10)

    # =========================================================
    # FOOTER
    # =========================================================
    footer = page_section.footer
    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_run(fp,
        f"Nipple Crime Theme Camp  |  {sop_number} {sop_title}  |  Ver {version}  |  Confidential — Internal Use Only",
        size=8, color=(128, 128, 128))

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    doc.save(output_path)
    print(f"Saved: {output_path}")


# =========================================================
# SOP Tr3 — Statement of Intent (BMorg)
# =========================================================

sections_tr3 = [
    (1, "1. Purpose", [
        "This SOP documents the annual process for completing and submitting Nipple Crime's "
        "Statement of Intent (SOI) to the Burning Man Organization (BMorg). The SOI is the "
        "formal application required for theme camp registration and playa placement. "
        "Submission is managed by the Treasurer in coordination with the President and VP.",
    ]),
    (1, "2. Scope", [
        "Applies to: Treasurer, President, VP, Infrastructure Director, Committees Director, HR Officer.",
        "Frequency: Annually (each Burn cycle).",
    ]),
    (1, "3. Timeline", [
        "- SOI portal opens: December – January (check burningman.org each year)",
        "- Internal info-gathering deadline: 2 weeks before portal opens",
        "- SOI submission deadline: February – March (verify annually)",
        "- Placement decision announced: April – May",
        "- Treasurer sets calendar reminders at each milestone",
    ]),
    (1, "4. Roles & Responsibilities", [
        "- Treasurer: Owns submission process; coordinates data gathering; maintains records",
        "- President: Reviews and approves SOI before submission",
        "- VP: Provides camp overview and strategic goals",
        "- Infrastructure Director: Provides power, vehicle, and physical infrastructure data",
        "- Committees Director: Provides list of interactive activities and public offerings",
        "- HR Officer: Provides confirmed participant count",
    ]),
    (1, "5. Information to Gather", [
        "The Treasurer sends a request to all parties (see Section 4) no later than 2 weeks before portal opens.",
        "",
        "Camp Identity (from President / VP):",
        "- Official camp name",
        "- Camp theme and description (public-facing, ~2–3 sentences)",
        "- Year established / returning camp status",
        "",
        "Participation (from HR Officer):",
        "- Total confirmed camp members",
        "- Estimated vehicle count",
        "- Number of structures / shade structures",
        "",
        "Interactive Activities (from Committees Director):",
        "- List of all public-facing activities and events",
        "- Any art installations or performances",
        "",
        "Infrastructure (from Infrastructure Director):",
        "- Total power requirement (kilowatts)",
        "- Generator count and capacity",
        "- Freshwater and greywater needs",
        "- Number and type of vehicles on playa",
        "- Shade structure dimensions and type",
        "",
        "Financial / Nonprofit (Treasurer owned):",
        "- 501(c)(3) / 509(a)(2) status confirmation",
        "- Dues and fee structure (for BMorg reference if requested)",
    ]),
    (1, "6. Submission Steps", [
        "1. Navigate to the BMorg Theme Camp SOI portal (burningman.org > Theme Camps > Registration)",
        "2. Log in with the Nipple Crime camp account (credentials held by Treasurer and President)",
        "3. Select 'Start New Statement of Intent' or 'Update Returning Camp'",
        "4. Complete all required fields using the information gathered in Section 5",
        "5. In the 'Camp Description' field, use the approved language from the President",
        "6. Review all entries for accuracy",
        "7. Share draft with President and VP for approval (minimum 48 hours before deadline)",
        "8. Incorporate any revisions",
        "9. Submit and immediately save / screenshot the confirmation page",
        "10. Record the BMorg confirmation number in the Records Log (see Section 8)",
    ]),
    (1, "7. Post-Submission", [
        "- Monitor the Nipple Crime email inbox for BMorg follow-up requests",
        "- Respond to any BMorg questions within 48 hours",
        "- Notify the full Board when placement decision is received",
        "- If waitlisted: follow BMorg waitlist instructions and notify Board immediately",
        "- If denied: convene Board discussion within 1 week to determine next steps",
    ]),
    (1, "8. Record Keeping", [
        "- Save submitted SOI as PDF (screenshot confirmation page if PDF not available)",
        "- File location: [Shared Drive > Treasurer > BMorg > SOI > YYYY]",
        "- Log entry in the BMorg Records Log:",
        "    Submission date, confirmation number, submitted by, portal URL, placement result",
        "- Retain records for minimum 5 years per nonprofit compliance requirements",
    ]),
    (1, "9. Contacts", [
        "- BMorg Theme Camp Services: burningman.org (contact form under Theme Camps section)",
        "- President: Reece Dassinger",
        "- VP / SOP Owner: Chris Reddin",
        "- Treasurer / Process Owner: Isabel Hoy",
    ]),
    (1, "10. Revision History", [
        "- v1.0 | 2026-03-03 | Initial draft | Chris Reddin",
    ]),
]

create_sop(
    output_path="Standard Operating Procedures/Tr3 Statement of Intent.docx",
    sop_number="Tr3",
    sop_title="Statement of Intent (BMorg)",
    department="Treasurer",
    version="1.0",
    effective_date="2026-03-03",
    last_updated="2026-03-03",
    sections=sections_tr3,
    nc_logo_path="Images/NC logo.png",
    bm_logo_path="Images/BM logo.jpg",
)
