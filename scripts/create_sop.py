"""
Nipple Crime SOP Generator
Creates a formatted DOCX SOP with left-only logo header.
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
):
    doc = Document()

    # --- Page margins ---
    page_section = doc.sections[0]
    page_section.top_margin = Inches(0.5)
    page_section.bottom_margin = Inches(0.75)
    page_section.left_margin = Inches(1)
    page_section.right_margin = Inches(1)

    # =========================================================
    # HEADER: left-aligned NC logo only
    # =========================================================
    logo_para = doc.add_paragraph()
    logo_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    if nc_logo_path and os.path.exists(nc_logo_path):
        logo_para.add_run().add_picture(nc_logo_path, height=Inches(1.0))
    else:
        add_run(logo_para, "NIPPLE CRIME", bold=True, size=20)

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
            elif len(line) > 2 and line[0].isdigit() and line[1] == '.':
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
        "Submission is managed by the Treasurer in coordination with the President.",
    ]),
    (1, "2. Timeline", [
        "- SOI portal opens: December – January (check burningman.org each year)",
        "- SOI submission deadline: February – March (verify annually)",
        "- Steward Sale ticket allocations announced: second week of February",
        "- Conditional Late Season Directed tickets (new camps) announced: late spring",
        "- Placement decisions announced: April – May",
        "- Treasurer sets calendar reminders at each milestone",
    ]),
    (1, "3. Submission Steps", [
        "1. Go to burningman.org > Participate > Theme Camps > Statement of Intent",
        "2. Log in using the Nipple Crime camp account (credentials held by President and Treasurer)",
        "3. Select 'Update Returning Camp' (Nipple Crime is a returning placed camp)",
        "4. Complete all required fields — see Section 4 for 2026 reference answers",
        "5. Review all entries with President before submitting",
        "6. Submit form",
        "7. Save / screenshot the BMorg confirmation email for records (see Section 5)",
        "8. If updates are needed after submission, email placement@burningman.org with subject line:",
        "   [Last Sector] Camp Name - SOI Updates   (example: [3:00] Nipple Crime - SOI Updates)",
    ]),
    (1, "4. 2026 Submission Reference", [
        "The following answers were submitted for the 2026 Burn cycle. Update each year as needed.",
        "",
        "CAMP CONTACT",
        "- First Name: Reece",
        "- Last Name: Dassinger",
        "- Email: reece@nipplecrime.org",
        "- Playa Name: Kirkland Cowboy",
        "- Phone: 775 247 4493",
        "- Address: 5592 Spandrell Cir, Sparks, NV 89436",
        "",
        "BURNING MAN HISTORY",
        "- Years attended: 2014–2024",
        "- Projects & Affiliations: Theme Camp, Mutant Vehicle, Art Installation",
        "- Description: 2014 participated in theme camp; 2015 built bar, designed/coordinated theme camp; "
        "2016–2023 Nipple Crime camp lead; built art car 2023",
        "",
        "CAMP DETAILS",
        "- Camp category: Theme Camp",
        "- Returning camp: Yes",
        "- Camp name: Nipple Crime",
        "- Last placed year: 2025",
        "- Will complete PCQ: Yes (existing placed camp)",
        "",
        "INTERACTIVITY",
        "- Interactivity change: Reduced Interactivity",
        "- Description: Reducing one or two larger events; more open bar and fire poof time to reduce workload",
        "- Center Camp Activation: No",
        "",
        "MUTANT VEHICLES",
        "- Hosting MVs: Yes",
        "- MV Details: Great Sax, SirKiss, Green Zebra",
        "",
        "POPULATION",
        "- Number of campmates: 105",
        "- Population change: About the same",
        "",
        "TICKETS",
        "- Number of Steward Sale tickets requested: 44",
        "",
        "ADDITIONAL CONTACTS",
        "- Additional Contact: Isabel Hoy — izhoy@yahoo.com",
        "- Sustainability Contact: Chris Reddin — creddin1@hotmail.com",
    ]),
    (1, "5. BMorg Confirmation (2026)", [
        "Upon successful submission BMorg sends the following confirmation:",
        "",
        "\"Thank you for submitting The Theme Camp Statement of Intent. We'll be reviewing these "
        "in the upcoming weeks and will reach out if we have any questions or need clarification. "
        "Stewards Sale allocations will be announced by the second week of February. Conditional "
        "Late Season Directed tickets for new theme camps will be announced in the late spring.\"",
        "",
        "\"If you are a theme camp in good standing and are taking 2026 off, we will note that "
        "for future access. Please plan to complete a SOI when you return!\"",
        "",
        "\"If you have updates after submitting, please email placement@burningman.org with the "
        "last sector your camp was placed in, camp name, and 'SOI Updates' in the subject line. "
        "Example: [3:00] Camp Buttercup - SOI Updates.\"",
        "",
        "Save this confirmation email to: [Shared Drive > Treasurer > BMorg > SOI > YYYY]",
    ]),
    (1, "6. Record Keeping", [
        "- Save submitted SOI confirmation email as PDF",
        "- File location: [Shared Drive > Treasurer > BMorg > SOI > YYYY]",
        "- Log: submission date, submitted by, any BMorg correspondence, placement result",
        "- Retain records for minimum 5 years per nonprofit compliance requirements",
    ]),
    (1, "7. Contacts", [
        "- BMorg Placement: placement@burningman.org",
        "- President / Primary Submitter: Reece Dassinger — reece@nipplecrime.org",
        "- Treasurer: Isabel Hoy — izhoy@yahoo.com",
        "- VP / SOP Owner: Chris Reddin — creddin1@hotmail.com",
    ]),
    (1, "8. Revision History", [
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
)
