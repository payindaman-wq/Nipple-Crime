"""
Nipple Crime SOP Generator
Creates a formatted DOCX SOP matching the master template (Tr3).
Usage: python scripts/create_sop.py
"""

from docx import Document
from docx.shared import Inches, Pt, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import lxml.etree as etree
import os

NS_WP  = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
NS_A   = 'http://schemas.openxmlformats.org/drawingml/2006/main'
NS_PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
NS_R   = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'


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


def add_floating_image(paragraph, image_path, width_emu, height_emu, pos_h_emu, pos_v_emu):
    """
    Add a floating (anchored) image to a paragraph by:
    1. Adding it inline to register the relationship and get the rId
    2. Replacing the inline XML with an anchor at the specified position
    """
    run = paragraph.add_run()
    run.add_picture(image_path, width=Emu(width_emu), height=Emu(height_emu))

    w_drawing = run._element.find(qn('w:drawing'))
    inline = w_drawing.find('{%s}inline' % NS_WP)

    # Extract rId from blip element
    blip = inline.find('.//{%s}blip' % NS_A)
    r_id = blip.get('{%s}embed' % NS_R)

    anchor_xml = (
        '<wp:anchor xmlns:wp="{wp}" xmlns:a="{a}" xmlns:pic="{pic}" xmlns:r="{r}" '
        'distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="3" '
        'behindDoc="0" locked="0" layoutInCell="0" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/>'
        '<wp:positionH relativeFrom="column"><wp:posOffset>{ph}</wp:posOffset></wp:positionH>'
        '<wp:positionV relativeFrom="paragraph"><wp:posOffset>{pv}</wp:posOffset></wp:positionV>'
        '<wp:extent cx="{cx}" cy="{cy}"/>'
        '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        '<wp:wrapSquare wrapText="largest"/>'
        '<wp:docPr id="2" name="BMlogo"/>'
        '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<pic:pic><pic:nvPicPr>'
        '<pic:cNvPr id="2" name="BMlogo"/>'
        '<pic:cNvPicPr><a:picLocks noChangeAspect="1" noChangeArrowheads="1"/></pic:cNvPicPr>'
        '</pic:nvPicPr>'
        '<pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        '<pic:spPr bwMode="auto">'
        '<a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>'
        '</pic:spPr>'
        '</pic:pic></a:graphicData></a:graphic>'
        '</wp:anchor>'
    ).format(
        wp=NS_WP, a=NS_A, pic=NS_PIC, r=NS_R,
        ph=pos_h_emu, pv=pos_v_emu,
        cx=width_emu, cy=height_emu,
        rid=r_id,
    )

    anchor = etree.fromstring(anchor_xml)
    w_drawing.remove(inline)
    w_drawing.append(anchor)


def create_sop(
    output_path,
    sop_number,
    sop_title,
    department,
    version,
    effective_date,
    last_updated,
    sections,
    nc_logo_path="Images/NC logo.png",
    bm_logo_path="Images/BM logo.jpg",
):
    doc = Document()

    page_section = doc.sections[0]
    page_section.top_margin    = Inches(0.5)
    page_section.bottom_margin = Inches(0.75)
    page_section.left_margin   = Inches(1)
    page_section.right_margin  = Inches(1)

    # =========================================================
    # HEADER: NC logo inline left + BM logo floating right
    # Dimensions and positions match the Tr3 master template exactly.
    #   NC logo  — inline:  ~3.16" wide x 0.90" tall
    #   BM logo  — anchored: ~1.67" wide x 1.19" tall
    #              H offset from column: ~5.28" (4,833,620 EMU)
    #              V offset from paragraph: -0.14" (-123,825 EMU)
    # =========================================================
    logo_para = doc.add_paragraph()
    logo_para.alignment = WD_ALIGN_PARAGRAPH.LEFT

    if nc_logo_path and os.path.exists(nc_logo_path):
        logo_para.add_run().add_picture(nc_logo_path, height=Inches(0.9))
    else:
        add_run(logo_para, "NIPPLE CRIME", bold=True, size=20)

    if bm_logo_path and os.path.exists(bm_logo_path):
        add_floating_image(
            logo_para,
            bm_logo_path,
            width_emu=1529080,
            height_emu=1091565,
            pos_h_emu=4833620,
            pos_v_emu=-123825,
        )

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
    add_run(lup, "Last Updated: %s" % last_updated, bold=True, size=9)

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
        "Nipple Crime Theme Camp  |  %s %s  |  Ver %s  |  Confidential — Internal Use Only"
        % (sop_number, sop_title, version),
        size=8, color=(128, 128, 128))

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    doc.save(output_path)
    print("Saved: %s" % output_path)


# =========================================================
# SOP Tr3 — Statement of Intent (BMorg)
# NOTE: Tr3 master template was manually finalized. Do not regenerate.
# =========================================================

# =========================================================
# SOP Tr4 — Mutant Vehicle Statement of Intent (BMorg)
# =========================================================

sections_tr4 = [
    (1, "1. Purpose", [
        "This SOP documents the annual process for completing and submitting Nipple Crime's "
        "Mutant Vehicle Statement of Intent (MVSOI) to the Burning Man Organization (BMorg). "
        "The MVSOI is required for all mutant vehicles seeking to operate on the playa and to "
        "be considered for the Stewards Ticket Sale allocation. Primary vehicle: SirKiss.",
    ]),
    (1, "2. Timeline", [
        "- MVSOI portal opens: December – January (check burningman.org each year)",
        "- Submission deadline: February – March (verify annually)",
        "- Stewards Sale ticket allocations announced: second week of February",
        "- On-playa DMV licensing: upon arrival at Burning Man each year",
        "- Set calendar reminders at each milestone",
    ]),
    (1, "3. Submission Steps", [
        "1. Go to burningman.org > Participate > Mutant Vehicles > Statement of Intent",
        "2. Log in using the Nipple Crime account (credentials held by President and Treasurer)",
        "3. Select 'Previously Applied — We've applied to bring this vehicle before'",
        "4. Complete all required fields — see Section 4 for 2026 reference answers",
        "5. Review entries before submitting",
        "6. Submit form",
        "7. Save / screenshot the BMorg confirmation email for records (see Section 5)",
    ]),
    (1, "4. 2026 Submission Reference", [
        "The following answers were submitted for the 2026 Burn cycle. Update each year as needed.",
        "",
        "CONTACT",
        "- First Name: Reece",
        "- Last Name: Dassinger",
        "- Email: Leadership@nipplecrime.org",
        "",
        "VEHICLE",
        "- Primary Mutant Vehicle Name: SirKiss",
        "- Most Recent Placed Camp Name: Nipple Crime",
        "- Vehicle Status: Previously Applied — We've applied to bring this vehicle before",
        "- Most Recent Year Licensed at On-Playa DMV: 2025",
        "",
        "INTERACTIVITY",
        "- Participatory Aspects: Offering Rides, Music/Sound System, DJ Platform",
        "- Description: We have a sound tech, DJ lineups, crowd management, walkers, and we allow "
        "anyone on the art car at all times. We are handicap accessible. We have been in the zip line.",
        "",
        "STATUS FOR 2026",
        "- Active / Requesting Access: We plan to bring our Mutant Vehicle in 2026 and would like "
        "to be considered for the Stewards Ticket Sale this year.",
        "",
        "MULTIPLE VEHICLES",
        "- Registering more than one MV: No",
        "",
        "TICKET INFORMATION",
        "- Total crew and passenger capacity of all MVs: 60",
        "- Tickets requested for MV Crew and Support Team: 18",
        "  (Note: tickets sold in pairs; MV crew only — theme camp support tickets requested separately through Placement)",
        "",
        "CAMPING PLANS",
        "- Part of Another Placed Camp: Team members camping with Nipple Crime (placed theme camp "
        "submitting its own placement request)",
    ]),
    (1, "5. Record Keeping", [
        "- Save BMorg confirmation email as PDF",
        "- File location: [Shared Drive > Treasurer > BMorg > MVSOI > YYYY]",
        "- Log: submission date, submitted by, BMorg correspondence, ticket allocation result",
        "- Retain records for minimum 5 years per nonprofit compliance requirements",
    ]),
    (1, "6. Contacts", [
        "- BMorg Mutant Vehicle Team: burningman.org (contact via MV portal)",
        "- President / Primary Submitter: Reece Dassinger — reece@nipplecrime.org",
        "- Art Car Supervisor (SirKiss): Anthony Tolosano",
        "- Treasurer: Isabel Hoy — izhoy@yahoo.com",
        "- VP / SOP Owner: Chris Reddin — creddin1@hotmail.com",
    ]),
    (1, "7. Revision History", [
        "- v1.0 | 2026-03-03 | Initial draft | Chris Reddin",
    ]),
]

create_sop(
    output_path="Standard Operating Procedures/Tr4 Mutant Vehicle Statement of Intent.docx",
    sop_number="Tr4",
    sop_title="Mutant Vehicle Statement of Intent (BMorg)",
    department="Treasurer",
    version="1.0",
    effective_date="2026-03-03",
    last_updated="2026-03-03",
    sections=sections_tr4,
)
