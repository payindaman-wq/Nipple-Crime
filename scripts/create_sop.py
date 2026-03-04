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

# =========================================================
# SOP Com1 — Slack
# =========================================================

sections_com1 = [
    (1, "1. Purpose", [
        "This SOP covers how Nipple Crime uses Slack as its central internal communications "
        "platform. It defines the channel structure, posting rules, and the meeting recap "
        "process managed by the Communications Officer.",
    ]),
    (1, "2. Channel Structure", [
        "- #announcements — Camp-wide announcements; replies in thread only",
        "- #leadership-board — Board-level discussion (board members only)",
        "- #leadership-general — All leadership roles",
        "- #committee-food — Kitchen & Bar committee",
        "- #committee-power — Power Grid team",
        "- #committee-lnt — Leave No Trace team",
        "- #help-wanted — Open asks and volunteer recruitment",
        "Additional committee channels created as needed using the #committee- prefix.",
    ]),
    (1, "3. Posting Rules", [
        "- Announcements: one top-level post, all replies in thread",
        "- Decisions: summarize in one 'Decision:' message and pin it to the channel",
        "- Pin the following in relevant channels at all times:",
        "  - Interest / labor form link",
        "  - Dues payment link (when available)",
        "  - Org chart",
        "  - Meeting recap index",
    ]),
    (1, "4. Meeting Recap SOP", [
        "Owner: Communications Officer",
        "Deadline: Post within 24 hours of any board or leadership meeting.",
        "Post to: #leadership-board in Slack (and optionally email the board).",
        "",
        "TEMPLATE (copy/paste each recap):",
        "",
        "Meeting: Board Meeting -- YYYY-MM-DD",
        "Attendees: [list names]",
        "",
        "Key Decisions:",
        "- Decision 1 (who decided, any constraints)",
        "- Decision 2",
        "",
        "Action Items (Owner | Action | Due Date | Dependencies):",
        "- [Owner] -- [action] -- [due date] -- [dependencies]",
        "",
        "Risks / Watch-outs:",
        "- [Risk] -- [mitigation owner] -- [next check-in date]",
        "",
        "Next Meeting:",
        "- Date/time + what must be ready by then",
    ]),
    (1, "5. Contacts", [
        "- Communications Officer (Slack owner): Daemon Wyner",
        "- President (approves major messaging): Reece Dassinger",
        "- VP / SOP Owner: Chris Reddin",
    ]),
    (1, "6. Revision History", [
        "- v1.0 | 2026-03-03 | Initial draft | Chris Reddin",
    ]),
]

create_sop(
    output_path="Standard Operating Procedures/Com1 Slack.docx",
    sop_number="Com1",
    sop_title="Slack",
    department="Communications",
    version="1.0",
    effective_date="2026-03-03",
    last_updated="2026-03-03",
    sections=sections_com1,
)

# =========================================================
# SOP Com2 — Sakari
# =========================================================

sections_com2 = [
    (1, "1. Purpose", [
        "This SOP covers how Nipple Crime uses Sakari for SMS outreach to camp members. "
        "Sakari is integrated with HubSpot, allowing bulk and 1:1 texts to be sent "
        "directly to HubSpot contact lists.",
    ]),
    (1, "2. Access & Integration Setup", [
        "- Confirm Sakari is connected to HubSpot (Settings > Integrations in Sakari)",
        "- When connected, HubSpot contact lists sync and a Sakari SMS card appears on each contact record",
        "- Credentials held by: Communications Officer and President",
    ]),
    (1, "3. Compliance (Opt-Out)", [
        "- Opt-out is carrier keyword-driven: members text STOP to opt out",
        "- Always include 'Reply STOP to opt out' on first outreach of each year",
        "- Do not override carrier opt-out behavior",
        "- Pricing is driven by SMS segments (longer messages or special characters cost more -- keep messages short)",
    ]),
    (1, "4. Send a Bulk SMS Campaign", [
        "Use this to text the whole membership (e.g., annual interest check-in).",
        "",
        "1. In Sakari, go to Campaigns > Create Campaign",
        "2. Complete the campaign stages: Details, Contacts, Conditions, Messaging, Schedule",
        "3. In Contacts, select the synced HubSpot list (e.g., '2026 - Interest: Yes/Maybe/Unknown')",
        "4. Write message -- keep it short, include the form link",
        "5. Preview estimated cost: three-dot menu > Preview before sending",
        "6. Schedule at a reasonable hour; avoid repeat sends to non-responders too quickly",
        "7. Send",
        "",
        "RECOMMENDED TEMPLATE (first outreach of year):",
        "NC 2026 check-in: are you coming + can you help? Fill this out: [FORM LINK]",
        "Reply STOP to opt out.",
    ]),
    (1, "5. Send a 1:1 SMS (from HubSpot)", [
        "Use this for personal follow-up with potential leaders or unclear respondents.",
        "",
        "1. In HubSpot, open the person's Contact record",
        "2. Find the Sakari SMS card/module",
        "3. Click Send SMS, write your message, send",
        "4. Replies are tracked on the contact record",
        "",
        "RECOMMENDED TEMPLATE (role recruitment):",
        "Hey -- you marked that you're open to leadership. We still need a [Role] Lead.",
        "Are you open to it (or co-leading)? I can send a simple checklist.",
        "",
        "Troubleshooting: If the SMS card is missing, confirm the Sakari-HubSpot integration",
        "is active and the card is enabled in HubSpot's integration settings.",
    ]),
    (1, "6. Contacts", [
        "- Communications Officer (Sakari owner): Daemon Wyner",
        "- President: Reece Dassinger",
        "- VP / SOP Owner: Chris Reddin",
    ]),
    (1, "7. Revision History", [
        "- v1.0 | 2026-03-03 | Initial draft | Chris Reddin",
    ]),
]

create_sop(
    output_path="Standard Operating Procedures/Com2 Sakari.docx",
    sop_number="Com2",
    sop_title="Sakari (SMS)",
    department="Communications",
    version="1.0",
    effective_date="2026-03-03",
    last_updated="2026-03-03",
    sections=sections_com2,
)

# =========================================================
# SOP Com3 — HubSpot
# =========================================================

sections_com3 = [
    (1, "1. Purpose", [
        "This SOP covers how Nipple Crime uses HubSpot for contact management, forms, "
        "email marketing (Nipple News), and website upkeep. HubSpot is the central CRM "
        "and outbound communications platform managed by the Communications Officer.",
    ]),
    (1, "2. Access & Permissions", [
        "Before doing anything else, verify you have the correct permissions.",
        "- Minimum required: Super Admin (or equivalent) access",
        "- Required access areas:",
        "  - CRM (Contacts)",
        "  - Marketing > Forms",
        "  - Marketing > Email",
        "  - Content / Website Pages",
        "  - Commerce (optional; usually Treasurer-owned)",
        "- To check: Settings > Users & Teams in HubSpot",
        "- If you cannot create/edit properties, lists, or forms -- fix permissions before troubleshooting anything else",
        "- Note: Property-level access can be restricted separately from general contact access",
    ]),
    (1, "3. Data Model", [
        "- Every person is a Contact record",
        "- Contact properties store structured answers (e.g., 'Interested in camp 2026?', 'Volunteer hours', 'Leadership role')",
        "- Forms collect and update properties at scale",
        "- Lists/Segments (active or static) target the right people for outreach",
    ]),
    (1, "4. Naming Conventions", [
        "Use consistent naming so properties and lists stay organized across years.",
        "",
        "PROPERTIES (prefix with year):",
        "- 2026 - Interested in Camp",
        "- 2026 - Volunteer Time Available",
        "- 2026 - Leadership Interest",
        "- 2026 - Leadership Role (Assigned)",
        "",
        "LISTS/SEGMENTS:",
        "- 2026 - Interest: Yes/Maybe/Unknown",
        "- 2026 - Interest: No",
        "- 2026 - Potential Leaders (Interested, No Role Assigned)",
        "- 2026 - Needs Follow-up (No response yet)",
        "",
        "List type rule:",
        "- Active lists: anything ongoing (auto-updates as responses come in)",
        "- Static lists: one-time snapshots (everyone invited to X at time of send)",
    ]),
    (1, "5. Create a Contact Property", [
        "Use this to add a structured field (e.g., a dropdown) for survey answers.",
        "",
        "1. In HubSpot, click Settings (gear icon)",
        "2. Go to Data Management > Properties",
        "3. Choose object type: Contact",
        "4. Click Create Property",
        "5. Set:",
        "   - Property label: e.g., 2026 - Interested in Camp",
        "   - Field type: Dropdown select",
        "   - Options: Yes, Maybe, No",
        "6. Save",
        "",
        "The property is now available in forms, contact views, and list filters.",
        "Troubleshooting: If 'Create property' is missing or throws errors, check Super Admin permissions.",
    ]),
    (1, "6. Build the Annual Interest / Labor / Leadership Form", [
        "Clone last year's form rather than building from scratch.",
        "",
        "1. Go to Marketing > Forms",
        "2. Find last year's 'Existing Member Interest' form",
        "3. Clone it",
        "4. Update text to 2026; keep it short",
        "5. Required fields:",
        "   - Email (required)",
        "   - 2026 - Interested in Camp (dropdown: Yes / Maybe / No)",
        "   - 2026 - Volunteer Time Available (dropdown)",
        "   - 2026 - Leadership Interest (dropdown)",
        "6. Recommended dropdown options for Volunteer Time Available:",
        "   - 1-2 hours/week",
        "   - 3-4 hours/week",
        "   - 4+ hours/week",
        "   - I have time mainly close to the burn",
        "   - No time / can't help this year",
        "7. Recommended dropdown options for Leadership Interest:",
        "   - Yes -- I'm interested in a leadership role",
        "   - Maybe -- talk to me",
        "   - No -- not this year",
        "8. Publish the form",
        "9. Get the share link: Marketing > Forms > hover form > Actions > Share > Copy link",
        "10. Quality check: submit it yourself once and confirm the properties updated on your contact record",
    ]),
    (1, "7. Create a Segment / List for Targeting", [
        "1. Go to CRM > Lists (Segments)",
        "2. Create a Contact-based list",
        "3. Choose list type (Active or Static -- see Section 4)",
        "4. Build filters",
        "",
        "EXAMPLE A: '2026 - Interest: Yes/Maybe/Unknown' (Active list)",
        "Filter: 2026 - Interested in Camp is any of: Yes, Maybe",
        "(and/or is unknown/empty for non-responders)",
        "",
        "EXAMPLE B: 'Potential Leaders (Willing + No Role Assigned)'",
        "Prereq: create '2026 - Leadership Role (Assigned)' property first",
        "Filter 1: 2026 - Leadership Interest is any of: Yes, Maybe",
        "Filter 2: AND 2026 - Leadership Role (Assigned) is unknown/empty",
        "Save as: '2026 - Potential Leaders (No Role Assigned)'",
    ]),
    (1, "8. Send Nipple News Email", [
        "Clone last year's closest matching email rather than building from scratch.",
        "",
        "1. Go to Marketing > Email",
        "2. Find last year's matching email (e.g., 'Are you coming?')",
        "3. Clone",
        "4. Edit: subject line (clear + action-oriented), body (short sections, bold deadlines), insert form link",
        "5. Preview as a real contact; send a test email to yourself",
        "6. Set Send to list: '2026 - Interest: Yes/Maybe/Unknown'",
        "7. Enable web version ('view in browser') so you can share the link on Facebook or via SMS",
        "8. Send",
        "",
        "WEB VERSION SHARE TEMPLATE (SMS/Facebook):",
        "Nipple News #1 is out -- full details here: [WEB VERSION LINK]",
    ]),
    (1, "9. Get the Web Version Link", [
        "1. Open the email in HubSpot",
        "2. Go to email details / web version area",
        "3. Copy the web version URL",
        "Use this link for: Facebook group posts, SMS 'read more' links, members who don't read email.",
    ]),
    (1, "10. Website Upkeep (HubSpot CMS)", [
        "Keep Sign Up, Donate, and other key pages current each year.",
        "",
        "1. Go to Content > Website Pages",
        "2. Click the page to edit",
        "3. Update text, dates, and buttons",
        "4. Publish (or schedule publish)",
        "",
        "If temporarily disabling a button (e.g., donations not open yet):",
        "- Remove the button URL, OR",
        "- Change the button URL to the homepage to avoid a dead link",
    ]),
    (1, "11. Payment Links (Know Where They Live)", [
        "Payment links are owned by the Treasurer but Comms touches them when updating pages or emails.",
        "Location: HubSpot > Commerce > Payment Links",
        "",
        "Comms responsibilities:",
        "- Update broken or mislabeled buttons on web pages (notify Treasurer when doing so)",
        "- Do NOT change pricing or payment plan logic without explicit Treasurer approval",
    ]),
    (1, "12. Handoff Checklist (for Next Comms Officer)", [
        "Keep this updated throughout the year so handoff is smooth.",
        "",
        "- List of active properties created this year (names + purpose)",
        "- List/segment names used",
        "- Which form is the canonical interest form for this year",
        "- Message templates:",
        "  - First interest SMS",
        "  - Leadership recruitment SMS",
        "  - Nipple News email skeleton",
        "- Where web version links live (and how to generate them)",
        "- Website pages most often updated (Sign Up, Donate, etc.)",
        "- Who to contact for permissions: President / Super Admin",
    ]),
    (1, "13. Contacts", [
        "- Communications Officer (HubSpot owner): Daemon Wyner",
        "- President (approves major messaging + payment decisions): Reece Dassinger",
        "- Treasurer (payment links, dues language): Isabel Hoy",
        "- VP / SOP Owner: Chris Reddin",
    ]),
    (1, "14. Revision History", [
        "- v1.0 | 2026-03-03 | Initial draft | Chris Reddin",
    ]),
]

create_sop(
    output_path="Standard Operating Procedures/Com3 HubSpot.docx",
    sop_number="Com3",
    sop_title="HubSpot",
    department="Communications",
    version="1.0",
    effective_date="2026-03-03",
    last_updated="2026-03-03",
    sections=sections_com3,
)
