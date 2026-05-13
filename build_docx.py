"""Build Beacon-Capabilities-Statement.docx with python-docx.

Mirrors the HTML one-pager: letterhead with yellow B mark, lede,
three core-capability cards, a warehousing strip, meta row, and dark CTA bar.
"""
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn, nsmap
from docx.oxml import OxmlElement

# Brand
YELLOW = "F5B82E"
CREAM = "FCF6E6"
NAVY = "1B2738"
CARD_BG = "FFFDF5"


def set_cell_bg(cell, hex_color):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tc_pr.append(shd)


def set_cell_border(cell, **kwargs):
    """kwargs like top={'sz': 8, 'color': '1B2738', 'val': 'single'}"""
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_borders = tc_pr.find(qn("w:tcBorders"))
    if tc_borders is None:
        tc_borders = OxmlElement("w:tcBorders")
        tc_pr.append(tc_borders)
    for edge in ("top", "left", "bottom", "right"):
        if edge in kwargs:
            spec = kwargs[edge]
            existing = tc_borders.find(qn(f"w:{edge}"))
            if existing is not None:
                tc_borders.remove(existing)
            el = OxmlElement(f"w:{edge}")
            el.set(qn("w:val"), spec.get("val", "single"))
            el.set(qn("w:sz"), str(spec.get("sz", 4)))
            el.set(qn("w:space"), "0")
            el.set(qn("w:color"), spec.get("color", "000000"))
            tc_borders.append(el)


def add_run(p, text, *, size=10, bold=False, color=NAVY, font="Helvetica", letter_spacing=None, caps=False):
    r = p.add_run(text)
    r.font.name = font
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = RGBColor.from_string(color)
    if letter_spacing is not None:
        rPr = r._element.get_or_add_rPr()
        spacing = OxmlElement("w:spacing")
        spacing.set(qn("w:val"), str(letter_spacing))
        rPr.append(spacing)
    if caps:
        rPr = r._element.get_or_add_rPr()
        caps_el = OxmlElement("w:caps")
        caps_el.set(qn("w:val"), "1")
        rPr.append(caps_el)
    return r


def remove_paragraph_spacing(p):
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = 1.15


def set_page_bg(doc, hex_color):
    settings = doc.settings.element
    bg = OxmlElement("w:background")
    bg.set(qn("w:color"), hex_color)
    # background must be the first child of document; but for settings approach:
    doc.element.insert(0, bg)
    # Also enable display background in settings
    display_bg = OxmlElement("w:displayBackgroundShape")
    settings.append(display_bg)


def main():
    doc = Document()

    # Page setup: letter, narrow margins
    section = doc.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)
    section.left_margin = Inches(0.55)
    section.right_margin = Inches(0.55)

    set_page_bg(doc, CREAM)

    # Base style
    style = doc.styles["Normal"]
    style.font.name = "Helvetica"
    style.font.size = Pt(10)
    style.font.color.rgb = RGBColor.from_string(NAVY)

    # ----- LETTERHEAD (2 cols: brand left, address right) -----
    head = doc.add_table(rows=1, cols=2)
    head.autofit = False
    head.columns[0].width = Inches(4.6)
    head.columns[1].width = Inches(2.8)

    # Brand cell: nested table for B mark + name
    brand_cell = head.cell(0, 0)
    brand_cell.width = Inches(4.6)
    bp = brand_cell.paragraphs[0]
    remove_paragraph_spacing(bp)

    brand_inner = brand_cell.add_table(rows=1, cols=2)
    brand_inner.autofit = False
    brand_inner.columns[0].width = Inches(0.75)
    brand_inner.columns[1].width = Inches(3.7)
    b_mark = brand_inner.cell(0, 0)
    b_mark.width = Inches(0.75)
    set_cell_bg(b_mark, YELLOW)
    bm_p = b_mark.paragraphs[0]
    bm_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    remove_paragraph_spacing(bm_p)
    bm_p.paragraph_format.space_before = Pt(2)
    add_run(bm_p, "B", size=34, bold=True, color=NAVY, font="Helvetica")
    b_mark.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    name_cell = brand_inner.cell(0, 1)
    name_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    eyebrow_p = name_cell.paragraphs[0]
    remove_paragraph_spacing(eyebrow_p)
    add_run(eyebrow_p, "CAPABILITIES STATEMENT", size=8, bold=True, color=NAVY, letter_spacing=40)
    name_p = name_cell.add_paragraph()
    remove_paragraph_spacing(name_p)
    add_run(name_p, "Beacon Manufacturing", size=20, bold=True, color=NAVY)

    # Remove borders on brand_inner
    for row in brand_inner.rows:
        for c in row.cells:
            set_cell_border(c,
                top={'val': 'nil'}, left={'val': 'nil'},
                bottom={'val': 'nil'}, right={'val': 'nil'})

    # Address cell
    addr = head.cell(0, 1)
    addr.width = Inches(2.8)
    addr.vertical_alignment = WD_ALIGN_VERTICAL.TOP
    ap = addr.paragraphs[0]
    ap.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    remove_paragraph_spacing(ap)
    add_run(ap, "DETROIT, MICHIGAN", size=9, bold=True, color=NAVY, letter_spacing=30)
    for line in ["2200 Beaufait Avenue", "Detroit, MI 48207",
                 "hello@beaconmfg.com · (313) 555-0184", "beaconmfg.com"]:
        p = addr.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        remove_paragraph_spacing(p)
        add_run(p, line, size=9, color=NAVY)

    # Letterhead bottom border (navy)
    for c in head.rows[0].cells:
        set_cell_border(c,
            top={'val': 'nil'}, left={'val': 'nil'}, right={'val': 'nil'},
            bottom={'val': 'single', 'sz': 16, 'color': NAVY})

    # Spacer
    sp = doc.add_paragraph()
    remove_paragraph_spacing(sp)
    sp.paragraph_format.space_after = Pt(8)

    # ----- LEDE -----
    h1 = doc.add_paragraph()
    remove_paragraph_spacing(h1)
    add_run(h1, "Metal fabrication, assembly, and prototyping ", size=18, bold=True, color=NAVY)
    add_run(h1, "—", size=18, bold=True, color=YELLOW)
    add_run(h1, " made in Detroit.", size=18, bold=True, color=NAVY)

    lede = doc.add_paragraph()
    remove_paragraph_spacing(lede)
    lede.paragraph_format.space_before = Pt(4)
    lede.paragraph_format.space_after = Pt(6)
    add_run(lede,
        "Beacon Manufacturing is a Detroit-based contract manufacturer pairing decades of "
        "skilled-trade craftsmanship with modern digital fabrication. From one-off prototypes "
        "to recurring production runs, we cut, form, weld, finish, assemble, kit, and ship — "
        "all under one roof. Short lead times, tight tolerances, and a single point of "
        "accountability for engineers, founders, and procurement teams.",
        size=10, color=NAVY)

    # ----- CORE CAPABILITIES title -----
    st = doc.add_paragraph()
    remove_paragraph_spacing(st)
    st.paragraph_format.space_before = Pt(6)
    st.paragraph_format.space_after = Pt(4)
    add_run(st, "CORE CAPABILITIES", size=9, bold=True, color=NAVY, letter_spacing=50)

    # ----- THREE CARDS -----
    cards = doc.add_table(rows=1, cols=3)
    cards.autofit = False
    col_w = Inches(2.43)
    for col in cards.columns:
        col.width = col_w

    card_data = [
        ("Fabrication",
         "Precision sheet-metal and structural metalwork from raw stock to finished part.",
         ["Fiber laser cutting (up to 1/2\")",
          "CNC press brake forming",
          "MIG / TIG / spot welding",
          "Powder coat & wet paint",
          "Grinding, deburring, polishing"]),
        ("Assembly",
         "Sub-assembly and finished-goods build with mechanical, electrical, and pneumatic integration.",
         ["Mechanical & fastener assembly",
          "Electrical harness & panel build",
          "Pneumatic / hydraulic integration",
          "Functional test & QA",
          "Kitting & final packaging"]),
        ("Prototyping & DigiFab",
         "Rapid iteration with digital fabrication for fast turns on form, fit, and function.",
         ["3D printing (FDM, SLA, SLS)",
          "CNC routing & milling",
          "Laser engraving & cutting",
          "DFM review & design support",
          "Low-volume pilot runs"]),
    ]

    for i, (title, blurb, items) in enumerate(card_data):
        cell = cards.cell(0, i)
        cell.width = col_w
        set_cell_bg(cell, CARD_BG)
        set_cell_border(cell,
            top={'val': 'single', 'sz': 4, 'color': 'E5DDC0'},
            right={'val': 'single', 'sz': 4, 'color': 'E5DDC0'},
            bottom={'val': 'single', 'sz': 4, 'color': 'E5DDC0'},
            left={'val': 'single', 'sz': 36, 'color': YELLOW})
        # Title
        cp0 = cell.paragraphs[0]
        remove_paragraph_spacing(cp0)
        cp0.paragraph_format.space_before = Pt(2)
        add_run(cp0, title, size=12, bold=True, color=NAVY)
        # Blurb
        bp_ = cell.add_paragraph()
        remove_paragraph_spacing(bp_)
        bp_.paragraph_format.space_before = Pt(2)
        bp_.paragraph_format.space_after = Pt(4)
        add_run(bp_, blurb, size=9, color=NAVY)
        # Bullets
        for item in items:
            ip = cell.add_paragraph()
            remove_paragraph_spacing(ip)
            ip.paragraph_format.left_indent = Inches(0.08)
            add_run(ip, "■ ", size=8, color=YELLOW, bold=True)
            add_run(ip, item, size=9, color=NAVY)
        # trailing padding
        pad = cell.add_paragraph()
        remove_paragraph_spacing(pad)
        pad.paragraph_format.space_after = Pt(2)

    # ----- WAREHOUSING title -----
    st2 = doc.add_paragraph()
    remove_paragraph_spacing(st2)
    st2.paragraph_format.space_before = Pt(8)
    st2.paragraph_format.space_after = Pt(4)
    add_run(st2, "WAREHOUSING & LOGISTICS", size=9, bold=True, color=NAVY, letter_spacing=50)

    # ----- STRIP (2 cols inside one card-like row) -----
    strip = doc.add_table(rows=1, cols=2)
    strip.autofit = False
    strip.columns[0].width = Inches(1.7)
    strip.columns[1].width = Inches(5.7)

    s_left = strip.cell(0, 0)
    s_right = strip.cell(0, 1)
    s_left.width = Inches(1.7)
    s_right.width = Inches(5.7)
    set_cell_bg(s_left, CARD_BG)
    set_cell_bg(s_right, CARD_BG)
    set_cell_border(s_left,
        top={'val': 'single', 'sz': 4, 'color': 'E5DDC0'},
        bottom={'val': 'single', 'sz': 4, 'color': 'E5DDC0'},
        right={'val': 'nil'},
        left={'val': 'single', 'sz': 36, 'color': NAVY})
    set_cell_border(s_right,
        top={'val': 'single', 'sz': 4, 'color': 'E5DDC0'},
        bottom={'val': 'single', 'sz': 4, 'color': 'E5DDC0'},
        left={'val': 'nil'},
        right={'val': 'single', 'sz': 4, 'color': 'E5DDC0'})

    sl_p = s_left.paragraphs[0]
    remove_paragraph_spacing(sl_p)
    sl_p.paragraph_format.space_before = Pt(2)
    add_run(sl_p, "One roof, end to end.", size=11, bold=True, color=NAVY)
    sl_p2 = s_left.add_paragraph()
    remove_paragraph_spacing(sl_p2)
    add_run(sl_p2, "42,000 sq ft Detroit facility", size=8, color=NAVY)

    sr_p = s_right.paragraphs[0]
    remove_paragraph_spacing(sr_p)
    sr_p.paragraph_format.space_before = Pt(2)
    add_run(sr_p,
        "On-site warehousing, inventory programs, and outbound logistics keep your supply "
        "chain tight. We hold stock, manage Kanban replenishment, and ship direct to your "
        "line, your customer, or your DC.",
        size=9, color=NAVY)

    # Pills (yellow filled + outline)
    pills_p = s_right.add_paragraph()
    remove_paragraph_spacing(pills_p)
    pills_p.paragraph_format.space_before = Pt(4)
    filled = ["PICK & PACK", "KITTING", "KANBAN", "VMI", "INVENTORY MGMT"]
    outline = ["LTL / PARCEL", "DROP SHIP", "CROSS-DOCK"]
    for i, label in enumerate(filled):
        r = add_run(pills_p, f"  {label}  ", size=8, bold=True, color=NAVY, letter_spacing=20)
        rPr = r._element.get_or_add_rPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), YELLOW)
        rPr.append(shd)
        if i != len(filled) - 1 or outline:
            add_run(pills_p, "  ", size=8)
    for i, label in enumerate(outline):
        add_run(pills_p, f"[ {label} ]", size=8, bold=True, color=NAVY, letter_spacing=20)
        if i != len(outline) - 1:
            add_run(pills_p, "  ", size=8)

    # ----- META ROW -----
    meta = doc.add_table(rows=1, cols=3)
    meta.autofit = False
    for col in meta.columns:
        col.width = Inches(2.43)
    meta_data = [
        ("INDUSTRIES",
         "Automotive · Mobility · Industrial OEM · Consumer hardware · Architectural · Defense"),
        ("CERTIFICATIONS",
         "ISO 9001:2015 · AWS D1.1 welders · ITAR registered · Detroit-based small business"),
        ("NAICS",
         "332710 · 332312 · 332999 · 333249 · 488991"),
    ]
    for i, (label, value) in enumerate(meta_data):
        c = meta.cell(0, i)
        c.width = Inches(2.43)
        set_cell_border(c,
            top={'val': 'single', 'sz': 16, 'color': NAVY},
            left={'val': 'nil'}, right={'val': 'nil'}, bottom={'val': 'nil'})
        p0 = c.paragraphs[0]
        remove_paragraph_spacing(p0)
        p0.paragraph_format.space_before = Pt(4)
        add_run(p0, label, size=8, bold=True, color=NAVY, letter_spacing=40)
        p1 = c.add_paragraph()
        remove_paragraph_spacing(p1)
        add_run(p1, value, size=9, color=NAVY)

    # spacer push to bottom
    spacer = doc.add_paragraph()
    remove_paragraph_spacing(spacer)
    spacer.paragraph_format.space_after = Pt(10)

    # ----- CTA BAR -----
    cta = doc.add_table(rows=1, cols=2)
    cta.autofit = False
    cta.columns[0].width = Inches(4.5)
    cta.columns[1].width = Inches(2.9)
    cta_l = cta.cell(0, 0)
    cta_r = cta.cell(0, 1)
    cta_l.width = Inches(4.5)
    cta_r.width = Inches(2.9)
    set_cell_bg(cta_l, NAVY)
    set_cell_bg(cta_r, NAVY)
    set_cell_border(cta_l,
        left={'val': 'single', 'sz': 48, 'color': YELLOW},
        top={'val': 'nil'}, bottom={'val': 'nil'}, right={'val': 'nil'})
    set_cell_border(cta_r,
        left={'val': 'nil'}, top={'val': 'nil'}, bottom={'val': 'nil'}, right={'val': 'nil'})

    cl_p = cta_l.paragraphs[0]
    remove_paragraph_spacing(cl_p)
    cl_p.paragraph_format.space_before = Pt(6)
    add_run(cl_p, "LET'S BUILD IT", size=9, bold=True, color=YELLOW, letter_spacing=50)
    cl_p2 = cta_l.add_paragraph()
    remove_paragraph_spacing(cl_p2)
    cl_p2.paragraph_format.space_after = Pt(6)
    add_run(cl_p2, "Send a print. Get a quote in 48 hours.", size=13, bold=True, color=CREAM)

    cr_p = cta_r.paragraphs[0]
    cr_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    remove_paragraph_spacing(cr_p)
    cr_p.paragraph_format.space_before = Pt(8)
    add_run(cr_p, "quotes@beaconmfg.com", size=11, bold=True, color=YELLOW)
    cr_p2 = cta_r.add_paragraph()
    cr_p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    remove_paragraph_spacing(cr_p2)
    cr_p2.paragraph_format.space_after = Pt(6)
    add_run(cr_p2, "(313) 555-0184 · beaconmfg.com", size=10, color=CREAM)

    doc.save("Beacon-Capabilities-Statement.docx")
    print("DOCX generated")


if __name__ == "__main__":
    main()
