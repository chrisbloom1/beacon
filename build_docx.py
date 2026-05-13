"""Build Beacon-Capabilities-Statement.docx with python-docx.

Mirrors capabilities.html: dark teal masthead with yellow BEACON wordmark
block, lede with real Beacon mark image, three core-capability cards with
yellow top accents and dark index tabs, dark warehousing strip with source
bullets, yellow CTA bar, dark stamp footer.

Brand: yellow #fdb602, cream paper #fdf6e4, deep teal ink #022429, sand
#c9b789. Display: Barlow Condensed. Body: Figtree.
"""
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

YELLOW = "FDB602"
CREAM = "FDF6E4"
TEAL = "022429"
SAND = "C9B789"
WHITE = "FFFFFF"

DISPLAY = "Barlow Condensed"
BODY = "Figtree"


def set_cell_bg(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def set_cell_border(cell, **kwargs):
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = tcPr.find(qn("w:tcBorders"))
    if tcBorders is None:
        tcBorders = OxmlElement("w:tcBorders")
        tcPr.append(tcBorders)
    for edge in ("top", "left", "bottom", "right"):
        if edge in kwargs:
            spec = kwargs[edge]
            existing = tcBorders.find(qn(f"w:{edge}"))
            if existing is not None:
                tcBorders.remove(existing)
            el = OxmlElement(f"w:{edge}")
            el.set(qn("w:val"), spec.get("val", "single"))
            el.set(qn("w:sz"), str(spec.get("sz", 4)))
            el.set(qn("w:space"), "0")
            el.set(qn("w:color"), spec.get("color", "000000"))
            tcBorders.append(el)


def set_cell_margins(cell, top=80, bottom=80, left=120, right=120):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for edge, val in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        m = OxmlElement(f"w:{edge}")
        m.set(qn("w:w"), str(val))
        m.set(qn("w:type"), "dxa")
        tcMar.append(m)
    tcPr.append(tcMar)


def remove_table_borders(table):
    tblPr = table._element.tblPr
    tblBorders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        b = OxmlElement(f"w:{edge}")
        b.set(qn("w:val"), "nil")
        tblBorders.append(b)
    tblPr.append(tblBorders)


def add_run(p, text, *, size=10, bold=False, italic=False, color=TEAL,
            font=BODY, letter_spacing=None, shade=None):
    r = p.add_run(text)
    rPr = r._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), font)
    rFonts.set(qn("w:hAnsi"), font)
    rFonts.set(qn("w:cs"), font)
    r.font.name = font
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.italic = italic
    r.font.color.rgb = RGBColor.from_string(color)
    if letter_spacing is not None:
        spacing = OxmlElement("w:spacing")
        spacing.set(qn("w:val"), str(letter_spacing))
        rPr.append(spacing)
    if shade is not None:
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), shade)
        rPr.append(shd)
    return r


def tighten(p, before=0, after=0, line=1.15):
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line


def set_page_bg(doc, hex_color):
    bg = OxmlElement("w:background")
    bg.set(qn("w:color"), hex_color)
    doc.element.insert(0, bg)
    display_bg = OxmlElement("w:displayBackgroundShape")
    doc.settings.element.append(display_bg)


def main():
    doc = Document()
    section = doc.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    section.top_margin = Inches(0)
    section.bottom_margin = Inches(0)
    section.left_margin = Inches(0)
    section.right_margin = Inches(0)
    set_page_bg(doc, CREAM)

    base = doc.styles["Normal"]
    base.font.name = BODY
    base.font.size = Pt(10)
    base.font.color.rgb = RGBColor.from_string(TEAL)

    # ============= MASTHEAD =============
    head = doc.add_table(rows=1, cols=3)
    head.autofit = False
    head.columns[0].width = Inches(1.85)
    head.columns[1].width = Inches(3.65)
    head.columns[2].width = Inches(3.0)
    remove_table_borders(head)

    # Wordmark cell — yellow box with BEACON
    wm = head.cell(0, 0)
    wm.width = Inches(1.85)
    set_cell_bg(wm, TEAL)
    set_cell_margins(wm, top=180, bottom=180, left=420, right=120)
    wm.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    # Nested table to make a yellow chip inside teal masthead
    chip_tbl = wm.add_table(rows=1, cols=1)
    chip_tbl.autofit = False
    chip_tbl.columns[0].width = Inches(1.3)
    remove_table_borders(chip_tbl)
    chip = chip_tbl.cell(0, 0)
    chip.width = Inches(1.3)
    set_cell_bg(chip, YELLOW)
    set_cell_margins(chip, top=140, bottom=100, left=160, right=160)
    cp = chip.paragraphs[0]
    cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tighten(cp)
    add_run(cp, "BEACON", size=30, bold=True, color=TEAL, font=DISPLAY,
            letter_spacing=20)

    # Middle: eyebrow + doc title + sub
    mid = head.cell(0, 1)
    mid.width = Inches(3.65)
    set_cell_bg(mid, TEAL)
    set_cell_margins(mid, top=240, bottom=180, left=240, right=120)
    mid.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    mp0 = mid.paragraphs[0]
    tighten(mp0)
    add_run(mp0, "CAPABILITIES STATEMENT", size=11, bold=True, color=YELLOW,
            font=DISPLAY, letter_spacing=80)
    mp1 = mid.add_paragraph()
    tighten(mp1, before=2)
    add_run(mp1, "BEACON MANUFACTURING", size=20, bold=True, color=CREAM,
            font=DISPLAY)
    mp2 = mid.add_paragraph()
    tighten(mp2, before=3)
    add_run(mp2, "BUILDER-LED MFG. IN DETROIT · V2", size=8, bold=True,
            color=SAND, font=BODY, letter_spacing=50)

    # Right: address block on teal
    rt = head.cell(0, 2)
    rt.width = Inches(3.0)
    set_cell_bg(rt, TEAL)
    set_cell_margins(rt, top=240, bottom=180, left=120, right=320)
    rt.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    for line in ["2050 15th St", "Detroit, MI 48216"]:
        p = rt.add_paragraph() if line != "2050 15th St" else rt.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        tighten(p)
        add_run(p, line, size=9.5, color=CREAM, font=BODY)
    url_p = rt.add_paragraph()
    url_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    tighten(url_p, before=2)
    add_run(url_p, "BEACONMFG.US", size=14, bold=True, color=YELLOW,
            font=DISPLAY, letter_spacing=30)
    tag_p = rt.add_paragraph()
    tag_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    tighten(tag_p, before=2)
    add_run(tag_p, "LOCAL MFG. · 100% ONSHORE", size=8, bold=True,
            color=SAND, font=BODY, letter_spacing=40)

    # Yellow accent strip below masthead (single-row table)
    accent = doc.add_table(rows=1, cols=1)
    accent.autofit = False
    accent.columns[0].width = Inches(8.5)
    remove_table_borders(accent)
    ac = accent.cell(0, 0)
    ac.width = Inches(8.5)
    set_cell_bg(ac, YELLOW)
    set_cell_margins(ac, top=40, bottom=40, left=0, right=0)
    ap = ac.paragraphs[0]
    tighten(ap)
    # empty bar acts as the 8px yellow rule

    # ============= PAPER (inset content) =============
    # Use a single-column table that has white-cream cell to inset content
    paper = doc.add_table(rows=1, cols=1)
    paper.autofit = False
    paper.columns[0].width = Inches(8.5)
    remove_table_borders(paper)
    pcell = paper.cell(0, 0)
    pcell.width = Inches(8.5)
    set_cell_bg(pcell, CREAM)
    set_cell_margins(pcell, top=260, bottom=120, left=720, right=720)

    # Build content inside pcell. Need to add nested tables/paragraphs.

    # ---- LEDE (image + text in nested 2-col table)
    lede_tbl = pcell.add_table(rows=1, cols=2)
    lede_tbl.autofit = False
    lede_tbl.columns[0].width = Inches(0.95)
    lede_tbl.columns[1].width = Inches(6.4)
    remove_table_borders(lede_tbl)

    mark_cell = lede_tbl.cell(0, 0)
    mark_cell.width = Inches(0.95)
    mark_cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
    mark_p = mark_cell.paragraphs[0]
    tighten(mark_p)
    mark_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    try:
        mark_p.add_run().add_picture("assets/beacon-mark.png", width=Inches(0.82))
    except Exception:
        set_cell_bg(mark_cell, YELLOW)
        add_run(mark_p, "B", size=36, bold=True, color=TEAL, font=DISPLAY)

    lcell = lede_tbl.cell(0, 1)
    lcell.width = Inches(6.4)
    lcell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
    h1 = lcell.paragraphs[0]
    tighten(h1)
    add_run(h1, "METAL FABRICATION, ASSEMBLY, AND ", size=22, bold=True,
            color=TEAL, font=DISPLAY)
    add_run(h1, "  RAPID PROTOTYPING.  ", size=22, bold=True, color=TEAL,
            font=DISPLAY, shade=YELLOW)

    lp = lcell.add_paragraph()
    tighten(lp, before=4, line=1.45)
    add_run(lp, "Beacon is a builder-led contract manufacturer in Detroit ",
            size=10.5, bold=True, color=TEAL)
    add_run(lp, "delivering metal fabrication, assembly, and rapid prototyping "
                "— from one-off prototypes to repeatable short runs and full "
                "production. ", size=10.5, color=TEAL)
    add_run(lp, "  Bring us a problem.  ", size=10.5, bold=True, color=TEAL,
            shade=YELLOW)

    # spacer inside paper cell
    pad1 = pcell.add_paragraph()
    tighten(pad1, after=6)

    # ---- CORE CAPABILITIES title (bar + label + sub)
    title = pcell.add_paragraph()
    tighten(title, before=4, after=2)
    add_run(title, "▮▮ ", size=14, bold=True, color=YELLOW, font=BODY)
    add_run(title, "CORE CAPABILITIES   ", size=15, bold=True, color=TEAL,
            font=DISPLAY, letter_spacing=70)
    add_run(title, "Fabrication · Assembly · Prototyping", size=9,
            italic=True, color=TEAL)

    # underline rule
    rule = pcell.add_paragraph()
    tighten(rule, after=6)
    pPr = rule._element.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "12")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), TEAL)
    pBdr.append(bottom)
    pPr.append(pBdr)

    # ---- CARDS
    cards = pcell.add_table(rows=1, cols=3)
    cards.autofit = False
    col_w = Inches(2.43)
    for col in cards.columns:
        col.width = col_w
    remove_table_borders(cards)

    cards_data = [
        ("01 · FAB", "FABRICATION",
         "Cutting, forming, and welding for metal parts and structures.",
         [("Laser cutting", " — tube + flat"),
          ("Forming", " — brake press, tube bending, punch & press"),
          ("Welding", " — MIG / TIG / laser + robotic"),
          ("Short runs", " — jigs & fixtures for repeatability")]),
        ("02 · ASM", "ASSEMBLY",
         "Kitting, sub-assembly, full builds — and everything in between.",
         [(None, "Kitting + sub-assembly + final assembly"),
          (None, "Complete build-ups with CAD / 3xD support"),
          (None, "Torque & fastener standards · EOL test"),
          (None, "Quality checks + rework when needed")]),
        ("03 · PROTO", "PROTOTYPING & DIGIFAB",
         "Fast prototypes and fixtures, CAD to part.",
         [("3D printing", " — FDM, SLA, SLS (EOS)"),
          (None, "Quick jigs, check-fixtures, POCs, mockups"),
          (None, "Same-week iteration on form, fit, function")]),
    ]

    for i, (num, title_, blurb, items) in enumerate(cards_data):
        cell = cards.cell(0, i)
        cell.width = col_w
        set_cell_bg(cell, WHITE)
        set_cell_border(cell,
            top={"val": "single", "sz": 36, "color": YELLOW},
            right={"val": "single", "sz": 4, "color": "E8DFC0"},
            bottom={"val": "single", "sz": 4, "color": "E8DFC0"},
            left={"val": "single", "sz": 4, "color": "E8DFC0"})
        set_cell_margins(cell, top=200, bottom=180, left=200, right=200)

        # index tab (text on teal, top of card)
        p0 = cell.paragraphs[0]
        tighten(p0)
        add_run(p0, f"  {num}  ", size=9, bold=True, color=YELLOW,
                font=DISPLAY, letter_spacing=40, shade=TEAL)

        ph = cell.add_paragraph()
        tighten(ph, before=2, after=2)
        add_run(ph, title_, size=18, bold=True, color=TEAL, font=DISPLAY)

        pt = cell.add_paragraph()
        tighten(pt, after=4, line=1.4)
        add_run(pt, blurb, size=9, italic=True, color=TEAL)

        for bold_part, rest in items:
            ip = cell.add_paragraph()
            tighten(ip, line=1.4)
            ip.paragraph_format.left_indent = Inches(0.08)
            add_run(ip, "■ ", size=8, bold=True, color=YELLOW, font=BODY)
            if bold_part:
                add_run(ip, bold_part, size=9.5, bold=True, color=TEAL)
                add_run(ip, rest, size=9.5, color=TEAL)
            else:
                add_run(ip, rest, size=9.5, color=TEAL)

    pad2 = pcell.add_paragraph()
    tighten(pad2, after=4)

    # ---- WAREHOUSING title
    title2 = pcell.add_paragraph()
    tighten(title2, before=8, after=2)
    add_run(title2, "▮▮ ", size=14, bold=True, color=YELLOW, font=BODY)
    add_run(title2, "WAREHOUSING & LOGISTICS   ", size=15, bold=True,
            color=TEAL, font=DISPLAY, letter_spacing=70)
    add_run(title2, "Products in, products out — without the chaos.",
            size=9, italic=True, color=TEAL)

    rule2 = pcell.add_paragraph()
    tighten(rule2, after=6)
    pPr2 = rule2._element.get_or_add_pPr()
    pBdr2 = OxmlElement("w:pBdr")
    b2 = OxmlElement("w:bottom")
    b2.set(qn("w:val"), "single")
    b2.set(qn("w:sz"), "12")
    b2.set(qn("w:space"), "1")
    b2.set(qn("w:color"), TEAL)
    pBdr2.append(b2)
    pPr2.append(pBdr2)

    # ---- STRIP (dark teal, 2 cols)
    strip = pcell.add_table(rows=1, cols=2)
    strip.autofit = False
    strip.columns[0].width = Inches(2.1)
    strip.columns[1].width = Inches(5.2)
    remove_table_borders(strip)

    sl = strip.cell(0, 0)
    sr = strip.cell(0, 1)
    sl.width = Inches(2.1)
    sr.width = Inches(5.2)
    set_cell_bg(sl, TEAL)
    set_cell_bg(sr, TEAL)
    set_cell_margins(sl, top=200, bottom=200, left=220, right=160)
    set_cell_margins(sr, top=200, bottom=200, left=200, right=200)
    set_cell_border(sl,
        left={"val": "single", "sz": 48, "color": YELLOW},
        top={"val": "nil"}, bottom={"val": "nil"}, right={"val": "nil"})
    set_cell_border(sr,
        left={"val": "single", "sz": 6, "color": YELLOW},
        top={"val": "nil"}, bottom={"val": "nil"}, right={"val": "nil"})

    slp = sl.paragraphs[0]
    tighten(slp)
    add_run(slp, "ONE ROOF.", size=18, bold=True, color=CREAM, font=DISPLAY)
    slp2 = sl.add_paragraph()
    tighten(slp2)
    add_run(slp2, "END TO END.", size=18, bold=True, color=CREAM, font=DISPLAY)
    slp3 = sl.add_paragraph()
    tighten(slp3, before=4)
    add_run(slp3, "Detroit · 2050 15th St", size=8, italic=True, color=SAND)

    bullets = [
        ("Receiving, inventory, and part management", ""),
        ("Parcel + freight", " — including no-box, roll-on / roll-off"),
        ("Flexible storage", " on the floor"),
    ]
    first = True
    for bold_part, rest in bullets:
        if first:
            ip = sr.paragraphs[0]
            first = False
        else:
            ip = sr.add_paragraph()
        tighten(ip, line=1.5)
        ip.paragraph_format.left_indent = Inches(0.1)
        add_run(ip, "■ ", size=9, bold=True, color=YELLOW, font=BODY)
        add_run(ip, bold_part, size=10, bold=True, color=CREAM)
        if rest:
            add_run(ip, rest, size=10, color=CREAM)

    # ============= CTA (after paper inset) =============
    # Black-bordered yellow CTA, full-width
    cta = doc.add_table(rows=1, cols=2)
    cta.autofit = False
    cta.columns[0].width = Inches(5.2)
    cta.columns[1].width = Inches(3.3)
    remove_table_borders(cta)
    cl = cta.cell(0, 0)
    cr = cta.cell(0, 1)
    cl.width = Inches(5.2)
    cr.width = Inches(3.3)
    set_cell_bg(cl, YELLOW)
    set_cell_bg(cr, YELLOW)
    set_cell_margins(cl, top=260, bottom=260, left=720, right=160)
    set_cell_margins(cr, top=260, bottom=260, left=160, right=720)
    set_cell_border(cl,
        top={"val": "single", "sz": 48, "color": TEAL},
        left={"val": "nil"}, bottom={"val": "nil"}, right={"val": "nil"})
    set_cell_border(cr,
        top={"val": "single", "sz": 48, "color": TEAL},
        left={"val": "nil"}, bottom={"val": "nil"}, right={"val": "nil"})

    clp = cl.paragraphs[0]
    tighten(clp)
    add_run(clp, "HAVE A PART, KIT, OR ASSEMBLY IN MIND?",
            size=13, bold=True, color=TEAL, font=DISPLAY, letter_spacing=60)
    clp2 = cl.add_paragraph()
    tighten(clp2, before=4, line=1.2)
    add_run(clp2, "Send a print, a CAD file, or a napkin sketch. ",
            size=14, color=TEAL, font=DISPLAY)
    add_run(clp2, "We'll quote, prototype, and put it on a line.",
            size=14, bold=True, color=TEAL, font=DISPLAY)

    crp = cr.paragraphs[0]
    crp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    tighten(crp)
    add_run(crp, "BEACONMFG.US", size=22, bold=True, color=TEAL,
            font=DISPLAY, letter_spacing=30)
    crp2 = cr.add_paragraph()
    crp2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    tighten(crp2, before=2)
    add_run(crp2, "2050 15th St · Detroit, MI 48216",
            size=9, bold=True, color=TEAL)

    # ============= STAMP =============
    stamp = doc.add_table(rows=1, cols=3)
    stamp.autofit = False
    stamp.columns[0].width = Inches(2.5)
    stamp.columns[1].width = Inches(3.5)
    stamp.columns[2].width = Inches(2.5)
    remove_table_borders(stamp)
    for i in range(3):
        c = stamp.cell(0, i)
        set_cell_bg(c, TEAL)
        set_cell_margins(c, top=120, bottom=120, left=720 if i == 0 else 80,
                          right=720 if i == 2 else 80)
    s0 = stamp.cell(0, 0).paragraphs[0]
    tighten(s0)
    s0.alignment = WD_ALIGN_PARAGRAPH.LEFT
    add_run(s0, "BEACON © 2026", size=7.5, bold=True, color=SAND, font=BODY,
            letter_spacing=40)
    s1 = stamp.cell(0, 1).paragraphs[0]
    tighten(s1)
    s1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_run(s1, "LOCAL MFG. · 100% ONSHORE AMERICAN INDUSTRY",
            size=7.5, bold=True, color=YELLOW, font=BODY, letter_spacing=40)
    s2 = stamp.cell(0, 2).paragraphs[0]
    tighten(s2)
    s2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    add_run(s2, "CAPABILITIES STATEMENT · V2", size=7.5, bold=True,
            color=SAND, font=BODY, letter_spacing=40)

    doc.save("Beacon-Capabilities-Statement.docx")
    print("DOCX generated")


if __name__ == "__main__":
    main()
