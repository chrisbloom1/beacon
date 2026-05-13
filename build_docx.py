"""Build Beacon-Capabilities-Statement.docx with python-docx.

Uses Beacon's real brand: yellow #fdb602, cream #fdf6e4, deep teal #022429,
sand #c9b789. Display: Barlow Condensed (fallback Impact). Body: Figtree
(fallback Helvetica/Arial). Mirrors the HTML one-pager.
"""
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# Brand
YELLOW = "FDB602"
CREAM = "FDF6E4"
TEAL = "022429"
SAND = "C9B789"
WHITE = "FFFFFF"

DISPLAY_FONT = "Barlow Condensed"
DISPLAY_FALLBACK = "Impact"
BODY_FONT = "Figtree"
BODY_FALLBACK = "Helvetica"


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
    tbl = table._element
    tblPr = tbl.tblPr
    tblBorders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        b = OxmlElement(f"w:{edge}")
        b.set(qn("w:val"), "nil")
        tblBorders.append(b)
    tblPr.append(tblBorders)


def add_run(p, text, *, size=10, bold=False, italic=False, color=TEAL,
            font=BODY_FONT, letter_spacing=None, caps=False, shade=None):
    r = p.add_run(text)
    r.font.name = font
    # Set complex script + east asian names too so Word respects font
    rPr = r._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), font)
    rFonts.set(qn("w:hAnsi"), font)
    rFonts.set(qn("w:cs"), font)
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.italic = italic
    r.font.color.rgb = RGBColor.from_string(color)
    if letter_spacing is not None:
        spacing = OxmlElement("w:spacing")
        spacing.set(qn("w:val"), str(letter_spacing))
        rPr.append(spacing)
    if caps:
        caps_el = OxmlElement("w:caps")
        caps_el.set(qn("w:val"), "1")
        rPr.append(caps_el)
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
    settings = doc.settings.element
    display_bg = OxmlElement("w:displayBackgroundShape")
    settings.append(display_bg)


def main():
    doc = Document()
    section = doc.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    section.top_margin = Inches(0.45)
    section.bottom_margin = Inches(0.4)
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)
    set_page_bg(doc, CREAM)

    # Base style → Figtree
    base = doc.styles["Normal"]
    base.font.name = BODY_FONT
    base.font.size = Pt(10)
    base.font.color.rgb = RGBColor.from_string(TEAL)

    # ============= LETTERHEAD =============
    head = doc.add_table(rows=1, cols=2)
    head.autofit = False
    head.columns[0].width = Inches(4.7)
    head.columns[1].width = Inches(2.8)
    remove_table_borders(head)

    brand_cell = head.cell(0, 0)
    brand_cell.width = Inches(4.7)
    bp = brand_cell.paragraphs[0]
    tighten(bp)

    # Nested table: mark + name
    bi = brand_cell.add_table(rows=1, cols=2)
    bi.autofit = False
    bi.columns[0].width = Inches(0.78)
    bi.columns[1].width = Inches(3.85)
    remove_table_borders(bi)

    mark = bi.cell(0, 0)
    mark.width = Inches(0.78)
    mark.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    # Insert the beacon-mark.png
    mp = mark.paragraphs[0]
    tighten(mp)
    mp.alignment = WD_ALIGN_PARAGRAPH.LEFT
    try:
        mp.add_run().add_picture("assets/beacon-mark.png", width=Inches(0.65))
    except Exception:
        # Fallback: yellow B
        set_cell_bg(mark, YELLOW)
        add_run(mp, "B", size=30, bold=True, color=TEAL, font=DISPLAY_FONT)
        mp.alignment = WD_ALIGN_PARAGRAPH.CENTER

    name_cell = bi.cell(0, 1)
    name_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    e = name_cell.paragraphs[0]
    tighten(e)
    add_run(e, "CAPABILITIES STATEMENT · V2", size=8, bold=True, color=TEAL,
            font=BODY_FONT, letter_spacing=50)
    n = name_cell.add_paragraph()
    tighten(n, before=1)
    add_run(n, "BEACON ", size=26, bold=True, color=TEAL, font=DISPLAY_FONT)
    add_run(n, "MANUFACTURING", size=26, bold=False, color=TEAL, font=DISPLAY_FONT)

    # Address
    addr = head.cell(0, 1)
    addr.width = Inches(2.8)
    ap = addr.paragraphs[0]
    ap.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    tighten(ap)
    add_run(ap, "DETROIT · MICHIGAN", size=11, bold=True, color=TEAL,
            font=DISPLAY_FONT, letter_spacing=40)
    for line in ["2050 15th St", "Detroit, MI 48216",
                 "Local Mfg. · 100% Onshore"]:
        p = addr.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        tighten(p)
        add_run(p, line, size=9, color=TEAL)
    # Yellow URL chip
    url_p = addr.add_paragraph()
    url_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    tighten(url_p, before=3)
    add_run(url_p, "  beaconmfg.us  ", size=10, bold=True, color=TEAL,
            font=BODY_FONT, shade=YELLOW)

    # Letterhead bottom border
    for c in head.rows[0].cells:
        set_cell_border(c,
            top={"val": "nil"}, left={"val": "nil"}, right={"val": "nil"},
            bottom={"val": "single", "sz": 24, "color": TEAL})

    sp = doc.add_paragraph()
    tighten(sp, after=6)

    # ============= LEDE =============
    h1 = doc.add_paragraph()
    tighten(h1, after=2)
    add_run(h1, "BUILDER-LED ", size=26, bold=True, color=TEAL, font=DISPLAY_FONT)
    add_run(h1, "  CONTRACT MANUFACTURING  ", size=26, bold=True,
            color=TEAL, font=DISPLAY_FONT, shade=YELLOW)
    h1b = doc.add_paragraph()
    tighten(h1b, after=4)
    add_run(h1b, "MADE IN DETROIT.", size=26, bold=True, color=TEAL, font=DISPLAY_FONT)

    lede = doc.add_paragraph()
    tighten(lede, before=2, after=4, line=1.4)
    add_run(lede,
        "Beacon is a builder-led contract manufacturer in Detroit ",
        size=10.5, bold=True, color=TEAL)
    add_run(lede,
        "delivering metal fabrication, assembly, and rapid prototyping — from "
        "one-off prototypes to repeatable short runs and full production. ",
        size=10.5, color=TEAL)
    add_run(lede, "Bring us a problem. ", size=10.5, bold=True, color=TEAL)
    add_run(lede, "We'll quote it, prototype it, and put it on a line.",
            size=10.5, color=TEAL)

    # ============= CORE CAPABILITIES TITLE =============
    title_tbl = doc.add_table(rows=1, cols=1)
    title_tbl.autofit = False
    remove_table_borders(title_tbl)
    tc = title_tbl.cell(0, 0)
    tcp = tc.paragraphs[0]
    tighten(tcp, before=6, after=2)
    add_run(tcp, "▮ ", size=14, bold=True, color=YELLOW, font=BODY_FONT)
    add_run(tcp, "CORE CAPABILITIES   ", size=13, bold=True, color=TEAL,
            font=DISPLAY_FONT, letter_spacing=70)
    add_run(tcp, "Fabrication · Assembly · Prototyping", size=9,
            italic=True, color=TEAL)

    # ============= CARDS =============
    cards = doc.add_table(rows=1, cols=3)
    cards.autofit = False
    col_w = Inches(2.46)
    for col in cards.columns:
        col.width = col_w
    remove_table_borders(cards)

    card_data = [
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
          (None, "Same-week iteration on form, fit, function"),
          (None, "Design-for-manufacture review & support")]),
    ]

    for i, (num, title, blurb, items) in enumerate(card_data):
        cell = cards.cell(0, i)
        cell.width = col_w
        set_cell_bg(cell, WHITE)
        set_cell_border(cell,
            top={"val": "single", "sz": 4, "color": "E8DFC0"},
            right={"val": "single", "sz": 4, "color": "E8DFC0"},
            bottom={"val": "single", "sz": 4, "color": "E8DFC0"},
            left={"val": "single", "sz": 48, "color": YELLOW})
        set_cell_margins(cell, top=140, bottom=140, left=180, right=180)

        p0 = cell.paragraphs[0]
        tighten(p0)
        add_run(p0, num, size=10, bold=True, color=TEAL,
                font=DISPLAY_FONT, letter_spacing=40)

        ph = cell.add_paragraph()
        tighten(ph, before=1, after=2)
        add_run(ph, title, size=17, bold=True, color=TEAL, font=DISPLAY_FONT)

        pt = cell.add_paragraph()
        tighten(pt, after=4, line=1.35)
        add_run(pt, blurb, size=9, italic=True, color=TEAL)

        for bold_part, rest in items:
            ip = cell.add_paragraph()
            tighten(ip, line=1.4)
            ip.paragraph_format.left_indent = Inches(0.08)
            add_run(ip, "■ ", size=8, bold=True, color=YELLOW, font=BODY_FONT)
            if bold_part:
                add_run(ip, bold_part, size=9.5, bold=True, color=TEAL)
                add_run(ip, rest, size=9.5, color=TEAL)
            else:
                add_run(ip, rest, size=9.5, color=TEAL)

    # ============= WAREHOUSING TITLE =============
    sp2 = doc.add_paragraph()
    tighten(sp2, after=2)

    title_tbl2 = doc.add_table(rows=1, cols=1)
    title_tbl2.autofit = False
    remove_table_borders(title_tbl2)
    t2 = title_tbl2.cell(0, 0)
    t2p = t2.paragraphs[0]
    tighten(t2p, before=6, after=2)
    add_run(t2p, "▮ ", size=14, bold=True, color=YELLOW, font=BODY_FONT)
    add_run(t2p, "WAREHOUSING & LOGISTICS   ", size=13, bold=True,
            color=TEAL, font=DISPLAY_FONT, letter_spacing=70)
    add_run(t2p, "Products in, products out — without the chaos.",
            size=9, italic=True, color=TEAL)

    # ============= STRIP (dark) =============
    strip = doc.add_table(rows=1, cols=2)
    strip.autofit = False
    strip.columns[0].width = Inches(1.85)
    strip.columns[1].width = Inches(5.65)
    remove_table_borders(strip)

    sl = strip.cell(0, 0)
    sr = strip.cell(0, 1)
    sl.width = Inches(1.85)
    sr.width = Inches(5.65)
    set_cell_bg(sl, TEAL)
    set_cell_bg(sr, TEAL)
    set_cell_margins(sl, top=140, bottom=140, left=160, right=140)
    set_cell_margins(sr, top=140, bottom=140, left=140, right=160)
    set_cell_border(sl,
        left={"val": "single", "sz": 48, "color": YELLOW},
        top={"val": "nil"}, bottom={"val": "nil"}, right={"val": "nil"})
    set_cell_border(sr,
        left={"val": "nil"}, top={"val": "nil"},
        bottom={"val": "nil"}, right={"val": "nil"})

    sl_p = sl.paragraphs[0]
    tighten(sl_p)
    add_run(sl_p, "ONE ROOF.", size=16, bold=True, color=CREAM, font=DISPLAY_FONT)
    sl_p2 = sl.add_paragraph()
    tighten(sl_p2, after=2)
    add_run(sl_p2, "END TO END.", size=16, bold=True, color=CREAM, font=DISPLAY_FONT)
    sl_p3 = sl.add_paragraph()
    tighten(sl_p3)
    add_run(sl_p3, "Receiving · inventory · outbound.", size=8,
            italic=True, color=SAND)

    sr_p = sr.paragraphs[0]
    tighten(sr_p, line=1.4)
    add_run(sr_p,
        "Receiving, inventory, and part management with flexible storage on "
        "the floor. Parcel and freight outbound — including no-box, roll-on / "
        "roll-off for awkward builds. We keep your supply tight so the line "
        "never starves.", size=9.5, color=CREAM)

    pills_p = sr.add_paragraph()
    tighten(pills_p, before=5)
    filled = ["RECEIVING", "INVENTORY", "PART MGMT", "FLEXIBLE STORAGE"]
    outline = ["PARCEL", "FREIGHT", "ROLL-ON / ROLL-OFF"]
    for i, label in enumerate(filled):
        add_run(pills_p, f"  {label}  ", size=8, bold=True, color=TEAL,
                font=BODY_FONT, letter_spacing=20, shade=YELLOW)
        add_run(pills_p, "  ", size=8)
    for i, label in enumerate(outline):
        add_run(pills_p, f"[ {label} ]", size=8, bold=True, color=YELLOW,
                font=BODY_FONT, letter_spacing=20)
        if i != len(outline) - 1:
            add_run(pills_p, "  ", size=8)

    # ============= META ROW =============
    sp3 = doc.add_paragraph()
    tighten(sp3, after=2)
    meta = doc.add_table(rows=1, cols=3)
    meta.autofit = False
    for col in meta.columns:
        col.width = Inches(2.46)
    remove_table_borders(meta)
    meta_data = [
        ("MADE FOR",
         "Hardware founders · Industrial OEMs · Mobility · Architectural · Defense & gov · Consumer goods"),
        ("HOW WE WORK",
         "Send a print, a CAD file, or a napkin sketch. We quote, prototype, and put it on a line."),
        ("ETHOS",
         "Local Mfg. · 100% Onshore American Industry · Builder-led · Detroit"),
    ]
    for i, (label, value) in enumerate(meta_data):
        c = meta.cell(0, i)
        c.width = Inches(2.46)
        set_cell_border(c,
            top={"val": "single", "sz": 18, "color": TEAL},
            left={"val": "nil"}, right={"val": "nil"}, bottom={"val": "nil"})
        set_cell_margins(c, top=100, bottom=40, left=40, right=80)
        p0 = c.paragraphs[0]
        tighten(p0)
        add_run(p0, label, size=10, bold=True, color=TEAL,
                font=DISPLAY_FONT, letter_spacing=50)
        p1 = c.add_paragraph()
        tighten(p1, line=1.35)
        add_run(p1, value, size=9, color=TEAL)

    # spacer push CTA toward bottom
    for _ in range(2):
        pad = doc.add_paragraph()
        tighten(pad, after=6)

    # ============= CTA =============
    cta = doc.add_table(rows=1, cols=2)
    cta.autofit = False
    cta.columns[0].width = Inches(4.6)
    cta.columns[1].width = Inches(2.9)
    remove_table_borders(cta)
    cl = cta.cell(0, 0)
    cr = cta.cell(0, 1)
    cl.width = Inches(4.6)
    cr.width = Inches(2.9)
    set_cell_bg(cl, YELLOW)
    set_cell_bg(cr, YELLOW)
    set_cell_margins(cl, top=180, bottom=180, left=220, right=140)
    set_cell_margins(cr, top=180, bottom=180, left=140, right=200)
    set_cell_border(cl,
        left={"val": "single", "sz": 64, "color": TEAL},
        top={"val": "nil"}, bottom={"val": "nil"}, right={"val": "nil"})
    set_cell_border(cr,
        left={"val": "nil"}, top={"val": "nil"},
        bottom={"val": "nil"}, right={"val": "nil"})

    clp = cl.paragraphs[0]
    tighten(clp)
    add_run(clp, "HAVE A PART, KIT, OR ASSEMBLY IN MIND?",
            size=10, bold=True, color=TEAL, font=DISPLAY_FONT, letter_spacing=50)
    clp2 = cl.add_paragraph()
    tighten(clp2, before=2)
    add_run(clp2, "SEND A PRINT. WE'LL PUT IT ON A LINE.",
            size=20, bold=True, color=TEAL, font=DISPLAY_FONT)

    crp = cr.paragraphs[0]
    crp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    tighten(crp, before=4)
    add_run(crp, "BEACONMFG.US", size=14, bold=True, color=TEAL,
            font=DISPLAY_FONT, letter_spacing=50)
    crp2 = cr.add_paragraph()
    crp2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    tighten(crp2, before=1)
    add_run(crp2, "2050 15th St · Detroit, MI 48216",
            size=9, color=TEAL)

    # Stamp footer line
    stamp = doc.add_paragraph()
    tighten(stamp, before=4)
    add_run(stamp, "BEACON © 2026     ", size=7.5, bold=True, color=TEAL,
            font=BODY_FONT, letter_spacing=30)
    add_run(stamp, "·     LOCAL MFG. · 100% ONSHORE AMERICAN INDUSTRY     ",
            size=7.5, bold=True, color=TEAL, font=BODY_FONT, letter_spacing=30)
    add_run(stamp, "·     CAPABILITIES STATEMENT · V2",
            size=7.5, bold=True, color=TEAL, font=BODY_FONT, letter_spacing=30)

    doc.save("Beacon-Capabilities-Statement.docx")
    print("DOCX generated")


if __name__ == "__main__":
    main()
