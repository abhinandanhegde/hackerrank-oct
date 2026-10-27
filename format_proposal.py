from copy import deepcopy
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

SOURCE = r"C:\Users\abhinandan\Downloads\SDG_Photography_Competition_2026_Proposal.docx"
OUTPUT = r"C:\Users\abhinandan\Desktop\hackerrank\SDG_Photography_Competition_2026_Proposal_Polished.docx"

NAVY = "17365D"
PALE_BLUE = "D9EAF7"
LIGHT_BORDER = "D9D9D9"


def set_cell_shading(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_cell_border(cell, color=LIGHT_BORDER):
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = tc_pr.first_child_found_in("w:tcBorders")
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        tag = qn(f"w:{edge}")
        node = borders.find(tag)
        if node is None:
            node = OxmlElement(f"w:{edge}")
            borders.append(node)
        node.set(qn("w:val"), "single")
        node.set(qn("w:sz"), "6")
        node.set(qn("w:color"), color)


def set_cell_margins(cell, top=110, start=120, bottom=110, end=120):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for side, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tc_mar.find(qn(f"w:{side}"))
        if node is None:
            node = OxmlElement(f"w:{side}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def set_font(run, size=10.5, bold=None, color="000000", italic=None):
    run.font.name = "Aptos"
    run._element.rPr.rFonts.set(qn("w:ascii"), "Aptos")
    run._element.rPr.rFonts.set(qn("w:hAnsi"), "Aptos")
    run.font.size = Pt(size)
    run.font.color.rgb = RGBColor.from_string(color)
    if bold is not None:
        run.font.bold = bold
    if italic is not None:
        run.font.italic = italic


def format_paragraph(p, size=10.5, before=0, after=6, line=1.12, align=None):
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line
    if align is not None:
        p.alignment = align
    for run in p.runs:
        set_font(run, size=size)


doc = Document(SOURCE)
for section in doc.sections:
    section.top_margin = Inches(0.72)
    section.bottom_margin = Inches(0.72)
    section.left_margin = Inches(0.75)
    section.right_margin = Inches(0.75)

# Base typography and spacing.
styles = doc.styles
styles["Normal"].font.name = "Aptos"
styles["Normal"]._element.rPr.rFonts.set(qn("w:ascii"), "Aptos")
styles["Normal"]._element.rPr.rFonts.set(qn("w:hAnsi"), "Aptos")
styles["Normal"].font.size = Pt(10.5)

# Cover/title block.
title = doc.paragraphs[0]
title.style = styles["Title"]
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
title.paragraph_format.space_before = Pt(6)
title.paragraph_format.space_after = Pt(4)
for run in title.runs:
    set_font(run, size=20, bold=True)

subtitle = doc.paragraphs[1]
subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
subtitle.paragraph_format.space_before = Pt(0)
subtitle.paragraph_format.space_after = Pt(18)
for run in subtitle.runs:
    set_font(run, size=13, italic=True)

heading_indexes = {3, 5, 7, 9, 11, 12}
for i, p in enumerate(doc.paragraphs):
    if not p.text.strip() or i in (0, 1):
        continue
    if i in heading_indexes:
        p.style = styles["Heading 1"]
        p.paragraph_format.space_before = Pt(15)
        p.paragraph_format.space_after = Pt(7)
        p.paragraph_format.keep_with_next = True
        for run in p.runs:
            set_font(run, size=13, bold=True)
    elif i in (13, 14, 15):
        format_paragraph(p, size=10.5, before=0, after=5, line=1.12)
    else:
        format_paragraph(p, size=10.5, before=0, after=7, line=1.12)

# Tables: readable padding, explicit borders, deliberate widths/alignment.
for ti, table in enumerate(doc.tables):
    table.style = "Table Grid"
    table.autofit = False
    ncols = len(table.columns)
    for ri, row in enumerate(table.rows):
        for ci, cell in enumerate(row.cells):
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            set_cell_margins(cell)
            set_cell_border(cell)
            for p in cell.paragraphs:
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(3)
                p.paragraph_format.line_spacing = 1.06
                if ti == 4 and ci > 0:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                else:
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                for run in p.runs:
                    set_font(run, size=9.5)
            # Key/value tables use a clear key column.
            if ncols == 2 and ci == 0:
                set_cell_shading(cell, PALE_BLUE)
                for p in cell.paragraphs:
                    for run in p.runs:
                        set_font(run, size=9.5, bold=True)

    # Scoring matrix has a distinct header and total row.
    if ti == 4:
        for cell in table.rows[0].cells:
            set_cell_shading(cell, NAVY)
            for p in cell.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in p.runs:
                    set_font(run, size=9.5, bold=True, color="FFFFFF")
        for cell in table.rows[-1].cells:
            set_cell_shading(cell, PALE_BLUE)
            for p in cell.paragraphs:
                for run in p.runs:
                    set_font(run, size=9.5, bold=True)
        widths = [Inches(3.7), Inches(0.92), Inches(0.92), Inches(0.92)]
    elif ncols == 2:
        widths = [Inches(2.0), Inches(4.45)]
    else:
        widths = [Inches(6.45 / ncols)] * ncols
    for row in table.rows:
        for ci, width in enumerate(widths):
            row.cells[ci].width = width

    # Keep the heading above its following table but avoid forced blank pages.
    if ti < len(doc.tables) - 1:
        for cell in table.rows[-1].cells:
            for p in cell.paragraphs:
                p.paragraph_format.space_after = Pt(5)

doc.core_properties.title = "SDG Photography Competition 2026 Proposal"
doc.save(OUTPUT)
print(OUTPUT)
