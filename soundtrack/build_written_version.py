from pathlib import Path
import re

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor


ROOT = Path(__file__).resolve().parent
SOURCE = ROOT / "written_version.md"
OUTPUT = ROOT / "love_over_time_written_version.docx"


def set_cell_margins(cell, top=90, start=120, bottom=90, end=120):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = tcPr.first_child_found_in("w:tcMar")
    if tcMar is None:
        tcMar = OxmlElement("w:tcMar")
        tcPr.append(tcMar)
    for m, v in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tcMar.find(qn(f"w:{m}"))
        if node is None:
            node = OxmlElement(f"w:{m}")
            tcMar.append(node)
        node.set(qn("w:w"), str(v))
        node.set(qn("w:type"), "dxa")


def add_page_number(paragraph):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run("Love Over Time  ·  ")
    run.font.size = Pt(8)
    run.font.color.rgb = RGBColor(102, 113, 125)
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), "PAGE")
    paragraph._p.append(fld)


def add_markdown_runs(paragraph, text):
    tokens = re.split(r"(\*[^*]+\*)", text)
    for token in tokens:
        if not token:
            continue
        if token.startswith("*") and token.endswith("*"):
            run = paragraph.add_run(token[1:-1])
            run.italic = True
        else:
            paragraph.add_run(token)


def set_run_font(run, name="Aptos", size=11, color="17202A"):
    run.font.name = name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    run.font.size = Pt(size)
    run.font.color.rgb = RGBColor.from_string(color)


def style_document(doc):
    styles = doc.styles
    normal = styles["Normal"]
    normal.font.name = "Aptos"
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Aptos")
    normal.font.size = Pt(11)
    normal.font.color.rgb = RGBColor(23, 32, 42)
    normal.paragraph_format.space_after = Pt(8)
    normal.paragraph_format.line_spacing = 1.15

    title = styles["Title"]
    title.font.name = "Georgia"
    title._element.rPr.rFonts.set(qn("w:eastAsia"), "Georgia")
    title.font.size = Pt(26)
    title.font.bold = False
    title.font.color.rgb = RGBColor(0, 0, 0)
    title.paragraph_format.space_after = Pt(6)

    subtitle = styles["Subtitle"]
    subtitle.font.name = "Georgia"
    subtitle._element.rPr.rFonts.set(qn("w:eastAsia"), "Georgia")
    subtitle.font.size = Pt(14)
    subtitle.font.italic = False
    subtitle.font.color.rgb = RGBColor(80, 87, 94)
    subtitle.paragraph_format.space_after = Pt(22)

    for name, size, before, after in (("Heading 1", 15, 20, 7), ("Heading 2", 12, 14, 5), ("Heading 3", 10, 12, 5)):
        style = styles[name]
        style.font.name = "Aptos Display"
        style._element.rPr.rFonts.set(qn("w:eastAsia"), "Aptos Display")
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.color.rgb = RGBColor(0, 0, 0)
        style.paragraph_format.space_before = Pt(before)
        style.paragraph_format.space_after = Pt(after)
        style.paragraph_format.keep_with_next = True


def add_paragraph(doc, text="", style=None, align=None):
    p = doc.add_paragraph(style=style)
    if align is not None:
        p.alignment = align
    add_markdown_runs(p, text)
    for run in p.runs:
        set_run_font(run)
    return p


def main():
    lines = SOURCE.read_text(encoding="utf-8").splitlines()
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Inches(0.85)
    section.bottom_margin = Inches(0.8)
    section.left_margin = Inches(0.95)
    section.right_margin = Inches(0.95)
    style_document(doc)

    footer = section.footer.paragraphs[0]
    add_page_number(footer)

    in_references = False
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        if not line:
            i += 1
            continue
        if line.startswith("# "):
            p = doc.add_paragraph(style="Title")
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(line[2:].strip())
            set_run_font(run, "Georgia", 26, "000000")
        elif line.startswith("## "):
            p = doc.add_paragraph(style="Heading 1")
            run = p.add_run(line[3:].strip())
            set_run_font(run, "Aptos Display", 15, "000000")
            in_references = line[3:].strip() == "Course Materials Referenced"
        elif line.startswith("### "):
            p = doc.add_paragraph(style="Heading 2")
            run = p.add_run(line[4:].strip())
            set_run_font(run, "Aptos Display", 12, "000000")
        elif re.match(r"^\d+\. ", line):
            p = doc.add_paragraph(style="List Number")
            add_markdown_runs(p, re.sub(r"^\d+\. ", "", line))
            for run in p.runs:
                set_run_font(run)
        else:
            p = doc.add_paragraph()
            add_markdown_runs(p, line)
            for run in p.runs:
                set_run_font(run, "Aptos", 10.5 if in_references else 11)
            if in_references:
                p.paragraph_format.space_after = Pt(3)
                p.paragraph_format.line_spacing = 1.0
        i += 1

    # Make the short title/artist headings visually distinct while keeping them as ordinary headings.
    for p in doc.paragraphs:
        if p.style.name == "Heading 1" and "—" in p.text:
            for run in p.runs:
                run.font.color.rgb = RGBColor(198, 110, 78)

    doc.core_properties.title = "Love Over Time"
    doc.core_properties.subject = "A soundtrack of romantic relationships across the life course"
    doc.core_properties.author = "Olivia Huang"
    doc.core_properties.keywords = "HDFS 225, close relationships, soundtrack, life course"
    doc.save(OUTPUT)
    print(OUTPUT)


if __name__ == "__main__":
    main()
