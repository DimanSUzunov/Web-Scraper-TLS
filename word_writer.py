"""
UPDATED
Module for extracting and writing HTML content blocks to a Word document.
Includes helpers for formatting, hyperlinks, tables, and images.
"""

from typing import Any, Optional
from config import DEFAULT_HEADING, DEFAULT_IMAGE_WIDTH, NOTICE_COLORS
from html_utils import insert_image
from fetcher import FAQ, Button

from docx.shared import Pt, Inches, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from bs4 import Tag, NavigableString
    
def add_faq_to_doc(doc, faq: FAQ):
    """Adds FAQ title with hyperlink to Word document."""

    para = doc.add_paragraph()

    # Use your existing helper (this is the correct way)
    add_hyperlink(para, faq.title, faq.url)

    # Optional visual label
    run = para.add_run(" <FAQ>")
    run.bold = True
    run.font.color.rgb = RGBColor(0, 102, 204)

    para.paragraph_format.space_after = Pt(6)


# --- Formatting Helpers ---
def restart_numbering(paragraph, num_id: int, level: int = 0) -> None:
    """Restarts list numbering for a paragraph."""
    p = paragraph._p
    pPr = p.get_or_add_pPr()
    numPr = OxmlElement('w:numPr')
    ilvl = OxmlElement('w:ilvl')
    ilvl.set(qn('w:val'), str(level))
    numPr.append(ilvl)
    numId = OxmlElement('w:numId')
    numId.set(qn('w:val'), str(num_id))
    numPr.append(numId)
    pPr.append(numPr)

def is_bold_tag(tag: Any) -> bool:
    """Checks if a tag should be rendered as bold."""
    bold_classes = {'bold', 'strong', 'fw-bold', 'font-weight-bold', 'font-weight-700'}
    if isinstance(tag, NavigableString):
        return False
    return (
        tag.name in {'strong', 'b'} or
        'font-weight: bold' in str(tag.get('style', '')).lower() or
        bool(set(tag.get('class', [])) & bold_classes))

# --- Hyperlink Helper ---
def add_hyperlink(paragraph, text: str, url: str) -> None:
    """Adds a clickable hyperlink to a paragraph."""
    part = paragraph.part
    r_id = part.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True)
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)
    run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    color = OxmlElement('w:color')
    color.set(qn('w:val'), "0000FF")
    u = OxmlElement('w:u')
    u.set(qn('w:val'), "single")
    rPr.append(color)
    rPr.append(u)
    run.append(rPr)
    t = OxmlElement('w:t')
    t.text = text
    run.append(t)
    hyperlink.append(run)
    paragraph._p.append(hyperlink)

    return paragraph


def add_button_to_doc(doc, button: Button):
    """Adds a button as hyperlink with <Button> label."""

    para = doc.add_paragraph()

    # clickable link
    add_hyperlink(para, button.text, button.url)

    # <Button> label
    run = para.add_run(" <Button>")
    run.bold = True
    run.font.color.rgb = RGBColor(0, 102, 204)

    para.paragraph_format.space_after = Pt(6)


# --- Inline and Block Formatting ---
def walk_inline(para, elem: Any, inherited_bold: bool = False, font_color: Optional[str] = None) -> None:
    """Recursively adds inline content to a paragraph, handling bold and color."""
    if isinstance(elem, NavigableString):
        if elem.strip():
            run = para.add_run(str(elem))
            if inherited_bold:
                run.bold = True
            if font_color:
                run.font.color.rgb = RGBColor.from_string(font_color)
    elif isinstance(elem, Tag):
        if elem.name == 'a' and elem.get('href'):
            add_hyperlink(para, elem.get_text(strip=True), elem['href'])
            return
        this_bold = inherited_bold or is_bold_tag(elem)
        for child in elem.children:
            walk_inline(para, child, this_bold, font_color)

def add_paragraph_with_formatting(doc, parent_elem: Any, font_color: Optional[str] = None) -> Any:
    """Adds a formatted paragraph to the document from a parent HTML element."""
    para = doc.add_paragraph()
    walk_inline(para, parent_elem, font_color=font_color)
    return para

def add_table_with_borders(doc, table_tag: Any) -> None:
    """Adds a table with borders to the document from an HTML table tag."""
    rows = table_tag.find_all("tr")
    if not rows:
        return
    num_cols = max(len(row.find_all(["td", "th"])) for row in rows)
    table = doc.add_table(rows=0, cols=num_cols)
    table.style = 'Table Grid'
    for row in rows:
        cells = row.find_all(["td", "th"])
        row_cells = table.add_row().cells
        for i, cell in enumerate(cells):
            para = row_cells[i].paragraphs[0]
            walk_inline(para, cell)
            if cell.name == "th":
                for run in para.runs:
                    run.bold = True

# --- Main Extraction Function ---
def extract_text_elements(doc, container: Any, num_id_counter: int) -> int:
    """
    Extracts and writes HTML elements from a container into the Word document.
    Handles headers, paragraphs, lists, images, tables, notices, and accordions.
    Returns the updated num_id_counter for list numbering.
    """
    # Handle buttons attached to block
    if container is not None and hasattr(container, "buttons") and container.buttons:
        for button in container.buttons:
            add_button_to_doc(doc, button)    

    # Render images attached to block
    if container is not None and hasattr(container, "images") and container.images:
        for image in container.images:
            if image.url.startswith("<"):
                # Inline SVG string
                from html_utils import insert_inline_svg
                insert_inline_svg(doc, image.url)
            else:
                # Regular image URL
                insert_image(doc, image.url)

    def process_element(elem):
        nonlocal num_id_counter
        if isinstance(elem, NavigableString):
            return
        # Skip sitemap menu blocks
        class_list = elem.get('class', [])
        if any(cls.startswith('frame-type-menu_sitemap') for cls in class_list):
            return
        if elem.name in ['h1', 'h2', 'h3', 'h4']:
            level = {'h1': 1, 'h2': 2, 'h3': 3, 'h4': 4}.get(elem.name, 2)
            para = doc.add_paragraph()
            run = para.add_run(f"{elem.get_text(strip=True)} <header {level}>")
            run.bold = True
        elif elem.name == 'p':
            add_paragraph_with_formatting(doc, elem)
        elif elem.name == 'ul':
            for li in elem.find_all("li", recursive=False):
                add_paragraph_with_formatting(doc, li).style = 'List Bullet'
        elif elem.name == 'ol':
            current_num_id = num_id_counter
            num_id_counter += 1
            for li in elem.find_all("li", recursive=False):
                para = add_paragraph_with_formatting(doc, li)
                para.style = 'List Number'
                restart_numbering(para, current_num_id, level=0)
        elif elem.name == 'img':
            src = elem.get("src")
            if src and src.startswith("http"):
                insert_image(doc, src)
        elif elem.name == 'table':
            add_table_with_borders(doc, elem)
        elif elem.name == 'br':
            doc.add_paragraph()
        
        elif elem.name == 'div' and 'notice' in elem.get('class', []):
            classes = elem.get('class', [])

            color = None
            label = "Note"  # default

            for notice_class, notice_color in NOTICE_COLORS.items():
                if notice_class in classes:
                    color = notice_color

                    # Map class to correct label
                    if "red" in notice_class:
                        label = "Bug"
                    elif "yellow" in notice_class:
                        label = "Warning"
                    elif "blue" in notice_class:
                        label = "Note"

                    break

            # Create paragraph
            para = doc.add_paragraph()

            # Add bold label with colon
            label_run = para.add_run(f"{label}: ")
            label_run.bold = True

            if color:
                label_run.font.color.rgb = RGBColor.from_string(color)

            # Add notice text
            walk_inline(para, elem, inherited_bold=False, font_color=color)

            # Optional spacing for readability
            para.paragraph_format.space_after = Pt(6)


        elif elem.name == 'dl' and 'accordion' in elem.get('class', []):
            dt_tags = elem.find_all("dt", class_="accordion__title")
            for dt in dt_tags:
                title_text = "+ " + dt.get_text(strip=True) + " <Accordion>"
                dd = dt.find_next_sibling("dd", class_="accordion__content")
                if not dd:
                    continue
                p = doc.add_paragraph()
                run = p.add_run(title_text)
                run.bold = True
                run.font.size = Pt(16)
                run.font.name = 'Calibri'
                run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Calibri')
                for child in dd.children:
                    if isinstance(child, NavigableString):
                        continue
                    num_id_counter = extract_text_elements(doc, child, num_id_counter)
        else:
            for child in elem.children:
                process_element(child)

    process_element(container)

    # Add FAQs AFTER processing the block
    if container is not None and hasattr(container, "faqs") and container.faqs:

        # optional heading
        heading = doc.add_paragraph("Frequently Asked Questions")
        heading.runs[0].bold = True
        heading.runs[0].font.size = Pt(16)

        for faq in container.faqs:
            add_faq_to_doc(doc, faq)


    return num_id_counter
