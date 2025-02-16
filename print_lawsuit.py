#!/usr/bin/env python3
import argparse
import re
import pickle
import os
from collections import OrderedDict

# PDF-related imports
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib.utils import ImageReader

# DOCX-related imports
import docx
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Inches

###############################################################################
#  LAWSUIT CLASS
###############################################################################
class Lawsuit:
    """
    A container for Lawsuit data, with each attribute stored as an OrderedDict:
      - header:   an OrderedDict for storing top-level metadata.
      - sections: an OrderedDict with each key as "HEADING_NUMBER HEADING_TEXT"
                  (e.g. "1. INTRODUCTION" or "I. INTRODUCTION"),
                  and each value as the full text of that heading.
      - exhibits: an OrderedDict keyed by a simple index string ("1", "2", etc.).
                  Each value is another OrderedDict with 'caption' and 'image_path'.
      - documents: an OrderedDict for storing entire detected bracketed documents.

    This class also automatically stores case_information and law_firm_information
    from command line arguments.
    """

    def __init__(
        self,
        sections=None,
        exhibits=None,
        header=None,
        documents=None,
        case_information="",
        law_firm_information=""
    ):
        if sections is None:
            sections = OrderedDict()
        if exhibits is None:
            exhibits = OrderedDict()
        if header is None:
            header = OrderedDict()
        if documents is None:
            documents = OrderedDict()

        # Ensure they are all OrderedDict:
        self.sections = OrderedDict(sections)
        self.exhibits = OrderedDict(exhibits)
        self.header   = OrderedDict(header)
        self.documents = OrderedDict(documents)

        # Store the command-line-provided info
        self.case_information     = case_information
        self.law_firm_information = law_firm_information

    def __repr__(self):
        """
        Print the Lawsuit object clearly, showing all keys and values in each OrderedDict fully,
        as well as the case_information and law_firm_information fields.
        """
        header_str = "\n".join([f"  {k}: {v}" for k, v in self.header.items()])
        sections_str = "\n".join([f"  {sec_key}: {sec_value}" for sec_key, sec_value in self.sections.items()])
        exhibits_str = []
        for ex_key, ex_data in self.exhibits.items():
            ex_inner = "\n      ".join([f"{ik}: {iv}" for ik, iv in ex_data.items()])
            exhibits_str.append(f"  {ex_key}:\n      {ex_inner}")
        exhibits_str = "\n".join(exhibits_str)

        documents_str = []
        for doc_key, doc_text in self.documents.items():
            documents_str.append(f"  {doc_key}:\n      {doc_text}")
        documents_str = "\n".join(documents_str)

        return (
            "Lawsuit Object:\n\n"
            "CASE INFORMATION:\n"
            f"  {self.case_information}\n\n"
            "LAW FIRM INFORMATION:\n"
            f"  {self.law_firm_information}\n\n"
            "HEADER:\n"
            f"{header_str}\n\n"
            "SECTIONS:\n"
            f"{sections_str}\n\n"
            "EXHIBITS:\n"
            f"{exhibits_str}\n\n"
            "DOCUMENTS:\n"
            f"{documents_str}\n"
        )

###############################################################################
#  HELPER FUNCTIONS (PDF)
###############################################################################
def is_line_all_caps(line_str):
    """Returns True if the line contains at least one uppercase letter
       and no lowercase letters (a-z)."""
    import string
    if not re.search(r'[A-Z]', line_str):
        return False
    return not re.search(r'[a-z]', line_str)

def draw_firm_name_vertical_center(pdf_canvas, text, page_width, page_height):
    """
    Draws the firm name vertically, centered along the page height on the left side.
    """
    pdf_canvas.saveState()
    pdf_canvas.setFont("Times-Bold", 10)
    text_width = pdf_canvas.stringWidth(text, "Times-Bold", 10)
    x_pos = 0.2 * inch
    y_center = page_height / 2.0
    y_pos = y_center - (text_width / 2.0)
    pdf_canvas.translate(x_pos, y_pos)
    pdf_canvas.rotate(90)
    pdf_canvas.drawString(0, 0, text)
    pdf_canvas.restoreState()

def wrap_text_to_lines(pdf_canvas, full_text, font_name, font_size, max_width):
    """
    Splits a large text into a list of (line_string, ended_full_line) pairs,
    respecting max_width so that text does not overflow horizontally.
    """
    pdf_canvas.setFont(font_name, font_size)
    paragraphs = full_text.split('\n')
    all_lines = []
    for paragraph in paragraphs:
        words = paragraph.split()
        if not words:
            # Empty line
            all_lines.append(("", False))
            continue
        current_line = ""
        for word in words:
            if not current_line:
                test_line = word
            else:
                test_line = current_line + " " + word

            if pdf_canvas.stringWidth(test_line, font_name, font_size) <= max_width:
                current_line = test_line
            else:
                # The line overflowed
                all_lines.append((current_line, True))
                current_line = word
        if current_line:
            all_lines.append((current_line, False))
    return all_lines

def draw_exhibit_page(
    pdf_canvas,
    page_width,
    page_height,
    firm_name,
    case_name,
    exhibit_caption,
    exhibit_image,
    page_number,
    total_pages
):
    """
    Draws a single exhibit page with bounding box, firm/case name, exhibit caption at top,
    and the exhibit image scaled to fill the remaining space.
    """
    # Bounding box
    pdf_canvas.setLineWidth(2)
    pdf_canvas.rect(0.5 * inch, 0.5 * inch, page_width - 1.0 * inch, page_height - 1.3 * inch)

    # Vertical firm name
    draw_firm_name_vertical_center(pdf_canvas, firm_name, page_width, page_height)

    # Case name at top center
    pdf_canvas.setFont("Times-Bold", 12)
    pdf_canvas.drawCentredString(page_width / 2.0, page_height - 0.5 * inch, case_name)

    # Horizontal line
    pdf_canvas.setLineWidth(1)
    pdf_canvas.line(0.5 * inch, page_height - 0.6 * inch, page_width - 0.5 * inch, page_height - 0.6 * inch)

    # Draw the exhibit caption
    pdf_canvas.setFont("Times-Roman", 10)
    top_margin = page_height - 1.2 * inch
    left_margin = 1.2 * inch
    line_spacing = 0.25 * inch
    max_caption_width = page_width - (2 * left_margin)

    # Wrap the caption
    wrapped_caption_lines = wrap_text_to_lines(pdf_canvas, exhibit_caption, "Times-Roman", 10, max_caption_width)
    current_y = top_margin
    for (cap_line, _) in wrapped_caption_lines:
        pdf_canvas.drawString(left_margin, current_y, cap_line)
        current_y -= line_spacing

    # We only want 1 blank line below the last caption line before the image:
    spacing_lines = 1
    top_of_image_area = current_y - (spacing_lines * line_spacing)
    bottom_of_image_area = 0.5 * inch  # bounding box bottom
    if top_of_image_area < bottom_of_image_area:
        top_of_image_area = bottom_of_image_area

    available_height = top_of_image_area - bottom_of_image_area
    available_width = (page_width - 1.0 * inch)  # bounding box from 0.5 inch to page_width-0.5 inch

    try:
        img_reader = ImageReader(exhibit_image)
        img_width, img_height = img_reader.getSize()
    except Exception as e:
        # If image loading fails, notify in the middle of the page
        pdf_canvas.setFont("Times-Italic", 10)
        pdf_canvas.drawCentredString(
            page_width / 2.0,
            page_height / 2.0,
            f"Unable to load image: {exhibit_image} Error: {e}"
        )
    else:
        # Scale the image
        scale = min(available_width / img_width, available_height / img_height, 1.0)
        new_width = img_width * scale
        new_height = img_height * scale

        # Center horizontally
        x_img = 0.5 * inch + (available_width - new_width) / 2.0
        y_img_bottom = bottom_of_image_area

        pdf_canvas.drawImage(
            img_reader,
            x_img,
            y_img_bottom,
            width=new_width,
            height=new_height,
            preserveAspectRatio=True,
            anchor='c'
        )

    # Footer with page number
    pdf_canvas.setFont("Times-Italic", 9)
    footer_text = f"Page {page_number} of {total_pages}"
    pdf_canvas.drawCentredString(page_width / 2.0, 0.5 * inch - 0.1 * inch, footer_text)

###############################################################################
#  DETECTING LEGAL-TITLE BLOCKS
###############################################################################
def is_full_equals_line(line_str):
    """
    Returns True if the line is composed entirely of '=' (with optional whitespace)
    and has at least a few '=' characters.
    """
    stripped = line_str.strip()
    if len(stripped) < 5:
        return False
    return bool(re.match(r'^[=]+$', stripped))

def detect_legal_title_blocks(lines):
    """
    Given a list of lines, detects blocks bracketed by lines of '========...'.
    We remove those bracket lines and treat everything in between them
    as a single "legal_page_title_block" which we will place on its own PDF page.

    Yields:
      ("legal_page_title_block", [lines_in_between])  for bracketed blocks
      ("normal_line", line)                           for normal lines
    """
    i = 0
    n = len(lines)
    while i < n:
        if is_full_equals_line(lines[i]):
            # found top bracket
            j = i + 1
            inner_lines = []
            found_bottom = False
            while j < n:
                if is_full_equals_line(lines[j]):
                    # found the closing bracket
                    found_bottom = True
                    j += 1  # skip bottom bracket
                    break
                else:
                    inner_lines.append(lines[j])
                j += 1

            if found_bottom:
                # yield the block
                yield ("legal_page_title_block", inner_lines)
                i = j
                continue
            else:
                # no matching bottom bracket; treat as normal line
                yield ("normal_line", lines[i])
                i += 1
        else:
            yield ("normal_line", lines[i])
            i += 1

###############################################################################
#  PAGE-DRAWING FOR TEXT SEGMENTS
###############################################################################
def draw_legal_page_title_block(
    pdf_canvas,
    page_width,
    page_height,
    block_lines,
    firm_name,
    case_name,
    page_number,
    total_pages,
):
    """
    Draws one bracketed block as a stand-alone page (ensuring "every detected page starts on a new page").
    The text is displayed big, bold, centered on the page, with normal bounding box, etc.
    """
    # Page bounding box
    pdf_canvas.setLineWidth(2)
    pdf_canvas.rect(0.5 * inch, 0.5 * inch, page_width - 1.0 * inch, page_height - 1.3 * inch)

    # Vertical firm name
    draw_firm_name_vertical_center(pdf_canvas, firm_name, page_width, page_height)

    # Case name at top center
    pdf_canvas.setFont("Times-Bold", 12)
    pdf_canvas.drawCentredString(page_width / 2.0, page_height - 0.5 * inch, case_name)

    # Horizontal line
    pdf_canvas.setLineWidth(1)
    pdf_canvas.line(0.5 * inch, page_height - 0.6 * inch, page_width - 0.5 * inch, page_height - 0.6 * inch)

    # Now draw block lines in big bold style, centered
    pdf_canvas.setFont("Times-Bold", 14)

    line_spacing = 0.3 * inch
    y_text = page_height - 1.5 * inch
    for line_str in block_lines:
        pdf_canvas.drawCentredString(page_width / 2.0, y_text, line_str)
        y_text -= line_spacing

    # Footer
    pdf_canvas.setFont("Times-Italic", 9)
    footer_text = f"Page {page_number} of {total_pages}"
    pdf_canvas.drawCentredString(page_width / 2.0, 0.5 * inch - 0.1 * inch, footer_text)

def draw_page_of_segments(
    pdf_canvas,
    page_width,
    page_height,
    segments,
    start_index,
    max_lines_per_page,
    firm_name,
    case_name,
    page_number,
    total_pages,
    line_offset_x,
    line_offset_y,
    line_spacing,
    heading_positions
):
    """
    Draws up to max_lines_per_page items from the 'segments' onto the PDF page,
    each with line numbers on the far left/right (unless it's a forced new-page block).

    Returns the index of the next segment that hasn't been drawn yet.
    """
    # Page bounding box
    pdf_canvas.setLineWidth(2)
    pdf_canvas.rect(0.5 * inch, 0.5 * inch, page_width - 1.0 * inch, page_height - 1.3 * inch)

    # Vertical firm name
    draw_firm_name_vertical_center(pdf_canvas, firm_name, page_width, page_height)

    # Case name at top center
    pdf_canvas.setFont("Times-Bold", 12)
    pdf_canvas.drawCentredString(page_width / 2.0, page_height - 0.5 * inch, case_name)

    # Horizontal line under case name
    pdf_canvas.setLineWidth(1)
    pdf_canvas.line(0.5 * inch, page_height - 0.6 * inch, page_width - 0.5 * inch, page_height - 0.6 * inch)

    end_index = start_index
    current_line_count = 0
    y_text = line_offset_y

    while end_index < len(segments) and current_line_count < max_lines_per_page:
        seg = segments[end_index]

        # If this segment is a "legal_page_title_block" forcing new page:
        if seg.get("page_always_new"):
            # If we haven't printed anything on this page yet, we can draw it immediately here,
            # otherwise we return so that the main loop will start a fresh page for it.
            if current_line_count > 0:
                # We must finish this page now so that on the next call we start fresh.
                break
            else:
                # Draw the single block here on a new page
                block_lines = seg["lines"]
                draw_legal_page_title_block(
                    pdf_canvas,
                    page_width,
                    page_height,
                    block_lines,
                    firm_name,
                    case_name,
                    page_number,
                    total_pages,
                )
                end_index += 1
                # We used this entire page for the bracket block, so we are done with it.
                return end_index

        # Otherwise, normal line-based segment
        line_number = end_index + 1
        # line numbers on left and right
        pdf_canvas.setFont("Times-Roman", 10)
        pdf_canvas.drawString(line_offset_x - 0.6 * inch, y_text, str(line_number))
        pdf_canvas.drawString(page_width - 0.4 * inch, y_text, str(line_number))

        # If heading => record for table of contents
        if seg["is_heading"] or seg["is_subheading"]:
            heading_positions.append(
                (
                    seg["text"],
                    page_number,
                    line_number,
                    seg["is_subheading"]
                )
            )

        # Draw text according to alignment
        pdf_canvas.setFont(seg["font_name"], seg["font_size"])
        if seg["alignment"] == "center":
            left_boundary = line_offset_x
            right_boundary = page_width - 0.5 * inch
            mid_x = (left_boundary + right_boundary) / 2.0
            pdf_canvas.drawCentredString(mid_x, y_text, seg["text"])
        else:
            pdf_canvas.drawString(line_offset_x, y_text, seg["text"])

        y_text -= line_spacing
        current_line_count += 1
        end_index += 1

    # Footer
    pdf_canvas.setFont("Times-Italic", 9)
    footer_text = f"Page {page_number} of {total_pages}"
    pdf_canvas.drawCentredString(page_width / 2.0, 0.5 * inch - 0.1 * inch, footer_text)

    return end_index

###############################################################################
#  TABLE OF CONTENTS (PDF)
###############################################################################
def generate_index_pdf(index_filename, firm_name, case_name, heading_positions):
    """
    Generates a table of contents PDF (index_filename) that lists each section/subsection
    in order encountered in the main PDF (heading_positions). Next to each entry,
    prints the page#:line#.
    """
    pdf_canvas = canvas.Canvas(index_filename, pagesize=letter)
    pdf_canvas.setTitle("Table of Contents")

    page_width, page_height = letter
    top_margin = 1.0 * inch
    bottom_margin = 1.0 * inch
    left_margin = 1.0 * inch
    right_margin = 0.5 * inch
    line_spacing = 0.25 * inch

    # We'll measure text widths with a temporary canvas
    temp_canvas = canvas.Canvas("dummy.pdf", pagesize=letter)

    def wrap_text(linestr, font_name, font_size, maxwidth):
        temp_canvas.setFont(font_name, font_size)
        return wrap_text_to_lines(temp_canvas, linestr, font_name, font_size, maxwidth)

    max_entry_width = page_width - left_margin - 1.5 * inch

    # Flatten out headings, wrapping as needed
    flattened_lines = []
    for (heading_text, pg_num, ln_num, is_sub) in heading_positions:
        if is_sub:
            font_name = "Times-Roman"
            font_size = 9
        else:
            font_name = "Times-Bold"
            font_size = 10

        wrapped = wrap_text(heading_text, font_name, font_size, max_entry_width)
        text_lines = [w[0] for w in wrapped] if wrapped else [""]

        for i, txt_line in enumerate(text_lines):
            flattened_lines.append(
                (
                    txt_line,
                    pg_num,
                    ln_num,
                    is_sub,
                    (i == 0)  # is_first_line
                )
            )

    usable_height = page_height - (top_margin + bottom_margin) - 1.0 * inch
    max_lines_per_page = int(usable_height // line_spacing)

    total_lines = len(flattened_lines)
    total_index_pages = max(1, (total_lines + max_lines_per_page - 1) // max_lines_per_page)

    i = 0
    current_page_index = 1

    while i < total_lines:
        # Page bounding box
        pdf_canvas.setLineWidth(2)
        pdf_canvas.rect(0.5 * inch, 0.5 * inch, page_width - 1.0 * inch, page_height - 1.3 * inch)

        # Vertical firm name
        draw_firm_name_vertical_center(pdf_canvas, firm_name, page_width, page_height)

        # Case name at top center
        pdf_canvas.setFont("Times-Bold", 12)
        pdf_canvas.drawCentredString(page_width / 2.0, page_height - 0.5 * inch, case_name)

        # Horizontal line
        pdf_canvas.setLineWidth(1)
        pdf_canvas.line(0.5 * inch, page_height - 0.6 * inch, page_width - 0.5 * inch, page_height - 0.6 * inch)

        # Title: TABLE OF CONTENTS
        pdf_canvas.setFont("Times-Bold", 14)
        pdf_canvas.drawCentredString(page_width / 2.0, page_height - 0.75 * inch, "TABLE OF CONTENTS")

        x_text = left_margin
        y_text = page_height - top_margin - 0.75 * inch
        lines_on_this_page = 0

        while i < total_lines and lines_on_this_page < max_lines_per_page:
            (line_text, pg_num, ln_num, is_sub, is_first_line) = flattened_lines[i]
            if is_sub:
                font_name = "Times-Roman"
                font_size = 9
            else:
                font_name = "Times-Bold"
                font_size = 10

            pdf_canvas.setFont(font_name, font_size)
            pdf_canvas.drawString(x_text, y_text, line_text)

            # Print "page:line" on the right only on the first wrapped line
            if is_first_line:
                label_str = f"{pg_num}:{ln_num}"
                pdf_canvas.drawRightString(page_width - right_margin - 0.2 * inch, y_text, label_str)

            y_text -= line_spacing
            i += 1
            lines_on_this_page += 1

        # Footer
        pdf_canvas.setFont("Times-Italic", 9)
        footer_text = f"Index Page {current_page_index} of {total_index_pages}"
        pdf_canvas.drawCentredString(page_width / 2.0, 0.5 * inch - 0.1 * inch, footer_text)

        if i < total_lines:
            pdf_canvas.showPage()
            current_page_index += 1
        else:
            break

    pdf_canvas.save()

###############################################################################
#  DOCX GENERATION (COMPLAINT + TOC)
###############################################################################
def generate_complaint_docx(docx_filename, firm_name, case_name, header_od, sections_od, heading_styles):
    """
    Generates a Word document version of the complaint text.
    - The 'header_od["content"]' is placed at the beginning.
    - Any bracketed blocks (legal page title blocks) become big bold paragraphs, centered.
    - Headings are bold or subheading style; normal body lines are standard paragraphs.
    """
    doc = Document()

    # Base font style
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(12)

    # Insert the firm/case name at top center
    top_par = doc.add_paragraph()
    top_par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = top_par.add_run(f"{firm_name} | {case_name}\n")
    run.bold = True
    run.font.size = Pt(14)

    # Step 1: handle the header content
    header_content = header_od.get("content", "")
    header_lines = header_content.splitlines()

    buffer_of_lines = []
    for kind, block_lines in detect_legal_title_blocks(header_lines):
        if kind == "legal_page_title_block":
            # flush normal lines first
            if buffer_of_lines:
                for line in buffer_of_lines:
                    line_stripped = line.strip()
                    para = doc.add_paragraph()
                    if is_line_all_caps(line_stripped):
                        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    else:
                        para.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    para.add_run(line_stripped)
                buffer_of_lines = []

            # Now add the block lines as big bold
            for line in block_lines:
                para = doc.add_paragraph()
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                runx = para.add_run(line.strip())
                runx.bold = True
                runx.font.size = Pt(14)

        else:
            buffer_of_lines.append(block_lines)

    # Flush any leftover normal lines in header
    if buffer_of_lines:
        for line in buffer_of_lines:
            line_stripped = line.strip()
            para = doc.add_paragraph()
            if is_line_all_caps(line_stripped):
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            para.add_run(line_stripped)
        buffer_of_lines = []

    # Step 2: sections
    for section_key, section_body in sections_od.items():
        style_type = heading_styles.get(section_key, "section")

        # blank line
        doc.add_paragraph()

        # heading paragraph
        heading_para = doc.add_paragraph()
        heading_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if style_type == "section":
            run = heading_para.add_run(section_key)
            run.bold = True
            run.font.size = Pt(12)
        else:
            run = heading_para.add_run(section_key)
            run.bold = False
            run.font.size = Pt(11)

        # handle body lines, checking bracket blocks
        body_lines = section_body.splitlines()
        normal_buffer = []
        for kind, block_lines in detect_legal_title_blocks(body_lines):
            if kind == "legal_page_title_block":
                # flush normal lines first
                if normal_buffer:
                    for bline in normal_buffer:
                        bline_str = bline.strip()
                        para = doc.add_paragraph()
                        para.alignment = WD_ALIGN_PARAGRAPH.LEFT
                        rr = para.add_run(bline_str)
                        if style_type == "section":
                            rr.font.size = Pt(12)
                        else:
                            rr.font.size = Pt(11)
                    normal_buffer = []
                # now add the bracket block lines in big bold
                for linex in block_lines:
                    para = doc.add_paragraph()
                    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    runx = para.add_run(linex.strip())
                    runx.bold = True
                    runx.font.size = Pt(14)
            else:
                normal_buffer.append(block_lines)

        # flush leftover normal lines
        if normal_buffer:
            for bline in normal_buffer:
                bline_str = bline.strip()
                para = doc.add_paragraph()
                para.alignment = WD_ALIGN_PARAGRAPH.LEFT
                rr = para.add_run(bline_str)
                if style_type == "section":
                    rr.font.size = Pt(12)
                else:
                    rr.font.size = Pt(11)
            normal_buffer = []

    doc.save(docx_filename)
    print(f"DOCX complaint saved as: {docx_filename}")

def generate_toc_docx(docx_filename, firm_name, case_name, heading_positions):
    """
    Generates a Table of Contents in DOCX form (page#:line#), with each main heading and subheading.
    """
    doc = Document()

    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(12)

    # Top heading
    top_par = doc.add_paragraph()
    top_par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = top_par.add_run(f"{firm_name} | {case_name}\nTABLE OF CONTENTS\n")
    run.bold = True
    run.font.size = Pt(14)

    # Create a table for the TOC
    table = doc.add_table(rows=0, cols=2)
    table.autofit = True

    for (heading_text, pg_num, ln_num, is_sub) in heading_positions:
        row_cells = table.add_row().cells
        left_cell = row_cells[0]
        right_cell = row_cells[1]

        if is_sub:
            this_font_size = 11
            this_bold = False
        else:
            this_font_size = 12
            this_bold = True

        # Left cell: heading text
        left_par = left_cell.paragraphs[0]
        run_left = left_par.add_run(heading_text)
        run_left.font.size = Pt(this_font_size)
        run_left.bold = this_bold

        # Right cell: "page:line"
        right_par = right_cell.paragraphs[0]
        run_right = right_par.add_run(f"{pg_num}:{ln_num}")
        run_right.font.size = Pt(this_font_size)
        run_right.bold = False
        right_par.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    doc.save(docx_filename)
    print(f"Table of Contents DOCX saved as: {docx_filename}")

###############################################################################
#  ADDITIONAL PARSER FOR MULTIPLE DOCUMENTS
###############################################################################
def parse_documents_from_text(raw_text):
    """
    Parse multiple bracketed documents (separated by lines of '=====...') from the raw text.
    Each matched pair of full-equals lines forms one document. Those bracket lines are removed,
    and the content in between is considered a single document.
    Returns a list of document-strings (without the bracket lines).
    """
    lines = raw_text.splitlines()
    docs = []
    i = 0
    n = len(lines)

    while i < n:
        if is_full_equals_line(lines[i]):
            j = i + 1
            doc_lines = []
            while j < n and not is_full_equals_line(lines[j]):
                doc_lines.append(lines[j])
                j += 1
            if j < n:
                # found a bottom bracket
                docs.append("\n".join(doc_lines))
                i = j + 1
            else:
                # no matching bottom bracket
                break
        else:
            i += 1

    return docs

###############################################################################
#  PARSING HEADER AND SECTIONS
###############################################################################
def parse_header_and_sections(raw_text):
    """
    Anything before the first valid heading is stored in header['content'].

    A valid heading must:
        1) Match a pattern allowing either Roman numerals or digits followed by a dot,
           repeated one or more times, e.g. "I. ", "1. ", "II.1. ", etc.
        2) The text portion after that must be all-caps to qualify as a heading.
    """
    header_od = OrderedDict()
    sections_od = OrderedDict()

    heading_pattern = re.compile(r'^((?:[IVXLCDM]+\.|[0-9]+\.)+)\s+(.*)$', re.IGNORECASE)
    lines = raw_text.splitlines()
    idx = 0
    header_lines = []

    # find first heading
    while idx < len(lines):
        line = lines[idx].rstrip('\n').rstrip('\r')
        m = heading_pattern.match(line)
        if m:
            heading_number = m.group(1).strip()
            heading_title = m.group(2).strip()
            if is_line_all_caps(heading_title):
                break
        header_lines.append(line)
        idx += 1

    header_od["content"] = "\n".join(header_lines)

    current_heading_key = None
    current_body_lines = []

    while idx < len(lines):
        line = lines[idx].rstrip('\n').rstrip('\r')
        match_heading = heading_pattern.match(line)
        if match_heading:
            heading_number = match_heading.group(1).strip()
            heading_title = match_heading.group(2).strip()
            if is_line_all_caps(heading_title):
                # finalize old heading
                if current_heading_key is not None:
                    sections_od[current_heading_key] = "\n".join(current_body_lines)
                current_body_lines = []
                if heading_number.endswith('.'):
                    heading_number = heading_number[:-1]
                current_heading_key = f"{heading_number} {heading_title}"
            else:
                current_body_lines.append(line)
        else:
            current_body_lines.append(line)
        idx += 1

    if current_heading_key is not None:
        sections_od[current_heading_key] = "\n".join(current_body_lines)

    return header_od, sections_od

def classify_headings(sections_od):
    """
    Distinguish main sections vs. subsections:
      - If the numeric portion has more than one dot (e.g. "1.1", "II.1", etc.), it's a subheading.
      - Otherwise, it's a main section.
    """
    heading_styles = {}
    for full_key in sections_od.keys():
        parts = full_key.split(None, 1)
        if len(parts) == 2:
            heading_number, _heading_text = parts[0], parts[1]
        else:
            heading_number = parts[0]
        dot_count = heading_number.count('.')
        if dot_count > 1:
            heading_styles[full_key] = "subsection"
        else:
            heading_styles[full_key] = "section"
    return heading_styles

###############################################################################
#  BUILDING SEGMENTS
###############################################################################
def prepare_main_pdf_segments(header_text, sections_od, heading_styles, pdf_canvas, max_text_width):
    """
    Create a list of segments. Each segment is a dict describing how to render that line or block:
      {
        "text": <string line>,
        "font_name": "Times-Roman" or "Times-Bold",
        "font_size": 10 or 9,
        "alignment": "left" or "center",
        "is_heading": bool,
        "is_subheading": bool
      }
      ...OR...
      {
        "legal_page_title": True,
        "page_always_new": True,
        "lines": [strings in bracketed block]
      }

    We detect bracketed blocks (legal_page_title_block) and ensure each one
    will start on its own page by using "page_always_new": True.

    We also wrap normal lines to fit max_text_width.
    """
    segments = []

    # 1) handle header lines (and bracketed blocks in them)
    header_lines = header_text.splitlines()
    normal_buffer = []

    def flush_normal_buffer():
        for line in normal_buffer:
            line_str = line.strip()
            if not line_str:
                # blank line
                segments.append({
                    "text": "",
                    "font_name": "Times-Roman",
                    "font_size": 10,
                    "alignment": "left",
                    "is_heading": False,
                    "is_subheading": False
                })
            elif is_line_all_caps(line_str):
                # center it
                wrapped = wrap_text_to_lines(pdf_canvas, line_str, "Times-Roman", 10, max_text_width)
                for (wl, _) in wrapped:
                    segments.append({
                        "text": wl,
                        "font_name": "Times-Roman",
                        "font_size": 10,
                        "alignment": "center",
                        "is_heading": False,
                        "is_subheading": False
                    })
            else:
                # left
                wrapped = wrap_text_to_lines(pdf_canvas, line_str, "Times-Roman", 10, max_text_width)
                for (wl, _) in wrapped:
                    segments.append({
                        "text": wl,
                        "font_name": "Times-Roman",
                        "font_size": 10,
                        "alignment": "left",
                        "is_heading": False,
                        "is_subheading": False
                    })
        normal_buffer.clear()

    for kind, block_lines in detect_legal_title_blocks(header_lines):
        if kind == "legal_page_title_block":
            if normal_buffer:
                flush_normal_buffer()
            lines_cleaned = [ln.strip() for ln in block_lines]
            segments.append({
                "legal_page_title": True,
                "page_always_new": True,  # ensure it starts on a new page
                "lines": lines_cleaned
            })
        else:
            normal_buffer.append(block_lines)

    if normal_buffer:
        flush_normal_buffer()

    # 2) handle each section
    for section_key, section_body in sections_od.items():
        style = heading_styles.get(section_key, "section")
        if style == "section":
            heading_font_name = "Times-Bold"
            heading_font_size = 10
            body_font_name = "Times-Roman"
            body_font_size = 10
            is_heading = True
            is_subheading = False
        else:
            heading_font_name = "Times-Roman"
            heading_font_size = 9
            body_font_name = "Times-Roman"
            body_font_size = 9
            is_heading = False
            is_subheading = True

        # Add a blank line
        segments.append({
            "text": "",
            "font_name": body_font_name,
            "font_size": body_font_size,
            "alignment": "left",
            "is_heading": False,
            "is_subheading": False
        })

        # Heading line(s) (wrapped if needed)
        heading_wrapped = wrap_text_to_lines(
            pdf_canvas, section_key, heading_font_name, heading_font_size, max_text_width
        )
        for (wl, _) in heading_wrapped:
            segments.append({
                "text": wl,
                "font_name": heading_font_name,
                "font_size": heading_font_size,
                "alignment": "center",
                "is_heading": is_heading,
                "is_subheading": is_subheading
            })

        # Then body lines + possible bracket blocks
        lines_of_body = section_body.splitlines()
        normal_buffer_sec = []

        def flush_section_buffer():
            for line in normal_buffer_sec:
                line_str = line.strip()
                if not line_str:
                    segments.append({
                        "text": "",
                        "font_name": body_font_name,
                        "font_size": body_font_size,
                        "alignment": "left",
                        "is_heading": False,
                        "is_subheading": False
                    })
                else:
                    wrapped = wrap_text_to_lines(
                        pdf_canvas, line_str, body_font_name, body_font_size, max_text_width
                    )
                    for (wl, _) in wrapped:
                        segments.append({
                            "text": wl,
                            "font_name": body_font_name,
                            "font_size": body_font_size,
                            "alignment": "left",
                            "is_heading": False,
                            "is_subheading": False
                        })
            normal_buffer_sec.clear()

        for kind, block_lines in detect_legal_title_blocks(lines_of_body):
            if kind == "legal_page_title_block":
                if normal_buffer_sec:
                    flush_section_buffer()
                lines_cleaned = [ln.strip() for ln in block_lines]
                segments.append({
                    "legal_page_title": True,
                    "page_always_new": True,  # force a new page for bracket block
                    "lines": lines_cleaned
                })
            else:
                normal_buffer_sec.append(block_lines)

        if normal_buffer_sec:
            flush_section_buffer()

    return segments

###############################################################################
#  MAIN PDF GENERATION
###############################################################################
def generate_legal_document(
    firm_name,
    case_name,
    output_filename,
    header_od,
    sections_od,
    exhibits,
    heading_positions
):
    """
    Generate the main PDF with line-numbered text (including bracket-block pages).
    Then append exhibits. Also produce a DOCX version of the same content.
    """
    page_width, page_height = letter
    pdf_canvas = canvas.Canvas(output_filename, pagesize=letter)
    pdf_canvas.setTitle("Legal Document")
    pdf_canvas.setAuthor(firm_name)
    pdf_canvas.setSubject(case_name)
    pdf_canvas.setCreator("Legal PDF Generator")

    heading_styles = classify_headings(sections_od)

    top_margin = 1.0 * inch
    bottom_margin = 1.0 * inch
    left_margin = 1.2 * inch
    right_margin = 0.5 * inch
    line_spacing = 0.25 * inch

    usable_height = page_height - (top_margin + bottom_margin)
    max_lines_per_page = int(usable_height // line_spacing)
    line_offset_x = left_margin
    line_offset_y = page_height - top_margin
    max_text_width = page_width - right_margin - line_offset_x - 0.2 * inch

    # Build segments for main content
    segments = prepare_main_pdf_segments(
        header_text=header_od.get("content", ""),
        sections_od=sections_od,
        heading_styles=heading_styles,
        pdf_canvas=pdf_canvas,
        max_text_width=max_text_width
    )

    # Count how many pages the text segments will require
    current_index = 0
    text_pages = 0
    total_segments = len(segments)

    while current_index < total_segments:
        seg = segments[current_index]
        if seg.get("page_always_new"):
            # This block alone will consume one full page
            text_pages += 1
            current_index += 1
        else:
            # We can fit up to max_lines_per_page segments on this page
            lines_used = 0
            local_i = current_index
            while local_i < total_segments and lines_used < max_lines_per_page:
                s = segments[local_i]
                if s.get("page_always_new"):
                    # must stop to start new page for that block
                    break
                lines_used += 1
                local_i += 1
            text_pages += 1
            current_index = local_i

    # The total number of exhibit pages is the number of exhibits
    exhibit_pages = len(exhibits)
    total_pages = text_pages + exhibit_pages

    # Actually render the text segments
    page_number = 1
    current_index = 0
    while current_index < total_segments:
        # Start a new page
        next_index = draw_page_of_segments(
            pdf_canvas=pdf_canvas,
            page_width=page_width,
            page_height=page_height,
            segments=segments,
            start_index=current_index,
            max_lines_per_page=max_lines_per_page,
            firm_name=firm_name,
            case_name=case_name,
            page_number=page_number,
            total_pages=total_pages,
            line_offset_x=line_offset_x,
            line_offset_y=line_offset_y,
            line_spacing=line_spacing,
            heading_positions=heading_positions
        )
        pdf_canvas.showPage()
        page_number += 1
        current_index = next_index

    # Render each exhibit on its own page
    for (caption, image_path) in exhibits:
        draw_exhibit_page(
            pdf_canvas=pdf_canvas,
            page_width=page_width,
            page_height=page_height,
            firm_name=firm_name,
            case_name=case_name,
            exhibit_caption=caption,
            exhibit_image=image_path,
            page_number=page_number,
            total_pages=total_pages,
        )
        pdf_canvas.showPage()
        page_number += 1

    pdf_canvas.save()

    # Also generate DOCX
    generate_complaint_docx(
        docx_filename=os.path.splitext(output_filename)[0] + ".docx",
        firm_name=firm_name,
        case_name=case_name,
        header_od=header_od,
        sections_od=sections_od,
        heading_styles=heading_styles
    )

###############################################################################
#  MAIN
###############################################################################
def main():
    parser = argparse.ArgumentParser(
        description=(
            "Generate a legal-style PDF (with bracketed blocks each on a new page), "
            "line numbering, exhibits, a separate table-of-contents PDF, and DOCX files. "
            "Also pickles the Lawsuit object if requested."
        )
    )
    parser.add_argument("--firm_name", required=True,
                        help="Firm name placed vertically on pages.")
    parser.add_argument("--case", required=True,
                        help="Case name placed horizontally at the top.")
    parser.add_argument("--output", default="lawsuit.pdf",
                        help="Output PDF filename for the main document (default: lawsuit.pdf).")
    parser.add_argument("--file", required=True,
                        help="Path to a UTF-8 text file containing the body (with sections/subsections).")
    parser.add_argument("--exhibits", nargs='+', default=[],
                        help="Exhibit caption-text-file and image-file pairs, e.g. --exhibits cap1.txt img1.png cap2.txt img2.png.")
    parser.add_argument("--index", default="index.pdf",
                        help="PDF filename for the table of contents (default: index.pdf).")
    parser.add_argument("--pickle", nargs='?', const=None,
                        help=(
                            "Optional path to store the Lawsuit object in pickle format. "
                            "If no path is given, defaults to 'lawsuit.pickle'."
                        ))

    args = parser.parse_args()

    # Read the raw text from the file
    with open(args.file, 'r', encoding='utf-8') as f:
        raw_text = f.read()

    # Parse out header and sections
    header_od, sections_od = parse_header_and_sections(raw_text)

    # Build exhibits
    if len(args.exhibits) % 2 != 0:
        raise ValueError("Exhibits must be in pairs: CAPTION_FILE IMAGE_FILE")

    exhibits_od = OrderedDict()
    ex_index = 1
    for i in range(0, len(args.exhibits), 2):
        cap_file = args.exhibits[i]
        image_file = args.exhibits[i + 1]
        with open(cap_file, 'r', encoding='utf-8') as cfp:
            cap_text = cfp.read()
        exhibits_od[str(ex_index)] = OrderedDict([
            ('caption', cap_text),
            ('image_path', image_file)
        ])
        ex_index += 1

    # Add sample metadata to the header (you can adjust or remove this as needed)
    header_od["DocumentTitle"] = "Complaint for Damages"
    header_od["DateFiled"] = "2025-02-14"
    header_od["Court"] = "Sample Court"

    # Parse bracketed documents from raw_text (if any). Each bracket-block pair is considered a separate document.
    found_documents = parse_documents_from_text(raw_text)
    documents_od = OrderedDict()
    if found_documents:
        for idx, doc_text in enumerate(found_documents, start=1):
            documents_od[str(idx)] = doc_text

    # Create Lawsuit object, including the newly parsed documents
    lawsuit_obj = Lawsuit(
        sections=sections_od,
        exhibits=exhibits_od,
        header=header_od,
        documents=documents_od,
        case_information=args.case,
        law_firm_information=args.firm_name
    )

    # Convert exhibits for PDF generation
    exhibits_for_pdf = []
    for key, val in lawsuit_obj.exhibits.items():
        exhibits_for_pdf.append((val["caption"], val["image_path"]))

    # Track headings for TOC
    heading_positions = []

    # Generate main PDF + DOCX
    generate_legal_document(
        firm_name=args.firm_name,
        case_name=args.case,
        output_filename=args.output,
        header_od=header_od,
        sections_od=sections_od,
        exhibits=exhibits_for_pdf,
        heading_positions=heading_positions
    )

    # Generate TOC PDF + DOCX
    generate_index_pdf(
        index_filename=args.index,
        firm_name=args.firm_name,
        case_name=args.case,
        heading_positions=heading_positions
    )
    index_docx = os.path.splitext(args.index)[0] + ".docx"
    generate_toc_docx(
        docx_filename=index_docx,
        firm_name=args.firm_name,
        case_name=args.case,
        heading_positions=heading_positions
    )

    # Optionally pickle
    if args.pickle is not None:
        pickle_filename = args.pickle if args.pickle else "lawsuit.pickle"
        with open(pickle_filename, "wb") as pf:
            pickle.dump(lawsuit_obj, pf)
        pkl_path = pickle_filename
    else:
        pkl_path = "Not saved (not requested)."

    # Summary
    print(f"PDF generated: {args.output}")
    print(f"DOCX Complaint generated: {os.path.splitext(args.output)[0] + '.docx'}")
    print(f"Index PDF generated: {args.index}")
    print(f"Index DOCX generated: {index_docx}")
    print(f"Lawsuit object saved to: {pkl_path}\n")
    print("Dumped Lawsuit object:")
    print(lawsuit_obj)

if __name__ == "__main__":
    main()