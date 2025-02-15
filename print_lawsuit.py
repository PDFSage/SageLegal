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
from docx.shared import Pt

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

    This class also automatically stores case_information and law_firm_information
    from command line arguments.
    """

    def __init__(self, sections=None, exhibits=None, header=None,
                 case_information="", law_firm_information=""):
        if sections is None:
            sections = OrderedDict()
        if exhibits is None:
            exhibits = OrderedDict()
        if header is None:
            header = OrderedDict()

        # Ensure they are all OrderedDict:
        self.sections = OrderedDict(sections)
        self.exhibits = OrderedDict(exhibits)
        self.header   = OrderedDict(header)

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
            f"{exhibits_str}\n"
        )

###############################################################################
#  HELPER FUNCTIONS (PDF)
###############################################################################
def is_line_all_caps(line_str):
    """Returns True if the line contains at least one uppercase letter
       and no lowercase letters (a-z)."""
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
    respecting max_width so that text does not overflow.
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

def draw_exhibit_page(pdf_canvas,
                      page_width,
                      page_height,
                      firm_name,
                      case_name,
                      exhibit_caption,
                      exhibit_image,
                      page_number,
                      total_pages):
    """
    Draws a single exhibit page with bounding box, firm/case name, exhibit caption at top,
    and the exhibit image centered below.
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

    # Prepare to place the exhibit image
    spacing_lines = 2
    margin = 1.0 * inch
    max_img_width = page_width - 2 * margin
    max_img_height = page_height - 2 * margin

    # Try loading the image
    try:
        img_reader = ImageReader(exhibit_image)
        img_width, img_height = img_reader.getSize()
    except Exception as e:
        pdf_canvas.setFont("Times-Italic", 10)
        pdf_canvas.drawCentredString(
            page_width / 2.0,
            page_height / 2.0,
            f"Unable to load image: {exhibit_image} Error: {e}"
        )
    else:
        scale = min(max_img_width / img_width, max_img_height / img_height, 1.0)
        new_width = img_width * scale
        new_height = img_height * scale

        y_img_top = current_y - (spacing_lines * line_spacing)
        y_img_bottom = y_img_top - new_height
        bottom_margin = 1.0 * inch

        # If the image goes too low, adjust upward
        if y_img_bottom < bottom_margin:
            y_img_bottom = bottom_margin
            y_img_top = y_img_bottom + new_height

        x_img = (page_width - new_width) / 2.0
        pdf_canvas.drawImage(img_reader,
                             x_img,
                             y_img_bottom,
                             width=new_width,
                             height=new_height,
                             preserveAspectRatio=True,
                             anchor='c')

    # Footer with page number
    pdf_canvas.setFont("Times-Italic", 9)
    footer_text = f"Page {page_number} of {total_pages}"
    pdf_canvas.drawCentredString(page_width / 2.0, 0.5 * inch - 0.1 * inch, footer_text)

###############################################################################
#  BUILDING THE MAIN PDF CONTENT
###############################################################################
def prepare_main_pdf_segments(header_text, sections_od, heading_styles):
    """
    Prepares a list of segments for the main PDF from the given header text
    and sections. Each segment is a dict:
        {
            "text": <string>,
            "font_name": "Times-Roman" or "Times-Bold",
            "font_size": 10 or 9,
            "alignment": "left" or "center",
            "is_heading": True/False,
            "is_subheading": True/False
        }

    Requirements:
      - Within the 'header_text' lines, center-align any lines that are all caps,
        otherwise left-align, using normal 10 pt Times-Roman.
      - Main section headings => Bold, 10 pt, center-aligned, preceded by a blank line.
      - Subsection headings => Not bold, 9 pt, center-aligned, preceded by a blank line.
      - Body text for main sections => Normal 10 pt.
      - Body text for subsections => Normal 9 pt.
    """
    segments = []

    # 1) Handle the header text lines
    header_lines = header_text.splitlines()
    for line in header_lines:
        line_stripped = line.strip()
        if not line_stripped:
            segments.append({
                "text": "",
                "font_name": "Times-Roman",
                "font_size": 10,
                "alignment": "left",
                "is_heading": False,
                "is_subheading": False
            })
            continue

        if is_line_all_caps(line_stripped):
            segments.append({
                "text": line_stripped,
                "font_name": "Times-Roman",
                "font_size": 10,
                "alignment": "center",
                "is_heading": False,
                "is_subheading": False
            })
        else:
            segments.append({
                "text": line_stripped,
                "font_name": "Times-Roman",
                "font_size": 10,
                "alignment": "left",
                "is_heading": False,
                "is_subheading": False
            })

    # 2) Handle each section
    for section_key, section_body in sections_od.items():
        style = heading_styles.get(section_key, "section")
        if style == "section":
            # main section
            heading_font_name = "Times-Bold"
            heading_font_size = 10
            body_font_name = "Times-Roman"
            body_font_size = 10
            is_heading = True
            is_subheading = False
        else:
            # subsection
            heading_font_name = "Times-Roman"  # not bold
            heading_font_size = 9             # smaller
            body_font_name = "Times-Roman"
            body_font_size = 9
            is_heading = False
            is_subheading = True

        # Insert a blank line before the heading
        segments.append({
            "text": "",
            "font_name": heading_font_name,
            "font_size": heading_font_size,
            "alignment": "left",
            "is_heading": False,
            "is_subheading": False
        })

        # Add the heading segment (center aligned)
        segments.append({
            "text": section_key,
            "font_name": heading_font_name,
            "font_size": heading_font_size,
            "alignment": "center",
            "is_heading": is_heading,
            "is_subheading": is_subheading
        })

        # Add the body segments
        body_lines = section_body.splitlines()
        for body_line in body_lines:
            segments.append({
                "text": body_line,
                "font_name": body_font_name,
                "font_size": body_font_size,
                "alignment": "left",
                "is_heading": False,
                "is_subheading": False
            })

    return segments

def wrap_segments(pdf_canvas, segments, max_width):
    """
    Given a list of segment dicts (with text, font_name, font_size, alignment, etc.),
    expand/wrap each segment's text to produce a list of final lines.

    Returns a list of dicts:
        {
            "text": <string>,
            "font_name": ...,
            "font_size": ...,
            "alignment": "left" or "center",
            "is_heading": ...,
            "is_subheading": ...,
        }

    Each line is guaranteed to fit in max_width for the specified font.
    """
    wrapped_output = []
    for seg in segments:
        text = seg["text"]
        font_name = seg["font_name"]
        font_size = seg["font_size"]
        alignment = seg["alignment"]
        is_heading = seg["is_heading"]
        is_subheading = seg["is_subheading"]

        lines = wrap_text_to_lines(pdf_canvas, text, font_name, font_size, max_width)
        if not lines:
            # Possibly an empty line
            wrapped_output.append({
                "text": "",
                "font_name": font_name,
                "font_size": font_size,
                "alignment": alignment,
                "is_heading": is_heading,
                "is_subheading": is_subheading
            })
            continue

        for (wrapped_line, _) in lines:
            wrapped_output.append({
                "text": wrapped_line,
                "font_name": font_name,
                "font_size": font_size,
                "alignment": alignment,
                "is_heading": is_heading,
                "is_subheading": is_subheading
            })

    return wrapped_output

def draw_page_of_segments(pdf_canvas,
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
                          heading_positions):
    """
    Draws up to max_lines_per_page items from the 'segments' list onto the PDF page,
    each with a line number on the far left and far right.

    heading_positions is a list; whenever we encounter a heading or subheading segment,
    we record (text, page_number, line_number, is_subheading) so that we can build
    the Table of Contents.
    """
    # Bounding box
    pdf_canvas.setLineWidth(2)
    pdf_canvas.rect(0.5 * inch, 0.5 * inch, page_width - 1.0 * inch, page_height - 1.3 * inch)

    # Vertical firm name
    draw_firm_name_vertical_center(pdf_canvas, firm_name, page_width, page_height)

    # Case name at top center
    pdf_canvas.setFont("Times-Bold", 12)
    pdf_canvas.drawCentredString(page_width / 2.0, page_height - 0.5 * inch, case_name)

    # Horizontal line below case name
    pdf_canvas.setLineWidth(1)
    pdf_canvas.line(0.5 * inch, page_height - 0.6 * inch, page_width - 0.5 * inch, page_height - 0.6 * inch)

    end_index = min(start_index + max_lines_per_page, len(segments))
    y_text = line_offset_y

    for i in range(start_index, end_index):
        line_number = i + 1  # 1-based line numbering
        seg = segments[i]

        # Left line number
        pdf_canvas.setFont("Times-Roman", 10)
        pdf_canvas.drawString(line_offset_x - 0.6 * inch, y_text, str(line_number))
        # Right line number
        pdf_canvas.drawString(page_width - 0.4 * inch, y_text, str(line_number))

        # Record heading info if this segment is a heading or subheading
        if seg["is_heading"] or seg["is_subheading"]:
            heading_positions.append((
                seg["text"],       # heading text
                page_number,       # current page
                line_number,       # line number
                seg["is_subheading"]
            ))

        # Draw the actual text
        pdf_canvas.setFont(seg["font_name"], seg["font_size"])
        if seg["alignment"] == "center":
            left_boundary  = line_offset_x
            right_boundary = page_width - 0.5 * inch
            mid_x = (left_boundary + right_boundary) / 2.0
            pdf_canvas.drawCentredString(mid_x, y_text, seg["text"])
        else:
            pdf_canvas.drawString(line_offset_x, y_text, seg["text"])

        y_text -= line_spacing

    # Footer with page number
    pdf_canvas.setFont("Times-Italic", 9)
    footer_text = f"Page {page_number} of {total_pages}"
    pdf_canvas.drawCentredString(page_width / 2.0, 0.5 * inch - 0.1 * inch, footer_text)

    return end_index

###############################################################################
#  GENERATE TABLE OF CONTENTS (PDF)
###############################################################################
def generate_index_pdf(index_filename, firm_name, case_name, heading_positions):
    """
    Generates a table of contents PDF (index_filename) that lists each section or subsection
    in the order encountered in the main PDF (heading_positions). Next to each entry,
    prints the page number and line number. Subsections are smaller font than main sections.
    """
    pdf_canvas = canvas.Canvas(index_filename, pagesize=letter)
    pdf_canvas.setTitle("Table of Contents")

    page_width, page_height = letter
    top_margin = 1.0 * inch
    bottom_margin = 1.0 * inch
    left_margin = 1.0 * inch
    right_margin = 0.5 * inch
    line_spacing = 0.25 * inch

    # We'll measure text widths with a temporary canvas:
    temp_canvas = canvas.Canvas("dummy.pdf", pagesize=letter)

    # Flatten out headings to account for wrapping
    max_entry_width = page_width - left_margin - 1.5 * inch

    flattened_lines = []
    for (heading_text, page_num, ln_num, is_sub) in heading_positions:
        if is_sub:
            font_name = "Times-Roman"
            font_size = 9
        else:
            font_name = "Times-Bold"
            font_size = 10

        wrapped = wrap_text_to_lines(temp_canvas, heading_text, font_name, font_size, max_entry_width)
        text_lines = [w[0] for w in wrapped] if wrapped else [""]

        for i, txt_line in enumerate(text_lines):
            flattened_lines.append((
                txt_line,
                page_num,
                ln_num,
                is_sub,
                (i == 0)  # is this the first line of the heading text?
            ))

    usable_height = page_height - (top_margin + bottom_margin) - 1.0 * inch
    max_lines_per_page = int(usable_height // line_spacing)

    total_lines = len(flattened_lines)
    total_index_pages = max(1, (total_lines + max_lines_per_page - 1) // max_lines_per_page)

    i = 0
    current_page_index = 1

    while i < total_lines:
        # Draw bounding box
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

            if is_first_line:
                # Print "page:line" on the right
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
#  GENERATE DOCX VERSIONS (COMPLAINT + TABLE OF CONTENTS)
###############################################################################
def generate_complaint_docx(docx_filename, firm_name, case_name, header_od, sections_od, heading_styles):
    """
    Generates a Word document version of the complaint (similar to the PDF content).
    We'll simply place:
      - The 'header_od["content"]' at the beginning (lines as paragraphs).
      - Each section heading as either main heading or subheading in bold/smaller font.
      - Each section body as standard paragraphs.
    """
    doc = Document()

    # Define base font style
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

    # Insert lines from header_od["content"]
    header_text = header_od.get("content", "")
    for line in header_text.splitlines():
        para = doc.add_paragraph()
        line_stripped = line.strip()
        # If line is all caps, we center it; else left-align
        if is_line_all_caps(line_stripped):
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        else:
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
        para.add_run(line_stripped)

    # Now add each section
    for section_key, section_body in sections_od.items():
        style_type = heading_styles.get(section_key, "section")

        # blank line before heading
        doc.add_paragraph()

        # heading
        heading_para = doc.add_paragraph()
        heading_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

        if style_type == "section":
            run = heading_para.add_run(section_key)
            run.bold = True
            run.font.size = Pt(12)
        else:
            # Subsection
            run = heading_para.add_run(section_key)
            # We'll do normal font but slightly smaller
            run.bold = False
            run.font.size = Pt(11)

        # body
        for body_line in section_body.splitlines():
            body_para = doc.add_paragraph()
            body_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run_body = body_para.add_run(body_line)
            if style_type == "section":
                run_body.font.size = Pt(12)
            else:
                run_body.font.size = Pt(11)

    doc.save(docx_filename)
    print(f"DOCX complaint saved as: {docx_filename}")

def generate_toc_docx(docx_filename, firm_name, case_name, heading_positions):
    """
    Generates a Table of Contents in DOCX form (similar to the PDF index).
    We'll simply list each heading with (page:line) on the right. Subsections
    will be smaller or non-bold.
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

    # List headings in order encountered
    for (heading_text, pg_num, ln_num, is_sub) in heading_positions:
        if is_sub:
            this_font_size = 11
            this_bold = False
        else:
            this_font_size = 12
            this_bold = True

        para = doc.add_paragraph()
        para_format = para.paragraph_format
        # We'll do a simple alignment left for text, put page:line to the far right
        tab_stops = para.tabs
        # Add a right tab stop around 6.5 inches (roughly page width margins)
        tab_stops.add_tab_stop(docx.shared.Inches(6.5), alignment=WD_ALIGN_PARAGRAPH.RIGHT)

        run_heading = para.add_run(heading_text)
        run_heading.font.size = Pt(this_font_size)
        run_heading.bold = this_bold

        # Add the "page:line" as a separate run, preceded by a tab
        run_tab = para.add_run("\t")  # to jump to the right tab stop
        run_pgline = para.add_run(f"{pg_num}:{ln_num}")
        run_pgline.font.size = Pt(this_font_size)
        run_pgline.bold = False

    doc.save(docx_filename)
    print(f"Table of Contents DOCX saved as: {docx_filename}")

###############################################################################
#  PARSING THE INPUT TEXT (HEADER, SECTIONS/SUBSECTIONS)
###############################################################################
def parse_header_and_sections(raw_text):
    """
    Anything before the first valid heading is stored in header['content'].

    A valid heading must:
        1) Match a pattern allowing either Roman numerals or digits followed by a dot,
           repeated one or more times, e.g. "I. ", "1. ", "II.1. ", etc.
        2) The text portion after the heading number must be all-caps to qualify as a heading.
    Returns:
      header_od, sections_od
    """
    header_od = OrderedDict()
    sections_od = OrderedDict()

    # Updated pattern to match either digits or Roman numerals followed by a dot,
    # repeated one or more times, then some space, then the heading text.
    # Example matches: "I. INTRO", "II. PARTIES", "1. INTRO", "1.1 SUBINTRO", "III.1.1 SOMETHING"
    heading_pattern = re.compile(r'^((?:[IVXLCDM]+\.|[0-9]+\.)+)\s+(.*)$', re.IGNORECASE)

    lines = raw_text.splitlines()
    idx = 0
    header_lines = []

    while idx < len(lines):
        line = lines[idx].rstrip('\n').rstrip('\r')
        m = heading_pattern.match(line)
        if m:
            heading_number = m.group(1).strip()
            heading_title = m.group(2).strip()
            # Check for all-caps in the heading title to confirm it's a valid heading
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

            # Confirm heading title is all caps to treat it as a heading
            if is_line_all_caps(heading_title):
                # Save previous heading's body if we had one
                if current_heading_key is not None:
                    sections_od[current_heading_key] = "\n".join(current_body_lines)
                current_body_lines = []

                # Clean trailing dot if any
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
      If the numeric portion has more than one dot (e.g. "1.1", "II.1", etc.), it's a subheading.
      Otherwise, it's a main section.
    Return a dict: { full_key: "section" or "subsection" }
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
#  MAIN PDF GENERATION WRAPPER
###############################################################################
def generate_legal_document(firm_name,
                            case_name,
                            output_filename,
                            header_od,
                            sections_od,
                            text_body,
                            exhibits,
                            heading_positions):
    """
    Generates the main PDF with line-numbered text from the combined
    header + sections, then appends exhibits.

    heading_positions is a list for storing (heading_text, page#, line#, is_subsection).
    """
    page_width, page_height = letter
    pdf_canvas = canvas.Canvas(output_filename, pagesize=letter)

    pdf_canvas.setTitle("Legal Document")
    pdf_canvas.setAuthor(firm_name)
    pdf_canvas.setSubject(case_name)
    pdf_canvas.setCreator("Legal PDF Generator")

    # 1) Determine main vs. sub sections
    heading_styles = classify_headings(sections_od)

    # 2) Prepare segment objects from the header + sections
    segments = prepare_main_pdf_segments(
        header_text=header_od.get("content", ""),
        sections_od=sections_od,
        heading_styles=heading_styles
    )

    # 3) Wrap them to ensure no overflow
    top_margin = 1.0 * inch
    bottom_margin = 1.0 * inch
    left_margin = 1.2 * inch
    right_margin = 0.5 * inch
    line_spacing = 0.25 * inch

    usable_height = page_height - (top_margin + bottom_margin)
    lines_that_fit = int(usable_height // line_spacing)
    max_lines_per_page = lines_that_fit - 2  # Some internal spacing in the box

    line_offset_x = left_margin
    line_offset_y = page_height - top_margin
    max_text_width = page_width - right_margin - line_offset_x - 0.2 * inch

    wrapped_segments = wrap_segments(pdf_canvas, segments, max_text_width)

    total_text_lines = len(wrapped_segments)
    text_pages = max(1, (total_text_lines + max_lines_per_page - 1) // max_lines_per_page)
    exhibit_pages = len(exhibits)
    total_pages = text_pages + exhibit_pages

    current_index = 0
    page_number = 1

    # Draw the main text pages
    while current_index < total_text_lines:
        next_index = draw_page_of_segments(
            pdf_canvas=pdf_canvas,
            page_width=page_width,
            page_height=page_height,
            segments=wrapped_segments,
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
        current_index = next_index
        page_number += 1
        if current_index < total_text_lines:
            pdf_canvas.showPage()

    # Draw exhibits (one page per exhibit)
    for (caption, image_path) in exhibits:
        pdf_canvas.showPage()
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
        page_number += 1

    pdf_canvas.save()

    # Once the PDF is done, also create the docx version of the complaint (similar text).
    heading_styles = classify_headings(sections_od)
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
        description="Generate a professional legal-style PDF with firm/case names, line numbering, "
                    "and optional exhibits, plus a separate table-of-contents PDF and corresponding DOCX files. "
                    "Also pickles the Lawsuit object."
    )
    parser.add_argument("--firm_name", required=True, help="Firm name placed vertically on pages.")
    parser.add_argument("--case", required=True, help="Case name placed horizontally at the top.")
    parser.add_argument("--output", default="lawsuit.pdf",
                        help="Output PDF filename for the main document (default lawsuit.pdf).")
    parser.add_argument("--file", required=True,
                        help="Path to a UTF-8 text file containing the body (with sections/subsections).")
    parser.add_argument("--exhibits", nargs='+', default=[],
                        help="Exhibit caption-text-file/image-file pairs.")
    parser.add_argument("--index", default="index.pdf",
                        help="PDF filename for the table of contents (default index.pdf).")
    parser.add_argument("--pickle", nargs='?', const=None,
                        help="Optional path to store the Lawsuit object in pickle format. "
                             "If no path is given, defaults to 'lawsuit.pickle'.")

    args = parser.parse_args()

    # Read the raw text from the file
    with open(args.file, 'r', encoding='utf-8') as f:
        raw_text = f.read()

    # Parse out the header and sections
    header_od, sections_od = parse_header_and_sections(raw_text)

    # Prepare exhibits in an OrderedDict
    if len(args.exhibits) % 2 != 0:
        raise ValueError("Exhibits must be provided in pairs: CAPTION_FILE IMAGE_FILE")

    exhibits_od = OrderedDict()
    exhibit_index = 1
    for i in range(0, len(args.exhibits), 2):
        cap_file = args.exhibits[i]
        image_path = args.exhibits[i + 1]

        with open(cap_file, 'r', encoding='utf-8') as cf:
            caption_text = cf.read()

        exhibits_od[str(exhibit_index)] = OrderedDict([
            ('caption', caption_text),
            ('image_path', image_path),
        ])
        exhibit_index += 1

    # Example extra metadata in header if needed
    header_od["DocumentTitle"] = "Complaint for Damages"
    header_od["DateFiled"] = "2025-02-14"
    header_od["Court"] = "Sample Court"

    # Build our Lawsuit object, storing the new fields automatically
    lawsuit_obj = Lawsuit(
        sections=sections_od,
        exhibits=exhibits_od,
        header=header_od,
        case_information=args.case,
        law_firm_information=args.firm_name
    )

    # Convert exhibits to pass to PDF generator
    exhibits_for_pdf = []
    for ex_key, ex_data in lawsuit_obj.exhibits.items():
        exhibits_for_pdf.append((ex_data["caption"], ex_data["image_path"]))

    # We'll track heading info for the TOC
    heading_positions = []

    # 1) Generate the main PDF (and also a corresponding .docx)
    generate_legal_document(
        firm_name=args.firm_name,
        case_name=args.case,
        output_filename=args.output,
        header_od=header_od,
        sections_od=sections_od,
        text_body=raw_text,
        exhibits=exhibits_for_pdf,
        heading_positions=heading_positions
    )

    # 2) Generate the index (table of contents) PDF
    generate_index_pdf(
        index_filename=args.index,
        firm_name=args.firm_name,
        case_name=args.case,
        heading_positions=heading_positions
    )

    # Also generate a DOCX version of the table of contents
    index_docx_filename = os.path.splitext(args.index)[0] + ".docx"
    generate_toc_docx(
        docx_filename=index_docx_filename,
        firm_name=args.firm_name,
        case_name=args.case,
        heading_positions=heading_positions
    )

    # 3) Pickle the Lawsuit object if requested
    if args.pickle is not None:
        pickle_filename = args.pickle if args.pickle else "lawsuit.pickle"
        with open(pickle_filename, 'wb') as pf:
            pickle.dump(lawsuit_obj, pf)
        pkl_path = pickle_filename
    else:
        pkl_path = "Not saved (not requested)."

    # Print summary
    print(f"PDF generated: {args.output}")
    docx_main = os.path.splitext(args.output)[0] + ".docx"
    print(f"DOCX Complaint generated: {docx_main}")
    print(f"Index PDF generated: {args.index}")
    print(f"Index DOCX generated: {index_docx_filename}")
    print(f"Lawsuit object saved to: {pkl_path}\n")

    # Print the Lawsuit object
    print("Dumped Lawsuit object:")
    print(lawsuit_obj)


if __name__ == "__main__":
    main()