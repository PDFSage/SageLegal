#!/usr/bin/env python3

import argparse
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch

def wrap_text_to_lines(pdf_canvas, full_text, font_name, font_size, max_width):
    """
    Splits a large text into a list of (line_string, ended_full_line) pairs.
    Each tuple is (str, bool), where:
      str  = the actual text line,
      bool = True if this line was forced at max width, False if it ended early.
    This is used later to detect lines that "end early" vs. lines that
    went all the way to the end.
    """
    pdf_canvas.setFont(font_name, font_size)

    paragraphs = full_text.split('\n')
    all_lines = []

    for paragraph in paragraphs:
        words = paragraph.split()
        if not words:
            # Blank line or empty paragraph
            all_lines.append(("", False))  # preserve empty line
            continue

        current_line = ""
        for word in words:
            if not current_line:
                current_line = word
            else:
                test_line = current_line + " " + word
                if pdf_canvas.stringWidth(test_line, font_name, font_size) <= max_width:
                    current_line = test_line
                else:
                    # The current line is at or near the max width (forced break)
                    all_lines.append((current_line, True))
                    current_line = word

        if current_line:
            # This line ended early (not forced by max width)
            all_lines.append((current_line, False))

    return all_lines

def draw_firm_name_vertical_center(pdf_canvas, text, page_width, page_height):
    """
    Draws the firm name vertically, centered along the page height (left margin).
    Moves the text near the left edge.
    """
    pdf_canvas.saveState()
    pdf_canvas.setFont("Times-Bold", 10)

    text_width = pdf_canvas.stringWidth(text, "Times-Bold", 10)
    
    # Position near the left edge
    x_pos = 0.2 * inch
    
    # Center vertically
    y_center = page_height / 2.0
    y_pos = y_center - (text_width / 2.0)
    
    # Translate, rotate 90°, then draw
    pdf_canvas.translate(x_pos, y_pos)
    pdf_canvas.rotate(90)
    pdf_canvas.drawString(0, 0, text)
    pdf_canvas.restoreState()

def draw_page_content(
    pdf_canvas,
    page_width,
    page_height,
    line_tuples,
    start_line_index,
    max_lines_per_page,
    firm_name,
    case_name,
    line_offset_x,
    line_offset_y,
    line_spacing,
    page_number,
    total_pages
):
    """
    Draw up to `max_lines_per_page` lines on the given canvas, starting
    from `start_line_index` in `line_tuples`. Each element in line_tuples
    is (line_text, ended_full_line_bool).

    Adds:
      - Firm name vertically on the left, centered along the page's height
      - Case name at top center, with a horizontal rule below it
      - Numbered lines on both left edge and right edge
      - Centering logic for short lines that follow another short line
      - BUT: If it's the first line of a new page and the previous line ended early, do NOT center.
      - Page numbers at bottom center
      - A bounding box around the page for a professional look

    Returns the index of the next line after this page.
    """

    # Draw an outer bounding box, lowered so it doesn't overlap the case name
    pdf_canvas.setLineWidth(2)
    pdf_canvas.rect(
        0.5 * inch,                  # left
        0.5 * inch,                  # bottom
        page_width - 1.0 * inch,     # width
        page_height - 1.3 * inch     # height (lowered top edge)
    )

    # Firm name on the left
    draw_firm_name_vertical_center(pdf_canvas, firm_name, page_width, page_height)
    
    # Case name at top center in bold, slightly larger
    pdf_canvas.setFont("Times-Bold", 12)
    pdf_canvas.drawCentredString(page_width / 2.0, page_height - 0.5 * inch, case_name)
    
    # Horizontal line under the case name
    pdf_canvas.setLineWidth(1)
    pdf_canvas.line(
        0.5 * inch, 
        page_height - 0.6 * inch, 
        page_width - 0.5 * inch, 
        page_height - 0.6 * inch
    )
    
    # Reset font for the body text
    pdf_canvas.setFont("Times-Roman", 10)

    # Starting position for the body text
    x_text = line_offset_x
    y_text = line_offset_y
    
    end_line_index = min(start_line_index + max_lines_per_page, len(line_tuples))
    
    for i in range(start_line_index, end_line_index):
        line_number_str = f"{i + 1}"
        text_line, forced_break = line_tuples[i]
        
        # Left numbering
        pdf_canvas.drawString(x_text - 0.6 * inch, y_text, line_number_str)
        # Right numbering
        pdf_canvas.drawString(page_width - 0.4 * inch, y_text, line_number_str)
        
        # We want to figure out if the current line ended early or not.
        line_ended_early = not forced_break

        # The previous line ended early if i>0 and the stored bool was False
        # (because forced_break == True means it reached full width).
        if i > 0:
            _, prev_forced_break = line_tuples[i - 1]
            prev_line_ended_early = not prev_forced_break
        else:
            prev_line_ended_early = False  # No previous line at all

        # Check if this is the first line on the current page
        line_is_first_on_page = (i == start_line_index)

        # Logic:
        # - If a line ends early and directly follows another short line (on the same page), center it.
        # - Normally, if it's the first line of a page and ends early, we center it.
        # - EXCEPTION: If it's the first line of a new page AND the previous page's last line ended early,
        #              we do NOT center (left-align instead).
        if line_ended_early:
            # If it's the first line on the page and the previous line ended early,
            # do NOT center.
            if line_is_first_on_page and prev_line_ended_early and i > 0:
                pdf_canvas.drawString(x_text, y_text, text_line)
            # If it's the first line on the page, but the last line from the previous page was forced
            # or there's no previous line at all, then center.
            elif line_is_first_on_page:
                pdf_canvas.drawCentredString(page_width / 2.0, y_text, text_line)
            # If it's not the first line on the page and the previous line was also short, center it.
            elif prev_line_ended_early:
                pdf_canvas.drawCentredString(page_width / 2.0, y_text, text_line)
            else:
                # Normal short line that doesn't follow another short line => left align
                pdf_canvas.drawString(x_text, y_text, text_line)
        else:
            # forced_break == True => it used the entire width => always left align
            pdf_canvas.drawString(x_text, y_text, text_line)
        
        y_text -= line_spacing

    # Footer: page number at bottom center
    pdf_canvas.setFont("Times-Italic", 9)
    footer_text = f"Page {page_number} of {total_pages}"
    pdf_canvas.drawCentredString(page_width / 2.0, 0.5 * inch - 0.1 * inch, footer_text)

    return end_line_index

def generate_legal_document(firm_name, case_name, output_filename, text_body):
    """
    Generates a legal-style PDF document with:
      - Vertical firm name on left
      - Case name top center, bold, with a horizontal rule
      - Numbered lines on left and right
      - Wrapped text, Times-Roman, size 10
      - Centering logic for short lines
      - Exception: If the first line on a new page follows a short line on the previous page,
                   do NOT center it.
      - Page numbering at bottom center
      - A bounding box around each page
    """
    page_width, page_height = letter

    # Create the PDF canvas with metadata
    pdf_canvas = canvas.Canvas(output_filename, pagesize=letter)
    pdf_canvas.setTitle("Legal Document")
    pdf_canvas.setAuthor(firm_name)
    pdf_canvas.setSubject(case_name)
    pdf_canvas.setCreator("Legal PDF Generator")

    # Margins
    top_margin = 1.0 * inch
    bottom_margin = 1.0 * inch
    left_margin = 1.2 * inch
    right_margin = 0.5 * inch

    # Line spacing
    line_spacing = 0.25 * inch

    # Calculate how many lines fit in the vertical space,
    # leaving a little extra at the bottom for the footer
    usable_height = page_height - (top_margin + bottom_margin)
    lines_that_fit = int(usable_height // line_spacing)
    max_lines_per_page = lines_that_fit - 2  # reserve a couple lines for safety

    # Coordinates for text start
    line_offset_x = left_margin
    line_offset_y = page_height - top_margin

    # Maximum width for wrapped text (account for margins)
    max_text_width = page_width - right_margin - line_offset_x - 0.2 * inch

    # Convert the raw text into wrapped lines
    wrapped_lines = wrap_text_to_lines(
        pdf_canvas,
        text_body,
        font_name="Times-Roman",
        font_size=10,
        max_width=max_text_width
    )

    total_lines = len(wrapped_lines)
    total_pages = (total_lines + max_lines_per_page - 1) // max_lines_per_page
    
    current_line_index = 0
    page_number = 1

    # Paginate
    while current_line_index < total_lines:
        next_line_index = draw_page_content(
            pdf_canvas,
            page_width,
            page_height,
            wrapped_lines,
            current_line_index,
            max_lines_per_page,
            firm_name,
            case_name,
            line_offset_x,
            line_offset_y,
            line_spacing,
            page_number,
            total_pages
        )
        current_line_index = next_line_index
        page_number += 1
        
        # If there are more lines to draw, start a new page
        if current_line_index < total_lines:
            pdf_canvas.showPage()
    
    # Save the final PDF
    pdf_canvas.save()

def main():
    parser = argparse.ArgumentParser(
        description="Generate a professional legal-style PDF with firm name, case name, and custom line centering rules."
    )
    parser.add_argument(
        "--firm_name",
        required=True,
        help="Firm name to be placed vertically near the left edge, centered vertically."
    )
    parser.add_argument(
        "--case",
        required=True,
        help="Case name to be placed horizontally at the top center of each page."
    )
    parser.add_argument(
        "--output",
        required=True,
        help="Output PDF filename."
    )
    parser.add_argument(
        "--file",
        required=True,
        help="Path to a text file (UTF-8) containing the body of the document."
    )
    
    args = parser.parse_args()
    
    # Read the text from file
    with open(args.file, 'r', encoding='utf-8') as f:
        text_body = f.read()
    
    generate_legal_document(
        firm_name=args.firm_name,
        case_name=args.case,
        output_filename=args.output,
        text_body=text_body
    )

if __name__ == "__main__":
    main()