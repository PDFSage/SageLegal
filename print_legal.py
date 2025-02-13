#!/usr/bin/env python3

import argparse
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch

def wrap_text_to_lines(pdf_canvas, full_text, font_name, font_size, max_width):
    """
    Splits a large text into a list of (line_string, ended_full_line) pairs.
    Each tuple is:
       (str, bool)
    where `str` is the actual text line,
          `bool` indicates whether this line was forced at max width (True)
                 or it ended early because there were no more words (False).

    This is used later to detect lines that "end early" vs. lines that
    "went all the way to the end".
    """
    pdf_canvas.setFont(font_name, font_size)

    paragraphs = full_text.split('\n')
    all_lines = []

    for paragraph in paragraphs:
        words = paragraph.split()
        if not words:
            # Blank line or empty paragraph
            all_lines.append(("", False))  # preserve empty line, mark as ended early
            continue

        current_line = ""
        for word in words:
            if not current_line:
                # Start the line with the current word
                current_line = word
            else:
                # Test if adding this word would exceed max_width
                test_line = current_line + " " + word
                if pdf_canvas.stringWidth(test_line, font_name, font_size) <= max_width:
                    current_line = test_line
                else:
                    # The current line is at or near the max width (forced break)
                    all_lines.append((current_line, True))
                    current_line = word

        # After the loop, if there's something in current_line, add it
        # Since we didn't force a break here, it ended early
        if current_line:
            all_lines.append((current_line, False))

    return all_lines

def draw_firm_name_vertical_center(pdf_canvas, text, page_width, page_height):
    """
    Draws the firm name vertically, centered along the page height (left margin).
    Moves the text near the left edge.
    """
    pdf_canvas.saveState()
    pdf_canvas.setFont("Times-Roman", 10)

    text_width = pdf_canvas.stringWidth(text, "Times-Roman", 10)
    
    # Move firm name near the left edge:
    x_pos = 0.2 * inch
    
    # Vertically center:
    y_center = (page_height / 2.0)
    y_pos = y_center - (text_width / 2.0)
    
    # Translate, then rotate 90° to place text vertically
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
    line_spacing
):
    """
    Draw up to `max_lines_per_page` lines on the given canvas, starting
    from `start_line_index` in `line_tuples`. Each element in line_tuples
    is (line_text, ended_full_line_bool).

    Adds line numbers on both left and right edges.
    Places the firm name vertically centered on the left side,
    and the case name at the top center horizontally.

    Returns the next line index after drawing (to continue on subsequent pages).

    Additional requirement:
    - Move left line numbering further left (compared to original code).
    - For a line that ends early (ended_full_line_bool = False), and if
      the previous line also ended early, center that line.
      (Interpreted as a possible "title line".)
    """
    pdf_canvas.setFont("Times-Roman", 10)
    
    # Firm name on the left
    draw_firm_name_vertical_center(pdf_canvas, firm_name, page_width, page_height)
    
    # Case name at top center
    pdf_canvas.drawCentredString(page_width / 2.0, page_height - 0.5 * inch, case_name)
    
    # Starting position for body lines
    x_text = line_offset_x
    y_text = line_offset_y
    
    end_line_index = min(start_line_index + max_lines_per_page, len(line_tuples))
    
    for i in range(start_line_index, end_line_index):
        line_number_str = f"{i + 1}"
        text_line, forced_break = line_tuples[i]
        
        # Move left numbering further to the left compared to original (e.g. -0.6 inch)
        # Adjust as desired (here we do -0.6 inch from x_text)
        pdf_canvas.drawString(x_text - 0.6 * inch, y_text, line_number_str)
        
        # Right line number (near the far right)
        pdf_canvas.drawString(page_width - 0.4 * inch, y_text, line_number_str)
        
        # Check if this line is short, and also the previous line was short
        # (meaning the previous line ended early as well).
        # If so, we center this line.
        if i > 0:
            _, prev_forced_break = line_tuples[i - 1]
        else:
            prev_forced_break = True  # no previous line, so default to True

        if (not forced_break) and (not prev_forced_break):
            # Both this line and the previous line ended early: center it
            pdf_canvas.drawCentredString(page_width / 2.0, y_text, text_line)
        else:
            # Normal line draw (left aligned)
            pdf_canvas.drawString(x_text, y_text, text_line)
        
        # Move up for next line
        y_text -= line_spacing
    
    return end_line_index

def generate_legal_document(firm_name, case_name, output_filename, text_body):
    """
    Generates a legal-style PDF document:
     - Firm name placed vertically near the left edge, centered vertically.
     - Case name placed horizontally at the top center of each page.
     - Numbered lines on both left (near firm name) and far right edges.
     - Automatically wraps the text to fit the page width.
     - Uses Times-Roman font, size 10.
     - Left line numbering is moved further left.
     - If a line ends early and is also preceded by a line that ends early,
       that line is centered (treated as a title-like line).
    """
    page_width, page_height = letter
    
    # Create PDF
    pdf_canvas = canvas.Canvas(output_filename, pagesize=letter)
    
    # Layout parameters
    max_lines_per_page = 40
    line_spacing = 0.25 * inch
    
    # Text starts further in from the left so there's space for firm name and line numbers
    line_offset_x = 1.2 * inch  # left margin for the actual text
    line_offset_y = page_height - 1.0 * inch  # top margin for first line
    
    # We'll consider the right line number around (page_width - 0.4 inch),
    # so define a safe text width:
    max_text_width = (page_width - 0.4 * inch) - line_offset_x - 0.2 * inch

    # Wrap the text, get (line_text, forced_break) pairs
    wrapped_lines = wrap_text_to_lines(
        pdf_canvas,
        text_body,
        font_name="Times-Roman",
        font_size=10,
        max_width=max_text_width
    )
    
    current_line_index = 0
    total_lines = len(wrapped_lines)

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
            line_spacing
        )
        current_line_index = next_line_index
        
        if current_line_index < total_lines:
            pdf_canvas.showPage()
    
    # Save file
    pdf_canvas.save()

def main():
    parser = argparse.ArgumentParser(
        description="Generate a legal-style PDF with a firm name, case name, and line numbering."
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
    
    # Read the UTF-8 text from file
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