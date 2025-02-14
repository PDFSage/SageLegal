#!/usr/bin/env python3

import argparse
import re
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib.utils import ImageReader

def wrap_text_to_lines(pdf_canvas, full_text, font_name, font_size, max_width):
    """
    Splits a large text into a list of (line_string, ended_full_line) pairs.
    Each tuple is (str, bool), where:
      str  = the actual text line,
      bool = True if this line was forced at max width, False if it ended early.
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
      - Special rule for the first page:
          If the very first line on the first page (page_number == 1 and the first line on that page)
          ends early (not forced_break), it should be centered (unless it starts with a number/bullet).
      - For all other lines, center them only if both the current line
        and the preceding line ended early (unless it starts with a number/bullet).
      - Page numbers at bottom center
      - A bounding box around the page for a more professional look

    Returns the index of the next line after this page.

    Note: Lines that start with a number or bullet point ('*', '-', '•', or digit)
    should always be left-aligned, regardless of any "ended early" checks.
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
        
        # Check the previous line's forced_break status
        if i > 0:
            _, prev_forced_break = line_tuples[i - 1]
        else:
            prev_forced_break = True  # No preceding line

        # Decide whether to center this line:
        # 1) If it's the first page and this is the first line on that page, and it ended early
        # 2) Or if this line and the previous line ended early (consecutive short lines)
        # BUT only if the line does NOT start with a number/bullet.
        # Any line starting with a digit or one of these bullets ('*', '-', '•')
        # should remain left-aligned, ignoring the centering rules.

        # Quick check for "starts with number or bullet"
        if re.match(r'^[0-9\*\-•]', text_line.strip()):
            # Force left alignment
            pdf_canvas.drawString(x_text, y_text, text_line)
        else:
            center_first_line_on_first_page = (
                page_number == 1
                and i == start_line_index
                and (not forced_break)
            )
            center_consecutive_short_lines = (
                (not forced_break)
                and (not prev_forced_break)
                and i > 0
            )

            if center_first_line_on_first_page or center_consecutive_short_lines:
                pdf_canvas.drawCentredString(page_width / 2.0, y_text, text_line)
            else:
                pdf_canvas.drawString(x_text, y_text, text_line)
        
        y_text -= line_spacing

    # Footer: page number at bottom center
    pdf_canvas.setFont("Times-Italic", 9)
    footer_text = f"Page {page_number} of {total_pages}"
    pdf_canvas.drawCentredString(page_width / 2.0, 0.5 * inch - 0.1 * inch, footer_text)

    return end_line_index

def draw_wrapped_caption(
    pdf_canvas,
    page_width,
    top_margin,
    left_margin,
    caption_text,
    font_name,
    font_size,
    line_spacing
):
    """
    Wraps the exhibit caption if it's long and draws it at the
    top of the page, left-aligned, following the same line wrap
    style as the main body text.
    Returns the final y-coordinate after drawing the caption.
    """
    # Determine the maximum width for wrapping
    max_caption_width = page_width - (2 * left_margin)
    # Wrap the caption
    wrapped_caption_lines = wrap_text_to_lines(
        pdf_canvas,
        caption_text,
        font_name=font_name,
        font_size=font_size,
        max_width=max_caption_width
    )

    current_y = top_margin
    pdf_canvas.setFont(font_name, font_size)

    for (line_text, _) in wrapped_caption_lines:
        pdf_canvas.drawString(left_margin, current_y, line_text)
        current_y -= line_spacing

    return current_y

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
    Draws a single exhibit page with the same styling:
      - Bounding box
      - Firm name vertically on the left
      - Case name at top center
      - Page numbering at bottom center
      - Exhibit caption near the top (wrapped)
      - The exhibit image placed directly beneath the caption with 1-2 lines of spacing.
    """
    # Draw an outer bounding box
    pdf_canvas.setLineWidth(2)
    pdf_canvas.rect(
        0.5 * inch,              # left
        0.5 * inch,              # bottom
        page_width - 1.0 * inch, # width
        page_height - 1.3 * inch # height
    )

    # Firm name on the left
    draw_firm_name_vertical_center(pdf_canvas, firm_name, page_width, page_height)

    # Case name at top center in bold
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

    # Now draw the exhibit caption, wrapped, a bit lower to avoid clashing with the top
    pdf_canvas.setFont("Times-Roman", 10)
    top_margin = page_height - 1.2 * inch
    left_margin = 1.2 * inch
    line_spacing = 0.25 * inch

    # Draw the wrapped caption; get the new y position
    new_y = draw_wrapped_caption(
        pdf_canvas,
        page_width,
        top_margin,
        left_margin,
        exhibit_caption,
        font_name="Times-Roman",
        font_size=10,
        line_spacing=line_spacing
    )

    # Calculate image placement directly below caption with a bit of spacing
    spacing_lines = 2
    margin = 1.0 * inch
    max_img_width = page_width - 2 * margin
    max_img_height = page_height - 2 * margin

    try:
        img_reader = ImageReader(exhibit_image)
        img_width, img_height = img_reader.getSize()
    except Exception as e:
        # If there's an error reading the image, place a message
        pdf_canvas.setFont("Times-Italic", 10)
        pdf_canvas.drawCentredString(
            page_width / 2.0, 
            page_height / 2.0, 
            f"Unable to load image: {exhibit_image}\nError: {e}"
        )
    else:
        # Scale proportionally
        scale = min(max_img_width / img_width, max_img_height / img_height, 1.0)
        new_width = img_width * scale
        new_height = img_height * scale

        # The top of the image (y_img_top) is 2 line-spacings below the final caption line
        y_img_top = new_y - (spacing_lines * line_spacing)
        y_img_bottom = y_img_top - new_height

        # Ensure it doesn't drop below our chosen bottom margin
        bottom_margin = 1.0 * inch
        if y_img_bottom < bottom_margin:
            y_img_bottom = bottom_margin
            y_img_top = y_img_bottom + new_height

        # Center the image horizontally
        x_img = (page_width - new_width) / 2.0

        pdf_canvas.drawImage(
            img_reader,
            x_img,
            y_img_bottom,
            width=new_width,
            height=new_height,
            preserveAspectRatio=True,
            anchor='c'
        )

    # Footer: page number
    pdf_canvas.setFont("Times-Italic", 9)
    footer_text = f"Page {page_number} of {total_pages}"
    pdf_canvas.drawCentredString(page_width / 2.0, 0.5 * inch - 0.1 * inch, footer_text)

def generate_legal_document(
    firm_name,
    case_name,
    output_filename,
    text_body,
    exhibits
):
    """
    Generates a legal-style PDF document with:
      - Vertical firm name on the left
      - Case name top center, bold, with a horizontal rule
      - Numbered lines on left and right for the text
      - Wrapped text, Times-Roman, size 10
      - On the first page only, the first line is centered if it ends early (unless it starts with a number/bullet)
      - For subsequent lines, center them only if both they and their preceding line ended early (unless it starts with a number/bullet)
      - Page numbering at bottom center ("Page X of Y")
      - A bounding box around each page
      - Following the main text, any number of exhibits (each on its own page),
        where the exhibit caption is also wrapped and placed a bit lower to avoid clashing at the top,
        then the exhibit image is inserted directly underneath with some spacing.

    Also, any line that starts with a number or bullet ('*', '-', '•') is always left-aligned.
    """
    page_width, page_height = letter

    # Create the PDF canvas and set metadata
    pdf_canvas = canvas.Canvas(output_filename, pagesize=letter)
    pdf_canvas.setTitle("Legal Document")
    pdf_canvas.setAuthor(firm_name)
    pdf_canvas.setSubject(case_name)
    pdf_canvas.setCreator("Legal PDF Generator")

    # Margins and spacing
    top_margin = 1.0 * inch
    bottom_margin = 1.0 * inch
    left_margin = 1.2 * inch
    right_margin = 0.5 * inch
    line_spacing = 0.25 * inch

    # Calculate how many lines fit into the vertical space
    usable_height = page_height - (top_margin + bottom_margin)
    lines_that_fit = int(usable_height // line_spacing)
    # Reserve a couple lines for spacing / boundary
    max_lines_per_page = lines_that_fit - 2

    # Coordinates for text start
    line_offset_x = left_margin
    line_offset_y = page_height - top_margin

    # Maximum width for wrapped text
    max_text_width = page_width - right_margin - line_offset_x - 0.2 * inch

    # Wrap the main text
    wrapped_lines = wrap_text_to_lines(
        pdf_canvas,
        text_body,
        font_name="Times-Roman",
        font_size=10,
        max_width=max_text_width
    )

    total_lines = len(wrapped_lines)
    # How many pages for the main text?
    text_pages = (total_lines + max_lines_per_page - 1) // max_lines_per_page

    # One page per exhibit
    exhibit_pages = len(exhibits)

    # Total pages = pages for text + pages for exhibits
    total_pages = text_pages + exhibit_pages

    current_line_index = 0
    page_number = 1

    # --- Paginate the main text ---
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
        
        # Only advance to a new page if there's more text
        if current_line_index < total_lines:
            pdf_canvas.showPage()

    # --- Add exhibits, each on its own fresh page ---
    for (caption, image_path) in exhibits:
        pdf_canvas.showPage()
        draw_exhibit_page(
            pdf_canvas,
            page_width,
            page_height,
            firm_name,
            case_name,
            caption,
            image_path,
            page_number,
            total_pages
        )
        page_number += 1

    # Finalize the PDF
    pdf_canvas.save()

def main():
    parser = argparse.ArgumentParser(
        description="Generate a professional legal-style PDF with a firm name, case name, line numbering, and optional exhibits."
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
    #
    # We now support passing multiple exhibit caption/image pairs in one go, like:
    #   --exhibits e1 pokimane.jpg e2 pokimanellc.jpg e3 johndoemirror.jpg
    #
    # This is done by specifying nargs='+' and then processing them in pairs.
    #
    parser.add_argument(
        "--exhibits",
        nargs='+',
        default=[],
        help=(
            "Exhibit caption-text-file / image-file pairs. Example:\n"
            "  --exhibits e1 pokimane.jpg e2 pokimanellc.jpg e3 johndoemirror.jpg\n"
            "Each pair is (CAPTION_FILE, IMAGE_PATH)."
        ),
    )
    
    args = parser.parse_args()
    
    # Read the main document text
    with open(args.file, 'r', encoding='utf-8') as f:
        text_body = f.read()
    
    # The exhibits list must come in pairs
    exhibits_cli_list = args.exhibits
    if len(exhibits_cli_list) % 2 != 0:
        raise ValueError(
            "Exhibits must be provided in pairs: caption_file image_file "
            "(e.g., e1 pokimane.jpg e2 pokimanellc.jpg)"
        )

    exhibits = []
    for i in range(0, len(exhibits_cli_list), 2):
        cap_file = exhibits_cli_list[i]
        image_path = exhibits_cli_list[i+1]
        with open(cap_file, 'r', encoding='utf-8') as cf:
            caption_text = cf.read()
        exhibits.append((caption_text, image_path))

    generate_legal_document(
        firm_name=args.firm_name,
        case_name=args.case,
        output_filename=args.output,
        text_body=text_body,
        exhibits=exhibits
    )

if __name__ == "__main__":
    main()