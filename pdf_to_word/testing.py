import fitz
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Spacer, Paragraph
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from io import BytesIO
from docx import Document

def generate_heading_list(doc):
    heading_list = {}

    for paragraph in doc.paragraphs:
        # Check if the paragraph style is 'Heading 1' or 'Heading 2'
        if paragraph.style.name in ['Heading 1', 'Heading 2']:
            heading_text = paragraph.text.strip()
            heading_level = paragraph.style.name
            heading_list[heading_text] = [heading_level]

    return heading_list

def get_heading_page_numbers_in_pdf(pdf_path, headings):
    doc = fitz.open(pdf_path)
    heading_page_numbers = {heading: [] for heading in headings}

    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        text = page.get_text("text")

        for heading_text in headings:
            if heading_text in text and page_num + 1 not in heading_page_numbers[heading_text]:
                headings[heading_text].append(page_num)



def create_index_page(pdf_path, output_pdf, headings_with_pages):
    try:
        # Open the PDF
        doc = fitz.open(pdf_path)
        first_page = doc[0]

        # Set styles
        title_fontsize = 20
        normal_fontsize = 12
        line_spacing = 12
        heading1_fontsize = 14
        heading2_fontsize = 12

        # Calculate center alignment for "TABLE OF CONTENTS" title
        page_width = first_page.rect.width
        text_width = fitz.get_text_length("TABLE OF CONTENTS", fontsize=title_fontsize)
        x_title = (page_width - text_width) / 2
        y_title = 50  # Starting y-coordinate for the title

        # Add the index title with center alignment
        first_page.insert_text((x_title, y_title), "TABLE OF CONTENTS", fontsize=title_fontsize, set_simple=True)

        # Starting y-coordinate for entries below the title
        y = y_title + title_fontsize + line_spacing * 2


        # Function to add index entry with left alignment
        def add_index_entry(page, x, y, text, page_number, font_size, line_spacing, level):
            text_width = fitz.get_text_length(text, fontsize=font_size)
            page_number_width = fitz.get_text_length(str(page_number), fontsize=font_size)
            
            # Calculate dots count based on available space
            if level == 'Heading 1':
                dots_count = max(1, int((page_width - x - text_width - page_number_width - 30) / 4))  # Adjust 30 for spacing
            elif level == 'Heading 2':
                dots_count = max(1, int((page_width - x - text_width - page_number_width)- 10 / 4))  # Adjust 40 for additional indent
            
            # Insert text, dots, and page number with right-aligned page number
            page.insert_text((x, y), text, fontsize=font_size, set_simple=True, render_mode=0)
            
            # Calculate the position for dots
            dots_x = x + text_width + 5  # Adjust 5 for space between text and dots
            page.insert_text((dots_x, y), "." * dots_count, fontsize=font_size, set_simple=True, render_mode=0)
            
            # Calculate the position for right-aligned page number
            page_number_x = page_width - 30 - page_number_width  # Adjust 30 for space before page number
            page.insert_text((page_number_x, y), str(page_number), fontsize=font_size, set_simple=True, render_mode=0)

        # Add the headings and page numbers with left alignment
        for heading, (level, page_number) in headings_with_pages.items():
            if level == 'Heading 1':
                add_index_entry(first_page, x_title - 140, y, heading, page_number, heading1_fontsize, line_spacing,level)
            elif level == 'Heading 2':
                add_index_entry(first_page, x_title - 120, y, heading, page_number, heading2_fontsize, line_spacing,level)
            y += normal_fontsize + line_spacing

        # Save the updated PDF
        doc.save(output_pdf)
        doc.close()

        print(f"Index page created and saved to {output_pdf}")

    except Exception as e:
        print(f"Error: {e}")




# Example usage:
# headings_with_pages = {"Chapter 1": [1, 2, 3], "Chapter 2": [4, 5, 6]}
# doc_path = 'Temp_folder/modified_SDCIT.docx'
# doc = Document(doc_path)
# headings = generate_heading_list(doc)

# pdf_path = 'Final.pdf'
# get_heading_page_numbers_in_pdf(pdf_path, headings)

# create_index_page("Final.pdf", "output.pdf", headings)
