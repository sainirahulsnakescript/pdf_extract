import fitz
import os
import io
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.colors import black

def create_blank_page_with_logo_and_watermark(logo=True):
    
        watermark_text="FLEXXON CONFIDENTIAL"
        header_logo_path='flexon_logo.png'
        buffer = io.BytesIO()
        c = canvas.Canvas(buffer, pagesize=A4)
        width, height = A4

        # Calculate space needed for watermark text
        watermark_font_size = 50
        text_width = c.stringWidth(watermark_text, "Helvetica", watermark_font_size)
        text_height = watermark_font_size # Add some padding

        # Draw border around the page
        border_padding =15
        border_width =1.1  # Adjust border width as needed
        c.setStrokeColorRGB(0, 0, 0)  # Border color (black in this case)
        c.setLineWidth(border_width)
        c.rect(border_padding, border_padding, width - 2*border_padding, height - 2*border_padding)

        # Calculate position for logo
        logo_width = 140
        logo_height = 45
        logo_x = width - logo_width - border_padding-380
        logo_y = height - logo_height - border_padding-3  # Position near top with padding
        #draw logo
        if logo == True:
            c.drawImage(header_logo_path, logo_x, logo_y, width=logo_width, height=logo_height)
        c.setFillColorRGB(0.8,0.8,0.8, alpha=0.3)  # Adjust alpha value as needed (0.0 for fully transparent, 1.0 for fully opaque)

        available_space = logo_y - border_padding   
        # Add top margin for data
        top_margin = max(100, available_space // 3)  # Adjust as needed
        bottom_margin = max(100, available_space // 3)  # Adjust as needed


        #clacuate the remaing space for pdf data
        available_space -= top_margin + bottom_margin

        # Move the origin to the top-left corner of the content area
        c.translate(border_padding, border_padding + bottom_margin + available_space)

        # Move the origin back to the bottom-left corner of the page for watermark
        c.translate(0, -(border_padding + bottom_margin + available_space))

        # Move the origin to the center of the page for rotation
        c.translate(width / 2, height / 2)
        # Rotate the canvas
        c.rotate(45)
        # Calculate centered position for watermark text with margins
        text_width = c.stringWidth(watermark_text, "Helvetica", watermark_font_size)
        x_position = (width - text_width) / 2 -10
        y_position = available_space / 2 -160  # Adjusted y-position for centering

        # Draw outlined watermark text with margins
        c.setFont("Helvetica", watermark_font_size)
        c.setLineWidth(5)  # Adjust outline width as needed
        c.setStrokeColor(black)  # Outline color
        c.drawCentredString(x_position, y_position, watermark_text)

        # Draw filled watermark text on top of outline
        c.setFillColorRGB(0,0,0, alpha=0.23)
        c.drawCentredString(x_position, y_position, watermark_text)

        c.showPage()
        c.save()
        buffer.seek(0)   
        return buffer.getvalue()# Return the PDF content as bytes

# pdf_content = create_blank_page_with_logo_and_watermark()
# with open('output.pdf', 'wb') as f:
#     f.write(pdf_content)


def apply_watermark(new_pdf_path1 = 'modified_SDCIT.pdf'):
    doc = fitz.open(new_pdf_path1)
    
    template_pdf_content = create_blank_page_with_logo_and_watermark(logo=True)
    template_doc = fitz.open("pdf", template_pdf_content)
    template_page = template_doc.load_page(0)
    template_doc_without_logo = fitz.open("pdf", create_blank_page_with_logo_and_watermark(logo=False))

    output_pdf = fitz.open()

    for page_num in range(doc.page_count):
        page = doc.load_page(page_num)
        rect = page.rect
        
        new_page = output_pdf.new_page(width=rect.width, height=rect.height)

        new_page.show_pdf_page(rect, doc, page_num)
        if page_num == 0:
            new_page.show_pdf_page(rect, template_doc_without_logo, 0)
        else:
           new_page.show_pdf_page(rect, template_doc, 0)

    new_pdf_path = os.path.splitext(new_pdf_path1)[0] + "_templated.pdf"
    
    output_pdf.save(new_pdf_path)
    return new_pdf_path


apply_watermark()