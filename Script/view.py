from .helper import *
from docx import Document
from docx2pdf import convert
import os
import shutil
import random
import string


def Get_New_PDF(input_pdf_path,Header_image_path):
    try:
        Folder_name = 'output/'+ ''.join(random.choices(string.ascii_letters + string.digits, k=10))
        output_docx = Folder_name+'/modified_SDCIT.docx'
        without_matermark_pdf = Folder_name+'/Without_watermark.pdf'
        output_pdf = Folder_name+'/output.pdf'
        final_pdf = Folder_name+'/Final.pdf'
        os.makedirs(Folder_name,exist_ok=True)
        width_cm = 5.00
        height_cm = 1.44
        without_header_footer_file_path = remove_header_footer(input_pdf_path,Folder_name+'/without_header_footer.pdf',70, 70)
        docx_path =  convert_docx_to_pdf_windows(without_header_footer_file_path,Folder_name)
        doc = Document(docx_path)
        remove_header_footer(doc)
        Format_doc(doc)
        start_each_heading_from_new_page(doc)
        start_each_heading_from_new_line(doc)
        # Remove content above the first heading
        remove_content_above_first_heading(doc)
        headings = [paragraph.text.strip().split('\n')[0] for paragraph in doc.paragraphs if is_heading(paragraph)]
        headings_to_remove = prompt_for_headings_to_remove(headings)
        version = input('Please Version of Document: ')
        if headings_to_remove:
            remove_headings_with_content(doc,headings_to_remove)
        set_page_size_to_a4(doc)
        # Add header and footer
        remove_empty_and_excessive_spaces(doc)
        start_each_heading1_from_new_page(doc)
        replace_heading_numbering(doc)
        add_header_with_image_size(doc, Header_image_path,width_cm,height_cm )
        add_page_border(doc)
        create_index_of_heading(doc)
        add_footer_with_page_number(doc)
        doc.save(output_docx)
        update_toc_with_win32(output_docx)
        convert_docx_to_pdf(output_docx,without_matermark_pdf)
        add_custom_page_at_start(without_matermark_pdf,output_pdf,Header_image_path,version)
        add_watermark_to_pdf(output_pdf,final_pdf)
        print('Final PDf At', final_pdf)
        print('Docx At:', output_docx)
    except Exception as e:
        print('Error',e)




Get_New_PDF('Static\maker datasheet.pdf','Static\\flexon_logo.png')
# Get_New_PDF('Static\Marker datasheet With Watermark.pdf','Static\\flexon_logo.png')