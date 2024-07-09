# from docx import Document
# from docx.shared import Pt
# from docx.oxml.ns import nsdecls
# from docx.oxml import OxmlElement
# import re
# from docx2pdf import convert

# # Function definitions

# def is_heading(paragraph):
#     # Check if paragraph starts with an integer followed by a dot and a space or an integer followed by a space
#     if not re.match(r'^\d+(\.| ) ', paragraph.text):
#         return False

#     # Check if text is bold and size is >= 12
#     for run in paragraph.runs:
#         if run.bold and run.font.size and run.font.size.pt >= 12:
#             return True
#     return False

# def replace_heading_numbering(doc):
#     current_number = 1
#     for paragraph in doc.paragraphs:
#         if is_heading(paragraph):
#             heading_text = paragraph.text.strip()
#             expected_numbering = f"{current_number}. "
            
#             # Check if current heading matches the expected numbering
#             if heading_text.startswith(expected_numbering):
#                 current_number += 1
#             else:
#                 # Replace the numbering if it doesn't match
#                 new_heading_text = f"{current_number}. {heading_text.split('. ', 1)[-1]}"
#                 paragraph.text = new_heading_text
#                 current_number += 1
            
#             # Apply bold and size 14 to heading
#             for run in paragraph.runs:
#                 run.bold = True
#                 run.font.size = Pt(14)

# def remove_content_above_first_heading(doc):
#     first_heading_index = None

#     # Find the first heading
#     for i, paragraph in enumerate(doc.paragraphs):
#         if is_heading(paragraph):
#             first_heading_index = i
#             break

#     if first_heading_index is not None:
#         # Remove all elements before the first heading
#         elements_to_remove = []
#         for i, element in enumerate(doc.element.body):
#             if i < first_heading_index:
#                 elements_to_remove.append(element)
#             else:
#                 break
        
#         for element in elements_to_remove:
#             element.getparent().remove(element)

# def remove_heading(doc, heading_text):
#     found_heading = False
#     for i, paragraph in enumerate(doc.paragraphs):
#         if is_heading(paragraph) and paragraph.text.strip() == heading_text:
#             found_heading = True
#             doc.paragraphs[i].clear()
#             doc.paragraphs[i].text = ""  # Clear heading text

#             # Remove any consecutive empty paragraphs
#             for j in range(i + 1, len(doc.paragraphs)):
#                 if doc.paragraphs[j].text.strip() == "":
#                     doc.paragraphs[j].clear()
#                     doc.paragraphs[j].text = ""
#                 else:
#                     break
            
#             break
    
#     if found_heading:
#         replace_heading_numbering(doc)

# def prompt_for_headings_to_remove(headings):
#     print("Available Headings:")
#     for idx, heading in enumerate(headings, start=1):
#         print(f"{idx}. {heading}")

#     choices = input("Enter the numbers of the headings you want to remove separated by commas (e.g., '1, 3') or enter '0' to skip: ")
#     if choices.strip() == '0':
#         return []

#     selected_indices = []
#     try:
#         selected_indices = [int(idx.strip()) for idx in choices.split(',')]
#         invalid_indices = [idx for idx in selected_indices if idx < 1 or idx > len(headings)]
#         if invalid_indices:
#             print(f"Invalid choices: {', '.join(map(str, invalid_indices))}. Please enter valid numbers.")
#             return prompt_for_headings_to_remove(headings)
#         return [headings[idx - 1] for idx in selected_indices]
#     except ValueError:
#         print("Invalid input. Please enter numbers separated by commas.")
#         return prompt_for_headings_to_remove(headings)

# # Example usage
# input_docx = 'SDCIT.docx'

# # Load the document
# doc = Document(input_docx)

# # Remove content above the first heading
# remove_content_above_first_heading(doc)

# # Replace heading numbering in series and apply bold and size 14
# replace_heading_numbering(doc)

# # Extract all headings
# headings = [paragraph.text.strip() for paragraph in doc.paragraphs if is_heading(paragraph)]

# # Prompt user for headings to remove
# headings_to_remove = prompt_for_headings_to_remove(headings)

# if headings_to_remove:
#     for heading_text in headings_to_remove:
#         # Remove heading and update numbering
#         remove_heading(doc, heading_text)

#     # Save the modified document
#     output_docx = 'modified_SDCIT.docx'
#     output_pdf = 'modified_SDCIT.pdf'
#     doc.save(output_docx)


#     convert(output_docx, output_pdf)

#     print(f"Headings removed and numbering adjusted in '{output_docx}'")
# else:
#     print("No headings selected for removal. Exiting.")


import sys
import os
import comtypes.client

wdFormatPDF = 17

in_file = 'Temp_folder/modified_SDCIT.docx'
out_file = 'Final.pdf'

word = comtypes.client.CreateObject('Word.Application')
doc = word.Documents.Open(in_file)
doc.SaveAs(out_file, FileFormat=wdFormatPDF)
doc.Close()
word.Quit()