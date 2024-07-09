from pdf2docx import Converter, parse
pdf_file_path = 'Maker Datasheet.pdf'
Docx_file_path = 'SDCIT.docx'

# parse(pdf_file=pdf_file_path,docx_file=Docx_file_path)
cv = Converter(pdf_file_path)
cv.convert(Docx_file_path)
cv.close()



# from docx import Document
# import os
# import re

# def is_heading(paragraph):
#     # Check if paragraph starts with an integer followed by a dot and a space or an integer followed by a space
#     if not re.match(r'^\d+(\.| ) ', paragraph.text):
#         return False

#     # Check if text is bold and size is >= 12
#     for run in paragraph.runs:
#         if run.bold and run.font.size and run.font.size.pt >= 12:
#             return True
#     return False

# def remove_content_under_heading(doc, heading_to_remove):
#     remove_next = False
#     for paragraph in doc.paragraphs:
#         if is_heading(paragraph) and paragraph.text.strip() == heading_to_remove:
#             # Found the heading to remove content under
#             remove_next = True
#         elif remove_next:
#             # Remove content under this heading until next heading or end of document
#             if is_heading(paragraph):
#                 break  # Stop removing content at the next heading
#             else:
#                 paragraph.clear()  # Clear content of non-heading paragraphs

# # Function to prompt user for heading to remove
# def prompt_for_heading_to_remove(headings):
#     print("Available Headings:")
#     for idx, heading in enumerate(headings, start=1):
#         print(f"{idx}. {heading}")
    
#     while True:
#         choice = input("Enter the number of the heading you want to remove (or '0' to skip): ")
#         try:
#             choice_idx = int(choice)
#             if 0 <= choice_idx <= len(headings):
#                 if choice_idx == 0:
#                     return None  # User chose to skip
#                 return headings[choice_idx - 1]  # Return selected heading
#             else:
#                 print("Invalid choice. Please enter a valid number.")
#         except ValueError:
#             print("Invalid input. Please enter a number.")

# # Example usage
# input_docx = 'SDCIT.docx'

# # Load the document
# doc = Document(input_docx)

# # Extract all headings
# headings = [paragraph.text.strip() for paragraph in doc.paragraphs if is_heading(paragraph)]

# # Prompt user for heading to remove
# heading_to_remove = prompt_for_heading_to_remove(headings)

# if heading_to_remove:
#     # Remove content under specified heading
#     remove_content_under_heading(doc, heading_to_remove)

#     # Save the modified document
#     output_docx = 'modified_SDCIT.docx'
#     doc.save(output_docx)

#     print(f"Content under heading '{heading_to_remove}' removed in '{output_docx}'")
# else:
#     print("No heading selected for removal. Exiting.")
 