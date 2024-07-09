# import docx
# from docx import Document
# from docx.oxml import OxmlElement
# from docx.oxml.ns import qn
# from docx2pdf import convert

# # Load your existing document
# doc_name = "modified_SDCIT.docx"
# document = Document(doc_name)


# # Create a new paragraph for the heading "Table of Contents"
# toc_heading = document.add_paragraph()
# toc_run = toc_heading.add_run("Table of Contents")
# toc_run.bold = True
# toc_run.font.size = docx.shared.Pt(16)  # Adjust font size if needed
# toc_run.font.name = 'Arial'  # Adjust font family if needed
# toc_heading.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
# toc_heading.space_after = docx.shared.Pt(12)  # Adjust spacing after heading

# # Create a new paragraph for the TOC in a temporary document
# temp_doc = docx.Document()
# toc_paragraph = temp_doc.add_paragraph()
# run = toc_paragraph.add_run()

# # Define the TOC field elements
# fldChar = OxmlElement('w:fldChar')
# fldChar.set(qn('w:fldCharType'), 'begin')

# instrText = OxmlElement('w:instrText')
# instrText.set(qn('xml:space'), 'preserve')
# instrText.text = 'TOC \\o "1-2" \\h \\z \\u'

# fldChar2 = OxmlElement('w:fldChar')
# fldChar2.set(qn('w:fldCharType'), 'separate')

# fldChar3 = OxmlElement('w:updateFields')
# fldChar3.set(qn('w:val'), 'true')

# fldChar4 = OxmlElement('w:fldChar')
# fldChar4.set(qn('w:fldCharType'), 'end')

# # Append the TOC field elements to the run element
# r_element = run._r
# r_element.append(fldChar)
# r_element.append(instrText)
# r_element.append(fldChar2)
# r_element.append(fldChar3)
# r_element.append(fldChar4)

# # Get the XML of the new paragraph
# toc_xml = toc_paragraph._element

# # Extract the body element of the main document
# body = document._element.body

# # Insert the heading and TOC paragraph XML at the beginning of the body
# body.insert(0, toc_heading._element)
# body.insert(1, toc_xml)

# # Save the modified document with TOC
# modified_doc_name = "modified_document_with_TOC.docx"
# document.save(modified_doc_name)


# new_doc = Document(modified_doc_name)

# # Convert the document to PDF
# pdf_name = "Finals.pdf"
# # convert(modified_doc_name, pdf_name)

# print(f"Table of Contents inserted into '{modified_doc_name}' and converted to '{pdf_name}'.")



import os as os
import sys as sys
sys.path.append(os.path.abspath("wrappers/python/XmlDocx"))
import XmlDocx as XmlDocx
document = XmlDocx.XmlDocx("water/config.xml")
document.setDocumentProperties("water/settings.xml")
document.addContent("water/content.xml")
document.setXmlDocxPath("modified_SDCIT.docx")
document.render()
