import re
from docx import Document
def is_sub_subheading(paragraph):
    for line in paragraph.text.splitlines():
        if re.match(r'^\d+\.\d+\.\d+', line):
            print(line)
            if paragraph.style.name == 'Heading 3':
                return True
            for run in paragraph.runs:
                if run.text in line or run.bold:
                    return True
    return False


doc = Document(r'output\0eS15dhUsz\maker datasheet.docx')

for para in doc.paragraphs:
    is_sub_subheading(para)