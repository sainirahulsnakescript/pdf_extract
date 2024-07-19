# PDF to DOCX Conversion and Formatting Script

This project provides a Python script to convert PDF files to DOCX, perform various formatting tasks on the DOCX files, and then convert them back to PDF. The script also allows adding custom headers, footers, and watermarks to the final PDF.

## Features

- Convert PDF files to DOCX format.
- Remove headers and footers from the input PDF.
- Apply various formatting changes to the DOCX file:
  - Start each heading from a new page.
  - Start each heading from a new line.
  - Remove content above the first heading.
  - Remove specified headings and their content.
  - Set page size to A4.
  - Remove empty and excessive spaces.
  - Replace heading numbering.
  - Add custom headers with images.
  - Add page borders.
  - Create an index of headings.
  - Add footers with page numbers.
- Convert the modified DOCX file back to PDF.
- Add a custom page and watermark to the final PDF.

## Requirements

- Python 3.7 or higher
- `python-docx`
- `docx2pdf`
- `shutil`
- `random`
- `string`
- `os`
- `win32com.client`

## Installation

1. Clone the repository:
    ```sh
    git clone https://github.com/yourusername/pdf-docx-formatting.git
    ```

2. Install the required packages:
    ```sh
    pip install python-docx docx2pdf pywin32
    ```

## Usage

1. Place your input PDF and header image in the `Static` folder.
2. Run the script with the following command:
    ```sh
    python Script/view.py
    ```

3. Follow the prompts to complete the formatting tasks.

## Code Overview

### `view.py`

The `Get_New_PDF` function in `view.py` orchestrates the entire process. It:
- Creates a unique output folder.
- Removes headers and footers from the input PDF.
- Converts the PDF to DOCX format.
- Applies various formatting changes to the DOCX file.
- Adds headers, footers, page borders, and other elements.
- Updates the table of contents.
- Converts the modified DOCX back to PDF.
- Adds a custom page and watermark to the final PDF.

### `helper.py`

Contains helper functions used by `view.py` to perform specific tasks such as:
- `remove_header_footerssss`
- `convert_docx_to_pdf_windows`
- `remove_header_footer`
- `Format_doc`
- `start_each_heading_from_new_page`
- `start_each_heading_from_new_line`
- `remove_content_above_first_heading`
- `is_heading`
- `prompt_for_headings_to_remove`
- `remove_headings_with_content`
- `set_page_size_to_a4`
- `remove_empty_and_excessive_spaces`
- `start_each_heading1_from_new_page`
- `replace_heading_numbering`
- `add_header_with_image_size`
- `add_page_border`
- `create_index_of_heading`
- `add_footer_with_page_number`
- `update_toc_with_win32`
- `convert_docx_to_pdf`
- `add_custom_page_at_start`
- `add_watermark_to_pdf`
