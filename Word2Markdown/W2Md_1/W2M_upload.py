import os
import re
import docx2txt
from docx import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph
from pathlib import Path
from tkinter import Tk, filedialog

def extract_images(docx_path, images_dir):
    # 提取文档中的图片
    text = docx2txt.process(docx_path, images_dir)
    image_files = os.listdir(images_dir)
    image_files.sort()
    return image_files

def convert_table_to_markdown(table):
    rows = table.rows
    table_md = []
    for i, row in enumerate(rows):
        cells = row.cells
        row_md = "| " + " | ".join(cell.text.strip() for cell in cells) + " |"
        table_md.append(row_md)
        if i == 0:  # Add the header separator
            separator = "| " + " | ".join("---" for _ in cells) + " |"
            table_md.append(separator)
    return "\n".join(table_md)

def convert_docx_to_markdown(file_path):
    # Read the document
    doc = Document(file_path)

    # Extract the file name without extension
    file_name = os.path.basename(file_path)
    file_base = os.path.splitext(file_name)[0]

    # Create a directory for the markdown file and images
    output_dir = f"./generate_data/{file_base}"
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Path for the markdown file
    md_file_path = os.path.join(output_dir, "output.md")

    # Initialize markdown content
    md_content = []

    # Create a directory for images
    images_dir = Path(output_dir) / "Images"
    images_dir.mkdir(parents=True, exist_ok=True)

    # Extract images
    image_files = extract_images(file_path, images_dir)

    # Function to handle different heading levels
    def handle_heading(text, level):
        if level > 6:
            level = 6
        return f"{'#' * level} {text}"

    # Function to check if the paragraph contains an image
    def contains_image(paragraph):
        for run in paragraph.runs:
            if "graphicData" in run._element.xml:
                return True
        return False

    image_index = 0
    content_started = False
    first_secondary_found = False
    primary_title = []

    for element in doc.element.body:
        if isinstance(element, CT_Tbl):
            table = Table(element, doc)
            table_md = convert_table_to_markdown(table)
            md_content.append(table_md)
        elif isinstance(element, CT_P):
            para = Paragraph(element, doc)
            text = para.text.strip()

            # Skip empty paragraphs
            if not text and not contains_image(para):
                continue

            # Skip the first page content
            if not content_started:
                if re.match(r'^\d+\s*[^.\d].*$', text):
                    content_started = True
                else:
                    continue

            # Filter out titles with date format or containing "-"
            if re.match(r'^\d{4}-\d{2}-\d{2}', text) or '-' in text:
                continue

            # If first secondary title found, consider previous lines as primary title
            if re.match(r'^\d+\s*[^.\d].*$', text):
                if not first_secondary_found:
                    first_secondary_found = True
                    primary_title_text = ' '.join(primary_title).strip()
                    # 只保留空格后面的文字
                    if primary_title_text and ' ' in primary_title_text:
                        primary_title_text = primary_title_text.split(' ', 1)[1].strip()
                    if primary_title_text:
                        print(f"Primary title matched: {primary_title_text}")  # 调试输出
                        md_content.append(handle_heading(primary_title_text, 1))
                level = 2
                md_content.append(handle_heading(text, level))
            elif re.match(r'^\d+\.\d+[^.]*$', text):
                match = re.match(r'^(\d+\.\d+)(.*)$', text)
                if match:
                    level = 3
                    md_content.append(handle_heading(match.group(1), level))
                    remaining_text = match.group(2).strip()
                    if remaining_text:
                        md_content.append(remaining_text)
            elif re.match(r'^\d+\.\d+\.\d+[^.]*$', text):
                match = re.match(r'^(\d+\.\d+\.\d+)(.*)$', text)
                if match:
                    level = 4
                    md_content.append(handle_heading(match.group(1), level))
                    remaining_text = match.group(2).strip()
                    if remaining_text:
                        md_content.append(remaining_text)
            elif re.match(r'^\d+\.\d+\.\d+\.\d+[^.]*$', text):
                match = re.match(r'^(\d+\.\d+\.\d+\.\d+)(.*)$', text)
                if match:
                    level = 5
                    md_content.append(handle_heading(match.group(1), level))
                    remaining_text = match.group(2).strip()
                    if remaining_text:
                        md_content.append(remaining_text)
            elif re.match(r'^\d+\.\d+\.\d+\.\d+\.\d+[^.]*$', text):
                match = re.match(r'^(\d+\.\d+\.\d+\.\d+\.\d+)(.*)$', text)
                if match:
                    level = 6
                    md_content.append(handle_heading(match.group(1), level))
                    remaining_text = match.group(2).strip()
                    if remaining_text:
                        md_content.append(remaining_text)
            elif re.match(r'^\d+\.\d+\.\d+\.\d+\.\d+\.\d+[^.]*$', text):
                match = re.match(r'^(\d+\.\d+\.\d+\.\d+\.\d+\.\d+)(.*)$', text)
                if match:
                    level = 7
                    md_content.append(handle_heading(match.group(1), level))
                    remaining_text = match.group(2).strip()
                    if remaining_text:
                        md_content.append(remaining_text)
            else:
                if not first_secondary_found:
                    primary_title.append(text)
                else:
                    md_content.append(text)

            # Insert image if the paragraph contains an image
            if contains_image(para) and image_index < len(image_files):
                image_filename = image_files[image_index]
                md_content.append(f"![image_{image_index + 1}](./Images/{image_filename})")
                image_index += 1

    # Add primary title if not already added
    if primary_title and not first_secondary_found:
        primary_title_text = ' '.join(primary_title).strip()
        # 只保留空格后面的文字
        if primary_title_text and ' ' in primary_title_text:
            primary_title_text = primary_title_text.split(' ', 1)[1].strip()
        if primary_title_text:
            print(f"Primary title matched: {primary_title_text}")  # 调试输出
            md_content.append(handle_heading(primary_title_text, 1))
    elif not primary_title:
        print("No primary title matched")  # 调试输出

    # Write markdown content to file
    with open(md_file_path, 'w', encoding='utf-8') as md_file:
        md_file.write('\n\n'.join(md_content))

def select_files():
    root = Tk()
    root.withdraw()  # Hide the root window
    file_paths = filedialog.askopenfilenames(
        title="Select DOCX files",
        filetypes=(("DOCX files", "*.docx"), ("All files", "*.*"))
    )
    return file_paths

if __name__ == "__main__":
    file_paths = select_files()
    if file_paths:
        for file_path in file_paths:
            convert_docx_to_markdown(file_path)
    else:
        print("No files selected.")
