import os
import re
from docx import Document
from docx.oxml import parse_xml
from docx2pdf import convert
from pdf2docx import parse

# 检索PDF文件
path = r'C:\\Users\Administrator\Desktop'
pdf_file_path = os.path.join(path, 'CoverPage.pdf')
new_pdf_file_path = os.path.join(path, 'CoverPage1.pdf')

# 检查文件是否存在
if not os.path.isfile(pdf_file_path):
    print(f"在路径{path}下面没有CoverPage.pdf")
else:
    # 转换PDF到Word
    docx_file_path = os.path.join(path, 'CoverPage.doc')
    parse(pdf_file_path, docx_file_path)

    # 打开docx文件
    doc = Document(docx_file_path)

    # 匹配正则表达式
    regex = r'[A-Za-z]+\s?\d\.\d, \d{2}-?[A-Za-z]{3}-?\d{4}'
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            text = run.text
            print(text)
            if re.search(regex, text):
                for match in re.finditer(regex, text):
                    start, end = match.span()
                    original_text = text[start:end]
                    replaced_text = "Version 5.1, 20-Dec-2023"
                    run.text = text.replace(original_text, replaced_text)  # 直接替换文本
                    print(run.text)
                    run.bold = False  # 设置加粗
                    # 设置缩进
                    element = run._element
                    if element.getprevious() is None:
                        element.addprevious(parse_xml(r'<w:ind w:firstLine="720" />'))

    # 保存为新的docx文件
    new_docx_file_path = os.path.join(path, 'CoverPage1.docx')
    doc.save(new_docx_file_path)

    # 转换新的docx文件为PDF
    convert(new_docx_file_path, new_pdf_file_path)
