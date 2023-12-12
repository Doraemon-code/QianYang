import os
from docx2pdf import convert
from PyPDF2 import PdfReader, PdfWriter

# 自定义目录路径
pdf_directory = "C:\\Users\\WeiqianYu\\Desktop\\006\\CRF\\1.1"
docx_directory = "C:\\Users\\WeiqianYu\\Desktop\\006"

# 定义查找包含特定字符串的PDF文件的函数
def find_matching_pdfs(directory, search_string):
    matching_files = []
    for filename in os.listdir(directory):
        if filename.endswith(".pdf"):
            if search_string in filename:
                matching_files.append(os.path.join(directory, filename))
    return matching_files

# 查找包含特定字符串的PDF文件
annotated_crf_pdfs = find_matching_pdfs(pdf_directory, "Annotated CRF")
blank_crf_pdfs = find_matching_pdfs(pdf_directory, "Blank CRF")

# 添加CoverPage作为PDF文件的首页
def add_cover_page(pdf_files, cover_page_path):
    for pdf_file in pdf_files:
        with open(pdf_file, "rb") as file:
            reader = PdfReader(file)
            cover_page = PdfReader(open(cover_page_path, "rb"))
            writer = PdfWriter()

            # 将CoverPage添加为PDF的第一页
            writer.append_pages_from_reader(cover_page)
            writer.append_pages_from_reader(reader)

            # 保存新的PDF文件
            with open(os.path.join(pdf_directory, f"{os.path.basename(pdf_file)}"), "wb") as new_file:
                writer.write(new_file)

# 判断docx路径下是否存在pdf格式的CoverPage，如果存在则添加CoverPage作为Annotated CRF和Blank CRF的首页。如果不存在则转换CoverPage的docx文档为PDF格式.
if os.path.exists(os.path.join(docx_directory, 'CoverPage.pdf')):   
    add_cover_page(annotated_crf_pdfs, os.path.join(docx_directory, "CoverPage.pdf"))
    add_cover_page(blank_crf_pdfs, os.path.join(docx_directory, "CoverPage.pdf"))
else:
    convert(os.path.join(docx_directory, "CoverPage.docx"), os.path.join(docx_directory, "CoverPage.pdf"))
    add_cover_page(annotated_crf_pdfs, os.path.join(docx_directory, "CoverPage.pdf"))
    add_cover_page(blank_crf_pdfs, os.path.join(docx_directory, "CoverPage.pdf"))