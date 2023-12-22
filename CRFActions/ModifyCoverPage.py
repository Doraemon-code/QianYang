import re
import PyPDF2

# 打开PDF文件
with open('your_file.pdf', 'rb') as file:
    reader = PyPDF2.PdfFileReader(file)
    
    # 选择你想要的页面（例如第一页）
    page = reader.getPage(0)
    
    # 从该页中提取文本
    text = page.extractText()
    
    # 使用正则表达式匹配字符串
    pattern = r'V\d+\.\d+, \d{2}[A-Za-z]{3}\d{4}'
    matches = re.findall(pattern, text)
    
    # 输出匹配的字符串
    for match in matches:
        print(match)

    # 将匹配的字符串替换为自定义的字符串
    replacement = 'your_replacement_string'
    new_text = re.sub(pattern, replacement, text)

# 将新的文本写入新的PDF文件
with open('new_file.pdf', 'wb') as file:
    writer = PyPDF2.PdfFileWriter()
    
    # 创建一个新的页面并将新的文本添加到该页
    page = PyPDF2.pdf.PageObject.createBlankPage(None, 210, 297)  # 创建一个空白页
    page.extractText(new_text)  # 将新的文本添加到该页
    writer.addPage(page)  # 将该页添加到PDF文件
    
    writer.write(file)
