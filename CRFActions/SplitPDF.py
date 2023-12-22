import os
import PyPDF2

def split_pdf(input_path, output_path, ranges, custom_names=None):
    with open(input_path, 'rb') as infile:
        reader = PyPDF2.PdfReader(infile)
        for i, page_range in enumerate(ranges):
            output = PyPDF2.PdfWriter()
            pages = []
            if '-' in page_range:
                start, end = map(int, page_range.split('-'))
                for page_num in range(start-1, end):
                    pages.append(reader.pages[page_num])
            else:
                page_num = int(page_range) - 1
                pages.append(reader.pages[page_num])

            for page in pages:
                output.add_page(page)

            output_filename = custom_names[i] if custom_names and i < len(custom_names) else f"output_{i + 1}.pdf"
            with open(os.path.join(output_path, output_filename), 'wb') as outfile:
                output.write(outfile)

#下面的参数为文件名的共同部分，通常不同更改
blankCRF = "CRF UAT Completion Checklist"
anotatedCRF = "Annotated CRF UAT Completion Checklist"
editcheck = "Edit Check UAT Completion Checklist"

#源PDF路径与拆分文件的输出路径，对于同一个项目通常也不需要修改
input_path  = "C:\\Users\\WeiqianYu\\Desktop\\"
output_path = "C:\\Users\\WeiqianYu\\Desktop\\"

filename = "收据_2023-12-15_171514.pdf"
ranges = ["1", "2", "3"]
study = "ONC-392-001"
version_info = "V1.2_20231222"

BlankCRF = study + " " + blankCRF + "_" + version_info
AnotatedCRF = study + " " +  anotatedCRF + "_" + version_info
EditCheck = study + " " +  editcheck + "_" + version_info
input_path = os.path.join(input_path,filename)

custom_names = [BlankCRF, AnotatedCRF, EditCheck]

split_pdf(input_path, output_path, ranges, custom_names)
