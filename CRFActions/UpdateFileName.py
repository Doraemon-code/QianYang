import os
import re

# 文件夹路径
folder_path = "C:\\Users\\WeiqianYu\\Desktop\\006\\CRF\\1.1"

# 版本号和日期
version_number = "V1.1"
date = "08Dec2023"

# 获取文件列表
file_list = os.listdir(folder_path)

# 遍历文件
for filename in file_list:
    # 检查文件是否为pdf格式
    if filename.endswith(".pdf") or filename.endswith(".xlsx"):
        # 使用正则表达式匹配文件名
        match = re.match(r'^(.*?)_(.*?)_(\d+\.\d+\.\d+\.\d+)_(\d+)\.(pdf|xlsx)$', filename)
        if match:
            # 从匹配中获取不同部分的信息
            project_name = match.group(1)
            middle_part = match.group(2)
            version_info = match.group(4)
            
            # 根据要求修改中间部分
            if middle_part == "AnnotatedCRF":
                middle_part = "Annotated CRF"
            elif middle_part == "BlankUniqueCRF":
                middle_part = "Blank Unique CRF"
            elif middle_part == "BlankCRF":
                middle_part = "Blank CRF"
            elif middle_part == "AnnotatedUniqueCRF":
                middle_part = "Annotated Unique CRF"
            elif middle_part == "DBSpec":
                middle_part = "Database Specification"
            elif middle_part == "EditCheck":
                middle_part = "Edit Check"    
            
            # 构建新的文件名
            new_filename = f"{project_name}_{middle_part}_{version_number}_{date}{os.path.splitext(filename)[-1]}"
            
            # 构建旧的文件路径和新的文件路径
            old_file_path = os.path.join(folder_path, filename)
            new_file_path = os.path.join(folder_path, new_filename)
            
            # 重命名文件
            os.rename(old_file_path, new_file_path)
            print(f"文件 {filename} 已重命名为 {new_filename}")
        else:
            print(f"文件 {filename} 不符合命名规范")
