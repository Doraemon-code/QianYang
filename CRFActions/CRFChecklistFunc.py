# -*- coding: utf-8 -*-
"""
Created on Fri Dec 15 14:07:59 2023

@author: QianYang
"""

import openpyxl
import pandas as pd
import pdfplumber
import openpyxl.styles as styles
from openpyxl.utils.dataframe import dataframe_to_rows
import tkinter as tk
from tkinter import filedialog


def getFileLoc(*files):
    root = tk.Tk()
    root.attributes('-topmost', 1)
    root.withdraw()
    file_paths = []
    for file in files:
        file_path = filedialog.askopenfilename(title=file)
        file_paths.append(file_path)
    return tuple(file_paths)

def getFormName(page_text, lable_subj, lable_formName):
    output = None
    text = page_text
    lines = text.splitlines()
    trigger = 0
    for line in lines:
        if lable_subj in line:
            trigger = 1
            continue
        if trigger == 1:
            trigger = 2
            continue
        if trigger == 2:
            output = line.strip()
            break
    if output != None:
        if lable_formName in output:
            key, value = output.split(lable_formName)
            formName = key.strip()
        else:
            formName = output
    try:
        return formName
    except:
        return None
    
def simplifyNumList(num_list):
    ranges = []
    range_out = []
    start = num_list[0]
    end = num_list[0]
    for num in num_list[1:]:
        if num == end + 1:
            end = num
        else:
            ranges.append((start, end))
            start = num
            end = num
    ranges.append((start, end))
    for start,end in ranges:
        if start != end:
            range_out.append(str(start) + "-" + str(end))
        else:
            range_out.append(str(start))
    return range_out

def getFormList(path):
    crf = None
    Lable_begin = "Protocol #"
    lable_formName = "dataset ="
    lable_subj = "Site ID"
    
    with pdfplumber.open(path) as crf:
        Form_Dic = {}
        #从PDF取page
        for page in crf.pages: 
            if Lable_begin in page.extract_text(): 
            
                #满足条件就把PDF header 和 page 添加到字典
                formName = getFormName(page.extract_text(), lable_subj, lable_formName)
                Form_Dic.setdefault(formName, []).append(page.page_number)
                
        if Form_Dic:
            #对字典遍历，将页码变成页码范围
            for key,value in Form_Dic.items():
                Form_Dic[key] =",".join((simplifyNumList(value)))
                
        # #将字典的header和page整合起来，变成一个列表
        # Form_List = [f"{key}\n{value}" for key, value in Form_Dic.items()]
        return Form_Dic
    
def OutFile( folder_path, outName, *args):
    out_path = folder_path + outName + '.xlsx'
    with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
        workbook = writer.book
        sheets = []
        font = styles.Font(name='Arial', size=11)
        for i in range(len(args)//2):
            input_Dic = args[2*i]
            sheetName = args[2*i+1]
            dict_list = [f"{key}\n{value}" for key, value in input_Dic.items()]
            df = pd.DataFrame(dict_list, columns=['CRF page*'])
            if i == 0:
                df['Dataset' + "\n" + '(Question, OID, Name, Datatype, Code list, SAS Field)'], df['Codelist' + "\n" 
                    + '(Datatype, Label, Code, Referenced Variable)'], df['Variable Name' + "\n" 
                    + '(Label, Datatype, Length, Format)'], df['Result'],df['Initial'] = None, None, None, None, None
            else:
                df['Dataset' + "\n" + '(Question, OID, Name, Datatype, Code list, SAS Field)'], df['Codelist' + "\n" 
                    + '(Datatype, Label, Code, Referenced Variable)'], df['Variable Name' + "\n" 
                    + '(Label, Datatype, Length, Format)'], df['Spelling check'], df['Result'],df['Initial'] = None, None, None, None, None, None
            sheet = workbook.create_sheet(title=sheetName)
            sheets.append(sheet)
            for row in dataframe_to_rows(df, index=False, header=True):
                sheet.append(row)        
            for col in sheet.columns:
                for cell in col:
                    cell.font = font
        for sheet in sheets:
            workbook.active = workbook.index(sheet)
        workbook.close()