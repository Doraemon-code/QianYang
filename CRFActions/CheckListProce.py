# -*- coding: utf-8 -*-
"""
Created on Tue Oct 31 09:22:48 2023

@author: QianYang
"""
#all import outside the function

import openpyxl
import pandas as pd
import pdfplumber
import openpyxl.styles as styles
from openpyxl.utils.dataframe import dataframe_to_rows
import tkinter as tk
from tkinter import filedialog
    
# 01 - get any file location that selected
def getFileLoc(*files):
    root = tk.Tk()
    root.attributes('-topmost', 1)
    root.withdraw()
    file_paths = []
    for file in files:
        file_path = filedialog.askopenfilename(title=file)
        file_paths.append(file_path)
    return tuple(file_paths)

def GetWorksheetName(workbook_path, worksheet_name):
    workbook = openpyxl.load_workbook(workbook_path)    
    worksheet = None
    for sheet in workbook.worksheets:
        if sheet.title == worksheet_name:
            worksheet = sheet
            break
    return worksheet

def getFormList(crf,Lable_begin,lable_formName,trigger):
    Form_List = []
    for page in crf.pages:
        if Lable_begin in page.extract_text(): # bigining after frid forms
            formName = getFormName(trigger, page.extract_text(), lable_formName)
            Form_List.append(formName)
    Form_List = [x for x in set(Form_List) if x is not None]
    Form_List = sorted(Form_List)
    return Form_List

def getFormName(trigger_input, page_text, lable):
    output = None
    text = page_text
    lines = text.splitlines()
    trigger = 0
    for line in lines:
        if trigger_input :
            if lable == line:
                trigger = 1
                continue
            if trigger == 1:
                trigger = 2
                continue
            if trigger == 2:
                output = line.strip()
                break
        else:
            if lable in line:
                key, value = line.split(lable)
                output = key.strip()
                break
    return output

def getPageRange(trigger_input, pages, Lable_begin,listForm, lable):
    trigger = False
    output = []
    for page in pages:
        if Lable_begin in page.extract_text():
            formName = getFormName(trigger_input, page.extract_text(), lable)
            # get the first page number, and then to the next page derrectly
            if listForm == formName and not trigger:
                firstPage = str(page.page_number)
                trigger = True
                continue
            # when first page got, entered by trigger, and once last page valued,will find the next first page for the following pages
            if listForm != formName and trigger:
                lastPage = str(page.page_number-1)
                trigger = False
                # once last page valued, the first page must valued
                if firstPage == lastPage:
                    output.append(firstPage) 
                else:
                    output.append(firstPage + " - " + lastPage)
    return output

def CheckListCreator(crfType, acrf_path):
    #"Protocol #" can be the unique icon to make the table pages and form pages apart,and "Site ID Subject ID" can be the trigger to get form name line
    Lable_begin = "Protocol #"
    lable_formName = "dataset ="
    lable_subj = "Site ID Subject ID"
    crf=pdfplumber.open(acrf_path)
    # initail a dic to collect everty form name and pages
    Page_Dic = {}
    # branch
    if "blank" in crfType:
        trigger = True
        Form_List = getFormList(crf,Lable_begin,lable_subj,trigger)
        for form in Form_List:
                pageRange =getPageRange(trigger, crf.pages, Lable_begin, form, lable_subj)
                Page_Dic[form] = ", ".join(pageRange)  
    else:
        trigger = False
        Form_List = getFormList(crf,Lable_begin,lable_formName,trigger)
        for form in Form_List:
                pageRange =getPageRange(trigger, crf.pages, Lable_begin, form, lable_formName)
                Page_Dic[form] = ", ".join(pageRange)  
    return Page_Dic

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
                   
