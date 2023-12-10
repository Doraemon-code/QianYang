# -*- coding: utf-8 -*-
"""
Created on Mon Dec  4 11:08:12 2023

@author: QianYang
"""

from CheckListProce import CheckListProce
checkList = CheckListProce()
import re

class CheckListCreator():
    def __init__(self):
        pass
    
def Creator(self):
    CRF_List = ["aCRF File", "blank CRF File"]
    CRF_File_Path = checkList.getFileLoc(*CRF_List)
          
    #get all the page dictionary
    page_List = []
    for path in CRF_File_Path:
        dic = checkList.CheckListCreator(path)
        page_List.append(dic)
    
    export_path=CRF_File_Path[0].replace(re.search("AnnotatedUniqueCRF(.*?).pdf",CRF_File_Path[0]).group().strip("\""),"")
    
    if len(page_List) == 2:
        i = 0
        checkList.OutFile(export_path, "CheckList", page_List[i], CRF_List[i],page_List[i+1], CRF_List[i+1])
