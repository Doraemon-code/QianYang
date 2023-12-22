# -*- coding: utf-8 -*-
"""
Created on Mon Dec  4 11:08:12 2023

@author: QianYang
"""

import CheckListProce 
import re

class CRFCheckListCreator():
    def __init__(self):
        pass

    @staticmethod
    def Creator():
        CRF_List = ["aCRF File", "blank CRF File"]
        CRF_File_Path = CheckListProce.getFileLoc(*CRF_List)
        CRF_Dic = {}
        for i in range(len(CRF_List)):
            CRF_Dic[CRF_List[i]] = CRF_File_Path[i]              
        page_List = []
        for key, value in CRF_Dic.items():
            dic = CheckListProce.CheckListCreator(str(key), str(value))
            page_List.append(dic)
        export_path=CRF_File_Path[0].replace(re.search(r'AnnotatedCRF_(.*?).pdf',CRF_File_Path[0]).group().strip("\""),"")                                 
        if len(page_List) == 2:
            i = 0
            CheckListProce.OutFile(export_path, "CheckList", page_List[i], CRF_List[i],page_List[i+1], CRF_List[i+1])
