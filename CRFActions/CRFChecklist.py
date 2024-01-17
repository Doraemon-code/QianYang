# -*- coding: utf-8 -*-
"""
Created on Fri Dec 15 14:51:27 2023

@author: QianYang
"""

import re
import CRFChecklistFunc as CRFChecklistFunc              
CRF_List = ["aCRF File", "blank CRF File"]
CRF_File_Path = CRFChecklistFunc.getFileLoc(*CRF_List)
page_List = []
for path in CRF_File_Path:
    dic = CRFChecklistFunc.getFormList(path)
    page_List.append(dic)
export_path=CRF_File_Path[0].replace(re.search(r'AnnotatedCRF_(.*?).pdf',CRF_File_Path[0]).group().strip("\""),"")                                 
if len(page_List) == 2:
    CRFChecklistFunc.OutFile(export_path, "CheckList", page_List[0], CRF_List[0],page_List[1], CRF_List[1])