# -*- coding: utf-8 -*-
"""
Created on Tue May 14 20:26:33 2019

@author: Cpt.SriRaja
"""

import  xlsxwriter
from pyexcel_xls import save_data
data = {"sheet 1 " :[ [1,2,3],[2,3,4]]}
save_data("E:\\data12.xls", data)

workbook =  xlsxwriter.Workbook("E:\\data13.xlsx")
worksheet = workbook.add_worksheet()
worksheet.write("A1" , "hai")

workbook.close()