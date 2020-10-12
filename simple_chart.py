# -*- coding: utf-8 -*-
"""
Created on Wed May 15 16:28:54 2019

@author: Cpt.SriRaja
"""


import xlsxwriter

workbook = xlsxwriter.Workbook("E:\\pie_chart.xlsx")

worksheet = workbook.add_worksheet()
chart = workbook.add_chart({'type': 'pie'})

data = [
    ['Pass', 'Fail'],
    [90, 10],
]

worksheet.write_column('A1', data[0])
worksheet.write_column('B1', data[1])

chart.add_series({
    'categories': '=Sheet1!$A$1:$A$2',
    'values':     '=Sheet1!$B$1:$B$2'
    
})

worksheet.insert_chart('C3', chart)

workbook.close()