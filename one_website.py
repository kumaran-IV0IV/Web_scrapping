# -*- coding: utf-8 -*-
"""
Created on Wed May 15 17:10:41 2019

@author: Cpt.SriRaja
"""



import xlsxwriter
from bs4 import BeautifulSoup
import requests
page = requests.get("https://www.python.org/")

soup = BeautifulSoup(page.content, "html.parser")

for script in soup(["script"],["style"]):
    script.extract()
    
text = soup.get_text()



#for line in text.splitlines():
lst1 = text.strip()
lst2 = lst1.split(" ") 
kewrd_1 = lst2.count("the")
kewrd_2 =lst2.count("is")
kewrd_3 = lst2.count("to")

workbook =  xlsxwriter.Workbook("E:\\data14.xlsx")
worksheet = workbook.add_worksheet()

row = 0
col = 0

D1 = {"the": kewrd_1, "is": kewrd_2, "to": kewrd_3 }

  
# Iterate over the data and write it out row by row. 
for key, value in D1.items(): 
    worksheet.write(row, col, key) 
    worksheet.write(row, col + 1, value) 
    row += 1
    
chart = workbook.add_chart({'type':'pie'})

chart.add_series({
    'categories': '=Sheet1!$A$1:$A$3',
    'values':     '=Sheet1!$B$1:$B$3'
    
})

worksheet.insert_chart('C4' , chart)
  
workbook.close() 
   