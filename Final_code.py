# -*- coding: utf-8 -*-
"""
Created on Mon May 13 21:11:56 2019

@author: Cpt.SriRaja
"""
import xlsxwriter
from bs4 import BeautifulSoup
import requests

workbook =  xlsxwriter.Workbook("E:\\scrapped_data.xlsx")
worksheet1 = workbook.add_worksheet()
worksheet2 = workbook.add_worksheet()
worksheet3 = workbook.add_worksheet()

D1 = {}
class scrap:
    def __init__(self,url):
        self.page = requests.get(url)
        
        soup = BeautifulSoup(self.page.content, "html.parser")
        
        for script in soup(["script"],["style"]):
            script.extract()
            
        self.text = soup.get_text()
        
        
    def data(self):
        #for line in text.splitlines():
        lst1 = self.text.strip()
        lst2 = lst1.split(" ") 
        kewrd_1 = lst2.count("the")
        kewrd_2 =lst2.count("is")
        kewrd_3 = lst2.count("to")
        
        D1 = {"the": kewrd_1, "is": kewrd_2, "to": kewrd_3 }
        
        return D1
        
		
url1 = scrap("https://www.python.org/")
url2 = scrap("https://www.random.org/")
url3 = scrap("https://en.wikipedia.org/wiki/Wikipedia")


#*************************Worksheet_1*******************************#

row = 0
col = 0

data1 = url1.data()
print(data1)
  
# Iterate over the data and write it out row by row. 
for key, value in data1.items(): 
    worksheet1.write(row, col, key) 
    worksheet1.write(row, col + 1, value) 
    row += 1
    
chart = workbook.add_chart({'type':'pie'})

chart.add_series({
    'categories': '=Sheet1!$A$1:$A$3',
    'values':     '=Sheet1!$B$1:$B$3'
    
})

worksheet1.insert_chart('C4' , chart)
  
#*************************Worksheet_2*******************************#

row = 0
col = 0


data2 = url2.data()
print(data2)  
# Iterate over the data and write it out row by row. 
for key, value in data2.items(): 
    worksheet2.write(row, col, key) 
    worksheet2.write(row, col + 1, value) 
    row += 1
    
chart = workbook.add_chart({'type':'pie'})

chart.add_series({
    'categories': '=Sheet2!$A$1:$A$3',
    'values':     '=Sheet2!$B$1:$B$3'
    
})

worksheet2.insert_chart('C4' , chart)
  
#*************************Worksheet_3*******************************#

data3 = url3.data()
print(data3)
row = 0
col = 0

# Iterate over the data and write it out row by row. 
for key, value in data3.items(): 
    worksheet3.write(row, col, key) 
    worksheet3.write(row, col + 1, value) 
    row += 1
    
chart = workbook.add_chart({'type':'pie'})

chart.add_series({
    'categories': '=Sheet2!$A$1:$A$3',
    'values':     '=Sheet2!$B$1:$B$3'
    
})

worksheet3.insert_chart('C4' , chart)
  
workbook.close()


       