# -*- coding: utf-8 -*-
"""
Created on Mon Aug 19 15:21:23 2019

@author: aniksinh
"""

import xml.etree.ElementTree as ET
import xlrd

wb = xlrd.open_workbook(r'C:\Users\aniksinh\Desktop\xml\xml.xlsx') 
sheet = wb.sheet_by_index(0)

tree = ET.parse(r'C:\Users\aniksinh\Desktop\xml\MRO57.xml')
root = tree.getroot()

data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(1,sheet.nrows)]
print(sheet.nrows)
for i in range(sheet.nrows-1):
    qmnum=int(data[i][0])
    tdline=data[i][1]
    leveranspunkt=data[i][2]
    serialno=data[i][3]
    tag=['QMNUM','TDLINE','LEVERANSPUNKT','SERIALNO']
    
    for form in root.iter(tag[0]):
        form.text = str(qmnum)
    
    for form in root.iter(tag[1]):
        form.text = str(tdline)
    
    for form in root.iter(tag[2]):
        form.text = str(leveranspunkt)
    
    for form in root.iter(tag[3]):
        form.text = str(serialno)
    
    tree.write(r'C:\Users\aniksinh\Desktop\xml\new\MRO57'+'('+str(i)+')'+'.xml', xml_declaration=True, method="xml", encoding="UTF-8")