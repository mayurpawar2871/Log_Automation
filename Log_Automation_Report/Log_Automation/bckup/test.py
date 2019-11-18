#https://www.programcreek.com/python/example/6932/xlwt.Workbook
#https://translate.googleusercontent.com/translate_c?depth=1&hl=en&prev=search&rurl=translate.google.com&sl=ja&sp=nmt4&u=https://stackoverflow.com/questions/2719884/pivots-using-pyexcelerator-xlrd&xid=17259,15700022,15700186,15700190,15700256,15700259,15700262,15700265,15700271,15700280,15700283&usg=ALkJrhiZXBogexZN8Mp9yeechhvOaVibaA
#https://hackernoon.com/working-with-spreadsheets-using-python-part-1-380a120387f

import glob
import xlwt
import xlrd
import os
import xlutils


from xlwt import Workbook 
from xlutils.copy import copy

#overwrite xls sheet
rb = xlrd.open_workbook('Main_Report.xlsx')   #new add  
wb = copy(rb)  #new add  
sheet = wb.get_sheet(0)  #new add  
sheet2 = rb.sheet_by_index(0)


for i in range(4,sheet2.nrows):
    keyword_counter=sheet2.cell_value(i, 4)    
    print(" keyword_counter=", keyword_counter)
