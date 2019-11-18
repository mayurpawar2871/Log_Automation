import glob
import xlwt
import xlrd
import os



from xlwt import Workbook 
from xlutils.copy import copy
from xlrd import open_workbook

#overwrite xls sheet
rb = xlrd.open_workbook('Main_Report.xlsx')   #new add  
wb = copy(rb)  #new add  
sheet = wb.get_sheet(0)  #new add  
sheet2 = rb.sheet_by_index(0)

#loc = ("/home/sim/Documents/SHM/Testing/Scan_report/Input/Main_Report.xlsx") 
#wb = xlrd.open_workbook(loc)       #new add  
#sheet = wb.sheet_by_index(0)       #new add  


keyword_counter=0   #new add  
count=35   #initialize to to start printing report from below
pass_count=0
i=4
for i in range(sheet2.nrows):     #new
    keyword_counter=sheet2.cell_value(i, 4)    #new
    print(keyword_counter)

