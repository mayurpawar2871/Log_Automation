import xlrd
import xlwt
import os

from xlwt import Workbook 
from xlutils.copy import copy
from xlrd import open_workbook


#loc = ("/home/sim/Documents/SHM/Testing/Scan_report/Input/test.xlsx") 

rb = xlrd.open_workbook('Main_Report.xlsx')
wb = copy(rb)
sheet = wb.get_sheet(0)
sheet.write(2, 16, 'asas')
wb.save('Main_Report.xlsx')

