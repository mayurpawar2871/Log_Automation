#DEV BY Mayur
# lop keyword to all text n update in Main_Report.xls

#https://python-forum.io/Thread-Iterate-over-excel-sheets-using-xlrd
# This log automation is freezed now on 8 th NOV 2019
# WORKING WITH  Main_Report.xls

import glob
import xlwt
import xlrd
import os

#to store all filename .txt 
#DECLARE EMPTY DICTINARY
files = {}

from xlwt import Workbook 
from xlutils.copy import copy
from xlrd import open_workbook

#overwrite xls sheet
rb = xlrd.open_workbook('Main_Report.xlsx')   #new add  
wb = copy(rb)  #new add  
sheet = wb.get_sheet(0)  #new add  
sheet2 = rb.sheet_by_index(0)

count=-1

#enterstring
#variable = raw_input('ENTER INPUT:')
keyword_counter=0   #new add  


#Wildcard glob operator  will match LOGS those have .txt extension.
fileinfo=glob.glob("/home/sim/Documents/SHM/Testing/Log_Automation_Report/Log_Automation/Logs_long_run/*.txt")
#fileinfo=os.path.basename("/home/mp/PROJ_HW/P18_UBUNUT/SUB_PROJ1/Long_Run_test_459/2019_09_05/*.txt")
#print(fileinfo)

#open file xls 
for filename in fileinfo:
    with open(filename, "r") as file:
        if filename in files:
           continue
        files[filename] = file.read()

#sheet.write(22, 0,'NUMBER OF LOGS')
sheet.write(22, 1,'KEYWORD')
sheet.write(2, 5,'HOW MANY FAILED')
sheet.write(22, 3,'LOG PATH ')


count=23   #initialize to to start printing report from below
pass_count_inc=3

for i in range(4,sheet2.nrows):     #new
    keyword_counter=sheet2.cell_value(i, 4)    #new
    print(" keyword_counter=", keyword_counter)
    pass_count=0
    pass_count_inc +=1

    for filename, text in files.items():
        #print("filename=",filename)
        #count_times=text.count(keyword_counter)
        
        sub_index = text.find(keyword_counter)
        #print("The position of 'contains' word: ", sub_index)
        
        if sub_index == -1:
         print("FAIL")
     	 #count += 1
     	 #sheet.write(count, 0,count-35)
     	 #sheet.write(count, 1,'NOT FOUND')
     	 #sheet.write(count, 2, count_times)
     	 #sheet.write(count, 2, filename)
     	 #wb.save('Main_Report.xlsx')
     	 #save in workbook
        else:
     	  #print("PASS")
     	 count += 1
     	 pass_count += 1
     	 #sheet.write(count, 0,count-35)
         sheet.write(count,1,keyword_counter)
         sheet.write(pass_count_inc, 5,pass_count)
     	 #sheet.write(count,1,'FOUND')  #replaced with keyword_counter
     	 #sheet.write(count, 2, count_times)
     	 sheet.write(count, 3, filename)

     	 #save in workbook
     	 #row.write(0,'PASS')
     	 wb.save('Main_Report.xlsx')

        if pass_count == 0:
         sheet.write(pass_count_inc, 5,pass_count)
         wb.save('Main_Report.xlsx')
        

  
