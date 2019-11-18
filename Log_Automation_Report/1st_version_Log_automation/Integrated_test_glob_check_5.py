#DEV BY Mayur
#working with 2019_09_05 & take character as input 
#string count functionality added
# counts no of times occured in each log n print in report
# also prints file name & path for each log
# prints TOTAL COUNT 

#LIMITATION - prints only individual report, MAIN report summzry manual work entry needed from individial report 


import glob
import xlwt
import xlrd
import os

#to store all filename .txt 
#DECLARE EMPTY DICTINARY
files = {}

#create xls sheet
from xlwt import Workbook 
workbook = Workbook()
sheet = workbook.add_sheet('Sheet_1', cell_overwrite_ok=True)

count=-1

#enterstring
variable = raw_input('ENTER INPUT:')

sheet.read(0, 11,'TOTAL FOUND')
print("INPUT PRINT=",filename)


#Wildcard glob operator  will match LOGS those have .txt extension.
fileinfo=glob.glob("/home/sim/Documents/SHM/Testing/Scan_report/logs_25_27/*.txt")
#fileinfo=os.path.basename("/home/mp/PROJ_HW/P18_UBUNUT/SUB_PROJ1/Long_Run_test_459/2019_09_05/*.txt")
#print(fileinfo)

#open file xls 
for filename in fileinfo:
    with open(filename, "r") as file:
        if filename in files:
           continue
        files[filename] = file.read()

sheet.write(0, 0,'NUMBER OF LOGS')
sheet.write(0, 1,'TEST')
sheet.write(0, 2,'HOW MANY PASSED')
sheet.write(0, 3,'LOG PATH ')
sheet.write(0, 11,'TOTAL FOUND')

count=0  #initialize to1 to start printing report from 2 nd row onwards
pass_count=0

for filename, text in files.items():
    print("filename=",filename)
    #print("=" * 80)
    #print(text)
    
    count_times=text.count(variable)
    sub_index = text.find(variable)
    print("The position of 'contains' word: ", sub_index)

    if sub_index == -1:
     print("FAIL")
     count += 1
     sheet.write(count, 0,count)
     sheet.write(count, 1,'NOT FOUND')
     sheet.write(count, 2, count_times)
     sheet.write(count, 3, filename)
     workbook.save('Report.xls')
     #save in workbook
    else:
     print("PASS")
     count += 1
     pass_count += 1
     sheet.write(count, 0,count)
     sheet.write(count,1,'FOUND')
     sheet.write(count, 2, count_times)
     sheet.write(count, 3, filename)
     sheet.write(0, 12,pass_count)
     #save in workbook
     #row.write(0,'PASS')
     workbook.save('Report.xls')

    print("COUNT=",count)
    if pass_count == 0:
     sheet.write(0, 12,pass_count)
     workbook.save('Report.xls')
        

  
