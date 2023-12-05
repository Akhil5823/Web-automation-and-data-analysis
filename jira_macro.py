import xlwings as xw
import glob
import os.path
# import  jpype
from zipfile import ZipFile   
# import  asposecells
# jpype.startJVM() 
import win32com.client as win32
# from asposecells.api import Workbook
import conn_site
import time
import os

file1 = open(r"Kaarunya Files Final 10 Feb\CONFIG_2.txt","r")
dow=file1.readline()


# # .....................................................................................................................................
#calling jira server

conn_site.jira()

#finds the recently downloaded jira zipped folder
folder_path = r'C:\Users\ebinakh\Downloads'
file_type = r'\*zip'
files = glob.glob(folder_path + file_type)
max_file = max(files, key=os.path.getctime)
print(max_file) 

#extract the zipped folder
# loading the temp.zip and creating a zip object
# with ZipFile(max_file,'r') as zObject:
#     # Extracting all the members of the zip 
#     # into a specific location.
#     zObject.extractall(
#         path=r"C:\Users\ebinakh\Downloads")

try:
    # conn_site.jira()
    print('jira file downloaded')
    pass
except:
    print("can't get latest jira file.")
time.sleep(5)
# print('jira file downloaded')

# finds the recently downloaded file

folder_path = "\r"+dow
print(folder_path)
file_type = r'\*xls'
print(file_type)
files = glob.glob(folder_path + file_type)

max_file = max(files, key=os.path.getctime)
print(max_file)
print("Please wait,converting to xlsx........")

#converts it into xlsx format
fname = max_file
excel = win32.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Open(fname)
wb.SaveAs(fname+"x", FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
wb.Close()                               #FileFormat = 56 is for .xls extension
excel.Application.Quit()
print("xlsx file created")


#finds the converted file
folder_path = dow
file_type = r'\*xlsx'
files = glob.glob(folder_path + file_type)
max_file = max(files, key=os.path.getctime)
print(max_file)

# max_file=r'Kaarunya Files Final 10 Feb\Hub 1 GAIA proj (eTeamProject) 2023-07-20T13_01_39+0200.xlsx'


# #copy the data from Output1.xlsx to the worksheet in which we are applying the macro
# workbook1 = Workbook(r"C:\Users\ebinakh\OneDrive - Ericsson\Documents\Ak_ericsson\tab_to_pow\Kaarunya Files Final 10 Feb\test_jira_pull2.xlsm") #destination
# workbook2 = Workbook(max_file) #source
# workbook1.getWorksheets().get("general_report").copy(workbook2.getWorksheets().get(0))
# workbook1.save(r"C:\Users\ebinakh\OneDrive - Ericsson\Documents\Ak_ericsson\tab_to_pow\Kaarunya Files Final 10 Feb\test_jira_pull2.xlsm")


import shutil
print("stareted copying")
original = max_file
target = r"Kaarunya Files Final 10 Feb\test_jira_pull2.xlsx"
shutil.copyfile(original, target)
time.sleep(5)
print("copied to dest")



#Jira macro applied on jira file
try:
    print('3')
    wb1=  xw.Book(r"Kaarunya Files Final 10 Feb\test_jira_pull2.xlsx")
    print("kk")
    for sheet in wb1.sheets:
        if sheet.name!="general_report":
            sheet.delete()
            print("sheet deleted")
    # wb1.activate("general_report")
    wb2=xw.Book(r'Kaarunya Files Final 10 Feb\UATRJ-All-In-One-FEB_1.xlsm')
    print("applying")
    macro1=wb2.macro("Module11.JIRA_GAIA")
    wb1.activate("general_report")
    macro1()
    # workbook1 = Workbook(r"C:\Users\ebinakh\OneDrive - Ericsson\Documents\Ak_ericsson\tab_to_pow\Kaarunya Files Final 10 Feb\JIRA_data.xlsx")  
    # workbook1.save()
    wb1.close()
    wb2.close()

    print("Jira macro applied successfully !!")
except Exception as e:
    print(f"An error occurred: {e}")
