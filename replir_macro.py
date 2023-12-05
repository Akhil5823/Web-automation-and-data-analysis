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

#replir file downloading
try:
    conn_site.replir()
    pass
except:
    print("can't download latest replir file.")
folder_path = r'C:\Users\ebinakh\Downloads'
file_type = r'\*xlsm'
files = glob.glob(folder_path + file_type)
max_file = max(files, key=os.path.getctime)
print("replir file downloaded: "+max_file)


# #copying data to the macro applying file
# workbook1 = Workbook(r"Kaarunya Files Final 10 Feb\test_replir.xlsm") #destination
# workbook2 = Workbook(max_file) #source
# workbook1.getWorksheets().get("Header").copy(workbook2.getWorksheets().get(0))
# workbook1.getWorksheets().get("Data").copy(workbook2.getWorksheets().get(1))
# workbook1.save(r"C:\Users\ebinakh\OneDrive - Ericsson\Documents\Ak_ericsson\tab_to_pow\Kaarunya Files Final 10 Feb\test_replir.xlsm")



# # #allocation macro on replir file
# wb=  xw.Book(r"C:\Users\ebinakh\OneDrive - Ericsson\Documents\Ak_ericsson\tab_to_pow\Kaarunya Files Final 10 Feb\test_replir.xlsm")
# for sheet in wb.sheets:
#     if 'Evaluation Warning' in sheet.name:
#         sheet.delete()
#         # print("sheet deleted")
# wb.activate("Data")
# macro1=wb.macro("Module2.Allocation_macro_LATEST_testing")
# macro1()

try:
    print('3')
    wb1=  xw.Book(max_file)
    print("kk")
    for sheet in wb1.sheets:
        if 'Evaluation Warning' in sheet.name:
            sheet.delete()
            print("sheet deleted")
    # wb1.activate("general_report")
    wb2=xw.Book(r'Kaarunya Files Final 10 Feb\UATRJ-All-In-One-FEB_1.xlsm')
    print("applying")
    macro1=wb2.macro("Module2.Allocation_macro_LATEST_testing")
    wb1.activate("Data")
    macro1()
    wb1.close()
    wb2.close()
    print("allocation macro applied on replir file !!")
    # print("utilisation macro applied successfully !!")
except Exception as e:
    print(f"An error occurred: {e}")


time.sleep(5)

# print("allocation macro applied on replir file !!")