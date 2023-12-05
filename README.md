# Web-automation-and-data-analysis
Automates the proccess of applying macros on excel sheets and automates the proccess of fetching data from different servers.
Automating the process of applying macros.

One time setup:
1)	Go to Kaarunya Files Final 10 Feb and open UATRJ-All-In-One-FEB_1.xlsm. 
2)	Open vba – go to module 4 – change ‘file location’ to absolute path of this excel file.
3)	In the ‘Kaarunya Files Final 10 Feb’ folder open CONFIG_2.txt  and add the absolute path of your downloads folder in the first line.
4)	In the CONFIG_1.txt add your signam id in the first line and password in the next line.
5)	Install the following webdrivers and put it inside the main folder.(It is used for the web automation part.)
a.	Link for firefox webdrive: https://sourceforge.net/projects/geckodriver.mirror/
b.	Link for edge webdriver: https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/ 
(download the latest stable version)
*extract before putting it in the folder

Overview:
1)	There are 3 folders for each data sources : jira_macro , replir_macro , export_macro.
2)	Each of these folders have a .exe application which upon running will download the required files-apply the macros and give the output file in case of jira and replir.
3)	In the case of export, obtain the latest export file, then go to the ‘Kaarunya Files Final 10 Feb’ folder inside the ‘export_macro’ folder and replace the export1.xlsx with the new export file and rename it as export1.xlsx
*Note: make sure Data sheet comes before pivot table in export1.xlsx
Usage:
1)	The process can be started by running the .exe files in the respective folders. (export_macro.exe, jira_macro.exe, replir_macro.exe) A terminal screen will appear showing the progress.
2)	Incase there is any problem downloading the replier or jira , just download it manually to downloads folder.
3)	The following excel files are produced after running the exe and can be connected to power bi.
a.	Allocation_data.xlsx
b.	 amonth.xlsx,ytd.xlsx
c.	 JIRA_Data.xlsx
d.	SA_designation_data.xlsx
e.	 Utilisation_data.xlsx
f.	LGO.xlsx


