import xlrd
import pymssql
import os
import csv

# List file
ExportFiles = 0
ExportData = []

# Add r before the path to avoid char '\'
path = r'D:\Documents\Projects\iWorkflow\Enhancement\iClaims\BatchData\ToChris\0116'

files = os.listdir(path)
for file in files:
    ExportFiles += 1
    exFileName = os.path.join(path, file)
    csv_reader = csv.reader(open(exFileName, encoding='utf-8'))
    for row in csv_reader:
        ExportData.append(row)

#print(ExportData)
print("Total Export files are: {0}".format(ExportFiles))
print("Total Export rows are: {0}".format(ExportData.__len__()))

'''
#BU Dictionary
CLBU={'CLDCN':'CLCC','ASGCN':'ASG','CLC':'CLC','CLCEC':'CLC','CLCNC':'CLC','CLCSC':'CLC','CLCSW':'CLC','CMACN':'CMA'}

#Read outsocpe company data from excel
outSocpeData=xlrd.open_workbook('D:\\TFS\\Documentation\\Projects\\iWorkflow\\030.需求开发\\iWorkflow iClaims\\UserInformation\\CMA&ASG OUTSCOPE PROPERTIES.xlsx')
table=outSocpeData.sheet_by_name(u'Sheet1')
nrows=table.nrows


#open sql server connection & cursor
#server="GAOLAN" #UAT
server="10.154.128.185" #Production
user="IWFAdmin"
#password="3edcVFR$"
password="password@123456"
#dbname="CapitalWorkflowCenter_UAT"
dbname="CapitaWorkflowCenter"
con=pymssql.connect(server,user,password,dbname)
cursor=con.cursor(as_dict=True)
'''
