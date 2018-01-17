import xlrd
import xlwt
import pymssql
import os

path=r'D:\TFS\Documentation\Projects\iWorkflow\030.需求开发\iWorkflow iClaims\UserInformation'
filename='CMA&ASG OUTSCOPE PROPERTIES.xlsx'
fullpath=os.path.join(path,filename)

#Read outsocpe company data from excel
outSocpeData=xlrd.open_workbook(fullpath)
table=outSocpeData.sheet_by_index(0)
nrows=table.nrows
outScopeComID=[]
for i in range(1,nrows):
    id=table.cell(i,3).value
    outScopeComID.append(id)

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

sqlCom="select * from igrowemployee where isActive=1 and CHRISID is not null and CostCenterCode=''"
cursor.execute(sqlCom)
dbRows=cursor.fetchall()
resultMsg='All user have CostCenter.'
missingData=[]
for row in dbRows:
    #Chris_Company_Code at index 13
    comId=str(row['Chris_Company_Code']).ljust(3)
    if(comId not in outScopeComID):
        missingData.append(row)

#close
cursor.close()
con.close()

#write missing data in excel
f=xlwt.Workbook()
sheet1=f.add_sheet(u'MissingCostCenter',cell_overwrite_ok=True)
#add column name
for columnName in missingData:
    i=0
    for item in columnName:
        sheet1.write(0,i,item)
        i+=1

rowIndex=1
for missingRow in missingData:
    columnIndex=0
    for column,item in missingRow.items():
        sheet1.write(rowIndex,columnIndex,item)
        columnIndex+=1
    rowIndex+=1
f.save('D:\\MissingData.xls')

print("Missing rows are :{0}".format(missingData.__len__()))