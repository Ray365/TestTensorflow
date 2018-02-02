import time
import pymssql
import xlwt

# sqlComTxt = "select ApplicantUserDisplayNameCN,ClaimUserDisplayNameCN,u.UserCode,ProcessName,ProjectName,SerialNumber,Title,Purpose,v.CreatedBy,* from vw_GetBizInstancesInfo v left join K2_SystemUser u on v.ClaimUserAccount=u.UserAccount   where SerialNumber in ('1001-FSS-PC-00069401','1038-LEAS-EC-00025941','1038-LEAS-PC-00069766','1038-LEAS-PC-00069767','2064-CASH-PC-00070875','3040-BDAM-PC-00071402')"

# Check duplication ADUserID in iGrowPeople
# sqlComTxt="select * from IGrowEmployee where ADUserID in (select ADUserID from IGrowEmployee group by ADUserID having COUNT(*)>1) order by ADUserID"

# Check duplication employee in k2_system
sqlComTxt = "select UserID Account, UserCode ChrisID,* from K2_SystemUser where UserCode<>'' and UserID<>'' and IsActive=1 and UserCode in (select UserCode from K2_SystemUser K left join IGrowEmployee I on K.UserCode=I.CHRISID group by UserCode having count(*)>1) order by UserDisplayNameEN"
sqlComTxt = "select UserID Account, UserCode ChrisID,* from K2_SystemUser where IsActive=1 order by UserDisplayNameEN"
# open sql server connection & cursor
server = "GAOLAN"  # UAT
user = "IWFAdmin"
password = "3edcVFR$"
dbname = "CapitalWorkflowCenter_UAT"
con = pymssql.connect(server, user, password, dbname)
cursor = con.cursor(as_dict=True)

cursor.execute(sqlComTxt)
dbRows = cursor.fetchall()

outputData = []
for row in dbRows:
    outputData.append(row)

# close
cursor.close()
con.close()

# write missing data in excel
f = xlwt.Workbook()
sheet1 = f.add_sheet(u'Result', cell_overwrite_ok=True)
# add column name
for columnName in outputData:
    i = 0
    for item in columnName:
        sheet1.write(0, i, item)
        i += 1

rowIndex = 1
for rowData in outputData:
    columnIndex = 0
    for column, item in rowData.items():
        sheet1.write(rowIndex, columnIndex, item)
        columnIndex += 1
    rowIndex += 1
f.save(r'D:\ResultData{0}.xls'.format(time.strftime('%Y-%m-%d_%H%M%S', time.localtime(time.time()))))
print('End at:{0}'.format(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))))
