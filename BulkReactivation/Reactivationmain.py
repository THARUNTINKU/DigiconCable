from Digicon.BulkReactivation.Reactivation import Reactivation
from Digicon.BulkReactivation.ReactivationExcel import getfilterdpaidlist,getallvcno
import time
import xlwt


objreactivate = Reactivation()
# Step-1: Login
objreactivate.login()
# Step-2: Get list choose all or selected from Y
# reactivationlist = getallvcno()
reactivationlist = getfilterdpaidlist()
notreactivatedlist = []

workbook = xlwt.Workbook('I:/Digicon/Digicon/NotReactivated.xlsx')
sheet = workbook.add_sheet('NotReactivated2')
count = 0

for i in reactivationlist:
    objreactivate.searchbarclick(i)
    resset = objreactivate.reactivatebutton()
    if resset != 'Y':
        notreactivatedlist.append(i)
        sheet.write(0, count, i)
        count = count + 1
    time.sleep(1)

workbook.save('I:/Digicon/Digicon/NotReactivated.xlsx')
print("Not Reactivated_List: ", notreactivatedlist)
print("Count: ", len(notreactivatedlist))

