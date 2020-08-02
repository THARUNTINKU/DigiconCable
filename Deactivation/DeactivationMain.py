from Digicon.Deactivation.DeactivationService import Deactivation
from Digicon.Deactivation.DeactivationExcel import getallvcno,getfilterdpaidlist
import time
import xlwt


objdeactivate = Deactivation()
# Step-1: Login
objdeactivate.login()
# Step-2: Get list choose all or selected from Y
deactivationlist = getallvcno()
# reactivationlist = getfilterdpaidlist()
alreadydeactivatedlist = []

workbook = xlwt.Workbook()
sheet = workbook.add_sheet('NotReactivated')
count = 0

for i in deactivationlist:
    objdeactivate.searchbarclick(i)
    resset = objdeactivate.deactivatebutton2()
    if resset != 'Y':
        alreadydeactivatedlist.append(i)
        sheet.write(0, count, i)
        count = count + 1
    time.sleep(1)

workbook.save('I:/Digicon/Digicon/AutoGeneration/AlreadyDeactivated.xlsx')
print("Already Deactivated_List: ", alreadydeactivatedlist)
print("Count: ", len(alreadydeactivatedlist))

