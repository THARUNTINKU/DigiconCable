from Digicon.FetchPackage.FetchPackageExcel import generateActiveDict, set_packagedetailslist

# 1. GET PACKAGE LIST KEY AND VALUE PAIR FROM EXCEL
package_dict = generateActiveDict()
# print("totaldict>>", totaldict)
# print("Filter count: ", len(totaldict))

# 2. PUT PACKAGE DETAILS TO MAIN EXCEL
set_packagedetailslist(package_dict)