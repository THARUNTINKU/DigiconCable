import xlrd


def getfilterdpaidlist():
    workbook = xlrd.open_workbook("D:/Personal/Cable/Digicon/INDRA_SAT_VISION.xlsx")
    # Read data from Excel sheet named "Check"
    filter_paid = []
    flag = 0
    # "Vallur-1", "Vallur-2"
    sheet_list = ["SankarK", "Pudhur", "Muslim St", "A.Colony", "Kamaraj Nagar", "Etteiyampatti"]
    for sheetname in sheet_list:
        sheet = workbook.sheet_by_name(sheetname)
        for curr_row in range(1, sheet.nrows):
            row_data = []
            for curr_col in range(0, 6):
                # Read the data in the current cell
                data = sheet.cell_value(curr_row, curr_col)
                # print("curr_row", curr_row, "curr_col", curr_col, "Data: ", data)
                if curr_col == 2 and (data == 'P' or data == 'p'):
                    # print(data)
                    flag = 1
                row_data.append(data)
            if flag == 1:
                flag = 0
                filter_paid.append(row_data)

    filtered_vcno = []
    for i in range(len(filter_paid)):
        if filter_paid[i][5] != '':
            filtered_vcno.append(filter_paid[i][5])

    print("Filtered_vcno: ", filtered_vcno)
    print("Filter count: ", len(filtered_vcno))
    return filtered_vcno


def getdeactivatelist():
    workbook = xlrd.open_workbook("D:/Personal/Cable/Digicon/INDRA_SAT_VISION.xlsx")
    # Read data from Excel sheet named "Check"
    deactivate_list = []
    dflag = 0
    # "Vallur-1", "Vallur-2"
    sheet_list = ["Vallur-1", "Vallur-2", "SankarK", "Pudhur", "Muslim St", "A.Colony", "Kamaraj Nagar", "Etteiyampatti"]
    for sheetname in sheet_list:
        sheet = workbook.sheet_by_name(sheetname)
        for curr_row in range(1, sheet.nrows):
            row_data = []
            for curr_col in range(0, 6):
                # Read the data in the current cell
                data = sheet.cell_value(curr_row, curr_col)
                # print("curr_row", curr_row, "curr_col", curr_col, "Data: ", data)
                if curr_col == 2 and (data == 'D' or data == 'd'):
                    # print(data)
                    dflag = 1
                row_data.append(data)
            if dflag == 1:
                dflag = 0
                deactivate_list.append(row_data)

    deactivate_vcno = []
    for i in range(len(deactivate_list)):
        if deactivate_list[i][4] != '':
            deactivate_vcno.append(deactivate_list[i][4])

    print("Deactivate_vcno: ", deactivate_vcno)
    print("Deactivate count: ", len(deactivate_vcno))
    return deactivate_vcno


getfilterdpaidlist()
# getdeactivatelist()