import xlrd


def getallvcno():
    workbook = xlrd.open_workbook("M:/Digicon/Reactivation.xlsx")
    sheet = workbook.sheet_by_name("ReactivationAll")
    filtered_vcno = []
    flag = 1
    for curr_row in range(1, sheet.nrows):
        row_data = []
        for curr_col in range(0, 3):
            # Read the data in the current cell
            data = sheet.cell_value(curr_row, curr_col)
            # print(data)
            # print(type(data))
            if curr_col == 2 and (data == 'Y' or data == 'y'):
                flag = 0
            row_data.append(data)
        if flag == 1:
            filtered_vcno.append(row_data[1])
        else:
            flag = 1

    print("Filtered_vcno: ", filtered_vcno)
    print("Total: ", len(filtered_vcno))
    return filtered_vcno


def getfilterdpaidlist():
    workbook = xlrd.open_workbook("I:/Digicon/Digicon/INDRA_SAT_VISION.xlsx")
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
                if curr_col == 2 and (data == 'Y' or data == 'y'):
                    # print(data)
                    flag = 1
                row_data.append(data)
            if flag == 1:
                flag = 0
                filter_paid.append(row_data)

    filtered_vcno = []
    for i in range(len(filter_paid)):
        if filter_paid[i][4] != '':
            filtered_vcno.append(filter_paid[i][4])

    print("Filtered_vcno: ", filtered_vcno)
    print("Filter count: ", len(filtered_vcno))
    return filtered_vcno

# getallvcno()
getfilterdpaidlist()