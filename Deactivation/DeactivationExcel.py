import xlrd


def getfilterdpaidlist():
    workbook = xlrd.open_workbook("M:/Digicon/INDRA_SAT_VISION.xlsx")
    # Read data from Excel sheet named "Check"
    sheet = workbook.sheet_by_name("Sheet1")
    filter_paid = []
    flag = 0
    for curr_row in range(1, sheet.nrows):
        row_data = []
        for curr_col in range(0, 5):
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

    count = 0
    filtered_vcno = []
    for i in range(len(filter_paid)):
        if filter_paid[i][3] != '':
            filtered_vcno.append(filter_paid[i][3])

    print("filtered_vcno: ", filtered_vcno)
    print("Count: ", count)
    return filtered_vcno


def getallvcno():
    workbook = xlrd.open_workbook("I:/Digicon/Digicon/Deactivation.xlsx")
    sheet = workbook.sheet_by_name("DeactivationAll")
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


getallvcno()