import tabula
import pandas as pd
from openpyxl import *
from openpyxl.styles import Border, Side


def start(file):
    table1 = tabula.read_pdf("reports_pdf/" + file, pages='all')
    data1 = pd.DataFrame(table1[0])


    # column names dict
    dict1 = {
        "Date": "Date",
        "Total": "Total Occ.",
        "Arr.": "Arr. Rooms",
        "Comp.": "Comp. Rooms",
        "House": "House Use",
        "Deduct": "Deduct Indiv.",
        "Non-Ded.": "Non-Ded. Indiv.",
        "Deduct.1": "Deduct Group",
        "Non-Ded..1": "Non-Ded. Group",
        "Occ.%": "Occupancy-%",
        "Room Revenue": "Room Revenue",
        "Average Rate": "Average Rate",
        "Dep.": "Dep. Rooms",
        "Day Use": "Day Use Rooms",
        "No Show": "No Show Rooms",
        "OOO": "OOO Rooms",
        "Adl. &": "Adl. & Chl."}


    # changing columns names, removing row 0, deleting commas, changing "str" to "int
    data1.rename(columns=dict1,
                 inplace=True)
    data1 = data1.drop(labels=0, axis=0)
    data1 = data1.replace(',','', regex=True)


    # part one %
    data1["Occupancy-%"] = data1["Occupancy-%"].replace('%', '', regex=True)


    # converting columns to numeric
    data1[["Occupancy-%", "Total Occ.", "Arr. Rooms", "Comp. Rooms", "House Use", "Deduct Indiv.", "Non-Ded. Indiv.",
           "Deduct Group", "Non-Ded. Group", "Room Revenue", "Average Rate", "Dep. Rooms", "Day Use Rooms", "No Show Rooms",
           "OOO Rooms", "Adl. & Chl."]] = data1[["Occupancy-%", "Total Occ.", "Arr. Rooms", "Comp. Rooms", "House Use",
                                                 "Deduct Indiv.", "Non-Ded. Indiv.", "Deduct Group", "Non-Ded. Group",
                                                 "Room Revenue", "Average Rate", "Dep. Rooms", "Day Use Rooms",
                                                 "No Show Rooms", "OOO Rooms", "Adl. & Chl."]].apply(pd.to_numeric)


    # part 2 %
    data1["Occupancy-%"] = data1["Occupancy-%"] / 100


    # finding subtotal and total rows
    list1 = data1["Date"].to_list()
    list2 = []
    y = 1
    for x in list1:
        if x == "Subtotal":
            list2.append(y)
        y += 1
    y = 2
    for x in list1:
        if x == "Total":
            total = y
        y += 1


    # saving dataframe to excel and cells formatting
    df = data1
    file = file[:len(file) - 3]
    writer = pd.ExcelWriter("reports_excel/" + file + "xlsx", engine="xlsxwriter")
    df.to_excel(writer, sheet_name='sheetName', index=False, na_rep='')

    for column in df:
        column_length = max(df[column].astype(str).map(len).max(), len(column))
        col_idx = df.columns.get_loc(column)
        writer.sheets['sheetName'].set_column(col_idx, col_idx, column_length)

    workbook = writer.book
    worksheet = writer.sheets['sheetName']

    format1 = workbook.add_format({'num_format': '0.00%', "bold": "True"})
    worksheet.set_column('J:J', 11, format1) # procenty
    format2 = workbook.add_format({'num_format': '#,###.00'})
    worksheet.set_column('K:K', 14, format2) # room revenue
    format3 = workbook.add_format({"bold": "True"})
    worksheet.set_column('L:L', 11, format3) # average revenue

    writer.save()


    # openpyxl - borders and move total
    wb = load_workbook(filename="reports_excel/" + file + "xlsx")
    ws = wb["sheetName"]
    ws.move_range("A"+str(total)+":"+"Q"+str(total), rows=1, cols=0)
    total += 1

    def set_border(ws, cell_range, style2):
        rows = ws[cell_range]
        side = Side(border_style=style2, color="FF000000")

        rows = list(rows)  # we convert iterator to list for simplicity, but it's not memory efficient solution
        max_y = len(rows) - 1  # index of the last row
        for pos_y, cells in enumerate(rows):
            max_x = len(cells) - 1  # index of the last cell
            for pos_x, cell in enumerate(cells):
                border = Border(
                    left=cell.border.left,
                    right=cell.border.right,
                    top=cell.border.top,
                    bottom=cell.border.bottom
                )
                if pos_x == 0:
                    border.left = side
                if pos_x == max_x:
                    border.right = side
                if pos_y == 0:
                    border.top = side
                if pos_y == max_y:
                    border.bottom = side

                # set new border only if it's one of the edge cells
                if pos_x == 0 or pos_x == max_x or pos_y == 0 or pos_y == max_y:
                    cell.border = border

    set_border(ws, "A"+str(list2[0]+1)+":"+"Q"+str(list2[0]+1), "thin")
    set_border(ws, "A"+str(list2[1]+1)+":"+"Q"+str(list2[1]+1), "thin")
    set_border(ws, "A"+str(total)+":"+"Q"+str(total), "medium")

    wb.save(filename="reports_excel/" + file + "xlsx")