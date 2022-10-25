import tabula
import pandas as pd


def start(file):
    table1 = tabula.read_pdf("reports_pdf/" + file, pages='all', stream=True)
    data1 = pd.DataFrame(table1[0])

    if len(data1.columns) > 4:
        dict1 = {
            "Unnamed: 0": "",
            "2022": "2022 DAY",
            "2022.1": "2022 MONTH",
            "2022.2": "2022 YEAR",
            "2021": "2021 DAY",
            "2021.1": "2021 MONTH",
            "2021.2": "2021 YEAR",
        }
    else:
        dict1 = {
            "Unnamed: 0": "",
            "2022": "2022 DAY",
            "2022.1": "2022 MONTH",
            "2022.2": "2022 YEAR",
        }

    # changing columns names, removing row 0, deleting commas, changing "str" to "int
    data1.rename(columns=dict1,
                 inplace=True)
    data1 = data1.drop(labels=0, axis=0)
    data1 = data1.replace(',', '', regex=True)

    # changing columns to numeric
    if len(data1.columns) > 4:
        data1[["2022 DAY", "2022 MONTH", "2022 YEAR", "2021 DAY", "2021 MONTH", "2021 YEAR"]] = \
            data1[["2022 DAY", "2022 MONTH", "2022 YEAR", "2021 DAY", "2021 MONTH", "2021 YEAR"]].apply(pd.to_numeric)
    else:
        data1[["2022 DAY", "2022 MONTH", "2022 YEAR"]] = \
            data1[["2022 DAY", "2022 MONTH", "2022 YEAR"]].apply(pd.to_numeric)

    # saving dataframe to excel and cells formatting
    df = data1
    file = file[:len(file) - 3]
    writer = pd.ExcelWriter("reports_excel/" + file + "xlsx", engine="xlsxwriter")
    df.to_excel(writer, sheet_name="Managers", index=False, na_rep='')


    worksheet = writer.sheets["Managers"]

    # formatting columns and rows
    worksheet.set_column('A:A', 35)
    worksheet.set_column('B:G', 13)

    for g in range(0, 51):
        worksheet.set_row(g, 17)

    writer.save()