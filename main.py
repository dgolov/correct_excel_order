from openpyxl import load_workbook
import json


def open_excel(path: str, sheet_name: str) -> (tuple, None):
    print('Loading document...')
    wb = load_workbook(path)
    print(type(wb))
    try:
        sheet = wb.get_sheet_by_name(sheet_name)
        print(type(sheet))
        return wb, sheet
    except KeyError:
        return None


def correcting_file(wb, sheet, row_key: int, row_value: int) -> None:
    print('Parsing document...')
    for row in range(1, sheet.max_row):
        result = data.get(sheet.cell(row=row, column=row_key).value)
        if result:
            sheet.cell(row=row, column=row_value).value = result

    wb.save('./correct_order.xlsx')
    print('Success!')
    return


with open('data2.json', 'r') as file:
    print('Open json file')
    data = json.load(file)


try:
    order = open_excel(
        path='./mba/order.xlsx',
        sheet_name="Выгрузка"
    )
    correcting_file(*order, row_key=8, row_value=6)
except TypeError:
    print('Incorrect sheet name')
