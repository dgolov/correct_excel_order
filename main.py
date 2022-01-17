from openpyxl import load_workbook
import json


def open_excel():
    print('Загрузка документа...')
    wb = load_workbook('./test2.xlsx')
    try:
        sheet = wb.get_sheet_by_name('Выгрузка')
        return wb, sheet
    except KeyError:
        return None


def correcting_file(wb, sheet):
    print('Парсинг документа...')
    for row in range(1, sheet.max_row):
        result = data.get(sheet.cell(row=row, column=8).value)
        if result:
            sheet.cell(row=row, column=6).value = result

    wb.save('./test3.xlsx')
    print('Успешно!')


with open('data.json', 'r') as file:
    data = json.load(file)


try:
    work_book, sheet_name = open_excel()
    correcting_file(work_book, sheet_name)
except TypeError:
    print('Неверное имя "Выгрузка"')
