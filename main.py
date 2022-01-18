from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException

import json
import sys


def open_excel(path: str, sheet_name: str) -> (tuple, None):
    """ Open excel order
    :param path: path in arg 1
    :param sheet_name: sheet name where to look
    :return:
    """
    print('Loading document...')
    try:
        wb = load_workbook(path)
    except InvalidFileException:
        print(f'No such file or directory: {path}')
        return None
    try:
        sheet = wb.get_sheet_by_name(sheet_name)
        return wb, sheet
    except KeyError:
        print('Incorrect sheet name')
        return None


def correcting_file(wb, sheet, coll_key: int, coll_value: int, data: json) -> None:
    """ Read end correct order
    :param wb: work book object
    :param sheet: sheet object
    :param coll_key: column number to search by key
    :param coll_value: column number for replacement
    :param data: correct data
    :return:
    """
    print('Parsing document...')
    for row in range(1, sheet.max_row):
        result_to_replace = data.get(sheet.cell(row=row, column=coll_key).value)
        if result_to_replace:
            sheet.cell(row=row, column=coll_value).value = result_to_replace

    wb.save('./correct_order.xlsx')
    return


def load_json(path: str) -> (json, None):
    """ load correct data in json file
    :param path: path in arg 2
    :return: json data
    """
    print('Open json file')
    try:
        with open(path, 'r') as file:
            data = json.load(file)
            return data
    except FileNotFoundError:
        print(f'No such file or directory: {path}')
        return None


def run(path_to_excel: str, path_to_json: str, sheet: str) -> int:
    """ Entry point
    :param path_to_excel: path in arg 1
    :param sheet:  path in arg 2
    :param path_to_json:  path in arg 3
    :return:
    """
    correct_data = load_json(path_to_json)
    if not correct_data:
        return 0
    try:
        order = open_excel(
            path=path_to_excel,
            sheet_name=sheet
        )
        correcting_file(*order, coll_key=8, coll_value=6, data=correct_data)
        return 1
    except TypeError:
        return 0


if __name__ == '__main__':
    if len(sys.argv) == 2 and (sys.argv[1] == '-h' or sys.argv[1] == '--help'):
        print('python main.py <path to excel> <sheet name> <path to json>')
    elif len(sys.argv) < 3:
        print('Missing number of arguments passed (requires 3)')
    else:
        result = run(path_to_excel=sys.argv[1], sheet=sys.argv[2], path_to_json=sys.argv[3])
        print('Success!') if result else print('Fail')
