# correct_excel_order

Корректировка Excel документов по файлу json. Пробегается по строкам ищет ключ в заданой ячейке и меняет значение в другой заданной ячейке

RUN:

    python main.py <path to excel> <sheet name> <path to json>
    

Example:
   
    python main.py ./mba/order.xlsx Выгрузка data2.json

