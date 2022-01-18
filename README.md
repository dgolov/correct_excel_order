# correct_excel_order
 Программа имеет два режима

1. Корректировка Excel документов по файлу json. Пробегается по строкам ищет ключ в заданой ячейке и меняет значение в другой заданной ячейке
2. Перенос данных json в Excel таблицу

----------
RUN:

-Correcting:


    python main.py <path to excel> <sheet name> <path to json>
    
    
-Make:
    
    
    python main.py <path to excel> <make mode> <path to json>

    

----------
Examples:
   
    python main.py ./mba/order.xlsx Выгрузка data2.json
    
    python main.py ./mba/order.xlsx -m data2.json
    
    python main.py ./mba/order.xlsx --make data2.json

    
Help:
    
    python main.py -h 
    
    OR
   
    python main.py --help


