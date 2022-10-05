import datetime
import os
import openpyxl
def raed_excel():

    path = os.path.abspath(os.path.curdir)
    if not os.path.exists(f'{path}/documents'):
        os.makedirs(f'{path}/documents')

    path = os.path.abspath(f'{path}/documents')

    dir_list = [os.path.join(path, x) for x in os.listdir(path)]
    qq = datetime.datetime.now() - datetime.datetime.now()
    if len(dir_list) >= 1:
        # Создадим список из путей к файлам и дат их создания.
        date_list = [[x, os.path.getctime(x)] for x in dir_list]

        # Отсортируем список по дате создания в обратном порядке
        sort_date_list = sorted(date_list, key=lambda x: x[1], reverse=True)

        # Выведем первый элемент списка. Он и будет самым последним по дате
        list_excel = sort_date_list[0][0]


    file_to_read = openpyxl.load_workbook(list_excel, data_only=True)
    sheet = file_to_read['Таблица']
    n=0
    qqq=datetime.datetime.now() - datetime.datetime.now()
    for row in range(2, sheet.max_row + 1):

        if 'None' not in str(sheet[row][10].value) and '\n' not in sheet[row][10].value:

            if str(sheet[row][10].value) == 'Смирнов Виктор':
                n+=1
                qq += sheet[row][14].value
                qqq+=sheet[row][16].value
                print(qq, sheet[row][14].value,sheet[row][16].value)
    print(qq)
    print(qqq)
    print(qq/n)
    print(qqq / n)
raed_excel()