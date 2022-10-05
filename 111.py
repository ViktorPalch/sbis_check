from aiogram import Bot, types
from aiogram.utils import executor
from aiogram.dispatcher import Dispatcher
from aiogram.types import ReplyKeyboardRemove, \
    ReplyKeyboardMarkup, KeyboardButton, \
    InlineKeyboardMarkup, InlineKeyboardButton
import openpyxl
import os

import datetime
import logging
from asyncio import get_event_loop
import time
import sys


def raed_excel():
    path = os.path.abspath(os.path.curdir)
    if not os.path.exists(f'{path}/documents'):
        os.makedirs(f'{path}/documents')

    path = os.path.abspath(f'{path}/documents')

    dir_list = [os.path.join(path, x) for x in os.listdir(path)]

    if len(dir_list) >= 1:
        # Создадим список из путей к файлам и дат их создания.
        date_list = [[x, os.path.getctime(x)] for x in dir_list]

        # Отсортируем список по дате создания в обратном порядке
        sort_date_list = sorted(date_list, key=lambda x: x[1], reverse=True)

        # Выведем первый элемент списка. Он и будет самым последним по дате
        list_excel = sort_date_list[0][0]

    file_to_read = openpyxl.load_workbook(list_excel, data_only=True)
    sheet = file_to_read['Таблица']
    s = {}
    for row in range(2, sheet.max_row + 1):

        if 'None' not in str(sheet[row][10].value) and '\n' not in sheet[row][10].value:

            if str(sheet[row][10].value) not in s.keys():

                if len(str(sheet[row][14].value)) == 14 or len(str(sheet[row][14].value)) == 15:
                    s14 = datetime.datetime.strptime(str(sheet[row][14].value), '%H:%M:%S.%f')
                    s14 = s14.strftime("%H:%M:%S.%f")

                elif len(str(sheet[row][14].value)) == 7:
                    s14 = datetime.datetime.strptime(str(sheet[row][14].value), '%H:%M:%S')
                    s14 = s14.strftime("%H:%M:%S")

                if len(str(sheet[row][16].value)) == 14 or len(str(sheet[row][16].value)) == 15:
                    s16 = datetime.datetime.strptime(str(sheet[row][16].value), '%H:%M:%S.%f')
                    s16 = s16.strftime("%H:%M:%S.%f")

                elif len(str(sheet[row][16].value)) == 7:
                    s16 = datetime.datetime.strptime(str(sheet[row][16].value), '%H:%M:%S')
                    s16 = s16.strftime("%H:%M:%S")

                elif len(str(sheet[row][16].value)) >= 20:

                    s16 = datetime.datetime.strptime((str(sheet[row][16].value))[0]+(str(sheet[row][16].value)[7:]), '%d %H:%M:%S.%f')
                    s16 = s16.strftime('%d %H:%M:%S.%f')

                s.update({sheet[row][10].value: [s14, s16, 1]})


            else:

                if len(str(s[sheet[row][10].value][0])) == 14 and len(s[sheet[row][10].value][0]) != 7 or len(
                        str(s[sheet[row][10].value][0])) == 15:
                    d14 = datetime.datetime.strptime(s[sheet[row][10].value][0], "%H:%M:%S.%f")
                    d14 = datetime.timedelta(hours=d14.hour, minutes=d14.minute, seconds=d14.second,
                                             microseconds=d14.microsecond)


                elif len(str(s[sheet[row][10].value][0])) == 7:
                    d14 = datetime.datetime.strptime(s[sheet[row][10].value][0], '%H:%M:%S')
                    d14 = datetime.timedelta(hours=d14.hour, minutes=d14.minute, seconds=d14.second)

                if len(str(s[sheet[row][10].value][1])) == 14 and len((s[sheet[row][10].value][1])) != 7 or len(
                        str(s[sheet[row][10].value][1])) == 15:

                    d16 = datetime.datetime.strptime(s[sheet[row][10].value][1], "%H:%M:%S.%f")
                    d16 = datetime.timedelta(hours=d16.hour, minutes=d16.minute, seconds=d16.second,
                                             microseconds=d16.microsecond)
                elif len(str(s[sheet[row][10].value][1])) == 7:
                    d16 = datetime.datetime.strptime(s[sheet[row][10].value][1], "%H:%M:%S")
                    d16 = datetime.timedelta(hours=d16.hour, minutes=d16.minute, seconds=d16.second)

                elif len(str(s[sheet[row][10].value][1])) == 18:
                    d16 = datetime.datetime.strptime(s[sheet[row][10].value][1], "%d %H:%M:%S.%f")
                    d16 = datetime.timedelta(days=d16.day, hours=d16.hour, minutes=d16.minute, seconds=d16.second,
                                             microseconds=d16.microsecond)


                if len(str(sheet[row][14].value)) == 14 or len(str(sheet[row][14].value)) == 15:
                    s14 = datetime.datetime.strptime(str(sheet[row][14].value), "%H:%M:%S.%f")
                    s14 = (datetime.timedelta(hours=s14.hour, minutes=s14.minute, seconds=s14.second,
                                              microseconds=s14.microsecond))

                elif len(str(sheet[row][14].value)) == 7:
                    s14 = datetime.datetime.strptime(str(sheet[row][14].value), "%H:%M:%S")
                    s14 = (datetime.timedelta(hours=s14.hour, minutes=s14.minute, seconds=s14.second))

                if len(str(sheet[row][16].value)) == 14 or len(str(sheet[row][16].value)) == 15:
                    s16 = datetime.datetime.strptime(str(sheet[row][16].value), "%H:%M:%S.%f")
                    s16 = (datetime.timedelta(hours=s16.hour, minutes=s16.minute, seconds=s16.second,
                                              microseconds=s16.microsecond))
                elif len(str(sheet[row][16].value)) == 7:
                    s16 = datetime.datetime.strptime(str(sheet[row][16].value), "%H:%M:%S")
                    s16 = (datetime.timedelta(hours=s16.hour, minutes=s16.minute, seconds=s16.second))

                s1 = d14 + s14
                s2 = str(d16 + s16)
                print(type(s2))
                s3 = s[sheet[row][10].value][2] + 1
                print(s2)
                s[sheet[row][10].value] = [str(s1), str(s2), s3]

    for i in s.items():
        print(i)

    # for i in s.items():
    #
    #     if len(str(i[1][0])) == 14:
    #         sss = datetime.datetime.strptime(i[1][0], "%H:%M:%S.%f")
    #         sss = datetime.timedelta(hours=sss.hour, minutes=sss.minute, seconds=sss.second,
    #                                  microseconds=sss.microsecond)
    #     elif len(str(i[1][0])) == 7:
    #         sss = datetime.datetime.strptime(i[1][0], "%H:%M:%S")
    #         sss = datetime.timedelta(hours=sss.hour, minutes=sss.minute, seconds=sss.second)
    #
    #     else:
    #         print(len(str(i[1][0])),'sss')
    #
    #     if len(str(i[1][1])) == 14:
    #         ddd = datetime.datetime.strptime(i[1][1], "%H:%M:%S.%f")
    #         ddd = datetime.timedelta(hours=ddd.hour, minutes=ddd.minute, seconds=ddd.second,
    #                                  microseconds=ddd.microsecond)
    #     elif len(str(i[1][1])) == 7:
    #         ddd = datetime.datetime.strptime(i[1][1], "%H:%M:%S")
    #         ddd = datetime.timedelta(hours=ddd.hour, minutes=ddd.minute, seconds=ddd.second)
    #
    #     else:
    #         print(len(str(i[1][1])),'ddd')
    #
    #
    #     sss = str(sss / int(i[1][2]))
    #     ddd = str(ddd / int(i[1][2]))
    #     sss = datetime.datetime.strptime(sss, '%H:%M:%S.%f')
    #     sss = sss.strftime("%H:%M:%S")
    #     ddd = datetime.datetime.strptime(ddd, '%H:%M:%S.%f')
    #     ddd = ddd.strftime("%H:%M:%S")
    #     s[i[0]] = [sss, ddd, i[1][2]]
    # ss = []
    # for i in sorted(s.items(), key=lambda item: item[1][2], reverse=True):
    #     ss.append(f'{i[1][2]} - {i[1][0]} - {i[1][1]} - {i[0]}')
    # for i in ss:
    #     print(i)


raed_excel()
