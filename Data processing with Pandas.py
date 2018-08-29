# -*- coding: utf-8 -*-

import pandas as pd
import xlwt
##from pandas import ExcelWriter



def open_file(name):
    """ Считывает данные из файла .xls и передает их в переменную data
    (DataFrame)."""
    data = pd.read_excel('data/{}.xls'.format(name))
    return data


def processing(data):
    """ Подсчитывает количество уникальных ответов на вопрос и передает
    результаты в переменную new_data (Series)."""

##  Передалать Series в DataFrame!
    new_data = pd.Series()

    for i in data.columns:
        new_data[i] = data[i].value_counts().reset_index()
        
    return new_data


def recording_file(name, data, new_data):
    """ Записывает переменную new_data в файл .xls"""

##    Реализовать запись в файл средствами самого Pandas!
##    with ExcelWriter('{}.xls'.format(name + '_Обработанный')) as writer:
##        data.to_excel(writer, 'Лист1', index=False)
##        new_data.to_excel(writer, 'Лист2', index=False)
##        writer.save()

    sheet = xlwt.Workbook()
    ws = sheet.add_sheet('Лист1')

    # Записывает в каждый второй столбец первой строки названия вопросов
    for i_index, i_value in enumerate(data.columns):
        ws.write(0, i_index * 2, i_value)

        # Записывает столбец-ответ и столбец-количество ответов для вопроса
        for j_index, j_value in enumerate(new_data[i_value].values):
            ws.write(j_index + 1, i_index * 2, str(j_value[0]))
            ws.write(j_index + 1, i_index * 2 + 1, int(j_value[1]))

    sheet.save(name + '_Ответы' + '.xls')


def main():
    name = input("Введите имя файла для обработки \n")
    data = open_file(name)
    new_data = processing(data)
    recording_file(name, data, new_data)



if __name__ == '__main__':
    main()
