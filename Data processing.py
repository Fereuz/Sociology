# -*- coding: utf-8 -*-

import xlrd, xlwt

def opening_and_reading_file(name):
    """ Открывает файл и получает данные из .xls и записывает их в
    переменную data.

        data = [[str, ..., str], ..., [str, ..., str]]
            data[0] - названия столбцов(номера вопросов)
            data[i] - i-ая строка файла, содержащая ответы на вопросы
    """

    rb = xlrd.open_workbook('data/{}.xls'.format(name), formatting_info = True)
    sheet = rb.sheet_by_index(0)
    data = [sheet.row_values(rownum) for rownum in range(sheet.nrows)]

    return data

def data_processing(data):
    """ Получает из data вариантры ответов и сами ответы на вопрос"""
    
    variants_answers = [[] for i in range(len(data[0]))]
    answers = [[] for i in range(len(data[0]))]

    # Получаем ответы на вопрос
    for i in range(len(data[0])):
        for j in range(1, len(data)): 
            try:
                answers[i].append(int(data[j][i]))
            except ValueError:
                if data[j][i] == '':
                    answers[i].append(0)
                else:
                    a = data[j][i].split(', ')
                    for k in a:
                        try:
                            answers[i].append(int(k))
                        except ValueError:
                            answers[i].append(k)

    # Получаем возможные (уникальные) варианты ответов на вопрос
    for i in range(len(answers)):
        for j in range(len(answers[i])):
            if answers[i][j] not in variants_answers[i]:
                variants_answers[i].append(answers[i][j])

    return answers, variants_answers

def counter(answers, variants_answers):
    """ Суммирует ответы на вопрос и записывает их в переменную additional_data,
    где идет чередование столбцов с вариантами ответов и количеством этих
    ответов.

        additional_data = [[str, ..., str], ..., [str, ..., str]]
            additional_data[i] - варианты ответов
                i - even numbers
            additional_data[j] - количество этих ответов
                j - odd numbers
    """

    additional_data = [[] for i in range(2 * len(data[0]))]

    for i in range(len(answers)):
        for j in variants_answers[i]:
            a = answers[i].count(j)
            additional_data[i * 2].append(j)
            additional_data[i * 2 + 1].append(a)

    return additional_data

def recording_file(name, data, additional_data):
    """ Записывает данные из переменной data в новый файл и добавляет в конце
    данные из переменной additional_data.
    """
    
    sheet = xlwt.Workbook()
    ws = sheet.add_sheet('Лист1')

    for i in range(len(data)):
        for j in range(len(data[i])):
            ws.write(i, j, data[i][j])

    for i in range(len(data[0])):
        ws.write(len(data) + 3, i * 2, data[0][i])

    k = True
    for i in range(len(additional_data)):
        if k == True:
            for j in range(len(additional_data[i])):
                ws.write(j + len(data) + 4, i, str(additional_data[i][j]))
            k = False
        else:
            for j in range(len(additional_data[i])):
                ws.write(j + len(data) + 4, i, additional_data[i][j])
            k = True

    sheet.save(name + '_Перезаписанный' + '.xls')



name = input("Введите имя файла: \n")
data = opening_and_reading_file(name)
answers, variants_answers = data_processing(data)
additional_data = counter(answers, variants_answers)
recording_file(name, data, additional_data)
