#!/usr/bin/python3
# Version 0.1.0
'''
Скрипт для преобразования электронной таблицы в текст, форматированный для RedMine

Необходим Python2.6 и выше

Модули, от которых зависит:
pyexcel
pyexcel-xls
pyexcel-ods

Список поддерживаемых форматов: .odt .xls

На вход скрипту подается электронная таблица (например, файл в формате .ods). Он преобразовывается в формат,
который воспринимает RedMine для построения своих таблиц и сохраняется в файл .txt либо выводится на экран.

Для корректной работы необходимо соблюдать верную разметку таблицы в исходном файле:
* Если нужно объединить несколько ячеек по горизонтали, то сначала оставляем необходимое количество пустых ячеек,
  в последней ячейке пишем нужные данные;
* Если в одном исходном файле несколько таблиц, то их нужно разделить, записав в строку между ними разделительную
  последовательность HORIZONTAL_TABLE_SPLITTER, по умолчанию: #HSPLIT;
* Если ячейку нужно оставить пустой, то в нее следует записать последовательность EMPTY_CELL, по умолчанию: #EMPTY;
* Если нужно слить ячейки вертикально, то значение итоговой ячейки пишем сверху, затем в необходимое количество ячеек
  ниже записываем последовательность VERTICAL_JOIN, по умолчанию #VJOIN;

Пример вида таблицы, подаваемой на вход:
+-------+------+--+--+--+--+
|a1     |a2    |a3|  |  |a4|
+-------+------+--+--+--+--+
|#EMPTY |#VJOIN|  |b4|b5|b6|
+-------+------+--+--+--+--+
|#HSPLIT|      |  |  |  |  |
+-------+------+--+--+--+--+
|#STR   |string|  |  |  |  |
+-------+------+--+--+--+--+
|e1     |e2    |e3|  |  |e4|
+-------+------+--+--+--+--+
|#EMPTY |#VJOIN|  |g4|g5|g6|
+-------+------+--+--+--+--+
|#HSPLIT|      |  |  |  |  |
+-------+------+--+--+--+--+
'''


import pyexcel as pe
import argparse
import collections


# Разделитель между двумя разными таблицами
HORIZONTAL_TABLE_SPLITTER = "#HSPLIT"
EMPTY_CELL = "#EMPTY"
STRING_LINE = "#STR"
VERTICAL_JOIN = "#VJOIN"
# TODO Поддержать новые управляющие значения
VERTICAL_TABLE_SPLITTER = "#VSPLIT"


#-------------------------------------------------------------------------------------------------------------------------
# Считываем данные из файла excel и возвращаем их в виде списка [[строка, номер, 1],[строка, номер, 2],[строка, номер, 3]]
#-------------------------------------------------------------------------------------------------------------------------
def get_raw_table_data(filename : str) -> list(list()):
    '''
    Функция вычитывает данные ячеек из файла таблицы и преобразовывает их к виду:
    [[строка, номер, 1],[строка, номер, 2],[строка, номер, 3]]

    Входные параметры:
    filename - путь к файлу;

    Возвращает список списков данных ячеек файла
    '''
    sheet = pe.get_sheet(file_name = filename)
    return sheet.to_array()


#-------------------------------------------------------------------------------------------------------------------------
# Разделяем лист с несколькими таблицами на отдельные таблицы
#-------------------------------------------------------------------------------------------------------------------------
def split_raw_table_data_for_tables(raw_data : list(list())) -> list(list(list())):
    '''
    Функция делит весь полученный лист на отдельные таблицы. Прочитывает весь файл, в нем ищет HORIZONTAL_TABLE_SPLITTER, если
    находит, то сохраняет все вычитанное как отдельную таблицу вида [[строка, номер, 1], [срока, номер, 2]]
    
    Входные параметры:
    raw_data - список списков, который содержит все данные, вычитанные из файла

    Возвращает трехмерный список вида [Таблица1, Таблица2]
    '''
    tables = []
    table = []
    for line in raw_data:
        if HORIZONTAL_TABLE_SPLITTER in line:
            table.append(line)
            tables.append(table)
            table = []
        else:
            table.append(line)
    if table != []:
        tables.append(table)
    return tables


#-------------------------------------------------------------------------------------------------------------------------
# Конвертирование набора таблиц в синтаксис редмайна
#-------------------------------------------------------------------------------------------------------------------------
def convert_few_tables_in_sheet_2_redmine(tables : list(list(list())), bold_rows_number : int, bold_columns_number : int) -> str:
    '''
    Функция используется для конвертирования листа электронной таблицы, в котором содержится несколько отдельных таблиц
    в строку, отформатированную для разметки redmine

    Входные параметры:
    tables - трехмерный список, содеражащий в себе несколько разделенных таблиц
    bold_rows_number - количество верхних строк, которые нужно сделать жирными
    bold_columns_number - количество левых столбцов, которые нужно сделать жирными

    Возвращает строку, отформатированную в синтаксисе RedMine
    '''
    result_str = ''
    for table in tables:
        result_str += convert_raw_data_2_redmine(table, bold_rows_number, bold_columns_number)
    
    return result_str


#-------------------------------------------------------------------------------------------------------------------------
# Преобразовываем данные в формат, пригодный для Redmine
# На текущий момент возможно только горизонтальное слияние
#-------------------------------------------------------------------------------------------------------------------------
def convert_raw_data_2_redmine(data : list(list()), bold_rows_number : int, bold_columns_number : int) -> str:
    '''
    Функция преобразует сырое содержание таблицы из списка списков в строку, отформатированную
    для разметки redmine

    Входные параметры:
    data - список списков сырых значений ячеек таблицы;
    bold_rows_number - количество верхних строк, которые нужно сделать жирными
    bold_columns_number - количество левых столбцов, которые нужно сделать жирными

    Возвращает строку, отформатированную в синтаксисе RedMine
    '''
    result_str = ''
    # Делаем нужное количество строк и столбцов ячейки жирными
    bold_upper_rows(bold_rows_number, data)
    bold_left_columns(bold_columns_number, data)
    # Вертикальное слияние
    vertical_join(data)
    # Перебираем строки
    for line in data:
            result_str += convert_raw_line_2_redmine(line)
    # Убираем множества пустых строк
    while '|\n|\n' in result_str:
        result_str = result_str.replace('|\n|\n', '|\n\n')
        result_str = result_str.replace('\n\n', '\n')
    # Заменяем табличный разделитель на пустую строку
    result_str = result_str.replace(HORIZONTAL_TABLE_SPLITTER, '')
    
    return result_str


#-------------------------------------------------------------------------------------------------------------------------
# Конвертируем сырую строку таблицы в формат, пригодный для редмайн
#-------------------------------------------------------------------------------------------------------------------------
def convert_raw_line_2_redmine(line : list) -> str:
    '''
    Функция преобразует одну строку в вид, пригодный для разметки редмайн

    Входные параметры:
    line - список значений ячеек строки;

    Возвращает строку, содержащую значения ячеек, отформатированных в синтаксисе RedMine
    '''
    merge_count = 1
    merge_flag = False
    result_str = '|'

    # Обрабатываем ячейки, которые должны быть пустыми
    while EMPTY_CELL in line:
        ind = line.index(EMPTY_CELL)
        line[ind] = ' '
    # Обрабатываем строку, содержащую горизонтальный разделитель
    if HORIZONTAL_TABLE_SPLITTER in line:
        return HORIZONTAL_TABLE_SPLITTER + "\n"
    # Обрабатываем строку, которая должна быть просто строкой
    elif STRING_LINE in line:
        ind = line.index(STRING_LINE)
        try:
            return line[ind + 1] + "\n"
        except:
            return ''
    elif '' in line:
        # Если в строке есть пустые ячейки, значит нужно провести
        # Объединение ячеек. Для этого нужно перебрать все ячейки строки
        for cell in line:
            if cell == '':
                merge_count += 1
                merge_flag = True
            else:
                if merge_flag:
                    result_str = f"{result_str}\{merge_count}=.{cell}|"
                    merge_flag = False
                    merge_count = 1
                else:
                    result_str = f"{result_str}{cell}|"
        return result_str + '\n'
    else:
        # Если в строке нет пустых ячеек или управляющих конструкций, то объединение строк не требуется
        # В таком случае все ячейки можно записать в строку через разделитель "|"
        return '|' + '|'.join(line) + '|\n'


#-------------------------------------------------------------------------------------------------------------------------
# Вертикальное слияние ячеек
#-------------------------------------------------------------------------------------------------------------------------
def vertical_join(raw_table : list(list())):
    '''
    Функция объединяет ячейки вертикально

    Входные данные:
    raw_table - список списков, содержащий таблицу в сыром виде;

    Ничего не возвращает, изменяет таблицу
    '''
    # Процесс мержинга в указанный номер строки
    merging = False
    # Номер строки куда мержим
    merging_row_index = 0
    # Номера столбцов куда мержим и количество сливаемых ячеек
    merging_cell_in_row = {}
    # Число сливаемых ячеек
    merging_count = 1

    # Перебираем строки в таблице
    for line in raw_table:
        # Если в строке есть последовательность вертикального слияния и это не верхняя строка
        if VERTICAL_JOIN in line and raw_data.index(line) != 0:
            merging = True
            # Заменяем все ячейки вертикального слияния в строке на "#DELETE"
            # Запоминаем индексы столбцов для слияния и количество ячеек слияния
            while VERTICAL_JOIN in line:
                ind = line.index(VERTICAL_JOIN)
                merging_cell_in_row[ind] = merging_cell_in_row.get(ind, 0) + 1
                line[ind] = "#DELETE"
            # Удаляем все ячейки, помеченные как "#DELETE"
            while "#DELETE" in line:
                ind = line.index("#DELETE")
                line.pop(ind)
        else:
            if merging:
                merging = False
                # Заполняем синтаксис слияния
                for cell_index in merging_cell_in_row.keys():
                    raw_table[merging_row_index][cell_index] = f"/{merging_cell_in_row.get(cell_index, 0) + 1}.{raw_table[merging_row_index][cell_index]}"
                merging_cell_in_row = {}
            else:
                # получаем номер строки, куда сливаем
                merging_row_index = raw_table.index(line)


#-------------------------------------------------------------------------------------------------------------------------
# Сохраняем получившееся в файл
#-------------------------------------------------------------------------------------------------------------------------
def save_red_data(red_data : str, filename : str) -> None:
    '''
    Функция сохраняет отфоматированную таблицу в файл

    Входные параметры:
    red_data - строка, содержащая таблицу, отформатированную в синтаксис RedMine;
    filename - имя файла, в который будут сохранены данные;

    Не возвращает никакое значение
    '''
    with open(filename, 'w') as f:
        f.write(red_data)
        f.close()


#-------------------------------------------------------------------------------------------------------------------------
# Сделать жирными некоторое количество верхних строк
#-------------------------------------------------------------------------------------------------------------------------
def bold_upper_rows(number : int, raw_data : list(list())) -> None:
    '''
    Функция добавляет синтаксис жирного текста (*bold text*) к указанному количеству верхних строк.

    Входные параметры:
    number - число верхних строк, которые необходимо сделать жирными;
    raw_data - содержимое таблицы в сыром виде, представляет собой список списков значений ячеек; 

    Ничего не возвращает, только изменяет переданную таблицу
    '''
    # Если первые строки это именно строки, то пропускаем
    for line in raw_data:
        if STRING_LINE in line:
            continue
        # Получаем индекс строки и делаем строки жирными
        start = raw_data.index(line)
        end = start + number
        if end > len(raw_data):
            end = len(raw_data)
        for i in range(start, end):
            for cell in raw_data[i]:
                # Пропускаем если ячейка уже жирная
                if cell.startswith('*') and cell.endswith('*'):
                    continue
                # Пропускаем если ячейка пустая
                elif cell == '':
                    continue
                # Пропускаем если ячейка содержит управляющую конструкцию
                elif cell_in_keywords(cell):
                    continue
                ind = raw_data[i].index(cell)
                raw_data[i][ind] = f"*{cell.strip()}*"
        return None


#-------------------------------------------------------------------------------------------------------------------------
# Сделать жирными некоторое количество левых столбцов
#-------------------------------------------------------------------------------------------------------------------------
def bold_left_columns(number : int, raw_data : tuple(list())) -> None:
    '''
    Функция добавляет синтаксис жирного текста (*bold text*) к указанному количеству левых столбцов.

    Входные параметры:
    number - число левых столбцов, которые необходимо сделать жирными;
    raw_data - содержимое таблицы в сыром виде, представляет собой список списков значений ячеек; 

    Ничего не возвращает, только изменяет переданную таблицу
    '''    
    for line in raw_data:
        # Если первые строки это именно строки, то пропускаем
        if STRING_LINE in line:
            continue
        # Делаем столбцы жирными
        for i in range(number):
            # Пропускаем если ячейка уже жирная
            if line[i].startswith('*') and line[i].endswith('*'):
                continue
            # Пропускаем если ячейка пустая
            elif line[i] == '':
                continue
            # Пропускаем если ячейка содержит управляющую конструкцию
            elif cell_in_keywords(line[i]):
                continue
            line[i] = f"*{line[i].strip()}*"


#-------------------------------------------------------------------------------------------------------------------------
# Проверка вхождения значения ячейки в управляющие ключевые слова
#-------------------------------------------------------------------------------------------------------------------------
def cell_in_keywords(cell : str) -> bool:
    '''
    Функция проверяет входит ли значение ячейки в список зарезервированных ключевых слов

    Входные параметры:
    cell - строка, содержащая значение ячейки таблицы;

    Возвращает булевое значение
    '''
    keywords = [HORIZONTAL_TABLE_SPLITTER, VERTICAL_TABLE_SPLITTER, EMPTY_CELL, STRING_LINE, VERTICAL_JOIN]
    return cell in keywords


if __name__ == '__main__':
    # Проверка на наличие аргумента при запуске скрипта
    parser = argparse.ArgumentParser(description = 'e-table to RedMine Converter')
    parser.add_argument('file', help = 'e-table file for convert to RedMine synt')
    parser.add_argument('--out', help = 'File for saving text result. If not defind, result will be printed', default = None)
    parser.add_argument('--br', help = 'Bold few upest rows. Enter the number of upper rows to bold them', type = int, default = 0)
    parser.add_argument('--bc', help = 'Bold few left columns. Enter the number of left columns to bold them', type = int, default = 0)
    parser.add_argument('-o','--OneTable', help = 'Point this if onle one table in etable file. Use for less work time', action = 'store_true', default = False)

    args = parser.parse_args()
    out_file = args.out
    file = args.file
    bold_row_number = args.br
    bold_columns_number = args.bc
    one_table = args.OneTable

    # Вычитываем данные из файла
    raw_data = get_raw_table_data(file)
    if one_table:
        red_data = convert_raw_data_2_redmine(raw_data, bold_row_number, bold_columns_number)
    else:
        splited_tables = split_raw_table_data_for_tables(raw_data) 
        red_data = convert_few_tables_in_sheet_2_redmine(splited_tables, bold_row_number, bold_columns_number)
    if out_file is None:
        print(red_data)
    else:
        save_red_data(red_data, out_file)