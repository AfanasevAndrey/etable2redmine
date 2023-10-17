Скрипт для преобразования электронной таблицы в текст, форматированный для RedMine

Необходим Python2.6 и выше

Модули, от которых зависит:
* pyexcel
* pyexcel-xls
* pyexcel-ods

Список поддерживаемых форматов: 
* .odt
* .xls

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
```
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
```

Пример запуска:
```
python3 etable2redmine.py --br 3 --bc 1 table_file.ods
```
Для удобства используются аргументы командой строки:
```
usage: etable2redmine.py [-h] [--out OUT] [--br BR] [--bc BC] [-o] file

e-table to RedMine Converter

positional arguments:
  file            e-table file for convert to RedMine synt

options:
  -h, --help      show this help message and exit
  --out OUT       File for saving text result. If not defind, result will be printed
  --br BR         Bold few upest rows. Enter the number of upper rows to bold them
  --bc BC         Bold few left columns. Enter the number of left columns to bold them
  -o, --OneTable  Point this if onle one table in etable file. Use for less work time
```
