TUI app для обработки xlsx файлов.

Инструкция для использования скриптов еще не переработанных в TUI.

1с_transform_sheet:
1. Создать папки input_files, output_files
2. В папку input_files сохранить xlsx файлы Приложение №9 Юр и Список:
 - Список это выгрузка реестра точек поставок в сеть ЮР лиц из 1с.
3. Вызвать через cmd uv run .\1с_transform_sheet\main.py askue или rider

p2_readings:
1. Создать папки input_files, output_files
2. В папку input_files сохранить xlsx файл p2_readings.xlsx
 - p2_readings.xlsx - это Отчет новые показания 2 из Пирамида2
3. Вызвать через cmd uv run .\p2_readings\main.py private или legal
