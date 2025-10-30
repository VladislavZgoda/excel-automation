add_missing_readings:
1. Создать папки input_files и output_files
2. В папку input_files сохранить xlsx файлы read_data и write_data
  - read_data.xlsx - Отчет "Новые показания" из Пирамида 2
  - write_data.xlsx - Приложение №9
3. Создать bat файл для запуска скрипта в cmd
  К примеру:
  """
  @py.exe D:\repos\excel-automation\add_missing_readings\main.py %*
  pause
  """
