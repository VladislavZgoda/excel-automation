add_missing_readings:
1. Создать папки input_files и output_files.
2. В папку input_files сохранить xlsx файлы read_data и write_data:
  - read_data.xlsx - Отчет "Новые показания" из Пирамида 2
  - write_data.xlsx - Приложение №9
3. Создать bat файл для запуска скрипта в cmd.
  К примеру:
  """
  @py.exe D:\repos\excel-automation\add_missing_readings\main.py %*
  pause
  """

process_legal_entities:
1. Создать папки input_files, output_files и templates.
2. В папку input_files сохранить xlsx файлы meter_readings и current_meter_readings:
   - meter_readings.xlsx - Отчет "Новые показания" из Пирамида 2
   - current_meter_reading.xlsx - Выгрузка показаний из Пирамида 2 с А+ текущие
3. В папку templates сохранить ведомости без показаний.
Создать bat файл для запуска скрипта в cmd.
  К примеру:
  """
  @py.exe D:\repos\excel-automation\process_legal_entities\main.py %*
  pause
  """
