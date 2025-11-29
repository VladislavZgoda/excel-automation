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
2. В папку input_files сохранить xlsx файлы meter_readings, current_meter_readings и matritca_readings:
   - meter_readings.xlsx - Отчет "Новые показания" из Пирамида 2
   - current_meter_reading.xlsx - Выгрузка показаний из Пирамида 2 с А+ текущие
   - matritca_readings.xlsx  - Выгрузка показаний из Sims client.
3. В папку templates сохранить ведомости без показаний.
4. Создать bat файл для запуска скрипта в cmd.
  К примеру:
  """
  @py.exe D:\repos\excel-automation\process_legal_entities\main.py %*
  pause
  """

process_matritca_readings:
1. Создать папки input_files, output_files
2. В папку input_files сохранить xlsx файл matritca_readings:
   - matritca_readings - Выгрузка показаний из Sims client.
   Формат: #	Код потребителя	Серийный №	Дата	Активная энергия, импорт, тариф1	Активная энергия, импорт, тариф2	Активная энергия, импорт, тариф3	Активная энергия, импорт	Адрес	Наименование точки учета	Тип устройства
3. Создать bat файл для запуска скрипта в cmd.
  К примеру:
  """
  @py.exe D:\repos\excel-automation\process_legal_entities\main.py private(или legal)%*
  pause
  """
