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
   - current_meter_readings.xlsx - Выгрузка показаний из Пирамида 2 с А+ текущие
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
  @py.exe D:\repos\excel-automation\process_legal_entities\main.py private(или legal) %*
  pause
  """

process_one_zone_meters:
1. Создать папки input_files, output_files
2. В папку input_files сохранить xlsx файлы matritca_readings и one_zone_meters:
   - matritca_readings - Выгрузка показаний из Sims client.
   Формат: #	Код потребителя	Серийный №	Дата	Активная энергия, импорт, тариф1	Активная энергия, импорт, тариф2	Активная энергия, импорт, тариф3	Активная энергия, импорт	Адрес	Наименование точки учета	Тип устройства
  - one_zone_meters - Серийные номера однозонных ПУ модели NP7x. Берутся из 1С энергетика при создании ведомости для загрузки показаний.
  Формат: В столбец "А" список серийных номеров
3. Создать bat файл для запуска скрипта в cmd.
  К примеру:
  """
  @py.exe D:\repos\excel-automation\process_one_zone_meters\main.py %*
  pause
  """

microgeneration:
1. Создать папки input_files, output_files, templates
2. В папку input_files сохранить xlsx файлы matritca_readings:
 - matritca_readings - Выгрузка импорта-экспорта из Sims client.
 Формат: #	"Код потребителя"	"Серийный №"	Дата	Активная энергия, импорт, тариф1	Активная энергия, импорт, тариф2	Активная энергия, импорт, тариф3	Активная энергия, импорт	Активная энергия, экспорт, тариф1	Активная энергия, экспорт, тариф2	Активная энергия, экспорт, тариф3	Активная энергия экспорт	Адрес	"Наименование точки учета"	"Тип устройства"	
3. В папку templates положить шаблоны с ТУ:
 - private для быта
 - legal для юридических лиц
Формат: "№ п/п"	Л/С	Номер_ПУ	Дата	"Т1 импорт"	"Т2 импорт"	"Т3 импорт"	"Т сумм импорт"	"Т1 экспорт"	"Т2 экспорт"	"Т3 экспорт"	Т" сумм экспорт"	Адрес	"ФИО абонента"	Дата_АСКУЭ
