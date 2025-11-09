import re
from sys import exit
from pathlib import Path
from typing import TypedDict
from time import strftime, localtime

from openpyxl import load_workbook
from openpyxl.styles import Alignment

script_dir = Path(__file__).resolve().parent

# meter_readings.xlsx - Отчет "Новые показания" из Пирамида 2.
meter_readings_path = script_dir / "input_files" / "meter_readings.xlsx"

try:
    wb_meter_readings = load_workbook(meter_readings_path)
    ws_meter_readings = wb_meter_readings["Данные"]
except FileNotFoundError:
    print("FileNotFoundError: The file 'meter_readings.xlsx' not found.")
    exit(1)
except KeyError:
    print("KeyError: The sheet 'Данные' was not found.")
    exit(1)

# current_meter_reading.xlsx - Выгрузка показаний из Пирамида 2 с А+ текущие.
current_meter_readings_path = (script_dir /
                               "input_files" / "current_meter_readings.xlsx")

try:
    wb_current_meter_readings = load_workbook(current_meter_readings_path)
    ws_current_meter_readings = wb_current_meter_readings["Sheet"]
except FileNotFoundError:
    print("FileNotFoundError: "
          "The file 'current_meter_readings.xlsx' not found.")
    exit(1)
except KeyError:
    print("KeyError: The sheet 'Sheet' was not found.")
    exit(1)

current_date = strftime("%d.%m.%Y", localtime())
meter_readings_date = ws_meter_readings["K6"].value

alignment_date = Alignment(horizontal="center", vertical="center")
alignment_value = Alignment(horizontal="right", vertical="center")

MeterData = TypedDict(
    "MeterData",
    {
        "readings": int | float,
        "date": str,
    },
)

# key - Серийный номер ПУ.
meter_readings: dict[str, MeterData] = {}

for row in range(7, ws_meter_readings.max_row + 1):
    str_row_number = str(row)
    readings = ws_meter_readings["K" + str_row_number].value

    if not isinstance(readings, (int, float)):
        continue

    serial_number = ws_meter_readings["E" + str_row_number].value
    meter_readings[serial_number] = {
        "readings": readings,
        "date": meter_readings_date,
    }

serial_number_regex = re.compile(r"\d{6,8}")

for row in range(3, ws_current_meter_readings.max_row + 1):
    str_row_number = str(row)
    serial_number_cell = ws_current_meter_readings["A" + str_row_number].value
    serial_number = serial_number_regex.search(serial_number_cell)

    if serial_number is None:
        continue

    readings = ws_current_meter_readings["C" + str_row_number].value

    if not isinstance(readings, (int, float)):
        continue

    meter_readings[serial_number.group()] = {
        "readings": readings,
        "date": current_date,
    }

assets_dir = Path(script_dir / "assets")

for file in assets_dir.iterdir():
    wb = load_workbook(file)
    ws = wb.active

    if ws is None:
        continue

    for row in range(3, ws.max_row + 1):
        str_row_number = str(row)
        serial_number = str(ws["C" + str_row_number].value)

        if serial_number not in meter_readings:
            continue

        meter_data = meter_readings[serial_number]
        ws["H" + str_row_number].value = round(meter_data["readings"], 2)
        ws["H" + str_row_number].alignment = alignment_value
        ws["D" + str_row_number].value = meter_data["date"]
        ws["D" + str_row_number].alignment = alignment_date
        ws["K" + str_row_number].value = current_date
        ws["K" + str_row_number].alignment = alignment_date

    wb.save(script_dir / "output_files" / file.name)
