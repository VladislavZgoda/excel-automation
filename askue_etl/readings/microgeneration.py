from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import cast

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet


@dataclass(frozen=True)
class MeterData:
    date: datetime
    T1_import: float | None
    T2_import: float | None
    T_import: float | None
    T1_export: float | None
    T2_export: float | None
    T_export: float | None


type MeterReadings = dict[str, MeterData]
type Meters = set[str]


def prepare_readings(readings_path: Path, template_path: Path) -> MeterReadings:
    meters = _get_meters_from_template(template_path)
    meter_readings = _get_readings(readings_path, meters)

    return meter_readings


def _get_meters_from_template(template_path: Path) -> Meters:
    wb = load_workbook(template_path, read_only=True)
    ws = cast(Worksheet, wb.active)

    meters: Meters = set()

    for row in range(3, ws.max_row + 1):
        str_row = str(row)
        meter = str(ws["C" + str_row].value)
        meters.add(meter)

    wb.close()
    return meters


def _get_readings(readings_path: Path, meters: Meters) -> MeterReadings:
    wb = load_workbook(readings_path, read_only=True)
    ws = cast(Worksheet, wb.active)

    meter_readings: MeterReadings = {}

    for row in range(3, ws.max_row + 1):
        str_row = str(row)
        meter = str(ws["C" + str_row].value).zfill(8)

        if meter not in meters:
            continue

        meter_readings[meter] = MeterData(
            date=ws["D" + str_row].value,
            T1_import=ws["E" + str_row].value,
            T2_import=ws["F" + str_row].value,
            T_import=ws["H" + str_row].value,
            T1_export=ws["I" + str_row].value,
            T2_export=ws["J" + str_row].value,
            T_export=ws["L" + str_row].value,
        )

    wb.close()
    return meter_readings
