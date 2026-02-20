import re
from pathlib import Path
from pprint import pprint

import pandas as pd

script_dir = Path(__file__).resolve().parent
p2_readings_path = script_dir / "input_files" / "p2_readings.xlsx"
output_path = script_dir / "output_files"


def extract_date(header: str) -> str:
    pattern = r"^(?:Показание на начало периода|Показание на конец периода) (.*)$"
    m = re.match(pattern, header)

    if m:
        return m.group(1)
    return header


df = pd.read_excel(p2_readings_path, header=[0, 1], skiprows=4)

level0_headers = df.columns.get_level_values(0)
level1_headers = df.columns.get_level_values(1)
new_level0_headers = [extract_date(val) for val in level0_headers]

df.columns = pd.MultiIndex.from_arrays([new_level0_headers, level1_headers])


def copy_by_date(row: pd.Series) -> pd.Series:
    current_date: str = row[("Дата последних показаний", "Unnamed: 8_level_1")]
    columns: list[str] = ["Т1", "Т2", "Т3", "Общая"]

    for col in columns:
        row[col] = row[(current_date, col)]

    return row

df: pd.DataFrame = df.apply(copy_by_date, axis="columns")

df = df[
    [
        "№пп",
        "Серийный номер",
        "Дата последних показаний",
        "Т1",
        "Т2",
        "Т3",
        "Общая",
        "Точка учёта",
        "Абонент",
        "Тип",
        "Наименование балансовой группы",
    ]
]

df.columns = df.columns.droplevel(1)

df.columns = [
    "№ п/п",
    "Номер_ПУ",
    "Дата",
    "Т1",
    "Т2",
    "Т3",
    "Т сумм",
    "Адрес",
    "ФИО абонента",
    "Тип ПУ",
    "ТП",
]

df.insert(1, "Л/С", None)
df["Л/С"] = df["ФИО абонента"].str.extract(r"(\d{12})")[0]

# [А-ЯЁ][а-яё]+ - первая часть фамилии.
# (?:-[А-ЯЁ][а-яё]+)? - необязательный дефис со второй частью фамилии, когда из двух частей.
# \s+ - один или несколько пробелов между фамилией и инициалами.
# [А-ЯЁ]\. - первый инициал с точкой.
# (?:\s?[А-ЯЁ]\.)? - возможный пробел и второй инициал с точкой.
pattern = r"^([А-ЯЁ][а-яё]+(?:-[А-ЯЁ][а-яё]+)?\s+[А-ЯЁ]\.(?:\s?[А-ЯЁ]\.)?)"
df["ФИО абонента"] = df["ФИО абонента"].str.extract(pattern)

# Убрать модель ПУ и S/N.
df["Адрес"] = df["Адрес"].str.split("\\").apply(lambda x: ", ".join(x[:-1]))

# Добавить 0 к началу серийного номера, если он из 7 цифр.
df["Номер_ПУ"] = df["Номер_ПУ"].apply(str).str.zfill(8)

df["Т1"] = df["Т1"].round(2).replace("н/д", None)
df["Т2"] = df["Т2"].round(2).replace("н/д", None)
df["Т3"] = df["Т3"].round(2).replace("н/д", None)
df["Т сумм"] = df["Т сумм"].round(2).replace("н/д", None)

df.to_excel(output_path / "test.xlsx")

# pprint(df.columns)
