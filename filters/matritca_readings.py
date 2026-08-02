from datetime import datetime
from pathlib import Path
from typing import Literal
from zoneinfo import ZoneInfo

import polars as pl
from natsort import index_natsorted

BalanceGroupType = Literal["Быт", "Юр"]


def filterReadings(
    readings_path: Path, balance_group: BalanceGroupType
) -> pl.DataFrame:
    askue_date = datetime.now(ZoneInfo("Europe/Moscow")).strftime("%d.%m.%Y")
    consumer_number_filter = "230700" if balance_group == "Быт" else "230710"

    filters = [pl.col("Л/С").str.starts_with(consumer_number_filter)]
    if balance_group == "Быт":
        filters.append(pl.col("ФИО абонента") != "ОДПУ")

    return (
        pl.read_excel(readings_path, read_options={"header_row": 1})
        .head(-1)
        .rename(
            {
                "#": "№ п/п",
                "Код потребителя": "Л/С",
                "Серийный №": "Номер_ПУ",
                "Активная энергия, импорт, тариф1": "Т1",
                "Активная энергия, импорт, тариф2": "Т2",
                "Активная энергия, импорт, тариф3": "Т3",
                "Активная энергия, импорт": "Т сумм",
                "Наименование точки учета": "ФИО абонента",
                "Тип устройства": "Тип ПУ",
            }
        )
        .with_columns(pl.col("Л/С").str.extract(r"(\d{12})"))
        .filter(*filters)
        .with_columns(
            pl.lit(askue_date, dtype=pl.String).alias("Дата_АСКУЭ"),
            pl.lit("УСПД").alias("Способ снятия показаний"),
            pl.col("Дата").dt.to_string("%d.%m.%Y"),
            pl.col("Адрес").str.extract(r"(ТП-\d{1,3})").alias("ТП"),
        )
        .pipe(lambda df: df.select(pl.all().gather(index_natsorted(df["ТП"]))))
        .with_columns(
            # Добавить нумерацию для строк.
            pl.int_range(1, pl.len() + 1).alias("№ п/п"),
            pl.col("Т1", "Т2", "Т3", "Т сумм").cast(pl.Float64, strict=False).round(2),
            # Добавить 0 к началу серийного номера, если он из 7 цифр.
            pl.col("Номер_ПУ").str.zfill(8),
        )
    )
