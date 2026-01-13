from pathlib import Path
from datetime import date

import polars as pl
from xlsxwriter import Workbook
from natsort import index_natsorted

script_dir = Path(__file__).resolve().parent
matritca_readings_path = script_dir / "input_files" / "matritca_readings.xlsx"
one_zone_meters_path = script_dir / "input_files" / "one_zone_meters.xlsx"

askue_date = date.today().strftime("%d.%m.%Y")

df_one_zone_meters = pl.read_excel(
    one_zone_meters_path, read_options={"header_row": None}
)

df = (
    pl.read_excel(matritca_readings_path, read_options={"header_row": 1})
    .slice(0, -1)  # Удалить последнюю строку
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
)

df = df.with_columns(pl.col("Л/С").str.strip_chars("\n\r\t")).filter(
    pl.col("Л/С").str.starts_with("230700"), pl.col("ФИО абонента") != "ОДПУ"
)

df = (
    df.with_columns(
        pl.lit(askue_date).alias("Дата_АСКУЭ"),
        pl.lit("УСПД").alias("Способ снятия показаний"),
        pl.col("Адрес").str.extract(r"(ТП-\d{1,3})").alias("ТП"),
    )
    # Естественный порядок сортировки
    .pipe(lambda df: df.select(pl.all().gather(index_natsorted(df["ТП"]))))
)

df = df.with_columns(
    # Добавить нумерацию для строк.
    pl.int_range(1, pl.len() + 1).alias("№ п/п"),
    pl.col("Т1", "Т2", "Т3", "Т сумм").cast(pl.Float64, strict=False).round(2),
    # Добавить 0 к началу серийного номера, если он из 7 цифр.
    pl.col("Номер_ПУ").str.zfill(8),
)

is_difference = pl.col("Т сумм").sub(pl.col("Т1").add(pl.col("Т2"))).abs().gt(1)

is_one_zone = pl.col("Номер_ПУ").is_in(
    df_one_zone_meters.get_column("column_1").implode()
)

condition = is_one_zone & is_difference

df = df.with_columns(
    (pl.when(condition).then(pl.col("Т сумм")).otherwise(pl.col("Т1"))).alias("Т1"),
    (pl.when(condition).then(0).otherwise(pl.col("Т2"))).alias("Т2"),
)

df_supplement_nine = df.select(
    "№ п/п",
    "Л/С",
    "Номер_ПУ",
    "Дата",
    "Т1",
    "Т2",
    "Т3",
    "Т сумм",
    "Адрес",
    "ФИО абонента",
    "Дата_АСКУЭ",
    "Тип ПУ",
    "Способ снятия показаний",
    "ТП",
)

font_styles = {
    "font_name": "Times New Roman",
    "font_size": 10,
}

border_styles = {
    "border": 1,
    "border_color": "black",
}

alignment_center = {
    "align": "center",
    "valign": "vcenter",
}

alignment_left = {
    "align": "left",
    "valign": "vcenter",
}

alignment_right = {
    "align": "right",
    "valign": "vcenter",
}

header_styles = {
    "font_name": "Times New Roman",
    "font_size": 12,
    **border_styles,
    **alignment_center,
}

merge_styles = {
    "font_name": "Times New Roman",
    "font_size": 14,
    **alignment_center,
}


shared_column_formats = {
    "№ п/п": {
        **font_styles,
        **border_styles,
        **alignment_center,
    },
    "Л/С": {
        **font_styles,
        **border_styles,
        **alignment_center,
    },
    "Номер_ПУ": {
        **font_styles,
        **border_styles,
        **alignment_center,
    },
    "Дата": {
        **font_styles,
        **border_styles,
        **alignment_center,
    },
    "Т1": {
        **font_styles,
        **border_styles,
        **alignment_right,
    },
    "Т2": {
        **font_styles,
        **border_styles,
        **alignment_right,
    },
    "Т3": {
        **font_styles,
        **border_styles,
        **alignment_right,
    },
    "Т сумм": {
        **font_styles,
        **border_styles,
        **alignment_right,
    },
    "Адрес": {
        **font_styles,
        **border_styles,
        **alignment_left,
    },
    "ФИО абонента": {
        **font_styles,
        **border_styles,
        **alignment_left,
    },
}

supplement_nine_path = (
    script_dir / "output_files" / f"Приложение №9 Быт {askue_date}.xlsx"
)

with Workbook(supplement_nine_path) as wb:
    ws = wb.add_worksheet()
    ws.name = "Быт"

    merge_format = wb.add_format({**merge_styles})

    ws.merge_range(
        "A1:N1",
        data="Ведомость дистанционного снятия показаний посредствам АСКУЭ и ридера",
        cell_format=merge_format,
    )

    df_supplement_nine.write_excel(
        workbook=wb,
        worksheet=ws,
        position="A2",
        float_precision=2,
        autofit=True,
        column_widths={"№ п/п": 40},
        dtype_formats={pl.Date: "dd.mm.yyyy;@", pl.Int64: "@", pl.Float64: "@"},
        header_format={**header_styles},
        column_formats={
            **shared_column_formats,
            "Дата_АСКУЭ": {
                **font_styles,
                **border_styles,
                **alignment_center,
            },
            "Тип ПУ": {
                **font_styles,
                **border_styles,
                **alignment_center,
            },
            "Способ снятия показаний": {
                **font_styles,
                **border_styles,
                **alignment_center,
            },
            "ТП": {
                **font_styles,
                **border_styles,
                **alignment_center,
            },
        },
    )

df_askue_register = df_supplement_nine.select(
    pl.all().exclude(
        "Дата_АСКУЭ",
        "Тип ПУ",
        "Способ снятия показаний",
        "ТП",
    )
).with_columns(
    pl.lit(None).alias("Ведомость_КС"),
    pl.lit("Згода В.Г.").alias("Контролер"),
)

askue_register_path = script_dir / "output_files" / f"АСКУЭ Быт {askue_date}.xlsx"

with Workbook(askue_register_path) as wb:
    ws = wb.add_worksheet()
    ws.name = "БЫТ"

    merge_format = wb.add_format({**merge_styles})

    ws.merge_range(
        "A1:L1",
        data="Ведомость дистанционного снятия показаний посредствам АСКУЭ и ридера",
        cell_format=merge_format,
    )

    df_askue_register.write_excel(
        workbook=wb,
        worksheet=ws,
        position="A2",
        float_precision=2,
        autofit=True,
        column_widths={"№ п/п": 40},
        dtype_formats={pl.Date: "dd.mm.yyyy;@", pl.Int64: "@", pl.Float64: "@"},
        header_format={**header_styles},
        column_formats={
            **shared_column_formats,
            "Ведомость_КС": {
                **border_styles,
            },
            "Контролер": {
                **font_styles,
                **border_styles,
                **alignment_center,
            },
        },
    )
