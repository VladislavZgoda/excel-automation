from pathlib import Path

import polars as pl

script_dir = Path(__file__).resolve().parent
supplement_nine_path = script_dir / "input_files" / "Приложение №9 Юр.xlsx"
export_1c = script_dir / "input_files" / "Список.xlsx"

df_supplement_nine = pl.read_excel(supplement_nine_path, read_options={"header_row": 1})

df_1c = (
    df_supplement_nine.filter(pl.col("Способ снятия показаний") != "Ридер")
    .select(
        "№ п/п",
        "Л/С",
        pl.lit(None).alias("Наименование договора"),
        pl.lit(None).alias("Код ТУ"),
        pl.lit(None).alias("Наименование ТУ"),
        "Номер_ПУ",
        "Дата",
        "Т1",
        "Т2",
        "Т3",
        "Т сумм",
        pl.lit(None).alias("Филиал сбыт"),
    )
    .rename({"Л/С": "№ договора", "Т сумм": "Т_сумм", "Дата": "Дата КСП"})
    .filter(pl.col("№ договора").str.starts_with("230710"))
)

# 28 - столбец с s/n ПУ, 4 - Код ТУ.
df_export_1c = pl.read_excel(export_1c, read_options={"header_row": 1}).select(
    pl.col("28"), pl.col("4")
)

df_1c = (
    df_1c.join(
        df_export_1c.select(["28", "4"]), left_on="Номер_ПУ", right_on="28", how="left"
    )
    .select(
        "№ п/п",
        "№ договора",
        "Наименование договора",
        pl.col("4").alias("Код ТУ"),
        "Наименование ТУ",
        "Номер_ПУ",
        "Дата КСП",
        "Т1",
        "Т2",
        "Т3",
        "Т_сумм",
        "Филиал сбыт",
    )
    .with_columns(
        pl.col("Дата КСП").str.strptime(pl.Date, format="%d.%m.%Y").alias("date")
    )
)

max_year = df_1c.select(pl.col("date").dt.year().max()).item()

max_month = (
    df_1c.filter(pl.col("date").dt.year() == max_year)
    .select(pl.col("date").dt.month().max())
    .item()
)

max_month_cond = (pl.col("date").dt.year() == max_year) & (
    pl.col("date").dt.month() == max_month
)

new_date_expr = (pl.col("date").dt.month_start() - pl.duration(days=1)).dt.strftime(
    "%d.%m.%Y"
)

df_1c = df_1c.with_columns(
    pl.when(max_month_cond)
    .then(new_date_expr)
    .otherwise(pl.col("Дата КСП"))
    .alias("Дата КСП")
).drop("date")

df_1c.write_excel(
    script_dir / "output_files" / "исуэ.xlsx",
    float_precision=2,
    autofit=True,
)
