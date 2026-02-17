from pathlib import Path
from sys import argv, exit

import polars as pl

if len(argv) < 2:
    print(
        "Error: Missing argument! Example: 'uv run 1с_transform_sheet/main.py askue | rider'."
    )
    exit(1)
elif argv[1] not in ["askue", "rider"]:
    print("Error: The script argument is invalid! It must be 'askue' or 'rider'.")
    exit(1)

is_askue = argv[1] == "askue"

script_dir = Path(__file__).resolve().parent
supplement_nine_path = script_dir / "input_files" / "Приложение №9 Юр.xlsx"
export_1c = script_dir / "input_files" / "Список.xlsx"

df_supplement_nine = pl.read_excel(
    supplement_nine_path,
    read_options={"header_row": 1},
    schema_overrides={
        "Т1": pl.String,
        "Т2": pl.String,
        "Т сумм": pl.String,
    },
)

readings_source_cond = (
    pl.col("Способ снятия показаний") != "Ридер"
    if is_askue
    else pl.col("Способ снятия показаний") == "Ридер"
)

df_1c = (
    df_supplement_nine.filter(readings_source_cond)
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
    .filter(
        pl.col("№ договора").str.starts_with("230710")
        | pl.col("№ договора").str.starts_with("230760")
    )
)

# 28 - столбец с s/n ПУ, 4 - Код ТУ.
df_export_1c = pl.read_excel(export_1c, read_options={"header_row": 1}).select(
    pl.col("28"), pl.col("4")
)

df_1c = df_1c.join(
    df_export_1c.select(["28", "4"]), left_on="Номер_ПУ", right_on="28", how="left"
).select(
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

if is_askue:
    df_1c = df_1c.with_columns(
        pl.col("Дата КСП")
        .str.strptime(pl.Date, format="%d.%m.%Y", strict=False)
        .alias("date"),
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

if not is_askue:
    df_1c = df_1c.with_columns(
        pl.col("Дата КСП")
        .str.to_date(format="%Y-%m-%d", exact=False)
        .dt.strftime("%d.%m.%Y")
    )

output_path = (
    script_dir / "output_files" / "исуэ.xlsx"
    if is_askue
    else script_dir / "output_files" / "ридер.xlsx"
)

df_1c.write_excel(
    output_path,
    float_precision=2,
    autofit=True,
)
