from dataclasses import dataclass
from io import BytesIO
from typing import cast

import polars as pl
from polars.selectors import Selector
from xlsxwriter import Workbook
from xlsxwriter.format import Format

from askue_etl.matritca_readings import BalanceGroupType

ColumnFormats = dict[str | Selector | tuple[str | Selector, ...], Format]


def create_wb_reports(
    ridings: pl.DataFrame, balance_group: BalanceGroupType
) -> tuple[BytesIO, BytesIO]:
    buffer_register = BytesIO()
    buffer_supplement_nine = BytesIO()

    merge_styles = {
        "font_name": "Times New Roman",
        "font_size": 14,
        "align": "center",
        "valign": "vcenter",
    }

    df_supplement_nine = ridings.select(
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

    with Workbook(buffer_supplement_nine, {"in_memory": True}) as wb:
        ws = cast(Workbook.worksheet_class, wb.add_worksheet(balance_group))
        merge_format = wb.add_format({**merge_styles})

        ws.merge_range(
            "A1:N1",
            data="Ведомость дистанционного снятия показаний посредствам АСКУЭ и ридера",
            cell_format=merge_format,
        )

        formats = _build_common_formats(wb)

        column_formats: ColumnFormats = {
            "№ п/п": formats.common,
            "Л/С": formats.common,
            "Номер_ПУ": formats.common,
            "Дата": formats.date,
            "Т1": formats.right,
            "Т2": formats.right,
            "Т3": formats.right,
            "Т сумм": formats.right,
            "Адрес": formats.left,
            "ФИО абонента": formats.left,
            "Дата_АСКУЭ": formats.date,
            "Тип ПУ": formats.common,
            "Способ снятия показаний": formats.common,
            "ТП": formats.common,
        }

        df_supplement_nine.write_excel(
            workbook=wb,
            worksheet=ws,
            position="A2",
            float_precision=2,
            autofit=True,
            column_widths={"№ п/п": 40, "ФИО абонента": 200, "Адрес": 180},
            dtype_formats={pl.Int64: "@", pl.Float64: "@"},
            header_format=formats.header,
            column_formats=column_formats,
        )

    df_register = df_supplement_nine.select(
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

    with Workbook(buffer_register, {"in_memory": True}) as wb:
        ws = cast(Workbook.worksheet_class, wb.add_worksheet(balance_group))

        merge_format = wb.add_format({**merge_styles})

        ws.merge_range(
            "A1:L1",
            data="Ведомость дистанционного снятия показаний посредствам АСКУЭ и ридера",
            cell_format=merge_format,
        )

        formats = _build_common_formats(wb)

        column_formats: ColumnFormats = {
            "№ п/п": formats.common,
            "Л/С": formats.common,
            "Номер_ПУ": formats.common,
            "Дата": formats.date,
            "Т1": formats.right,
            "Т2": formats.right,
            "Т3": formats.right,
            "Т сумм": formats.right,
            "Адрес": formats.left,
            "ФИО абонента": formats.left,
            "Ведомость_КС": formats.border_only,
            "Контролер": formats.common,
        }

        df_register.write_excel(
            workbook=wb,
            worksheet=ws,
            position="A2",
            float_precision=2,
            autofit=True,
            column_widths={"№ п/п": 40, "ФИО абонента": 200, "Адрес": 180},
            dtype_formats={pl.Int64: "@", pl.Float64: "@"},
            header_format=formats.header,
            column_formats=column_formats,
        )

    return buffer_register, buffer_supplement_nine


@dataclass
class CommonFormats:
    merge: Format
    header: dict[str, object]
    common: Format
    left: Format
    right: Format
    date: Format
    border_only: Format


# Format привязан к книге, в которой создан, поэтому вызывать эту
# функцию нужно один раз на каждый Workbook.
def _build_common_formats(wb: Workbook) -> CommonFormats:
    font_styles = {"font_name": "Times New Roman", "font_size": 10}
    border_styles = {"border": 1, "border_color": "black"}
    alignment_center = {"align": "center", "valign": "vcenter"}
    alignment_left = {"align": "left", "valign": "vcenter", "text_wrap": True}
    alignment_right = {"align": "right", "valign": "vcenter"}

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

    return CommonFormats(
        merge=wb.add_format({**merge_styles}),
        header={**header_styles},
        common=wb.add_format({**font_styles, **border_styles, **alignment_center}),
        left=wb.add_format({**font_styles, **border_styles, **alignment_left}),
        right=wb.add_format({**font_styles, **border_styles, **alignment_right}),
        date=wb.add_format(
            {**font_styles, **border_styles, **alignment_center, "num_format": "@"}
        ),
        border_only=wb.add_format({**border_styles}),
    )
