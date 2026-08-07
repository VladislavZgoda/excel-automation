import asyncio
import zipfile
from datetime import datetime
from io import BytesIO
from pathlib import Path
from typing import cast, get_args
from zoneinfo import ZoneInfo

from textual import on, work
from textual.app import ComposeResult
from textual.containers import Container
from textual.reactive import var
from textual.widgets import (
    Button,
    Select,
    Static,
)
from textual_fspicker import FileSave, Filters

from filters.matritca_readings import BalanceGroupType, filterReadings
from widgets.file_picker import FILE_LOCATION, FilePathSelected, FilePicker
from write_to_excel.matritca_readings import create_wb_reports


class MatritcaReadingsPanel(Container):
    readings_path: var[Path | None] = var(None)
    balance_group: var[BalanceGroupType | None] = var(None)
    wb_reports: var[tuple[BytesIO, BytesIO] | None] = var(None)

    def compose(self) -> ComposeResult:
        yield Static("Трансформировать экспорт из Sims в формат для 1С, Приложение №9.")
        yield Static("Выберете балансовую группу", id="select-btn")
        yield Select.from_values(
            get_args(BalanceGroupType), value="Быт", allow_blank=False
        )
        yield FilePicker(
            "Выберите файл с показаниями",
            "meter-readings",
        )
        yield Button(
            "Обработать данные",
            id="process-data-btn",
            classes="action-btn",
            variant="default",
            disabled=True,
        )
        yield Button(
            "Сохранить отчёты",
            id="save-file-btn",
            classes="action-btn",
            variant="default",
            disabled=True,
        )

    def on_file_path_selected(self, event: FilePathSelected) -> None:
        match event.picker_id:
            case "meter-readings":
                self.readings_path = event.file_path
            case _:
                return

        self._check_and_enable_process_data_btn()

    @on(Select.Changed)
    def select_changed(self, event: Select.Changed) -> None:
        self.balance_group = cast(BalanceGroupType, event.value)

    @on(Button.Pressed, "#process-data-btn")
    @work(thread=True)
    def handle_process_data_btn(self) -> None:
        balance_group = self.balance_group

        if self.readings_path is None:
            raise ValueError("Требуется объект Path, получен None.")
        if balance_group is None:
            raise TypeError(
                f"Требуется тип {BalanceGroupType.__str__()}, получен None."
            )
        if balance_group not in get_args(BalanceGroupType):
            raise ValueError(
                f"Требуется значение 'Быт' или 'Юр', получен {balance_group}."
            )

        self.app.call_from_thread(self._on_process_data_start)
        filtered_readings = filterReadings(self.readings_path, balance_group)
        wb_reports = create_wb_reports(filtered_readings, balance_group)

        self.app.call_from_thread(self._on_process_data_done, wb_reports)
        self.notify("Данные обработаны.", timeout=10)

    @on(Button.Pressed, "#save-file-btn")
    @work
    async def handle_save_btn(self) -> None:
        if self.wb_reports is None:
            raise ValueError("Требуется объект tuple [BytesIO, BytesIO], получен None.")
        if save_path := await self.app.push_screen_wait(
            FileSave(
                FILE_LOCATION,
                filters=Filters(("ZIP", lambda p: p.suffix.lower() == ".zip")),
            ),
        ):
            register_buf, supplement_nine_buf = self.wb_reports

            await asyncio.to_thread(
                self._write_zip,
                save_path.with_suffix(".zip"),
                register_buf.getvalue(),
                supplement_nine_buf.getvalue(),
            )

            self.notify("Файл сохранён.", timeout=10)

    def _check_and_enable_process_data_btn(self) -> None:
        if self.readings_path is None:
            return

        process_data_btn = self.query_one("#process-data-btn", Button)
        process_data_btn.disabled = False
        process_data_btn.variant = "primary"

    def _on_process_data_start(self) -> None:
        self.query_one("#process-data-btn", Button).loading = True

    def _on_process_data_done(self, wb_reports: tuple[BytesIO, BytesIO]) -> None:
        self.wb_reports = wb_reports
        self.query_one("#process-data-btn", Button).loading = False
        save_file_btn = self.query_one("#save-file-btn", Button)
        save_file_btn.disabled = False
        save_file_btn.variant = "success"

    def _write_zip(
        self,
        path: Path,
        register: bytes,
        supplement: bytes,
    ) -> None:
        askue_date = datetime.now(ZoneInfo("Europe/Moscow")).strftime("%d.%m.%Y")

        with (
            open(path, "wb") as f,
            zipfile.ZipFile(f, "w", zipfile.ZIP_DEFLATED) as zf,
        ):
            zf.writestr(f"АСКУЭ {self.balance_group} {askue_date}.xlsx", register)
            zf.writestr(f"Приложение №9 {self.balance_group}.xlsx", supplement)
