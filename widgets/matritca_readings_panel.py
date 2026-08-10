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
from textual.widgets import Button, Checkbox, Select, Static
from textual_fspicker import FileSave, Filters

from askue_etl.readings.matritca_readings import BalanceGroupType, prepare_readings
from askue_etl.reports.matritca_readings import write_readings_reports
from widgets.file_picker import FILE_LOCATION, FilePathSelected, FilePicker


class MatritcaReadingsPanel(Container):
    readings_path: var[Path | None] = var(None)
    list_1c_path: var[Path | None] = var(None)
    balance_group: var[BalanceGroupType | None] = var(None)
    readings_reports: var[tuple[BytesIO, BytesIO] | None] = var(None)

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
        yield Checkbox("Изменить однозонные ПУ")
        yield FilePicker(
            "Выберите файл с выгрузкой 1С",
            picker_id="list-1C",
            id="list-1c-picker",
            disabled=True,
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
            case "list-1C":
                self.list_1c_path = event.file_path
            case _:
                return

        self._reset_readings_reports()
        self._check_and_enable_process_data_btn()

    def on_checkbox_changed(self, event: Checkbox.Changed) -> None:
        list_1c_picker = self.query_one("#list-1c-picker", FilePicker)
        list_1c_picker.disabled = not event.value

        if not event.value:
            self.list_1c_path = None
            list_1c_picker.reset()

    def on_select_changed(self, event: Select.Changed) -> None:
        self.balance_group = cast(BalanceGroupType, event.value)

        checkbox = self.query_one(Checkbox)
        is_private = self.balance_group == "Быт"

        checkbox.disabled = not is_private

        if not is_private:
            checkbox.value = False

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
        prepared_readings = prepare_readings(
            self.readings_path, balance_group, self.list_1c_path
        )
        readings_reports = write_readings_reports(prepared_readings, balance_group)

        self.app.call_from_thread(self._on_process_data_done, readings_reports)
        self.notify("Данные обработаны.", timeout=10)

    @on(Button.Pressed, "#save-file-btn")
    @work
    async def handle_save_btn(self) -> None:
        if self.readings_reports is None:
            raise ValueError("Требуется объект tuple [BytesIO, BytesIO], получен None.")
        if save_path := await self.app.push_screen_wait(
            FileSave(
                FILE_LOCATION,
                filters=Filters(("ZIP", lambda p: p.suffix.lower() == ".zip")),
            ),
        ):
            register_buf, supplement_nine_buf = self.readings_reports

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

    def _on_process_data_done(self, readings_reports: tuple[BytesIO, BytesIO]) -> None:
        self.readings_reports = readings_reports
        self.query_one("#process-data-btn", Button).loading = False
        save_file_btn = self.query_one("#save-file-btn", Button)
        save_file_btn.disabled = False
        save_file_btn.variant = "success"

    def _reset_readings_reports(self) -> None:
        self.readings_reports = None

        save_file_btn = self.query_one("#save-file-btn", Button)
        save_file_btn.disabled = True
        save_file_btn.variant = "default"

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
