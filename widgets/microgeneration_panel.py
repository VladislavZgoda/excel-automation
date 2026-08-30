from pathlib import Path

from openpyxl import Workbook
from textual import on, work
from textual.app import ComposeResult
from textual.containers import Container
from textual.reactive import var
from textual.widgets import Button, Static
from textual_fspicker import FileSave, Filters

from askue_etl.common.validation import require_not_none
from askue_etl.readings.microgeneration import prepare_readings
from askue_etl.reports.microgeneration import write_readings_report

from .file_picker import FILE_LOCATION, FilePathSelected, FilePicker


class MicrogenerationPanel(Container):
    readings_path: var[Path | None] = var(None)
    template_path: var[Path | None] = var(None)
    workbook: var[Workbook | None] = var(None)

    def compose(self) -> ComposeResult:
        yield Static("Заполнить шаблон данными по микрогенерации.")
        yield FilePicker("Выберите файл с Sims показаниями", "readings")
        yield FilePicker("Выберите файл с шаблоном", "template")
        yield Button(
            "Обработать данные",
            id="process-data-btn",
            classes="action-btn",
            disabled=True,
        )
        yield Button(
            "Сохранить ведомость",
            id="save-file-btn",
            classes="action-btn",
            disabled=True,
        )

    def on_file_path_selected(self, event: FilePathSelected) -> None:
        match event.picker_id:
            case "readings":
                self.readings_path = event.file_path
            case "template":
                self.template_path = event.file_path
            case _:
                return

        self._reset_readings_reports()
        self._check_and_enable_process_data_btn()

    @on(Button.Pressed, "#process-data-btn")
    @work(thread=True)
    def handle_process_data_btn(self) -> None:
        readings_path = require_not_none(self.readings_path, "readings_path")
        template_path = require_not_none(self.template_path, "template_path")

        self.app.call_from_thread(self._on_process_data_start)

        meter_readings = prepare_readings(readings_path, template_path)
        workbook = self.workbook = write_readings_report(template_path, meter_readings)

        self.app.call_from_thread(self._on_process_data_done, workbook)
        self.notify("Ведомость сформирована.")

    @on(Button.Pressed, "#save-file-btn")
    @work
    async def handle_save_btn(self) -> None:
        wb = require_not_none(self.workbook, "workbook")

        if save_path := await self.app.push_screen_wait(
            FileSave(
                FILE_LOCATION,
                filters=Filters(("XLSX", lambda p: p.suffix.lower() == ".xlsx")),
            ),
        ):
            wb.save(save_path.with_suffix(".xlsx"))
            self._reset_state()
            self.notify("Файл сохранён.")

    def _on_process_data_start(self) -> None:
        self.query_one("#process-data-btn", Button).loading = True

    def _on_process_data_done(self, workbook: Workbook) -> None:
        self.workbook = workbook

        save_file_btn = self.query_one("#save-file-btn", Button)
        save_file_btn.disabled = False
        save_file_btn.variant = "success"
        self.query_one("#process-data-btn", Button).loading = False

    def _check_and_enable_process_data_btn(self) -> None:
        if self.readings_path is None or self.template_path is None:
            return

        process_data_btn = self.query_one("#process-data-btn", Button)
        process_data_btn.disabled = False
        process_data_btn.variant = "primary"

    def _reset_readings_reports(self) -> None:
        self.workbook = None

        save_file_btn = self.query_one("#save-file-btn", Button)
        save_file_btn.disabled = True
        save_file_btn.variant = "default"

    def _reset_state(self) -> None:
        self.readings_path = None
        self.template_path = None
        self.workbook = None

        for picker in self.query(FilePicker):
            picker.reset()

        process_data_btn = self.query_one("#process-data-btn", Button)
        process_data_btn.disabled = True
        process_data_btn.variant = "default"

        save_file_btn = self.query_one("#save-file-btn", Button)
        save_file_btn.disabled = True
        save_file_btn.variant = "default"
