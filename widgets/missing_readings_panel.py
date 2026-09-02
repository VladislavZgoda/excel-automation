from pathlib import Path

from textual import on
from textual.app import ComposeResult
from textual.containers import Container
from textual.reactive import var
from textual.widgets import Button, Static

from askue_etl.common.validation import require_not_none
from askue_etl.readings.missing_readings import prepare_readings
from askue_etl.reports.missing_readings import write_readings_report

from .file_picker import FilePathSelected, FilePicker


class MissingReadingsPanel(Container):
    readings_path: var[Path | None] = var(None)
    report_path: var[Path | None] = var(None)

    def compose(self) -> ComposeResult:
        yield Static("Добавить в Приложение №9 по юр. лицам отсутствующие показания.")
        yield FilePicker('Выберете отчёт "Новые показания" из П2.', "new-readings")
        yield FilePicker("Выберете Приложение №9.", "report")
        yield Button(
            "Добавить показания",
            id="process-data-btn",
            classes="action-btn",
            disabled=True,
        )

    def on_file_path_selected(self, event: FilePathSelected) -> None:
        match event.picker_id:
            case "new-readings":
                self.readings_path = event.file_path
            case "report":
                self.report_path = event.file_path
            case _:
                return

        self._check_and_enable_process_data_btn()

    @on(Button.Pressed, "#process-data-btn")
    def handle_process_data_btn(self) -> None:
        readings_path = require_not_none(self.readings_path, "readings_path")
        report_path = require_not_none(self.report_path, "report_path")

        meter_readings = prepare_readings(readings_path, report_path)
        wb = write_readings_report(report_path, meter_readings)

    def _check_and_enable_process_data_btn(self) -> None:
        if self.readings_path is None or self.report_path is None:
            return

        process_data_btn = self.query_one("#process-data-btn", Button)
        process_data_btn.disabled = False
        process_data_btn.variant = "primary"
