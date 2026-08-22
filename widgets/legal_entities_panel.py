from pathlib import Path

from textual import on, work
from textual.app import ComposeResult
from textual.containers import Container
from textual.reactive import var
from textual.widgets import Button, Static

from askue_etl.readings.legal_readings import prepare_readings
from askue_etl.reports.legal_readings import write_readings_reports
from widgets.file_picker import FilePathSelected, FilePicker
from widgets.folder_picker import FolderPathSelected, FolderPicker


class LegalEntitiesPanel(Container):
    sims_readings_path: var[Path | None] = var(None)
    p2_readings_path: var[Path | None] = var(None)
    p2_current_readings: var[Path | None] = var(None)
    template_folder_path: var[Path | None] = var(None)
    reports_folder_path: var[Path | None] = var(None)

    def compose(self) -> ComposeResult:
        yield Static("Перенести данные из экспорта Sims и П2 в отчётные ведомости.")
        yield FilePicker("Выберите файл с Sims показаниями", "sims-readings")
        yield FilePicker('Выберите файл "Новые показания" из П2', "p2-readings")
        yield FilePicker(
            "Выберите файл c текущими показаниями из П2", "p2-current-readings"
        )
        yield FolderPicker("Выберите папку с шаблонами ведомостей", "template-folder")
        yield FolderPicker("Выберите папку для сохранения ведомостей", "reports-folder")
        yield Button(
            "Сформировать ведомости",
            id="process-data-btn",
            classes="action-btn",
            disabled=True,
        )

    def on_file_path_selected(self, event: FilePathSelected) -> None:
        match event.picker_id:
            case "sims-readings":
                self.sims_readings_path = event.file_path
            case "p2-readings":
                self.p2_readings_path = event.file_path
            case "p2-current-readings":
                self.p2_current_readings = event.file_path
            case _:
                return

        self._check_and_enable_process_data_btn()

    def on_folder_path_selected(self, event: FolderPathSelected) -> None:
        match event.picker_id:
            case "template-folder":
                self.template_folder_path = event.folder_path
            case "reports-folder":
                self.reports_folder_path = event.folder_path
            case _:
                return

        self._check_and_enable_process_data_btn()

    @on(Button.Pressed, "#process-data-btn")
    @work(thread=True)
    def handle_process_data_btn(self) -> None:
        if self.sims_readings_path is None:
            raise ValueError("Требуется объект Path, получен None.")
        if self.p2_readings_path is None:
            raise ValueError("Требуется объект Path, получен None.")
        if self.p2_current_readings is None:
            raise ValueError("Требуется объект Path, получен None.")
        if self.template_folder_path is None:
            raise ValueError("Требуется объект Path, получен None.")
        if self.reports_folder_path is None:
            raise ValueError("Требуется объект Path, получен None.")

        self.app.call_from_thread(self._on_process_data_start)

        meter_readings = prepare_readings(
            self.sims_readings_path, self.p2_readings_path, self.p2_current_readings
        )

        write_readings_reports(
            meter_readings, self.template_folder_path, self.reports_folder_path
        )

        self.app.call_from_thread(self._on_process_data_done)
        self.notify("Ведомости созданы.", timeout=10)

    def _check_and_enable_process_data_btn(self) -> None:
        all_paths_selected = all(
            path is not None
            for path in (
                self.sims_readings_path,
                self.p2_readings_path,
                self.p2_current_readings,
                self.template_folder_path,
                self.reports_folder_path,
            )
        )

        if not all_paths_selected:
            return

        process_data_btn = self.query_one("#process-data-btn", Button)
        process_data_btn.disabled = False
        process_data_btn.variant = "primary"

    def _on_process_data_start(self) -> None:
        self.query_one("#process-data-btn", Button).loading = True

    def _on_process_data_done(self) -> None:
        self.sims_readings_path = None
        self.p2_readings_path = None
        self.p2_current_readings = None
        self.template_folder_path = None
        self.reports_folder_path = None

        for picker in self.query(FilePicker):
            picker.reset()

        for picker in self.query(FolderPicker):
            picker.reset()

        process_data_btn = self.query_one("#process-data-btn", Button)
        process_data_btn.loading = False
        process_data_btn.disabled = True
        process_data_btn.variant = "default"
