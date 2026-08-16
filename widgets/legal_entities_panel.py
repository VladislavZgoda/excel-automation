from pathlib import Path

from textual.app import ComposeResult
from textual.containers import Container
from textual.reactive import var
from textual.widgets import Button, Static

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
        yield Button(
            "Сохранить ведомости",
            id="save-file-btn",
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

    def on_folder_path_selected(self, event: FolderPathSelected) -> None:
        match event.picker_id:
            case "template-folder":
                self.template_folder_path = event.folder_path
            case "reports-folder":
                self.reports_folder_path = event.folder_path
            case _:
                return
