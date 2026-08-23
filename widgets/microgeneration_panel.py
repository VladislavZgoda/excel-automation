from pathlib import Path

from textual import on
from textual.app import ComposeResult
from textual.containers import Container
from textual.reactive import var
from textual.widgets import Button, Static

from askue_etl.common.validation import require_not_none
from widgets.file_picker import FilePathSelected, FilePicker


class MicrogenerationPanel(Container):
    readings_path: var[Path | None] = var(None)
    template_path: var[Path | None] = var(None)

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

        self._check_and_enable_process_data_btn()

    @on(Button.Pressed, "#process-data-btn")
    def handle_process_data_btn(self) -> None:
        readings_path = require_not_none(self.readings_path, "readings_path")
        template_path = require_not_none(self.template_path, "template_path")

    @on(Button.Pressed, "#save-file-btn")
    def handle_save_btn(self) -> None:
        pass

    def _check_and_enable_process_data_btn(self) -> None:
        if self.readings_path is None or self.template_path is None:
            return

        process_data_btn = self.query_one("#process-data-btn", Button)
        process_data_btn.disabled = False
        process_data_btn.variant = "primary"
