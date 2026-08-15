from textual.app import ComposeResult
from textual.containers import Container
from textual.widgets import Button, Static

from widgets.file_picker import FilePicker


class LegalEntitiesPanel(Container):
    def compose(self) -> ComposeResult:
        yield Static("Перенести данные из экспорта Sims и П2 в отчётные ведомости.")
        yield FilePicker("Выберите файл с Sims показаниями", "sims-readings")
        yield FilePicker('Выберите файл "Новые показания" из П2', "p2-readings")
        yield FilePicker(
            "Выберите файл c текущими показаниями из П2", "p2-current-readings"
        )
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
