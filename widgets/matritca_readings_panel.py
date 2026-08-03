from pathlib import Path
from typing import cast, get_args

from textual import on
from textual.app import ComposeResult
from textual.containers import Container
from textual.reactive import var
from textual.widgets import (
    Button,
    Select,
    Static,
)

from filters.matritca_readings import BalanceGroupType, filterReadings
from widgets.file_picker import FilePathSelected, FilePicker
from write_to_excel.matritca_readings import create_wb_reports


class MatritcaReadingsPanel(Container):
    readings_path: var[Path | None] = var(None)
    balance_group: var[BalanceGroupType | None] = var(None)

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
            "Обработать данные", id="process-data-btn", variant="default", disabled=True
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

        filtered_readings = filterReadings(self.readings_path, balance_group)
        wb_reports = create_wb_reports(filtered_readings, balance_group)

    def _check_and_enable_process_data_btn(self) -> None:
        if self.readings_path is None:
            return

        process_data_btn = self.query_one("#process-data-btn", Button)
        process_data_btn.disabled = False
        process_data_btn.variant = "primary"
