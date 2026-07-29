from typing import ClassVar

from textual.app import App, ComposeResult
from textual.containers import Container, Horizontal
from textual.widgets import (
    Button,
    ContentSwitcher,
    Footer,
    Header,
    Label,
    ListItem,
    ListView,
    Static,
)


class MyApp(App):
    BINDINGS: ClassVar[list[tuple[str, str, str]]] = [
        ("d", "toggle_dark", "Включить/выключить темный режим"),
    ]

    CSS_PATH = "styles.tcss"
    SUB_TITLE = "Трансформация данных из Пирамида 2, Sims для импорта в 1С и отчётов."
    TITLE = "Автоматизация EXCEL"

    def compose(self) -> ComposeResult:
        yield Header()
        with Horizontal():
            yield Sidebar()
            with ContentSwitcher(initial="panel_matritca_readings"):
                yield MatritcaReadingsPanel(id="panel_matritca_readings")
        yield Footer()

    def on_mount(self) -> None:
        self.theme = "monokai"

    def on_list_view_selected(self, event: ListView.Selected) -> None:
        self.query_one(ContentSwitcher).current = event.item.id

    def action_toggle_dark(self) -> None:
        self.theme = "monokai" if self.theme == "atom-one-light" else "atom-one-light"


class Sidebar(ListView):
    def compose(self) -> ComposeResult:
        yield ListItem(Label("Обработать ПУ Матрица"), id="panel_matritca_readings")


class MatritcaReadingsPanel(Container):
    def compose(self) -> ComposeResult:
        yield Static("Трансформировать экспорт из Sims в формат для 1С, Приложение №9.")


if __name__ == "__main__":
    MyApp().run()
