import platform
from pathlib import Path

from textual import on, work
from textual.app import ComposeResult, Widget
from textual.message import Message
from textual.widgets import Button, Label
from textual_fspicker import FileOpen, Filters

FILE_LOCATION = (
    Path.home() / "Desktop" if platform.system() == "Windows" else Path.home()
)

FILE_FILTER = Filters(("XLSX", lambda p: p.suffix.lower() == ".xlsx"))


class FilePathSelected(Message):
    def __init__(self, file_path: Path, picker_id: str) -> None:
        self.file_path = file_path
        self.picker_id = picker_id
        super().__init__()


class FilePicker(Widget):
    def __init__(
        self,
        button_text: str,
        picker_id: str,
        id: str | None = None,
        disabled: bool = False,
    ) -> None:
        self.button_text = button_text
        self.picker_id = picker_id
        super().__init__(id=id, disabled=disabled)

    def compose(self) -> ComposeResult:
        initial_variant = "default" if self.disabled else "warning"
        yield Button(
            self.button_text, variant=initial_variant, id=f"{self.picker_id}-btn"
        )
        yield Label(id=f"{self.picker_id}-label", variant="success")

    def watch_disabled(self, disabled: bool) -> None:
        btn = self.query_one(f"#{self.picker_id}-btn", Button)
        btn.variant = "default" if disabled else "warning"

    @on(Button.Pressed)
    @work
    async def open_file(self, event: Button.Pressed) -> None:
        if event.button.id != f"{self.picker_id}-btn":
            return

        if file_opened := await self.app.push_screen_wait(
            FileOpen(FILE_LOCATION, filters=FILE_FILTER)
        ):
            self.query_one(f"#{self.picker_id}-label", Label).update(file_opened.name)
            self.query_one(f"#{self.picker_id}-btn", Button).variant = "success"
            self.post_message(FilePathSelected(file_opened, self.picker_id))

    def reset(self) -> None:
        self.query_one(f"#{self.picker_id}-label", Label).update("")
