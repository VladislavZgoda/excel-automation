import platform
from pathlib import Path

from textual import on, work
from textual.app import ComposeResult
from textual.message import Message
from textual.widget import Widget
from textual.widgets import Button, Label
from textual_fspicker import SelectDirectory

FOLDER_LOCATION = (
    Path.home() / "Desktop" if platform.system() == "Windows" else Path.home()
)


class FolderPathSelected(Message):
    def __init__(self, folder_path: Path, picker_id: str) -> None:
        self.folder_path = folder_path
        self.picker_id = picker_id
        super().__init__()


class FolderPicker(Widget):
    def __init__(
        self,
        button_text: str,
        picker_id: str,
    ) -> None:
        self.button_text = button_text
        self.picker_id = picker_id
        super().__init__()

    def compose(self) -> ComposeResult:
        yield Button(self.button_text, variant="warning", id=f"{self.picker_id}-btn")
        yield Label(id=f"{self.picker_id}-label", variant="success")

    @on(Button.Pressed)
    @work
    async def open_folder(self, event: Button.Pressed) -> None:
        if event.button.id != f"{self.picker_id}-btn":
            return

        if folder_opened := await self.app.push_screen_wait(
            SelectDirectory(FOLDER_LOCATION)
        ):
            self.query_one(f"#{self.picker_id}-label", Label).update(str(folder_opened))
            self.query_one(f"#{self.picker_id}-btn", Button).variant = "success"
            self.post_message(FolderPathSelected(folder_opened, self.picker_id))

    def reset(self) -> None:
        self.query_one(f"#{self.picker_id}-label", Label).update("")
        self.query_one(f"#{self.picker_id}-btn", Button).variant = "warning"
