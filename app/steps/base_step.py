from PySide6.QtWidgets import QWidget, QTextBrowser, QVBoxLayout, QHBoxLayout
from resources import read_text_resource

class BaseStep(QWidget):
    def __init__(self, parent=None, controller=None) -> None:
        super().__init__(parent)
        
        self.controller = controller
        self.main_layout = QVBoxLayout(self)
        self.main_layout.setContentsMargins(0, 0, 0, 0)

        self._init_ui_rows()
        self._build_ui_rows()

        self.text_link_binds = {}

        # "Delta One" means a row after all of the rows build from subclasses.
        row_delta_one = QHBoxLayout()
        row_delta_one.setContentsMargins(10, 10, 10, 10)
        self.main_layout.addLayout(row_delta_one)

        self._init_notes_filename()
        self._build_notes_widget()
        if self.notes_widget:
            row_delta_one.addWidget(self.notes_widget)
        
    def _handle_link_click(self, url) -> None:
        if url.scheme() == "method":
            if self.text_link_binds.get(url.path()):
                self.text_link_binds[url.path()]()

    def _init_ui_rows(self) -> None:
        self.ui_rows = []

    def _build_ui_rows(self) -> None:
        for row in self.ui_rows:
            self.main_layout.addLayout(row)
    
    def _init_notes_filename(self) -> str:
        return "none.html"

    def _build_notes_widget(self) -> None:
        self.notes_widget = QTextBrowser()
        self.notes_widget.setReadOnly(True)
        self.notes_widget.setOpenLinks(False)
        self.notes_widget.setMinimumHeight(420)
        self.notes_widget.setStyleSheet("border: none;")
        self.notes_widget.anchorClicked.connect(self._handle_link_click)

        html_content: str = read_text_resource("steps", "notes", self._init_notes_filename())
        self.notes_widget.setHtml(html_content)
    
    def _next_step(self) -> None:
        pass
    
    def _previous_step(self) -> None:
        pass