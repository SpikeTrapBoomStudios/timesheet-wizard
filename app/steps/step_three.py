from .base_step import BaseStep
from PySide6.QtWidgets import QHBoxLayout, QLabel, QLineEdit, QPushButton


class StepThree(BaseStep):
    def __init__(self, parent=None, controller=None) -> None:
        super().__init__(parent, controller)

        # Row 0: Title
        row_0: QHBoxLayout = self.ui_rows[0]
        row_0.setContentsMargins(10, 10, 10, 10)
        title_label = QLabel("Review and Save")
        title_label.setStyleSheet("font-size: 18px; font-weight: bold;")
        row_0.addWidget(title_label)

        # Row 1: Source file
        row_1: QHBoxLayout = self.ui_rows[1]
        row_1.setContentsMargins(10, 10, 10, 10)

        source_file_label = QLabel("Source File: ")
        row_1.addWidget(source_file_label)

        self.source_file_display = QLineEdit()
        self.source_file_display.setReadOnly(True)
        self.source_file_display.setPlaceholderText("No file selected")
        self.source_file_display.setMinimumWidth(240)
        row_1.addWidget(self.source_file_display)

        # Row 2: Sheet name
        row_2: QHBoxLayout = self.ui_rows[2]
        row_2.setContentsMargins(10, 10, 10, 10)

        sheet_name_label = QLabel("Sheet Name: ")
        row_2.addWidget(sheet_name_label)

        self.sheet_name_display = QLineEdit()
        self.sheet_name_display.setReadOnly(True)
        self.sheet_name_display.setPlaceholderText("None Selected")
        self.sheet_name_display.setMinimumWidth(150)
        row_2.addWidget(self.sheet_name_display)

        # Row 3: Start date
        row_3: QHBoxLayout = self.ui_rows[3]
        row_3.setContentsMargins(10, 10, 10, 10)

        start_date_label = QLabel("Selected START date (mm/dd/yyyy): ")
        row_3.addWidget(start_date_label)

        self.start_date_display = QLineEdit()
        self.start_date_display.setReadOnly(True)
        self.start_date_display.setPlaceholderText("No date selected")
        self.start_date_display.setMinimumWidth(150)
        row_3.addWidget(self.start_date_display)

        # Row 4: Advanced settings button
        row_4: QHBoxLayout = self.ui_rows[4]
        row_4.setContentsMargins(10, 10, 10, 10)

        advanced_settings_button: QPushButton = QPushButton("Advanced Settings")
        advanced_settings_button.clicked.connect(self._show_advanced_settings)
        row_4.addWidget(advanced_settings_button)

        self.layout().addStretch()

    def _show_advanced_settings(self) -> None:
        self.controller.step_controller.setCurrentIndex(3)
        self.controller._next_clicked()

    def showEvent(self, event) -> None:
        self.source_file_display.setText(self.controller.source_path or "")
        self.sheet_name_display.setText(self.controller.source_sheet_name or "")
        self.start_date_display.setText(self.controller.start_date or "")
        super().showEvent(event)

    def _init_ui_rows(self) -> None:
        self.ui_rows: list[QHBoxLayout] = [QHBoxLayout(), QHBoxLayout(), QHBoxLayout(), QHBoxLayout(), QHBoxLayout()]

    def _build_notes_widget(self) -> None:
        self.notes_widget = None
