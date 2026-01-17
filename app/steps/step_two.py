from .base_step import BaseStep
from PySide6.QtWidgets import QHBoxLayout, QLabel, QLineEdit, QCalendarWidget, QWidget
from PySide6.QtGui import QTextCharFormat
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QMessageBox


class StepTwo(BaseStep):
    def __init__(self, parent=None, controller=None) -> None:
        super().__init__(parent, controller)

        # Row 0: Title
        row_0: QHBoxLayout = self.ui_rows[0]
        row_0.setContentsMargins(0, 0, 0, 0)
        title_label = QLabel("Start Date")
        title_label.setStyleSheet("font-size: 18px; font-weight: bold;")
        row_0.addWidget(title_label)

        # Row 1: Date input label and line edit
        row_1: QHBoxLayout = self.ui_rows[1]
        row_1.setContentsMargins(10, 0, 10, 0)

        start_date_label = QLabel("Selected START date (mm/dd/yyyy): ")
        row_1.addWidget(start_date_label)

        self.start_date_display = QLineEdit()
        self.start_date_display.setReadOnly(True)
        self.start_date_display.setPlaceholderText("No date selected")
        self.start_date_display.setMaximumWidth(150)
        row_1.addWidget(self.start_date_display)
        row_1.addStretch()

        # Row 2: Calendar widget
        row_2: QHBoxLayout = self.ui_rows[2]
        row_2.setContentsMargins(10, 10, 10, 10)

        self.calendar = QCalendarWidget()
        self.calendar.clicked.connect(self._on_date_selected)
        
        # Remove red coloring from weekends
        weekend_format = QTextCharFormat()
        self.calendar.setWeekdayTextFormat(Qt.Saturday, weekend_format)
        self.calendar.setWeekdayTextFormat(Qt.Sunday, weekend_format)
        
        calendar_container = QWidget()
        calendar_layout = QHBoxLayout(calendar_container)
        calendar_layout.addStretch()
        calendar_layout.addWidget(self.calendar)
        calendar_layout.addStretch()
        
        row_2.addWidget(calendar_container)

        # Add stretch to prevent vertical expansion
        self.layout().addStretch()

    def _on_date_selected(self, date) -> None:
        self.controller.start_date = date.toString("MM/dd/yyyy")
        self.start_date_display.setText(self.controller.start_date)

    def _init_ui_rows(self) -> None:
        self.ui_rows: list[QHBoxLayout] = [QHBoxLayout(), QHBoxLayout(), QHBoxLayout()]

    def _init_notes_filename(self) -> str:
        return "step_two.html"

    def _next_step(self) -> None:
        if self.controller.start_date:
            self.controller.step_controller.setCurrentIndex(
                self.controller.step_controller.currentIndex() + 1
            )
        else:
            QMessageBox.warning(self, "Warning", "Please select a start date before proceeding.")