from PySide6.QtCore import Qt, Signal, QEvent
from PySide6.QtGui import QFont, QShowEvent
from PySide6.QtWidgets import (
    QVBoxLayout, QHBoxLayout, QLabel, QSpinBox,
    QPushButton, QCheckBox, QGridLayout
)
from .base_step import BaseStep
import math
import openpyxl
from openpyxl import Workbook
from resources import resource_path
import os


class AdvancedSettingsStep(BaseStep):
    go_back = Signal()

    def __init__(self, parent=None, controller=None):
        super().__init__(parent, controller)

        # Row 0: Title
        row_0: QHBoxLayout = self.ui_rows[0]
        row_0.setContentsMargins(0, 0, 0, 0)
        title = QLabel("Advanced Settings")
        title.setStyleSheet("font-size: 18px; font-weight: bold;")
        row_0.addWidget(title)

        # Row 1: Settings Grid
        row_1: QHBoxLayout = self.ui_rows[1]
        row_1.setContentsMargins(10, 0, 10, 0)
        
        settings_layout = QGridLayout()
        settings_layout.setSpacing(10)

        # Start Row
        start_label = QLabel("Start Row:")
        self.start_spinbox = QSpinBox()
        self.start_spinbox.setValue(2)
        self.start_spinbox.setMinimum(2)
        self.start_spinbox.setMaximum(1000000000)
        self.start_spinbox.valueChanged.connect(self.update_end_minimum)
        start_reset_btn = QPushButton("Reset")
        start_reset_btn.clicked.connect(self.reset_start_row)

        settings_layout.addWidget(start_label, 0, 0)
        settings_layout.addWidget(self.start_spinbox, 0, 1)
        settings_layout.addWidget(start_reset_btn, 0, 2)

        # End Row
        end_label = QLabel("End Row:")
        self.end_spinbox = QSpinBox()
        max_row = self._get_max_row()
        self.end_spinbox.setValue(max_row)
        self.end_spinbox.setMinimum(2)
        self.end_spinbox.setMaximum(max_row)
        self.end_spinbox.valueChanged.connect(self.update_start_maximum)
        end_reset_btn = QPushButton("Reset")
        end_reset_btn.clicked.connect(self.reset_end_row)
        self.unlimited_checkbox = QCheckBox("Unlimited")
        self.unlimited_checkbox.setChecked(True)
        self.unlimited_checkbox.stateChanged.connect(self.toggle_end_spinbox)

        settings_layout.addWidget(end_label, 1, 0)
        settings_layout.addWidget(self.end_spinbox, 1, 1)
        settings_layout.addWidget(end_reset_btn, 1, 2)
        settings_layout.addWidget(self.unlimited_checkbox, 1, 3)

        row_1.addLayout(settings_layout)

        # Row 2: Inclusive note
        row_2: QHBoxLayout = self.ui_rows[2]
        row_2.setContentsMargins(10, 0, 10, 0)
        inclusive_note = QLabel("Row range is inclusive.")
        inclusive_note.setStyleSheet("font-style: italic; color: #666;")
        row_2.addWidget(inclusive_note)

        # Row 3: Save and Continue Button
        row_3: QHBoxLayout = self.ui_rows[3]
        row_3.setContentsMargins(10, 0, 10, 10)
        row_3.addStretch()
        save_button = QPushButton("Save and Continue")
        save_button.clicked.connect(self._save_and_continue)
        row_3.addWidget(save_button)
        row_3.addStretch()

        # Add stretch to prevent vertical expansion
        self.layout().addStretch()

        self.toggle_end_spinbox()

    def _save_and_continue(self) -> None:
        start_row: int = self.start_spinbox.value()
        end_row: int = self.end_spinbox.value() if not self.unlimited_checkbox.isChecked() else math.inf
        self.controller.row_range = (start_row, end_row)
        self.controller._previous_clicked()

    def _init_ui_rows(self) -> None:
        self.ui_rows: list[QHBoxLayout] = [QHBoxLayout(), QHBoxLayout(), QHBoxLayout(), QHBoxLayout()]

    def reset_start_row(self):
        self.start_spinbox.setValue(2)

    def reset_end_row(self):
        self.unlimited_checkbox.setChecked(True)

    def toggle_end_spinbox(self):
        self.end_spinbox.setEnabled(not self.unlimited_checkbox.isChecked())
    
    def update_end_minimum(self):
        self.end_spinbox.setMinimum(self.start_spinbox.value())
    
    def update_start_maximum(self):
        if not self.unlimited_checkbox.isChecked():
            self.start_spinbox.setMaximum(self.end_spinbox.value())
    
    def _get_max_row(self) -> int:
        try:
            source_path = self.controller.source_path
            source_sheet_name = self.controller.source_sheet_name
            
            if not source_path or not source_sheet_name:
                return 1000000000
            
            src_wb: Workbook = openpyxl.load_workbook(source_path)
            src_ws = src_wb[source_sheet_name]
            max_row = src_ws.max_row
            src_wb.close()
            
            return max_row if max_row and max_row >= 2 else 2
        except Exception:
            return 1000000000
    
    def showEvent(self, event: QShowEvent) -> None:
        super().showEvent(event)
        if self.controller.source_path and self.controller.source_sheet_name:
            max_row = self._get_max_row()
            self.end_spinbox.setMaximum(max_row)
            if self.unlimited_checkbox.isChecked():
                self.end_spinbox.setValue(max_row)
    
    def _build_notes_widget(self):
        self.notes_widget = None