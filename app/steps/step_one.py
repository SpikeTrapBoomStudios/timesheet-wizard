from .base_step import BaseStep
from PySide6.QtWidgets import QFileDialog, QHBoxLayout, QLabel, QPushButton, QLineEdit, QComboBox, QMessageBox
from PySide6.QtWidgets import QSizePolicy
import openpyxl
from resources import resource_path


class StepOne(BaseStep):
    def __init__(self, parent=None, controller=None) -> None:
        super().__init__(parent, controller)

        # Set up <a> link bindings
        self.text_link_binds["download_example_timesheet"] = self._download_example_timesheet

        # Row 0: Title
        row_0: QHBoxLayout = self.ui_rows[0]
        row_0.setContentsMargins(0, 0, 0, 0)
        title_label = QLabel("Source File and Sheet")
        title_label.setStyleSheet("font-size: 18px; font-weight: bold;")
        row_0.addWidget(title_label)

        # Row 1: Source file selection
        row_1: QHBoxLayout = self.ui_rows[1]
        row_1.setContentsMargins(10, 0, 10, 0)

        select_source_label = QLabel("Source File: ")
        row_1.addWidget(select_source_label)

        self.source_file_display = QLineEdit()
        self.source_file_display.setReadOnly(True)
        self.source_file_display.setPlaceholderText("No file selected")
        self.source_file_display.setMinimumWidth(240)
        row_1.addWidget(self.source_file_display)

        select_source_button = QPushButton("Select File")
        select_source_button.clicked.connect(self._select_source_file)
        row_1.addWidget(select_source_button)
        
        # Row 2: Sheet name selection
        row_2: QHBoxLayout = self.ui_rows[2]
        row_2.setContentsMargins(10, 0, 10, 10)

        self.sheet_name_label = QLabel("Source Sheet Name: ")
        row_2.addWidget(self.sheet_name_label)
        self.sheet_name_label.hide()

        self.sheet_name_combobox = QComboBox()
        self.sheet_name_combobox.setMinimumWidth(50)
        self.sheet_name_combobox.addItem("None Selected")
        combobox_sp = QSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Fixed)
        combobox_sp.setRetainSizeWhenHidden(True)
        self.sheet_name_combobox.setSizePolicy(combobox_sp)
        self.sheet_name_combobox.hide()
        row_2.addWidget(self.sheet_name_combobox)

        # Add stretch to prevent vertical expansion
        self.layout().addStretch()
    
    def _download_example_timesheet(self) -> None:
        save_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Example Timesheet As",
            "example_timesheet.xlsx",
            "Excel Files (*.xlsx *.xlsm)"
        )

        if save_path:
            example_path = resource_path("assets", "source_sheet_example.xlsx")
            with open(example_path, 'rb') as src_file:
                with open(save_path, 'wb') as dest_file:
                    dest_file.write(src_file.read())
            QMessageBox.information(self, "Download Complete", f"Example timesheet saved to:\n{save_path}")

    def _select_source_file(self) -> None:
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Source File",
            "",
            "Excel Files (*.xlsx *.xlsm)"
        )

        if file_path:
            self.controller.source_path = file_path
            self.source_file_display.setText(self.controller.source_path)

            self.sheet_name_combobox.clear()
            self.sheet_name_combobox.addItem("None Selected")

            self.sheet_name_label.show()
            self.sheet_name_combobox.show()

            workbook: openpyxl.Workbook = openpyxl.load_workbook(self.controller.source_path, read_only=True)
            for sheet_name in workbook.sheetnames:
                self.sheet_name_combobox.addItem(sheet_name)
            workbook.close()

    def _init_ui_rows(self) -> None:
        self.ui_rows: list[QHBoxLayout] = [QHBoxLayout(), QHBoxLayout(), QHBoxLayout()]
    
    def _init_notes_filename(self) -> str:
        return "step_one.html"
    
    def _next_step(self) -> None:
        if self.controller.source_path and self.sheet_name_combobox.currentText() != "None Selected":
            self.controller.source_sheet_name = self.sheet_name_combobox.currentText()
            self.controller.step_controller.setCurrentIndex(
                self.controller.step_controller.currentIndex() + 1
            )
        elif not self.controller.source_path:
            QMessageBox.warning(self, "Warning", "Please select a source file before proceeding.")
        elif self.sheet_name_combobox.currentText() == "None Selected":
            QMessageBox.warning(self, "Warning", "Please select a source sheet name before proceeding.")