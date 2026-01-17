import sys
from PySide6.QtGui import QAction, Qt, QIcon
from PySide6.QtWidgets import (
    QApplication, QLabel, QMainWindow, QMenuBar, QMenu, QPushButton,
    QHBoxLayout, QStatusBar, QWidget, QStackedWidget, QFileDialog
)
from steps.step_one import StepOne
from steps.step_two import StepTwo
from steps.step_three import StepThree
from steps.advanced_settings import AdvancedSettingsStep
from timesheet_creator import create_timesheets
import math
from resources import resource_path
from PySide6.QtWidgets import QMessageBox
import os
import platform


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Timesheet Wizard")
        self.setWindowIcon(QIcon(str(resource_path("assets", "file_icon", "icon.png"))))

        self.setMinimumSize(700, 380)
        window_width = 800
        window_height = 600
        
        screen_width: int = self.screen().size().width()
        screen_height: int = self.screen().size().height()

        x: int = (screen_width - window_width) // 2
        y: int = (screen_height - window_height) // 2
        
        self.setGeometry(x, y, window_width, window_height)

        self.source_path: str = ""
        self.source_sheet_name: str = ""
        self.start_date: str = ""
        self.row_range: tuple[int, int] = (2, math.inf)

        self.root = QWidget()
        self.root_layout = QHBoxLayout(self.root)
        self.root_layout.setContentsMargins(0, 0, 0, 0)
        self.setCentralWidget(self.root)
        self._build_content_widget()
        self._build_menu_bar()
        self._build_status_bar()
    
    def _build_menu_bar(self) -> None:
        menu_bar: QMenuBar = self.menuBar()

        file_menu: QMenu = menu_bar.addMenu("&File")

        quit_action = QAction("Quit", self)
        quit_action.setShortcut("Ctrl+Q")
        quit_action.triggered.connect(self.close)

        file_menu.addAction(quit_action)
    
    # Content widget is where each step will have its widgets in
    def _build_content_widget(self) -> None:
        content_widget = QWidget()
        layout = QHBoxLayout(content_widget)
        layout.setContentsMargins(10, 10, 10, 10)

        self.step_controller = QStackedWidget()

        for Step in (StepOne, StepTwo, StepThree, AdvancedSettingsStep):
            step_instance = Step(controller=self)
            self.step_controller.addWidget(step_instance)

        layout.addWidget(self.step_controller, alignment=Qt.AlignTop)
        self.step_controller.setCurrentIndex(0)

        self.root_layout.addWidget(content_widget)
    
    def _toggle_buttons_based_on_step(self) -> None:
        self.status_label.setText(f"Step {self.step_controller.currentIndex() + 1} of {self.step_controller.count() - 1}")
        if self.step_controller.currentIndex() > 0:
                self.prev_button.show()
        else:
            self.prev_button.hide()
        if self.step_controller.currentIndex() == 2:
            # Review and Save step
            self.finish_button.show()
            self.next_button.hide()
        else:
            self.finish_button.hide()
            self.next_button.show()
        if self.step_controller.currentIndex() == 3:
            # Advanced settings step
            self.finish_button.hide()
            self.prev_button.hide()
            self.next_button.hide()

    def _next_clicked(self) -> None:
        current_step: QWidget = self.step_controller.widget(self.step_controller.currentIndex())
        if current_step:
            current_step._next_step()
            self._toggle_buttons_based_on_step()
    
    def _previous_clicked(self) -> None:
        current_step_index: int = self.step_controller.currentIndex()
        if current_step_index > 0:
            self.step_controller.setCurrentIndex(current_step_index - 1)
            self._toggle_buttons_based_on_step()

    def _save_and_finish(self) -> None:
        current_step_index: int = self.step_controller.currentIndex()
        if current_step_index == 2:
            user_save_path, _ = QFileDialog.getSaveFileName(
                self,
                "Save Generated Timesheets As",
                "generated_timesheets.xlsx",
                "Excel Files (*.xlsx);;All Files (*)"
            )
            if user_save_path.strip() == "":
                return

            success = create_timesheets(
                source_file_path = self.source_path,
                source_sheet_name = self.source_sheet_name,
                template_file_path = resource_path("assets", "timesheet_template.xlsx"),
                output_file_path=user_save_path,
                start_date=self.start_date,
                row_range=self.row_range
            )

            if success:
                to_pdf_success = self._excel_to_pdf(
                    excel_file_path=user_save_path,
                    output_pdf_path=os.path.splitext(user_save_path)[0] + ".pdf"
                )
                if not to_pdf_success:
                    QMessageBox.warning(
                        self,
                        "PDF Conversion Failed",
                        "Timesheets were created successfully, but converting to PDF failed."
                    )

                self.close()
            else:
                QMessageBox.critical(self, "Error", "An error occurred while creating the timesheets.")
    
    # Risky method, that's why there are so many try/excepts. Windows-only.
    def _excel_to_pdf(self, excel_file_path: str, output_pdf_path: str) -> bool:
        if not platform.system() == "Windows":
            QMessageBox.warning(self, "Error converting to PDF", "Unable to convert the timesheets to PDF. Note: This is a windows only feature currently.")
            return False
        try:
            import win32com.client
        except:
            print("Could not import win32com")
            return False
        try:
            excel_path = os.path.abspath(excel_file_path)
            pdf_path = os.path.abspath(output_pdf_path)
            
            if not os.path.exists(excel_path):
                QMessageBox.critical(self, "Error", f"Excel file not found: {excel_path}")
                return False
            
            try:
                excel = win32com.client.Dispatch("Excel.Application")
            except Exception as e:
                QMessageBox.critical(
                    self, 
                    "Excel Not Available", 
                    "Failed to open Microsoft Excel.\n\n"
                    "Please ensure that:\n"
                    "• Microsoft Excel is installed on your computer\n"
                    "• Excel is up to date\n"
                    "• Excel is properly registered with Windows\n\n"
                    f"Technical details: {str(e)}"
                )
                return False
            
            excel.Visible = False
            excel.DisplayAlerts = False
            
            try:
                wb = excel.Workbooks.Open(excel_path)
                wb.ExportAsFixedFormat(0, pdf_path)
                wb.Close(False)
                excel.Quit()
                return True
            except Exception as e:
                QMessageBox.critical(
                    self, 
                    "PDF Conversion Error", 
                    f"Failed to convert Excel to PDF.\n\n"
                    f"This may occur if:\n"
                    f"• The Excel file is corrupted\n"
                    f"• The output path is invalid\n"
                    f"• Excel encountered an internal error\n\n"
                    f"Technical details: {str(e)}"
                )
                excel.Quit()
                return False
            
        except Exception as e:
            QMessageBox.critical(self, "Unexpected Error", f"An unexpected error occurred:\n{str(e)}")
            try:
                if 'excel' in locals():
                    excel.Quit()
            except:
                pass
            return False
    
    # Status bar houses previous/next buttons as well as a status label
    def _build_status_bar(self) -> None:
        status_bar: QStatusBar = self.statusBar()
        status_bar.setSizeGripEnabled(False)

        # Create a container widget for right-aligned buttons
        right_widget = QWidget()
        right_layout = QHBoxLayout(right_widget)

        self.prev_button = QPushButton("Previous")
        self.prev_button.clicked.connect(self._previous_clicked)

        self.next_button = QPushButton("Next")
        self.next_button.clicked.connect(self._next_clicked)

        self.finish_button = QPushButton("Finish")
        self.finish_button.clicked.connect(self._save_and_finish)

        right_layout.addWidget(self.prev_button)
        self.prev_button.hide()
        right_layout.addWidget(self.next_button)
        right_layout.addWidget(self.finish_button)
        self.finish_button.hide()

        # Create a container widget for left-aligned status label
        left_widget = QWidget()
        left_layout = QHBoxLayout(left_widget)
        left_layout.setContentsMargins(10, 0, 0, 0)

        self.status_label = QLabel("Step 0 of 3")
        self.status_label.setStyleSheet("font-size: 16px;")
        left_layout.addWidget(self.status_label)

        status_bar.addPermanentWidget(left_widget, 1)
        status_bar.addPermanentWidget(right_widget)

        self._toggle_buttons_based_on_step()

if __name__ == "__main__":
    import ctypes
    try:
        myappid = 'stbstudios.timesheetwizard.app.1.0'
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
    except:
        pass

    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())