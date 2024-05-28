import sys
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, 
                             QComboBox, QMessageBox, QFileDialog, QCheckBox)
from PyQt5.QtGui import QFont
from PyQt5.QtCore import Qt
from threading import Thread
import subprocess
import os

class CustomButton(QPushButton):
    def __init__(self, text, parent=None):
        super().__init__(text, parent)
        self.setStyleSheet("""
            QPushButton {
                background-color: #e63946;
                color: white;
                border: none;
                padding: 10px;
                font-size: 16px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #d62828;
            }
        """)

class ToggleSwitch(QCheckBox):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setText("Dark Mode")
        self.setStyleSheet("""
            QCheckBox {
                font-size: 16px;
                color: white;
                background-color: #2e2e2e;
                border: 1px solid #555555;
                border-radius: 15px;
                padding: 10px;
            }
            QCheckBox::indicator {
                width: 0px;
                height: 0px;
            }
        """)
        self.setFixedSize(120, 40)
        self.setTristate(False)
        self.stateChanged.connect(self.updateAppearance)

    def updateAppearance(self, state):
        if self.isChecked():
            self.setText("Light Mode")
            self.setStyleSheet("""
                QCheckBox {
                    font-size: 16px;
                    color: black;
                    background-color: #f0f0f0;
                    border: 1px solid #cccccc;
                    border-radius: 15px;
                    padding: 10px;
                }
                QCheckBox::indicator {
                    width: 0px;
                    height: 0px;
                }
            """)
        else:
            self.setText("Dark Mode")
            self.setStyleSheet("""
                QCheckBox {
                    font-size: 16px;
                    color: white;
                    background-color: #2e2e2e;
                    border: 1px solid #555555;
                    border-radius: 15px;
                    padding: 10px;
                }
                QCheckBox::indicator {
                    width: 0px;
                    height: 0px;
                }
            """)
        self.parent().toggle_theme()

class SeleniumAutomationApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.excel_path = ''

    def initUI(self):
        self.setWindowTitle('Hyperlink Automation')
        self.setStyleSheet("background-color: #2e2e2e; color: white;")
        layout = QVBoxLayout()

        # Excel file selection
        self.select_file_button = CustomButton('Select Excel File', self)
        self.select_file_button.clicked.connect(self.select_excel_file)
        layout.addWidget(self.select_file_button)
        
        # Excel file path display
        self.excel_path_label = QLabel('No file selected')
        self.excel_path_label.setStyleSheet("font-size: 14px; padding: 5px; border: 1px solid #555555; border-radius: 5px; background-color: #3e3e3e;")
        layout.addWidget(self.excel_path_label)

        # Manufacturer dropdown
        self.manufacturer_dropdown = QComboBox(self)
        self.manufacturer_dropdown.addItems(["Acura", "BMW", "Audi", "Chevrolet", "Chevy"])  # Add all your manufacturers here
        self.manufacturer_dropdown.setStyleSheet("background-color: #3e3e3e; color: white; padding: 5px; border: 1px solid #555555; border-radius: 5px;")
        layout.addWidget(self.manufacturer_dropdown)
        
        # Theme switch section
        theme_switch_section = QHBoxLayout()
        
        self.theme_toggle = ToggleSwitch(self)
        theme_switch_section.addWidget(self.theme_toggle)
        
        layout.addLayout(theme_switch_section)
        
        # Start button
        self.start_button = CustomButton('Start Automation', self)
        self.start_button.clicked.connect(self.start_automation)
        layout.addWidget(self.start_button)
        
        self.setLayout(layout)
        self.resize(400, 300)

    def select_excel_file(self):
        self.excel_path, _ = QFileDialog.getOpenFileName(self, 'Open file', 'C:/Users/', "Excel files (*.xlsx *.xls)")
        if self.excel_path:
            self.excel_path_label.setText(self.excel_path)
        else:
            self.excel_path_label.setText('No file selected')

    def toggle_theme(self):
        if self.theme_toggle.isChecked():
            self.setStyleSheet("background-color: #ffffff; color: black;")
            self.manufacturer_dropdown.setStyleSheet("background-color: #f0f0f0; color: black; padding: 5px; border: 1px solid #cccccc; border-radius: 5px;")
            self.excel_path_label.setStyleSheet("font-size: 14px; padding: 5px; border: 1px solid #cccccc; border-radius: 5px; background-color: #f0f0f0;")
        else:
            self.setStyleSheet("background-color: #2e2e2e; color: white;")
            self.manufacturer_dropdown.setStyleSheet("background-color: #3e3e3e; color: white; padding: 5px; border: 1px solid #555555; border-radius: 5px;")
            self.excel_path_label.setStyleSheet("font-size: 14px; padding: 5px; border: 1px solid #555555; border-radius: 5px; background-color: #3e3e3e;")

    def start_automation(self):
        manufacturer = self.manufacturer_dropdown.currentText()
        confirm_message = f"You have selected {manufacturer}. Are you sure? This can take some time. Please close your Google Chrome before continuing as it will crash the program, continue?"
        confirm = QMessageBox.question(self, 'Confirmation', confirm_message, QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

        if confirm == QMessageBox.Yes and self.excel_path:
            script_path = os.path.join(os.path.dirname(__file__), f"{manufacturer}.py")
            Thread(target=lambda: subprocess.run(["python", script_path, self.excel_path], check=True)).start()
        elif not self.excel_path:
            QMessageBox.warning(self, 'Warning', "Please select an Excel file first.", QMessageBox.Ok)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = SeleniumAutomationApp()
    ex.show()
    sys.exit(app.exec_())
