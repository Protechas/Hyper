import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QComboBox, QMessageBox, QFileDialog
from threading import Thread
import subprocess
import os

class SeleniumAutomationApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.excel_path = ''

    def initUI(self):
        self.setWindowTitle('Hyperlink Automation')
        layout = QVBoxLayout()

        # Manufacturer dropdown
        self.manufacturer_dropdown = QComboBox(self)
        self.manufacturer_dropdown.addItems(["Acura", "Audi", "BMW", "Chevrolet"])  # Add all your manufacturers here
        layout.addWidget(self.manufacturer_dropdown)

        # Start button
        self.start_button = QPushButton('Start Automation', self)
        self.start_button.clicked.connect(self.start_automation)
        layout.addWidget(self.start_button)

        # Excel file selection
        self.select_file_button = QPushButton('Select Excel File', self)
        self.select_file_button.clicked.connect(self.select_excel_file)
        layout.addWidget(self.select_file_button)

        self.setLayout(layout)
        self.resize(400, 200)

    def select_excel_file(self):
        self.excel_path, _ = QFileDialog.getOpenFileName(self, 'Open file', 'C:/Users/', "Excel files (*.xlsx *.xls)")
        if self.excel_path:
            print(f"Selected file: {self.excel_path}")

    def start_automation(self):
        manufacturer = self.manufacturer_dropdown.currentText()
        confirm_message = f"You have selected {manufacturer}. Are you sure? This can take some time, continue?"
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