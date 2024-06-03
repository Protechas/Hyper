import sys
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, 
                             QTreeWidget, QTreeWidgetItem, QMessageBox, QFileDialog, QCheckBox)
from PyQt5.QtGui import QFont
from PyQt5.QtCore import Qt
from threading import Thread
import subprocess
import os

class CustomButton(QPushButton):
    def __init__(self, text, color, parent=None):
        super().__init__(text, parent)
        self.color = color
        self.setStyleSheet(f"""
            QPushButton {{
                background-color: {color};
                color: white;
                border: none;
                padding: 10px;
                font-size: 16px;
                border-radius: 5px;
            }}
            QPushButton:hover {{
                background-color: {self.darken_color(color)};
            }}
        """)

    def darken_color(self, color):
        hex_color = color.lstrip('#')
        rgb = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
        darkened_rgb = tuple(max(0, min(255, int(c * 0.85))) for c in rgb)
        return '#{:02x}{:02x}{:02x}'.format(*darkened_rgb)

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
        self.excel_paths = []

    def initUI(self):
        self.setWindowTitle('Hyperlink Automation')
        self.setStyleSheet("background-color: #2e2e2e; color: white;")
        layout = QVBoxLayout()

        # Excel file selection
        file_selection_layout = QHBoxLayout()
        self.select_file_button = CustomButton('Select Excel Files', '#e63946', self)
        self.select_file_button.clicked.connect(self.select_excel_files)
        file_selection_layout.addWidget(self.select_file_button)

        # Excel file path display
        self.excel_path_label = QLabel('No files selected')
        self.excel_path_label.setStyleSheet("font-size: 14px; padding: 5px; border: 1px solid #555555; border-radius: 5px; background-color: #3e3e3e;")
        file_selection_layout.addWidget(self.excel_path_label)

        self.activate_full_automation_button = CustomButton('Activate Full Automation', '#e3b505', self)
        self.activate_full_automation_button.clicked.connect(self.activate_full_automation)
        file_selection_layout.addWidget(self.activate_full_automation_button)

        layout.addLayout(file_selection_layout)

        # Manufacturer tree widget with checkboxes
        manufacturer_selection_layout = QHBoxLayout()
        self.manufacturer_tree = QTreeWidget(self)
        self.manufacturer_tree.setHeaderHidden(True)
        self.manufacturer_tree.setStyleSheet("background-color: #3e3e3e; color: white; border: 1px solid #555555; border-radius: 5px;")
        manufacturers = ["Acura", "Alfa Romeo", "Audi", "BMW", "Brightdrop", "Buick", "Cadillac", "Chevrolet", "Dodge", 
                         "Fiat", "Ford", "Genesis", "GMC", "Honda", "Hyundai", "Infiniti", "Jaguar", "Kia", "Lexus", 
                         "Mazda", "Mini", "Mitsubishi", "Nissan", "Porsche", "Ram", "Rolls Royce", "Subaru", "Toyota", 
                         "Volkswagen", "Volvo"]
        for manufacturer in manufacturers:
            item = QTreeWidgetItem(self.manufacturer_tree)
            item.setText(0, manufacturer)
            item.setCheckState(0, Qt.Unchecked)
        manufacturer_selection_layout.addWidget(self.manufacturer_tree)

        # Select All button
        self.select_all_button = CustomButton('Select All', '#e3b505', self)
        self.select_all_button.clicked.connect(self.select_all)
        manufacturer_selection_layout.addWidget(self.select_all_button)

        layout.addLayout(manufacturer_selection_layout)

        # Theme switch section
        theme_switch_section = QHBoxLayout()
        
        self.theme_toggle = ToggleSwitch(self)
        theme_switch_section.addWidget(self.theme_toggle)
        
        layout.addLayout(theme_switch_section)
        
        # Start button
        self.start_button = CustomButton('Start Automation', '#e63946', self)
        self.start_button.clicked.connect(self.start_automation)
        layout.addWidget(self.start_button)
        
        self.setLayout(layout)
        self.resize(600, 400)

    def select_excel_files(self):
        self.excel_paths, _ = QFileDialog.getOpenFileNames(self, 'Open files', 'C:/Users/', "Excel files (*.xlsx *.xls)")
        if self.excel_paths:
            self.excel_path_label.setText("\n".join([f"{i + 1}. {os.path.basename(path)}" for i, path in enumerate(self.excel_paths)]))
        else:
            self.excel_path_label.setText('No files selected')

    def toggle_theme(self):
        if self.theme_toggle.isChecked():
            self.setStyleSheet("background-color: #ffffff; color: black;")
            self.manufacturer_tree.setStyleSheet("background-color: #f0f0f0; color: black; border: 1px solid #cccccc; border-radius: 5px;")
            self.excel_path_label.setStyleSheet("font-size: 14px; padding: 5px; border: 1px solid #cccccc; border-radius: 5px; background-color: #f0f0f0;")
        else:
            self.setStyleSheet("background-color: #2e2e2e; color: white;")
            self.manufacturer_tree.setStyleSheet("background-color: #3e3e3e; color: white; border: 1px solid #555555; border-radius: 5px;")
            self.excel_path_label.setStyleSheet("font-size: 14px; padding: 5px; border: 1px solid #555555; border-radius: 5px; background-color: #3e3e3e;")

    def start_automation(self):
        selected_manufacturers = []
        for i in range(self.manufacturer_tree.topLevelItemCount()):
            item = self.manufacturer_tree.topLevelItem(i)
            if item.checkState(0) == Qt.Checked:
                selected_manufacturers.append(item.text(0))

        if self.excel_paths and selected_manufacturers:
            confirm_message = "You have selected the following manufacturers and Excel files:\n\n"
            for i, manufacturer in enumerate(selected_manufacturers):
                confirm_message += f"{i + 1}. {manufacturer}\n"
            confirm_message += "\nPlease ensure the order is correct. Continue?"

            confirm = QMessageBox.question(self, 'Confirmation', confirm_message, QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

            if confirm == QMessageBox.Yes:
                for excel_path in self.excel_paths:
                    for manufacturer in selected_manufacturers:
                        script_path = os.path.join(os.path.dirname(__file__), f"{manufacturer}.py")
                        Thread(target=lambda: subprocess.run(["python", script_path, excel_path], check=True)).start()
            else:
                QMessageBox.warning(self, 'Warning', "Automation process canceled.", QMessageBox.Ok)
        else:
            QMessageBox.warning(self, 'Warning', "Please select Excel files and manufacturers first.", QMessageBox.Ok)

    def activate_full_automation(self):
        if not self.excel_paths:
            QMessageBox.warning(self, 'Warning', "Please select Excel files first.", QMessageBox.Ok)
            return

        selected_manufacturers = []
        for i in range(self.manufacturer_tree.topLevelItemCount()):
            item = self.manufacturer_tree.topLevelItem(i)
            if item.checkState(0) == Qt.Checked:
                selected_manufacturers.append(item.text(0))
        
        if not selected_manufacturers:
            QMessageBox.warning(self, 'Warning', "Please select manufacturers first.", QMessageBox.Ok)
            return
        
        confirm_message = ("WARNING!!! This will take a LONG time to complete, ETA N/A as of yet. "
                           "Please prepare to not touch your computer for a period of time. "
                           "Also ensure that every Excel file is put in the proper order or this will mess all the Longsheets up. Continue?")
        confirm = QMessageBox.question(self, 'Confirmation', confirm_message, QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

        if confirm == QMessageBox.Yes:
            for excel_path in self.excel_paths:
                for manufacturer in selected_manufacturers:
                    script_path = os.path.join(os.path.dirname(__file__), f"{manufacturer}.py")
                    Thread(target=lambda: subprocess.run(["python", script_path, excel_path], check=True)).start()

    def select_all(self):
        select_all_checked = True
        for i in range(self.manufacturer_tree.topLevelItemCount()):
            item = self.manufacturer_tree.topLevelItem(i)
            if item.checkState(0) != Qt.Checked:
                select_all_checked = False
                break
        
        for i in range(self.manufacturer_tree.topLevelItemCount()):
            item = self.manufacturer_tree.topLevelItem(i)
            item.setCheckState(0, Qt.Checked if not select_all_checked else Qt.Unchecked)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = SeleniumAutomationApp()
    ex.show()
    sys.exit(app.exec_())
