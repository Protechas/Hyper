import sys
from PyQt5.QtWidgets import (QApplication, QDialog, QPlainTextEdit, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, 
                             QTreeWidget, QTreeWidgetItem, QMessageBox, QFileDialog, QCheckBox)
from PyQt5.QtGui import QFont
from PyQt5.QtCore import Qt,pyqtSignal,QThread
from threading import Thread
import subprocess
from time import sleep
import os

class WorkerThread(QThread):
    output_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(str)

    def __init__(self, command, manufacturer, parent=None):
        super(WorkerThread, self).__init__(parent)
        self.command = command
        self.manufacturer = manufacturer

    def run(self):
        # Set up the environment to disable buffering
        env = os.environ.copy()
        env["PYTHONUNBUFFERED"] = "1"

        # Run the subprocess with unbuffered output
        process = subprocess.Popen(self.command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, bufsize=1, universal_newlines=True, env=env)
        
        # Read stdout line by line and emit each line
        for stdout_line in iter(process.stdout.readline, ""):
            self.output_signal.emit(stdout_line.strip())
        process.stdout.close()

        # Wait for the process to finish and emit any error lines
        process.wait()
        if process.returncode != 0:
            for stderr_line in iter(process.stderr.readline, ""):
                self.output_signal.emit(stderr_line.strip())
        
        process.stderr.close()
        self.finished_signal.emit(self.manufacturer)  # Emit when a manufacturer is finished

class TerminalDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Terminal Output")
        self.setGeometry(100, 100, 600, 400)

        self.layout = QVBoxLayout()
        self.terminal_output = QPlainTextEdit()
        self.terminal_output.setReadOnly(True)
        self.layout.addWidget(self.terminal_output)

        self.setLayout(self.layout)

    def append_output(self, text):
        self.terminal_output.appendPlainText(text)
        self.terminal_output.ensureCursorVisible()


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
        self.manufacturer_links = {
            "Acura": "https://calibercollision.sharepoint.com/:f:/g/enterpriseprojects/VehicleServiceInformation/ElvYzTNE4KRAoLa1yoXXfG8Bts2R8_lFxBc7fm3XcCxyZg?e=Q2evq5",
            "Alfa Romeo": "https://calibercollision.sharepoint.com/:f:/g/enterpriseprojects/VehicleServiceInformation/EhNnruVI995AiljPoEYp-BMB6XOIhOC9JX4XN6FDXtBbbQ?e=StAVfM",
            "Audi": "https://calibercollision.sharepoint.com/:f:/g/enterpriseprojects/VehicleServiceInformation/EhCY-x-ICyJIof9fIcOCI4oBywQvzN2daFMHuS1c8hNhAA?e=HkCg2Y",
            "BMW": "https://calibercollision.sharepoint.com/:f:/g/enterpriseprojects/VehicleServiceInformation/ErWRCoawOG1IjuclCmUAbz0BZJbsYESgRYnyGZ32Cln_fA?e=vWmHVa",
            "Brightdrop": "https://calibercollision.sharepoint.com/:f:/g/enterpriseprojects/VehicleServiceInformation/Eozt6QvbAgpBpvW9FJwf1JsBeXDmZ6zAhEuEKWjjO-BTAQ?e=Z9QhK4",
            "Buick": "https://calibercollision.sharepoint.com/:f:/g/enterpriseprojects/VehicleServiceInformation/ElnBWMWkI-tAkghaeYaQV4UB9ajyxUla_T70kTMGD9W-7Q?e=I8FlfX",
            "Cadillac": "https://calibercollision.sharepoint.com/:f:/g/enterpriseprojects/VehicleServiceInformation/EpAg-XzfjaZOhz1wFq4oigkBr0ecZO61gtY2h4Yc3ZXv3g?e=fbTwrN",
            "Chevrolet": "https://calibercollision.sharepoint.com/:f:/g/enterpriseprojects/VehicleServiceInformation/Ej2hmGaQ0gpPm_AcAVat14oBzmKOH3DmWQX6rUqIBIkMzQ?e=0ncLlc",
            "Chrysler": "https://calibercollision.sharepoint.com/:f:/g/enterpriseprojects/VehicleServiceInformation/Ek-IrbBt8M1KirZFMw_WRWcB5PutCD-6G_74ngyR9sa-0Q?e=CDAXq3",
            "Dodge": "https://calibercollision.sharepoint.com/:f:/g/enterpriseprojects/VehicleServiceInformation/EtHvOVtuVXZMvNvPhl53pzABfMg2bujv7oqxqn_8Yoe3uw?e=FN0HFG",
            "Fiat": "https://calibercollision.sharepoint.com/:f:/g/enterpriseprojects/VehicleServiceInformation/EjJ9n4qYcYVDi8O6Wo-ottgBin-8czx56lHJSkEJqsQPCg?e=N8f2bN",
            "Ford": "https://calibercollision.sharepoint.com/:f:/g/enterpriseprojects/VehicleServiceInformation/Em8bPa15EA5GktOKhvjOZvkBcBKUNAat9udR377A9QywqA?e=1RzygZ",
            "Genesis": "https://calibercollision.sharepoint.com/:f:/g/enterpriseprojects/VehicleServiceInformation/EpZx-JEEwA1Ku4_QcRc-AdUBaq80GNKpJ220mqDh3K6hSw?e=mvydbg",
            "GMC": "https://calibercollision.sharepoint.com/:f:/g/enterpriseprojects/VehicleServiceInformation/EpZx-JEEwA1Ku4_QcRc-AdUBaq80GNKpJ220mqDh3K6hSw?e=RjLqeo",
            "Honda": "https://calibercollision.sharepoint.com/:f:/g/enterpriseprojects/VehicleServiceInformation/El5TH2K3oUFLvYsLiCLddfsBZedXJsd0cccD_PdEd8VeZw?e=C1jer2",
            "Hyundai": "https://calibercollision.sharepoint.com/:f:/g/enterpriseprojects/VehicleServiceInformation/Ej3OmwyxcvpCgKSgRitDBRQBogmGhAXl02fTSWZPUv8fTg?e=Qc5d9L",
            "Infiniti": "https://calibercollision.sharepoint.com/:f:/g/enterpriseprojects/VehicleServiceInformation/EnluXQ2j1ApLgKz8sSLlZZ4BdsCFiXn7QDgDurAYIJucjA?e=1VsA8u",
            "Jaguar": "https://calibercollision.sharepoint.com/:f:/g/enterpriseprojects/VehicleServiceInformation/Elcyy5IrpfFLhZLZuGcOlRcBL9_1AsfTg5US8y5I-ukHZQ?e=3ppWTI",
            "Jeep": "https://calibercollision.sharepoint.com/:f:/g/enterpriseprojects/VehicleServiceInformation/EteEBfd6I15LnVR7jVoC1vQBk7TB1PpYPWTeOdNVxX9WxA?e=VaxhkI",
            "Kia": "https://calibercollision.sharepoint.com/:f:/g/enterpriseprojects/VehicleServiceInformation/EmeBljSMc45AtY8TPaG3LzABPyukP0RIhCH6AnX1g7tc9w?e=N4Aaey",
            "Land Rover": "https://calibercollision.sharepoint.com/:f:/g/enterpriseprojects/VehicleServiceInformation/Em_X8FfBGmZNm3XhfSe4aFgBEIVhlMpv6NyMlya7FIssNQ?e=NDVpGP",
            "Lexus": "https://calibercollision.sharepoint.com/:f:/g/enterpriseprojects/VehicleServiceInformation/EoWwGnZ3tyxGjHifp8REvzsB6bt-IS1vJWbBOQXqpzfdTA?e=l2jI53",
            "Lincoln": "https://calibercollision.sharepoint.com/:f:/g/enterpriseprojects/VehicleServiceInformation/EiPhzaQo7RVJo0RIonH1iBUB3-uaZ5R5SbwSlCwKkdCo3A?e=mYLBce",
            "Mazda": "https://calibercollision.sharepoint.com/:f:/g/enterpriseprojects/VehicleServiceInformation/En4bjRzkrgpEoqNs75NUqw0BT11FRfgVaeUI5sFJR9g-lQ?e=5cBCNl",
            "Mercedes": "https://calibercollision.sharepoint.com/:f:/g/enterpriseprojects/VehicleServiceInformation/EqPtFOl1KyZJqrM1-hvKFNsBHYyQg3SIRSG_u_GdulNfmA?e=c1WkH7",
            "Mini": "https://calibercollision.sharepoint.com/:f:/g/enterpriseprojects/VehicleServiceInformation/EsuCQ3jb4SdGn_z2nctNatIBxbBh9AfszpoLNuHeyeRzaQ?e=kB1Wpa",
            "Mitsubishi": "https://calibercollision.sharepoint.com/:f:/g/enterpriseprojects/VehicleServiceInformation/Ep4sBEBP2G9EpVOl4yii-B4BI0BTrQ4SsTb-3eL_aaqDBg?e=ojV9Ku",
            "Nissan": "https://calibercollision.sharepoint.com/:f:/g/enterpriseprojects/VehicleServiceInformation/EuNJjABNTxhOtV0xKwFIA40BRfK2HZ_5DghIAwz8lZkGqw?e=zOs16F",
            "Porsche": "https://calibercollision.sharepoint.com/:f:/g/enterpriseprojects/VehicleServiceInformation/EsfyKyNjxahEgDvI9veqMY8BJPXSIS7DivN9zExS4Tw3TA?e=Gzb9SX",
            "Ram": "https://calibercollision.sharepoint.com/:f:/g/enterpriseprojects/VehicleServiceInformation/Eqk8MFOiaG5OvRTEu_hoHF4BfIgtHcOkRgzVAd98tvD0rw?e=pwU2uh",
            "Rolls Royce": "https://calibercollision.sharepoint.com/:f:/g/enterpriseprojects/VehicleServiceInformation/EpKJpSlNaltJkVW6rcOaMwYBa813qB32RB9MkpvxntyHPA?e=cbtNeB",
            "Subaru": "https://calibercollision.sharepoint.com/:f:/g/enterpriseprojects/VehicleServiceInformation/Etpws1skaeZHqm9doHVREFMBvCRb0Rm8Oj1mKo0HBruAxw?e=j9VRqN",
            "Toyota": "https://calibercollision.sharepoint.com/:f:/g/enterpriseprojects/VehicleServiceInformation/EiB53aPXartJhkxyWzL5AFABZQsY3x-XDWPXQCqgFIrvoQ?e=m4DrKQ",
            "Volkswagen": "https://calibercollision.sharepoint.com/:f:/g/enterpriseprojects/VehicleServiceInformation/Esry4bJndglOg1tzeCO63-kBojqzcx6mt0PZRNiIrHtaXw?e=0xdogQ",
            "Volvo": "https://calibercollision.sharepoint.com/:f:/g/enterpriseprojects/VehicleServiceInformation/Enaf3N8_gq1EvPizK5mLL4gBBiklHgRi_JiQV7QGE2j-Vg?e=IIhhTu",                        
            # Add other manufacturer links here
        }
        self.completed_manufacturers = []
        self.threads = []
        
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
        manufacturers = ["Acura", "Alfa Romeo", "Audi", "BMW", "Brightdrop", "Buick", "Cadillac", "Chevrolet","Chrysler", "Dodge", 
                         "Fiat", "Ford", "Genesis", "GMC", "Honda", "Hyundai", "Infiniti", "Jaguar", "Jeep", "Kia", "Lexus", "Land Rover", "Lincoln", 
                         "Mazda", "Mercedes", "Mini", "Mitsubishi", "Nissan", "Porsche", "Ram", "Rolls Royce", "Subaru", "Toyota", 
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
            self.excel_paths = [path.strip() for path in self.excel_paths]  # Ensure no leading/trailing spaces
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
                # Show the terminal window
                self.terminal = TerminalDialog(self)
                self.terminal.show()

                self.selected_manufacturers = selected_manufacturers
                self.current_index = 0
                self.process_next_manufacturer()

            else:
                QMessageBox.warning(self, 'Warning', "Automation process canceled.", QMessageBox.Ok)
        else:
            QMessageBox.warning(self, 'Warning', "Please select Excel files and manufacturers first.", QMessageBox.Ok)

    def process_next_manufacturer(self):
        if self.current_index < len(self.selected_manufacturers):
            manufacturer = self.selected_manufacturers[self.current_index]
            excel_path = self.excel_paths[self.current_index]
            sharepoint_link = self.manufacturer_links.get(manufacturer)

            if sharepoint_link:
                script_path = os.path.join(os.path.dirname(__file__), "SharepointExtractor.py")
                excel_path = excel_path.strip()
                sharepoint_link = sharepoint_link.strip()
                args = ["python", script_path, sharepoint_link, excel_path]

                # Run the command in a thread and show the output in the terminal
                thread = WorkerThread(args, manufacturer)
                thread.output_signal.connect(self.terminal.append_output)
                thread.finished_signal.connect(self.on_manufacturer_finished)
                thread.start()
                self.threads.append(thread)

    def on_manufacturer_finished(self, manufacturer):
        # Mark manufacturer as completed
        self.completed_manufacturers.append(manufacturer)

        # Show success message in terminal for this manufacturer
        self.terminal.append_output(f"Completed {manufacturer}. Waiting 10 seconds before next manufacturer...")

        # Wait for 10 seconds before starting the next manufacturer
        sleep(10)

        # Move to the next manufacturer
        self.current_index += 1
        self.process_next_manufacturer()

        # If all manufacturers are completed, show a completion message
        if self.current_index >= len(self.selected_manufacturers):
            completed_message = "The Following Manufacturers have been completed:\n"
            completed_message += "\n".join(self.completed_manufacturers)
            QMessageBox.information(self, 'Completed', completed_message, QMessageBox.Ok)
            self.terminal.append_output("All manufacturers processed successfully.")

            
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
            # Show the terminal window
            self.terminal = TerminalDialog(self)
            self.terminal.show()

            # Start processing manufacturers one by one
            for manufacturer, excel_path in zip(selected_manufacturers, self.excel_paths):
                sharepoint_link = self.manufacturer_links.get(manufacturer)
                if sharepoint_link:
                    script_path = os.path.join(os.path.dirname(__file__), "SharepointExtractor.py")
                    excel_path = excel_path.strip()
                    sharepoint_link = sharepoint_link.strip()
                    args = ["python", script_path, sharepoint_link, excel_path]

                    # Run the command in a thread and show the output in the terminal
                    thread = WorkerThread(args, manufacturer)
                    thread.output_signal.connect(self.terminal.append_output)
                    thread.finished_signal.connect(self.on_manufacturer_finished)
                    thread.start()
                    self.threads.append(thread)
        else:
            QMessageBox.warning(self, 'Warning', "Full automation process canceled.", QMessageBox.Ok)


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
