import sys
from PyQt5.QtWidgets import (QApplication, QDialog, QPlainTextEdit, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, 
                             QTreeWidget, QTreeWidgetItem, QMessageBox, QFileDialog, QCheckBox, QScrollArea)
from PyQt5.QtGui import QFont
from PyQt5.QtCore import Qt,pyqtSignal,QThread
from threading import Thread
import subprocess
from time import sleep
import os

#Adds Terminal infoormation
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
        process = subprocess.Popen(
            self.command,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            bufsize=1,
            universal_newlines=True,
            encoding='utf-8',  # 💡 Add this to avoid UnicodeDecodeError
            errors='replace',  # 💡 Replaces unknown characters with � instead of crashing
            env=env
)

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
            "Acura": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/Egwph7U2M7tMgy4U82m8HVEBbeB3CxoibZz9zFww6iBZqw?e=l6ekEO",
            "Alfa Romeo": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/EvO_UqobQBJOrwQAGefJrNgB4YDcOAAtQy_Y578hKRJE9A?e=73mDgy",
            "Audi": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/Ek0hoMxpf-RKgEkcFE7q4cgBz-OHaRSh6B5OSRnMVOPLKw?e=AEmzrm",
            "BMW": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/EiXISHFadJVPh0GzAp9RvXgBJ9u-Y1QcpDAfgttL87t9cQ?e=mLUNPd",
            "Brightdrop": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/Esr1s_-xkRlMr9SDAGPK6qoBM92UVxBXnHgYyXSYUSLzcQ?e=MOh0KB",
            "Buick": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/Ertv57aXnodKl9TFNSovFvEBvq-7X1ctOg0K5yH1Xj8VPA?e=hBQFy1",
            "Cadillac": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/ElXaigJrO7VGjrIxQDgZtX4BlSydGdiUGPabxGNiEw8SsA?e=JTiTNf",
            "Chevrolet": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/EseY-o8uBStGuO6Vz1DBlsMBaPVd97tw-CmkcANhFQju2A?e=ABN4Qu",
            "Chrysler": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/EnFQfBk769lBnt4QoykmWiIB_3qmkBAy0dIkWoEELbpfrA?e=BDvm7K",
            "Dodge": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/EuWilX-Oxj9OhJXI1rkXR6kBT7JEIwh12CaDN1rxcGQOLA?e=JSsxpq",
            "Fiat": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/EnS87Jrn5gdJtbuiEY4LkBkBkD-aFjNiR54RhIL8ApivPQ?e=7Iag4Z",
            "Ford": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/ErpPrcU6itZOuXT1D1m9nNYB-5FC0XhUZsGOsVoV_Js4Pw?e=taa6Bp",
            "Genesis": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/EsTgaheoczdOgE4iJKaMNlAB7tC_R8edA35MadVBZc7kbg?e=fUIbym",
            "GMC": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/EvjRi9A-4udLumIKpBpoJFwBkRbglOwe3W6C5obtbl40qw?e=tkmjxy",
            "Honda": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/Et92iX3A2mhOojkKNdjrWF8BsJPFFBO_gWP5Q84KO3nfiA?e=oiV1Xe",
            "Hyundai": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/EqZPS-XofL1Gsov7b0cEZtEBXKRtqoa3H1GNbA2cqVoTOw?e=TNHJ4N",
            "Infiniti": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/EsM2wf_e2chBjdBwkVDsi4EBCTlNibezvEvXx1PdCW4PWw?e=nFXG0Q",
            "Jaguar": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/Eu1NKfOfm1FCub0XFzLRaysBAD8H7eJs0Htf9tHoE3uA1A?e=g4Yceg",
            "Jeep": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/EnRqXg8biOdFgl0uvZ9-BzwBdg3s6QD9AVqeHkLcFoYy6g?e=B3quWV",
            "Kia": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/EiSaF12paSFDukUZHI98LtcBpyvMEf6qzkf_B9pyYYKrrA?e=w6OrxP",
            "Land Rover": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/EgcVzgpah55IiMG2F0a_-OsBS4sfkaqdiHgNeczziH8wvA?e=XbmPyP",
            "Lexus": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/EofqjO0OAy9BooaRObpgCUkB1VN8bewjLtB1NIOPKwmwhA?e=eg9buz",
            "Lincoln": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/EuI2AalhjZ9NjykawVp1RuIB7S3INdHJSsrTbRYUN9QUaQ?e=OyZHt3",
            "Mazda": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/Evw3r94DVIRHkOZD10Jro5wB9PGNleSRg1SDjQ3zG5x-ig?e=iNZAYe",
            "Mercedes": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/EuuT3bDUP1lEvbjM0hkk-FQBpfQIj4fl5rwcI9FYF1MnYA?e=NCEj6K",
            "Mini": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/ErE803EUR89EuNPaCYdPKmYBp-MqQCwRdH7aoPZ0cmdh2Q?e=cZRPvS",
            "Mitsubishi": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/EnWx8oYtnadMsJ2KsflP2uoBfDyj1XXdDVu5JPu9xMt13g?e=1Y08r2",
            "Nissan": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/El80Z4DzjwhAjZ7FMNFFtsABmrPLL_MjOaKMPnj0NF25UQ?e=HeQ9Ob",
            "Porsche": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/EvqAAY7ZWuxOgiNIOG_5RKYBfiW9eyAGSedUi9ZnDMCK8g?e=z4fZmB",
            "Ram": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/Er4y2ZHuq7RMrM8KMy5aSy0B53UPrZtUexV2apOYE-VdFw?e=RbjqgX",
            "Rolls Royce": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/EuUx8qWb8a5Lg69dZiqgPrkBElO3gQAuaLTZOvKdlOIkJg?e=3bPP4d",
            "Subaru": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/Etn9UVVyAthIiHPcwKnvQzgB8wrkm3qwyQOWaIOd-CJZYA?e=zSG8Mw",
            "Tesla": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/EnLNge4G5CZDoa4Ec9t5ESwB-1MP-MOXranHT_DBuv6ZHg?e=GtwTDf",
            "Toyota": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/EonbuVPsLixOrWBA-LEmXpoBzVe-CeCreW_66jiroMFMHA?e=x6bV7i",
            "Volkswagen": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/EkZbGXTqw5RJo-QY3CtTtukBdVpwKTz-QeDFpus_pHRNDg?e=2cQSHi",
            "Volvo": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/EneJUPUGviJEjn0OHfyqQNYBG9fqQ5g23OS15-2KALJIbA?e=8weOwd",                        
            # Add other manufacturer links here
        }
        self.repair_links = {
            # Add Repair SI SharePoint links like:
            "Acura": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/EmmQxHGQ2wxNpIO04pNu_iIBqou8brkQWpKnjHPgSFT-CQ?e=LlDLJe",
        }

        self.completed_manufacturers = []
        self.threads = []
        
    def initUI(self):
        self.setWindowTitle('Hyperlink Automation')
        self.setStyleSheet("background-color: #2e2e2e; color: white;")
        layout = QVBoxLayout()
    
        # Excel file selection layout
        file_selection_layout = QHBoxLayout()
        self.select_file_button = CustomButton('Select Excel Files', '#e63946', self)
        self.select_file_button.clicked.connect(self.select_excel_files)
        file_selection_layout.addWidget(self.select_file_button)
    
        # Excel file path display
        self.excel_path_label = QLabel('No files selected')
        self.excel_path_label.setStyleSheet("font-size: 14px; padding: 5px; border: 1px solid #555555; border-radius: 5px; background-color: #3e3e3e;")
        file_selection_layout.addWidget(self.excel_path_label)
    
        layout.addLayout(file_selection_layout)
    
        # "Select All (Manufacturers)" and "Select All (ADAS Systems)" button layout
        select_all_buttons_layout = QHBoxLayout()
        self.select_all_manufacturers_button = CustomButton('Select All (Manufacturers)', '#e3b505', self)
        self.select_all_manufacturers_button.clicked.connect(self.select_all_manufacturers)
        select_all_buttons_layout.addWidget(self.select_all_manufacturers_button)
    
        self.select_all_adas_button = CustomButton('Select All (ADAS Systems)', '#e3b505', self)
        self.select_all_adas_button.clicked.connect(self.select_all_adas)
        select_all_buttons_layout.addWidget(self.select_all_adas_button)
    
        layout.addLayout(select_all_buttons_layout)
    
        # Manufacturer and ADAS selection layout
        manufacturer_selection_layout = QHBoxLayout()
    
        # Manufacturer tree widget with checkboxes
        manufacturer_list_layout = QVBoxLayout()
        self.manufacturer_tree = QTreeWidget(self)
        self.manufacturer_tree.setHeaderHidden(True)
        self.manufacturer_tree.setStyleSheet("background-color: #3e3e3e; color: white; border: 1px solid #555555; border-radius: 5px;")
        manufacturers = ["Acura", "Alfa Romeo", "Audi", "BMW", "Brightdrop", "Buick", "Cadillac", "Chevrolet", "Chrysler", "Dodge",
                         "Fiat", "Ford", "Genesis", "GMC", "Honda", "Hyundai", "Infiniti", "Jaguar", "Jeep", "Kia", "Lexus", 
                         "Land Rover", "Lincoln", "Mazda", "Mercedes", "Mini", "Mitsubishi", "Nissan", "Porsche", "Ram", 
                         "Rolls Royce", "Subaru", "Tesla", "Toyota", "Volkswagen", "Volvo"]
        for manufacturer in manufacturers:
            item = QTreeWidgetItem(self.manufacturer_tree)
            item.setText(0, manufacturer)
            item.setCheckState(0, Qt.Unchecked)
        manufacturer_list_layout.addWidget(self.manufacturer_tree)
        manufacturer_selection_layout.addLayout(manufacturer_list_layout)
    
        # ADAS Acronyms section
        adas_selection_layout = QVBoxLayout()
        adas_label = QLabel("ADAS Systems")
        adas_label.setStyleSheet("font-size: 14px; padding: 5px;")
        adas_selection_layout.addWidget(adas_label)
    
        adas_acronyms = ["ACC", "AEB", "AHL", "APA", "BSW", "BUC", "LKA", "LW", "NV", "SVC"]
        self.adas_checkboxes = []
        repair_systems = [
            "SAS", "YAW", "G-Force", "SWS", "AHL", "NV", "HUD", "SRS",
            "ESC", "SRS D&E", "SCI", "SRR", "HLI", "TPMS", "SBI",
            "EBDE (1)", "EBDE (2)", "HDE (1)", "HDE (2)", "LGR", "PSI", "WRL",
            "PCM", "TRANS", "AIR", "ABS", "BCM",
            "KEY", "FOB", "HVAC (1)", "HVAC (2)", "COOL", "HEAD (1)", "HEAD (2)"
        ]
        self.repair_checkboxes = []

        for adas in adas_acronyms:
            checkbox = QCheckBox(adas, self)
            checkbox.setStyleSheet("font-size: 12px; padding: 5px;")
            self.adas_checkboxes.append(checkbox)
            adas_selection_layout.addWidget(checkbox)
    
        manufacturer_selection_layout.addLayout(adas_selection_layout)
        layout.addLayout(manufacturer_selection_layout)
        
        
        
        # Repair Systems section
        repair_selection_layout = QVBoxLayout()
        repair_label = QLabel("Repair Systems")
        repair_label.setStyleSheet("font-size: 14px; padding: 5px;")
        repair_selection_layout.addWidget(repair_label)
        
        for system in repair_systems:
            checkbox = QCheckBox(system, self)
            checkbox.setStyleSheet("font-size: 12px; padding: 5px;")
            self.repair_checkboxes.append(checkbox)
            repair_selection_layout.addWidget(checkbox)
        
        manufacturer_selection_layout.addLayout(repair_selection_layout)
        

    
        # Theme switch section
        # Theme switch section
        theme_switch_section = QHBoxLayout()

        # ADAS / Repair SI Label and Toggle
        self.mode_label_left = QLabel("Repair SI")
        self.mode_label_left.setStyleSheet("font-size: 14px; padding: 10px;")
        theme_switch_section.addWidget(self.mode_label_left)

        self.si_mode_toggle = QCheckBox()
        self.si_mode_toggle.setFixedSize(60, 30)
        self.si_mode_toggle.setStyleSheet("""
            QCheckBox::indicator {
                width: 60px;
                height: 30px;
            }
            QCheckBox {
                background-color: #555;
                border-radius: 15px;
            }
        """)
        theme_switch_section.addWidget(self.si_mode_toggle)

        # Dark mode toggle
        theme_switch_section.addStretch()
        self.theme_toggle = ToggleSwitch(self)
        theme_switch_section.addWidget(self.theme_toggle)
        layout.addLayout(theme_switch_section)
    
        # Start button
        self.start_button = CustomButton('Start Automation', '#e63946', self)
        self.start_button.clicked.connect(self.start_automation)
        layout.addWidget(self.start_button)
    
        self.setLayout(layout)
        self.resize(600, 400)
    
    # Function to select/unselect all manufacturers
    def select_all_manufacturers(self):
        select_all_checked = True
        for i in range(self.manufacturer_tree.topLevelItemCount()):
            item = self.manufacturer_tree.topLevelItem(i)
            if item.checkState(0) != Qt.Checked:
                select_all_checked = False
                break
    
        for i in range(self.manufacturer_tree.topLevelItemCount()):
            item = self.manufacturer_tree.topLevelItem(i)
            item.setCheckState(0, Qt.Checked if not select_all_checked else Qt.Unchecked)
    
    # Function to select/unselect all ADAS systems
    def select_all_adas(self):
        select_all_checked = all(checkbox.isChecked() for checkbox in self.adas_checkboxes)
    
        for checkbox in self.adas_checkboxes:
            checkbox.setChecked(not select_all_checked)

    
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
    
        # Collect selected ADAS acronyms
        if self.si_mode_toggle.isChecked():
            selected_adas = [cb.text() for cb in self.repair_checkboxes if cb.isChecked()]
        else:
            selected_adas = [cb.text() for cb in self.adas_checkboxes if cb.isChecked()]

    
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
    
                # Set up for processing manufacturers
                self.selected_manufacturers = selected_manufacturers
                self.selected_adas = selected_adas  # Save the ADAS systems for use later
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
            
            if self.si_mode_toggle.isChecked():
                sharepoint_link = self.repair_links.get(manufacturer)
            else:
                sharepoint_link = self.manufacturer_links.get(manufacturer)

    
            if sharepoint_link:
                # Define the script path
                script_path = os.path.join(os.path.dirname(__file__), "SharepointExtractor.py")
    
                # Collect selected ADAS systems
                selected_adas = [checkbox.text() for checkbox in self.adas_checkboxes if checkbox.isChecked()]
    
                # Arguments for the subprocess
                adas_or_repair = [cb.text() for cb in (self.adas_checkboxes if not self.si_mode_toggle.isChecked() else self.repair_checkboxes) if cb.isChecked()]
                mode_flag = "repair" if self.si_mode_toggle.isChecked() else "adas"
                args = ["python", script_path, sharepoint_link, excel_path, ",".join(adas_or_repair), mode_flag]


    
                # Run the command in a thread
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
                
                if self.si_mode_toggle.isChecked():
                    sharepoint_link = self.repair_links.get(manufacturer)
                else:
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
