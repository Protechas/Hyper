import sys
from PyQt5.QtWidgets import (QApplication, QDialog, QPlainTextEdit, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, 
                             QTreeWidget, QTreeWidgetItem, QMessageBox, QFileDialog, QCheckBox, QScrollArea, QListWidget, QProgressBar)
from PyQt5.QtGui import QFont
from PyQt5.QtCore import Qt,pyqtSignal,QThread
from threading import Thread
import subprocess
import signal
import psutil
from time import sleep
import datetime
import os
import logging
import re

######################################################################################     Terminal & Certain GUI (Buttons and Switches) Code    ######################################################################################

# ── configure a “Logs” folder in Documents ──
LOG_DIR = os.path.join(os.path.expanduser("~"), "Documents", "Hyper Logs")
os.makedirs(LOG_DIR, exist_ok=True)

# ── log filename with timestamp ──
now = datetime.datetime.now()   # ← module.datetime.now()
log_file = os.path.join(
    LOG_DIR,
    now.strftime("Hyper_Log_%m_%d_%Y_%H-%M-%S.log")
)

# ── basicConfig writes to both file and console ──
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s: %(message)s",
    handlers=[
        logging.FileHandler(log_file, encoding="utf-8"),
        logging.StreamHandler(sys.stdout)
    ]
)

#Adds Terminal infoormation
class WorkerThread(QThread):
    finished_signal = pyqtSignal(str, bool)
    output_signal   = pyqtSignal(str)
    progress_signal = pyqtSignal(int)  # ← added for live percentage updates

    def __init__(self, command: list[str], manufacturer: str, parent=None):
        super().__init__(parent)
        self.command      = command
        self.manufacturer = manufacturer
        self.process      = None

    def run(self):
        # ── Prepare env ──
        env = os.environ.copy()
        env["PYTHONUNBUFFERED"] = "1"

        if os.name == "nt":
            creationflags = subprocess.CREATE_NEW_PROCESS_GROUP
            preexec_fn    = None
        else:
            creationflags = 0
            preexec_fn    = os.setsid

        self.process = subprocess.Popen(
            self.command,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            bufsize=1,
            universal_newlines=True,
            encoding="utf-8",
            errors="replace",
            env=env,
            creationflags=creationflags,
            preexec_fn=preexec_fn
        )

        # Estimate total line count (for basic progress calculation)
        estimated_total_lines = 300  # ← adjust this if you want more accuracy
        processed_lines = 0

        # ── Stream stdout ──
        for line in iter(self.process.stdout.readline, ""):
            if not line:
                break
            self.output_signal.emit(line.rstrip("\n"))
            processed_lines += 1

            # Emit progress signal
            percent = min(100, int((processed_lines / estimated_total_lines) * 100))
            self.progress_signal.emit(percent)

        self.process.stdout.close()

        # ── Wait & then stream stderr if error ──
        ret = self.process.wait()
        success = (ret == 0)
        if not success:
            for e in iter(self.process.stderr.readline, ""):
                if not e:
                    break
                self.output_signal.emit(e.rstrip("\n"))
        self.process.stderr.close()

        # Ensure final progress is 100%
        self.progress_signal.emit(100)

        self.finished_signal.emit(self.manufacturer, success)

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

class ModeSwitch(QCheckBox):
    def __init__(self, parent=None):
        super().__init__(parent)
        # no visible text on the switch itself
        self.setText("")
        self.setFixedSize(50, 25)
        self.setStyleSheet("""
            QCheckBox {
                background-color: #888;
                border-radius: 12px;
            }
            QCheckBox::indicator {
                width: 21px; height: 21px;
                border-radius: 10px;
                background-color: white;
                margin: 2px;
            }
            QCheckBox::indicator:checked {
                margin-left: 27px;
            }
            QCheckBox:disabled {
                background-color: #555;
            }               
        """)

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
        
######################################################################################     Main Application Code    ######################################################################################

class SeleniumAutomationApp(QWidget):
    def __init__(self):
        super().__init__()
        
        # 🆕 hide CM progress lines in terminal output
        self.hide_cm_progress_in_terminal = True
        
        self.initUI()
        self.terminal = None           
        self.excel_paths = []
                                   ########################################################     ADAS SI Links     ########################################################
        self.manufacturer_links = {
            "Acura": [
                "https://calibercollision.sharepoint.com/:f:/s/O365-DepartmentofInformationSoloutions/Er9Jvy1gtUBAtz59yCRcSmMBI6Z0VaIZGz8bAxHh10_NqQ?e=KSGOjN",# Documents (2012 - 2016)
                "https://calibercollision.sharepoint.com/:f:/s/O365-DepartmentofInformationSoloutions/Ek5WPRnM0plLqbOnqeK9DHMBzoVfYiOKx-KNylrDyPgyUQ?e=FnPH5f",# Documents (2017 - 2021)
                "https://calibercollision.sharepoint.com/:f:/s/O365-DepartmentofInformationSoloutions/EgH-RFx6tAJOl9NufaV1y4ABx75-yRACbaYYqXFg2IzK-g?e=wNjzRH" # Documents (2022 - 2026)
            ],
            "Alfa Romeo": [
                "https://calibercollision.sharepoint.com/:f:/s/O365-DepartmentofInformationSoloutions/ErOke5xzYSdJuxzA1RXrlTwBCJxKIitAemYutpEimASATg?e=hWdGVy",# Documents (2012 - 2016)
                "https://calibercollision.sharepoint.com/:f:/s/O365-DepartmentofInformationSoloutions/Er2MZy2hndVHowEc489Cwr0ByJuhjHVrSWBmHhBHsimnZA?e=H3sAVX",# Documents (2017 - 2021)
                "https://calibercollision.sharepoint.com/:f:/s/O365-DepartmentofInformationSoloutions/Et84MN80bJ9On_e2O19o4SkBh7gcFCWl0r9z__aOrDX1Og?e=0VYApY" # Documents (2022 - 2026)
            ],
            "Audi": [
                "https://sharepoint.com/.../Audi (2012 - 2016)",
                "https://sharepoint.com/.../Audi (2017 - 2021)",
                "https://sharepoint.com/.../Audi (2022 - 2026)"
            ],
            "BMW": [
                "https://sharepoint.com/.../BMW (2012 - 2016)",
                "https://sharepoint.com/.../BMW (2017 - 2021)",
                "https://sharepoint.com/.../BMW (2022 - 2026)"
            ],
            "Brightdrop": [
                "https://sharepoint.com/.../Brightdrop (2012 - 2016)",
                "https://sharepoint.com/.../Brightdrop (2017 - 2021)",
                "https://sharepoint.com/.../Brightdrop (2022 - 2026)"
            ],
            "Buick": [
                "https://sharepoint.com/.../Buick (2012 - 2016)",
                "https://sharepoint.com/.../Buick (2017 - 2021)",
                "https://sharepoint.com/.../Buick (2022 - 2026)"
            ],
            "Cadillac": [
                "https://sharepoint.com/.../Cadillac (2012 - 2016)",
                "https://sharepoint.com/.../Cadillac (2017 - 2021)",
                "https://sharepoint.com/.../Cadillac (2022 - 2026)"
            ],
            "Chevrolet": [
                "https://sharepoint.com/.../Chevrolet (2012 - 2016)",
                "https://sharepoint.com/.../Chevrolet (2017 - 2021)",
                "https://sharepoint.com/.../Chevrolet (2022 - 2026)"
            ],
            "Chrysler": [
                "https://sharepoint.com/.../Chrysler (2012 - 2016)",
                "https://sharepoint.com/.../Chrysler (2017 - 2021)",
                "https://sharepoint.com/.../Chrysler (2022 - 2026)"
            ],
            "Dodge": [
                "https://sharepoint.com/.../Dodge (2012 - 2016)",
                "https://sharepoint.com/.../Dodge (2017 - 2021)",
                "https://sharepoint.com/.../Dodge (2022 - 2026)"
            ],
            "Fiat": [
                "https://sharepoint.com/.../Fiat (2012 - 2016)",
                "https://sharepoint.com/.../Fiat (2017 - 2021)",
                "https://sharepoint.com/.../Fiat (2022 - 2026)"
            ],
            "Ford": [
                "https://sharepoint.com/.../Ford (2012 - 2016)",
                "https://sharepoint.com/.../Ford (2017 - 2021)",
                "https://sharepoint.com/.../Ford (2022 - 2026)"
            ],
            "Genesis": [
                "https://sharepoint.com/.../Genesis (2012 - 2016)",
                "https://sharepoint.com/.../Genesis (2017 - 2021)",
                "https://sharepoint.com/.../Genesis (2022 - 2026)"
            ],
            "GMC": [
                "https://sharepoint.com/.../GMC (2012 - 2016)",
                "https://sharepoint.com/.../GMC (2017 - 2021)",
                "https://sharepoint.com/.../GMC (2022 - 2026)"
            ],
            "Honda": [
                "https://sharepoint.com/.../Honda (2012 - 2016)",
                "https://sharepoint.com/.../Honda (2017 - 2021)",
                "https://sharepoint.com/.../Honda (2022 - 2026)"
            ],
            "Hyundai": [
                "https://sharepoint.com/.../Hyundai (2012 - 2016)",
                "https://sharepoint.com/.../Hyundai (2017 - 2021)",
                "https://sharepoint.com/.../Hyundai (2022 - 2026)"
            ],
            "Infiniti": [
                "https://sharepoint.com/.../Infiniti (2012 - 2016)",
                "https://sharepoint.com/.../Infiniti (2017 - 2021)",
                "https://sharepoint.com/.../Infiniti (2022 - 2026)"
            ],
            "Jaguar": [
                "https://sharepoint.com/.../Jaguar (2012 - 2016)",
                "https://sharepoint.com/.../Jaguar (2017 - 2021)",
                "https://sharepoint.com/.../Jaguar (2022 - 2026)"
            ],
            "Jeep": [
                "https://sharepoint.com/.../Jeep (2012 - 2016)",
                "https://sharepoint.com/.../Jeep (2017 - 2021)",
                "https://sharepoint.com/.../Jeep (2022 - 2026)"
            ],
            "Kia": [
                "https://sharepoint.com/.../Kia (2012 - 2016)",
                "https://sharepoint.com/.../Kia (2017 - 2021)",
                "https://sharepoint.com/.../Kia (2022 - 2026)"
            ],
            "Land Rover": [
                "https://sharepoint.com/.../Land Rover (2012 - 2016)",
                "https://sharepoint.com/.../Land Rover (2017 - 2021)",
                "https://sharepoint.com/.../Land Rover (2022 - 2026)"
            ],
            "Lexus": [
                "https://sharepoint.com/.../Lexus (2012 - 2016)",
                "https://sharepoint.com/.../Lexus (2017 - 2021)",
                "https://sharepoint.com/.../Lexus (2022 - 2026)"
            ],
            "Lincoln": [
                "https://sharepoint.com/.../Lincoln (2012 - 2016)",
                "https://sharepoint.com/.../Lincoln (2017 - 2021)",
                "https://sharepoint.com/.../Lincoln (2022 - 2026)"
            ],
            "Mazda": [
                "https://sharepoint.com/.../Mazda (2012 - 2016)",
                "https://sharepoint.com/.../Mazda (2017 - 2021)",
                "https://sharepoint.com/.../Mazda (2022 - 2026)"
            ],
            "Mercedes": [
                "https://sharepoint.com/.../Mercedes (2012 - 2016)",
                "https://sharepoint.com/.../Mercedes (2017 - 2021)",
                "https://sharepoint.com/.../Mercedes (2022 - 2026)"
            ],
            "Mini": [
                "https://sharepoint.com/.../Mini (2012 - 2016)",
                "https://sharepoint.com/.../Mini (2017 - 2021)",
                "https://sharepoint.com/.../Mini (2022 - 2026)"
            ],
            "Mitsubishi": [
                "https://sharepoint.com/.../Mitsubishi (2012 - 2016)",
                "https://sharepoint.com/.../Mitsubishi (2017 - 2021)",
                "https://sharepoint.com/.../Mitsubishi (2022 - 2026)"
            ],
            "Nissan": [
                "https://sharepoint.com/.../Nissan (2012 - 2016)",
                "https://sharepoint.com/.../Nissan (2017 - 2021)",
                "https://sharepoint.com/.../Nissan (2022 - 2026)"
            ],
            "Porsche": [
                "https://sharepoint.com/.../Porsche (2012 - 2016)",
                "https://sharepoint.com/.../Porsche (2017 - 2021)",
                "https://sharepoint.com/.../Porsche (2022 - 2026)"
            ],
            "Ram": [
                "https://sharepoint.com/.../Ram (2012 - 2016)",
                "https://sharepoint.com/.../Ram (2017 - 2021)",
                "https://sharepoint.com/.../Ram (2022 - 2026)"
            ],
            "Rolls Royce": [
                "https://sharepoint.com/.../Rolls Royce (2012 - 2016)",
                "https://sharepoint.com/.../Rolls Royce (2017 - 2021)",
                "https://sharepoint.com/.../Rolls Royce (2022 - 2026)"
            ],
            "Subaru": [
                "https://sharepoint.com/.../Subaru (2012 - 2016)",
                "https://sharepoint.com/.../Subaru (2017 - 2021)",
                "https://sharepoint.com/.../Subaru (2022 - 2026)"
            ],
            "Tesla": [
                "https://sharepoint.com/.../Tesla (2012 - 2016)",
                "https://sharepoint.com/.../Tesla (2017 - 2021)",
                "https://sharepoint.com/.../Tesla (2022 - 2026)"
            ],
            "Toyota": [
                "https://sharepoint.com/.../Toyota (2012 - 2016)",
                "https://sharepoint.com/.../Toyota (2017 - 2021)",
                "https://sharepoint.com/.../Toyota (2022 - 2026)"
            ],
            "Volkswagen": [
                "https://sharepoint.com/.../Volkswagen (2012 - 2016)",
                "https://sharepoint.com/.../Volkswagen (2017 - 2021)",
                "https://sharepoint.com/.../Volkswagen (2022 - 2026)"
            ],
            "Volvo": [
                "https://sharepoint.com/.../Volvo (2012 - 2016)",
                "https://sharepoint.com/.../Volvo (2017 - 2021)",
                "https://sharepoint.com/.../Volvo (2022 - 2026)"
            ]
        }
                             ########################################################     Repair SI Links     ########################################################
        self.repair_links = {
            "Acura": [
                "https://calibercollision.sharepoint.com/:f:/s/O365-DepartmentofInformationSoloutions/EihKJ1TNSRdNpsQv8r32FFsB3jkwS6DfqW4Mcff4NrOr6A?e=qVTz2q",# Documents (2012 - 2016)
                "https://calibercollision.sharepoint.com/:f:/s/O365-DepartmentofInformationSoloutions/Ek3lsuYY8cZEsJwsES_Q1KwBhX3TTKcKhB5C5mdXUNReDQ?e=GEE8Mv",# Documents (2017 - 2021)
                "https://calibercollision.sharepoint.com/:f:/s/O365-DepartmentofInformationSoloutions/Ei1SfDwSRlpKpSWMZ__bhdsB4VoeqkmzqFUDdb0anGPnbw?e=VJC52s" # Documents (2022 - 2026)
            ],
            "Alfa Romeo": [
                "https://calibercollision.sharepoint.com/:f:/s/O365-DepartmentofInformationSoloutions/ErOke5xzYSdJuxzA1RXrlTwBCJxKIitAemYutpEimASATg?e=hWdGVy",# Documents (2012 - 2016)
                "https://calibercollision.sharepoint.com/:f:/s/O365-DepartmentofInformationSoloutions/Er2MZy2hndVHowEc489Cwr0ByJuhjHVrSWBmHhBHsimnZA?e=H3sAVX",# Documents (2017 - 2021)
                "https://calibercollision.sharepoint.com/:f:/s/O365-DepartmentofInformationSoloutions/Et84MN80bJ9On_e2O19o4SkBh7gcFCWl0r9z__aOrDX1Og?e=0VYApY" # Documents (2022 - 2026)
            ],
            "Audi": [
                "https://sharepoint.com/.../Audi (2012 - 2016)",
                "https://sharepoint.com/.../Audi (2017 - 2021)",
                "https://sharepoint.com/.../Audi (2022 - 2026)"
            ],
            "BMW": [
                "https://sharepoint.com/.../BMW (2012 - 2016)",
                "https://sharepoint.com/.../BMW (2017 - 2021)",
                "https://sharepoint.com/.../BMW (2022 - 2026)"
            ],
            "Brightdrop": [
                "https://sharepoint.com/.../Brightdrop (2012 - 2016)",
                "https://sharepoint.com/.../Brightdrop (2017 - 2021)",
                "https://sharepoint.com/.../Brightdrop (2022 - 2026)"
            ],
            "Buick": [
                "https://sharepoint.com/.../Buick (2012 - 2016)",
                "https://sharepoint.com/.../Buick (2017 - 2021)",
                "https://sharepoint.com/.../Buick (2022 - 2026)"
            ],
            "Cadillac": [
                "https://sharepoint.com/.../Cadillac (2012 - 2016)",
                "https://sharepoint.com/.../Cadillac (2017 - 2021)",
                "https://sharepoint.com/.../Cadillac (2022 - 2026)"
            ],
            "Chevrolet": [
                "https://sharepoint.com/.../Chevrolet (2012 - 2016)",
                "https://sharepoint.com/.../Chevrolet (2017 - 2021)",
                "https://sharepoint.com/.../Chevrolet (2022 - 2026)"
            ],
            "Chrysler": [
                "https://sharepoint.com/.../Chrysler (2012 - 2016)",
                "https://sharepoint.com/.../Chrysler (2017 - 2021)",
                "https://sharepoint.com/.../Chrysler (2022 - 2026)"
            ],
            "Dodge": [
                "https://sharepoint.com/.../Dodge (2012 - 2016)",
                "https://sharepoint.com/.../Dodge (2017 - 2021)",
                "https://sharepoint.com/.../Dodge (2022 - 2026)"
            ],
            "Fiat": [
                "https://sharepoint.com/.../Fiat (2012 - 2016)",
                "https://sharepoint.com/.../Fiat (2017 - 2021)",
                "https://sharepoint.com/.../Fiat (2022 - 2026)"
            ],
            "Ford": [
                "https://sharepoint.com/.../Ford (2012 - 2016)",
                "https://sharepoint.com/.../Ford (2017 - 2021)",
                "https://sharepoint.com/.../Ford (2022 - 2026)"
            ],
            "Genesis": [
                "https://sharepoint.com/.../Genesis (2012 - 2016)",
                "https://sharepoint.com/.../Genesis (2017 - 2021)",
                "https://sharepoint.com/.../Genesis (2022 - 2026)"
            ],
            "GMC": [
                "https://sharepoint.com/.../GMC (2012 - 2016)",
                "https://sharepoint.com/.../GMC (2017 - 2021)",
                "https://sharepoint.com/.../GMC (2022 - 2026)"
            ],
            "Honda": [
                "https://sharepoint.com/.../Honda (2012 - 2016)",
                "https://sharepoint.com/.../Honda (2017 - 2021)",
                "https://sharepoint.com/.../Honda (2022 - 2026)"
            ],
            "Hyundai": [
                "https://sharepoint.com/.../Hyundai (2012 - 2016)",
                "https://sharepoint.com/.../Hyundai (2017 - 2021)",
                "https://sharepoint.com/.../Hyundai (2022 - 2026)"
            ],
            "Infiniti": [
                "https://sharepoint.com/.../Infiniti (2012 - 2016)",
                "https://sharepoint.com/.../Infiniti (2017 - 2021)",
                "https://sharepoint.com/.../Infiniti (2022 - 2026)"
            ],
            "Jaguar": [
                "https://sharepoint.com/.../Jaguar (2012 - 2016)",
                "https://sharepoint.com/.../Jaguar (2017 - 2021)",
                "https://sharepoint.com/.../Jaguar (2022 - 2026)"
            ],
            "Jeep": [
                "https://sharepoint.com/.../Jeep (2012 - 2016)",
                "https://sharepoint.com/.../Jeep (2017 - 2021)",
                "https://sharepoint.com/.../Jeep (2022 - 2026)"
            ],
            "Kia": [
                "https://sharepoint.com/.../Kia (2012 - 2016)",
                "https://sharepoint.com/.../Kia (2017 - 2021)",
                "https://sharepoint.com/.../Kia (2022 - 2026)"
            ],
            "Land Rover": [
                "https://sharepoint.com/.../Land Rover (2012 - 2016)",
                "https://sharepoint.com/.../Land Rover (2017 - 2021)",
                "https://sharepoint.com/.../Land Rover (2022 - 2026)"
            ],
            "Lexus": [
                "https://sharepoint.com/.../Lexus (2012 - 2016)",
                "https://sharepoint.com/.../Lexus (2017 - 2021)",
                "https://sharepoint.com/.../Lexus (2022 - 2026)"
            ],
            "Lincoln": [
                "https://sharepoint.com/.../Lincoln (2012 - 2016)",
                "https://sharepoint.com/.../Lincoln (2017 - 2021)",
                "https://sharepoint.com/.../Lincoln (2022 - 2026)"
            ],
            "Mazda": [
                "https://sharepoint.com/.../Mazda (2012 - 2016)",
                "https://sharepoint.com/.../Mazda (2017 - 2021)",
                "https://sharepoint.com/.../Mazda (2022 - 2026)"
            ],
            "Mercedes": [
                "https://sharepoint.com/.../Mercedes (2012 - 2016)",
                "https://sharepoint.com/.../Mercedes (2017 - 2021)",
                "https://sharepoint.com/.../Mercedes (2022 - 2026)"
            ],
            "Mini": [
                "https://sharepoint.com/.../Mini (2012 - 2016)",
                "https://sharepoint.com/.../Mini (2017 - 2021)",
                "https://sharepoint.com/.../Mini (2022 - 2026)"
            ],
            "Mitsubishi": [
                "https://sharepoint.com/.../Mitsubishi (2012 - 2016)",
                "https://sharepoint.com/.../Mitsubishi (2017 - 2021)",
                "https://sharepoint.com/.../Mitsubishi (2022 - 2026)"
            ],
            "Nissan": [
                "https://sharepoint.com/.../Nissan (2012 - 2016)",
                "https://sharepoint.com/.../Nissan (2017 - 2021)",
                "https://sharepoint.com/.../Nissan (2022 - 2026)"
            ],
            "Porsche": [
                "https://sharepoint.com/.../Porsche (2012 - 2016)",
                "https://sharepoint.com/.../Porsche (2017 - 2021)",
                "https://sharepoint.com/.../Porsche (2022 - 2026)"
            ],
            "Ram": [
                "https://sharepoint.com/.../Ram (2012 - 2016)",
                "https://sharepoint.com/.../Ram (2017 - 2021)",
                "https://sharepoint.com/.../Ram (2022 - 2026)"
            ],
            "Rolls Royce": [
                "https://sharepoint.com/.../Rolls Royce (2012 - 2016)",
                "https://sharepoint.com/.../Rolls Royce (2017 - 2021)",
                "https://sharepoint.com/.../Rolls Royce (2022 - 2026)"
            ],
            "Subaru": [
                "https://sharepoint.com/.../Subaru (2012 - 2016)",
                "https://sharepoint.com/.../Subaru (2017 - 2021)",
                "https://sharepoint.com/.../Subaru (2022 - 2026)"
            ],
            "Tesla": [
                "https://sharepoint.com/.../Tesla (2012 - 2016)",
                "https://sharepoint.com/.../Tesla (2017 - 2021)",
                "https://sharepoint.com/.../Tesla (2022 - 2026)"
            ],
            "Toyota": [
                "https://sharepoint.com/.../Toyota (2012 - 2016)",
                "https://sharepoint.com/.../Toyota (2017 - 2021)",
                "https://sharepoint.com/.../Toyota (2022 - 2026)"
            ],
            "Volkswagen": [
                "https://sharepoint.com/.../Volkswagen (2012 - 2016)",
                "https://sharepoint.com/.../Volkswagen (2017 - 2021)",
                "https://sharepoint.com/.../Volkswagen (2022 - 2026)"
            ],
            "Volvo": [
                "https://sharepoint.com/.../Volvo (2012 - 2016)",
                "https://sharepoint.com/.../Volvo (2017 - 2021)",
                "https://sharepoint.com/.../Volvo (2022 - 2026)"
            ]
        }

        # how many times to try each manufacturer before giving up
        self.max_attempts = 10

        # track how many times we've tried each one
        self.attempts = {}

        # lists for status
        self.completed_manufacturers = []
        self.failed_manufacturers    = []
        self.failed_excels           = []
        self.given_up_manufacturers  = []
        
        self.thread         = None    # your singular thread slot, if you have one
        self.threads        = []      # ← now you can safely append to self.threads
        self.is_running     = False
        self.stop_requested = False
        self.pause_requested= False
      
    def initUI(self):
        self.setWindowTitle('Hyper')
        self.setStyleSheet("background-color: #2e2e2e; color: white;")
        layout = QVBoxLayout()
    
        # Excel file selection layout
        file_selection_layout = QHBoxLayout()
        self.select_file_button = CustomButton('Select Excel Files', '#008000', self)
        self.select_file_button.clicked.connect(self.select_excel_files)
        file_selection_layout.addWidget(self.select_file_button)
    
        # Excel file path display
        self.excel_path_label = QLabel('No files selected')
        self.excel_path_label.setStyleSheet(
            "font-size: 14px; padding: 5px; "
            "border: 1px solid #555555; border-radius: 5px; "
            "background-color: #3e3e3e;"
        )
        
        # Excel file list (scrollable)
        self.excel_list = QListWidget(self)
        self.excel_list.setFixedHeight(100)   # tweak height as you like
        
        self.excel_list.setStyleSheet(
            "font-size: 14px; padding: 5px; "
            "background-color: #3e3e3e; color: white; "
            "border: 1px solid #555555; border-radius: 5px;"
        )
        self.excel_list.addItem('No files selected, please select files')
        file_selection_layout.addWidget(self.excel_list)
        layout.addLayout(file_selection_layout)
        self.si_mode_toggle = QCheckBox()
    
        # "Select All (Manufacturers)" and "Select All (ADAS Systems)" button layout
        select_all_buttons_layout = QHBoxLayout()
        self.select_all_manufacturers_button = CustomButton('Select All (Manufacturers)', '#e3b505', self)
        self.select_all_manufacturers_button.clicked.connect(self.select_all_manufacturers)
        select_all_buttons_layout.addWidget(self.select_all_manufacturers_button)
    
        self.select_all_adas_button = CustomButton('Select All (ADAS Systems)', '#e3b505', self)
        self.select_all_adas_button.clicked.connect(self.select_all_adas)
        select_all_buttons_layout.addWidget(self.select_all_adas_button)
        
        self.select_all_repair_button = CustomButton('Select All (Repair Systems)', '#e3b505', self)
        self.select_all_repair_button.clicked.connect(self.select_all_repair)
        select_all_buttons_layout.addWidget(self.select_all_repair_button)

        layout.addLayout(select_all_buttons_layout)
    
        # Manufacturer and ADAS selection layout
        manufacturer_selection_layout = QHBoxLayout()
    
        # Manufacturer tree widget with checkboxes
        manufacturer_list_layout = QVBoxLayout()
        
        # ▶ Label above the manufacturers list
        manufacturer_label = QLabel("Manufacturers")
        manufacturer_label.setAlignment(Qt.AlignHCenter)   # ⬅️ center text
        manufacturer_label.setStyleSheet("font-size: 14px; padding: 5px 6px;")
        manufacturer_list_layout.addWidget(manufacturer_label)
        
        self.manufacturer_tree = QTreeWidget(self)
        self.manufacturer_tree.setHeaderHidden(True)
        self.manufacturer_tree.setFixedWidth(200)  # 👈 Shift closer by narrowing it
        self.manufacturer_tree.setStyleSheet("""
            QTreeWidget {
                background-color: #3e3e3e;
                color: white;
                border: 1px solid #555555;
                border-radius: 5px;
                margin-left: 10px;  /* 👈 Fine-tune left shift */
            }
        """)
        
        # Manufacturer Check Boxes
        manufacturers = ["Acura", "Alfa Romeo", "Audi", "BMW", "Brightdrop", "Buick", "Cadillac", "Chevrolet", "Chrysler", "Dodge",
                         "Fiat", "Ford", "Genesis", "GMC", "Honda", "Hyundai", "Infiniti", "Jaguar", "Jeep", "Kia", "Land Rover", 
                         "Lexus", "Lincoln", "Mazda", "Mercedes", "Mini", "Mitsubishi", "Nissan", "Porsche", "Ram", 
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
        adas_label.setAlignment(Qt.AlignHCenter)    
        adas_label.setStyleSheet("font-size: 14px; padding: 5px;")
        adas_selection_layout.addWidget(adas_label)
    
        adas_acronyms = ["ACC", "AEB", "AHL", "APA", "BSW", "BUC", "LKA", "LW", "NV", "SVC", "WAMC"]
        self.adas_checkboxes = []
        repair_systems = [
            "SAS", "YAW", "G-Force", "SWS", "AHL", "NV", "HUD", "SRS", "SRA", 
            "ESC", "SRS D&E", "SCI", "SRR", "HLI", "TPMS", "SBI", "RC",
            "EBDE (1)", "EBDE (2)", "HDE (1)", "HDE (2)", "LGR", "PSI", "WRL",
            "PCM", "TRANS", "AIR", "ABS", "BCM","ODS","OCS","OCS2","OCS3","OCS4",
            "KEY", "FOB", "HVAC (1)", "HVAC (2)", "COOL", "HEAD (1)", "HEAD (2)",
        
            # Full Names
            "Steering Angle Sensor",
            "Yaw Rate Sensor",
            "G Force Sensor",
            "Seat Weight Sensor",
            "Adaptive Head Lamps",
            "Night Vision",
            "Heads Up Display",
            "Electronic Stability Control Relearn",
            "Airbag Disengagement/Engagement",
            "Steering Column Inspection",
            "Steering Rack Relearn",
            "Headlamp Initialization",
            "Tire Pressure Monitor Relearn",
            "Seat Belt Inspection",
            "Battery Disengagement",
            "Battery Engagement",
            "Hybrid Disengagement",
            "Hybrid Engagement",
            "Liftgate Relearn",
            "Power Seat Initialization",
            "Window Relearn",
            "Powertrain Control Module Program",
            "Transmission Control Module Program",
            "Airbag Control Module Program",
            "Antilock Brake Control Module Program",
            "Body Control Module Program",
            "Key Program",
            "Key FOB Relearn",
            "Heating, Air Conditioning, Ventilation EVAC",
            "Heating, Air Conditioning, Ventilation Recharge",
            "Coolant Services",
            "Headset Reset (Spring Style)",
            "Headset Reset (Squib Style)",
        ]

        self.repair_checkboxes = []

        for adas in adas_acronyms:
            checkbox = QCheckBox(adas, self)
            checkbox.setStyleSheet("font-size: 12px; padding: 5px;")
            self.adas_checkboxes.append(checkbox)
            adas_selection_layout.addWidget(checkbox)
    
        manufacturer_selection_layout.addLayout(adas_selection_layout)
        layout.addLayout(manufacturer_selection_layout)
                    
        # === Repair Systems Section (Label on top, Scrollable box underneath) ===
        
        # Vertical layout to hold both: label AND scrollable checkbox container
        repair_box_layout = QVBoxLayout()
        
        # Label (not scrollable)
        repair_label = QLabel("Repair Systems")
        repair_label.setAlignment(Qt.AlignHCenter)  
        repair_label.setFixedWidth(200)    
        repair_label.setStyleSheet("font-size: 14px; padding: 5px;")
        repair_box_layout.addWidget(repair_label)
        
        # Scrollable checkbox area (keep a reference so we can restyle it later)
        self.repair_scroll_area = QScrollArea()
        self.repair_scroll_area.setWidgetResizable(True)
        self.repair_scroll_area.setFixedWidth(180)
        self.repair_scroll_area.setStyleSheet(
            "background-color: #3e3e3e; border: 1px solid #555555; border-radius: 5px;"
        )
        
        repair_container = QWidget()
        repair_selection_layout = QVBoxLayout(repair_container)
        
        self.repair_checkboxes = []
        for system in repair_systems:
            checkbox = QCheckBox(system, self)
            checkbox.setStyleSheet("font-size: 12px; padding: 5px;")
            self.repair_checkboxes.append(checkbox)
            repair_selection_layout.addWidget(checkbox)
        
        self.repair_scroll_area.setWidget(repair_container)
        repair_box_layout.addWidget(self.repair_scroll_area)
        
        # Add the full repair module section to the right side
        manufacturer_selection_layout.addLayout(repair_box_layout)

        # Theme switch section
        theme_switch_section = QHBoxLayout()

        # ADAS / Repair SI Label and Toggle
        switch_layout = QHBoxLayout()
        switch_layout.setSpacing(8)
        
        self.label_adas   = QLabel("ADAS SI")
        self.label_repair = QLabel("Repair SI")
        for lbl in (self.label_adas, self.label_repair):
            lbl.setStyleSheet("font-size:14px; padding:5px;")
        
        self.mode_switch = ModeSwitch(self)
        # start unchecked => ADAS
        self.mode_switch.setChecked(False)
        self.mode_switch.stateChanged.connect(self.on_si_mode_toggled)
        
        switch_layout.addWidget(self.label_adas)
        switch_layout.addWidget(self.mode_switch)
        switch_layout.addWidget(self.label_repair)
        switch_layout.addStretch()
        
        layout.addLayout(switch_layout)
        
        # after creating self.si_mode_toggle …
        self.si_mode_toggle.stateChanged.connect(self.on_si_mode_toggled)

        # Excel Format Toggle Layout (OG / New)
        excel_mode_layout = QHBoxLayout()
        excel_mode_layout.setSpacing(8)
        
        label_og   = QLabel("OG")
        label_new  = QLabel("New")
        for lbl in (label_og, label_new):
            lbl.setStyleSheet("font-size:14px; padding:5px;")
        
        self.excel_mode_switch = ModeSwitch(self)
        self.excel_mode_switch.setChecked(True)  # Start in New mode
        
        excel_mode_layout.addWidget(label_og)
        excel_mode_layout.addWidget(self.excel_mode_switch)
        excel_mode_layout.addWidget(label_new)
        excel_mode_layout.addStretch()
        
        layout.addLayout(excel_mode_layout)

        # Dark mode toggle
        theme_switch_section.addStretch()
        self.theme_toggle = ToggleSwitch(self)
        theme_switch_section.addWidget(self.theme_toggle)
        layout.addLayout(theme_switch_section)
    
        # ── Clean up Mode checkbox ──
        self.cleanup_checkbox = QCheckBox("Broken Hyperlink Mode", self)
        self.cleanup_checkbox.setStyleSheet("font-size: 14px; padding: 5px;")
        layout.addWidget(self.cleanup_checkbox)    
    
        # ── Pause/Resume & Start/Stop Buttons ──
        self.pause_button = CustomButton('Pause Automation', '#e3a008', self)
        self.pause_button.clicked.connect(self.on_pause_resume)
        self.pause_button.setEnabled(False)
    
        self.start_button = CustomButton('Start Automation', '#008000', self)
        self.start_button.clicked.connect(self.on_start_stop)
    
        # Use a vertical layout here so Pause sits above Start
        self.button_layout = QVBoxLayout()  
        self.button_layout.addWidget(self.pause_button)
        self.button_layout.addWidget(self.start_button)
        layout.addLayout(self.button_layout)
    
        # ── Progress Bars ──
        self.current_manufacturer_label = QLabel("Current Manufacturer: None")
        self.current_manufacturer_label.setStyleSheet("font-size: 13px; padding: 5px;")
        self.current_manufacturer_progress = QProgressBar()
        self.current_manufacturer_progress.setMaximum(100)
        self.current_manufacturer_progress.setValue(0)
        self.current_manufacturer_progress.setFormat("%p%")
        self.current_manufacturer_progress.setStyleSheet("""
            QProgressBar {
                font-size: 12px;
                padding: 4px;
                color: black;           /* text color */
                text-align: center;     /* center the % */
            }
        """)
        
        # 🆕 Manufacturer Hyperlink Status
        self.manufacturer_hyperlink_label = QLabel("Manufacturer Hyperlinks Indexed: 0 / 0")
        self.manufacturer_hyperlink_label.setStyleSheet("font-size: 13px; padding: 5px;")
        self.manufacturer_hyperlink_bar = QProgressBar()
        self.manufacturer_hyperlink_bar.setMaximum(100)
        self.manufacturer_hyperlink_bar.setValue(0)
        self.manufacturer_hyperlink_bar.setFormat("%p%")
        self.manufacturer_hyperlink_bar.setStyleSheet("""
            QProgressBar {
                font-size: 12px;
                padding: 4px;
                color: black;
                text-align: center;
            }
        """)
        
        self.overall_progress_label = QLabel("Overall Progress: 0%")
        self.overall_progress_label.setStyleSheet("font-size: 13px; padding: 5px;")
        self.overall_progress_bar = QProgressBar()
        self.overall_progress_bar.setMaximum(100)
        self.overall_progress_bar.setValue(0)
        self.overall_progress_bar.setFormat("%p%")
        self.overall_progress_bar.setStyleSheet("""
            QProgressBar {
                font-size: 12px;
                padding: 4px;
                color: black;           /* text color */
                text-align: center;     /* center the % */
            }
        """)
        # Style used when manually stopped (red background + red chunk)
        self._bar_style_stopped = """
            QProgressBar {
                font-size: 12px;
                padding: 4px;
                color: white;                         /* % text */
                text-align: center;
                background-color: #4a0f0f;            /* bar background stays red even at 0% */
                border: 1px solid #aa4444;
                border-radius: 4px;
            }
            QProgressBar::chunk {
                background-color: #d32f2f;            /* the filled part */
                margin: 0px;
                border-radius: 2px;
            }
        """
        
        layout.addWidget(self.current_manufacturer_label)
        layout.addWidget(self.current_manufacturer_progress)
        
        # 🆕 Insert new hyperlink bar + label here
        layout.addWidget(self.manufacturer_hyperlink_label)
        layout.addWidget(self.manufacturer_hyperlink_bar)
        
        layout.addWidget(self.overall_progress_label)
        layout.addWidget(self.overall_progress_bar)      
        
        # after creating the bars
        self.current_manufacturer_progress.setObjectName("cmBar")
        self.manufacturer_hyperlink_bar.setObjectName("mhBar")
        self.overall_progress_bar.setObjectName("ovBar")
        
        # one stylesheet that preserves the grey groove; only the CHUNK turns red when stopped
        progress_css = """
        QProgressBar {
            font-size: 12px;
            padding: 4px;
            text-align: center;
            color: black;
            border: 1px solid #555555;
            border-radius: 4px;
            background: #e0e0e0;              /* normal groove */
        }
        QProgressBar::chunk {
            background-color: #19A602;         /* normal (green) fill */
        }
        
        /* when manually stopped, KEEP the same groove; only recolor the chunk.
           Use object-id selectors + !important to override any global red background. */
        QProgressBar#cmBar[stopped="true"],
        QProgressBar#mhBar[stopped="true"],
        QProgressBar#ovBar[stopped="true"] {
            background: #e0e0e0 !important;    /* cancel any red background rules */
        }
        
        QProgressBar#cmBar[stopped="true"]::chunk,
        QProgressBar#mhBar[stopped="true"]::chunk,
        QProgressBar#ovBar[stopped="true"]::chunk {
            background-color: #B30000 !important;  /* red fill only */
        }
        """
        
        # apply to each bar
        self.current_manufacturer_progress.setStyleSheet(progress_css)
        self.manufacturer_hyperlink_bar.setStyleSheet(progress_css)
        self.overall_progress_bar.setStyleSheet(progress_css)
        
        # after adding all widgets and layouts…
        self.si_mode_toggle.stateChanged.connect(self.on_si_mode_toggled)

        # set initial enabled/disabled state based on default toggle
        self.on_si_mode_toggled(self.mode_switch.checkState())

        self.setLayout(layout)
        self.resize(600, 400)

    def handle_extractor_output(self, line: str):
        """
        Consume stdout from SharepointExtractor. Update UI progress bars,
        but optionally suppress CM progress lines from the Terminal view.
        """
        # ── 1) Current Manufacturer progress coming as explicit "CM_PROGRESS a/b (p%)" ──
        m_cm = re.match(r"\s*CM_PROGRESS\s+(\d+)\s*/\s*(\d+)\s*\((\d+)%\)", line, re.IGNORECASE)
        if m_cm:
            done = int(m_cm.group(1))
            total = max(1, int(m_cm.group(2)))
            pct_from_text = int(m_cm.group(3))
            # trust the numbers coming from the extractor, but clamp
            pct = max(0, min(100, pct_from_text if 0 <= pct_from_text <= 100 else int(done/total*100)))
            self.current_manufacturer_progress.setValue(pct)
    
            # optionally suppress showing this in the terminal
            if self.hide_cm_progress_in_terminal:
                return
    
        # ── 2) Cleanup-mode progress (broken link repair) ──
        if getattr(self, '_cleanup_mode', False):
            # total to fix
            m_total = re.search(r'Total broken hyperlinks:\s*(\d+)', line)
            if m_total:
                self._initial_broken = int(m_total.group(1))
                self._fixed_count = 0
                if self.hide_cm_progress_in_terminal:
                    return
            # each item processed (we treat ✅/❌/“Fixed hyperlink for …” as a tick)
            if line.startswith(("Fixed hyperlink for", "✅", "❌")) and getattr(self, "_initial_broken", None):
                self._fixed_count += 1
                pct = int(self._fixed_count / max(1, self._initial_broken) * 100)
                self.current_manufacturer_progress.setValue(pct)
                if self.hide_cm_progress_in_terminal:
                    return
    
        # ── 3) Normal mode progress using "N Folders Remain" ──
        m_fr = re.search(r'(\d+)\s+Folders Remain', line)
        if m_fr:
            remaining = int(m_fr.group(1))
            if not hasattr(self, '_initial_folder_count') or self._initial_folder_count is None:
                self._initial_folder_count = remaining
            else:
                self._initial_folder_count = max(self._initial_folder_count, remaining)
            initial = max(1, self._initial_folder_count)
            pct = max(0, min(100, int((initial - remaining) / initial * 100)))
            self.current_manufacturer_progress.setValue(pct)
            # we **do** still show these lines normally (they're not CM_PROGRESS),
            # so no early return here.
    
        # finally, append anything not suppressed
        self.terminal.append_output(line)

    def mark_manual_stop(self):
        """
        Reflect a manual stop in labels and progress bars (works for Cleanup + Regular).
        Does NOT touch buttons; your on_start_stop() already handles swapping Start/Stop.
        """
        # Labels
        if hasattr(self, "current_manufacturer_label"):
            self.current_manufacturer_label.setText("Current Manufacturer: Manually Stopped")
        if hasattr(self, "manufacturer_hyperlink_label"):
            self.manufacturer_hyperlink_label.setText("Manufacturer Hyperlinks Indexed: Manually Stopped")
        if hasattr(self, "overall_progress_label"):
            self.overall_progress_label.setText("Overall Progress: Manually Stopped")
    
        # Progress bars → reset to 0
        for bar_name in ("current_manufacturer_progress", "manufacturer_hyperlink_bar", "overall_progress_bar"):
            bar = getattr(self, bar_name, None)
            if bar:
                try:
                    bar.setMaximum(100)  # in case it was set to a link-count
                    bar.setValue(0)
                except Exception:
                    pass
    
        # Clear any counters so next run starts fresh
        for attr in ("_initial_broken", "_fixed_count", "_initial_folder_count",
                     "_hyperlinks_total_links", "_hyperlinks_done_links"):
            if hasattr(self, attr):
                setattr(self, attr, None)
       
    def _style_single_bar(self, bar: QProgressBar, stopped: bool) -> None:
        if not bar:
            return
        bar.setStyleSheet(self._bar_style_stopped if stopped else self._bar_style_normal)
    
    def _apply_stopped_style_to_all_bars(self, stopped: bool):
        """
        Ensure bars use the baseline CSS (no lingering red background),
        then flip only the CHUNK to red via the 'stopped' dynamic property.
        """
        bars = (
            getattr(self, "current_manufacturer_progress", None),
            getattr(self, "manufacturer_hyperlink_bar", None),
            getattr(self, "overall_progress_bar", None),
        )
        for bar in bars:
            if not bar:
                continue
    
            # 1) HARD RESET any previous per-widget overrides that might have set a red background
            bar.setStyleSheet("")
            bar.setStyleSheet(getattr(self, "_progress_css", ""))
    
            # 2) Flip the property to switch the chunk color (groove stays grey)
            bar.setProperty("stopped", bool(stopped))
            bar.style().unpolish(bar)
            bar.style().polish(bar)
            bar.update()
    
            # 3) If value==0, flash 1→0 so you can immediately see the red chunk rule working
            if stopped and bar.value() == 0:
                bar.setValue(1)
                bar.setValue(0)
          
    def on_si_mode_toggled(self, state):
        """Enable one list & button, disable—and clear—the other."""
        is_repair = (state == Qt.Checked)
    
        # Repair group gets enabled; ADAS gets disabled & cleared
        for cb in self.repair_checkboxes:
            cb.setEnabled(is_repair)
        if not is_repair:
            for cb in self.repair_checkboxes:
                cb.setChecked(False)
        self.select_all_repair_button.setEnabled(is_repair)
        if not is_repair:
            self.select_all_repair_button.setChecked(False)
    
        # ADAS group gets enabled; Repair gets disabled & cleared
        for cb in self.adas_checkboxes:
            cb.setEnabled(not is_repair)
        if is_repair:
            for cb in self.adas_checkboxes:
                cb.setChecked(False)
        self.select_all_adas_button.setEnabled(not is_repair)
        if is_repair:
            self.select_all_adas_button.setChecked(False)
        # ✅ Disable Excel Format toggle if Repair SI is active
        # ✅ Disable Excel Format toggle and reset to OG when Repair SI is active
        if self.excel_mode_switch:
            if is_repair:
                self.excel_mode_switch.setChecked(False)   # ← Reset to OG
                self.excel_mode_switch.setEnabled(False)   # ← Gray out
            else:
                self.excel_mode_switch.setEnabled(True)

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

    def select_all_repair(self):
        select_all_checked = all(checkbox.isChecked() for checkbox in self.repair_checkboxes)
        for checkbox in self.repair_checkboxes:
            checkbox.setChecked(not select_all_checked)
    
    def select_excel_files(self):
        self.excel_paths, _ = QFileDialog.getOpenFileNames(
            self, 'Open files', 'C:/Users/', "Excel files (*.xlsx *.xls)"
        )
        if self.excel_paths:
            # Trim any stray whitespace
            self.excel_paths = [p.strip() for p in self.excel_paths]
    
            # 1) Show numbered filenames in the label
            numbered = [f"{i+1}. {os.path.basename(p)}"
                        for i, p in enumerate(self.excel_paths)]
            self.excel_path_label.setText("\n".join(numbered))
    
            # 2) Fill the scrollable list with the same numbering
            self.excel_list.clear()
            for i, p in enumerate(self.excel_paths):
                self.excel_list.addItem(f"{i+1}. {os.path.basename(p)}")
    
        else:
            # No files chosen: clear both widgets
            self.excel_path_label.setText('No files selected')
            self.excel_list.clear()
            self.excel_list.addItem('No files selected, please select files')
         
    def toggle_theme(self):
        if self.theme_toggle.isChecked():
            # light
            self.setStyleSheet("background-color: #ffffff; color: black;")
            self.manufacturer_tree.setStyleSheet(
                "background-color: #f0f0f0; color: black; "
                "border: 1px solid #cccccc; border-radius: 5px;"
            )
            self.excel_path_label.setStyleSheet(
                "font-size: 14px; padding: 5px; "
                "border: 1px solid #cccccc; border-radius: 5px; "
                "background-color: #f0f0f0;"
            )
            self.excel_list.setStyleSheet(
                "font-size: 14px; padding: 5px; "
                "border: 1px solid #cccccc; border-radius: 5px; "
                "background-color: #f0f0f0; color: black;"
            )
            self.repair_scroll_area.setStyleSheet(
                "background-color: #f0f0f0; "
                "border: 1px solid #cccccc; border-radius: 5px;"
            )
        else:
            # dark
            self.setStyleSheet("background-color: #2e2e2e; color: white;")
            self.manufacturer_tree.setStyleSheet(
                "background-color: #3e3e3e; color: white; "
                "border: 1px solid #555555; border-radius: 5px;"
            )
            self.excel_path_label.setStyleSheet(
                "font-size: 14px; padding: 5px; "
                "border: 1px solid #555555; border-radius: 5px; "
                "background-color: #3e3e3e;"
            )
            self.excel_list.setStyleSheet(
                "font-size: 14px; padding: 5px; "
                "border: 1px solid #555555; border-radius: 5px; "
                "background-color: #3e3e3e; color: white;"
            )
            self.repair_scroll_area.setStyleSheet(
                "background-color: #3e3e3e; "
                "border: 1px solid #555555; border-radius: 5px;"
            )

    def start_automation(self):
        # 1) gather selected manufacturers
        selected_manufacturers = []
        for i in range(self.manufacturer_tree.topLevelItemCount()):
            item = self.manufacturer_tree.topLevelItem(i)
            if item.checkState(0) == Qt.Checked:
                selected_manufacturers.append(item.text(0))

        # 2) gather selected systems based on the slide‐toggle
        if self.mode_switch.isChecked():   # Repair mode
            selected_systems = [cb.text() for cb in self.repair_checkboxes if cb.isChecked()]
        else:                              # ADAS mode
            selected_systems = [cb.text() for cb in self.adas_checkboxes if cb.isChecked()]

        # 3) sanity check
        if not (self.excel_paths and selected_manufacturers and selected_systems):
            QMessageBox.warning(self, 'Warning',
                "Please select Excel files, manufacturers, and at least one system.", QMessageBox.Ok)
            return

        # 4) confirm and kick off
        # — build Excel list
        excel_list = "\n".join(f"{i+1}. {os.path.basename(path)}"
                               for i, path in enumerate(self.excel_paths))
        # — build manufacturers list
        manu_list = "\n".join(f"{i+1}. {m}"
                              for i, m in enumerate(selected_manufacturers))
        
        cleanup_note = ""
        if self.cleanup_checkbox.isChecked():
            cleanup_note = (
                "\n\n⚠️ Broken Hyperlink Mode Activated:\n"
                "With this selected, it will ignore all the ADAS/Repair arguments\n"
                "and find the broken links. Based off of those results, it will\n"
                "find the matching links and repair them."
            )
        
        # detect Excel Format (ADAS SI / Repair SI)
        excel_format = "Repair SI" if self.mode_switch.isChecked() else "ADAS SI"
        
        # detect Version Format (OG / NEW)
        version_format = "NEW" if self.excel_mode_switch.isChecked() else "OG"
        
        confirm_message = (
            "Excel files selected:\n"
            f"{excel_list}\n\n"
            "Manufacturers selected:\n"
            f"{manu_list}\n\n"
            "Systems selected:\n"
            + ", ".join(selected_systems) + "\n\n"
            "Excel Format:\n"
            f"{excel_format}\n\n"
            "Version Format:\n"
            f"{version_format}"
            + cleanup_note + "\n\nContinue?"
        )
                
        

        if QMessageBox.question(self, 'Confirmation', confirm_message,
               QMessageBox.Yes | QMessageBox.No, QMessageBox.No) != QMessageBox.Yes:
            return
        
        # user clicked YES → mark running
        self.is_running     = True
        self.stop_requested = False
        
        # rip out the old “Start” button and insert a red “Stop Automation”
        # ── swap Start → Stop inside our vertical button_layout ──
        layout = self.button_layout
        layout.removeWidget(self.start_button)
        self.start_button.deleteLater()
        self.start_button = CustomButton("Stop Automation", "#e63946", self)
        self.start_button.clicked.connect(self.on_start_stop)
        layout.addWidget(self.start_button)
        

        # ── step 3: enable the Pause button when we start ──
        self.pause_button.setEnabled(True)
        self.pause_button.setText('Pause Automation')
        self.pause_requested = False

        # now proceed with the rest of your existing automation logic…

        # 5) stash for process_next_manufacturer
        self.selected_manufacturers = selected_manufacturers
        self.selected_systems       = selected_systems
        self.mode_flag              = "repair" if self.mode_switch.isChecked() else "adas"
        self.current_index          = 0
        
        # ── Initialize Overall Progress ──
        self.total_manufacturers = len(self.selected_manufacturers)
        # ── NEW: reset progress bars for new run ──
        self.overall_progress_bar.setValue(0)
        self.current_manufacturer_progress.setValue(0)
        self.overall_progress_label.setText(f"Overall Progress: 0 / {self.total_manufacturers}")
        self.current_manufacturer_label.setText("Current Manufacturer: None")

        # 6) show or reuse terminal & start
        if getattr(self, 'terminal', None) is None or not self.terminal.isVisible():
            self.terminal = TerminalDialog(self)

            # ── MONKEY‐PATCH for live logging ──
            _orig_append = self.terminal.append_output
            def _live_append(text: str):
                _orig_append(text)       # write to on‐screen terminal
                logging.info(text)       # write to logfile
            self.terminal.append_output = _live_append

        self.terminal.show()
        self.terminal.raise_()
        self.process_next_manufacturer()

    def process_next_manufacturer(self):
        # ── STOP BAILOUT ──
        if self.stop_requested:
            return
    
        if self.current_index >= len(self.selected_manufacturers):
            # 🆕 If cleanup mode, run final unresolved broken link removal
            if self.cleanup_checkbox.isChecked():
                try:
                    if hasattr(self, "extractor") and hasattr(self.extractor, "broken_entries"):
                        # Determine hyperlink column based on mode
                        if self.extractor.repair_mode and self.extractor.excel_mode == "og":
                            hyperlink_col = 8
                        elif not self.extractor.repair_mode and self.extractor.excel_mode == "og":
                            hyperlink_col = 12
                        elif not self.extractor.repair_mode and self.extractor.excel_mode == "new":
                            hyperlink_col = 11
                        else:
                            hyperlink_col = None
        
                        if hyperlink_col:
                            print("🧹 Finalizing cleanup — removing unresolved broken links...")
                            import openpyxl
                            wb = openpyxl.load_workbook(self.excel_paths[0])
                            ws = wb['Model Version']
                            removed_count = 0
                            for row, (yr, mk, mdl, sys) in self.extractor.broken_entries:
                                cell = ws.cell(row=row, column=hyperlink_col)
                                link_to_test = cell.hyperlink.target if cell.hyperlink else str(cell.value).strip() if cell.value else ""
                                if not link_to_test or link_to_test.lower() == "hyperlink not available":
                                    continue
                                if self.extractor.is_broken_sharepoint_link(link_to_test, file_name=sys):
                                    print(f"🗑 Removing unresolved broken link at row {row}: {link_to_test}")
                                    cell.value = None
                                    cell.hyperlink = None
                                    removed_count += 1
                            wb.save(self.excel_paths[0])
                            wb.close()
                            print(f"✅ Cleanup complete — removed {removed_count} unresolved links")
                except Exception as e:
                    print(f"⚠️ Final cleanup pass failed: {e}")
        
            # done!
            completed = "\n".join(sorted(self.selected_manufacturers, key=str.lower))
            QMessageBox.information(
                self,
                'Completed',
                f"The Following Manufacturers have been completed:\n{completed}",
                QMessageBox.Ok
            )
            return
        
            
        # Reset per‐manufacturer progress tracking
        self._initial_folder_count = None
    
        manufacturer = self.selected_manufacturers[self.current_index]
        self.current_manufacturer_label.setText(f"Current Manufacturer: {manufacturer}")
        self.current_manufacturer_progress.setValue(0)
    
        excel_path = self.excel_paths[self.current_index]
        link_dict = self.repair_links if self.mode_flag == "repair" else self.manufacturer_links
    
        # Get all SharePoint links for this manufacturer (could be 1 or many)
        sharepoint_links = link_dict.get(manufacturer)
        if not sharepoint_links:
            QMessageBox.warning(
                self,
                'Error',
                f"No SharePoint link found for {manufacturer} in {self.mode_flag} mode.",
                QMessageBox.Ok
            )
            return
    
        # Normalize to list if it's a single string
        if isinstance(sharepoint_links, str):
            sharepoint_links = [sharepoint_links]
    
        # 🆕 Cleanup Mode: Filter only the links matching the years from broken hyperlinks
        if self.cleanup_checkbox.isChecked():
            years_needed = self.get_broken_hyperlink_years_for_manufacturer(manufacturer)
            filtered_links = []
            for link in sharepoint_links:
                m = re.search(r'\((\d{4})\s*-\s*(\d{4})\)', link)
                if m:
                    start_year, end_year = int(m.group(1)), int(m.group(2))
                    if any(start_year <= y <= end_year for y in years_needed):
                        filtered_links.append(link)
            if filtered_links:
                sharepoint_links = filtered_links
    
        # Store state for multi-link handling
        self._multi_links        = sharepoint_links
        self._multi_link_index   = 0
        self._multi_excel_path   = excel_path
        self._multi_manufacturer = manufacturer
    
        # 🆕 Manufacturer hyperlink counter based on # of SharePoint links
        self._hyperlinks_total_links = len(self._multi_links)
        self._hyperlinks_done_links = 0
        self.update_manufacturer_progress_bar()
    
        # NEW: remember cleanup mode & reset its counters
        self._cleanup_mode = self.cleanup_checkbox.isChecked()
        if self._cleanup_mode:
            self._initial_broken = None
            self._fixed_count    = 0
    
        # Start the first sub-link run
        self.run_next_sub_link()
    
    def run_next_sub_link(self):
        if self._multi_link_index >= len(self._multi_links):
            # All links processed for this manufacturer
            self.on_manufacturer_finished(self._multi_manufacturer, True)
            return
    
        # Reset Current Manufacturer progress bar for this sub-link
        self.current_manufacturer_progress.setValue(0)
        self._initial_folder_count = None  # 🆕 Reset baseline for "Folders Remain"
    
        script_path = os.path.join(os.path.dirname(__file__), "SharepointExtractor.py")
        excel_mode = "new" if self.excel_mode_switch.isChecked() else "og"
    
        current_link = self._multi_links[self._multi_link_index]
        args = [
            sys.executable,
            script_path,
            current_link,
            self._multi_excel_path,
            ",".join(self.selected_systems),
            self.mode_flag,
            "cleanup" if self.cleanup_checkbox.isChecked() else "full",
            excel_mode
        ]
    
        self._cleanup_mode = self.cleanup_checkbox.isChecked()
        if self._cleanup_mode:
            self._initial_broken = None
            self._fixed_count = 0
    
        self.update_manufacturer_progress_bar()
    
        thread = WorkerThread(args, self._multi_manufacturer, parent=self)
        self.thread = thread
        thread.output_signal.connect(self.handle_extractor_output)
        thread.finished_signal.connect(self.on_sub_link_finished)
        thread.start()
        self.threads.append(thread)

    def finalize_cleanup_for_file(self, excel_path, broken_entries, hyperlink_col):
        import openpyxl
        wb = openpyxl.load_workbook(excel_path)
        ws = wb['Model Version']
    
        removed_count = 0
        for row, (yr, mk, mdl, sys) in broken_entries:
            cell = ws.cell(row=row, column=hyperlink_col)
    
            # Prefer the actual hyperlink target if it exists
            link_to_test = cell.hyperlink.target if cell.hyperlink else str(cell.value).strip() if cell.value else ""
    
            # Skip placeholders or empty links
            if not link_to_test or link_to_test.lower() == "hyperlink not available":
                continue
    
            # Check link validity
            if self.extractor.is_broken_sharepoint_link(link_to_test, file_name=sys):
                print(f"🗑 Removing unresolved broken link at row {row}: {link_to_test}")
                cell.value = None
                cell.hyperlink = None
                removed_count += 1
    
        wb.save(excel_path)
        wb.close()
        print(f"✅ Cleanup complete — removed {removed_count} unresolved links")
           
    def count_expected_hyperlinks_for_link(self, manufacturer, sharepoint_link):
        from openpyxl import load_workbook
        import re
    
        # Extract year range from link (assuming in format "(2012 - 2016)")
        m = re.search(r'\((\d{4})\s*-\s*(\d{4})\)', sharepoint_link)
        year_range = None
        if m:
            year_range = (int(m.group(1)), int(m.group(2)))
    
        wb = load_workbook(self.excel_paths[0])
        ws = wb.active
        count = 0
    
        for row in ws.iter_rows(min_row=2):
            year_val = row[0].value
            make_val = str(row[1].value).strip().lower() if row[1].value else ""
            link_val = row[11].hyperlink if len(row) > 11 else None  # Column L
    
            if make_val == manufacturer.lower():
                if year_range and isinstance(year_val, int) and not (year_range[0] <= year_val <= year_range[1]):
                    continue
                if not link_val:
                    count += 1
    
        return count
        
    def on_sub_link_finished(self, manufacturer, success):
        if not self.is_running:
            return
    
        if success:
            msg = f"✅ Finished SharePoint link {self._multi_link_index+1}/{len(self._multi_links)} for {manufacturer}"
            self.terminal.append_output(msg)
            logging.info(msg)
    
            # Increment completed links count & update bar
            self._hyperlinks_done_links += 1
            self.update_manufacturer_progress_bar()
    
            self._multi_link_index += 1
    
            if self._multi_link_index >= len(self._multi_links):
                # All links for this manufacturer complete
                self.manufacturer_hyperlink_label.setText("Manufacturer Hyperlinks Indexed: Completed")
                self.manufacturer_hyperlink_bar.setValue(self.manufacturer_hyperlink_bar.maximum())
                self.on_manufacturer_finished(manufacturer, True)
            else:
                # Move to next sub-link
                self.run_next_sub_link()
    
        else:
            msg = f"❌ SharePoint link {self._multi_link_index+1}/{len(self._multi_links)} for {manufacturer} failed"
            self.terminal.append_output(msg)
            logging.warning(msg)
            self.on_manufacturer_finished(manufacturer, False)
    
    def update_manufacturer_progress_bar(self):
        total_links = getattr(self, "_hyperlinks_total_links", 1)
        done_links = getattr(self, "_hyperlinks_done_links", 0)
    
        # Keep bar in sync
        self.manufacturer_hyperlink_bar.setMaximum(total_links)
        self.manufacturer_hyperlink_bar.setValue(done_links)
    
        # Always show x/y format
        self.manufacturer_hyperlink_label.setText(
            f"Manufacturer Hyperlinks Indexed: {done_links}/{total_links}"
        )
      
        
    def get_broken_hyperlink_years_for_manufacturer(self, manufacturer):
        years = set()
        from openpyxl import load_workbook
    
        wb = load_workbook(self.excel_paths[0])
        ws = wb.active
    
        for row in ws.iter_rows(min_row=2):
            year_val = row[0].value  # Assuming column A = Year
            make_val = str(row[1].value).strip().lower() if row[1].value else ""
            link_val = row[11].hyperlink if len(row) > 11 else None  # Assuming column L for hyperlink
    
            if make_val == manufacturer.lower():
                # No hyperlink OR broken link
                if not link_val or self.is_broken_sharepoint_link(link_val.target):
                    try:
                        years.add(int(year_val))
                    except (TypeError, ValueError):
                        pass
    
        return sorted(years)
    
    def on_manufacturer_finished(self, manufacturer, success):
        # ── HARD BAIL-OUT: if we're not running, stop immediately ──
        if not self.is_running:
            # Reset UI to Stopped state
            self.current_manufacturer_progress.setValue(0)
            self.overall_progress_bar.setValue(0)
            self.current_manufacturer_label.setText("Current Manufacturer: Manually Stopped")
            self.overall_progress_label.setText("Overall Progress: Manually Stopped")
    
            # Swap back to Start button
            layout = self.button_layout
            layout.removeWidget(self.start_button)
            self.start_button.deleteLater()
            self.start_button = CustomButton("Start Automation", "#008000", self)
            self.start_button.clicked.connect(self.on_start_stop)
            layout.addWidget(self.start_button)
    
            # Disable Pause
            self.pause_button.setEnabled(False)
            return
    
        # ── YOUR ORIGINAL LOGIC STARTS HERE ──
    
        # 1) count this run
        prev = self.attempts.get(manufacturer, 0)
        self.attempts[manufacturer] = prev + 1
        attempt_no = self.attempts[manufacturer]
    
        # 2) route based on success / attempt count
        if success:
            self.completed_manufacturers.append(manufacturer)
            msg = f"✅ {manufacturer} succeeded on attempt {attempt_no}."
            self.terminal.append_output(msg)
            logging.info(msg)
    
            # update overall on success
            finalized = len(self.completed_manufacturers) + len(self.given_up_manufacturers)
            percent   = int(finalized / self.total_manufacturers * 100)
            self.overall_progress_bar.setValue(percent)
            self.overall_progress_label.setText(
                f"Overall Progress: {finalized} / {self.total_manufacturers}"
            )
        else:
            if attempt_no < self.max_attempts:
                err_excel = self.excel_paths[self.current_index]
                self.failed_manufacturers.append(manufacturer)
                self.failed_excels.append(err_excel)
                msg = f"❗ {manufacturer} failed on attempt {attempt_no}; will retry later."
                self.terminal.append_output(msg)
                logging.warning(msg)
            else:
                self.given_up_manufacturers.append(manufacturer)
                msg = (
                    f"❌ {manufacturer} failed on attempt {attempt_no}; "
                    f"giving up after {self.max_attempts} tries."
                )
                self.terminal.append_output(msg)
                logging.error(msg)
    
                # update overall on final give-up
                finalized = len(self.completed_manufacturers) + len(self.given_up_manufacturers)
                percent   = int(finalized / self.total_manufacturers * 100)
                self.overall_progress_bar.setValue(percent)
                self.overall_progress_label.setText(
                    f"Overall Progress: {finalized} / {self.total_manufacturers}"
                )
    
        # 3) pause, then advance index
        msg = "⏱ Checking in 10s if i Need to run another Manufacturer…"
        self.terminal.append_output(msg)
        logging.info(msg)
        sleep(10)
        self.current_index += 1
    
        # 4) if still in this pass, keep going
        if self.current_index < len(self.selected_manufacturers):
            self.process_next_manufacturer()
            return
    
        # 5) end of pass: retry logic…
        if self.failed_manufacturers:
            retry_list = ", ".join(self.failed_manufacturers)
            self.terminal.append_output(f"🔄 Retrying: {retry_list}")
            sleep(10)
            self.selected_manufacturers = self.failed_manufacturers
            self.excel_paths            = self.failed_excels
            self.current_index          = 0
            self.failed_manufacturers   = []
            self.failed_excels          = []
            self.process_next_manufacturer()
            return
    
        # 6) final summary
        completed_sorted = sorted(self.completed_manufacturers, key=str.lower)
        given_up_sorted  = sorted(self.given_up_manufacturers,  key=str.lower)
    
        self.terminal.append_output("")
        self.terminal.append_output("🏁 All Manufacturers finished.")
        self.terminal.append_output(f"✅ Completed: {', '.join(completed_sorted)}")
        self.terminal.append_output(f"❌ Gave up:   {', '.join(given_up_sorted)}")
    
        # lock bars at 100%
        self.current_manufacturer_progress.setValue(100)
        self.overall_progress_bar.setValue(100)
        self.current_manufacturer_label.setText("Current Manufacturer: Complete")
        self.overall_progress_label.setText("Overall Progress: Complete")
        self.terminal.append_output("=" * 66)
    
        # reset tracking for next run
        self.completed_manufacturers    = []
        self.failed_manufacturers       = []
        self.failed_excels              = []
        self.given_up_manufacturers     = []
        self.attempts                   = {}
    
        # swap back to Start button
        layout = self.button_layout
        layout.removeWidget(self.start_button)
        self.start_button.deleteLater()
        self.start_button = CustomButton("Start Automation", "#008000", self)
        self.start_button.clicked.connect(self.on_start_stop)
        layout.addWidget(self.start_button)
        self.pause_button.setEnabled(False)
        self.is_running = False

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

    def closeEvent(self, event):
        # when the GUI closes, dump the terminal contents (if any) to Documents/Logs
        super().closeEvent(event)
     
    def _set_bar_stopped(self, bar, stopped: bool, name_hint: str):
        """
        Make a single QProgressBar turn red when stopped, while preserving its exact original look.
        Uses a per-widget stylesheet with higher specificity and !important so global styles can't override it.
        """
        if not bar:
            return
    
        # Cache original style once
        if not hasattr(bar, "_orig_stylesheet"):
            bar._orig_stylesheet = bar.styleSheet() or ""
    
        # Ensure we can target this bar specifically
        if not bar.objectName():
            bar.setObjectName(name_hint)
    
        if stopped:
            base = bar._orig_stylesheet
            obj  = bar.objectName()
            # Per‑widget rules (highest precedence) + !important
            red_css = (
                f'QProgressBar#{obj} {{ '
                f'  background-color: #470000 !important; '   # deep red even at 0%
                f'  color: white !important; '
                f'  border: 1px solid #aa4444; '
                f'  border-radius: 4px; '
                f'}} '
                f'QProgressBar#{obj}::chunk {{ '
                f'  background-color: #e53935 !important; '   # red fill when >0%
                f'}}'
            )
            bar.setStyleSheet(base + "\n" + red_css)
        else:
            # Restore exactly what the bar had before
            bar.setStyleSheet(bar._orig_stylesheet)
    
        # Force immediate repaint
        bar.style().unpolish(bar)
        bar.style().polish(bar)
        bar.update()
    
    def _style_bar(self, bar, *, stopped: bool, name_hint: str):
        """Per-widget stylesheet: groove stays grey; only the chunk color changes."""
        if not bar:
            return
        if not bar.objectName():
            bar.setObjectName(name_hint)
        obj = bar.objectName()
    
        normal_css = f"""
        QProgressBar#{obj} {{
            font-size: 12px; padding: 4px; text-align: center;
            color: black; border: 1px solid #555; border-radius: 4px;
            background: #e0e0e0;                /* grey groove */
        }}
        QProgressBar#{obj}::chunk {{
            background-color: #19A602;           /* GREEN fill in normal runs */
            margin: 0px; border-radius: 3px;
        }}
        """
    
        stopped_css = f"""
        QProgressBar#{obj} {{
            font-size: 12px; padding: 4px; text-align: center;
            color: black; border: 1px solid #555; border-radius: 4px;
            background: #e0e0e0 !important;      /* keep grey groove when stopped */
        }}
        QProgressBar#{obj}::chunk {{
            background-color: #B30000 !important;/* RED fill when stopped */
            margin: 0px; border-radius: 3px;
        }}
        """
    
        bar.setStyleSheet(stopped_css if stopped else normal_css)
        bar.style().unpolish(bar); bar.style().polish(bar); bar.update()
    
    def _force_zero_red(self, bar, enable: bool, full: bool = True):
        """
        When enable=True:
          - keep the text at "0%" (format override)
          - but set the value to full width (100%) so the red CHUNK fills the bar
        When enable=False:
          - restore the original format and value
        """
        if not bar:
            return
    
        if enable:
            if not hasattr(bar, "_orig_format"):
                bar._orig_format = bar.format()
            if not hasattr(bar, "_orig_value"):
                bar._orig_value = bar.value()
    
            bar.setFormat("0%")  # force the text to show 0%
            target = bar.maximum() if full else max(1, bar.maximum() // 100)
            bar.setValue(target)  # fill the bar (chunk draws full-width in red)
            bar._forced_zero_red = True
        else:
            if getattr(bar, "_forced_zero_red", False):
                bar.setFormat(getattr(bar, "_orig_format", "%p%"))
                bar.setValue(getattr(bar, "_orig_value", 0))
                bar._forced_zero_red = False
           
    def _apply_stopped_style_to_all_bars(self, stopped: bool):
        bars = (
            getattr(self, "current_manufacturer_progress", None),
            getattr(self, "manufacturer_hyperlink_bar", None),
            getattr(self, "overall_progress_bar", None),
        )
        for name_hint, bar in zip(("cmBar","mhBar","ovBar"), bars):
            self._style_bar(bar, stopped=stopped, name_hint=name_hint)
            # at manual stop, force a visible red sliver while showing "0%"
            self._force_zero_red(bar, enable=stopped)
       
    def on_start_stop(self):
        # — START path —
        if not self.is_running:
            # Try to start; this might show a confirmation and bail out.
            self.start_automation()
    
            # ⛔ If user clicked "No" (or start failed), do NOTHING to the bars/styles.
            # Leave the "Manually Stopped" red state intact.
            if not self.is_running:
                return
    
            # ✅ We are actually starting now → safe to restore normal styles
            self.pause_requested = False
            self.pause_button.setText('Pause Automation')
            self.pause_button.setEnabled(True)
    
            # Back to normal look for a fresh run
            self._apply_stopped_style_to_all_bars(False)
            return
    
        # — STOP path —   (unchanged)
        reply = QMessageBox.question(
            self,
            "Confirm Stop",
            "Are you sure you want to end this automation?",
            QMessageBox.Yes | QMessageBox.No
        )
        if reply != QMessageBox.Yes:
            return
    
        # If we were paused, first resume the subprocess so it can be cleanly terminated
        if self.pause_requested and self.thread is not None and hasattr(self.thread, "process"):
            try:
                proc = psutil.Process(self.thread.process.pid)
                proc.resume()
                for child in proc.children(recursive=True):
                    child.resume()
            except Exception as e:
                self.terminal.append_output(f"⚠️ Couldn’t resume before stopping: {e}")
    
        # clear pause flag and tell loops not to launch any more work
        self.pause_requested = False
        self.stop_requested  = True
    
        # 1) Ask the Python extractor to shut down nicely
        if self.thread is not None and hasattr(self.thread, "process"):
            try:
                if os.name == "nt":
                    self.thread.process.send_signal(signal.CTRL_BREAK_EVENT)
                else:
                    os.killpg(os.getpgid(self.thread.process.pid), signal.SIGTERM)
            except Exception:
                pass
    
            # 2) Fallback: kill Chrome/Chromedriver children and parent
            def kill_children(pid: int):
                try:
                    parent = psutil.Process(pid)
                except psutil.NoSuchProcess:
                    return
                for child in parent.children(recursive=True):
                    try:
                        pname = child.name().lower()
                    except psutil.NoSuchProcess:
                        continue
                    if "chrome" in pname or "chromedriver" in pname:
                        try: child.kill()
                        except psutil.NoSuchProcess: pass
                try: parent.kill()
                except psutil.NoSuchProcess: pass
    
            kill_children(self.thread.process.pid)
    
        # ── reset & disable Pause/Resume button when stopping ──
        self.pause_button.setText('Pause Automation')
        self.pause_button.setEnabled(False)
    
        # 3) Give it a moment, then report & swap button back
        sleep(1)
        self.terminal.append_output("❌ Hyperlink Automation has stopped.")
    
        # Show 'Manually Stopped' + reset and paint bars red
        if hasattr(self, "current_manufacturer_label"):
            self.current_manufacturer_label.setText("Current Manufacturer: Manually Stopped")
        if hasattr(self, "manufacturer_hyperlink_label"):
            self.manufacturer_hyperlink_label.setText("Manufacturer Hyperlinks Indexed: Manually Stopped")
        if hasattr(self, "overall_progress_label"):
            self.overall_progress_label.setText("Overall Progress: Manually Stopped")
    
        if hasattr(self, "current_manufacturer_progress"):
            self.current_manufacturer_progress.setValue(0)
        if hasattr(self, "manufacturer_hyperlink_bar"):
            self.manufacturer_hyperlink_bar.setValue(0)
        if hasattr(self, "overall_progress_bar"):
            self.overall_progress_bar.setValue(0)
    
        # turn bars red (and force a repaint at 0%)
        self._apply_stopped_style_to_all_bars(True)
    
        # ── swap back to a fresh “Start Automation” button ──
        layout = self.button_layout
        layout.removeWidget(self.start_button)
        self.start_button.deleteLater()
        self.start_button = CustomButton("Start Automation", "#008000", self)
        self.start_button.clicked.connect(self.on_start_stop)
        layout.addWidget(self.start_button)
    
        self.pause_button.setEnabled(False)
    
        # ── NEW: clear running state so clicks now start again ──
        self.is_running     = False
        self.stop_requested = False
    
        # Ensure it goes above the progress bars
        insert_index = layout.indexOf(self.current_manufacturer_label)
        layout.insertWidget(insert_index, self.start_button)
     
    def on_pause_resume(self):
        # only when running and we have a live subprocess
        if not self.is_running or self.thread is None or not hasattr(self.thread, "process"):
            return
    
        proc = psutil.Process(self.thread.process.pid)
    
        if not self.pause_requested:
            # ── PAUSE ──
            self.pause_requested = True
            self.pause_button.setText('Resume Automation')
            self.terminal.append_output("⏸️ Pausing automation…")
    
            # suspend main process & children
            try:
                proc.suspend()
                for child in proc.children(recursive=True):
                    child.suspend()
            except Exception as e:
                self.terminal.append_output(f"⚠️ Couldn’t pause process: {e}")
    
        else:
            # ── RESUME ──
            self.pause_requested = False
            self.pause_button.setText('Pause Automation')
            self.terminal.append_output("▶️ Resuming automation…")
    
            # resume main process & children
            try:
                proc.resume()
                for child in proc.children(recursive=True):
                    child.resume()
            except Exception as e:
                self.terminal.append_output(f"⚠️ Couldn’t resume process: {e}")
    
            # <-- no call to process_next_manufacturer() here!
            # the suspended extractor will continue its own loop

if __name__ == "__main__":
    try:
        app = QApplication(sys.argv)
        window = SeleniumAutomationApp()
        window.show()
        sys.exit(app.exec_())
    except Exception:
        logging.exception("Unhandled exception — crashing out")
        raise
