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

class SeleniumAutomationApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.terminal = None           
        self.excel_paths = []
        self.manufacturer_links = {
            # Add ADAS SI Sharepoint Links here
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
            # Add ADAS SI Sharepoint Links here
        }
        self.repair_links = {
            # Add Repair SI SharePoint links Here
            "Acura": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/EmmQxHGQ2wxNpIO04pNu_iIBqou8brkQWpKnjHPgSFT-CQ?e=qzTX1G",
            "Alfa Romeo": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/EqZpuA6CO9JEooJ3K1Ix5_wB_T0XPqsH77cK8fYmidKF7g?e=cQ1HDy",
            "Audi": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/EgRYXyielJhMrOCaEhv-snQBaKmKxSPpH03Xa-mensN5VA?e=qngPQ7",
            "BMW": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/ElQAa9HT-RNBit32x6C_iBoBYg8QoMuSTcPIVxzNzddyFw?e=9Cbokm",
            "Brightdrop": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/EjP4u3WI9hlCi3k2sVx68RkBOVPGQHrQ5s-w_bDvLk7xjg?e=XMl4Pi",
            "Buick": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/EpTQ19HJYTJClKfBL8sWZPcB_1siI7_HBUbGDxljMAPffQ?e=5DzybE",
            "Cadillac": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/Eq_HecuhDyRLpq1CA1HQwUEBKAJBbApj_kq7Ysp46tCyaQ?e=PVL9Wd",
            "Chevrolet": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/Es1Ts3bWdeBBiFIkV6u5rDcBzgzHitt0LluqN2MMKibXJQ?e=p3Xrbk",
            "Chrysler": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/Etc0KL9giEBMmfR3eUMN12cBBZm7i80tvroBK-KJvl6NHw?e=Tfaclo",
            "Dodge": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/Er-Fxe4uQ1FFnTse7VRcVncBvEBaFIEVjp4gkmdDLpzTPQ?e=wv5HnX",
            "Fiat": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/Eiu-7WxDdZ9JmRRmFQJmPcgBoxQVHGFJP1MHWYuf1uAwBA?e=irsMGM",
            "Ford": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/Eh1eSmaBWJVKnp8qcBURjpIBOmTqFYV3Kanzk7iOSFG_iw?e=sJc4aP",
            "Genesis": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/EjGjMPYyzY9Pvq3_WO8XImMB-cpFCpWDREhOzlCPSNk4tQ?e=c3vTcT",
            "GMC": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/EkLkAwwtpLNEmoxyhwREImcBruvb6Os-DdudL9B0KDvnuA?e=T2adym",
            "Honda": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/EkZKi51f0mVDpoRqr2SMAn8BPHq8zQWo_9xxa2Cs9tCidQ?e=yk4eo6",
            "Hyundai": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/EjE-7ITJcD9NproTitinWTQBdOZwCjRKZMwoIreBahxqFA?e=cUvMSs",
            "Infiniti": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/EouRS__qu0xJhiOESGAyrVsBImSJPRzO2GSeODMn0jDSNQ?e=mPrxEp",
            "Jaguar": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/Eu5B7lvGjhtElkEPsSV4gXgBrmmlR_CN3cBtKCYZEy195Q?e=M88RhW",
            "Jeep": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/EnHsDhbHq8FApbsfGAv4EjwB8pQwoF2F9EhkefRXvLbvYg?e=Nk6Ufa",
            "Kia": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/EtWreIYyuapKppkYtgsxkRAB2GA91hQKHzuWrm7a_QTWKw?e=WQm9aT",
            "Land Rover": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/EoULD7SYsmxBm7cV9A2QzF4BUQ3KxM8Jd3w3jSWBlNB4LQ?e=2pxQuz",
            "Lexus": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/Ej0l_oKrgVxDlQsymIC5x4QBum-9f1hVWyk1Y0XAnuyNWA?e=7k0KKm",
            "Lincoln": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/Epdie7oZKvVNvM286yyr-fUBAYncAvMRRgLDtTfOEs4Deg?e=YlsN0S",
            "Mazda": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/Ekm67apegXNKqvUV_-1M20sBBJ7jrtmOd5lH6eIW9XZD2w?e=q9XcYh",
            "Mercedes": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/EgF2pHCD-GlNowOZ-HFbQ0UBedkihUPwg8ivNqM2fnOBfQ?e=nosRRS",
            "Mini": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/ElDvB_z5c81AoWUrkfuIfBoB6HPFb4VH_IW-PTrD9nbSvQ?e=hekEFp",
            "Mitsubishi": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/EqABJkN9KGVNnOsXi-LBjYsBepqjm0i4LzXjgYfcq5LCKA?e=wdwrQy",
            "Nissan": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/EueIzuzK4c9Hifs0p10X3LkB8dBY6pCQpN4BX8OdzIhlSQ?e=kCWehA",
            "Porsche": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/Eh9mN-bex5FBsqGG3nCdmhAByV08wO9wj8pgh6CRhbY70Q?e=KOvjZZ",
            "Ram": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/EsVy318lU5RGqqhxeREcz64BzfWsDJc_D_SjAQMjHofFTQ?e=9BAdFz",
            # "Rolls Royce": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/EuUx8qWb8a5Lg69dZiqgPrkBElO3gQAuaLTZOvKdlOIkJg?e=3bPP4d", #Not available in Sharepoint
            "Subaru": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/EtmpOsrJmJ1EvKPd8YfjtCsBYtFit7XQ2Y375ccoEijoEA?e=N5gN9f",
            "Tesla": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/EjtvRkoo9NZCuRjBLWFK42EBu9DpqnQH8X3_2A-fOcprAw?e=Yh0QVO",
            "Toyota": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/ElWISB73T5FEt9UM9VetQEgBVfNhbRbbvRSN3Csn7Sn_hQ?e=RDwlIA",
            "Volkswagen": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/En6g2PWkUWBKnTLlDiQKlVcB7-rqXOeNk2nEo6-c97IXWw?e=4msFBk",
            "Volvo": "https://calibercollision.sharepoint.com/:f:/s/O365-Protech-InformationSolutions/EvnDnyzP0KdNrCzRruhMgVMB_LSqa_12qpp4bxFVZHZTWQ?e=8MJnHp", 
            # Add Repair SI Sharepoint Links here
        }

        # how many times to try each manufacturer before giving up
        self.max_attempts = 5

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
        self.setWindowTitle('Hyperlink Automation')
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
        self.manufacturer_tree = QTreeWidget(self)
        self.manufacturer_tree.setHeaderHidden(True)
        self.manufacturer_tree.setFixedWidth(260)  # 👈 Shift closer by narrowing it
        self.manufacturer_tree.setStyleSheet("""
            QTreeWidget {
                background-color: #3e3e3e;
                color: white;
                border: 1px solid #555555;
                border-radius: 5px;
                margin-left: 10px;  /* 👈 Fine-tune left shift */
            }
        """)

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
        adas_label.setStyleSheet("font-size: 14px; padding: 5px;")
        adas_selection_layout.addWidget(adas_label)
    
        adas_acronyms = ["ACC", "AEB", "AHL", "APA", "BSW", "BUC", "LKA", "LW", "NV", "SVC"]
        self.adas_checkboxes = []
        repair_systems = [
            "SAS", "YAW", "G-Force", "SWS", "AHL", "NV", "HUD", "SRS", "SRA", 
            "ESC", "SRS D&E", "SCI", "SRR", "HLI", "TPMS", "SBI", "RC",
            "EBDE (1)", "EBDE (2)", "HDE (1)", "HDE (2)", "LGR", "PSI", "WRL",
            "PCM", "TRANS", "AIR", "ABS", "BCM","ODS","OCS","OCS2","OCS3","OCS4",
            "KEY", "FOB", "HVAC (1)", "HVAC (2)", "COOL", "HEAD (1)", "HEAD (2)",
        
            # human-readable names
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

        # Dark mode toggle
        theme_switch_section.addStretch()
        self.theme_toggle = ToggleSwitch(self)
        theme_switch_section.addWidget(self.theme_toggle)
        layout.addLayout(theme_switch_section)
    
        # ── Clean up Mode checkbox ──
        self.cleanup_checkbox = QCheckBox("Broken Hyperlink Mode", self)
        self.cleanup_checkbox.setStyleSheet("font-size: 14px; padding: 5px;")
        layout.addWidget(self.cleanup_checkbox)
        
        # … after all your other widgets but *before* the progress bars …
    
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
    
        layout.addWidget(self.current_manufacturer_label)
        layout.addWidget(self.current_manufacturer_progress)
        layout.addWidget(self.overall_progress_label)
        layout.addWidget(self.overall_progress_bar)
    
    # … then finish with setLayout, resize, etc.

        
        

        # after adding all widgets and layouts…
        self.si_mode_toggle.stateChanged.connect(self.on_si_mode_toggled)

        # set initial enabled/disabled state based on default toggle
        self.on_si_mode_toggled(self.mode_switch.checkState())

        self.setLayout(layout)
        self.resize(600, 400)

    def handle_extractor_output(self, line: str):
        # always print to your on-screen terminal
        self.terminal.append_output(line)
    
        m = re.search(r'(\d+)\s+Folders Remain', line)
        if m:
            remaining = int(m.group(1))
    
            # record the very first—or any larger—remaining count we see
            if not hasattr(self, '_initial_folder_count') or self._initial_folder_count is None:
                self._initial_folder_count = remaining
            else:
                self._initial_folder_count = max(self._initial_folder_count, remaining)
    
            initial = self._initial_folder_count
            # compute percent done
            percent = int((initial - remaining) / initial * 100)
            # clamp
            percent = max(0, min(100, percent))
    
            self.current_manufacturer_progress.setValue(percent)

    
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
        
        confirm_message = (
            "Excel files selected:\n"
            f"{excel_list}\n\n"
            "Manufacturers selected:\n"
            f"{manu_list}\n\n"
            "Systems selected:\n"
            + ", ".join(selected_systems)
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
        sharepoint_link = link_dict.get(manufacturer)
    
        if not sharepoint_link:
            QMessageBox.warning(
                self,
                'Error',
                f"No SharePoint link found for {manufacturer} in {self.mode_flag} mode.",
                QMessageBox.Ok
            )
            return
    
        script_path = os.path.join(os.path.dirname(__file__), "SharepointExtractor.py")
        args = [
            sys.executable,
            script_path,
            sharepoint_link,
            excel_path,
            ",".join(self.selected_systems),
            self.mode_flag,
            "cleanup" if self.cleanup_checkbox.isChecked() else "full"
        ]
    
        thread = WorkerThread(args, manufacturer, parent=self)
        self.thread = thread
    
        # ── Connect to our custom handler for both terminal + progress parsing ──
        thread.output_signal.connect(self.handle_extractor_output)
    
        # ── Overall‐style progress remains wired to the bar directly ──
        # thread.progress_signal.connect(self.current_manufacturer_progress.setValue)
    
        thread.finished_signal.connect(self.on_manufacturer_finished)
        thread.start()
        self.threads.append(thread)

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
        
    def on_start_stop(self):
        # — START path —
        if not self.is_running:
            self.start_automation()
    
            # ── enable & reset Pause/Resume button when starting ──
            self.pause_requested = False
            self.pause_button.setText('Pause Automation')
            self.pause_button.setEnabled(True)
    
            return
    
        # — STOP path —
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
    
        # clear pause flag
        self.pause_requested = False
    
        # tell loops not to launch any more work
        self.stop_requested = True
    
        # 1) Ask the Python extractor to shut down nicely
        if self.thread is not None and hasattr(self.thread, "process"):
            try:
                # On Windows this sends CTRL+BREAK to the whole group
                if os.name == "nt":
                    self.thread.process.send_signal(signal.CTRL_BREAK_EVENT)
                else:
                    # On Unix, send SIGTERM to the entire session
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
                        try:
                            child.kill()
                        except psutil.NoSuchProcess:
                            pass
    
                try:
                    parent.kill()
                except psutil.NoSuchProcess:
                    pass
    
            kill_children(self.thread.process.pid)
    
        # ── reset & disable Pause/Resume button when stopping ──
        self.pause_button.setText('Pause Automation')
        self.pause_button.setEnabled(False)
    
        # 3) Give it a moment, then report & swap button back
        sleep(1)
        self.terminal.append_output("❌ Hyperlink Automation has stopped.")
        # Example in on_start_stop STOP path:
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
        insert_index = layout.indexOf(self.current_manufacturer_label)  # before labels and bars
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
