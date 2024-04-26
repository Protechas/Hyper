import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QComboBox, QMessageBox, QFileDialog
from threading import Thread
from selenium import webdriver
from selenium.webdriver import ActionChains
from openpyxl.styles import Font, Color
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.common.by import By
from openpyxl import load_workbook
import pyperclip
import time
import win32com.client as win32
from time import sleep

def add_hyperlink_to_excel(file_path, sheet_name, cell_address, url, display_text):
    wb = load_workbook(file_path)
    ws = wb[sheet_name]
    ws[cell_address].hyperlink = url
    ws[cell_address].value = display_text
    ws[cell_address].font = Font(color="0000FF", underline='single')
    wb.save(file_path)

class SeleniumAutomationApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.excel_path = ''
        self.driver = None
        self.wait = None
        self.action_chains = None

    def init_selenium(self):
        # Set up Chrome options
        chrome_options = Options()
        chrome_options.add_experimental_option("detach", True)
        
        # Set up the Chrome WebDriver
        self.driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
        self.wait = WebDriverWait(self.driver, 10)
        self.action_chains = ActionChains(self.driver)

    def initUI(self):
        self.setWindowTitle('Selenium Automation')
        layout = QVBoxLayout()

        # Manufacturer dropdown
        self.manufacturer_dropdown = QComboBox(self)
        self.manufacturer_dropdown.addItems(["Acura", "Audi", "BMW", "Chevrolet"])  # Add all your manufacturers here
        layout.addWidget(self.manufacturer_dropdown)

        # Start button
        self.start_button = QPushButton('Run Automation', self)
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
        confirm_message = f"You have selected {manufacturer}. Are you sure? This can take some time as it will be going through everything and refreshing links, continue?"
        confirm = QMessageBox.question(self, 'Confirmation', confirm_message, QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

        if confirm == QMessageBox.Yes and self.excel_path:
            self.init_selenium()
            Thread(target=lambda: self.run_manufacturer_script(manufacturer)).start()
        elif not self.excel_path:
            QMessageBox.warning(self, 'Warning', "Please select an Excel file first.", QMessageBox.Ok)

    def run_manufacturer_script(self, manufacturer):
        try:
            if manufacturer == "Acura":
                self.run_acura_script()
            # Add additional checks for other manufacturers
            # elif manufacturer == "BMW":
            #     self.run_bmw_script()
            # ... and so on for each manufacturer ...
        except Exception as e:
            QMessageBox.critical(self, 'Error', str(e), QMessageBox.Ok)
        finally:                  
            self.driver.quit()

    def run_acura_script(self):
        try:    
            # Collect URLs
            urls = []
            
            # Open a specific SharePoint page for Acura
            self.driver.get('https://calibercollision-my.sharepoint.com/:f:/g/personal/mark_klingenhofer_protechdfw_com/EjIo8sg9qXNEt6CDCKpeRGkBEY-67TppBRysHPrqdbNSmg')
        
            # Clicks Acura
            acura = self.wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="appRoot"]/div/div[2]/div/div/div[2]/div[2]/main/div/div/div[2]/div/div/div/div/div[2]/div/div/div/div[2]/div/div/div[3]/div/div[1]/span/span[1]/button')))
            acura.click()
        
            # Clicks 2012 Year
            acurayear2012 = self.wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="appRoot"]/div/div[2]/div/div/div[2]/div[2]/main/div/div/div[2]/div/div/div/div/div[2]/div/div/div/div[2]')))
            self.action_chains.double_click(acurayear2012).perform()
        
            # Clicks MDX Model
            acura2012mdx = self.wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="appRoot"]/div/div[2]/div/div/div[2]/div[2]/main/div/div/div[2]/div/div/div/div/div[2]/div/div/div/div[1]')))
            self.action_chains.double_click(acura2012mdx).perform()
            
            # Clicks on ACC Doc
            acura2012mdxacc = self.wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="appRoot"]/div/div[2]/div/div/div[2]/div[2]/main/div/div/div[2]/div/div/div/div/div[2]/div/div/div/div[1]')))
            self.action_chains.double_click(acura2012mdxacc).perform()

            time.sleep(3)
            
            document_url = self.driver.current_url  # This line is conceptual

            # Add hyperlink to Excel
            file_path = self.excel_path
            sheet_name = 'Sheet1'
            cell_address = 'L2'
            display_text = document_url  # or some other meaningful text
            add_hyperlink_to_excel(file_path, sheet_name, cell_address, document_url, display_text)

            # Continue with any other necessary steps, such as closing the browser
        except Exception as e:
            print(f"An error occurred: {e}")
            QMessageBox.critical(self, 'Error', str(e), QMessageBox.Ok)
        finally:
            # Decide whether to quit the driver based on your needs
            pass

    def safe_action(self, action, description, max_attempts=3):
        attempts = 0
        while attempts < max_attempts:
            try:
                action()
                return
            except (NoSuchElementException, TimeoutException) as e:
                print(f"Error performing {description}: {e}")
                time.sleep(2)
                attempts += 1
        raise Exception(f"Failed to perform {description} after {max_attempts} attempts")

# ... (rest of the PyQt5 application setup)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = SeleniumAutomationApp()
    ex.show()
    sys.exit(app.exec_())
