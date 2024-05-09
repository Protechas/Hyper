import sys
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font
import time
import io
from PIL import Image
import pytesseract
import re

# Set the path for Tesseract if not in PATH
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'  # Update this path as needed

def screenshot_and_get_text(driver):
    # Take screenshot using Selenium
    screenshot = driver.get_screenshot_as_png()
    image = Image.open(io.BytesIO(screenshot))
    # Use Tesseract to extract text
    text = pytesseract.image_to_string(image)
    return text

def find_row_in_excel(ws, year, make, model, adas_system):

    for row in ws.iter_rows(min_row=2, max_col=8):
        year_cell = row[0]  # Assuming year is in the first column
        make_cell = row[1]  # Assuming make is in the second column
        model_cell = row[2]  # Assuming model is in the third column
        adas_cell = row[7]  # Assuming ADAS system is in the eighth column

        # Debugging output to check what is being compared
        print(f"Checking row: Year={year_cell.value}, Make={make_cell.value}, Model={model_cell.value}, ADAS={adas_cell.value}")

        # Match year, make, model, and ADAS system; ensure to convert everything to string and strip any whitespace
        if str(year_cell.value).strip() == str(year).strip() and \
           str(make_cell.value).strip().lower() == make.lower().strip() and \
           str(model_cell.value).strip().lower() == model.lower().strip() and \
           adas_system.lower().strip() in str(adas_cell.value).lower().strip():
            # Return the cell in column L of the matched row
            return ws.cell(row=year_cell.row, column=12)  # Assumes column L is 12

    return None

def double_click_element(driver, wait, xpath):
    element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
    ActionChains(driver).double_click(element).perform()

def add_hyperlink_to_excel(file_path, sheet_name, cell_address, url, display_text):
    wb = load_workbook(file_path)
    ws = wb[sheet_name]
    ws[cell_address].hyperlink = url
    ws[cell_address].value = url
    ws[cell_address].font = Font(color="0000FF", underline='single')
    wb.save(file_path)
    
def get_document_url(driver, wait, document_xpath):
        double_click_element(driver, wait, document_xpath)  
        time.sleep(3)  
        document_url = driver.current_url
        time.sleep(3)  
        driver.back
        return document_url

def navigate_to_model(driver, wait, model_xpath):
     model_link = wait.until(EC.element_to_be_clickable((By.XPATH, model_xpath)))
     model_link.click()
     time.sleep(2)  # Wait for the model's page to load

def navigate_to_year(driver, wait, year_xpath):
     year_link = wait.until(EC.element_to_be_clickable((By.XPATH, year_xpath)))
     year_link.click()

def run_acura_script(excel_path):
    # Setup WebDriver
    options = webdriver.ChromeOptions()
    options.add_experimental_option("detach", True)
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    wait = WebDriverWait(driver, 10)
    action_chains = ActionChains(driver)
    
    # Your structured data
    years_models_documents = {
                           ###################
                           #                 #
                           #      2012       #
                           #                 #
                           ###################
        '2012': {
        'year_page_xpath': '//*[@data-automationid="ListCell"][2]',  
            'models': {
                'MDX': {
                    'model_page_xpath': '//*[@data-automationid="ListCell"][1]',  
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][1]', 
                        'AEB': '//*[@data-automationid="ListCell"][2]',  
                        'AHL': '//*[@data-automationid="ListCell"][6]',
                        'APA': '//*[@data-automationid="ListCell"][5]',
                        'BSW': '//*[@data-automationid="ListCell"][3]',
                        'BUC': '//*[@data-automationid="ListCell"][4]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]',                        
                    }
                },
                'RDX': {   #copy this Line v
                    'model_page_xpath': '//*[@data-automationid="ListCell"][2]',  
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][3]', 
                        'AEB': '//*[@data-automationid="ListCell"][5]',  
                        'AHL': '//*[@data-automationid="ListCell"][4]',
                        'APA': '//*[@data-automationid="ListCell"][2]',
                        'BSW': '//*[@data-automationid="ListCell"][6]',
                        'BUC': '//*[@data-automationid="ListCell"][1]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]',  
                    }
                },        #to this time ^
                'RL': {
                    'model_page_xpath': '//*[@data-automationid="ListCell"][3]', 
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][1]', 
                        'AEB': '//*[@data-automationid="ListCell"][2]',  
                        'AHL': '//*[@data-automationid="ListCell"][5]',
                        'APA': '//*[@data-automationid="ListCell"][3]',
                        'BSW': '//*[@data-automationid="ListCell"][6]',
                        'BUC': '//*[@data-automationid="ListCell"][4]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]', 
                    }
                },
                'TL': {
                    'model_page_xpath': '//*[@data-automationid="ListCell"][4]', 
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][4]', 
                        'AEB': '//*[@data-automationid="ListCell"][6]',  
                        'AHL': '//*[@data-automationid="ListCell"][5]',
                        'APA': '//*[@data-automationid="ListCell"][3]',
                        'BSW': '//*[@data-automationid="ListCell"][1]',
                        'BUC': '//*[@data-automationid="ListCell"][2]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]',  
                    }
                },
                'TSX': {
                    'model_page_xpath': '//*[@data-automationid="ListCell"][5]',
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][1]', 
                        'AEB': '//*[@data-automationid="ListCell"][5]',  
                        'AHL': '//*[@data-automationid="ListCell"][4]',
                        'APA': '//*[@data-automationid="ListCell"][2]',
                        'BSW': '//*[@data-automationid="ListCell"][6]',
                        'BUC': '//*[@data-automationid="ListCell"][1]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]',  
                    }
                },
                'ZDX': { 
                    'model_page_xpath': '//*[@data-automationid="ListCell"][6]', ########
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][1]', 
                        'AEB': '//*[@data-automationid="ListCell"][2]',  
                        'AHL': '//*[@data-automationid="ListCell"][6]',
                        'APA': '//*[@data-automationid="ListCell"][5]',
                        'BSW': '//*[@data-automationid="ListCell"][3]',
                        'BUC': '//*[@data-automationid="ListCell"][4]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]',  
                    }
                },                
            }
        },
        '2013': {
                           ###################
                           #                 #
                           #      2013       #
                           #                 #
                           ###################
        'year_page_xpath': '//*[@data-automationid="ListCell"][3]',  
            'models': {
                'ILX': {
                    'model_page_xpath': '//*[@data-automationid="ListCell"][1]',  
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][3]', 
                        'AEB': '//*[@data-automationid="ListCell"][5]',  
                        'AHL': '//*[@data-automationid="ListCell"][4]',
                        'APA': '//*[@data-automationid="ListCell"][2]',
                        'BSW': '//*[@data-automationid="ListCell"][6]',
                        'BUC': '//*[@data-automationid="ListCell"][1]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][9]',
                        'SVC': '//*[@data-automationid="ListCell"][11]',                        
                    }
                },
                'MDX': {   #copy this Line v
                    'model_page_xpath': '//*[@data-automationid="ListCell"][2]',  
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][1]', 
                        'AEB': '//*[@data-automationid="ListCell"][2]',  
                        'AHL': '//*[@data-automationid="ListCell"][6]',
                        'APA': '//*[@data-automationid="ListCell"][5]',
                        'BSW': '//*[@data-automationid="ListCell"][3]',
                        'BUC': '//*[@data-automationid="ListCell"][4]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]',  
                    }
                },        #to this time ^
                'RDX': {
                    'model_page_xpath': '//*[@data-automationid="ListCell"][3]', 
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][3]', 
                        'AEB': '//*[@data-automationid="ListCell"][5]',  
                        'AHL': '//*[@data-automationid="ListCell"][4]',
                        'APA': '//*[@data-automationid="ListCell"][2]',
                        'BSW': '//*[@data-automationid="ListCell"][6]',
                        'BUC': '//*[@data-automationid="ListCell"][1]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        #'NV': '//*[@data-automationid="ListCell"][%]',   <------ Needs to be added
                        'SVC': '//*[@data-automationid="ListCell"][10]', 
                    }
                },
                'TL': {
                    'model_page_xpath': '//*[@data-automationid="ListCell"][4]', 
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][4]', 
                        'AEB': '//*[@data-automationid="ListCell"][6]',  
                        'AHL': '//*[@data-automationid="ListCell"][5]',
                        'APA': '//*[@data-automationid="ListCell"][3]',
                        'BSW': '//*[@data-automationid="ListCell"][1]',
                        'BUC': '//*[@data-automationid="ListCell"][2]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        #'NV': '//*[@data-automationid="ListCell"][%]',   <------ Needs to be added
                        'SVC': '//*[@data-automationid="ListCell"][9]',  
                    }
                },
                'TSX': {
                    'model_page_xpath': '//*[@data-automationid="ListCell"][5]',
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][1]', 
                        'AEB': '//*[@data-automationid="ListCell"][5]',  
                        'AHL': '//*[@data-automationid="ListCell"][4]',
                        'APA': '//*[@data-automationid="ListCell"][2]',
                        'BSW': '//*[@data-automationid="ListCell"][6]',
                        'BUC': '//*[@data-automationid="ListCell"][1]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]',  
                    }
                },
                'ZDX': { 
                    'model_page_xpath': '//*[@data-automationid="ListCell"][6]', 
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][1]', 
                        'AEB': '//*[@data-automationid="ListCell"][2]',  
                        'AHL': '//*[@data-automationid="ListCell"][2]',
                        'APA': '//*[@data-automationid="ListCell"][2]',
                        'BSW': '//*[@data-automationid="ListCell"][3]',
                        'BUC': '//*[@data-automationid="ListCell"][2]',
                        'LKA': '//*[@data-automationid="ListCell"][2]',
                        'NV': '//*[@data-automationid="ListCell"][2]',
                        'SVC': '//*[@data-automationid="ListCell"][2]',  
                    }
                },                
            }
        },
        '2014': {
                           ###################
                           #                 #
                           #      2014       #
                           #                 #
                           ###################
        'year_page_xpath': '//*[@data-automationid="ListCell"][2]',  
            'models': {
                'MDX': {
                    'model_page_xpath': '//*[@data-automationid="ListCell"][1]',  
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][1]', 
                        'AEB': '//*[@data-automationid="ListCell"][2]',  
                        'AHL': '//*[@data-automationid="ListCell"][6]',
                        'APA': '//*[@data-automationid="ListCell"][5]',
                        'BSW': '//*[@data-automationid="ListCell"][3]',
                        'BUC': '//*[@data-automationid="ListCell"][4]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]',                        
                    }
                },
                'RDX': {   #copy this Line v
                    'model_page_xpath': '//*[@data-automationid="ListCell"][2]',  
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][3]', 
                        'AEB': '//*[@data-automationid="ListCell"][5]',  
                        'AHL': '//*[@data-automationid="ListCell"][4]',
                        'APA': '//*[@data-automationid="ListCell"][2]',
                        'BSW': '//*[@data-automationid="ListCell"][6]',
                        'BUC': '//*[@data-automationid="ListCell"][1]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]',  
                    }
                },        #to this time ^
                'RL': {
                    'model_page_xpath': '//*[@data-automationid="ListCell"][3]', 
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][1]', 
                        'AEB': '//*[@data-automationid="ListCell"][2]',  
                        'AHL': '//*[@data-automationid="ListCell"][5]',
                        'APA': '//*[@data-automationid="ListCell"][3]',
                        'BSW': '//*[@data-automationid="ListCell"][6]',
                        'BUC': '//*[@data-automationid="ListCell"][4]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]', 
                    }
                },
                'TL': {
                    'model_page_xpath': '//*[@data-automationid="ListCell"][4]', 
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][4]', 
                        'AEB': '//*[@data-automationid="ListCell"][6]',  
                        'AHL': '//*[@data-automationid="ListCell"][5]',
                        'APA': '//*[@data-automationid="ListCell"][3]',
                        'BSW': '//*[@data-automationid="ListCell"][1]',
                        'BUC': '//*[@data-automationid="ListCell"][2]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]',  
                    }
                },
                'TSX': {
                    'model_page_xpath': '//*[@data-automationid="ListCell"][5]',
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][1]', 
                        'AEB': '//*[@data-automationid="ListCell"][5]',  
                        'AHL': '//*[@data-automationid="ListCell"][4]',
                        'APA': '//*[@data-automationid="ListCell"][2]',
                        'BSW': '//*[@data-automationid="ListCell"][6]',
                        'BUC': '//*[@data-automationid="ListCell"][1]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]',  
                    }
                },
                'ZDX': { 
                    'model_page_xpath': '//*[@data-automationid="ListCell"][6]', 
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][1]', 
                        'AEB': '//*[@data-automationid="ListCell"][2]',  
                        'AHL': '//*[@data-automationid="ListCell"][2]',
                        'APA': '//*[@data-automationid="ListCell"][2]',
                        'BSW': '//*[@data-automationid="ListCell"][3]',
                        'BUC': '//*[@data-automationid="ListCell"][2]',
                        'LKA': '//*[@data-automationid="ListCell"][2]',
                        'NV': '//*[@data-automationid="ListCell"][2]',
                        'SVC': '//*[@data-automationid="ListCell"][2]',  
                    }
                },                
            }
        },
        '2015': {
                           ###################
                           #                 #
                           #      2015       #
                           #                 #
                           ###################  
                                     
        'year_page_xpath': '//*[@data-automationid="ListCell"][2]',  
            'models': {
                'MDX': {
                    'model_page_xpath': '//*[@data-automationid="ListCell"][1]',  
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][1]', 
                        'AEB': '//*[@data-automationid="ListCell"][2]',  
                        'AHL': '//*[@data-automationid="ListCell"][6]',
                        'APA': '//*[@data-automationid="ListCell"][5]',
                        'BSW': '//*[@data-automationid="ListCell"][3]',
                        'BUC': '//*[@data-automationid="ListCell"][4]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]',                        
                    }
                },
                'RDX': {   #copy this Line v
                    'model_page_xpath': '//*[@data-automationid="ListCell"][2]',  
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][3]', 
                        'AEB': '//*[@data-automationid="ListCell"][5]',  
                        'AHL': '//*[@data-automationid="ListCell"][4]',
                        'APA': '//*[@data-automationid="ListCell"][2]',
                        'BSW': '//*[@data-automationid="ListCell"][6]',
                        'BUC': '//*[@data-automationid="ListCell"][1]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]',  
                    }
                },        #to this time ^
                'RL': {
                    'model_page_xpath': '//*[@data-automationid="ListCell"][3]', 
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][1]', 
                        'AEB': '//*[@data-automationid="ListCell"][2]',  
                        'AHL': '//*[@data-automationid="ListCell"][5]',
                        'APA': '//*[@data-automationid="ListCell"][3]',
                        'BSW': '//*[@data-automationid="ListCell"][6]',
                        'BUC': '//*[@data-automationid="ListCell"][4]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]', 
                    }
                },
                'TL': {
                    'model_page_xpath': '//*[@data-automationid="ListCell"][4]', 
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][4]', 
                        'AEB': '//*[@data-automationid="ListCell"][6]',  
                        'AHL': '//*[@data-automationid="ListCell"][5]',
                        'APA': '//*[@data-automationid="ListCell"][3]',
                        'BSW': '//*[@data-automationid="ListCell"][1]',
                        'BUC': '//*[@data-automationid="ListCell"][2]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]',  
                    }
                },
                'TSX': {
                    'model_page_xpath': '//*[@data-automationid="ListCell"][5]',
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][1]', 
                        'AEB': '//*[@data-automationid="ListCell"][5]',  
                        'AHL': '//*[@data-automationid="ListCell"][4]',
                        'APA': '//*[@data-automationid="ListCell"][2]',
                        'BSW': '//*[@data-automationid="ListCell"][6]',
                        'BUC': '//*[@data-automationid="ListCell"][1]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]',  
                    }
                },
                'ZDX': { 
                    'model_page_xpath': '//*[@data-automationid="ListCell"][6]',
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][1]', 
                        'AEB': '//*[@data-automationid="ListCell"][2]',  
                        'AHL': '//*[@data-automationid="ListCell"][2]',
                        'APA': '//*[@data-automationid="ListCell"][2]',
                        'BSW': '//*[@data-automationid="ListCell"][3]',
                        'BUC': '//*[@data-automationid="ListCell"][2]',
                        'LKA': '//*[@data-automationid="ListCell"][2]',
                        'NV': '//*[@data-automationid="ListCell"][2]',
                        'SVC': '//*[@data-automationid="ListCell"][2]',  
                    }
                },                
            }
        },
        '2016': {
                           ###################
                           #                 #
                           #      2016       #
                           #                 #
                           ###################
            
        'year_page_xpath': '//*[@data-automationid="ListCell"][2]',  
            'models': {
                'MDX': {
                    'model_page_xpath': '//*[@data-automationid="ListCell"][1]',  
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][1]', 
                        'AEB': '//*[@data-automationid="ListCell"][2]',  
                        'AHL': '//*[@data-automationid="ListCell"][6]',
                        'APA': '//*[@data-automationid="ListCell"][5]',
                        'BSW': '//*[@data-automationid="ListCell"][3]',
                        'BUC': '//*[@data-automationid="ListCell"][4]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]',                        
                    }
                },
                'RDX': {   #copy this Line v
                    'model_page_xpath': '//*[@data-automationid="ListCell"][2]',  
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][3]', 
                        'AEB': '//*[@data-automationid="ListCell"][5]',  
                        'AHL': '//*[@data-automationid="ListCell"][4]',
                        'APA': '//*[@data-automationid="ListCell"][2]',
                        'BSW': '//*[@data-automationid="ListCell"][6]',
                        'BUC': '//*[@data-automationid="ListCell"][1]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]',  
                    }
                },        #to this time ^
                'RL': {
                    'model_page_xpath': '//*[@data-automationid="ListCell"][3]', 
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][1]', 
                        'AEB': '//*[@data-automationid="ListCell"][2]',  
                        'AHL': '//*[@data-automationid="ListCell"][5]',
                        'APA': '//*[@data-automationid="ListCell"][3]',
                        'BSW': '//*[@data-automationid="ListCell"][6]',
                        'BUC': '//*[@data-automationid="ListCell"][4]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]', 
                    }
                },
                'TL': {
                    'model_page_xpath': '//*[@data-automationid="ListCell"][4]', 
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][4]', 
                        'AEB': '//*[@data-automationid="ListCell"][6]',  
                        'AHL': '//*[@data-automationid="ListCell"][5]',
                        'APA': '//*[@data-automationid="ListCell"][3]',
                        'BSW': '//*[@data-automationid="ListCell"][1]',
                        'BUC': '//*[@data-automationid="ListCell"][2]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]',  
                    }
                },
                'TSX': {
                    'model_page_xpath': '//*[@data-automationid="ListCell"][5]',
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][1]', 
                        'AEB': '//*[@data-automationid="ListCell"][5]',  
                        'AHL': '//*[@data-automationid="ListCell"][4]',
                        'APA': '//*[@data-automationid="ListCell"][2]',
                        'BSW': '//*[@data-automationid="ListCell"][6]',
                        'BUC': '//*[@data-automationid="ListCell"][1]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]',  
                    }
                },
                'ZDX': { 
                    'model_page_xpath': '//*[@data-automationid="ListCell"][6]',
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][1]', 
                        'AEB': '//*[@data-automationid="ListCell"][2]',  
                        'AHL': '//*[@data-automationid="ListCell"][2]',
                        'APA': '//*[@data-automationid="ListCell"][2]',
                        'BSW': '//*[@data-automationid="ListCell"][3]',
                        'BUC': '//*[@data-automationid="ListCell"][2]',
                        'LKA': '//*[@data-automationid="ListCell"][2]',
                        'NV': '//*[@data-automationid="ListCell"][2]',
                        'SVC': '//*[@data-automationid="ListCell"][2]',  
                    }
                },                
            }
        },
        '2017': {
                           ###################
                           #                 #
                           #      2017       #
                           #                 #
                           ###################
        'year_page_xpath': '//*[@data-automationid="ListCell"][2]',  
            'models': {
                'MDX': {
                    'model_page_xpath': '//*[@data-automationid="ListCell"][1]',  
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][1]', 
                        'AEB': '//*[@data-automationid="ListCell"][2]',  
                        'AHL': '//*[@data-automationid="ListCell"][6]',
                        'APA': '//*[@data-automationid="ListCell"][5]',
                        'BSW': '//*[@data-automationid="ListCell"][3]',
                        'BUC': '//*[@data-automationid="ListCell"][4]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]',                        
                    }
                },
                'RDX': {   #copy this Line v
                    'model_page_xpath': '//*[@data-automationid="ListCell"][2]',  
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][3]', 
                        'AEB': '//*[@data-automationid="ListCell"][5]',  
                        'AHL': '//*[@data-automationid="ListCell"][4]',
                        'APA': '//*[@data-automationid="ListCell"][2]',
                        'BSW': '//*[@data-automationid="ListCell"][6]',
                        'BUC': '//*[@data-automationid="ListCell"][1]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]',  
                    }
                },        #to this time ^
                'RL': {
                    'model_page_xpath': '//*[@data-automationid="ListCell"][3]', 
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][1]', 
                        'AEB': '//*[@data-automationid="ListCell"][2]',  
                        'AHL': '//*[@data-automationid="ListCell"][5]',
                        'APA': '//*[@data-automationid="ListCell"][3]',
                        'BSW': '//*[@data-automationid="ListCell"][6]',
                        'BUC': '//*[@data-automationid="ListCell"][4]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]', 
                    }
                },
                'TL': {
                    'model_page_xpath': '//*[@data-automationid="ListCell"][4]', 
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][4]', 
                        'AEB': '//*[@data-automationid="ListCell"][6]',  
                        'AHL': '//*[@data-automationid="ListCell"][5]',
                        'APA': '//*[@data-automationid="ListCell"][3]',
                        'BSW': '//*[@data-automationid="ListCell"][1]',
                        'BUC': '//*[@data-automationid="ListCell"][2]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]',  
                    }
                },
                'TSX': {
                    'model_page_xpath': '//*[@data-automationid="ListCell"][5]',
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][1]', 
                        'AEB': '//*[@data-automationid="ListCell"][5]',  
                        'AHL': '//*[@data-automationid="ListCell"][4]',
                        'APA': '//*[@data-automationid="ListCell"][2]',
                        'BSW': '//*[@data-automationid="ListCell"][6]',
                        'BUC': '//*[@data-automationid="ListCell"][1]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]',  
                    }
                },
                'ZDX': { 
                    'model_page_xpath': '//*[@data-automationid="ListCell"][6]',
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][1]', 
                        'AEB': '//*[@data-automationid="ListCell"][2]',  
                        'AHL': '//*[@data-automationid="ListCell"][2]',
                        'APA': '//*[@data-automationid="ListCell"][2]',
                        'BSW': '//*[@data-automationid="ListCell"][3]',
                        'BUC': '//*[@data-automationid="ListCell"][2]',
                        'LKA': '//*[@data-automationid="ListCell"][2]',
                        'NV': '//*[@data-automationid="ListCell"][2]',
                        'SVC': '//*[@data-automationid="ListCell"][2]',  
                    }
                },                
            }
        },
        '2018': {
                           ###################
                           #                 #
                           #      2018       #
                           #                 #
                           ###################
            
        'year_page_xpath': '//*[@data-automationid="ListCell"][2]',  
            'models': {
                'MDX': {
                    'model_page_xpath': '//*[@data-automationid="ListCell"][1]',  
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][1]', 
                        'AEB': '//*[@data-automationid="ListCell"][2]',  
                        'AHL': '//*[@data-automationid="ListCell"][6]',
                        'APA': '//*[@data-automationid="ListCell"][5]',
                        'BSW': '//*[@data-automationid="ListCell"][3]',
                        'BUC': '//*[@data-automationid="ListCell"][4]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]',                        
                    }
                },
                'RDX': {   #copy this Line v
                    'model_page_xpath': '//*[@data-automationid="ListCell"][2]',  
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][3]', 
                        'AEB': '//*[@data-automationid="ListCell"][5]',  
                        'AHL': '//*[@data-automationid="ListCell"][4]',
                        'APA': '//*[@data-automationid="ListCell"][2]',
                        'BSW': '//*[@data-automationid="ListCell"][6]',
                        'BUC': '//*[@data-automationid="ListCell"][1]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]',  
                    }
                },        #to this time ^
                'RL': {
                    'model_page_xpath': '//*[@data-automationid="ListCell"][3]', 
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][1]', 
                        'AEB': '//*[@data-automationid="ListCell"][2]',  
                        'AHL': '//*[@data-automationid="ListCell"][5]',
                        'APA': '//*[@data-automationid="ListCell"][3]',
                        'BSW': '//*[@data-automationid="ListCell"][6]',
                        'BUC': '//*[@data-automationid="ListCell"][4]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]', 
                    }
                },
                'TL': {
                    'model_page_xpath': '//*[@data-automationid="ListCell"][4]', 
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][4]', 
                        'AEB': '//*[@data-automationid="ListCell"][6]',  
                        'AHL': '//*[@data-automationid="ListCell"][5]',
                        'APA': '//*[@data-automationid="ListCell"][3]',
                        'BSW': '//*[@data-automationid="ListCell"][1]',
                        'BUC': '//*[@data-automationid="ListCell"][2]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]',  
                    }
                },
                'TSX': {
                    'model_page_xpath': '//*[@data-automationid="ListCell"][5]',
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][1]', 
                        'AEB': '//*[@data-automationid="ListCell"][5]',  
                        'AHL': '//*[@data-automationid="ListCell"][4]',
                        'APA': '//*[@data-automationid="ListCell"][2]',
                        'BSW': '//*[@data-automationid="ListCell"][6]',
                        'BUC': '//*[@data-automationid="ListCell"][1]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]',  
                    }
                },
                'ZDX': { 
                    'model_page_xpath': '//*[@data-automationid="ListCell"][6]',
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][1]', 
                        'AEB': '//*[@data-automationid="ListCell"][2]',  
                        'AHL': '//*[@data-automationid="ListCell"][2]',
                        'APA': '//*[@data-automationid="ListCell"][2]',
                        'BSW': '//*[@data-automationid="ListCell"][3]',
                        'BUC': '//*[@data-automationid="ListCell"][2]',
                        'LKA': '//*[@data-automationid="ListCell"][2]',
                        'NV': '//*[@data-automationid="ListCell"][2]',
                        'SVC': '//*[@data-automationid="ListCell"][2]',  
                    }
                },                
            }
        },
        '2019': {
                           ###################
                           #                 #
                           #      2019       #
                           #                 #
                           ###################
        'year_page_xpath': '//*[@data-automationid="ListCell"][2]',  
            'models': {
                'MDX': {
                    'model_page_xpath': '//*[@data-automationid="ListCell"][1]',  
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][1]', 
                        'AEB': '//*[@data-automationid="ListCell"][2]',  
                        'AHL': '//*[@data-automationid="ListCell"][6]',
                        'APA': '//*[@data-automationid="ListCell"][5]',
                        'BSW': '//*[@data-automationid="ListCell"][3]',
                        'BUC': '//*[@data-automationid="ListCell"][4]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]',                        
                    }
                },
                'RDX': {   #copy this Line v
                    'model_page_xpath': '//*[@data-automationid="ListCell"][2]',  
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][3]', 
                        'AEB': '//*[@data-automationid="ListCell"][5]',  
                        'AHL': '//*[@data-automationid="ListCell"][4]',
                        'APA': '//*[@data-automationid="ListCell"][2]',
                        'BSW': '//*[@data-automationid="ListCell"][6]',
                        'BUC': '//*[@data-automationid="ListCell"][1]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]',  
                    }
                },        #to this time ^
                'RL': {
                    'model_page_xpath': '//*[@data-automationid="ListCell"][3]', 
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][1]', 
                        'AEB': '//*[@data-automationid="ListCell"][2]',  
                        'AHL': '//*[@data-automationid="ListCell"][5]',
                        'APA': '//*[@data-automationid="ListCell"][3]',
                        'BSW': '//*[@data-automationid="ListCell"][6]',
                        'BUC': '//*[@data-automationid="ListCell"][4]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]', 
                    }
                },
                'TL': {
                    'model_page_xpath': '//*[@data-automationid="ListCell"][4]', 
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][4]', 
                        'AEB': '//*[@data-automationid="ListCell"][6]',  
                        'AHL': '//*[@data-automationid="ListCell"][5]',
                        'APA': '//*[@data-automationid="ListCell"][3]',
                        'BSW': '//*[@data-automationid="ListCell"][1]',
                        'BUC': '//*[@data-automationid="ListCell"][2]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]',  
                    }
                },
                'TSX': {
                    'model_page_xpath': '//*[@data-automationid="ListCell"][5]',
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][1]', 
                        'AEB': '//*[@data-automationid="ListCell"][5]',  
                        'AHL': '//*[@data-automationid="ListCell"][4]',
                        'APA': '//*[@data-automationid="ListCell"][2]',
                        'BSW': '//*[@data-automationid="ListCell"][6]',
                        'BUC': '//*[@data-automationid="ListCell"][1]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]',  
                    }
                },
                'ZDX': { 
                    'model_page_xpath': '//*[@data-automationid="ListCell"][6]',
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][1]', 
                        'AEB': '//*[@data-automationid="ListCell"][2]',  
                        'AHL': '//*[@data-automationid="ListCell"][2]',
                        'APA': '//*[@data-automationid="ListCell"][2]',
                        'BSW': '//*[@data-automationid="ListCell"][3]',
                        'BUC': '//*[@data-automationid="ListCell"][2]',
                        'LKA': '//*[@data-automationid="ListCell"][2]',
                        'NV': '//*[@data-automationid="ListCell"][2]',
                        'SVC': '//*[@data-automationid="ListCell"][2]',  
                    }
                },                
            }
        },
        '2020': {
                           ###################
                           #                 #
                           #      2020       #
                           #                 #
                           ###################
        'year_page_xpath': '//*[@data-automationid="ListCell"][2]',  
            'models': {
                'MDX': {
                    'model_page_xpath': '//*[@data-automationid="ListCell"][1]',  
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][1]', 
                        'AEB': '//*[@data-automationid="ListCell"][2]',  
                        'AHL': '//*[@data-automationid="ListCell"][6]',
                        'APA': '//*[@data-automationid="ListCell"][5]',
                        'BSW': '//*[@data-automationid="ListCell"][3]',
                        'BUC': '//*[@data-automationid="ListCell"][4]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]',                        
                    }
                },
                'RDX': {   #copy this Line v
                    'model_page_xpath': '//*[@data-automationid="ListCell"][2]',  
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][3]', 
                        'AEB': '//*[@data-automationid="ListCell"][5]',  
                        'AHL': '//*[@data-automationid="ListCell"][4]',
                        'APA': '//*[@data-automationid="ListCell"][2]',
                        'BSW': '//*[@data-automationid="ListCell"][6]',
                        'BUC': '//*[@data-automationid="ListCell"][1]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]',  
                    }
                },        #to this time ^
                'RL': {
                    'model_page_xpath': '//*[@data-automationid="ListCell"][3]', 
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][1]', 
                        'AEB': '//*[@data-automationid="ListCell"][2]',  
                        'AHL': '//*[@data-automationid="ListCell"][5]',
                        'APA': '//*[@data-automationid="ListCell"][3]',
                        'BSW': '//*[@data-automationid="ListCell"][6]',
                        'BUC': '//*[@data-automationid="ListCell"][4]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]', 
                    }
                },
                'TL': {
                    'model_page_xpath': '//*[@data-automationid="ListCell"][4]', 
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][4]', 
                        'AEB': '//*[@data-automationid="ListCell"][6]',  
                        'AHL': '//*[@data-automationid="ListCell"][5]',
                        'APA': '//*[@data-automationid="ListCell"][3]',
                        'BSW': '//*[@data-automationid="ListCell"][1]',
                        'BUC': '//*[@data-automationid="ListCell"][2]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]',  
                    }
                },
                'TSX': {
                    'model_page_xpath': '//*[@data-automationid="ListCell"][5]',
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][1]', 
                        'AEB': '//*[@data-automationid="ListCell"][5]',  
                        'AHL': '//*[@data-automationid="ListCell"][4]',
                        'APA': '//*[@data-automationid="ListCell"][2]',
                        'BSW': '//*[@data-automationid="ListCell"][6]',
                        'BUC': '//*[@data-automationid="ListCell"][1]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]',  
                    }
                },
                'ZDX': { 
                    'model_page_xpath': '//*[@data-automationid="ListCell"][6]',
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][1]', 
                        'AEB': '//*[@data-automationid="ListCell"][2]',  
                        'AHL': '//*[@data-automationid="ListCell"][2]',
                        'APA': '//*[@data-automationid="ListCell"][2]',
                        'BSW': '//*[@data-automationid="ListCell"][3]',
                        'BUC': '//*[@data-automationid="ListCell"][2]',
                        'LKA': '//*[@data-automationid="ListCell"][2]',
                        'NV': '//*[@data-automationid="ListCell"][2]',
                        'SVC': '//*[@data-automationid="ListCell"][2]',  
                    }
                },                
            }
        },
        '2021': {
                           ###################
                           #                 #
                           #      2021       #
                           #                 #
                           ###################
        'year_page_xpath': '//*[@data-automationid="ListCell"][2]',  
            'models': {
                'MDX': {
                    'model_page_xpath': '//*[@data-automationid="ListCell"][1]',  
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][1]', 
                        'AEB': '//*[@data-automationid="ListCell"][2]',  
                        'AHL': '//*[@data-automationid="ListCell"][6]',
                        'APA': '//*[@data-automationid="ListCell"][5]',
                        'BSW': '//*[@data-automationid="ListCell"][3]',
                        'BUC': '//*[@data-automationid="ListCell"][4]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]',                        
                    }
                },
                'RDX': {   #copy this Line v
                    'model_page_xpath': '//*[@data-automationid="ListCell"][2]',  
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][3]', 
                        'AEB': '//*[@data-automationid="ListCell"][5]',  
                        'AHL': '//*[@data-automationid="ListCell"][4]',
                        'APA': '//*[@data-automationid="ListCell"][2]',
                        'BSW': '//*[@data-automationid="ListCell"][6]',
                        'BUC': '//*[@data-automationid="ListCell"][1]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]',  
                    }
                },        #to this time ^
                'RL': {
                    'model_page_xpath': '//*[@data-automationid="ListCell"][3]', 
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][1]', 
                        'AEB': '//*[@data-automationid="ListCell"][2]',  
                        'AHL': '//*[@data-automationid="ListCell"][5]',
                        'APA': '//*[@data-automationid="ListCell"][3]',
                        'BSW': '//*[@data-automationid="ListCell"][6]',
                        'BUC': '//*[@data-automationid="ListCell"][4]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]', 
                    }
                },
                'TL': {
                    'model_page_xpath': '//*[@data-automationid="ListCell"][4]', 
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][4]', 
                        'AEB': '//*[@data-automationid="ListCell"][6]',  
                        'AHL': '//*[@data-automationid="ListCell"][5]',
                        'APA': '//*[@data-automationid="ListCell"][3]',
                        'BSW': '//*[@data-automationid="ListCell"][1]',
                        'BUC': '//*[@data-automationid="ListCell"][2]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]',  
                    }
                },
                'TSX': {
                    'model_page_xpath': '//*[@data-automationid="ListCell"][5]',
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][1]', 
                        'AEB': '//*[@data-automationid="ListCell"][5]',  
                        'AHL': '//*[@data-automationid="ListCell"][4]',
                        'APA': '//*[@data-automationid="ListCell"][2]',
                        'BSW': '//*[@data-automationid="ListCell"][6]',
                        'BUC': '//*[@data-automationid="ListCell"][1]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]',  
                    }
                },
                'ZDX': { 
                    'model_page_xpath': '//*[@data-automationid="ListCell"][6]',
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][1]', 
                        'AEB': '//*[@data-automationid="ListCell"][2]',  
                        'AHL': '//*[@data-automationid="ListCell"][2]',
                        'APA': '//*[@data-automationid="ListCell"][2]',
                        'BSW': '//*[@data-automationid="ListCell"][3]',
                        'BUC': '//*[@data-automationid="ListCell"][2]',
                        'LKA': '//*[@data-automationid="ListCell"][2]',
                        'NV': '//*[@data-automationid="ListCell"][2]',
                        'SVC': '//*[@data-automationid="ListCell"][2]',  
                    }
                },                
            }
        },
        '2022': {
                           ###################
                           #                 #
                           #      2022       #
                           #                 #
                           ###################
            
        'year_page_xpath': '//*[@data-automationid="ListCell"][2]',  
            'models': {
                'MDX': {
                    'model_page_xpath': '//*[@data-automationid="ListCell"][1]',  
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][1]', 
                        'AEB': '//*[@data-automationid="ListCell"][2]',  
                        'AHL': '//*[@data-automationid="ListCell"][6]',
                        'APA': '//*[@data-automationid="ListCell"][5]',
                        'BSW': '//*[@data-automationid="ListCell"][3]',
                        'BUC': '//*[@data-automationid="ListCell"][4]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]',                        
                    }
                },
                'RDX': {   #copy this Line v
                    'model_page_xpath': '//*[@data-automationid="ListCell"][2]',  
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][3]', 
                        'AEB': '//*[@data-automationid="ListCell"][5]',  
                        'AHL': '//*[@data-automationid="ListCell"][4]',
                        'APA': '//*[@data-automationid="ListCell"][2]',
                        'BSW': '//*[@data-automationid="ListCell"][6]',
                        'BUC': '//*[@data-automationid="ListCell"][1]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]',  
                    }
                },        #to this time ^
                'RL': {
                    'model_page_xpath': '//*[@data-automationid="ListCell"][3]', 
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][1]', 
                        'AEB': '//*[@data-automationid="ListCell"][2]',  
                        'AHL': '//*[@data-automationid="ListCell"][5]',
                        'APA': '//*[@data-automationid="ListCell"][3]',
                        'BSW': '//*[@data-automationid="ListCell"][6]',
                        'BUC': '//*[@data-automationid="ListCell"][4]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]', 
                    }
                },
                'TL': {
                    'model_page_xpath': '//*[@data-automationid="ListCell"][4]', 
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][4]', 
                        'AEB': '//*[@data-automationid="ListCell"][6]',  
                        'AHL': '//*[@data-automationid="ListCell"][5]',
                        'APA': '//*[@data-automationid="ListCell"][3]',
                        'BSW': '//*[@data-automationid="ListCell"][1]',
                        'BUC': '//*[@data-automationid="ListCell"][2]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]',  
                    }
                },
                'TSX': {
                    'model_page_xpath': '//*[@data-automationid="ListCell"][5]',
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][1]', 
                        'AEB': '//*[@data-automationid="ListCell"][5]',  
                        'AHL': '//*[@data-automationid="ListCell"][4]',
                        'APA': '//*[@data-automationid="ListCell"][2]',
                        'BSW': '//*[@data-automationid="ListCell"][6]',
                        'BUC': '//*[@data-automationid="ListCell"][1]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]',  
                    }
                },
                'ZDX': { 
                    'model_page_xpath': '//*[@data-automationid="ListCell"][6]',
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][1]', 
                        'AEB': '//*[@data-automationid="ListCell"][2]',  
                        'AHL': '//*[@data-automationid="ListCell"][2]',
                        'APA': '//*[@data-automationid="ListCell"][2]',
                        'BSW': '//*[@data-automationid="ListCell"][3]',
                        'BUC': '//*[@data-automationid="ListCell"][2]',
                        'LKA': '//*[@data-automationid="ListCell"][2]',
                        'NV': '//*[@data-automationid="ListCell"][2]',
                        'SVC': '//*[@data-automationid="ListCell"][2]',  
                    }
                },                
            }
        },
        '2023': {
                           ###################
                           #                 #
                           #      2023       #
                           #                 #
                           ###################
            
        'year_page_xpath': '//*[@data-automationid="ListCell"][2]',  
            'models': {
                'MDX': {
                    'model_page_xpath': '//*[@data-automationid="ListCell"][1]',  
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][1]', 
                        'AEB': '//*[@data-automationid="ListCell"][2]',  
                        'AHL': '//*[@data-automationid="ListCell"][6]',
                        'APA': '//*[@data-automationid="ListCell"][5]',
                        'BSW': '//*[@data-automationid="ListCell"][3]',
                        'BUC': '//*[@data-automationid="ListCell"][4]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]',                        
                    }
                },
                'RDX': {   #copy this Line v
                    'model_page_xpath': '//*[@data-automationid="ListCell"][2]',  
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][3]', 
                        'AEB': '//*[@data-automationid="ListCell"][5]',  
                        'AHL': '//*[@data-automationid="ListCell"][4]',
                        'APA': '//*[@data-automationid="ListCell"][2]',
                        'BSW': '//*[@data-automationid="ListCell"][6]',
                        'BUC': '//*[@data-automationid="ListCell"][1]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]',  
                    }
                },        #to this time ^
                'RL': {
                    'model_page_xpath': '//*[@data-automationid="ListCell"][3]', 
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][1]', 
                        'AEB': '//*[@data-automationid="ListCell"][2]',  
                        'AHL': '//*[@data-automationid="ListCell"][5]',
                        'APA': '//*[@data-automationid="ListCell"][3]',
                        'BSW': '//*[@data-automationid="ListCell"][6]',
                        'BUC': '//*[@data-automationid="ListCell"][4]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]', 
                    }
                },
                'TL': {
                    'model_page_xpath': '//*[@data-automationid="ListCell"][4]', 
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][4]', 
                        'AEB': '//*[@data-automationid="ListCell"][6]',  
                        'AHL': '//*[@data-automationid="ListCell"][5]',
                        'APA': '//*[@data-automationid="ListCell"][3]',
                        'BSW': '//*[@data-automationid="ListCell"][1]',
                        'BUC': '//*[@data-automationid="ListCell"][2]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]',  
                    }
                },
                'TSX': {
                    'model_page_xpath': '//*[@data-automationid="ListCell"][5]',
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][1]', 
                        'AEB': '//*[@data-automationid="ListCell"][5]',  
                        'AHL': '//*[@data-automationid="ListCell"][4]',
                        'APA': '//*[@data-automationid="ListCell"][2]',
                        'BSW': '//*[@data-automationid="ListCell"][6]',
                        'BUC': '//*[@data-automationid="ListCell"][1]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]',  
                    }
                },
                'ZDX': { 
                    'model_page_xpath': '//*[@data-automationid="ListCell"][6]',
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][1]', 
                        'AEB': '//*[@data-automationid="ListCell"][2]',  
                        'AHL': '//*[@data-automationid="ListCell"][2]',
                        'APA': '//*[@data-automationid="ListCell"][2]',
                        'BSW': '//*[@data-automationid="ListCell"][3]',
                        'BUC': '//*[@data-automationid="ListCell"][2]',
                        'LKA': '//*[@data-automationid="ListCell"][2]',
                        'NV': '//*[@data-automationid="ListCell"][2]',
                        'SVC': '//*[@data-automationid="ListCell"][2]',  
                    }
                },                
            }
        },
        '2024': {
                           ###################
                           #                 #
                           #      2024       #
                           #                 #
                           ###################
        'year_page_xpath': '//*[@data-automationid="ListCell"][2]',  
            'models': {
                'MDX': {
                    'model_page_xpath': '//*[@data-automationid="ListCell"][1]',  
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][1]', 
                        'AEB': '//*[@data-automationid="ListCell"][2]',  
                        'AHL': '//*[@data-automationid="ListCell"][6]',
                        'APA': '//*[@data-automationid="ListCell"][5]',
                        'BSW': '//*[@data-automationid="ListCell"][3]',
                        'BUC': '//*[@data-automationid="ListCell"][4]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]',                        
                    }
                },
                'RDX': {   #copy this Line v
                    'model_page_xpath': '//*[@data-automationid="ListCell"][2]',  
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][3]', 
                        'AEB': '//*[@data-automationid="ListCell"][5]',  
                        'AHL': '//*[@data-automationid="ListCell"][4]',
                        'APA': '//*[@data-automationid="ListCell"][2]',
                        'BSW': '//*[@data-automationid="ListCell"][6]',
                        'BUC': '//*[@data-automationid="ListCell"][1]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]',  
                    }
                },        #to this time ^
                'RL': {
                    'model_page_xpath': '//*[@data-automationid="ListCell"][3]', 
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][1]', 
                        'AEB': '//*[@data-automationid="ListCell"][2]',  
                        'AHL': '//*[@data-automationid="ListCell"][5]',
                        'APA': '//*[@data-automationid="ListCell"][3]',
                        'BSW': '//*[@data-automationid="ListCell"][6]',
                        'BUC': '//*[@data-automationid="ListCell"][4]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]', 
                    }
                },
                'TL': {
                    'model_page_xpath': '//*[@data-automationid="ListCell"][4]', 
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][4]', 
                        'AEB': '//*[@data-automationid="ListCell"][6]',  
                        'AHL': '//*[@data-automationid="ListCell"][5]',
                        'APA': '//*[@data-automationid="ListCell"][3]',
                        'BSW': '//*[@data-automationid="ListCell"][1]',
                        'BUC': '//*[@data-automationid="ListCell"][2]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]',  
                    }
                },
                'TSX': {
                    'model_page_xpath': '//*[@data-automationid="ListCell"][5]',
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][1]', 
                        'AEB': '//*[@data-automationid="ListCell"][5]',  
                        'AHL': '//*[@data-automationid="ListCell"][4]',
                        'APA': '//*[@data-automationid="ListCell"][2]',
                        'BSW': '//*[@data-automationid="ListCell"][6]',
                        'BUC': '//*[@data-automationid="ListCell"][1]',
                        'LKA': '//*[@data-automationid="ListCell"][7]',
                        'NV': '//*[@data-automationid="ListCell"][8]',
                        'SVC': '//*[@data-automationid="ListCell"][10]',  
                    }
                },
                'ZDX': { 
                    'model_page_xpath': '//*[@data-automationid="ListCell"][6]',
                    'documents': {
                        'ACC': '//*[@data-automationid="ListCell"][1]', 
                        'AEB': '//*[@data-automationid="ListCell"][2]',  
                        'AHL': '//*[@data-automationid="ListCell"][2]',
                        'APA': '//*[@data-automationid="ListCell"][2]',
                        'BSW': '//*[@data-automationid="ListCell"][3]',
                        'BUC': '//*[@data-automationid="ListCell"][2]',
                        'LKA': '//*[@data-automationid="ListCell"][2]',
                        'NV': '//*[@data-automationid="ListCell"][2]',
                        'SVC': '//*[@data-automationid="ListCell"][2]',  
                    }
                },                
            }
        },
        # Copy New Years here, where the "#" is
    }

    try:
        # Navigate to the main SharePoint page for Acura
        print("Navigating to Acura's main SharePoint page...")
        driver.get('https://calibercollision-my.sharepoint.com/:f:/g/personal/mark_klingenhofer_protechdfw_com/EjIo8sg9qXNEt6CDCKpeRGkBEY-67TppBRysHPrqdbNSmg')
        
        # Clicks Acura
        print("Locating Acura link and clicking...")
        acura = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="appRoot"]/div/div[2]/div/div/div[2]/div[2]/main/div/div/div[2]/div/div/div/div/div[2]/div/div/div/div[2]/div/div/div[3]/div/div[1]/span/span[1]/button')))
        acura.click()
        time.sleep(1)
        adas_last_row = {}
        wb = load_workbook(excel_path)
        ws = wb['Sheet1']  # Correctly referencing the worksheet
      

        print(f"Workbook loaded successfully: {excel_path}")
    
        for year, data in years_models_documents.items():
            # Clicks the year
            print(f"Processing year: {year}")
            year_page_xpath = data['year_page_xpath']
            double_click_element(driver, wait, year_page_xpath)
            time.sleep(1)
            
            for model, model_data in data['models'].items():
                # Clicks the model
                print(f"Accessing model: {model}")
                model_page_xpath = model_data['model_page_xpath']
                double_click_element(driver, wait, model_page_xpath)
                time.sleep(1)
                
                for doc_name, doc_xpath in model_data['documents'].items():
                    double_click_element(driver, wait, model_page_xpath)
                    print(f"Retrieving document: {doc_name}")
                    document_url = get_document_url(driver, wait, doc_xpath)
                    
                      # Using adas_system correctly here
                    print(f"Document URL retrieved: {document_url}")
                    
                    # Take screenshot and get text
                    extracted_text = screenshot_and_get_text(driver)
    
                    # Correcting the function call
                    if doc_name in adas_last_row:
                        next_row = adas_last_row[doc_name] + 1
                    else:
                        cell = find_row_in_excel(ws, year, "Acura", model, doc_name)  # Correct function call
                        if cell:                           
                            next_row = cell.row
                        else:
                            next_row = ws.max_row + 1                           

                    print(f"Hyperlink for {doc_name} added at {cell.coordinate}")
                    wb.save(excel_path)
                    
                    # Go back to model page to get the next document's URL
                    print("Returning to model page...")
                    driver.back()
                
                # Goes back to the year's page to select the next model
                print("Returning to year page...")    
                driver.back()  # Ensure this takes you back to the correct page

            # Goes back to the Acura main page to select the next year
            print("Returning to Acura's main page...")    
            driver.back()  # Ensure this takes you back to the correct page
            
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        print("Script Sucessfully Completed, you may exit the program and terminal at any time ")
        driver.quit()
        
if __name__ == "__main__":
    excel_file_path = sys.argv[1]  # The Excel file path is expected as the first argument
    run_acura_script(excel_file_path)