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
        # Double click on the document link 
        double_click_element(driver, wait, document_xpath)
    
        # Wait for the new tab to appear and then switch to it
        time.sleep(3)  # Wait to ensure the new tab is loaded
 
        # Capture the URL from the new tab
        document_url = driver.current_url
        time.sleep(3)  # Wait to ensure the new tab is loaded
        driver.back

        return document_url

def navigate_to_model(driver, wait, model_xpath):
     model_link = wait.until(EC.element_to_be_clickable((By.XPATH, model_xpath)))
     model_link.click()
     time.sleep(2)  # Wait for the model's page to load

def navigate_to_year(driver, wait, year_xpath):

     # You might need to switch to iframe here if the element is inside an iframe
     # driver.switch_to.frame("frame_name_or_id")

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
                    'model_page_xpath': '//*[@data-automationid="ListCell"][6]',  ############# Continue from here
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
        '2013': {
        'year_page_xpath': '//*[@id="appRoot"]/div/div[2]/div/div/div[2]/div[2]/main/div/div/div[2]/div/div/div/div/div[2]/div/div/div/div[3]', 
            'models': {
                'MDX': {
                    'model_page_xpath': '//*[@id="appRoot"]/div/div[2]/div/div/div[2]/div[2]/main/div/div/div[2]/div/div/div/div/div[2]/div/div/div/div[1]',  
                    'documents': {
                        'ACC': '//*[@id="appRoot"]/div/div[2]/div/div/div[2]/div[2]/main/div/div/div[2]/div/div/div/div/div[2]/div/div/div/div[1]', 
                        'AEB': '//*[@id="appRoot"]/div/div[2]/div/div/div[2]/div[2]/main/div/div/div[2]/div/div/div/div/div[2]/div/div/div/div[2]',  
                        'AHL': '//*[@id="appRoot"]/div/div[2]/div/div/div[2]/div[2]/main/div/div/div[2]/div/div/div/div/div[2]/div/div/div/div[6]',
                        'APA': '//*[@id="appRoot"]/div/div[2]/div/div/div[2]/div[2]/main/div/div/div[2]/div/div/div/div/div[2]/div/div/div/div[5]',
                        'BSW': '//*[@id="appRoot"]/div/div[2]/div/div/div[2]/div[2]/main/div/div/div[2]/div/div/div/div/div[2]/div/div/div/div[3]',
                        'BUC': '//*[@id="appRoot"]/div/div[2]/div/div/div[2]/div[2]/main/div/div/div[2]/div/div/div/div/div[2]/div/div/div/div[4]',
                        'LKA': '//*[@id="appRoot"]/div/div[2]/div/div/div[2]/div[2]/main/div/div/div[2]/div/div/div/div/div[2]/div/div/div/div[7]',
                        'NV': '//*[@id="appRoot"]/div/div[2]/div/div/div[2]/div[2]/main/div/div/div[2]/div/div/div/div/div[2]/div/div/div/div[8]',
                        'SVC': '//*[@id="appRoot"]/div/div[2]/div/div/div[2]/div[2]/main/div/div/div[2]/div/div/div/div/div[2]/div/div/div/div[10]',
                    }
                },
                'RDX': {
                    'model_page_xpath': '//*[@id="appRoot"]/div/div[2]/div/div/div[2]/div[2]/main/div/div/div[2]/div/div/div/div/div[2]/div/div/div/div[2]',  
                    'documents': {
                        'ACC': '//*[@id="appRoot"]/div/div[2]/div/div/div[2]/div[2]/main/div/div/div[2]/div/div/div/div/div[2]/div/div/div/div[1]', 
                        'AEB': '//*[@id="appRoot"]/div/div[2]/div/div/div[2]/div[2]/main/div/div/div[2]/div/div/div/div/div[2]/div/div/div/div[2]',  
                        'AHL': '//*[@id="appRoot"]/div/div[2]/div/div/div[2]/div[2]/main/div/div/div[2]/div/div/div/div/div[2]/div/div/div/div[3]',
                        'APA': '//*[@id="appRoot"]/div/div[2]/div/div/div[2]/div[2]/main/div/div/div[2]/div/div/div/div/div[2]/div/div/div/div[4]',
                        'BSW': '//*[@id="appRoot"]/div/div[2]/div/div/div[2]/div[2]/main/div/div/div[2]/div/div/div/div/div[2]/div/div/div/div[1]',
                        'BUC': '//*[@id="appRoot"]/div/div[2]/div/div/div[2]/div[2]/main/div/div/div[2]/div/div/div/div/div[2]/div/div/div/div[1]',
                        'LKA': '//*[@id="appRoot"]/div/div[2]/div/div/div[2]/div[2]/main/div/div/div[2]/div/div/div/div/div[2]/div/div/div/div[1]',
                        'NV': '//*[@id="appRoot"]/div/div[2]/div/div/div[2]/div[2]/main/div/div/div[2]/div/div/div/div/div[2]/div/div/div/div[1]',
                        'SVC': '//*[@id="appRoot"]/div/div[2]/div/div/div[2]/div[2]/main/div/div/div[2]/div/div/div/div/div[2]/div/div/div/div[1]', 
                    }
                },
            }
        },
        # ... repeat this structure for other years
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
        current_row = 2  # Start at row 2, assuming row 1 has headers
        
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
                    print(f"Document URL retrieved: {document_url}")
                    
                    # Define the correct cell_address for each document
                    cell_address = f'L{current_row}'  # L column, next available row
                    add_hyperlink_to_excel(excel_path, 'Sheet1', cell_address, document_url, doc_name)
                    print (f"The URL for {doc_name} was placed in {cell_address}")
                    current_row += 1  # Move to the next row for the next document
                    
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