from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.common.by import By
from openpyxl import load_workbook
import pyperclip
import time

# Function to safely perform an action
def safe_action(action, description, max_attempts=3):
    attempts = 0
    while attempts < max_attempts:
        try:
            action()
            return
        except (NoSuchElementException, TimeoutException) as e:
            print(f"Error performing {description}: {e}")
            time.sleep(2)  # Wait before retrying
            attempts += 1
    raise Exception(f"Failed to perform {description} after {max_attempts} attempts")

# Set up Chrome options
chrome_options = Options()
chrome_options.add_experimental_option("detach", True)

# Set up the Chrome WebDriver
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

# Open a website
driver.get('https://calibercollision-my.sharepoint.com/:f:/g/personal/mark_klingenhofer_protechdfw_com/EjIo8sg9qXNEt6CDCKpeRGkBEY-67TppBRysHPrqdbNSmg')

wait = WebDriverWait(driver, 10)
action_chains = ActionChains(driver)

# Define the actions to perform
def click_acura():
    acura = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="appRoot"]/div/div[2]/div/div/div[2]/div[2]/main/div/div/div[2]/div/div/div/div/div[2]/div/div/div/div[2]/div/div/div[3]/div/div[1]/span/span[1]/button')))
    acura.click()

def double_click_year2012():
    acurayear2012 = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="appRoot"]/div/div[2]/div/div/div[2]/div[2]/main/div/div/div[2]/div/div/div/div/div[2]/div/div/div/div[2]')))
    action_chains.double_click(acurayear2012).perform()

def double_click_model():
    acura2012mdx = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="appRoot"]/div/div[2]/div/div/div[2]/div[2]/main/div/div/div[2]/div/div/div/div/div[2]/div/div/div/div[1]')))
    action_chains.double_click(acura2012mdx).perform()

def click_context_menu():
    context_menu_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="appRoot"]/div/div[2]/div/div/div[2]/div[2]/main/div/div/div[2]/div/div/div/div/div[2]/div/div/div/div[1]/div/div/div[3]/div/div[2]/div/button')))
    context_menu_button.click()

def click_open():
    open_button = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/div/div/div/div/div/div/ul/li[1]/button/div/span')))
    open_button.click()

def click_open_in_browser():
    open_in_browser_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Open in browser')]/ancestor::button")))
    open_in_browser_button.click()

# Perform actions with error handling
safe_action(click_acura, "clicking Acura")
safe_action(double_click_year2012, "double-clicking Year 2012")
safe_action(double_click_model, "double-clicking Model")
safe_action(click_context_menu, "clicking Context Menu")
safe_action(click_open, "clicking Open")
safe_action(click_open_in_browser, "clicking Open in Browser")

# Switch to new tab to get the URL
driver.switch_to.window(driver.window_handles[1])
document_url = driver.current_url

# Copy the URL to the clipboard
pyperclip.copy(document_url)

driver.close()  # Close the new tab
driver.switch_to.window(driver.window_handles[0])  # Switch back to the original tab

# Load the workbook and select the active worksheet
wb = load_workbook('path_to_your_excel_file.xlsx')
ws = wb.active

# Paste the URL into a particular cell, say 'A1'
cell_to_update = 'L2'  # Change this as needed
ws[cell_to_update] = pyperclip.paste()

# Save the workbook
wb.save('path_to_your_excel_file.xlsx')

# Wait for user input to close the browser
input("Press Enter to quit...")
driver.quit()