from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
import time

# Set up Chrome options
chrome_options = Options()
chrome_options.add_experimental_option("detach", True)

# Set up the Chrome WebDriver
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

# Open a website
driver.get('https://calibercollision-my.sharepoint.com/:f:/g/personal/mark_klingenhofer_protechdfw_com/EjIo8sg9qXNEt6CDCKpeRGkBEY-67TppBRysHPrqdbNSmg')

# Sleep for a moment to ensure the page loads
time.sleep(5)

# Initialize the ActionChains object
action_chains = ActionChains(driver)

# Using the given XPath to find and click on the button
acura = driver.find_element(By.XPATH, '//*[@id="appRoot"]/div/div[2]/div/div/div[2]/div[2]/main/div/div/div[2]/div/div/div/div/div[2]/div/div/div/div[2]/div/div/div[3]/div/div[1]/span/span[1]/button')
acura.click()

acurayear2012 = driver.find_element(By.XPATH, '//*[@id="appRoot"]/div/div[2]/div/div/div[2]/div[2]/main/div/div/div[2]/div/div/div/div/div[2]/div/div/div/div[2]')

# Perform the double click action on the acurayear2012 element
action_chains.double_click(acurayear2012).perform()

time.sleep(5)

# Wait for user input to close the browser
input("Press Enter to quit...")
driver.quit()

