import os
import re
import sys
import time
import openpyxl
import urllib.parse
import tkinter as tk
from enum import Enum
import win32clipboard
from tkinter import messagebox
from selenium import webdriver
from openpyxl.styles import Font
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.edge.webdriver import WebDriver
from selenium.webdriver.common.window import WindowTypes
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support import expected_conditions as EC


#####################################################################################################################################################

class SharepointExtractor: 
    """
    Class definition for a SharepointExtractor. 
    Navigates to a given sharepoint location and extracts all the links for files and folders inside of the given location
    """
  
    # Attributes for the SharepointExtractor instance
    sharepoint_make: str = None             # The make of the sharepoint folder in use
    sharepoint_link: str = None             # The link to the sharepoint location for the given make
    excel_file_path: str = None             # The fully qualified path to the excel file holding our ADAS SI
    selenium_driver: WebDriver = None       # The selenium driver instance for this session
    selenium_wait: WebDriverWait = None     # The default wait operator for the webdriver
    
    # Configuration attributes for the sharepoint module names and timeouts
    __MAX_WAIT_TIME__ = 120
    __DEBUG_RUN__ = False

    # Locators used to find objects on the sharepoint folder pages
    __ONEDRIVE_PAGE_NAME_LOCATOR__ = "//li[@data-automationid='breadcrumb-listitem']//div[@data-automationid='breadcrumb-crumb']"
    __ONEDRIVE_TABLE_LOCATOR__ = "//div[@data-automationid='list-page']"
    __ONEDRIVE_TABLE_ROW_LOCATOR__ = "./div[@role='row' and starts-with(@data-automationid, 'row-')]"

    # Collections of system names used for finding correct files and row locations
    __DEFINED_MODULE_NAMES__ = [ 'ACC', 'SCC', 'AEB', 'AHL', 'APA','BSW', 'BSW/RCTW', 'BSW-RCTW','BSW & RCTW','BSW RCTW','BSW-RCT W','BSW RCT W','BSM-RCTW','BSW-RTCW','BSW_RCTW','BCW-RCTW', 'BUC', 'LKA', 'LW', 'NV', 'SVC', 'WAMC' ]
    __ROW_SEARCH_TERMS__ = ['LKAS', 'FCW/LDW', 'Multipurpose', 'Cross Traffic Alert', 'Side Blind Zone Alert', 'Lane Change Alert', 'Blind Spot Warning (BSW)', 'Surround Vision Camera', 'Video Processing', 'Pending Further Research',]
    __ADAS_SYSTEMS_WHITELIST__ = [
        'FCW/LDW',
        'FCW-LDW',
        'Multipurpose Camera',
        'Multipurpose',
        'WL',
        '[WL]',
        'Forward Collision Warning/Lane Departure Warning (FCW/LDW)',
        'Blind Spot Warning (BSW)',
        'Cross Traffic Alert',
        ' Side Blind Zone Alert',
        'Side Blind Zone Alert',
        'Surround Vision Camera',
        'Video Processing'
    ]
    SPECIFIC_HYPERLINKS = {
    "2013 GMC Acadia (BSW-RCTW 1)-PL-PW072NLB.pdf": "L79",
    "2016 Acura RLX (LKA 1) [FCW-LDW].pdf": "L249",
    "2016 Acura RLX (LKA 1) [Multipurpose].pdf": "L250",
    "2012 Volkswagen CC (ACC 1).pdf": "L11",
    "2013 Volkswagen CC (ACC 1).pdf": "L83",
    "2014 Volkswagen CC (ACC 1).pdf": "L155",
    "2015 Volkswagen CC (ACC 1).pdf": "L227",
    "2016 Volkswagen CC (ACC 1).pdf": "L299",
    "2017 Volkswagen CC (ACC 1).pdf": "L371"
    # Add more mappings as needed
    }
    #################################################################################################################################################

    # Class objects holding information about files and folders in a given sharepoint location
    class EntryTypes(Enum):
        """Enumeration holding the different types of entries in a sharepoint location"""
        UNDEFINED = "N/A",
        FILE_ENTRY = "File",
        FOLDER_ENTRTY = "Folder"
    class SharepointEntry:
        """Base type for a file or folder in a sharepoint location"""
        
        # Attributes for a sharepoint entry object
        entry_name: str = None                                  # The name of the entry in the sharepoint location
        entry_link: str = None                                  # The link to the entry in the sharepoint location
        entry_heirarchy: str = None                             # The folder path/heirarchy to the entry in our sharepoint location        
        entry_type: 'SharepointExtractor.EntryType' = None      # The type of entry in the sharepoint location
        
        def __init__(self, name: str, heirarchy: str, link: str, type: 'SharepointExtractor.EntryType') -> 'SharepointExtractor.SharepointEntry':
            """
            CTOR for building a new SharepointEntry object
            
            ----------------------------------------------
            
            name: str
                The name of the entry in the sharepoint folder
            heirarchy: str  
                The heirachy path to the entry in the sharepoint folder
            link: str
                The link to the entry in the sharepoint folder
            type: SharepointExtractor.EntryType
                The type of entry in our sharepoint folder
            """
            
            # Assign properties of the entry to this instance and exit out 
            self.entry_name = name
            self.entry_link = link
            self.entry_type = type
            self.entry_heirarchy = heirarchy

    #################################################################################################################################################  

    def __init__(self, sharepoint_link: str, excel_file_path: str, debug_run: bool = False) -> 'SharepointExtractor':
        """
        CTOR for a new SharepointExtractor. Takes the link to the requested sharepoint location 
        and prepares to extract all file and folder links
        
        ----------------------------------------------
        
        sharepoint_link: str 
            The link to the sharepoint location for the given make
        excel_file_path: str 
            The fully qualified path to the excel file holding our ADAS SI
        debug_run: bool
            When true, we don't actually get any file links.
            Useful for quickly testing operations without waiting for links to generate.
            Defaults to false.
        """
        
        # Store attributes for the Extractor on this instance
        self.__DEBUG_RUN__ = debug_run
        self.sharepoint_link = sharepoint_link
        self.excel_file_path = excel_file_path
        self.selected_adas = sys.argv[3].split(",") if len(sys.argv) > 3 else []

        # Define the default wait timeout and setup a new selenium driver
        self.selenium_driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=self.__generate_chrome_options__())
        self.selenium_wait = WebDriverWait(self.selenium_driver, 10)

        # Navigate to the main SharePoint page for Acura
        print("Navigating to main SharePoint page link now...")
        self.selenium_driver.get(sharepoint_link)
 
        # Wait until the element with the specified XPath is found, or until 60 seconds have passed
        try: WebDriverWait(self.selenium_driver, self.__MAX_WAIT_TIME__).until(EC.presence_of_element_located((By.XPATH, self.__ONEDRIVE_TABLE_LOCATOR__)))
        except: 
            print(f"The element was not found within {self.__MAX_WAIT_TIME__} seconds.")
            raise Exception(f"ERROR! Failed to find valid login state after {self.__MAX_WAIT_TIME__} seconds!")
        
        # Find the make of the folder for the current sharepoint link and store it.
        self.sharepoint_make = self.selenium_driver.find_elements(By.XPATH, self.__ONEDRIVE_PAGE_NAME_LOCATOR__)[-1].get_attribute("innerText").strip()
        print(f"Configured new SharepointExtractor for {self.sharepoint_make} correctly!")
        
    def extract_contents(self) -> tuple[list, list]:
        """
        Extracts the file and folder links from the defined sharepoint location for the current extractor object.
        Returns a tuple of lists. The first list holds all of our SharepointEntry objects for the folders in the sharepoint,
        and the second list holds all of our SharepointEntry objects for the files in the sharepoint.
        If no ADAS systems are selected, processes all files.
        """
        # Index and store base folders and files here
        sharepoint_folders, sharepoint_files = self.__get_folder_rows__()
    
        # Compile regex patterns for the selected ADAS systems if any are selected
        adas_patterns = (
            [re.compile(rf"\({re.escape(adas)}\s*\d*\)", re.IGNORECASE) for adas in self.selected_adas]
            if self.selected_adas else None
        )
    
        # Initialize filtered files list
        filtered_files = []
    
        # Start indexing
        start_time = time.time()
        while len(sharepoint_folders) > 0:
            # Store the current folder value and navigate to it for indexing
            folder_link = sharepoint_folders.pop(0).entry_link
            child_folders, child_files = self.__get_folder_rows__(folder_link)
    
            # Add child folders for further processing
            sharepoint_folders.extend(child_folders)
    
            # Add all files if no ADAS systems are selected, otherwise filter them
            if adas_patterns:
                # Filter and add matching child files
                for file_entry in child_files:
                    if any(pattern.search(file_entry.entry_name) for pattern in adas_patterns):
                        filtered_files.append(file_entry)
                        
                    #else:
                        #print(f"Skipping file (not matching ADAS): {file_entry.entry_name}")
            else:
                # Add all files directly
                filtered_files.extend(child_files)
                #for file_entry in child_files:
                    #print(f"File added (no ADAS filter): {file_entry.entry_name}")
    
            # Log out how many child links and folders exist now
            print(f'{len(sharepoint_folders)} Folders Remain | {len(filtered_files)} Files Indexed')
    
        elapsed_time = time.time() - start_time
        print(f"Indexing routine took {elapsed_time:.2f} seconds.")
        return [sharepoint_folders, filtered_files]




    def populate_excel_file(self, file_entries: list) -> None:
        """
        Populates the excel file for the current make and stores all hyperlinks built in correct 
        locations.
    
        file_entries: list[SharepointEntry]
            The list of all file entries we're looking to put into our excel file
        """
    
        # Load the Excel file
        start_time = time.time()
        model_workbook = openpyxl.load_workbook(self.excel_file_path)
        model_worksheet = model_workbook['Model Version']
        print(f"Workbook loaded successfully: {self.excel_file_path}")
    
        # Setup trackers for correct row insertion during population
        current_model = ""
        adas_last_row = {}
    
        # Iterate through the filtered file entries
        for file_entry in file_entries:
            print(f"Processing file: {file_entry.entry_name}")
    
            # Pull the year and model for the file from the hierarchy
            hierarchy_segments = file_entry.entry_heirarchy.split('\\')
            if len(hierarchy_segments) < 3:
                print(f"Invalid entry hierarchy format: {file_entry.entry_heirarchy}")
                continue
    
            file_name = file_entry.entry_name
            file_model = hierarchy_segments[2]
            file_year = hierarchy_segments[1]
    
            # Check if ADAS last row needs to be reset or not
            if file_model != current_model:
                current_model = file_model
                adas_last_row = {}
    
            # Now update our excel file based on the values given for this entry
            if self.__update_excel_with_whitelist__(model_worksheet, file_name, file_entry.entry_link):
                continue
            self.__update_excel__(model_worksheet, file_year, file_model, file_name, file_entry.entry_link, adas_last_row, None)
    
        # Save the workbook after processing
        print(f"Saving updated changes to {self.sharepoint_make} sheet now...")
        model_workbook.save(self.excel_file_path)
        model_workbook.close()
    
        # Log the time taken to populate
        elapsed_time = time.time() - start_time
        print(f"Sheet population routine took {elapsed_time:.2f} seconds.")



    def __generate_chrome_options__(self) -> Options:
        """
        Configures Chrome to always use the existing user profile for multi-document processing.
        Returns a built set of Chrome options configured to use the existing profile.
        """
        
        # Define a new chrome options object and setup some default configuration
        chrome_options = Options()
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-extensions")  # Disable extensions to avoid conflicts
        chrome_options.add_argument("--disable-infobars")  # Disable infobars
        chrome_options.add_argument("--disable-browser-side-navigation")  # Disable side navigation issues
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")  # Avoid detection as bot
    
        # Always use the existing Chrome profile
        home_dir = os.path.expanduser("~")
        user_data_dir = os.path.join(home_dir, "AppData", "Local", "Google", "Chrome", "User Data")
        chrome_options.add_argument(f"--user-data-dir={user_data_dir}")
        profile_dir = "Default" 
        chrome_options.add_argument(f"--profile-directory={profile_dir}")
        
        # Return the updated options object
        return chrome_options
  
    
    def __is_row_folder__(self, row_element: WebElement) -> bool:
        # Get the row name from the aria-label attribute
        row_name = self.__get_row_name__(row_element)
        # If the row name contains a common file extension, assume it is a file.
        if re.search(r'\.(pdf|docx?|xlsx?|pptx?)$', row_name, re.IGNORECASE):
            return False
        # Otherwise, treat it as a folder.
        return True
     
    
    def __get_row_name__(self, row_element: WebElement) -> str:
        # Try to get the row name from the 'aria-label' attribute
        row_name = row_element.get_attribute("aria-label")
        if row_name and row_name.strip():
            return row_name.strip()
        # Fallback: return the text content if aria-label is not available
        return row_element.text.strip()
           
    
    def __get_unencrypted_link__(self, row_element: WebElement) -> str:
        """
        Generates the unencrypted link for a given row element.
    
        row_element: WebElement
            The row element for which the unencrypted link is to be generated.
    
        Returns:
        str
            The unencrypted link for the row element.
        """
        try:
            # Pull the folder name and add the name of it to our URL name
            base_url = self.selenium_driver.current_url.split("&p=true")[0]  # Current URL Split up for the path of the current folder
            row_name = self.__get_row_name__(row_element)                    # The name we're looking to open
            encoded_row_name = urllib.parse.quote(row_name)                  # URL-encode the row name to handle special characters
            plain_link = base_url + "%2F" + encoded_row_name                 # Relative folder URL based on drive layout

            # Return the built URL here
            return plain_link
    
        except IndexError as e:
            print(f"Error while generating unencrypted link: {e}")
            raise Exception(f"Failed to generate unencrypted link for row: {self.__get_row_name__(row_element)}")    
    def __get_encrypted_link__(self, row_element: WebElement) -> str:
                  
        # Debug run testing break out to speed things up
        if self.__DEBUG_RUN__:
            return f"Link For: {self.__get_row_name__(row_element)}"
        
        # Store a starting clipboard content value to ensure we get a new value during this method
        starting_clipboard_content = self.__get_clipboard_content__()
    
        # Find the selector element using the new locator that matches the row selection cell
        selector_element_locator = ".//div[@role='gridcell' and contains(@data-automationid, 'row-selection-')]"
        selector_element = WebDriverWait(row_element, self.__MAX_WAIT_TIME__).until(
            EC.presence_of_element_located((By.XPATH, selector_element_locator))
        )
        selector_element.click()
        
        # Attempt the share routine in a loop to retry when buttons don't appear correctly
        for retry_count in range(3):
            try:
                # Find the share button element using the new locator and click it
                row_element.find_element(By.XPATH, ".//button[@data-automationid='shareHeroId']").click()
                time.sleep(1.00)
                ActionChains(self.selenium_driver).send_keys(
                    Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.ENTER
                ).perform()
                time.sleep(1.25)
                ActionChains(self.selenium_driver).send_keys(
                    Keys.TAB, Keys.ARROW_DOWN, Keys.TAB, Keys.TAB, Keys.ENTER
                ).perform()           
                time.sleep(1.25)
                ActionChains(self.selenium_driver).send_keys(Keys.ENTER).perform()
                time.sleep(1.25)
                ActionChains(self.selenium_driver).send_keys(Keys.ESCAPE).perform()
                
                # Break this loop if this logic completes correctly
                break
            except:
                # Check if we can retry or not
                if retry_count == 3:
                    raise Exception("ERROR! Failed to open the share dialog for the current entry!")
                # Wait a moment before retrying
                time.sleep(1.0)
        
        # Unselect the element for the row 
        time.sleep(1.00)
        selector_element.click()
        
        # Make sure the link value is changed here. If it's not, run this routine again
        encrypted_file_link = self.__get_clipboard_content__()
        if encrypted_file_link == starting_clipboard_content:
            return self.__get_encrypted_link__(row_element)
    
        # Return the stored link from the clipboard
        return encrypted_file_link
         
    def __get_clipboard_content__(self) -> str:
            """
            Local helper method used to pull clipboard content for generated links
            Returns the link generated by onedrive
            """
            
            # Pull the clipboard content and store it, then dump the link contents out of it
            for retry_count in range(3):
            
                # Open the clipboard and pull our file link   
                try:    
                    win32clipboard.OpenClipboard()
                    encrypted_file_link = win32clipboard.GetClipboardData()
                    win32clipboard.CloseClipboard()

                    # Return the link generated here
                    return encrypted_file_link

                # On failures, retry opening the clipboard if possible
                except:
                
                    # Check if we can retry or not
                    if retry_count == 3:
                        raise Exception("ERROR! Failed to open the clipboard!")
                
                    # Wait a moment before retrying to open the clipboard 
                    win32clipboard.CloseClipboard()
                    time.sleep(1.0)    
    def __get_entry_heirarchy__(self, row_element: WebElement) -> str:
        # Find all breadcrumb elements
        title_elements = self.selenium_driver.find_elements(By.XPATH, self.__ONEDRIVE_PAGE_NAME_LOCATOR__)
        try:
            # Use a case-insensitive substring match rather than exact equality
            matching_element = next(title_element for title_element in title_elements
                                    if self.sharepoint_make.lower() in title_element.get_attribute("innerText").lower())
            title_index = title_elements.index(matching_element)
        except StopIteration:
            # If no match is found, fall back to the first element
            title_index = 0
    
        child_elements = title_elements[title_index:]
        entry_heirarchy = ""
        for child_element in child_elements:
            folder_name = child_element.get_attribute("innerText").strip()
            # (Keep your renaming logic here...)
            if folder_name == "RS3":
                folder_name = "RS 3"
            elif folder_name == "RS5":
                folder_name = "RS 5"
            elif folder_name == "RS6":
                folder_name = "RS 6"
            elif folder_name == "RS7":
                folder_name = "RS 7"
            elif folder_name == "SQ3":
                folder_name = "SQ 3"
            elif folder_name == "SQ5":
                folder_name = "SQ 5"
            elif folder_name == "SQ7":
                folder_name = "SQ 7"
            elif folder_name == "SQ8":
                folder_name = "SQ 8"
            elif folder_name == "VERANO":
                folder_name = "Verano"   
            elif folder_name == "Trailbalzer":
                folder_name = "Trailblazer"  
            elif folder_name == "Savanna":
                folder_name = "Savana"   
            elif folder_name == "Clarity":
                folder_name = "CLARITY ELECTRIC"  
            elif folder_name == "Clarity Plug In":
                folder_name = "CLARITY PLUG-IN"   
            elif folder_name == "EX35":
                folder_name = "EX"  
            elif folder_name == "G37 Convertible":
                folder_name = "G Convertible"  
            elif folder_name == "G37 Coupe":
                folder_name = "G Coupe"
            elif folder_name == "G37 Sedan":
                folder_name = "G Sedan"
            elif folder_name == "QX56":
                folder_name = "QX"   
            elif folder_name == "Grand Cherokee WL":          
                folder_name = "Grand Cherokee"                
            elif folder_name == "Wrangler (JL)":            
                folder_name = "Wrangler"   
            elif folder_name == "Wrangler JL":
                folder_name = "Wrangler"   
            elif folder_name == "K5 [Optima]":
                folder_name = "K5"
            elif folder_name == "K7 [Cadenza]":
                folder_name = "K7"   
            elif folder_name == "New Range Rover":
                folder_name = "Range Rover" 
            elif folder_name == "New Range Rover Evoque":
                folder_name = "Evoque" 
            elif folder_name == "New Range Rover Sport":
                folder_name = "Sport"    
            elif folder_name == "Range Rover Sport":
                folder_name = "Sport"     
            elif folder_name == "Range Rover Velar":
                folder_name = "Velar"     
            elif folder_name == "RCF":
                folder_name = "RC F"  
            elif folder_name == "CX3":
                folder_name = "CX-3" 
            elif folder_name == "CX30":
                folder_name = "CX-30" 
            elif folder_name == "CX5":
                folder_name = "CX-5" 
            elif folder_name == "CX50":
                folder_name = "CX-50"
            elif folder_name == "CX9":
                folder_name = "CX-9" 
            elif folder_name == "MX30":
                folder_name = "MX-30"   
            elif folder_name == "MX5":
                folder_name = "MX-5" 
            elif folder_name == "Mazda 2":
                folder_name = "2"  
            elif folder_name == "Mazda 3":
                folder_name = "3"   
            elif folder_name == "Mazda 5":
                folder_name = "5" 
            elif folder_name == "Mazda 6":
                folder_name = "6"  
            elif folder_name == "F54 Clubman":
                folder_name = "Clubman"  
            elif folder_name == "F55 Hardtop 4 Door":
                folder_name = "Hardtop 4D"
            elif folder_name == "F56 Hardtop 2 Door":
                folder_name = "Hardtop 2D" 
            elif folder_name == "F57 Convertible":
                folder_name = "Convertible"  
            elif folder_name == "F60 Countryman":
                folder_name = "Countryman" 
            elif folder_name == "Panamera 971":
                folder_name = "Panamera"
            elif folder_name == "Culinan":
                folder_name = "Cullinan"       
            elif folder_name == "RAV 4":
                folder_name = "RAV4"                 
            entry_heirarchy += folder_name + "\\"
    
        entry_heirarchy += self.__get_row_name__(row_element)
        return entry_heirarchy

        
    def __get_folder_rows__(self, row_link: str = None) -> tuple[list, list]:
        """
        Indexes the folders and files within a SharePoint directory.
        Filters files based on the selected ADAS systems. If no ADAS systems are selected, processes all files.
        """
    
        # If a link is provided to this method, navigate to it first before indexing
        if row_link is not None:
            self.selenium_driver.get(row_link)
    
        # Store lists for output files/lists
        indexed_files = []
        indexed_folders = []
    
        # Compile regex patterns for selected ADAS systems if any are selected
        adas_patterns = (
            [re.compile(rf"\({re.escape(adas)}\s*\d*\)", re.IGNORECASE) for adas in self.selected_adas]
            if self.selected_adas else None
        )
    
        # Look for the table element. If it doesn't appear in 5 seconds, assume no rows appeared in the folder
        try:
            WebDriverWait(self.selenium_driver, 2.5).until(
                EC.presence_of_element_located((By.XPATH, self.__ONEDRIVE_TABLE_LOCATOR__))
            )
        except:
            # Pull the title of the page to log out that nothing was found inside the current folder
            page_title = self.selenium_driver.find_elements(By.XPATH, self.__ONEDRIVE_PAGE_NAME_LOCATOR__)[-1].get_attribute(
                "innerText"
            ).strip()
            print(f"No folders/files found inside folder {page_title}")
            return [indexed_folders, indexed_files]
    
        # Find all page elements for the lists of files/folders and iterate them one by one
        page_elements = self.selenium_driver.find_elements(By.XPATH, self.__ONEDRIVE_TABLE_LOCATOR__)
        for page_element in page_elements:
            # Wait for the table rows to appear
            WebDriverWait(page_element, 15).until(
                EC.presence_of_element_located((By.XPATH, self.__ONEDRIVE_TABLE_ROW_LOCATOR__))
            )
            page_title = self.selenium_driver.find_elements(By.XPATH, self.__ONEDRIVE_PAGE_NAME_LOCATOR__)[-1].get_attribute(
                "innerText"
            ).strip()
            table_elements = page_element.find_elements(By.XPATH, self.__ONEDRIVE_TABLE_ROW_LOCATOR__)
    
            # Iterate the row elements in the table and decide if they should be included or not based on name/type of entry
            for row_element in table_elements:
                # Pull the name of the folder and the hierarchy of it for use later on
                entry_name = self.__get_row_name__(row_element)
                entry_hierarchy = self.__get_entry_heirarchy__(row_element)
    
                # Filter out entries with old/part/no in their names
                if entry_name.lower().startswith("no"):
                    continue
                if any(value in entry_name.lower() for value in ["old", "part", "replacement", "data"]) and entry_name:
                    continue
    
                # For folders, check if we need to store it as a folder or if the folder is a segmented file set
                if self.__is_row_folder__(row_element):
                    if page_title == self.sharepoint_make and re.search(r"\d{4}", entry_name) is None:
                        continue
    
                    # Pull the link for the folder and check its contents
                    folder_link = self.__get_unencrypted_link__(row_element)
                    if re.search("|".join(self.__DEFINED_MODULE_NAMES__), entry_name) is not None:
                        # Open a new tab and navigate to the folder being checked
                        self.selenium_driver.switch_to.new_window(WindowTypes.TAB)
                        self.selenium_driver.get(folder_link)
    
                        # Find all the child folders/files for the current row entry
                        try:
                            sub_table_element = WebDriverWait(self.selenium_driver, 25).until(
                                EC.presence_of_element_located((By.XPATH, self.__ONEDRIVE_TABLE_LOCATOR__))
                            )
                            sub_table_rows = sub_table_element.find_elements(By.XPATH, self.__ONEDRIVE_TABLE_ROW_LOCATOR__)
                            sub_table_entries = [self.__get_row_name__(sub_table_row) for sub_table_row in sub_table_rows]
                        except:
                            sub_table_entries = []
    
                        # Close the tab for the child folder being indexed
                        self.selenium_driver.close()
                        self.selenium_driver.switch_to.window(self.selenium_driver.window_handles[0])
    
                        # Check if this is a segmented file. If so, store this folder as a file
                        if any(
                            re.search(r"([PpAaRrTt]{4})|(\d+\s{0,}\.[^\s]+)", sub_entry_name)
                            for sub_entry_name in sub_table_entries
                        ):
                            folder_link = self.__get_encrypted_link__(row_element)
                            indexed_files.append(
                                SharepointExtractor.SharepointEntry(
                                    entry_name, entry_hierarchy, folder_link, SharepointExtractor.EntryTypes.FOLDER_ENTRTY
                                )
                            )
                            continue
    
                    # If the folder does not contain a module name, store it as a folder to be indexed later on
                    indexed_folders.append(
                        SharepointExtractor.SharepointEntry(
                            entry_name, entry_hierarchy, folder_link, SharepointExtractor.EntryTypes.FOLDER_ENTRTY
                        )
                    )
                    continue
    
                # For files, either add all files or filter by ADAS systems
                if adas_patterns:
                    # Filter files based on ADAS systems
                    if not any(pattern.search(entry_name) for pattern in adas_patterns):
                        
                        continue
                else:
                    # Log added files when no ADAS filter is applied
                    print(f"File added (no ADAS filter): {entry_name}")
    
                # When we find valid files, get the encrypted link and store it
                file_link = self.__get_encrypted_link__(row_element)
                indexed_files.append(
                    SharepointExtractor.SharepointEntry(
                        entry_name, entry_hierarchy, file_link, SharepointExtractor.EntryTypes.FILE_ENTRY
                    )
                )
    
        # Return the completed list of files and folders
        return [indexed_folders, indexed_files]



       
    def __update_excel_with_whitelist__(self, ws, entry_name, document_url):
        normalized_entry_name = entry_name.replace("(", "").replace(")", "").replace("-", "/").replace("[", "").replace("]", "").replace("WL", "").replace("Multipurpose", "Multipurpose Camera").replace("-PL-PW072NLB", " Side Blind Zone Alert").replace("forward Collision Warning/Lane Departure Warning (FCW/LDW)", "FCW/LDW").strip().upper()
        for row in ws.iter_rows(min_row=2, min_col=3, max_col=3):
            cell_value = str(row[0].value).strip().upper()
    
            # Restrict based on selected ADAS acronyms
            if self.selected_adas and cell_value not in self.selected_adas:
                continue
    
            if cell_value in self.__ADAS_SYSTEMS_WHITELIST__:
                if cell_value in normalized_entry_name:
                    cell = ws.cell(row=row[0].row, column=12)
                    cell.hyperlink = document_url
                    cell.value = document_url
                    cell.font = Font(color="0000FF", underline='single')
                    print(f"Hyperlink for {entry_name} added at {cell.coordinate}")
                    return True
        return False
    
    def __update_excel__(self, ws, year, model, doc_name, document_url, adas_last_row, cell_address=None):
        if self.selected_adas and not any(adas in doc_name.upper() for adas in self.selected_adas):
            return
        # Check if the document name has a specific cell address
        if doc_name in self.SPECIFIC_HYPERLINKS:
            cell = ws[self.SPECIFIC_HYPERLINKS[doc_name]]
        else:
            cell, error_message = self.__find_row_in_excel__(ws, year, self.sharepoint_make, model, doc_name)

        if not cell:
            if cell_address:
                cell = ws[cell_address]
            else:
                row = ws.max_row + 1
                if doc_name in adas_last_row:
                    row = adas_last_row[doc_name] + 1
                else:
                    adas_last_row[doc_name] = row
                cell = ws.cell(row=row, column=12)

            # Add the error message in column H if no matching cell was found
            error_cell = ws.cell(row=cell.row, column=8)
            error_cell.value = error_message or "General Placement Error"
            error_cell.font = Font(color="FF0000")  # Optional: Red color for the error text

        cell.hyperlink = document_url
        cell.value = document_url
        cell.font = Font(color="0000FF", underline='single')
        adas_last_row[doc_name] = cell.row
        print(f"Hyperlink for {doc_name} added at {cell.coordinate}")
        
    def __find_row_in_excel__(self, ws, year, make, model, file_name):
        # Initialize error tracking (no longer needed for error display)
        year_error, make_error, model_error, adas_error = True, True, True, True

        # Extract information from the file name using regex patterns
        extracted_year = re.search(r'\d{4}', file_name)
        extracted_make = self.sharepoint_make
        extracted_model = re.search(r'\b(?:Zevo 600|Other Model Names)\b', file_name)  # Modify the regex to capture your models

        # Extract ADAS systems based on the predefined ADAS system names
        extracted_adas_systems = [adas for adas in self.__DEFINED_MODULE_NAMES__ if adas in file_name.upper()]

        extracted_year = extracted_year.group(0) if extracted_year else "Unknown Year"
        extracted_model = extracted_model.group(0) if extracted_model else model  # Use the model passed if extraction fails
        extracted_adas_systems_str = ", ".join(extracted_adas_systems) if extracted_adas_systems else "Unknown ADAS"

        adas_file_name = file_name.replace(year, "").replace(make, "").replace(model, "").replace("[", "").replace("]", "").replace("WL", "").replace("BSM-RCTW", "BSW-RCTW")
        adas_file_name = adas_file_name.replace(model, "").replace("(", "").replace(")", "").replace("[", "").replace("]", "").replace("WL", "").replace("BSW-RCT W", "BSW-RCTW").replace("BSW-RSTW", "BSW-RCTW").replace("BCW-RCTW", "BSW-RCTW").replace("BSW-RTCW", "BSW-RCTW").replace("BSM-RCTW", "BSW-RCTW").replace("BSW_RCTW", "BSW-RCTW").replace("SCC", "ACC").replace("RR31 Culinan", "Culinan").replace("RR6 Dawn", "Dawn").replace("RR21 Ghost", "Ghost").replace("-PL-PW072NLB", "Side Blind Zone Alert").replace("BSW & RCTW", "BSW-RCTW").replace("-", "/").strip().upper()

        normalization_patterns = [
            (r'(RS)(\d)', r'\1 \2'),
            (r'(SQ)(\d)', r'\1 \2'),
            (r'BSW RCTW', r'BSW/RCTW'),
            (r'BSW-RCT W', r'BSW/RCTW'),
            (r'BSW-RCT W', r'BSW/RCTW')
        ]

        for pattern, replacement in normalization_patterns:
            adas_file_name = re.sub(pattern, replacement, adas_file_name)

        # Iterate through the worksheet rows
        for row in ws.iter_rows(min_row=2, max_col=8):
            year_value = str(row[0].value).strip() if row[0].value is not None else ''
            make_value = str(row[1].value).replace("audi", "Audi").strip() if row[1].value is not None else ''
            model_value = str(row[2].value).replace("RS3", "RS 3").replace("RS5", "RS 5").replace("RS6", "RS 6").replace("RS7", "RS 7").replace("SQ5", "SQ 5").replace("Super Duty F-250", "F-250 SUPER DUTY").replace("Super Duty F-350", "F-350 SUPER DUTY").replace("Super Duty F-450", "F-450 SUPER DUTY").replace("Super Duty F-550", "F-550 SUPER DUTY").replace("Super Duty F-600", "F-600 SUPER DUTY").replace("MACH-E", "Mustang Mach-E ").replace("G Convertable", "G Convertible").replace("Carnival MPV", "Carnival").replace("RANGE ROVER VELAR", "VELAR").replace("RANGE ROVER SPORT", "SPORT").replace("Range Rover Sport", "SPORT").replace("RANGE ROVER EVOQUE", "EVOQUE").replace("MX5", "MX-5").strip() if row[2].value is not None else ''
            adas_value = str(row[4].value).replace("%", "").replace("(", "").replace(")", "").replace("-", "/").replace("SCC 1", "ACC").replace(".pdf", "").strip() if row[4].value is not None else ''

            year_error = year_value.upper() != year.upper()
            make_error = make_value.upper() != make.upper()
            model_error = model_value.upper() != model.upper()
            adas_error = adas_value.upper() not in adas_file_name.upper()

            if year_error or make_error or model_error or adas_error:
                continue

            # If a matching cell is found, add the hyperlink only (don't add the file name to column K)
            for term_index, term in enumerate(self.__ROW_SEARCH_TERMS__):
                if term.upper() in adas_file_name:
                    return ws.cell(row=row[0].row, column=12), None

            return ws.cell(row=row[0].row, column=12), None

        return None, file_name

        ## Throw an exception when we fail to find a row for the current file name given
        # raise Exception(f"ERROR! Failed to find row for file: {file_name}!\nYear: {year}\nMake: {make}\nModel: {model}")           

#####################################################################################################################################################

if __name__ == '__main__':   
    
    # (Individual File testing without GUI, take away the # to perform whichever is needed)) 
    #excel_file_path = r'C:\Users\dromero3\Desktop\Excel Documents\Toyota Pre-Qual Long Sheet v6.3.xlsx'
    #sharepoint_link = 'https://calibercollision.sharepoint.com/:f:/g/enterpriseprojects/VehicleServiceInformation/EiB53aPXartJhkxyWzL5AFABZQsY3x-XDWPXQCqgFIrvoQ?e=m4DrKQ'
    #debug_run = True

    # (Usage with GUI, take away the # to perform whichever is needed)        
    sharepoint_link = sys.argv[1]
    excel_file_path = sys.argv[2]
    debug_run = False

    # Build a new sharepoint extractor with configuration values as defined above
    extractor = SharepointExtractor(sharepoint_link, excel_file_path, debug_run)

    print("="*100)
    extracted_folders, extracted_files = extractor.extract_contents()

    print("="*100)
    extractor.populate_excel_file(extracted_files)

    print("="*100)
    print(f"Extraction and population for {extractor.sharepoint_make} is complete!")