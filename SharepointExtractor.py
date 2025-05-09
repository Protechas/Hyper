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
import chromedriver_autoinstaller
from openpyxl.styles import Font
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException, ElementClickInterceptedException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
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
    __DEFINED_MODULE_NAMES__ = [
        'ACC', 'SCC', 'AEB', 'AHL', 'APA', 'BSW', 'BSW/RCTW', 'BSW-RCTW',
        'BSW & RCTW', 'BSW RCTW', 'BSW-RCT W', 'BSW RCT W', 'BSM-RCTW', 'BSW-RTCW', 'BSW_RCTW',
        'BCW-RCTW', 'BUC', 'LKA', 'LW', 'NV', 'SVC', 'WAMC',
    
        # 🔧 Repair SI modules added below
        'YAW', 'G-Force', 'SWS', 'HUD', 'SRS D&E', 'SCI', 'SRR', 'TPMS', 'SBI',
        'EBDE (1)', 'EBDE (2)', 'HDE (1)', 'HDE (2)', 'LGR', 'PSI', 'WRL',
        'PCM', 'TRANS', 'AIR', 'ABS', 'BCM', 'SAS', 'HLI', 'ESC','SRS',
        'KEY', 'FOB', 'HVAC (1)', 'HVAC (2)', 'COOL', 'HEAD (1)', 'HEAD (2)',
        "OCS","OCS2","OCS3","OCS4"
    ]

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
    "2016 Acura RLX (LKA 1) [FCW-LDW].pdf": "L248",
    "2016 Acura RLX (LKA 1) [Multipurpose].pdf": "L249",
    "2012 Volkswagen CC (ACC 1).pdf": "L11",
    "2013 Volkswagen CC (ACC 1).pdf": "L83",
    "2014 Volkswagen CC (ACC 1).pdf": "L155",
    "2015 Volkswagen CC (ACC 1).pdf": "L227",
    "2016 Volkswagen CC (ACC 1).pdf": "L299",
    "2017 Volkswagen CC (ACC 1).pdf": "L371",
    "2022 Kis Stinger (ACC 2)": "L956",
    "2023 Kia Niro EV (ACC 2)": "L1010",
    "2023 Kia Niro EV (AEB 2)": "L1011",
    "2023 Kia Niro EV (APA 1).pdf": "L1013",
    "2023 Kia Niro EV (BSW 1).pdf": "L1014",
    "2023 Kia Niro EV (BUC).pdf": "L1015",
    "2023 Kia Niro EV (LKA 1).pdf": "L1016",
    "2023 Kia Niro HEV (ACC 2)": "L1019",
    "2023 Kia Niro HEV (AEB 2)": "L1020",
     #"2023 Kia Niro HEV (ACC 2).pdf": "L155", Dupliacate entry
    "2023 Kia Niro HEV (APA 1).pdf": "L1022",
    "2023 Kia Niro HEV (BSW 1).pdf": "L1023",
    "2023 Kia Niro HEV (BUC).pdf": "L1024",
    "2023 Kia Niro HEV (LKA 1).pdf": "L1025",
    "2023 Kia Niro PHEV (ACC 2)": "L1028",
    "2023 Kia Niro PHEV (AEB 2)": "L1029",
    "2023 Kia Niro PHEV (APA 1).pdf": "L1031",
    "2023 Kia Niro PHEV (BSW 1).pdf": "L1032",
    "2023 Kia Niro PHEV (BUC).pdf": "L1033",
    "2023 Kia Niro PHEV (LKA 1).pdf": "L1034",
    "2023 Kia Sorento HEV (ACC 2)": "L1064",
    "2023 Kia Sorento HEV (AEB 2)": "L1065",
    "2023 Kia Sorento HEV (BSW 1)": "L1068",
    "2023 Kia Sorento HEV (SVC 1)": "L1072",
    "2023 Kia Sorento HEV (APA 1).pdf": "L1067",
    "2023 Kia Sorento HEV (BUC).pdf": "L1069",
    "2023 Kia Sorento HEV (LKA 1).pdf": "L1070",
    "2023 Kia Sorento PHEV (BSW 1)": "L1077",
    "2023 Kia Sorento PHEV (SVC 1)": "L1081",
    "2023 Kia Sorento PHEV (ACC 2).pdf": "L1073",
    "2023 Kia Sorento PHEV (AEB 2).pdf": "L1074",
    "2023 Kia Sorento PHEV (APA 1).pdf": "L1076",
    "2023 Kia Sorento PHEV (BUC).pdf": "L1078",
    "2023 Kia Sorento PHEV (LKA 1).pdf": "L1079",
    "2023 Kia Sportage (ACC 2)": "L1091",
    "2023 Kia Sportage (AEB 2)": "L1092",
    "2023 Kia Sportage (APA 1)": "L1094",
    "2023 Kia Sportage (SVC 1)": "L1099",
    "2023 Kia Sportage (BSW 1).pdf": "L1095",
    "2023 Kia Sportage (BUC).pdf": "L1096",
    "2023 Kia Sportage (LKA 1).pdf": "L1097",
    "2023 Kia Sportage NG5 (AEB 2)": "L1101", #misspealt, Its NQ5, not NG5
    "2023 Kia Sportage NQ5 (ACC 2)": "L1100",
    "2023 Kia Sportage NQ5 (APA 1)": "L1103",
    "2023 Kia Sportage NQ5 (BSW 1).pdf": "L1104",
    "2023 Kia Sportage NQ5 (BUC).pdf": "L1105",
    "2023 Kia Sportage NQ5 (LKA 1).pdf": "L1106",
    "2023 Kia Sportage NQ5 (SVC 1).pdf": "L1108",
    "2023 Kia Sportage PHEV (ACC 2)": "L1118",
    "2023 Kia Sportage PHEV (AEB 2)": "L1119",
    "2023 Kia Sportage PHEV (SVC 1)": "L1126",
    "2023 Kia Sportage PHEV (APA 1).pdf": "L1121",
    "2023 Kia Sportage PHEV (BSW 1).pdf": "L1122",
    "2023 Kia Sportage PHEV (BUC).pdf": "L123",
    "2023 Kia Sportage PHEV (LKA 1).pdf": "L1124", 
    "2024 Kia Carnival (ACC 2).pdf": "L1145", # MPV?????
    "2024 Kia Carnival (AEB 2).pdf": "L1146",
    "2024 Kia Carnival (APA 1).pdf": "L1148",
    "2024 Kia Carnival (BSW 1).pdf": "L1149",
    "2024 Kia Carnival (BUC).pdf": "L1150",
    "2024 Kia Carnival (LKA 1).pdf": "L1151",
    "2024 Kia Carnival (SVC 1).pdf": "L1153"
   
    # Add more mappings as needed
    }
    HYPERLINK_COLUMN_INDEX = 12  # Default is Column L (can change to 11 for K, etc.)

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

        self.mode = sys.argv[4] if len(sys.argv) > 4 else "adas"
        self.repair_mode = self.mode == "repair"
        self.selected_adas = sys.argv[3].split(",") if len(sys.argv) > 3 else []
        
        # Set correct column index
        self.HYPERLINK_COLUMN_INDEX = 8 if self.repair_mode else 12  # K for repair, L for ADAS
        
        
        # Store attributes for the Extractor on this instance
        self.__DEBUG_RUN__ = debug_run
        self.sharepoint_link = sharepoint_link
        self.excel_file_path = excel_file_path
        self.selected_adas = sys.argv[3].split(",") if len(sys.argv) > 3 else []

        # Define the default wait timeout and setup a new selenium driver
        # This will download (if needed) and add the correct chromedriver to PATH
        chromedriver_autoinstaller.install()
               
        # Then just start Chrome normally:
        self.selenium_driver = webdriver.Chrome(
            options=self.__generate_chrome_options__()
        )
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

        # ─── PAUSE TO LET THE FOLDER LIST SETTLE ───────────────────────────
        time.sleep(2.0)
        
        # Index and store base folders and files here
        sharepoint_folders, sharepoint_files = self.__get_folder_rows__()
    
        # Compile regex patterns for the selected ADAS systems if any are selected
        if self.repair_mode:
            adas_patterns = [re.compile(re.escape(rs), re.IGNORECASE) for rs in self.selected_adas] if self.selected_adas else None
        else:
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
    
            # ─── PAUSE TO LET THE FOLDER LIST SETTLE ───────────────────────────
            time.sleep(3.0)

            # Add child folders for further processing
            sharepoint_folders.extend(child_folders)
    
            # Add all files if no ADAS systems are selected, otherwise filter them
            if self.repair_mode and self.selected_adas:
                for file_entry in child_files:
                    entry_name = file_entry.entry_name
            
                    # Try to extract module from parentheses first
                    module_match = re.search(r'\((.*?)\)', entry_name)
                    if module_match:
                        file_module = module_match.group(1).strip().upper()
                    else:
                        # Fallback: try last word before .pdf
                        name_without_ext = os.path.splitext(entry_name)[0]
                        file_module = name_without_ext.split()[-1].strip().upper()
            
                    if file_module in [s.upper() for s in self.selected_adas]:
                        filtered_files.append(file_entry)
                    else:
                        print(f"Skipping {entry_name} — '{file_module}' not in selected: {self.selected_adas}")
            
            elif adas_patterns:
                for file_entry in child_files:
                    if any(pattern.search(file_entry.entry_name) for pattern in adas_patterns):
                        filtered_files.append(file_entry)
            
            else:
                filtered_files.extend(child_files)

    
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
        sheet_name = 'Model Version'
        if sheet_name not in model_workbook.sheetnames:
            print(f"WARNING: Sheet '{sheet_name}' not found. Defaulting to first sheet.")
            model_worksheet = model_workbook.active
            self.row_index = self.__build_row_index__(model_worksheet, self.repair_mode)
        else:
            model_worksheet = model_workbook[sheet_name]

        print(f"Workbook loaded successfully: {self.excel_file_path}")
    
        # Setup trackers for correct row insertion during population
        current_model = ""
        adas_last_row = {}
        self.row_index = self.__build_row_index__(model_worksheet, self.repair_mode)

        # Iterate through the filtered file entries
        for file_entry in file_entries:
            print(f"Processing file: {file_entry.entry_name}")
            file_name = file_entry.entry_name
        
            # === Year Extraction ===
            year_match = re.search(r'(20\d{2})', file_name)
            file_year = year_match.group(1) if year_match else "Unknown"
        
            # === Model Extraction ===
            base_name = re.sub(r'(20\d{2})', '', file_name)
            base_name = base_name.replace(".pdf", "").strip()
            base_name = re.sub(re.escape(self.sharepoint_make), "", base_name, flags=re.IGNORECASE).strip()
        
            model_tokens = []
            for token in base_name.split():
                # Stop if token starts with "(" or is a known module
                if token.startswith("(") or token.upper().strip("()[]") in self.__DEFINED_MODULE_NAMES__:
                    break
                model_tokens.append(token)

        
            file_model = " ".join(model_tokens).strip() if model_tokens else "Unknown"
        
            # === ✅ Fallback for Model from Hierarchy ===
            if file_model == "Unknown":
                segments = file_entry.entry_heirarchy.split("\\")
                if len(segments) > 1:
                    file_model = segments[-2]  # Usually the model folder
        
            # Check if ADAS last row needs to be reset or not
            if file_model != current_model:
                current_model = file_model
                adas_last_row = {}
        
            # Proceed with placing the hyperlink
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
        Configures Chrome to use a valid custom user profile.
        Chrome requires --user-data-dir to be a non-default location for automation.
        """
    
        chrome_options = Options()
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-extensions")
        chrome_options.add_argument("--disable-infobars")
        chrome_options.add_argument("--disable-browser-side-navigation")
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    
        # Define a dedicated automation user data directory
        home_dir = os.path.expanduser("~")
        automation_profile_base = os.path.join(home_dir, "ChromeAutomationProfiles")
    
        # Try Default first, then Profile 1
        profile_name = "Default"
        if not os.path.exists(os.path.join(home_dir, "AppData", "Local", "Google", "Chrome", "User Data", "Default")):
            profile_name = "Profile 1"
    
        # Copy the real profile to a custom folder if not already copied
        original_profile = os.path.join(home_dir, "AppData", "Local", "Google", "Chrome", "User Data", profile_name)
        target_profile = os.path.join(automation_profile_base, profile_name)
    
        if not os.path.exists(target_profile):
            import shutil
            shutil.copytree(original_profile, target_profile)
    
        chrome_options.add_argument(f"--user-data-dir={automation_profile_base}")
        chrome_options.add_argument(f"--profile-directory={profile_name}")
    
        print(f"[Chrome Profile] Using copied profile: {profile_name}")
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
        selector_locator = ".//div[@role='gridcell' and contains(@data-automationid, 'row-selection-')]"
        selector_element = WebDriverWait(row_element, self.__MAX_WAIT_TIME__)\
            .until(EC.presence_of_element_located((By.XPATH, selector_locator)))

        # ─── BEGIN CLICK WITH FALLBACK ───
        # 1) scroll into view
        self.selenium_driver.execute_script("arguments[0].scrollIntoView(true);", selector_element)
        # 2) try normal click, else JS click
        try:
            selector_element.click()
        except ElementClickInterceptedException:
            # if something (like the Share iframe) is covering it, JS-click bypasses it
            self.selenium_driver.execute_script("arguments[0].click();", selector_element)
        # ─── END CLICK WITH FALLBACK ─
        time.sleep(1.00)
        
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
            #elif folder_name == "SQ3":
             #   folder_name = "SQ 3"
            #elif folder_name == "SQ5":
             #   folder_name = "SQ 5"
            #elif folder_name == "SQ7":
             #   folder_name = "SQ 7"
            #elif folder_name == "SQ8":
             #   folder_name = "SQ 8"
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
        if row_link is not None:
            self.selenium_driver.get(row_link)
    
        indexed_files = []
        indexed_folders = []
    
        # Compile ADAS/Repair‐mode regex patterns
        if self.repair_mode and self.selected_adas:
            adas_patterns = [re.compile(re.escape(rs), re.IGNORECASE) for rs in self.selected_adas]
        elif self.selected_adas:
            adas_patterns = [
                re.compile(rf"\({re.escape(adas)}\s*\d*\)", re.IGNORECASE)
                for adas in self.selected_adas
            ]
        else:
            adas_patterns = None
    
        # ─── ROBUST WAIT FOR FOLDER ROWS ───
        # 1) wait for the table container to appear
        WebDriverWait(self.selenium_driver, self.__MAX_WAIT_TIME__).until(
            EC.presence_of_element_located((By.XPATH, self.__ONEDRIVE_TABLE_LOCATOR__))
        )
        # 2) wait for any loading spinner to vanish
        try:
            WebDriverWait(self.selenium_driver, 5).until(
                EC.invisibility_of_element_located((By.CSS_SELECTOR, ".loading-spinner"))
            )
        except TimeoutException:
            pass
    
        # 3) grab all table containers on the page
        page_elements = self.selenium_driver.find_elements(By.XPATH, self.__ONEDRIVE_TABLE_LOCATOR__)
    
        for page_element in page_elements:
            # Poll the number of rows until it stabilizes
            prev_count = -1
            stable = 0
            rows = []
            while stable < 3:
                rows = page_element.find_elements(By.XPATH, self.__ONEDRIVE_TABLE_ROW_LOCATOR__)
                count = len(rows)
                if count == prev_count:
                    stable += 1
                else:
                    stable = 0
                prev_count = count
                time.sleep(0.5)
    
            if not rows:
                print("No table rows found in folder; skipping...")
                continue
    
            # Get the folder title from the page header
            page_title = (
                self.selenium_driver
                .find_elements(By.XPATH, self.__ONEDRIVE_PAGE_NAME_LOCATOR__)[-1]
                .get_attribute("innerText")
                .strip()
            )
    
            # Now iterate the fully−loaded rows
            for row_element in rows:
                entry_name = self.__get_row_name__(row_element)
                entry_hierarchy = self.__get_entry_heirarchy__(row_element)
    
                if entry_name.lower().startswith("no"):
                    continue
                # skip unwanted terms
                ignore_terms = ["old", "part", "replacement", "data", "statement", "stament"]
                if any(term in entry_name.lower() for term in ignore_terms):
                    continue
    
                if self.__is_row_folder__(row_element):
                    if page_title == self.sharepoint_make and not re.search(r"\d{4}", entry_name):
                        continue
    
                    folder_link = self.__get_unencrypted_link__(row_element)
                    # special deep-folder check
                    if re.search("|".join(self.__DEFINED_MODULE_NAMES__), entry_name):
                        self.selenium_driver.switch_to.new_window(WindowTypes.TAB)
                        self.selenium_driver.get(folder_link)
                        try:
                            sub_table = WebDriverWait(self.selenium_driver, 25).until(
                                EC.presence_of_element_located((By.XPATH, self.__ONEDRIVE_TABLE_LOCATOR__))
                            )
                            sub_rows = sub_table.find_elements(By.XPATH, self.__ONEDRIVE_TABLE_ROW_LOCATOR__)
                            sub_names = [self.__get_row_name__(r) for r in sub_rows]
                        except:
                            sub_names = []
                        self.selenium_driver.close()
                        self.selenium_driver.switch_to.window(self.selenium_driver.window_handles[0])
    
                        if any(re.search(r"([PpAaRrTt]{4})|(\d+\s*\.[^\s]+)", name) for name in sub_names):
                            folder_link = self.__get_encrypted_link__(row_element)
                            indexed_files.append(
                                SharepointExtractor.SharepointEntry(
                                    entry_name,
                                    entry_hierarchy,
                                    folder_link,
                                    SharepointExtractor.EntryTypes.FOLDER_ENTRTY
                                )
                            )
                            continue
    
                    indexed_folders.append(
                        SharepointExtractor.SharepointEntry(
                            entry_name,
                            entry_hierarchy,
                            folder_link,
                            SharepointExtractor.EntryTypes.FOLDER_ENTRTY
                        )
                    )
                    continue
    
                # === 🔍 FILTERING STARTS HERE ===
                if self.selected_adas:
                    if self.repair_mode:
                        module_match = re.search(r'\((.*?)\)', entry_name)
                        if module_match:
                            file_module = module_match.group(1).strip().upper()
                        else:
                            file_module = os.path.splitext(entry_name)[0].split()[-1].upper()
    
                        if file_module not in [s.upper() for s in self.selected_adas]:
                            print(f"Skipping {entry_name} — '{file_module}' not in selected: {self.selected_adas}")
                            continue
                    else:
                        if not any(p.search(entry_name) for p in adas_patterns):
                            continue
                # === 🔍 FILTERING ENDS HERE ===
    
                file_link = self.__get_encrypted_link__(row_element)
                indexed_files.append(
                    SharepointExtractor.SharepointEntry(
                        entry_name,
                        entry_hierarchy,
                        file_link,
                        SharepointExtractor.EntryTypes.FILE_ENTRY
                    )
                )
    
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
                    cell = ws.cell(row=row[0].row, column=self.HYPERLINK_COLUMN_INDEX)
                    cell.hyperlink = document_url
                    cell.value = document_url
                    cell.font = Font(color="0000FF", underline='single')
                    print(f"Hyperlink for {entry_name} added at {cell.coordinate}")
                    return True
        return False
    
    def __update_excel__(self, ws, year, model, doc_name, document_url, adas_last_row, cell_address=None):
        # Skip filtering if in Repair mode
        if not self.repair_mode:
            if self.selected_adas and not any(adas in doc_name.upper() for adas in self.selected_adas):
                return
    
        # Try to find the correct Excel row for this system
        if doc_name in self.SPECIFIC_HYPERLINKS:
            cell = ws[self.SPECIFIC_HYPERLINKS[doc_name]]
            error_message = None
        else:
           cell, error_message = self.__find_row_in_excel__(
               ws, year, self.sharepoint_make, model, doc_name,
               repair_mode=self.repair_mode, row_index=self.row_index
           )

    
        # Create a unique key for tracking the row (includes system/module name now for Repair SI)
        if self.repair_mode:
            system_match = re.search(r"\((.*?)\)", doc_name)
            system_name = system_match.group(1).upper() if system_match else doc_name.split()[-1].upper()
            key = (year, self.sharepoint_make, model, system_name)
        else:
            key = (year, self.sharepoint_make, model, doc_name)
    
        if not cell:
            if cell_address:
                cell = ws[cell_address]
            else:
                row = ws.max_row + 1
                if key in adas_last_row:
                    row = adas_last_row[key] + 1
                else:
                    adas_last_row[key] = row
                cell = ws.cell(row=row, column=self.HYPERLINK_COLUMN_INDEX)
        
            # Place error info in the correct column: K (11) for ADAS, G (7) for Repair
            error_column = 11 if not self.repair_mode else 7
            error_cell = ws.cell(row=cell.row, column=error_column)
            error_cell.value = doc_name
            error_cell.font = Font(color="FF0000")

    
        cell.hyperlink = document_url
        cell.value = document_url
        cell.font = Font(color="0000FF", underline='single')
        adas_last_row[key] = cell.row  # Store row used to prevent duplicates
    
        print(f"Hyperlink for {doc_name} added at {cell.coordinate}")
        
    def __find_row_in_excel__(self, ws, year, make, model, file_name, repair_mode=False, row_index=None):
        def normalize_system_name(name):
            return re.sub(r"[^A-Z0-9]", "", name.upper()) if name else ''
    
        # Extract from file name
        extracted_year = re.search(r'\d{4}', file_name)
        extracted_make = self.sharepoint_make
        extracted_model = re.search(r'\b(?:Zevo 600|Other Model Names)\b', file_name)  # Modify as needed
    
        extracted_adas_systems = [adas for adas in self.__DEFINED_MODULE_NAMES__ if adas in file_name.upper()]
        extracted_year = extracted_year.group(0) if extracted_year else "Unknown Year"
        extracted_model = extracted_model.group(0) if extracted_model else model
        adas_file_name = file_name.replace(year, "").replace(make, "").replace(model, "")
        adas_file_name = re.sub(r"[\[\]()\-]", "", adas_file_name).replace("WL", "").replace("BSM-RCTW", "BSW-RCTW").strip().upper()
    
        normalization_patterns = [
            (r'(RS)(\d)', r'\1 \2'),
            (r'(SQ)(\d)', r'\1 \2'),
            (r'BSW RCTW', r'BSW/RCTW'),
            (r'BSW-RCT W', r'BSW/RCTW')
        ]
        for pattern, replacement in normalization_patterns:
            adas_file_name = re.sub(pattern, replacement, adas_file_name)
    
        # ⬇️ REPAIR MODE LOGIC
        if repair_mode:
            # Extract system name from file name
            system_match = re.search(r"\((.*?)\)", file_name)
            system_name = system_match.group(1).strip().upper() if system_match else file_name.split()[-1].strip().upper()
            normalized_system = re.sub(r"[^A-Z0-9]", "", system_name).upper().strip()
    
            key = (
                year.strip().upper(),
                make.strip().upper(),
                model.strip().upper(),
                normalized_system
            )
            
            # Debug output for validation
            #if row_index:
               # print(f"[DEBUG] Looking for key: {key}")
                #if key not in row_index:
                   # print(f"[DEBUG] Key not found in index.")
                    #print(f"[DEBUG] Available keys (sample): {list(row_index.keys())[:5]}")
            
                
            if row_index and key in row_index:
                row_num = row_index[key]
                return ws.cell(row=row_num, column=self.HYPERLINK_COLUMN_INDEX), None
            return None, file_name
    
        # ⬇️ ADAS LOGIC
        for row in ws.iter_rows(min_row=2, max_col=8):
            year_value = str(row[0].value).strip() if row[0].value is not None else ''
            make_value = str(row[1].value).replace("audi", "Audi").strip() if row[1].value is not None else ''
            model_value = str(row[2].value).replace("Super Duty F-250", "F-250 SUPER DUTY") \
                .replace("Super Duty F-350", "F-350 SUPER DUTY").replace("Super Duty F-450", "F-450 SUPER DUTY") \
                .replace("Super Duty F-550", "F-550 SUPER DUTY").replace("Super Duty F-600", "F-600 SUPER DUTY") \
                .replace("MACH-E", "Mustang Mach-E ").replace("G Convertable", "G Convertible") \
                .replace("Carnival MPV", "Carnival").replace("RANGE ROVER VELAR", "VELAR") \
                .replace("RANGE ROVER SPORT", "SPORT").replace("Range Rover Sport", "SPORT") \
                .replace("RANGE ROVER EVOQUE", "EVOQUE").replace("MX5", "MX-5").strip() if row[2].value is not None else ''
    
            adas_value = str(row[4].value).replace("%", "").replace("(", "").replace(")", "").replace("-", "/") \
                .replace("SCC 1", "ACC").replace(".pdf", "").strip() if row[4].value is not None else ''
    
            year_error = year_value.strip().upper() != year.strip().upper()
            make_error = make_value.strip().upper() != make.strip().upper()
            model_error = model_value.strip().upper() != model.strip().upper()
            adas_error = adas_value.strip().upper() not in adas_file_name.upper()
    
            if year_error or make_error or model_error or adas_error:
                continue
    
            for term in self.__ROW_SEARCH_TERMS__:
                if term.upper() in adas_file_name:
                    return ws.cell(row=row[0].row, column=self.HYPERLINK_COLUMN_INDEX), None
    
            return ws.cell(row=row[0].row, column=self.HYPERLINK_COLUMN_INDEX), None
    
        return None, file_name

       
    def __build_row_index__(self, ws, repair_mode=False):
        index = {}
        for row in ws.iter_rows(min_row=2, max_col=8):
            year = str(row[0].value).strip().upper() if row[0].value else ''
            make = str(row[1].value).strip().upper() if row[1].value else ''
            model = str(row[2].value).strip().upper() if row[2].value else ''
            
            if repair_mode:
                system = str(row[3].value).strip().upper() if row[3].value else ''
            else:
                system = str(row[4].value).strip().upper() if row[4].value else ''
    
            normalized_system = re.sub(r"[^A-Z0-9]", "", system)
            key = (year, make, model, normalized_system)
            index[key] = row[0].row
        return index
      

#####################################################################################################################################################

if __name__ == '__main__':   
    
    # (Individual File testing without GUI, take away the # to perform whichever is needed)) 
    #excel_file_path = r'C:\Users\dromero3\Desktop\Excel Documents\Toyota Pre-Qual Long Sheet v6.3.xlsx'
    #sharepoint_link = 'https://calibercollision.sharepoint.com/:f:/g/enterpriseprojects/VehicleServiceInformation/EiB53aPXartJhkxyWzL5AFABZQsY3x-XDWPXQCqgFIrvoQ?e=m4DrKQ'
    #debug_run = True

    # (Usage with GUI, take away the # to perform whichever is needed)        
    sharepoint_link = sys.argv[1]
    excel_file_path = sys.argv[2]
    debug_run = True

    # Build a new sharepoint extractor with configuration values as defined above
    extractor = SharepointExtractor(sharepoint_link, excel_file_path, debug_run)

    print("="*100)
    extracted_folders, extracted_files = extractor.extract_contents()

    print("="*100)
    extractor.populate_excel_file(extracted_files)

    print("="*100)
    print(f"Extraction and population for {extractor.sharepoint_make} is complete!")