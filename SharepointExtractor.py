# Imports for the SharepointExtractor class object
import os
import re
import time
import openpyxl
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
    __DEFINED_MODULE_NAMES__ = [ 'ACC', 'AEB', 'AHL', 'APA', 'BSW/RCTW', 'BSW-RCTW', 'BUC', 'LKA', 'NV', 'SVC', 'LW' ]

    # Locators used to find objects on the sharepoint folder pages
    __ONEDRIVE_PAGE_NAME_LOCATOR__ = "//li[contains(@data-automationid, 'breadcrumb-listitem')]"
    __ONEDRIVE_TABLE_LOCATOR__ = "//div[@data-automationid='list-pages']/div[contains(@id, 'virtualized-list')]"  
    __ONEDRIVE_TABLE_ROW_LOCATOR__ = "./div[contains(@data-automationid, 'row') and contains(@id, 'virtualized-list')]"
    __ONEDRIVE_TABLE_ROW_COLUMN_LOCATOR__ = "./div[@role='gridcell' and contains(@data-automationid, '$FIELD_NAME')]"

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

    def __init__(self, sharepoint_link: str, excel_file_path: str) -> 'SharepointExtractor':
        """
        CTOR for a new SharepointExtractor. Takes the link to the requested sharepoint location 
        and prepares to extract all file and folder links
        
        ----------------------------------------------
        
        sharepoint_link: str 
            The link to the sharepoint location for the given make
        excel_file_path: str 
            The fully qualified path to the excel file holding our ADAS SI
        """
        
        # Store attributes for the Extractor on this instance
        self.sharepoint_link = sharepoint_link
        self.excel_file_path = excel_file_path

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
        Extracts the file and folder links from the defined sharepoint location for the current extractor object
        Returns a tuple of lists. The first list holds all of our SharepointEntry objects for the folders in the sharepoint,
        and the second list holds all of our SharepointEntry objects for the folders in the sharepoint.
        """
        
        # Index and store base folders and files here then iterate them all
        sharepoint_folders, sharepoint_files = get_folder_results = self.__get_folder_rows__()

        # Iterate the contents of the base folders list as long as it has contents
        start_time = time.time()
        while len(sharepoint_folders) > 0:

            # Store the current folder value and navigate to it for indexing
            folder_link = sharepoint_folders.pop(0).entry_link
            child_folders, child_files = self.__get_folder_rows__(folder_link)
            
            # Add all of our links to the files and folders to our base lists
            for file_entry in child_files: sharepoint_files.append(file_entry)   
            for folder_entry in child_folders: sharepoint_folders.append(folder_entry)
           
            # Log out how many child links and folders exist now
            print(f'{len(sharepoint_folders)} Folders Remain | {len(sharepoint_files)} Files Indexed')

            # BREAK HERE FOR TESTING
            # if len(sharepoint_files) >= 10: break            

        # Build return lists for contents of folders and files        
        elapsed_time = time.time() - start_time
        print(f"Indexing routine took {elapsed_time} to complete")
        return [ sharepoint_folders, sharepoint_files ]
    def populate_excel_file(self, file_entries: list) -> None:
        """
        Populates the excel file for the current make and stores all hyperlinks built in correct 
        locations
        
        file_entries: list[SharepointEntry]
            The list of all file entries we're looking to put into our excel file
        """

        # Load our excel file from the path given 
        start_time = time.time()
        model_workbook = openpyxl.load_workbook(self.excel_file_path)
        model_worksheet = model_workbook['Model Version']  
        print(f"Workbook loaded successfully: {self.excel_file_path}")
 
        # Setup trackers for correct row insertion during population 
        current_model = ""
        adas_last_row = { }
        
        # Iterate all the file entries given and update the excel file accordingly
        for file_entry in file_entries:
            
            # Pull the year and model for the file from the heirarchy
            # Acura\\2015\\RDX\\FileName.ext
            file_name = file_entry.entry_name                           # FileName.ext
            file_model = file_entry.entry_heirarchy.split('\\')[-2]     # RDX
            file_year = file_entry.entry_heirarchy.split('\\')[-3]      # 2015

            # Check if ADAS last row needs to be reset or not
            if file_model != current_model:
                current_model = file_model
                adas_last_row = { }

            # Now update our excel file based on the values given for this entry
            self.__update_excel__(model_worksheet, file_year, file_model, file_name, file_entry.entry_link, adas_last_row, None)

        # Close the workbook once done populating information
        print(f"Saving updated changes to {self.sharepoint_make} sheet now...")
        model_workbook.save(self.excel_file_path)
        model_workbook.close()

        # Log out how long this routine took and exit this method
        elapsed_time = time.time() - start_time
        print(f"Sheet population routine took {elapsed_time} to complete")

    def __generate_chrome_options__(self) -> Options:
        """
        Asks the user if they want to use an existing chrome install or a new one
        Returns a built set of chrome options for the given configuration
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
        
        # Build a new root object for tkinter and ask the user for their choice
        root = tk.Tk()
        root.withdraw()
        use_existing_profile = messagebox.askyesno(
            "Chrome Instance",
            "Would you like to use the currently installed version of Chrome?\n"
            "Click 'Yes' for the currently installed version or 'No' for a new instance.") 
        root.destroy() 

        # Check the user option and update options if needed
        if not use_existing_profile: return chrome_options

        home_dir = os.path.expanduser("~")
        user_data_dir = os.path.join(home_dir, "AppData", "Local", "Google", "Chrome", "User Data")
        chrome_options.add_argument(f"user-data-dir={user_data_dir}")
        profile_dir = "Default" 
        chrome_options.add_argument(f"profile-directory={profile_dir}")
        
        # Return the updated options object
        return chrome_options    
    def __is_row_folder__(self, row_element: WebElement) -> bool:
            
        # Find the icon element and check if it's a folder or file
        icon_element_locator = self.__ONEDRIVE_TABLE_ROW_COLUMN_LOCATOR__.replace("$FIELD_NAME", "field-DocIcon")
        icon_element = WebDriverWait(row_element, self.__MAX_WAIT_TIME__)\
            .until(EC.presence_of_element_located((By.XPATH, icon_element_locator)))
            
        # Return true if this folder is in the name, false if it is not
        return "folder" in icon_element.accessible_name
    def __get_row_name__(self, row_element: WebElement) -> str:
            
        # Find the name column element and return the name for the row in use
        return row_element.get_attribute("aria-label").strip()
    def __get_folder_link__(self, row_element: WebElement) -> str:
            
        # Build and return a new URL for this row entry
        base_url = self.selenium_driver.current_url.replace("&ga=1", "")    # Base URL for the current page
        row_name = self.__get_row_name__(row_element)                       # The name we're looking to open
        row_link = base_url + "%2F" + row_name                              # Relative folder URL based on drive layout

        # Return the built URL here
        return row_link
    def __get_file_link__(self, row_element: WebElement) -> str:

        # Define some local helper functions to perform clipboard operations
        def __get_clipboard_content__() -> str:
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

        # Store a starting clipboard content value to ensure we get a new value during this method
        starting_clipboard_content = __get_clipboard_content__()

        # Find the selector element and try to click it here
        selector_element_locator = self.__ONEDRIVE_TABLE_LOCATOR__.replace("$FIELD_NAME", "row-selection")
        selector_element = WebDriverWait(row_element, self.__MAX_WAIT_TIME__)\
            .until(EC.presence_of_element_located((By.XPATH, selector_element_locator)))            
            
        # Pull the name element from the row and find child buttons for it
        name_element_locator = self.__ONEDRIVE_TABLE_ROW_COLUMN_LOCATOR__.replace("$FIELD_NAME", "field-LinkFilename")
        name_element = row_element.find_element(By.XPATH, name_element_locator) 
        ActionChains(self.selenium_driver).move_to_element_with_offset(name_element, 50, 0).perform()
            
        # Attempt the share routine in a loop to retry when buttons don't appear correctly
        for retry_count in range(3):

            try: 
                
                # Find the share button element and click it here. Setup share settings and copy the link to the clipboard
                name_element.find_element(By.XPATH, ".//button[@data-automationid='shareHeroId']").click()
                time.sleep(0.75)
                ActionChains(self.selenium_driver).send_keys(Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.ENTER).perform()
                time.sleep(0.75)
                ActionChains(self.selenium_driver).send_keys(Keys.ARROW_DOWN, Keys.TAB, Keys.ARROW_DOWN, Keys.TAB, Keys.ENTER).perform()           
                time.sleep(1.00)
                ActionChains(self.selenium_driver).send_keys(Keys.ENTER).perform()  
                time.sleep(0.75)
                ActionChains(self.selenium_driver).send_keys(Keys.ESCAPE).perform()                                     

                # Break this loop if this logic completes correctly
                break

            except:
                                    
                # Check if we can retry or not
                if retry_count == 3:
                    raise Exception("ERROR! Failed to open the share dialog for the current entry!")
                
                # Wait a moment before retrying to open the clipboard 
                time.sleep(1.0)
                
        # Make sure the link value is changed here. If it's not, run this routine again
        encrypted_file_link = __get_clipboard_content__()
        if encrypted_file_link == starting_clipboard_content: 
            return self.__get_file_link__(row_element)

        # Return the stored link from the clipboard
        return encrypted_file_link
    def __get_entry_heirarchy__(self, row_element: WebElement) -> str:
            
        # Find all of our title elements and check for the index of our make. Pull all values after that
        title_elements = self.selenium_driver.find_elements(By.XPATH, self.__ONEDRIVE_PAGE_NAME_LOCATOR__)   
        title_index = title_elements.index(next(title_element for title_element in title_elements if title_element.text == self.sharepoint_make))       
        child_elements = title_elements[title_index:]
        
        # Combine the name of the current folder plus the entry name for our output value
        entry_heirarchy = ""
        for child_element in child_elements: entry_heirarchy += child_element.text.strip() + "\\"
        entry_heirarchy += self.__get_row_name__(row_element)
        
        # Return the built heirarchy name value here
        return entry_heirarchy
    def __get_folder_rows__(self, row_link: str = None) -> tuple[list, list]:
            
        # Navigate to the next link if needed and find the title of the page
        if row_link != None: self.selenium_driver.get(row_link)
            
        # Find the parent table element and find all child rows in it
        table_element = WebDriverWait(self.selenium_driver, self.__MAX_WAIT_TIME__)\
            .until(EC.presence_of_element_located((By.XPATH, self.__ONEDRIVE_TABLE_LOCATOR__)))
        table_elements = table_element.find_elements(By.XPATH, self.__ONEDRIVE_TABLE_ROW_LOCATOR__)

        # Find our page title once the table content has appeard and see if this is a year or model page
        page_title = self.selenium_driver.find_elements(By.XPATH, self.__ONEDRIVE_PAGE_NAME_LOCATOR__)[-1].get_attribute("innerText").strip()
        is_year_folder = re.search("\\d{4}", page_title) != None        

        # Setup lists for the output files and folders
        indexed_files = [ ] 
        indexed_folders = [ ]       

        # Iterate all the rows and get link URLs for each one
        for row_element in table_elements:

            # Pull our row name before testing to filter rows we don't want
            entry_name = self.__get_row_name__(row_element)
            if "no" in entry_name.lower(): continue
            if "old" in entry_name.lower(): continue
                       
            # Check if this is a folder entry or not and make sure the name of the folder is a four digit year
            if not self.__is_row_folder__(row_element):

                # Confirm the file name matches/contains one of the requested modules
                if re.search("|".join(self.__DEFINED_MODULE_NAMES__), entry_name) == None: continue
    
                # Pull the link to our file and store the heirarchy for the entry
                file_link = self.__get_file_link__(row_element)
                file_heirarchy = self.__get_entry_heirarchy__(row_element)
                indexed_files.append(SharepointExtractor.SharepointEntry(entry_name, file_heirarchy, file_link, SharepointExtractor.EntryTypes.FILE_ENTRY))
                continue

            # Before pulling a folder link, make sure it's either a Model or Year folder
            # Some models have a space in them to see if this is a year page or not first
            if not is_year_folder and ' ' in entry_name: continue
            if re.search("\\d{4}|[^ \\n]+", entry_name) == None: continue

            # Store the URL for the row entry on our list and move on
            folder_link = self.__get_folder_link__(row_element)
            folder_heirarchy = self.__get_entry_heirarchy__(row_element)
            indexed_folders.append(SharepointExtractor.SharepointEntry(entry_name, folder_heirarchy, folder_link, SharepointExtractor.EntryTypes.FOLDER_ENTRTY))

        # Return our built list of indexed rows and elements here
        return [indexed_folders, indexed_files]       

    def __update_excel__(self, ws, year, model, doc_name, document_url, adas_last_row, cell_address=None):
        if cell_address:
            cell = ws[cell_address]
        else:
            cell = self.__find_row_in_excel__(ws, year, self.sharepoint_make, model, doc_name)
            if cell:
                row = cell.row
            else:
                row = ws.max_row + 1
            if doc_name in adas_last_row:
                row = adas_last_row[doc_name] + 1
            else:
                adas_last_row[doc_name] = row
            cell = ws.cell(row=row, column=12)

        cell.hyperlink = document_url
        cell.value = document_url
        cell.font = Font(color="0000FF", underline='single')
        adas_last_row[doc_name] = cell.row
        print(f"Hyperlink for {doc_name} added at {cell.coordinate}")
    def __find_row_in_excel__(self, ws, year, make, model, file_name):
        
        # Iterate the rows in the worksheet for the given column range starting at row 2
        for row in ws.iter_rows(min_row=2, max_col=8):

            # Store the properties for the current row and check if this file is for the row
            year_value = str(row[0].value).strip()
            make_value = str(row[1].value).strip()
            model_value = str(row[2].value).strip()
            adas_value = str(row[7].value).strip()          

            # Format the file and module name for correct ADAS module comparison routines
            adas_value = adas_value.replace("(", "").replace(")", "").upper().replace("-", "/").strip()
            adas_file_name = file_name.upper().replace("(", "").replace(")", "").replace("-", "/").strip()

            # Compare our values for the row to the given argument values
            if year_value != year: continue
            if make_value != make: continue
            if model_value != model: continue
            if adas_value not in adas_file_name: continue

            # Return the cell for the current row if all conditions are met
            row_number = row[0].row
            return ws.cell(row=row_number, column=12)

        # Return none if the requested values are not matched
        return None

#####################################################################################################################################################

if __name__ == '__main__':
    """
    Main entry point for a new SharepointExtractor. Takes the link and path to the excel file we need to use 
    for extracting content for the given make
    """
    
    # Build a test extractor and make sure everything can be setup correctly
    # TODO: Change the excel_file_path and sharepoint_link values to be pulled from sys.argv[x]
    excel_file_path = r'C:\Users\DRomero3\OneDrive - Caliber Collision\Acura Pre-Qual Long Sheet v5.44.xlsx'
    sharepoint_link = 'https://calibercollision-my.sharepoint.com/:f:/g/personal/mark_klingenhofer_protechdfw_com/El_B5eO677JOrCJs2XdDenEBfomiRKHT0bPKBAhrmEYCrA?e=URjvLR'
    extractor = SharepointExtractor(sharepoint_link, excel_file_path)

    # Exrtact our contents and return them here
    print("="*100)
    extracted_folders, extracted_files = extractor.extract_contents()   

    # Populate our excel file
    print("="*100)
    extractor.populate_excel_file(extracted_files)

    # Log out extraction is completed and exit this script
    print("="*100)
    print(f"Extraction and population for {extractor.sharepoint_make} is complete!")