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
    __ONEDRIVE_PAGE_NAME_LOCATOR__ = "//li//div[contains(@class, 'ms-TooltipHost') and @role='none']/div[@hidden]"
    __ONEDRIVE_TABLE_LOCATOR__ = ".//div[@role='presentation']/div[contains(@class, 'ms-List-page')]"  
    __ONEDRIVE_TABLE_ROW_LOCATOR__ = "./div[contains(@class, 'ms-List-cell') and contains(@role, 'presentation') and @data-list-index]"

    # Collections of system names used for finding correct files and row locations
    __DEFINED_MODULE_NAMES__ = [ 'ACC', 'AEB', 'AHL', 'APA', 'BSW/RCTW', 'BSW-RCTW','BSW RCTW','BSW-RCT W','BSW RCT W', 'BUC', 'LKA', 'NV', 'SVC' ]
    __ROW_SEARCH_TERMS__ = ['LKAS', 'FCW/LDW', 'Multipurpose', 'Cross Traffic Alert', 'Lane Change Alert', 'Side Blind Zone Alert', 'Surround Vision Camera', 'Video Processing', 'Pending Further Research',]
    __ADAS_SYSTEMS_WHITELIST__ = [
        'FCW/LDW',
        'FCW-LDW',
        'Multipurpose Camera',
        'Multipurpose',
        'Forward Collision Warning/Lane Departure Warning (FCW/LDW)',
        'Cross Traffic Alert',
        'Surround Vision Camera',
        'Video Processing'
    ]

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
        sharepoint_folders, sharepoint_files = self.__get_folder_rows__()

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
            # if len(sharepoint_files) >= 200: break            
        
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
            # Acura\\2014\\MDX\\2014 Acura MDX (LKA 1)\\FileName.ext
            file_name = file_entry.entry_name                                                   
            file_model = file_entry.entry_heirarchy.split('\\')[2]
            file_year = file_entry.entry_heirarchy.split('\\')[1]
            
            # Check if ADAS last row needs to be reset or not
            if file_model != current_model:
                current_model = file_model
                adas_last_row = { }

            # Now update our excel file based on the values given for this entry
            if self.__update_excel_with_whitelist__(model_worksheet, file_name, file_entry.entry_link): continue
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
        icon_element_locator = ".//div[contains(@class, 'fileTypeIconColumn')]//i"
        icon_element = WebDriverWait(row_element, self.__MAX_WAIT_TIME__)\
            .until(EC.presence_of_element_located((By.XPATH, icon_element_locator)))
            
        # Return true if this folder is in the name, false if it is not
        icon_class = icon_element.get_attribute("class")
        return "Folder" in icon_class    
    def __get_row_name__(self, row_element: WebElement) -> str:
            
        # Find the name column element and return the name for the row in use
        row_name_locator = ".//button[@data-automationid='FieldRenderer-name']"
        row_name_element = row_element.find_element(By.XPATH, row_name_locator)
        return row_name_element.text.strip()         
    def __get_unencrypted_link__(self, row_element: WebElement) -> str:
        
        # Pull the folder name and add the name of it to our URL name
        base_url = self.selenium_driver.current_url.split("&p=true")[0]     # Current URL Split up for the path of the current folder
        row_name = self.__get_row_name__(row_element)                       # The name we're looking to open
        plain_link = base_url + "%2F" + row_name                            # Relative folder URL based on drive layout

        # Return the built URL here
        return plain_link    
    def __get_encrypted_link__(self, row_element: WebElement) -> str:
              
        # Debug run testing break out to speed things up
        if (self.__DEBUG_RUN__): return f"Link For: {self.__get_row_name__(row_element)}"
        
        # Store a starting clipboard content value to ensure we get a new value during this method
        starting_clipboard_content = self.__get_clipboard_content__()

        # Find the selector element and try to click it here
        selector_element_locator = ".//div[@data-selection-toggle='true']"
        selector_element = WebDriverWait(row_element, self.__MAX_WAIT_TIME__)\
            .until(EC.presence_of_element_located((By.XPATH, selector_element_locator)))            
        selector_element.click()    
        
        # Attempt the share routine in a loop to retry when buttons don't appear correctly
        for retry_count in range(3):

            try: 
                
                # Find the share button element and click it here. Setup share settings and copy the link to the clipboard
                row_element.find_element(By.XPATH, ".//button[@data-automationid='FieldRender-ShareHero']").click()
                time.sleep(1.00)
                ActionChains(self.selenium_driver).send_keys(Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.ENTER).perform()
                time.sleep(1.25)
                ActionChains(self.selenium_driver).send_keys(Keys.ARROW_DOWN, Keys.TAB, Keys.ARROW_DOWN, Keys.TAB, Keys.TAB, Keys.ENTER, Keys.TAB, Keys.ENTER).perform()           
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
                
                # Wait a moment before retrying to open the clipboard 
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
        # Find all of our title elements and check for the index of our make. Pull all values after that
        title_elements = self.selenium_driver.find_elements(By.XPATH, self.__ONEDRIVE_PAGE_NAME_LOCATOR__)   
        title_index = title_elements.index(next(title_element for title_element in title_elements if title_element.get_attribute("innerText") == self.sharepoint_make))       
        child_elements = title_elements[title_index:]
    
        # Combine the name of the current folder plus the entry name for our output value
        entry_heirarchy = ""
        for child_element in child_elements:
            folder_name = child_element.get_attribute("innerText").strip()
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
            entry_heirarchy += folder_name + "\\"
    
        entry_heirarchy += self.__get_row_name__(row_element)
    
        # Return the built heirarchy name value here
        return entry_heirarchy
    
    def __get_folder_rows__(self, row_link: str = None) -> tuple[list, list]:

        # If a link is provided to this method, navigate to it first before indexing
        if row_link is not None:
            self.selenium_driver.get(row_link)

        # Store lists for output files/lists
        indexed_files = []
        indexed_folders = []      
        
        # Look for the table element. If it doesn't appear in 5 seconds, assume no rows appeared in the folder           
        try: WebDriverWait(self.selenium_driver, 2.5).until(EC.presence_of_element_located((By.XPATH, self.__ONEDRIVE_TABLE_LOCATOR__)))           
        except:            

            # Pull the title of the page to log out that nothing was found inside the current folder
            page_title = self.selenium_driver.find_elements(By.XPATH, self.__ONEDRIVE_PAGE_NAME_LOCATOR__)[-1].get_attribute("innerText").strip()
            print(f"No folders/files found inside folder {page_title}")
            return [indexed_files, indexed_folders]           
              
        # Find all page elements for the lists of files/folders and iterate them one by one
        page_elements = self.selenium_driver.find_elements(By.XPATH, self.__ONEDRIVE_TABLE_LOCATOR__)
        for page_element in page_elements:
            
            # Wait for the table rows to appear. If they don't appear, this should fail out since you always have rows when you have a table
            WebDriverWait(page_element, 5).until(EC.presence_of_element_located((By.XPATH, self.__ONEDRIVE_TABLE_ROW_LOCATOR__)))
            page_title = self.selenium_driver.find_elements(By.XPATH, self.__ONEDRIVE_PAGE_NAME_LOCATOR__)[-1].get_attribute("innerText").strip()
            table_elements = page_element.find_elements(By.XPATH, self.__ONEDRIVE_TABLE_ROW_LOCATOR__)

            # Iterate the row elements in the table and decide if they should be included or not based on name/type of entry
            for row_element in table_elements:

                # Pull the name of the folder and the heirarchy of it for use later on
                entry_name = self.__get_row_name__(row_element)
                entry_heirarchy = self.__get_entry_heirarchy__(row_element)
                       
                # Filter out entries with old/part/no in their names
                if entry_name.lower().startswith("no"):
                    continue
                if any(value in entry_name.lower() for value in ["old", "part"]) and entry_name:
                    continue
                
                # For folders, check if we need to store it as a folder of if the folder is a segmented file set
                if self.__is_row_folder__(row_element):
                    if page_title == self.sharepoint_make and re.search("\\d{4}", entry_name) is None:
                        continue

                    # Pull the link for the folder and open it in a new tab to peek at the contents of it
                    folder_link = self.__get_unencrypted_link__(row_element)
                    if re.search("|".join(self.__DEFINED_MODULE_NAMES__), entry_name) is not None:

                        # Open a new tab and navigate to the folder being checked
                        self.selenium_driver.switch_to.new_window(WindowTypes.TAB)
                        self.selenium_driver.get(folder_link)

                        # Find all the child folders/files for the current row entry
                        sub_table_element = WebDriverWait(self.selenium_driver, self.__MAX_WAIT_TIME__)\
                            .until(EC.presence_of_element_located((By.XPATH, self.__ONEDRIVE_TABLE_LOCATOR__)))
                        sub_table_rows = sub_table_element.find_elements(By.XPATH, self.__ONEDRIVE_TABLE_ROW_LOCATOR__)
                        sub_table_entries = [self.__get_row_name__(sub_table_row) for sub_table_row in sub_table_rows]

                        # Close the tab for the child folder being indexed
                        self.selenium_driver.close()
                        self.selenium_driver.switch_to.window(self.selenium_driver.window_handles[0])

                        # Check if this is a segmented file. If so, store this folder as a file
                        if any(re.search(r"([PpAaRrTt]{4})|(\d+\s{0,}\.[^\s]+)", sub_entry_name) for sub_entry_name in sub_table_entries):
                            folder_link = self.__get_encrypted_link__(row_element)
                            indexed_files.append(SharepointExtractor.SharepointEntry(entry_name, entry_heirarchy, folder_link, SharepointExtractor.EntryTypes.FOLDER_ENTRTY))
                            continue

                    # If the folder does not contain a module name, store it as a folder to be indexed later on
                    indexed_folders.append(SharepointExtractor.SharepointEntry(entry_name, entry_heirarchy, folder_link, SharepointExtractor.EntryTypes.FOLDER_ENTRTY))
                    continue

                # If the row is not a folder, check if it has a module name or not.
                if re.search("|".join(self.__DEFINED_MODULE_NAMES__), entry_name) is None:
                    continue

                # When we find valid module names, get the encrypted link for the folder and store it.
                file_link = self.__get_encrypted_link__(row_element)
                indexed_files.append(SharepointExtractor.SharepointEntry(entry_name, entry_heirarchy, file_link, SharepointExtractor.EntryTypes.FILE_ENTRY))
    
        # Return the completed list of files and folders
        return [indexed_folders, indexed_files]
       
    def __update_excel_with_whitelist__(self, ws, entry_name, document_url):
        normalized_entry_name = entry_name.replace("(", "").replace(")", "").replace("-", "/").replace("Multipurpose", "Multipurpose Camera").replace("forward Collision Warning/Lane Departure Warning (FCW/LDW)", "FCW/LDW").strip().upper()
        for row in ws.iter_rows(min_row=2, min_col=3, max_col=3):
            cell_value = str(row[0].value).strip().upper()
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

        cell = self.__find_row_in_excel__(ws, year, self.sharepoint_make, model, doc_name)

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

        cell.hyperlink = document_url
        cell.value = document_url
        cell.font = Font(color="0000FF", underline='single')
        adas_last_row[doc_name] = cell.row
        print(f"Hyperlink for {doc_name} added at {cell.coordinate}")  
        
    def __find_row_in_excel__(self, ws, year, make, model, file_name):
        
        # Remove the year make and model from the file name provided
        adas_file_name = file_name.replace(year, "").replace(make, "").replace(model, "")
        adas_file_name = adas_file_name.replace(model, "").replace("(", "").replace(")", "").replace("BSW-RCT W", "BSW-RCTW").replace("-", "/").strip().upper()

        # Apply specific normalization rules
        normalization_patterns = [
            (r'(RS)(\d)', r'\1 \2'),
            (r'(SQ)(\d)', r'\1 \2'),
            (r'BSW RCTW', r'BSW/RCTW'),
            (r'BSW-RCT W', r'BSW/RCTW'),
            (r'BSW RCT W', r'BSW/RCTW')
        ]
    
        for pattern, replacement in normalization_patterns:
            adas_file_name = re.sub(pattern, replacement, adas_file_name)

        for row in ws.iter_rows(min_row=2, max_col=8):
            year_value = str(row[0].value).strip() if row[0].value is not None else ''
            make_value = str(row[1].value).replace("audi", "Audi").strip() if row[1].value is not None else ''
            model_value = str(row[2].value).replace("RS3", "RS 3").replace("RS5", "RS 5").replace("RS6", "RS 6").replace("RS7", "RS 7").replace("SQ5", "SQ 5").strip() if row[2].value is not None else ''
            adas_value = str(row[4].value).replace("%", "").replace("-", "/").strip() if row[4].value is not None else ''

            if year_value.upper() != year.upper(): continue
            if make_value.upper() != make.upper(): continue
            if model_value.upper() != model.upper(): continue
            if adas_value.upper() not in adas_file_name: continue

            # Check the ADAS value and ensure proper placement
            for term_index, term in enumerate(self.__ROW_SEARCH_TERMS__):
                if term.upper() in adas_file_name:
                    return ws.cell(row=row[0].row + term_index, column=12)

            return ws.cell(row=row[0].row, column=12)      

        # Throw an exception when we fail to find a row for the current file name given
        # raise Exception(f"ERROR! Failed to find row for file: {file_name}!\nYear: {year}\nMake: {make}\nModel: {model}")           

#####################################################################################################################################################

if __name__ == '__main__':   
    
    # These values will be pulled from the call made by Hyper to boot this scripts
    excel_file_path = r'C:\Users\dromero3\Desktop\Excel Documents\Chrysler Pre-Qual Long Sheet v6.3.xlsx'
    sharepoint_link = 'https://calibercollision.sharepoint.com/:f:/g/enterpriseprojects/VehicleServiceInformation/Em3tiVe1A0pGuxCWxhsh6koBFqVQnpk3SgGCDMtPWlmdwQ?e=AUqPjM'
    debug_run = True

    # Build a new sharepoint extractor with configuration values as defined above
    extractor = SharepointExtractor(sharepoint_link, excel_file_path, debug_run)

    print("="*100)
    extracted_folders, extracted_files = extractor.extract_contents()   

    print("="*100)
    extractor.populate_excel_file(extracted_files)

    print("="*100)
    print(f"Extraction and population for {extractor.sharepoint_make} is complete!")