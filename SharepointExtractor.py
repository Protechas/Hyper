﻿import os
import re
import sys
import time
import shutil
import winreg
import platform
import openpyxl
import subprocess
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
import urllib.parse
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
        'BCW-RCTW', 'BUC', 'LKA', 'LW', 'NV', 'SVC', 'WAMC', 'FRS', 'PDS', 'RRS', 'WSC', 
    
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
    # ADAS SI Manual Placeholders
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
    "2024 Kia Carnival (SVC 1).pdf": "L1153",
    
    #Repair SI Manual Placeholders
    "2020 Kia Niro EV (DE PHEV)(G-Force)":    "H2842",
    "2020 Kia Niro EV (DE PHEV)(SAS)":        "H2840",
    "2020 Kia Niro EV (DE PHEV)(YAW)":        "H2841",
    "2020 Kia Niro EV (DE PHEV)(SWS).pdf":    "H2843",
    
    #2021 Kia Niro (Excel) and Below dosent match, but placing links anyway (it is named completely wrong in the Excel as a normal model should never reflect the EV version)
    "2021 Kia Niro EV (DE PHEV)(G-Force)":    "H2413",
    "2021 Kia Niro EV (DE PHEV)(SAS)":        "H2411",
    "2021 Kia Niro EV (DE PHEV)(YAW)":        "H2412",
    "2021 KIA Niro EV (DE PHEV)(ESC).pdf":    "H2418",
    "2021 KIA Niro EV (DE PHEV)(HLI).pdf":    "H2422",
    "2021 KIA Niro EV (DE PHEV)(SCI).pdf":    "H2420",
    "2021 KIA Niro EV (DE PHEV)(SRR).pdf":    "H2421",
    "2021 KIA Niro EV (DE PHEV)(SRS D&E).pdf":"H2419",
    "2021 Kia Niro EV (DE PHEV)(SWS).pdf":    "H2414",
    "2021 KIA Niro EV (DE PHEV)(TPMS).pdf":   "H2423",

    #2022 Kia Niro (Excel) and Below dosent match, but placing links anyway (it is named completely wrong in the Excel as a normal model should never reflect the EV version)
    "2022 Kia Niro EV (G-Force)":             "H2050",
    "2022 Kia Niro EV (SAS)":                 "H2048",
    "2022 Kia Niro EV (SRR)":                 "H2058",
    "2022 Kia Niro EV (YAW)":                 "H2049",
    "2022 Kia Niro EV (ESC).pdf":             "H2055",
    "2022 Kia Niro EV (HLI).pdf":             "H2059",
    "2022 Kia Niro EV (SCI).pdf":             "H2057",
    "2022 Kia Niro EV (SRS D&E).pdf":         "H2056",
    "2022 Kia Niro EV (SWS).pdf":             "H2051",
    "2022 Kia Niro EV (TPMS).pdf":            "H2060",

    "2023 Kia Niro (EV)(G-Force)":            "H1423",
    "2023 Kia Niro (EV)(SAS)":                "H1421",
    "2023 Kia Niro (EV)(YAW)":                "H1422",
    "2023 Kia Niro (EV) (SWS).pdf":           "H1424",
    "2023 Kia Niro (EV)(ESC).pdf":            "H1428",
    "2023 Kia Niro (EV)(HLI).pdf":            "H1432",
    "2023 Kia Niro (EV)(SCI).pdf":            "H1430",
    "2023 Kia Niro (EV)(SRR).pdf":            "H1431",
    "2023 Kia Niro (EV)(SRS D&E).pdf":        "H1429",
    "2023 Kia Niro (EV)(TPMS).pdf":           "H1433",

    "2023 Kia Niro (HEV)(G-Force)":           "H1456",
    "2023 Kia Niro (HEV)(YAW)":               "H1455",
    "2023 Kia Niro (HEV) (SAS).pdf":          "H1454",
    "2023 Kia Niro (HEV) (SWS).pdf":          "H1457",
    "2023 Kia Niro (HEV)(ESC).pdf":           "H1461",
    "2023 Kia Niro (HEV)(HLI).pdf":           "H1465",
    "2023 Kia Niro (HEV)(SCI).pdf":           "H1463",
    "2023 Kia Niro (HEV)(SRR).pdf":           "H1464",
    "2023 Kia Niro (HEV)(SRS D&E).pdf":       "H1462",
    "2023 Kia Niro (HEV)(TPMS).pdf":          "H1466",

    "2023 Kia Niro (PHEV)(G-Force)":          "H1489",
    "2023 Kia Niro (PHEV)(YAW)":              "H1488",
    "2023 Kia Niro (PHEV) (SAS).pdf":         "H1487",
    "2023 Kia Niro (PHEV) (SWS).pdf":         "H1490",
    "2023 Kia Niro (PHEV)(ESC).pdf":          "H1494",
    "2023 Kia Niro (PHEV)(HLI).pdf":          "H1498",
    "2023 Kia Niro (PHEV)(SCI).pdf":          "H1496",
    "2023 Kia Niro (PHEV)(SRR).pdf":          "H1497",
    "2023 Kia Niro (PHEV)(SRS D&E).pdf":      "H1495",
    "2023 Kia Niro (PHEV)(TPMS).pdf":         "H1499",

    "2023 Kia Sorento (HEV)(ESC)":            "H1626",
    "2023 Kia Sorento (HEV)(G-Force)":        "H1621",
    "2023 Kia Sorento (HEV)(SAS)":            "H1619",
    "2023 Kia Sorento (HEV)(SRR)":            "H1629",
    "2023 Kia Sorento (HEV)(YAW)":            "H1620",
    "2023 Kia Sorento (HEV)(HLI).pdf":        "H1630",
    "2023 Kia Sorento (HEV)(SCI).pdf":        "H1628",
    "2023 Kia Sorento (HEV)(SRS D&E).pdf":    "H1627",
    "2023 Kia Sorento (HEV)(SWS).pdf":        "H1622",
    "2023 Kia Sorento (HEV)(TPMS).pdf":       "H1631",

    "2023 Kia Sorento (PHEV)(SAS)":           "H1652",
    "2023 Kia Sorento PHEV (ESC)":            "H1659",
    "2023 Kia Sorento PHEV (SRR)":            "H1662",
    "2023 Kia Sorento (PHEV)(G-Force).pdf":   "H1654",
    "2023 Kia Sorento (PHEV)(HLI).pdf":       "H1663",
    "2023 Kia Sorento (PHEV)(SCI).pdf":       "H1661",
    "2023 Kia Sorento (PHEV)(SRS D&E).pdf":   "H1660",
    "2023 Kia Sorento (PHEV)(SWS).pdf":       "H1655",
    "2023 Kia Sorento (PHEV)(TPMS).pdf":      "H1664",
    "2023 Kia Sorento (PHEV)(YAW).pdf":       "H1653",

    "2023 Kia Sportage (HEV)(G-Force)":       "H1720",
    "2023 Kia Sportage (HEV)(YAW)":           "H1719",
    "2023 Kia Sportage HEV (SRR)":            "H1728",
    "2023 Kia Sportage (HEV)(ESC).pdf":       "H1725",
    "2023 Kia Sportage (HEV)(HLI).pdf":       "H1729",
    "2023 Kia Sportage (HEV)(SAS).pdf":       "H1718",
    "2023 Kia Sportage (HEV)(SCI).pdf":       "H1727",
    "2023 Kia Sportage (HEV)(SRS D&E).pdf":   "H1726",
    "2023 Kia Sportage (HEV)(SWS).pdf":       "H1721",
    "2023 Kia Sportage (HEV)(TPMS).pdf":      "H1730",

    "2023 Kia Sportage (NQ5)(ESC)":           "H1758",
    "2023 Kia Sportage (NQ5)(G-Force)":       "H1753",
    "2023 Kia Sportage (NQ5)(SAS)":           "H1751",
    "2023 Kia Sportage (NQ5)(SRR)":           "H1761",
    "2023 Kia Sportage (NQ5)(YAW)":           "H1752",
    "2023 Kia Sportage (NQ5)(HLI).pdf":       "H1762",
    "2023 Kia Sportage (NQ5)(SCI).pdf":       "H1760",
    "2023 Kia Sportage (NQ5)(SRS D&E).pdf":   "H1759",
    "2023 Kia Sportage (NQ5)(SWS).pdf":       "H1754",
    "2023 Kia Sportage (NQ5)(TPMS).pdf":      "H1763",

    "2023 Kia Sportage (NQ5A)(ESC)":          "H1791",
    "2023 Kia Sportage (NQ5A)(G-Force)":      "H1786",
    "2023 Kia Sportage (NQ5A)(SRR)":          "H1794",
    "2023 Kia Sportage (NQ5A)(YAW)":          "H1785",
    "2023 Kia Sportage (NQ5A)(HLI).pdf":      "H1795",
    "2023 Kia Sportage (NQ5A)(SAS).pdf":      "H1784",
    "2023 Kia Sportage (NQ5A)(SCI).pdf":      "H1793",
    "2023 Kia Sportage (NQ5A)(SRS D&E).pdf":  "H1792",
    "2023 Kia Sportage (NQ5A)(SWS).pdf":      "H1787",
    "2023 Kia Sportage (NQ5A)(TPMS).pdf":     "H1796",

    "2023 Kia Sportage (PHEV)(SRR)":          "H1827",
    "2023 Kia Sportage (PHEV)(ESC).pdf":      "H1824",
    "2023 Kia Sportage (PHEV)(G-Force).pdf":  "H1819",
    "2023 Kia Sportage (PHEV)(HLI).pdf":      "H1828",
    "2023 Kia Sportage (PHEV)(SAS).pdf":      "H1817",
    "2023 Kia Sportage (PHEV)(SCI).pdf":      "H1826",
    "2023 Kia Sportage (PHEV)(SRS D&E).pdf":  "H1825",
    "2023 Kia Sportage (PHEV)(SWS).pdf":      "H1820",
    "2023 Kia Sportage (PHEV)(TPMS).pdf":     "H1829",
    "2023 Kia Sportage (PHEV)(YAW).pdf":      "H1818"
   
    # Add more mappings as needed
    }

    REPAIR_SYNONYMS = {
        # Core driver-assist sensors
        "Steering Angle Sensor":           "SAS",
        "Yaw Rate Sensor":                 "YAW",
        "G Force Sensor":                  "G-Force",
        "Seat Weight Sensor":              "SWS",
        "Adaptive Head Lamps":             "AHL",
        "Night Vision":                    "NV",
        "Heads Up Display":                "HUD",
        "Electronic Stability Control Relearn": "ESC",
        "Airbag Disengagement/Engagement": "SRS D&E",
        "Steering Column Inspection":      "SCI",
        "Steering Rack Relearn":           "SRR",
        "Headlamp Initialization":         "HLI",
        "Tire Pressure Monitor Relearn":   "TPMS",
        "Seat Belt Inspection":            "SBI",

        # Electric/hybrid battery systems
        "Battery Disengagement":           "EBDE (1)",
        "Battery Engagement":              "EBDE (2)",
        "Hybrid Disengagement":            "HDE (1)",
        "Hybrid Engagement":               "HDE (2)",

        # Miscellaneous vehicle-level relearns
        "Liftgate Relearn":                "LGR",
        "Power Seat Initialization":       "PSI",
        "Window Relearn":                  "WRL",

        # Module programming routines
        "Powertrain Control Module Program":       "PCM",
        "Transmission Control Module Program":     "TRANS",
        "Airbag Control Module Program":           "AIR",
        "Antilock Brake Control Module Program":   "ABS",
        "Body Control Module Program":             "BCM",
        "Key Program":                             "KEY",
        "Key FOB Relearn":                         "FOB",

        # HVAC & coolant
        "Heating, Air Conditioning, Ventilation EVAC":    "HVAC (1)",
        "Heating, Air Conditioning, Ventilation Recharge":"HVAC (2)",
        "Coolant Services":                               "COOL",

        # Headset resets
        "Headset Reset (Spring Style)":       "HEAD (1)",
        "Headset Reset (Squib Style)":        "HEAD (2)",
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
        self.cleanup_mode = sys.argv[5] == "cleanup" if len(sys.argv) > 5 else False
        self.excel_mode = sys.argv[6] if len(sys.argv) > 6 else "og"
        self.broken_entries = []  # ← Store broken hyperlinks here for cleanup mode
        
        # Set correct column index
        # New mode: Column K (11), OG mode: Column L (12), or Repair: K (11)
        if self.repair_mode:
            self.HYPERLINK_COLUMN_INDEX = 8  # Column H (standard for repair mode)
        elif self.excel_mode == "new":
            self.HYPERLINK_COLUMN_INDEX = 11  # Column K
        else:
            self.HYPERLINK_COLUMN_INDEX = 12  # Column L (OG)
        
        
        
        # Store attributes for the Extractor on this instance
        self.__DEBUG_RUN__ = debug_run
        self.sharepoint_link = sharepoint_link
        self.excel_file_path = excel_file_path
        self.selected_adas = sys.argv[3].split(",") if len(sys.argv) > 3 else []

        # Check installed Chrome version
        def get_chrome_version():
            try:
                output = subprocess.check_output(
                    r'reg query "HKEY_CURRENT_USER\Software\Google\Chrome\BLBeacon" /v version',
                    shell=True, stderr=subprocess.DEVNULL, text=True
                )
                match = re.search(r"version\s+REG_SZ\s+([^\s]+)", output)
                return match.group(1) if match else None
            except Exception:
                return None
        
        def get_local_chromedriver_version():
            base = os.path.dirname(chromedriver_autoinstaller.__file__)
            if not os.path.exists(base):
                return None
            for folder in os.listdir(base):
                if folder.isdigit():
                    return folder
            return None
        
        chrome_version = get_chrome_version()
        driver_version = get_local_chromedriver_version()
        
        if chrome_version and driver_version:
            if not chrome_version.startswith(driver_version):
                mismatch_path = os.path.join(os.path.dirname(chromedriver_autoinstaller.__file__), driver_version)
                if os.path.exists(mismatch_path):
                    print(f"Deleting mismatched ChromeDriver v{driver_version} for Chrome v{chrome_version}")
                    shutil.rmtree(mismatch_path)
        
        # Always install the correct version
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
        
        if self.sharepoint_make.lower() == "toyota" and self.repair_mode:
           self.HYPERLINK_COLUMN_INDEX = 10  # Excel column J

    def extract_contents(self) -> tuple[list, list]:
        """
        Extracts the file and folder links from the defined sharepoint location for the current extractor object.
        Returns a tuple of lists. The first list holds all of our SharepointEntry objects for the folders in the sharepoint,
        and the second list holds all of our SharepointEntry objects for the files in the sharepoint.
        """
    
        time.sleep(2.0)
    
        if self.cleanup_mode:
            print("🔍 Clean up Mode: Navigating per broken link...")
    
            matched_files = []
    
            for _, (yr, mk, mdl, sys) in self.broken_entries:
                # Reverse map if Excel gave us a normalized string like "GFORCE"
                for desc, acronym in self.REPAIR_SYNONYMS.items():
                    normalized = acronym.replace(" ", "").replace("&", "").replace("-", "").upper()
                    if normalized == sys.strip().upper():
                        sys = acronym
                        break
            
                print(f"🔎 Seeking: {yr} ➝ {mdl} ➝ {sys}")
    
                # STEP 1: reset to root folder
                self.selenium_driver.get(self.sharepoint_link)
                time.sleep(2.0)
    
                # STEP 2: find year folder
                year_folders, _ = self.__get_folder_rows__()
                target_year = next((f for f in year_folders if yr.strip() == f.entry_name.strip()), None)
                if not target_year:
                    print(f"❌ Year folder '{yr}' not found.")
                    continue
                self.selenium_driver.get(target_year.entry_link)
                time.sleep(1.5)
    
                # STEP 3: find model folder
                model_folders, _ = self.__get_folder_rows__()
                target_model = next((f for f in model_folders if mdl.strip().upper() == f.entry_name.strip().upper()), None)
                if not target_model:
                    print(f"❌ Model folder '{mdl}' not found under year '{yr}'.")
                    continue
                self.selenium_driver.get(target_model.entry_link)
                time.sleep(1.5)
    
                # STEP 4: look for the file matching the system
                try:
                    # Locate SharePoint table rows directly from current model folder page
                    table = WebDriverWait(self.selenium_driver, 15).until(
                        EC.presence_of_element_located((By.XPATH, self.__ONEDRIVE_TABLE_LOCATOR__))
                    )
                    rows = table.find_elements(By.XPATH, self.__ONEDRIVE_TABLE_ROW_LOCATOR__)
                
                    for row in rows:
                        name = self.__get_row_name__(row)
                    
                        found_match = False
                    
                        if self.repair_mode:
                            # REUSE YOUR EXACT MATCH LOGIC FROM __get_folder_rows__
                            module_matches = re.findall(r'\((.*?)\)', name)
                            if module_matches:
                                for module in module_matches:
                                    if module.strip().upper() == sys.strip().upper():
                                        found_match = True
                                        break
                            else:
                                # fallback: look at last word
                                base = os.path.splitext(name)[0]
                                last = base.split()[-1].strip().upper()
                                if last == sys.strip().upper():
                                    found_match = True
                        else:
                            # ADAS match style
                            match_pattern = re.compile(rf"\({re.escape(sys)}\s*\d*\)", re.IGNORECASE)
                            if match_pattern.search(name):
                                found_match = True
                    
                        if not found_match:
                            continue

                
                        # Also check year
                        year_match = re.search(r'(20\d{2})', name)
                        file_year = year_match.group(1) if year_match else ""
                        if file_year.strip() != yr.strip():
                            continue
                
                        # Also check model (soft match)
                        clean_name = re.sub(r'\(.*?\)', '', name)
                        clean_name = re.sub(r'(20\d{2})', '', clean_name).replace(".pdf", "").strip().upper()
                        if mdl.strip().upper() not in clean_name:
                            continue
                
                        # ✅ MATCHED! Get encrypted link and add to matched_files
                        encrypted_link = self.__get_encrypted_link__(row)
                        matched_files.append(
                            SharepointExtractor.SharepointEntry(
                                name=name,
                                heirarchy=self.__get_entry_heirarchy__(row),
                                link=encrypted_link,
                                type=SharepointExtractor.EntryTypes.FILE_ENTRY
                            )
                        )
                        
                        print(f"✅ Direct match: {name}")
                        break
                    else:
                        print(f"❌ No direct match in visible rows: {mdl}, {sys}")
                
                except Exception as e:
                    print(f"⚠️ Failed to locate system file row: {e}")

    
            return [[], matched_files]

  
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
            
                    # — extract ALL parenthetical modules —
                    module_matches = re.findall(r'\((.*?)\)', entry_name)
                    found_match = False
            
                    if module_matches:
                        for module in module_matches:
                            module = module.strip().upper()
                            if module in [s.upper() for s in self.selected_adas]:
                                found_match = True
                                break
                    else:
                        # fallback: last word before extension
                        base = os.path.splitext(entry_name)[0]
                        last = base.split()[-1].strip().upper()
                        if last in [s.upper() for s in self.selected_adas]:
                            found_match = True
            
                    if found_match:
                        filtered_files.append(file_entry)
                    else:
                        print(f"Skipping {entry_name} — no matching system in {self.selected_adas}")


            
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

    def __simulate_entry_from_no_entry__(self, entry_name: str, real_link: str, heirarchy: str, sibling_files: list) -> 'SharepointExtractor.SharepointEntry':
        """
        Replaces 'No XYZ...' file with a simulated one using year/make/model from known good entries.
        """
        # Try to extract acronym (inside parentheses or inferred from known acronyms)
        acronym_match = re.search(r'\((.*?)\)', entry_name)
        if acronym_match:
            acronym = acronym_match.group(1)
        else:
            acronym = next((key for key in self.__DEFINED_MODULE_NAMES__ if key in entry_name.upper()), None)
        if not acronym:
            return None
    
        # Look for a similar file to simulate from
        for sibling in sibling_files:
            sibling_name = sibling.entry_name
            year_match = re.search(r'(20\d{2})', sibling_name)
            if not year_match:
                continue
            year = year_match.group(1)
    
            # Strip down to extract model like original logic
            base_name = re.sub(r'(20\d{2})', '', sibling_name)
            base_name = base_name.replace(".pdf", "").strip()
            base_name = re.sub(re.escape(self.sharepoint_make), "", base_name, flags=re.IGNORECASE).strip()
    
            tokens = []
            mod_names = {m.upper() for m in self.__DEFINED_MODULE_NAMES__}
            for token in base_name.split():
                if token.startswith("("):
                    content = token.strip("()")
                    if content.upper() in mod_names:
                        break
                    tokens.append(content)
                elif token.upper().strip("()[]") in mod_names:
                    break
                else:
                    tokens.append(token)
    
            model = " ".join(tokens)
            if model:
                new_name = f"{year} {self.sharepoint_make} {model} ({acronym})"
                return SharepointExtractor.SharepointEntry(
                    name=new_name,
                    heirarchy=heirarchy,
                    link=real_link,
                    type=SharepointExtractor.EntryTypes.FILE_ENTRY
                )
    
        return None


    def populate_excel_file(self, file_entries: list) -> None:
        """
        Populates the excel file for the current make and stores all hyperlinks built in correct 
        locations.
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
    
        # ── NEW: Detect cleanup mode and initialize list ──
        cleanup_mode = (len(sys.argv) > 5 and sys.argv[5] == "cleanup")
        self.cleanup_mode = cleanup_mode
        if cleanup_mode:
            self.broken_entries = []
    
        # Setup trackers for correct row insertion during population
        current_model = ""
        adas_last_row = {}
        self.row_index = self.__build_row_index__(model_worksheet, self.repair_mode)
    
        # ── Clean up Mode: Detect and clear broken links ──
        if self.cleanup_mode:
            print("🧹 Clean up Mode: Scanning for broken hyperlinks...")
    
            for key, row in self.row_index.items():
                cell = model_worksheet.cell(row=row, column=self.HYPERLINK_COLUMN_INDEX)
                url = str(cell.value).strip() if cell.value else None
                if not url:
                    continue
    
                # Ignore links that contain "part" in the name
                if "part" in url.lower():
                    print(f"⏩ Skipping 'part' link at row {row}: {url}")
                    continue
    
                if self.is_broken_sharepoint_link(url):
                    # Show which entry is being queued for fix
                    yr, mk, mdl, system_name = key
                    print(f"🔧 Broken link found → Year: {yr}, Make: {mk}, Model: {mdl}, System: {system_name}")
    
                    # Clear bad link
                    cell.value = None
                    cell.hyperlink = None
    
                    self.broken_entries.append((row, key))
    
            # ── NEW: announce total broken hyperlinks to fix ──
            total_broken = len(self.broken_entries)
            print(f"Total broken hyperlinks: {total_broken}")
    
            # ── After checking for broken links, reload SharePoint homepage ──
            print("🔄 Re-loading SharePoint root page to resume indexing...")
            self.selenium_driver.get(self.sharepoint_link)
            time.sleep(2.0)
    
    
        # Iterate through the filtered file entries
        # Iterate through the filtered file entries
        for file_entry in file_entries:
            print(f"Processing file: {file_entry.entry_name}")
            file_name = file_entry.entry_name
    
            # … your existing RENAMING logic …
            for desc, acr in self.REPAIR_SYNONYMS.items():
                pattern = f"({desc})"
                if pattern in file_name:
                    file_name = file_name.replace(pattern, f"({acr})")
                    file_entry.entry_name = file_name
                    break
    
            # === Year Extraction ===
            year_match = re.search(r'(20\d{2})', file_name)
            file_year = year_match.group(1) if year_match else "Unknown"
    
            # === Model Extraction ===
            base_name = re.sub(r'(20\d{2})', '', file_name)
            base_name = base_name.replace(".pdf", "").strip()
            base_name = re.sub(re.escape(self.sharepoint_make), "", base_name, flags=re.IGNORECASE).strip()
    
            model_tokens = []
            mod_names = {m.upper() for m in self.__DEFINED_MODULE_NAMES__}
    
            for token in base_name.split():
                if token.startswith("("):
                    content = token.strip("()")
                    if content.strip().upper() in mod_names:
                        break
                    else:
                        model_tokens.append(content)
                elif token.upper().strip("()[]") in mod_names:
                    break
                else:
                    model_tokens.append(token)
    
            file_model = " ".join(model_tokens).strip() if model_tokens else "Unknown"
    
            # ✅ Fallback for Model from Hierarchy
            if file_model == "Unknown":
                segments = file_entry.entry_heirarchy.split("\\")
                if len(segments) > 1:
                    file_model = segments[-2]
    
            # Reset model‐row tracker
            if file_model != current_model:
                current_model = file_model
                adas_last_row = {}
    
            # Place hyperlink
            if self.__update_excel_with_whitelist__(model_worksheet, file_name, file_entry.entry_link):
                if self.cleanup_mode:
                    print(f"Fixed hyperlink for: {file_entry.entry_name}")
                continue
    
            # **Now file_year and file_model are defined, no squiggles**
            self.__update_excel__(
                model_worksheet,
                file_year,
                file_model,
                file_name,
                file_entry.entry_link,
                adas_last_row,
                None
            )
    
            if self.cleanup_mode:
                print(f"Fixed hyperlink for: {file_entry.entry_name}")
    
    
        # Save the workbook
        print(f"Saving updated changes to {self.sharepoint_make} sheet now...")
        model_workbook.save(self.excel_file_path)
        model_workbook.close()
    
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
        # grab just the real name
        name = self.__get_row_name__(row_element).splitlines()[0]
        # if it’s got a .pdf/.docx/.xlsx/.pptx *anywhere* in that name, it’s a file
        if re.search(r'\.(pdf|docx?|xlsx?|pptx?)\b', name, re.IGNORECASE):
            return False
        return True

     
    
    def __get_row_name__(self, row_element: WebElement) -> str:
        """
        Read only the *first line* of the aria-label or innerText,
        so we never pull in date/author on subsequent prints.
        """
        raw = row_element.get_attribute("aria-label")
        if raw and raw.strip():
            # only the actual name, before any newline/date/author lines
            return raw.strip().splitlines()[0]

        # Fallback to the visible text, again only first line
        text = row_element.text.strip()
        return text.splitlines()[0]

   

    def __get_unencrypted_link__(self, row_element: WebElement) -> str:
        """
        Build the “AllItems.aspx?id=…” URL for this row, but:
          • Only use the first line of the aria-label (the real name)
          • Strip *exactly* one trailing “%2F” if present (no more)
          • Preserve everything after the first '&' (viewid, ga, noAuthRedirect…)
        """
        # 1) grab only the real name (no date/author)
        full_label = self.__get_row_name__(row_element)        # e.g. "2012\nJune 15…"
        item_name  = full_label.splitlines()[0]                # “2012”

        # 2) split off base & query
        current = self.selenium_driver.current_url
        try:
            base_url, query = current.split('?', 1)
        except ValueError:
            raise RuntimeError(f"Unexpected URL format: {current!r}")

        # 3) split into the id= piece and the rest
        id_part, rest = query.split('&', 1)
        key, old_val = id_part.split('=', 1)

        # 4) remove at most one trailing "%2F"
        if old_val.endswith('%2F'):
            base_id = old_val[:-3]
        else:
            base_id = old_val

        # 5) URL-encode only the item_name, tack it on
        encoded_name = urllib.parse.quote(item_name, safe='')
        new_val      = f"{base_id}%2F{encoded_name}"

        # 6) rebuild
        return f"{base_url}?{key}={new_val}&{rest}"

      
        
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
                    
    def is_broken_sharepoint_link(self, url: str) -> bool:
        try:
            original_tab = self.selenium_driver.current_window_handle
    
            # Open new tab and switch to it
            self.selenium_driver.execute_script("window.open('');")
            WebDriverWait(self.selenium_driver, 5).until(lambda d: len(d.window_handles) > 1)
    
            new_tabs = [h for h in self.selenium_driver.window_handles if h != original_tab]
            if not new_tabs:
                print("⚠️ New tab did not open.")
                return True
    
            new_tab = new_tabs[0]
            self.selenium_driver.switch_to.window(new_tab)
    
            # Load URL in new tab
            self.selenium_driver.get(url)
    
            WebDriverWait(self.selenium_driver, 10).until(
                EC.presence_of_element_located((By.TAG_NAME, "body"))
            )
    
            # Check for SharePoint error
            error_element = self.selenium_driver.find_elements(
                By.ID, "ctl00_PlaceHolderPageTitleInTitleArea_ErrorPageTitlePanel"
            )
            if error_element:
                if "something went wrong" in error_element[0].text.lower():
                    return True
    
            body_text = self.selenium_driver.find_element(By.TAG_NAME, "body").text.lower()
            if "sorry, something went wrong" in body_text:
                return True

            # 🔍 Check that the PDF viewer loaded the filename span
            try:
                WebDriverWait(self.selenium_driver, 5).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "span.fui-Text"))
                )
            except:
                print("❌ File viewer didn't load filename — treating as broken.")
                return True
    
            return False
    
        except Exception as e:
            print(f"⚠️ Error checking link: {url} → {e}")
            return True
    
        finally:
            # Close ONLY if on the new tab and it’s not the only tab open
            if self.selenium_driver.current_window_handle != original_tab:
                if len(self.selenium_driver.window_handles) > 1:
                    self.selenium_driver.close()
                    self.selenium_driver.switch_to.window(original_tab)


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
    
                    # — SPECIAL: if Repair mode & SAS is selected, grab the SAS folder itself —
                if self.repair_mode and 'SAS' in [s.upper() for s in self.selected_adas] \
                   and entry_name.strip().upper() == 'SAS':
                    # row_link is the folder URL we came in on
                    folder_url = row_link or self.selenium_driver.current_url
                    indexed_files.append(
                        SharepointExtractor.SharepointEntry(
                            entry_name,
                            entry_hierarchy,
                            folder_url,
                            SharepointExtractor.EntryTypes.FILE_ENTRY
                        )
                    )
                    continue

                # Special handling for "No ..." entries
                if entry_name.lower().startswith("no"):
                    simulated_entry = self.__simulate_entry_from_no_entry__(
                        entry_name,
                        self.__get_encrypted_link__(row_element),   # Get real SharePoint link
                        self.__get_entry_heirarchy__(row_element),
                        indexed_files
                    )
                    if simulated_entry:
                        indexed_files.append(simulated_entry)
                    continue  # Do not add the original "No ..." item
                                
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
                        module_matches = re.findall(r'\((.*?)\)', entry_name)
                        found_match = False
                    
                        if module_matches:
                            for module in module_matches:
                                module = module.strip().upper()
                                if module in [s.upper() for s in self.selected_adas]:
                                    found_match = True
                                    break
                        else:
                            name_without_ext = os.path.splitext(entry_name)[0]
                            last_word = name_without_ext.split()[-1].strip().upper()
                            if last_word in [s.upper() for s in self.selected_adas]:
                                found_match = True
                    
                        if not found_match:
                            print(f"Skipping {entry_name} — No matching system found in {self.selected_adas}")
                            continue

                    else:
                        if not any(p.search(entry_name) for p in adas_patterns):
                            continue
                # === 🔍 FILTERING ENDS HERE ===
    
                    # — SPECIAL: if file mentions MDPS, hyperlink the parent folder instead —
                if not self.__is_row_folder__(row_element) \
                   and any(phrase in entry_name.upper() for phrase in ['EXCEPT MDPS', 'MDPS ONLY']):
                    folder_url = row_link or self.selenium_driver.current_url
                    indexed_files.append(
                        SharepointExtractor.SharepointEntry(
                            entry_name,
                            entry_hierarchy,
                            folder_url,
                            SharepointExtractor.EntryTypes.FILE_ENTRY
                        )
                    )
                    continue


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
            module_matches = re.findall(r'\((.*?)\)', doc_name)
            system_name = None
            # pick the first module that’s in selected_adas
            for mod in module_matches:
                if mod.strip().upper() in [s.upper() for s in self.selected_adas]:
                    system_name = mod.strip().upper()
                    break
            # fallback to first parenthesis or last word
            if not system_name:
                if module_matches:
                    system_name = module_matches[0].strip().upper()
                else:
                    system_name = os.path.splitext(doc_name)[0].split()[-1].strip().upper()
    
            key = (year, self.sharepoint_make, model, system_name)
        else:
            key = (year, self.sharepoint_make, model, doc_name)
    
        # If we didn’t find a matching cell, create one at the bottom
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
    
            # ✅ Place RED NAME text in the correct column depending on mode
            if self.repair_mode:
                error_column = 7    # Column G for Repair mode
            elif self.excel_mode == "new":
                error_column = 10   # Column J for New mode
            else:
                error_column = 11   # Column K for OG mode
    
            error_cell = ws.cell(row=cell.row, column=error_column)
            error_cell.value = doc_name.splitlines()[0]
            error_cell.font = Font(color="FF0000")
    
        # ✅ Place the hyperlink itself
        cell.hyperlink = document_url
        cell.value = document_url
        cell.font = Font(color="0000FF", underline='single')
    
        # ✅ Track the row so we don’t add duplicates later
        adas_last_row[key] = cell.row
    
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
            module_matches = re.findall(r'\((.*?)\)', file_name)
            system_name = None
            for mod in module_matches:
                if mod.strip().upper() in [s.upper() for s in self.selected_adas]:
                    system_name = mod.strip().upper()
                    break
            if not system_name:
                if module_matches:
                    system_name = module_matches[0].strip().upper()
                else:
                    system_name = file_name.split()[-1].strip().upper()
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
        # Build once per file
        adas_file_name = file_name.replace(year, "").replace(make, "").replace(model, "")
        adas_file_name = re.sub(r"[\[\]()\-]", "", adas_file_name).replace("WL", "").replace("BSM-RCTW", "BSW-RCTW").strip().upper()
        
        for row in ws.iter_rows(min_row=2, max_col=20):
            if not any(cell.value for cell in row):
                continue  # skip empty rows
        
            year_value = str(row[0].value).strip() if row[0].value else ''
            make_value = str(row[1].value).replace("audi", "Audi").strip() if row[1].value else ''
            model_value = str(row[2].value).replace("Super Duty F-250", "F-250 SUPER DUTY") \
                .replace("Super Duty F-350", "F-350 SUPER DUTY").replace("Super Duty F-450", "F-450 SUPER DUTY") \
                .replace("Super Duty F-550", "F-550 SUPER DUTY").replace("Super Duty F-600", "F-600 SUPER DUTY") \
                .replace("MACH-E", "Mustang Mach-E ").replace("G Convertable", "G Convertible") \
                .replace("Carnival MPV", "Carnival").replace("RANGE ROVER VELAR", "VELAR") \
                .replace("RANGE ROVER SPORT", "SPORT").replace("Range Rover Sport", "SPORT") \
                .replace("RANGE ROVER EVOQUE", "EVOQUE").replace("MX5", "MX-5").strip() if row[2].value else ''
        
            # ADAS column (E vs T)
            if self.excel_mode == "new" and len(row) > 19 and row[19].value:
                adas_value = str(row[18].value).replace(".pdf", "").replace("(", "").replace(")", "").strip()
            elif len(row) > 4 and row[7].value:
                adas_value = str(row[7].value).replace("%", "").replace("(", "").replace(")", "").replace("-", "/") \
                    .replace("SCC 1", "ACC").replace(".pdf", "").strip()
            else:
                adas_value = ''
        
            # Compare
            if (
                year_value.strip().upper() == year.strip().upper()
                and make_value.strip().upper() == make.strip().upper()
                and model_value.strip().upper() == model.strip().upper()
                and adas_value.strip().upper() in adas_file_name
            ):
                #print(f"✅ MATCHED: {year_value} {make_value} {model_value} {adas_value}")
                return ws.cell(row=row[0].row, column=self.HYPERLINK_COLUMN_INDEX), None
        
        # If no match found
        print(f"❌ No match found in any row for {file_name}")
        return None, file_name

   
    def __build_row_index__(self, ws, repair_mode=False):
        index = {}
        for row in ws.iter_rows(min_row=2, max_col=8):
            year = str(row[0].value).strip().upper() if row[0].value else ''
            make = str(row[1].value).strip().upper() if row[1].value else ''
            model = str(row[2].value).strip().upper() if row[2].value else ''
            
            if repair_mode:
                if self.excel_mode == "new":
                    sys_cell = row[19]  # Column T
                elif self.sharepoint_make.lower() == "toyota":
                    sys_cell = row[4]   # Column E
                else:
                    sys_cell = row[3]   # Column D
            
                system = str(sys_cell.value).strip().upper() if sys_cell.value else ''
            else:
                if self.excel_mode == "new" and len(row) > 19:
                    system = str(row[19].value).strip().upper() if row[19].value else ''
                else:
                    system = str(row[4].value).strip().upper() if len(row) > 4 and row[4].value else ''
            

            normalized_system = re.sub(r"[^A-Z0-9]", "", system)
            key = (year, make, model, normalized_system)
            index[key] = row[0].row
        return index
      

#####################################################################################################################################################

if __name__ == '__main__':   

   # (Individual File testing without GUI, take away the # to perform whichever is needed)) 
   # excel_file_path = r'C:\Users\dromero3\Desktop\Excel Documents\Toyota Pre-Qual Long Sheet v6.3.xlsx'
   # sharepoint_link = 'https://calibercollision.sharepoint.com/:f:/g/enterpriseprojects/VehicleServiceInformation/EiB53aPXartJhkxyWzL5AFABZQsY3x-XDWPXQCqgFIrvoQ?e=m4DrKQ'
   # debug_run = True


    sharepoint_link = sys.argv[1]
    excel_file_path = sys.argv[2]
    debug_run = True
    
    extractor = SharepointExtractor(sharepoint_link, excel_file_path, debug_run)

    print("=" * 100)

    if extractor.cleanup_mode:
        # Clean up mode: find broken links → re-index only those
        extractor.populate_excel_file([])

        if extractor.broken_entries:
            print("🔁 Re-indexing SharePoint to replace broken links...")
            _, filtered_files = extractor.extract_contents()
            print(f"📥 Matched {len(filtered_files)} files for repair.")
            extractor.populate_excel_file(filtered_files)

    else:
        # Normal mode: index entire folder and populate Excel
        folders, files = extractor.extract_contents()
        extractor.populate_excel_file(files)

    print("=" * 100)
    print(f"Extraction and population for {extractor.sharepoint_make} is complete!")
