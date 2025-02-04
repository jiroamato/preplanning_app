import sys
import os
import PySimpleGUI as sg
from fillpdf import fillpdfs
from pathlib import Path
import logging
from datetime import datetime, date
from dateutil import parser
from dateutil.relativedelta import relativedelta
import re
import tkinter as tk
import locale
import time
import math
from PIL import Image

def get_absolute_path(relative_path):
    """Get absolute path for both development and compiled environments"""
    if getattr(sys, 'frozen', False):
        # Running in a bundle (compiled)
        base_path = sys._MEIPASS
    else:
        # Running in normal Python environment
        base_path = os.path.dirname(os.path.abspath(__file__))
    
    return os.path.join(base_path, relative_path)

def resize_image(image_path, size):
    """Resize image to fit within specified size while maintaining aspect ratio"""
    try:
        # Open the image directly (image_path is already absolute)
        img = Image.open(image_path)
        
        # Calculate aspect ratios
        target_ratio = size[0] / size[1]
        img_ratio = img.width / img.height
        
        if img_ratio > target_ratio:
            # Width is limiting factor
            new_width = size[0]
            new_height = int(new_width / img_ratio)
        else:
            # Height is limiting factor
            new_height = size[1]
            new_width = int(new_height * img_ratio)
            
        # Resize image
        img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
        
        # Create temp directory in the same directory as the script
        temp_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'temp')
        os.makedirs(temp_dir, exist_ok=True)
        
        # Create temp filename using just the base name of the original file
        temp_filename = f"temp_{os.path.basename(image_path)}"
        temp_path = os.path.join(temp_dir, temp_filename)
        
        # Save resized image
        img.save(temp_path)
        return temp_path
        
    except Exception as e:
        print(f"Error resizing image: {str(e)}")
        return image_path

class PDFTypes:
    PRE_ARRANGED_FULL_SERVICE = "Pre-Arranged Funeral Service Agreement - New.pdf"
    TRUSTAGE_APPLICATION = "Protector Plus TruStage Application form - New.pdf"
    PERSONAL_INFO_SHEET = "Personal Information Sheet - New.pdf"
    INSTRUCTIONS = "Instructions Concerning My Arrangements - New.pdf"
    JOURNEY_HOME = "Journey Home Enrollment Form - New.pdf"


class PDFAutofiller:
    TEXT_WIDTH = 40
    INPUT_WIDTH = 20
    DOLLAR_WIDTH = 10
    
    def __init__(self):
        logging.basicConfig(level=logging.INFO)
        
        sg.LOOK_AND_FEEL_TABLE['KFSTheme'] = {
                'BACKGROUND': '#f0f0f0',  # Light gray background
                'TEXT': '#000000',        # Black text
                'INPUT': '#ffffff',       # White input fields
                'TEXT_INPUT': '#000000',  # Black text in input fields
                'SCROLL': '#c7e0dc',      # Light teal for scrollbars
                'BUTTON': ('#000000', '#e7e7e7'),  # Black text on light gray buttons
                'PROGRESS': ('#01826B', '#D0D0D0'),
                'BORDER': 1,
                'SLIDER_DEPTH': 0,
                'PROGRESS_DEPTH': 0,
            }
        
        
        # Apply the theme
        sg.theme('KFSTheme')
    
        # Set font for all elements
        sg.set_options(font=("Helvetica", 10))
        
        self.base_path = self.get_base_path()
        self.setup_logging()
        
        # Add logo paths using existing base_path
        self.kearney_logo = os.path.join(self.base_path, "Logos", "Kearney Logo.png")
        self.burquitlam_logo = os.path.join(self.base_path, "Logos", "Burquitlam Logo.png")
        
        # Icon Path
        self.kearney_icon = os.path.join(self.base_path, "Logos", "Kearney Pin.ico")
        
        self.check_inactivity_timer = None
               
        self.phone_input_keys = [
            "-PHONE-", "Phone_3", "Representative Phone"
        ]
        
        self.discount_amount_keys = [('-DISCOUNT-AMT-', 0)
        ]
        
        # Listbox initialization
        self.viewing_listbox = None
        self.limousine_listbox = None
        self.casket_listbox = None
        self.urn_listbox = None
        self.keepsake_listbox = None
        self.crematorium_listbox = None
        self.other_3_listbox = None
        self.reception_facilities_listbox = None
        self.weekend_listbox = None
        
        self.dollar_input_keys = [
            "A1", "A2A", "A2B", "A2C", "A2D", "A3", "A4A", "A4B", "A4C",
            "A5A", "A5B", "A5C", "A5D", "A6", "A7", "A8", "A9A", "A9B",
            "A9C", "A9D", "A9E", "A9F", "A9G", "A9H", "Total A",
            "B1", "B2", "B3", "B4", "B5", "B6", "B7", "Total B",
            "C1", "C2", "C3", "C4", "C5", "C6", "C7", "C8", "C9", "C10", "Total C",
            "D1", "D2", "D3", "D4", "D5", "D6", "D7", "D8", "Total D",
            "Total \\(ABC\\)", "Discount", "GST", "PST", "Total D_2", "Grand Total",
            "3A Goods and Services", "3B MonumentMarker", "3C Other Expenses",
            "3D Final Documents Service", "3E Journey Home", "Total 3",
            "4A Single Pay", "4B Time Pay", "4C Single Pay Journey Home", "4D LPR", "Total 4 \\(ABCD\\)",
            ('-DISCOUNT-AMT-', 0)
        ]
        
        self.postal_input_keys = [
            "-POSTAL-", "Postal Code_4"
        ]
        
        self.last_value = {}
        
        self.location_data = {
            "Kearney Funeral Services (KFS)": {'ESTABLISHMENT_NAME': 'Kearney Funeral Services (KFS)', 'ESTABLISHMENT_EMAIL': 'Vancouver.Chapel@KearneyFS.com', 'ESTABLISHMENT_PHONE': '604-736-2668', 'ESTABLISHMENT_ADDRESS': '450 W 2nd Ave', 'ESTABLISHMENT_CITY': 'Vancouver', 'ESTABLISHMENT_PROVINCE': 'BC', 'ESTABLISHMENT_POSTAL_CODE': 'V5Y 1E2'},
            "Kearney Burnaby Chapel (KBC)": {'ESTABLISHMENT_NAME': 'Kearney Burnaby Chapel', 'ESTABLISHMENT_EMAIL': 'Burnaby@KearneyFS.com', 'ESTABLISHMENT_PHONE': '604-299-6889', 'ESTABLISHMENT_ADDRESS': '4715 Hastings St', 'ESTABLISHMENT_CITY': 'Burnaby', 'ESTABLISHMENT_PROVINCE': 'BC', 'ESTABLISHMENT_POSTAL_CODE': 'V5C 2K8'},
            "Kearney Burquitlam Funeral Home (BFH)": {'ESTABLISHMENT_NAME': 'Kearney Burquitlam Funeral Home', 'ESTABLISHMENT_EMAIL': 'Info@BurquitlamFuneralHome.ca', 'ESTABLISHMENT_PHONE': '604-936-9987', 'ESTABLISHMENT_ADDRESS': '102-200 Bernatchey St', 'ESTABLISHMENT_CITY': 'Coquitlam', 'ESTABLISHMENT_PROVINCE': 'BC', 'ESTABLISHMENT_POSTAL_CODE': 'V3K 0H8'},
            "Kearney Columbia-Bowell Chapel (CBC)": {'ESTABLISHMENT_NAME': 'Kearney Columbia-Bowell Chapel', 'ESTABLISHMENT_EMAIL': 'Columbia-Bowell@KearneyFS.com', 'ESTABLISHMENT_PHONE': '604-521-4881', 'ESTABLISHMENT_ADDRESS': '219 6th Street', 'ESTABLISHMENT_CITY': 'New Westminster', 'ESTABLISHMENT_PROVINCE': 'BC', 'ESTABLISHMENT_POSTAL_CODE': 'V3L 3A3'},
            "Kearney Cloverdale & South Surrey (CLO)": {'ESTABLISHMENT_NAME': 'Kearney Clovedale & South Surrey', 'ESTABLISHMENT_EMAIL': 'Cloverdale@KearneyFS.com', 'ESTABLISHMENT_PHONE': '604-574-2603', 'ESTABLISHMENT_ADDRESS': '17667 57th Ave', 'ESTABLISHMENT_CITY': 'Surrey', 'ESTABLISHMENT_PROVINCE': 'BC', 'ESTABLISHMENT_POSTAL_CODE': 'V3S 1H1'}
        }

        self.kearney_locations = list(self.location_data.keys())
        
        self.gst_fields = [
            "A1", "A2A", "A2B", "A2C", "A2D", "A3", "A4A", "A4B",
            "A4C", "A5A", "A5B", "A5C", "A5D", "A6", "A7", "A8",
            "A9A", "A9B", "A9C", "A9D", "A9E", "A9F", "A9G", "A9H",
            "B1", "B2", "B3", "B4", "B5", "B6", "B7", "C1", "C2",
            "C3", "C4", "C5", "C6", "C7", "C8", "C9", "C10"
        ]
        
        self.pst_fields = [
            "B1", "B2", "B3", "B4", "B5", "B6", "B7", "C4", "C7"
        ]

        # Add this new attribute to store package data
        self.packages = {
            "Full Funeral Church - Cremation": {
                'A1': 525.00, 'A2A': 2755.00, 'A3': 895.00, 'A4B': 365.00,
                'A5A': 365.00, 'A5B': 695.00, 'A9D': 295.00, 'A9E': 525.00,
                'Casket': 'Mazri','B1': 795.00, 'Urn': 'Basic Cardboard Urn', 'B2': 35.00, 'Crematorium': 'West Shore',
                'C2': 745.00, 'C5': 40.00, 'Other_2': 'Cadence Legacy Planner',
                'C9': 599.00, 'Death_Certificates_Quantity': '1', 'D7': 27.00, '3E Journey Home': 595.00
            },
            "Full Funeral Church - Burial": {
                'A1': 525.00, 'A2A': 2755.00, 'A3': 895.00, 'A4B': 365.00,
                'A5A': 365.00, 'A5B': 695.00, 'A9D': 295.00, 'A9E': 525.00, 'Casket': 'Mazri',
                'B1': 795.00, 'C5': 40.00, 'Other_2': 'Cadence Legacy Planner',
                'C9': 599.00, 'Death_Certificates_Quantity': '1', 'D7': 27.00, '3E Journey Home': 595.00
            },
            "Full Funeral Chapel - Cremation": {
                'A1': 525.00, 'A2A': 2755.00, 'A3': 895.00, 'A4A': 895.00, 'A4B': 365.00, 
                'A5A': 365.00, 'A5B': 695.00, 'A9D': 295.00, 'A9E': 525.00, 'Casket': 'Mazri',
                'B1': 795.00, 'Urn': 'Basic Cardboard Urn', 'B2': 35.00, 'Crematorium': 'West Shore',
                'C2': 745.00, 'C5': 40.00, 'Other_2': 'Cadence Legacy Planner', 'C9': 599.00, 
                'Death_Certificates_Quantity': 1, 'D7': 27.00, '3E Journey Home': 595.00
            },
            "Full Funeral Chapel - Burial": {
                'A1': 525.00, 'A2A': 2755.00, 'A3': 895.00, 'A4A': 895.00, 'A4B': 365.00,
                'A5A': 365.00, 'A5B': 695.00, 'A9D': 295.00, 'A9E': 525.00, 'Casket': 'Mazri',
                'B1': 795.00, 'C5': 40.00, 'Other_2': 'Cadence Legacy Planner', 'C9': 599.00,
                'Death_Certificates_Quantity': 1, 'D7': 27.00, '3E Journey Home': 595.00
            },
            "Full Funeral Chapel - Cremation, Reception": {
                'A1': 525.00, 'A2A': 2755.00, 'A3': 895.00, 'A4A': 895.00, 'A4B': 365.00,
                'A5A': 365.00, 'A5B': 695.00, 'Reception Facilities': 'KFS New West Reception Room Rental (Disp Dishes, Coffee, Tea)',
                'A8': 595.00, 'A9D': 295.00, 'A9E': 525.00, 'Casket': 'Mazri',
                'B1': 795.00, 'Urn': 'Basic Cardboard Urn', 'B2': 35.00, 'Crematorium': 'West Shore',
                'C2': 745.00, 'C5': 40.00, 'Other_2': 'Cadence Legacy Planner', 'C9': 599.00,
                'Death_Certificates_Quantity': 1, 'D7': 27.00, '3E Journey Home': 595.00
            },
            "Full Funeral Chapel - Burial, Reception": {
                'A1': 525.00, 'A2A': 2755.00, 'A3': 895.00, 'A4A': 895.00, 'A4B': 365.00, 
                'A5A': 365.00, 'A5B': 695.00, 'Reception Facilities': 'KFS New West Reception Room Rental (Disp Dishes, Coffee, Tea)',
                'A8': 595.00, 'A9D': 295.00, 'A9E': 525.00, 'Casket': 'Mazri',
                'B1': 795.00, 'C5': 40.00, 'Other_2': 'Cadence Legacy Planner', 'C9': 599.00, 
                'Death_Certificates_Quantity': 1, 'D7': 27.00, '3E Journey Home': 595.00
            },
            "Minimum Cremation - No Viewing": {
                'A1': 525.00, 'A2A': 765.00, 'A3': 595.00, 'A4B': 365.00,
                'Casket': 'Basic Cremation Container', 'B1': 395.00, 'Urn': 'Basic Cardboard Urn',
                'B2': 35.00, 'Crematorium': 'West Shore',
                'C2': 745.00, 'C5': 40.00, 'Other_2': 'Cadence Legacy Planner',
                'C9': 599.00, 'Death_Certificates_Quantity': 1,
                'D7': 27.00, '3E Journey Home': 595.00
            },
            "Minimum Cremation - With Viewing": {
                'A1': 525.00, 'A2A': 895.00, 'A3': 350.00, 'A4B': 365.00, 'A5A': 365.00, 'A5B': 695.00,
                'Evening Prayers or Visitation': 'Viewing at New West', 'A6': 500.00,
                'Casket': 'Mazri', 'B1': 795.00, 'Urn': 'Basic Cardboard Urn', 'B2': 35.00, 'Crematorium': 'West Shore',
                'C2': 745.00, 'C5': 40.00, 'Other_2': 'Cadence Legacy Planner',
                'C9': 599.00, 'Death_Certificates_Quantity': 1, 'D7': 27.00, '3E Journey Home': 595.00
            },
            "Graveside - No Viewing": {
                'A1': 525.00, 'A2A': 1195.00, 'A3': 595.00, 'A4B': 365.00,
                'A5A': 365.00, 'A9E': 525.00, 'Casket': 'Mazri', 'B1': 795.00,'C5': 40.00,
                'Other_2': 'Cadence Legacy Planner',
                'C9': 599.00, 'Death_Certificates_Quantity': 1, 'D7': 27.00, '3E Journey Home': 595.00
            },
            "Graveside - With Viewing": {
                'A1': 525.00, 'A2A': 1195.00, 'A3': 595.00, 'A4B': 365.00, 'A5A': 365.00,
                'A5B': 695.00, 'Evening Prayers or Visitation': 'Viewing at New West',
                'A6': 500.00, 'A9E': 525.00, 'Casket': 'Mazri', 'B1': 795.00,'C5': 40.00,
                'Other_2': 'Cadence Legacy Planner',
                'C9': 599.00, 'Death_Certificates_Quantity': 1, 'D7': 27.00, '3E Journey Home': 595.00
            },
            "Memorial Service - No Reception, No Viewing": {
                'A1': 525.00, 'A2A': 1895.00, 'A3': 795.00, 'A4A': 895.00, 'A4B': 365.00, 'Casket': 'Basic Cremation Container',
                'B1': 395.00, 'Urn': 'Basic Cardboard Urn', 'B2': 35.00, 'Crematorium': 'West Shore', 'C2': 745.00, 'C5': 40.00,
                'Other_2': 'Cadence Legacy Planner',
                'C9': 599.00, 'Death_Certificates_Quantity': 1, 'D7': 27.00, '3E Journey Home': 595.00
            },
            "Memorial Service - No Reception, With Viewing": {
                'A1': 525.00, 'A2A': 1895.00, 'A3': 795.00, 'A4A': 895.00, 'A4B': 365.00, 'A5B': 695.00,
                'Evening Prayers or Visitation': 'Viewing at New West', 'A6': 500.00,
                'Casket': 'Basic Cremation Container', 'B1': 395.00, 'Urn': 'Basic Cardboard Urn',
                'B2': 35.00, 'Crematorium': 'West Shore', 'C2': 745.00, 'C5': 40.00,
                'Other_2': 'Cadence Legacy Planner',
                'C9': 599.00, 'Death_Certificates_Quantity': 1, 'D7': 27.00, '3E Journey Home': 595.00
            },
            "Memorial Service - With Reception, No Viewing": {
                'A1': 525.00, 'A2A': 1895.00, 'A3': 795.00, 'A4A': 895.00, 'A4B': 365.00,
                'Reception Facilities': 'KFS New West Reception Room Rental (Disp Dishes, Coffee, Tea)', 'A8': 595.00,
                'Casket': 'Basic Cremation Container', 'B1': 395.00, 'Urn': 'Basic Cardboard Urn', 'B2': 35.00, 'Crematorium': 'West Shore',
                'C2': 745.00, 'C5': 40.00, 'Other_2': 'Cadence Legacy Planner',
                'C9': 599.00, 'Death_Certificates_Quantity': 1, 'D7': 27.00, '3E Journey Home': 595.00
            },
            "Memorial Service - With Reception, With Viewing": {
                'A1': 525.00, 'A2A': 1895.00, 'A3': 795.00, 'A4A': 895.00, 'A4B': 365.00, 'A5B': 695.00,
                'Evening Prayers or Visitation': 'Viewing at New West', 'A6': 500.00,
                'Reception Facilities': 'KFS New West Reception Room Rental (Disp Dishes, Coffee, Tea)', 'A8': 595.00,
                'Casket': 'Mazri', 'B1': 795.00, 'Urn': 'Basic Cardboard Urn', 'B2': 35.00, 'Crematorium': 'West Shore',
                'C2': 745.00, 'C5': 40.00, 'Other_2': 'Cadence Legacy Planner',
                'C9': 599.00, 'Death_Certificates_Quantity': 1, 'D7': 27.00, '3E Journey Home': 595.00
            },
            "Witness Cremation - No Viewing": {
                'A1': 525.00, 'A2A': 1295.00, 'A3': 595.00, 'A4B': 365.00, 'A5A': 365.00, 'A9E': 525.00, 'Casket': 'Basic Cremation Container',
                'B1': 395.00, 'Urn': 'Basic Cardboard Urn', 'B2': 35.00, 'Crematorium': 'Watch Start Cremation (Maple Ridge)',
                'C2': 1495.00, 'C5': 40.00, 'Other_2': 'Cadence Legacy Planner',
                'C9': 599.00, 'Death_Certificates_Quantity': 1, 'D7': 27.00, '3E Journey Home': 595.00
            },
            "Witness Cremation - With Viewing": {
                'A1': 525.00, 'A2A': 1295.00, 'A3': 595.00, 'A4B': 365.00, 'A5A': 365.00, 'A5B': 695.00,
                'A9E': 525.00, 'Evening Prayers or Visitation': 'Viewing at New West', 'A6': 500.00,
                'Casket': 'Mazri', 'B1': 795.00, 'Urn': 'Basic Cardboard Urn', 'B2': 35.00, 'Crematorium': 'Watch Start Cremation (Maple Ridge)',
                'C2': 1495.00, 'C5': 40.00, 'Other_2': 'Cadence Legacy Planner',
                'C9': 599.00, 'Death_Certificates_Quantity': 1, 'D7': 27.00, '3E Journey Home': 595.00
            },
            "Chapel Rental Reception - Cremation or Burial Elsewhere": {
                'A2A': 2495.00, 'A4A': 895.00, 'Reception Facilities': 'KFS New West Reception Room Rental (Disp Dishes, Coffee, Tea)',
                'A8': 595.00, 'Other_2': 'Cadence Legacy Planner', 'C9': 599.00,
                'Death_Certificates_Quantity': 1, 'D7': 27.00, '3E Journey Home': 595.00
            },
            "Ship Out International - Service at Church": {
                'A1': 525.00, 'A2A': 2755.00, 'A3': 895.00, 'A4B': 695.00, 'A5A': 365.00, 'A5B': 695.00,
                'A9B': 395.00, 'Casket': 'Misty Blue - 18 gauge steel',
                'B1': 3995.00, 'Other_1': 'Casket Shipping Outer Container', 'B7': 750.00,
                'C5': 40.00, 'Other_2': 'Cadence Legacy Planner',
                'C9': 599.00, 'Death_Certificates_Quantity': 1, 'D7': 27.00,
                'Other_4': 'Airfare Estimate', 'D8': 3500.00, '3E Journey Home': 595.00
            },
            "Ship Out International - Service at Chapel": {
                'A1': 525.00, 'A2A': 2755.00, 'A3': 895.00, 'A4A': 895.00, 'A4B': 695.00, 'A5A': 365.00, 'A5B': 695.00,
                'Reception Facilities': 'KFS New West Reception Room Rental (Disp Dishes, Coffee, Tea)', 'A8': 595.00,
                'A9B': 395.00, 'Casket': 'Misty Blue - 18 gauge steel',
                'B1': 3995.00, 'Other_1': 'Casket Shipping Outer Container', 'B7': 750.00,
                'C5': 40.00, 'Other_2': 'Cadence Legacy Planner', 'C9': 599.00,
                'Death_Certificates_Quantity': 1, 'D7': 27.00, '3E Journey Home': 595.00,
                'Other_4': 'Airfare Estimate', 'D8': 3500.00
            },
            "Ship Out International - No Service": {
                'A1': 525.00, 'A2A': 995.00, 'A3': 895.00, 'A4B': 695.00, 'A5A': 365.00, 'A5B': 695.00,
                'A9B': 395.00, 'Casket': 'Misty Blue - 18 gauge steel', 'B1': 3995.00,
                'Other_1': 'Casket Shipping Outer Container', 'B7': 750.00,
                'C5': 40.00, 'Other_2': 'Cadence Legacy Planner', 'C9': 599.00,
                'Death_Certificates_Quantity': 1, 'D7': 27.00,
                'Other_4': 'Airfare Estimate', 'D8': 3500.00
            }
        }
        
        self.viewings = {
            "Evening Prayers (includes Staff)": 695.00,
            "Viewing at New West": 500.00,
            "Viewing at Burnaby": 895.00,
            "Viewing at Vancouver": 895.00,
            "Viewing after 5PM (per hour)": 150.00,
            "Afterhours Visitation Staff": 695.00,
            "St Clare of Assisi Hall Viewing": 500.00,
            "St Clare of Assisi Hall Service": 1200.00
        }
        
        self.limousines = {
            "KFS Limousine (4 hour use)": 695.00,
            "KFS Horse Drawn Carriage": 2500.00
        }
        
        self.caskets = {
            "Basic Cremation Container": 395.00,
            "125 Oak Veneer": 2195.00,
            "Aspen Pine": 1795.00,
            "Moka Pine": 1795.00,
            "Brownsville Pine": 1995.00,
            "Brownsville Pine (oversize)": 3595.00,
            "Brownsville Cedar": 3595.00,
            "Brownsville Cedar (without blanket or paddle handles)": 2695.00,
            "60 Ash": 2695.00,
            "Windfield PC": 3560.00,
            "Wheat Maple": 2995.00,
            "Atlantic Poplar": 2995.00,
            "Stewart Ash": 3495.00,
            "Heavenly White Poplar": 3695.00,
            "Langara - Pawlownia": 3595.00,
            "Riley - Poplar": 3695.00,
            "700 Octagon Oak": 4195.00,
            "725 Octagon Oak": 4195.00,
            "Branson - Poplar": 3695.00,
            "Apostle Oak": 4395.00,
            "Stanley Oak": 4195.00,
            "Sherwood Oak": 3895.00,
            "Homeward - Poplar": 4295.00,
            "Basilica - Poplar": 4895.00,
            "Lotus - Poplar": 4695.00,
            "Michelangelo - Maple": 5795.00,
            "Carnaby Poplar": 3995.00,
            "Morgan Cherry": 8188.00,
            "Ambassador Full Couch - Mahogany": 7195.00,
            "168 Oak Full Couch": 5795.00,
            "Diplomat - Mahogany": 9888.00,
            "Executive Full Couch - Mahogany": 9888.00,
            "Dynasty - Mahogany": 14888.00,
            "Misty Blue - 18 gauge steel": 3995.00,
            "Capilano Blue - 18 gauge steel": 3995.00,
            "Nelson Brown - 18 gauge steel": 3995.00,
            "Deauville White - 18 gauge steel": 4395.00,
            "Aurora Copper - 32 oz Brushed Copper": 10495.00,
            "Natura": 3195.00,
            "Willow": 2995.00,
            "Bamboo shroud with use of Ceremonial willow carrier": 755.00,
            "Seagrass": 3495.00,
            "Eco Pine w/ rope handles and simple lining": 695.00,
            "Eco Pine w/ rope handles and simple lining (oversize)": 1095.00,
            "Mazri": 795.00,
            "Mazri (Oversized)": 1795.00,
            "30 Grey": 1075.00,
            "Genoa": 2195.00,
            "Navy Tabor": 2095.00,
            "White Ventura": 2195.00,
            "Dominion Maple Ceremonial (Rental casket)": 1795.00,
            "Northgate Oak Ceremonial (Rental casket)": 1355.00,
            "Burlington Cremation Container": 795.00,
            "McConnell Cremation Container": 995.00,
            "BP2 - BC Pine": 1195.00,
            "BP2 - BC Pine (oversize)": 1395.00,
            "BPO - BC Pine": 695.00,
            "BPO - BC Pine (oversize)": 795.00,
            "Basic Wooden (Oversize Wooden)": 495.00,
            "Cremation Tray (Minimum requirement)": 395.00
        }

        self.urns = {
            "Basic Cardboard Urn": 35.00,
            "Cherry/Maple/Walnut": 475.00,
            "LGV original series Natural/Red/Ebony": 895.00,
            "LGV ecologique series Natural/Red/Ebony": 695.00,
            "LGV Keepsake Natural/Red/Ebony": 195.00,
            "Sheesham": 250.00,
            "Sheesham Small": 85.00,
            "Providence Mahogany/Oak": 400.00,
            "Winchester Cherry/Oak": 475.00,
            "Hamilton Cherry/Natural": 175.00,
            "Hamilton Keepsake Cherry/Natural": 75.00,
            "Photo Urn Cherry/Natural/Black": 250.00,
            "Photo Urn Keepsake Cherry/Natural/Black": 95.00,
            "Metro Mantel Urn": 495.00,
            "Cathedral Ivory/Forest Green": 445.00,
            "Round Marble": 650.00,
            "Rectangular Marble": 595.00,
            "Modern Companion Urn Collection": 695.00,
            "Together Forever Companion (EA1007-E)": 1095.00,
            "Divine Square Companion": 1095.00,
            "Divine Heart Companion": 1195.00,
            "Aria Wheat/Bird/Tree/Butterfly/Rose": 485.00,
            "Aria Heart Wheat/Bird/Tree/Butterfly/Rose": 115.00,
            "Aria Keepsake Wheat/Bird/Tree/Butterfly/Rose": 65.00,
            "Classic Going Home": 395.00,
            "Classic Going Home Heart": 115.00,
            "Classic Going Home Keepsake": 65.00,
            "Mother of Pearl Aristocrat": 495.00,
            "Mother of Pearl Keepsake": 65.00,
            "Aristocrat Going Home": 445.00,
            "Aristocrat Going Home Heart": 115.00,
            "Quad Gold or Copper (Incl. 1 emblem)": 595.00,
            "Aristocrat Going Home Keepsake": 65.00,
            "Aristocrat Black & Gold": 495.00,
            "Aristocrat Black & Gold Keepsake": 65.00,
            "Classic Bronze/Pewter": 395.00,
            "Classic Heart Bronze/Pewter": 115.00,
            "Classic Keepsake Bronze/Pewter": 65.00,
            "Trinity Midnight Blue/Pearl White": 495.00,
            "Trinity Heart Midnight Blue/Pearl White": 115.00,
            "Trinity Keepsake Midnight Blue/Pearl White": 75.00,
            "Grecian Crimson Red/Rustic Bronze/Pewter": 495.00,
            "Grecian Keepsake Red/Bronze/Pewter": 85.00,
            "Monterey Purple/Blue/Ruby": 445.00,
            "Monterey Keepsake Purple/Blue/Ruby": 85.00,
            "Anoka Shimmering Grey/Blue": 515.00,
            "Anoka Keepsake Shimmering Grey/Blue": 85.00,
            "Celeste Charcoal/Pearl/Indigo": 445.00,
            "Celeste Heart Charcoal/Pearl/Indigo": 115.00,
            "Celeste Keepsake Charcoal/Pearl/Indigo": 75.00,
            "Blessing": 595.00,
            "Blessing Keepsake Pearl": 145.00,
            "Monarch Jali": 595.00,
            "Monarch Jali Keepsake": 145.00,
            "LoveBird Bronze/Midnight": 635.00,
            "LoveBird Keepsake Bronze/Midnight": 165.00,
            "LoveHeart Pearl/Red": 495.00,
            "LoveHeart Medium Blue/Pink": 295.00,
            "CuddleBear Medium Blue/Pink": 375.00,
            "Wings of Hope Pink/Blue/Pearl": 495.00,
            "Wing of Hope Medium Pink/Blue/Pearl": 295.00,
            "Saturn Bronze": 595.00,
            "Art Deco Classic": 495.00,
            "Art Deco Classic Keepsake": 165.00,
            "Adore Bronze/Midnight": 635.00,
            "Adore Keepsake Bronze/Midnight": 165.00,
            "Bright Silver Keepsake Heart": 165.00,
            "Bright Silver Keepsake Star": 165.00,
            "Brushed Gold Keepsake Star": 165.00,
            "Pride Rainbow": 395.00,
            "Pride Rainbow Heart Keepsake": 115.00,
            "Pride Rainbow Tealight Urn": 175.00,
            "Tealight Urn Bronze/Midnight": 175.00,
            "Pearl Tealight Urn Lavender/Pink/Blue": 175.00,
            "Princess Cat Bronze/Midnight/Pearl": 195.00,
            "Bright Silver Keepsake Paw": 95.00,
            "Tall Pet Tealight Urn Bronze/Midnight": 295.00,
            "Pet Tealight Urn Bronze/Midnight": 175.00,
            "Etienne Autumn/Butterfly/Rose": 545.00,
            "Etienne Keepsake Autumn/Butterfly/Rose": 135.00,
            "Essence Opal": 495.00,
            "Essence Cloisonné Heart Opal": 135.00,
            "Essence Keepsake Opal": 135.00,
            "Elite Pink/Blue": 545.00,
            "Elite Cloisonné Heart Pink/Blue": 135.00,
            "Elite Keepsake Pink/Blue": 135.00,
            "Scattering Tubes (Variety)": 185.00,
            "Scattering Tube Keepsake (Variety)": 65.00,
            "Journey Earth Aqua/Navy/Natural/Green": 295.00,
            "Journey Keepsake Aqua/Navy/Natural/Green": 95.00,
            "Bios Urn": 275.00,
            "Himalayan Rock Salt": 545.00,
            "Himalayan Salt Keepsake": 295.00,
            "Embrace Autumn Leaves": 295.00,
            "Turtle": 495.00,
            "Turtle Small Keepsake": 195.00,
            "Oceane Natural Urn": 545.00,
            "Oceane Natural Urn Keepsake": 195.00,
            "Bamboo Simplicity": 295.00,
            "Heart Urn (EA1002-E)": 545.00,
            "Heart Photo (EA4002-E)": 145.00,
            "Heart Keepsake (EA3002-E)": 145.00,
            "Heart Candle (EA2002-E)": 145.00,
            "Sunrays Candle (EA3005-E)": 145.00,
            "Seashells Candle (EA3006-E)": 145.00,
            "InFlight (EA1001)": 595.00,
            "InFlight Keepsake (EA3001)": 145.00,
            "Slate": 695.00,
            "Slate keepsake": 295.00,
            "Resin Urn": 445.00,
            "Arielle Heart Pink/Blue": 175.00,
            "Teddybear Pink/Blue": 145.00,
            "Heart Stands (each)": 20.00,
            "Temporary plastic urn": 40.00,
            "Together Forever (EA5007) holds no remains": 125.00,
            "Heart Plaque (EA5002-E)holds no remains": 145.00,
            "Engraving (Name and Dates)": 155.00,
            "Engraving (Per Line) Marble rectangle": 185.00,
            "Sticker Back Tags (name and dates)": 65.00,
            "Pendants (incl. satin string name and dates)": 95.00
        }
        
        self.crematorium = {
            "West Shore": 745.00,
            "Maple Ridge": 945.00,
            "Oversized - West Shore (300+)": 1045.00,
            "Oversized - Maple Ridge (300+)": 1255.00,
            "Watch Start Cremation (Maple Ridge)": 1495.00
        }
        
        self.other_3 = {
            "Rush Cremation Fee (72 hours from registration)": 350.00,
            "Scheduled Start Cremation Fee (no family)": 1245.00
        }
        
        self.reception_facilities = {
            "KFS New West Reception Room Rental (Disp Dishes, Coffee, Tea)": 595.00,
            "KFS New West Reception Room Rental (Glassware, Coffee, Tea)": 395.00,
            "KFS Vancouver Reception Room Rental (Facility Only)": 795.00
        }
        
        self.weekend = {
            "Sundays and Holiday Service Fee": 895.00,
            "Saturday Service Fee": 895.00
        }
        
        self.discounts = {
            "Cadence": 400.00,
            "Casket": 0.00,
            "Urn": 0.00,
            "Keepsake": 0.00,
            "Traditional Mourning Items": 0.00,
            "Memorial Stationary": 0.00,
            "Funeral Register": 0.00
        }
                
        self.payment_factors = {
            '0_to_54': {'3-year': 0.03150, '5-year': 0.01995, '10-year': 0.01155, '15-year': 0.008925, '20-year': 0.00735},
            '55_to_59': {'3-year': 0.03255, '5-year': 0.01995, '10-year': 0.01260, '15-year': 0.009975, '20-year': 0.00840},
            '60_to_64': {'3-year': 0.03255, '5-year': 0.02100, '10-year': 0.01365, '15-year': 0.01050, '20-year': 0.008925},
            '65': {'3-year': 0.03360, '5-year': 0.02100, '10-year': 0.01470, '15-year': 0.01155, '20-year': 0.009975},
            '66_to_69': {'3-year': 0.03360, '5-year': 0.02100, '10-year': 0.01470, '15-year': 0.01155, '20-year': None},
            '70': {'3-year': 0.03465, '5-year': 0.02205, '10-year': 0.01575, '15-year': 0.01260, '20-year': None},
            '71_to_74': {'3-year': 0.03465, '5-year': 0.02205, '10-year': 0.01575, '15-year': None, '20-year': None},
            '75': {'3-year': 0.03675, '5-year': 0.02310, '10-year': 0.01680, '15-year': None, '20-year': None},
            '76_to_79': {'3-year': 0.03675, '5-year': 0.02310, '10-year': None, '15-year': None, '20-year': None},
            '80': {'3-year': 0.03780, '5-year': 0.02625, '10-year': None, '15-year': None, '20-year': None},
            '81_to_82': {'3-year': 0.03780, '5-year': None, '10-year': None, '15-year': None, '20-year': None}
        }
        
        self.selected_package = ""

        personal_info_layout = [
            [sg.Text("Applicant Information", font=("Helvetica", 14, "bold"))],
            [sg.Text("First Name:"), sg.Input(key="-FIRST-", size=(20, 1)),
             sg.Text("Middle Name:"), sg.Input(key="-MIDDLE-", size=(20, 1)),
             sg.Text("Last Name:"), sg.Input(key="-LAST-", size=(20, 1))],
            [sg.Text("Birthdate:"), sg.Input(key="-BIRTHDATE-", size=(20, 1), enable_events=True),
             sg.Text("Age:"), sg.Input(key="-AGE-", size=(5, 1), default_text=""),
             sg.Text("Gender:"), sg.Input(key="-GENDER-", size=(20, 1)),
             sg.Text("SIN:"), sg.Input(key="SIN", size=(20, 1), enable_events=True)],
            [sg.Text("Phone:"), sg.Input(key="-PHONE-", size=(20, 1), enable_events=True),
             sg.Text("Email:"), sg.Input(key="-EMAIL-", size=(30, 1)),
             sg.Text("Occupation:"), sg.Input(key="-OCCUPATION-", size=(20, 1))],
            [sg.Text("Address:"), sg.Input(key="-ADDRESS-", size=(30, 1)),
             sg.Text("City:"), sg.Input(key="-CITY-", size=(20, 1)),
             sg.Text("Province:"), sg.Input(key="-PROVINCE-", size=(10, 1)),
             sg.Text("Postal Code:"), sg.Input(key="-POSTAL-", size=(10, 1), enable_events=True)],
            [sg.HorizontalSeparator()],
            [sg.Text("Beneficiary Information", font=("Helvetica", 14, "bold"))],
            [sg.Text("Full Name:"), sg.Input(key="Name", size=(30, 1)),
             sg.Text("Relationship:"), sg.Input(key="Relationship", size=(20, 1)),
             sg.Text("Phone:"), sg.Input(key="Phone_3", size=(20, 1), enable_events=True)],
            [sg.Text("Email:"), sg.Input(key="Email_3", size=(30, 1)),
             sg.Checkbox("Same address as Applicant", key="-SAME_ADDRESS-", enable_events=True)],
            [sg.Text("Address:"), sg.Input(key="Address \\(if different\\)", size=(30, 1)),
             sg.Text("City:"), sg.Input(key="City_4", size=(20, 1)),
             sg.Text("Province:"), sg.Input(key="Province_4", size=(10, 1)),
             sg.Text("Postal Code:"), sg.Input(key="Postal Code_4", size=(10, 1), enable_events=True)]
        ]

        section_a_layout = [
            # A. Professional Services
            [sg.Text("A. Professional Services", font=("Helvetica", 16, "bold")), sg.Text("(GST applicable)", font=("Helvetica", 10, "italic"))],
            [sg.Column([
                [sg.Text("1. Securing Release of Deceased and Transfer:", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("2. Services of Licensed Funeral Directors and Staff:", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("    Pallbearers:", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("    Alternate Day Interment:", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("    Alternate Day Interment 2:", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("3. Administration, Documentation & Registration:", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("4. Facilities and/or Equipment and Supplies:", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("    Sheltering of Remains:", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("    A/V Equipment:", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("5. Preparation Services", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("    a) Basic Sanitary Care, Dressing and Casketing:", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("    b) Embalming:", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("    c) Pacemaker Removal:", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("    d) Autopsy Care:", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("6. Evening Prayers or Visitation:", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("7. Weekend or Statutory Holiday:", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("8. Reception Facilities:", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("9. Vehicles", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("    Delivery of Cremated Remains:", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("    Transfer Vehicle for Transfer to Crematorium or Airport:", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("    Lead Vehicle:", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("    Service Vehicle:", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("    Funeral Coach:", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("    Limousine:", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("    Additional Limousines:", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("    Flower Van:", size=(self.TEXT_WIDTH, 1))],
            ]), sg.Column([
                [sg.Push(), sg.Text("$"), sg.Input(key="A1", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Text("$"), sg.Input(key="A2A", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Pallbearers", size=(self.INPUT_WIDTH, 1)), sg.Text("$"), sg.Input(key="A2B", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Alternate Day Interment 1", size=(self.INPUT_WIDTH, 1)), sg.Text("$"), sg.Input(key="A2C", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Alternate Day Interment 2", size=(self.INPUT_WIDTH, 1)), sg.Text("$"), sg.Input(key="A2D", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Text("$"), sg.Input(key="A3", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Text("$"), sg.Input(key="A4A", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Text("$"), sg.Input(key="A4B", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Text("$"), sg.Input(key="A4C", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Text("")],  # Empty space for "5. Preparation Services"
                [sg.Push(), sg.Text("$"), sg.Input(key="A5A", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Text("$"), sg.Input(key="A5B", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Pacemaker Removal", size=(self.INPUT_WIDTH, 1), enable_events=True), sg.Text("$"), sg.Input(key="A5C", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Autopsy Care", size=(self.INPUT_WIDTH, 1), enable_events=True), sg.Text("$"), sg.Input(key="A5D", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Evening Prayers or Visitation", size=(self.INPUT_WIDTH, 1), enable_events=True), sg.Text("$"), sg.Input(key="A6", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Weekend or Statutory Holiday", size=(self.INPUT_WIDTH, 1), enable_events=True), sg.Text("$"), sg.Input(key="A7", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Reception Facilities", size=(self.INPUT_WIDTH, 1), enable_events=True), sg.Text("$"), sg.Input(key="A8", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Text("")],  # Empty space for "9. Vehicles"
                [sg.Push(), sg.Input(key="Delivery of Cremated Remains", size=(self.INPUT_WIDTH, 1), enable_events=True), sg.Text("$"), sg.Input(key="A9A", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Transfer to Crematorium or Airport", size=(self.INPUT_WIDTH, 1), enable_events=True), sg.Text("$"), sg.Input(key="A9B", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Lead Vehicle", size=(self.INPUT_WIDTH, 1), enable_events=True), sg.Text("$"), sg.Input(key="A9C", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Service Vehicle", size=(self.INPUT_WIDTH, 1), enable_events=True), sg.Text("$"), sg.Input(key="A9D", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Funeral Coach", size=(self.INPUT_WIDTH, 1), enable_events=True), sg.Text("$"), sg.Input(key="A9E", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Limousine", size=(self.INPUT_WIDTH, 1), enable_events=True), sg.Text("$"), sg.Input(key="A9F", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Additional Limousines", size=(self.INPUT_WIDTH, 1), enable_events=True), sg.Text("$"), sg.Input(key="A9G", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Flower Van", size=(self.INPUT_WIDTH, 1), enable_events=True), sg.Text("$"), sg.Input(key="A9H", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
            ])],
            [sg.Text("TOTAL A:", font=("Helvetica", 10, "bold"), size=(self.TEXT_WIDTH, 1)), sg.Push(), sg.Text("$"), sg.Input(key="Total A", size=(self.DOLLAR_WIDTH, 1), enable_events=True)]
        ]
        
        section_b_layout = [
            # B. Merchandise
            [sg.Text("B. Merchandise", font=("Helvetica", 16, "bold")), sg.Text("(GST and/or PST applicable)", font=("Helvetica", 10, "italic"))],
            [sg.Column([
                [sg.Text("Casket:", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("Urn:", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("Keepsake:", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("Traditional Mourning Items:", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("Memorial Stationary:", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("Funeral Register:", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("Other:", size=(self.TEXT_WIDTH, 1))],
            ]), sg.Column([
                [sg.Push(), sg.Input(size=(40, 1),
                                     enable_events=True,
                                     key='Casket',
                                     background_color='white',
                                     text_color='black'),
                 sg.Text("$"), sg.Input(key="B1", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(size=(40, 1),
                                     enable_events=True,
                                     key='Urn',
                                     background_color='white',
                                     text_color='black'),
                 sg.Text("$"), sg.Input(key="B2", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(size=(40, 1),
                                     enable_events=True,
                                     key="Keepsake",
                                     background_color='white',
                                     text_color='black'),
                 sg.Text("$"), sg.Input(key="B3", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Traditional Mourning Items", size=(40 , 1)), sg.Text("$"), sg.Input(key="B4", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Memorial Stationary", size=(40, 1), visible=False, readonly=True),
                 sg.Text("Cards:"), sg.Input(key="Cards_Qty", size=(5, 1), enable_events=True), sg.Text("x $2.95 ="),
                 sg.Push(), sg.Text("$"), sg.Input(key="B5", size=(self.DOLLAR_WIDTH, 1), readonly=True)],
                [sg.Push(), sg.Text("Guest Book:"), sg.Input(key="Guest_Book_Qty", size=(5, 1), enable_events=True),
                 sg.Text("x $75.00 ="), sg.Push(), sg.Text("$"), sg.Input(key="B6", size=(self.DOLLAR_WIDTH, 1), readonly=True)],
                [sg.Push(), sg.Input(key="Other_1", size=(40, 1)), sg.Text("$"), sg.Input(key="B7", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
            ])],
            [sg.Text("TOTAL B:", font=("Helvetica", 10, "bold"), size=(self.TEXT_WIDTH, 1)), sg.Push(), sg.Text("$"), sg.Input(key="Total B", size=(self.DOLLAR_WIDTH, 1), enable_events=True)]
        ]
        
        section_c_layout = [
            # C. Cash Disbursements
            [sg.Text("C. Cash Disbursements", font=("Helvetica", 16, "bold")), sg.Text("(GST and/or PST applicable)", font=("Helvetica", 10, "italic"))],
            [sg.Column([
                [sg.Text("Cemetery:", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("Crematorium:", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("Obituary Notices:", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("Flowers:", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("CPBC Administration Fee:", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("Hostess:", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("Markers:", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("Catering:", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("Other:", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("Other 2:", size=(self.TEXT_WIDTH, 1))],
            ]), sg.Column([
                [sg.Push(), sg.Input(key="Cemetery", size=(40, 1), enable_events=True), sg.Text("$"), sg.Input(key="C1", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Crematorium", size=(40, 1), enable_events=True), sg.Text("$"), sg.Input(key="C2", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Obituary Notices", size=(40, 1), enable_events=True), sg.Text("$"), sg.Input(key="C3", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Flowers", size=(40, 1), enable_events=True), sg.Text("$"), sg.Input(key="C4", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="CPBC Administration Fee", size=(40, 1), enable_events=True), sg.Text("$"), sg.Input(key="C5", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Hostess", size=(40, 1), enable_events=True), sg.Text("$"), sg.Input(key="C6", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Markers", size=(40, 1), enable_events=True), sg.Text("$"), sg.Input(key="C7", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Catering", size=(40, 1), enable_events=True), sg.Text("$"), sg.Input(key="C8", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Other_2", size=(40, 1), enable_events=True), sg.Text("$"), sg.Input(key="C9", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Other_3", size=(40, 1), enable_events=True), sg.Text("$"), sg.Input(key="C10", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
            ])],
            [sg.Text("TOTAL C:", font=("Helvetica", 10, "bold"), size=(self.TEXT_WIDTH, 1)), sg.Push(), sg.Text("$"), sg.Input(key="Total C", size=(self.DOLLAR_WIDTH, 1), enable_events=True)]
        ]
        
        section_d_layout = [
            # D. Cash Disbursements
            [sg.Text("D. Cash Disbursements", font=("Helvetica", 16, "bold")), sg.Text("(GST exempt)", font=("Helvetica", 10, "italic"))],
            [sg.Column([
                [sg.Text("Clergy Honorarium:", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("Church Honorarium:", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("Altar Servers:", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("Organist:", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("Soloist:", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("Harpist:", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("Death Certificates:", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("Other:", size=(self.TEXT_WIDTH, 1))],
            ]), sg.Column([
                [sg.Push(), sg.Input(key="Clergy Honorarium", size=(40, 1)), sg.Text("$"), sg.Input(key="D1", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Church Honorarium", size=(40, 1)), sg.Text("$"), sg.Input(key="D2", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Altar Servers", size=(40, 1)), sg.Text("$"), sg.Input(key="D3", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Organist", size=(40, 1)), sg.Text("$"), sg.Input(key="D4", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Soloist", size=(40, 1)), sg.Text("$"), sg.Input(key="D5", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Harpist", size=(40, 1)), sg.Text("$"), sg.Input(key="D6", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(),sg.Text("Quantity:"), sg.Input(key="Death_Certificates_Quantity", size=(5, 1), enable_events=True),
                 sg.Text("x $27.00 ="), sg.Push(), sg.Text("$"), sg.Input(key="D7", size=(self.DOLLAR_WIDTH, 1), readonly=True)],
                [sg.Push(), sg.Input(key="Other_4", size=(self.INPUT_WIDTH, 1)), sg.Text("$"), sg.Input(key="D8", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
            ])],
            [sg.Text("TOTAL D:", font=("Helvetica", 10, "bold"), size=(self.TEXT_WIDTH, 1)), sg.Push(), sg.Text("$"), sg.Input(key="Total D", size=(self.DOLLAR_WIDTH, 1), enable_events=True)]
        ]
        
        discount_layout = self.create_discount_layout()
        
        total_charges_layout = [
            # TOTAL CHARGES
            [sg.Text("TOTAL CHARGES", font=("Helvetica", 16, "bold"))],
            [sg.Column([
                [sg.Text("TOTAL A, B, & C SECTIONS:", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("DISCOUNT:", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("G.S.T.:", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("P.S.T.:", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("TOTAL D:", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("GRAND TOTAL:", font=("Helvetica", 10, "bold"), size=(self.TEXT_WIDTH, 1))],
            ]), sg.Column([
                [sg.Push(), sg.Text("$"), sg.Input(key="Total \\(ABC\\)", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Text("$"), sg.Input(key="Discount", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Text("$"), sg.Input(key="GST", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Text("$"), sg.Input(key="PST", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Text("$"), sg.Input(key="Total D_2", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Text("$"), sg.Input(key="Grand Total", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
            ])]
        ]
        
        preplanned_amount_layout = [
            [sg.Text("Preplanned Amount", font=("Helvetica", 16, "bold"))],
            [sg.Column([
                [sg.Text("A. Goods and Services:", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("B. Monument/Marker:", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("C. Other Expenses:", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("D. Final Documents Service:", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("E. Journey Home (Time Pay Only):", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("Total (A+B+C+D+E):", size=(self.TEXT_WIDTH, 1))],
            ]), sg.Column([
                [sg.Push(), sg.Text("$"), sg.Input(key="3A Goods and Services", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Text("$"), sg.Input(key="3B MonumentMarker", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Text("$"), sg.Input(key="3C Other Expenses", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Text("$"), sg.Input(key="3D Final Documents Service", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Text("$"), sg.Input(key="3E Journey Home", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Text("$"), sg.Input(key="Total 3", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
            ])]
        ]
        
        pmt_select_layout = [
            [sg.Text("Payment Selection", font=("Helvetica", 16, "bold"))],
            [sg.Column([
                [sg.Text("Term:", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("A. Single Pay:", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("B. Time Pay:", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("C. Single Pay Journey Home:", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("D. LPR:", size=(self.TEXT_WIDTH, 1))],
                [sg.Text("Total (A+B+C+D):", size=(self.TEXT_WIDTH, 1))],
            ]), sg.Column([
                [sg.Push(), sg.Combo(["3-year", "5-year", "10-year", "15-year", "20-year"], 
                          key="Payment Term", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Text("$"), sg.Input(key="4A Single Pay", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Text("$"), sg.Input(key="4B Time Pay", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Text("$"), sg.Input(key="4C Single Pay Journey Home", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Text("$"), sg.Input(key="4D LPR", size=(self.DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Text("$"), sg.Input(key="Total 4 \\(ABCD\\)", size=(self.DOLLAR_WIDTH, 1), enable_events=True)]
            ])]
        ]
        
        monthly_payments_chart = [
            [sg.Text("Monthly Payment Options", font=("Helvetica", 16, "bold"), justification='center', expand_x=True)],
            [sg.Push(),
             sg.Table(
                values=[['0.00', '0.00', '0.00', '0.00', '0.00']],
                headings=['3-year', '5-year', '10-year', '15-year', '20-year'],
                auto_size_columns=False,
                col_widths=[20, 20, 20, 20, 20],
                justification='center',
                key='-MONTHLY_PAYMENTS_TABLE-',
                num_rows=1,
                row_height=50,
                display_row_numbers=False,
                enable_events=True,
                enable_click_events=True,
                pad=(10,5),
                expand_x=True,
                font=("Helvetica", 16)),
             sg.Push()],
            [sg.Text("*Note: Make sure there is an age value in 'Age' field to calculate the monthly payment", font=("Helvetica", 13, "italic"), justification='center', expand_x=True)]
        ]

        time_pay_layout = [
            [sg.Column(preplanned_amount_layout, expand_x=True)],
            [sg.HorizontalSeparator()],
            [sg.Checkbox("Single Pay Journey Home", key="-SINGLE_PAY_JH-", enable_events=True)],  # Added checkbox
            [sg.Column(pmt_select_layout, expand_x=True)],
            [sg.HorizontalSeparator()],
            [sg.Column(monthly_payments_chart, expand_x=True)]
        ]
        
        other_layout = [
            [sg.Text("Location", font=("Helvetica", 16, "bold"))],
            [sg.Text("Kearney Location:")],
            [sg.Combo(self.kearney_locations, key="Kearney Location", size=(40, 1), enable_events=True)],
            [sg.Text("Location Where Signed:")],
            [sg.Text("City:"), sg.Input(key="Signed City", size=(20, 1)),
             sg.Text("Province:"), sg.Input(key="Signed Province", size=(10, 1))]
        ]
        
        representative_layout = [
            [sg.Text("Representative Information", font=("Helvetica", 16, "bold"))],
            [sg.Text("First Name:"), sg.Input(key="Representative First Name", size=(20, 1))],
            [sg.Text("Middle Name:"), sg.Input(key="Representative Middle Name", size=(20, 1))],
            [sg.Text("Last Name:"), sg.Input(key="Representative Last Name", size=(20, 1))],
            [sg.Text("ID:"), sg.Input(key="Representative ID", size=(20, 1))],
            [sg.Text("Phone:"), sg.Input(key="Representative Phone", size=(20, 1), enable_events=True)],
            [sg.Text("Email:"), sg.Input(key="Representative Email", size=(20, 1), enable_events=True)]
        ]
        
        other_info_layout = [
            [sg.Column(other_layout, expand_x=True)],
            [sg.HorizontalSeparator()],
            [sg.Column(representative_layout, expand_x=True)]
        ]
        
        main_content = [
            [sg.Column([
                [sg.Image(source=resize_image(self.kearney_logo, (300, 80)), size=(300, 80)),
                 sg.Push(),
                 sg.Text("Client Forms", font=("Helvetica", 30, "bold"), justification='center', expand_x=True),
                 sg.Push(),
                 sg.Image(source=resize_image(self.burquitlam_logo, (300, 80)), size=(300, 80))],
            ], expand_x=True, expand_y=True)],
            [sg.HorizontalSeparator()],
            [sg.TabGroup([
                [
                    sg.Tab("Personal Info", [
                        [sg.Frame("Personal Information", personal_info_layout, expand_x=True, expand_y=True)]
                    ], expand_x=True, expand_y=True),
                    
                    sg.Tab("Packages", [
                        [sg.Frame("Packages", [
                            [sg.Text("Select Package:"), sg.Combo(list(self.packages.keys()), key="-PACKAGE-", enable_events=True)],
                            [sg.Text("Type of Service:"), sg.Input(key="Type of Service")],
                            [sg.Column(section_a_layout, scrollable=True, vertical_scroll_only=False, expand_x=True, expand_y=True),
                            sg.VerticalSeparator(),
                            sg.Column([
                                [sg.Frame("Section B", section_b_layout, expand_x=True)],
                                [sg.Frame("Section C", section_c_layout, expand_x=True)],
                                [sg.Frame("Section D", section_d_layout, expand_x=True)],
                                [sg.Frame("Discount", discount_layout, expand_x=True)],
                                [sg.Frame("Total Charges", total_charges_layout, expand_x=True)]
                            ], scrollable=True, vertical_scroll_only=False, expand_x=True, expand_y=True, key='-SECTIONS-COLUMN-')]
                        ], expand_x=True, expand_y=True)]
                    ], expand_x=True, expand_y=True),
                    
                    sg.Tab("Time Pay", [
                        [sg.Frame("Time Pay", time_pay_layout, expand_x=True, expand_y=True)]
                    ], expand_x=True, expand_y=True),
                    
                    sg.Tab("Other Info", [
                        [sg.Frame("Other Information", other_info_layout, expand_x=True, expand_y=True)]
                    ], expand_x=True, expand_y=True)
                ]
            ], expand_x=True)]
        ]

        layout = [
            [sg.Column(
                main_content,
                scrollable=True,
                vertical_scroll_only=False,
                expand_x=True,
                expand_y=True,
                key='-CONTENT-'
            )],
            [sg.HorizontalSeparator()],
            [sg.Column(
                [
                    [sg.Button("Autofill PDFs", expand_x=True),
                    sg.Button("Calculate Monthly Payment", key="-CALCULATE_MONTHLY_PAYMENT-", expand_x=True),
                    sg.Button("Refresh Form", key="-REFRESH-", expand_x=True),
                    sg.Button("Exit", expand_x=True)]
                ],
                expand_x=True,
                element_justification='center',
                key='-BUTTON-COLUMN-',
                pad=(0, 0)
            )]
        ]

        self.window = sg.Window("Kearney Client Forms",
                              layout,
                              resizable=True,
                              finalize=True,
                              margins=(0, 0),
                              icon=self.kearney_icon,
                              size=(800, 600))
        
        # Add the Tkinter binding for dynamic content width
        frame_id = self.window['-CONTENT-'].Widget.frame_id
        canvas = self.window['-CONTENT-'].Widget.canvas
        canvas.bind("<Configure>", lambda event, canvas=canvas, frame_id=frame_id: self.configure_scroll_frame(event, canvas, frame_id))
        
        self.window.set_min_size((800, 600))
        self.window.maximize()
        # Initialize window metadata for tracking discount rows
        self.window.metadata = 0
        self.pdf_paths = self.initialize_pdf_paths()
        
        for key in self.phone_input_keys:
            self.window[key].Widget.bind('<Key>', lambda e, key=key: self.validate_phone_input(key))
            self.last_value[key] = ''
            
        for key in self.dollar_input_keys:
            self.window[key].Widget.bind('<Key>', lambda e, key=key: self.validate_dollar_input(key))
            self.window[key].Widget.bind('<Return>', lambda e, key=key: self.handle_dollar_input(key))
            self.window[key].Widget.bind('<FocusOut>', lambda e, key=key: self.handle_dollar_input(key))
            self.last_value[key] = ''
            
            
            
        try:
            self.window.TKroot.protocol("WM_DELETE_WINDOW", self.on_closing)
        except Exception as e:
            logging.error(f"Error setting up TK root: {str(e)}")
            
    def on_closing(self):
        """Cleanup method for proper window closing"""
        try:
            if hasattr(self, 'check_inactivity_timer') and self.check_inactivity_timer:
                self.window.after_cancel(self.check_inactivity_timer)
            if hasattr(self, 'window'):
                self.window.close()
            sys.exit(0)  # Ensure complete program termination
        except Exception as e:
            logging.error(f"Error during cleanup: {str(e)}")
            sys.exit(1)

    def configure_scroll_frame(self, event, canvas, frame_id):
        """Configure the scrollable frame to expand with window width"""
        canvas.itemconfig(frame_id, width=canvas.winfo_width())


        # Add this list of all package input keys

        self.last_value = {key: '' for key in self.dollar_input_keys}
        locale.setlocale(locale.LC_ALL, '')  # Set the locale to the user's default
        
        self.viewing_listbox = None
        self.limousine_listbox = None
        self.casket_listbox = None
        self.urn_listbox = None
        self.keepsake_listbox = None
        self.crematorium_listbox = None
        self.other_3_listbox = None
        self.reception_facilities_listbox = None
        self.weekend_listbox = None
        
        
        self.last_activity_time = time.time()
        self.check_inactivity_timer = None
        
        self.current_monthly_payments = {}
        
    def calculate_cards(self, quantity):
        try:
            qty = int(quantity)
            total = qty * 2.95
            self.window["B5"].update(f"{total:.2f}")
            # Trigger calculation after updating values
            self.calculate_grand_total(self.get_current_values())
        except ValueError:
            self.window["B5"].update("")
            # Alsorigger calculation after updating values
            self.calculate_grand_total(self.get_current_values())
        
    def calculate_guest_books(self, quantity):
        try:
            qty = int(quantity)
            total = qty * 75.00
            self.window["B6"].update(f"{total:.2f}")
            # Trigger calculation after updating values
            self.calculate_grand_total(self.get_current_values())
        except ValueError:
            self.window["B6"].update("")
            # Alsorigger calculation after updating values
            self.calculate_grand_total(self.get_current_values())
            
    
        
    def handle_single_pay_jh_checkbox(self, values):
        """Handle the Single Pay Journey Home checkbox state change"""
        try:
            updated_values = values.copy()
            
            if values["-SINGLE_PAY_JH-"]:  # If checkbox is checked
                journey_home_value = values["3E Journey Home"]
                if journey_home_value:
                    # Update window and values dictionary
                    self.window["4C Single Pay Journey Home"].update(journey_home_value)
                    self.window["3E Journey Home"].update("")
                    updated_values["4C Single Pay Journey Home"] = journey_home_value
                    updated_values["3E Journey Home"] = ""
            else:  # If checkbox is unchecked
                single_pay_value = values["4C Single Pay Journey Home"]
                if single_pay_value:
                    # Update window and values dictionary
                    self.window["3E Journey Home"].update(single_pay_value)
                    self.window["4C Single Pay Journey Home"].update("")
                    updated_values["3E Journey Home"] = single_pay_value
                    updated_values["4C Single Pay Journey Home"] = ""
                    
            # Calculate totals with updated values
            self.calculate_section_3_total(updated_values)
            self.calculate_section_4_total(updated_values)
            if updated_values.get("-AGE-"):
                self.calculate_monthly_payment(updated_values)
                self.update_monthly_payments(updated_values)
        except Exception as e:
            logging.error(f"Error in handle_single_pay_jh_checkbox: {str(e)}")

    def calculate_section_3_total(self, values):
        """Calculate total for section 3"""
        try:
            total = sum([
                float(values["3A Goods and Services"].replace('$', '').replace(',', '') or 0),
                float(values["3B MonumentMarker"].replace('$', '').replace(',', '') or 0),
                float(values["3C Other Expenses"].replace('$', '').replace(',', '') or 0),
                float(values["3D Final Documents Service"].replace('$', '').replace(',', '') or 0),
                float(values["3E Journey Home"].replace('$', '').replace(',', '') or 0)
            ])
            self.window["Total 3"].update(locale.format_string('%.2f', total, grouping=True))
        except ValueError:
            pass

    def calculate_section_4_total(self, values):
        """Calculate total for section 4"""
        try:
            total = sum([
                float(values["4A Single Pay"].replace('$', '').replace(',', '') or 0),
                float(values["4B Time Pay"].replace('$', '').replace(',', '') or 0),
                float(values["4C Single Pay Journey Home"].replace('$', '').replace(',', '') or 0),
                float(values["4D LPR"].replace('$', '').replace(',', '') or 0)
            ])
            self.window["Total 4 \\(ABCD\\)"].update(locale.format_string('%.2f', total, grouping=True))
        except ValueError:
            pass
        
    def create_floating_listbox(self, for_casket=False, for_keepsake=False, for_viewing=False, for_limousine=False, for_crematorium=False, for_other_3=False, for_reception_facilities=False, for_weekend=False, for_urn=False):
        """Create the floating listbox and bind mouse events"""
        if for_weekend:
            if not self.weekend_listbox:
                input_widget = self.window['Weekend or Statutory Holiday'].Widget
                self.weekend_listbox = tk.Listbox(self.window.TKroot,
                                        width=50,
                                        height=10,
                                        background='white',
                                        foreground='black',
                                        font=('Helvetica', 10))
                self.weekend_listbox.bind('<<ListboxSelect>>',
                                         lambda e: self.on_listbox_select(e, for_weekend=True))
        elif for_reception_facilities:
            if not self.reception_facilities_listbox:
                input_widget = self.window['Reception Facilities'].Widget
                self.reception_facilities_listbox = tk.Listbox(self.window.TKroot,
                                        width=50,
                                        height=10,
                                        background='white',
                                        foreground='black',
                                        font=('Helvetica', 10))
                self.reception_facilities_listbox.bind('<<ListboxSelect>>',
                                         lambda e: self.on_listbox_select(e, for_reception_facilities=True))
        elif for_other_3:
            if not self.other_3_listbox:
                input_widget = self.window['Other_3'].Widget
                self.other_3_listbox = tk.Listbox(self.window.TKroot,
                                        width=50,
                                        height=10,
                                        background='white',
                                        foreground='black',
                                        font=('Helvetica', 10))
                self.other_3_listbox.bind('<<ListboxSelect>>',
                                         lambda e: self.on_listbox_select(e, for_other_3=True))
        elif for_crematorium:
            if not self.crematorium_listbox:
                input_widget = self.window['Crematorium'].Widget
                self.crematorium_listbox = tk.Listbox(self.window.TKroot,
                                        width=50,
                                        height=10,
                                        background='white',
                                        foreground='black',
                                        font=('Helvetica', 10))
                self.crematorium_listbox.bind('<<ListboxSelect>>',
                                         lambda e: self.on_listbox_select(e, for_crematorium=True))
        elif for_viewing:
            if not self.viewing_listbox:
                input_widget = self.window['Evening Prayers or Visitation'].Widget
                self.viewing_listbox = tk.Listbox(self.window.TKroot,
                                        width=50,
                                        height=10,
                                        background='white',
                                        foreground='black',
                                        font=('Helvetica', 10))
                self.viewing_listbox.bind('<<ListboxSelect>>',
                                         lambda e: self.on_listbox_select(e, for_viewing=True)) 
        elif for_limousine:
            if not self.limousine_listbox:
                input_widget = self.window['Limousine'].Widget
                self.limousine_listbox = tk.Listbox(self.window.TKroot,
                                        width=50,
                                        height=10,
                                        background='white',
                                        foreground='black',
                                        font=('Helvetica', 10))
                self.limousine_listbox.bind('<<ListboxSelect>>',
                                         lambda e: self.on_listbox_select(e, for_limousine=True))
        elif for_casket:
            if not self.casket_listbox:
                input_widget = self.window['Casket'].Widget
                self.casket_listbox = tk.Listbox(self.window.TKroot,
                                        width=50,
                                        height=10,
                                        background='white',
                                        foreground='black',
                                        font=('Helvetica', 10))
                self.casket_listbox.bind('<<ListboxSelect>>',
                                         lambda e: self.on_listbox_select(e, for_casket=True))  
        elif for_keepsake:
            if not self.keepsake_listbox:
                input_widget = self.window['Keepsake'].Widget
                self.keepsake_listbox = tk.Listbox(self.window.TKroot,
                                        width=50,
                                        height=10,
                                        background='white',
                                        foreground='black',
                                        font=('Helvetica', 10))
                self.keepsake_listbox.bind('<<ListboxSelect>>', 
                                           lambda e: self.on_listbox_select(e, for_keepsake=True))
        elif for_urn:  # Changed from else to elif
            if not self.urn_listbox:
                input_widget = self.window['Urn'].Widget
                self.urn_listbox = tk.Listbox(self.window.TKroot,
                                        width=50,
                                        height=10,
                                        background='white',
                                        foreground='black',
                                        font=('Helvetica', 10))
                self.urn_listbox.bind('<<ListboxSelect>>', 
                                     lambda e: self.on_listbox_select(e, for_urn=True))
                
                # Add focus-out binding to input fields
        if for_weekend:
            self.window['Weekend or Statutory Holiday'].Widget.bind('<FocusOut>', 
                lambda e: self.hide_listbox(for_weekend=True))
        elif for_reception_facilities:
            self.window['Reception Facilities'].Widget.bind('<FocusOut>', 
                lambda e: self.hide_listbox(for_reception_facilities=True))
        elif for_other_3:
            self.window['Other_3'].Widget.bind('<FocusOut>', 
                lambda e: self.hide_listbox(for_other_3=True))
        elif for_crematorium:
            self.window['Crematorium'].Widget.bind('<FocusOut>', 
                lambda e: self.hide_listbox(for_crematorium=True))
        elif for_viewing:
            self.window['Evening Prayers or Visitation'].Widget.bind('<FocusOut>', 
                lambda e: self.hide_listbox(for_viewing=True))
        elif for_limousine:
            self.window['Limousine'].Widget.bind('<FocusOut>', 
                lambda e: self.hide_listbox(for_limousine=True))
        elif for_casket:
            self.window['Casket'].Widget.bind('<FocusOut>', 
                lambda e: self.hide_listbox(for_casket=True))
        elif for_keepsake:
            self.window['Keepsake'].Widget.bind('<FocusOut>', 
                lambda e: self.hide_listbox(for_keepsake=True))
        elif for_urn:
            self.window['Urn'].Widget.bind('<FocusOut>', 
                lambda e: self.hide_listbox(for_urn=True))
                
        # Add mouse bindings for activity tracking
        if for_weekend and not self.weekend_listbox:
            self.weekend_listbox.bind('<Motion>', self.on_mouse_activity)
            self.weekend_listbox.bind('<Button-4>', self.on_mouse_activity)  # Mouse wheel up
            self.weekend_listbox.bind('<Button-5>', self.on_mouse_activity)  # Mouse wheel down
            self.weekend_listbox.bind('<MouseWheel>', self.on_mouse_activity)  # Windows mouse wheel
            self.weekend_listbox.bind('<B1-Motion>', self.on_mouse_activity)  # Drag with mouse button
        elif for_reception_facilities and not self.reception_facilities_listbox:
            self.reception_facilities_listbox.bind('<Motion>', self.on_mouse_activity)
            self.reception_facilities_listbox.bind('<Button-4>', self.on_mouse_activity)
            self.reception_facilities_listbox.bind('<Button-5>', self.on_mouse_activity)
            self.reception_facilities_listbox.bind('<MouseWheel>', self.on_mouse_activity)
            self.reception_facilities_listbox.bind('<B1-Motion>', self.on_mouse_activity)
        elif for_other_3 and not self.other_3_listbox:
            self.other_3_listbox.bind('<Motion>', self.on_mouse_activity)
            self.other_3_listbox.bind('<Button-4>', self.on_mouse_activity)
            self.other_3_listbox.bind('<Button-5>', self.on_mouse_activity)
            self.other_3_listbox.bind('<MouseWheel>', self.on_mouse_activity)
            self.other_3_listbox.bind('<B1-Motion>', self.on_mouse_activity)
        elif for_crematorium and not self.crematorium_listbox:
            self.crematorium_listbox.bind('<Motion>', self.on_mouse_activity)
            self.crematorium_listbox.bind('<Button-4>', self.on_mouse_activity)
            self.crematorium_listbox.bind('<Button-5>', self.on_mouse_activity)
            self.crematorium_listbox.bind('<MouseWheel>', self.on_mouse_activity)
            self.crematorium_listbox.bind('<B1-Motion>', self.on_mouse_activity)
        elif for_viewing and not self.viewing_listbox:
            self.viewing_listbox.bind('<Motion>', self.on_mouse_activity)
            self.viewing_listbox.bind('<Button-4>', self.on_mouse_activity)
            self.viewing_listbox.bind('<Button-5>', self.on_mouse_activity)
            self.viewing_listbox.bind('<MouseWheel>', self.on_mouse_activity)
            self.viewing_listbox.bind('<B1-Motion>', self.on_mouse_activity)
        elif for_limousine and not self.limousine_listbox:
            self.limousine_listbox.bind('<Motion>', self.on_mouse_activity)
            self.limousine_listbox.bind('<Button-4>', self.on_mouse_activity)
            self.limousine_listbox.bind('<Button-5>', self.on_mouse_activity)
            self.limousine_listbox.bind('<MouseWheel>', self.on_mouse_activity)
            self.limousine_listbox.bind('<B1-Motion>', self.on_mouse_activity)
        elif for_casket and not self.casket_listbox:
            self.casket_listbox.bind('<Motion>', self.on_mouse_activity)
            self.casket_listbox.bind('<Button-4>', self.on_mouse_activity)
            self.casket_listbox.bind('<Button-5>', self.on_mouse_activity)
            self.casket_listbox.bind('<MouseWheel>', self.on_mouse_activity)
            self.casket_listbox.bind('<B1-Motion>', self.on_mouse_activity)
        elif for_keepsake and not self.keepsake_listbox:
            self.keepsake_listbox.bind('<Motion>', self.on_mouse_activity)
            self.keepsake_listbox.bind('<Button-4>', self.on_mouse_activity)
            self.keepsake_listbox.bind('<Button-5>', self.on_mouse_activity)
            self.keepsake_listbox.bind('<MouseWheel>', self.on_mouse_activity)
            self.keepsake_listbox.bind('<B1-Motion>', self.on_mouse_activity)
        elif for_urn and not self.urn_listbox:
            self.urn_listbox.bind('<Motion>', self.on_mouse_activity)
            self.urn_listbox.bind('<Button-4>', self.on_mouse_activity)
            self.urn_listbox.bind('<Button-5>', self.on_mouse_activity)
            self.urn_listbox.bind('<MouseWheel>', self.on_mouse_activity)
            self.urn_listbox.bind('<B1-Motion>', self.on_mouse_activity)

    def on_mouse_activity(self, event):
        """Handle mouse movement and scrolling"""
        self.last_activity_time = time.time()
        return "break"  # Prevent event from propagating

   
    def handle_input(self, event, values, for_keepsake=False, for_casket=False, for_viewing=False, for_limousine=False, for_crematorium=False, for_other_3=False, for_reception_facilities=False, for_weekend=False, for_urn=False):
        """Handle input events and reset inactivity timer"""
        self.create_floating_listbox(for_casket, for_keepsake, for_viewing, for_limousine, for_crematorium, for_other_3, for_reception_facilities, for_weekend, for_urn)
        
        # Update last activity time
        self.last_activity_time = time.time()
        
        if for_weekend:
            search_text = values['Weekend or Statutory Holiday'].lower()
            listbox = self.weekend_listbox
            items_dict = self.weekend
            input_key = 'Weekend or Statutory Holiday'
        elif for_reception_facilities:
            search_text = values['Reception Facilities'].lower()
            listbox = self.reception_facilities_listbox
            items_dict = self.reception_facilities
            input_key = 'Reception Facilities'
        elif for_other_3:
            search_text = values['Other_3'].lower()
            listbox = self.other_3_listbox
            items_dict = self.other_3
            input_key = 'Other_3'
        elif for_crematorium:
            search_text = values['Crematorium'].lower()
            listbox = self.crematorium_listbox
            items_dict = self.crematorium
            input_key = 'Crematorium'
        elif for_limousine:
            search_text = values['Limousine'].lower()
            listbox = self.limousine_listbox
            items_dict = self.limousines
            input_key = 'Limousine'
        elif for_viewing:
            search_text = values['Evening Prayers or Visitation'].lower()
            listbox = self.viewing_listbox
            items_dict = self.viewings
            input_key = 'Evening Prayers or Visitation'
        elif for_casket:
            search_text = values['Casket'].lower()
            listbox = self.casket_listbox
            items_dict = self.caskets
            input_key = 'Casket'
        elif for_keepsake:
            search_text = values['Keepsake'].lower()
            listbox = self.keepsake_listbox
            items_dict = self.urns
            input_key = 'Keepsake'
        elif for_urn:
            search_text = values['Urn'].lower()
            listbox = self.urn_listbox
            items_dict = self.urns
            input_key = 'Urn'
        else:
            return
            
        # Check if listbox exists before trying to delete
        if listbox is None:
            return
            
        # Filter and update listbox
        try:
            listbox.delete(0, tk.END)
            filtered_items = [item for item in items_dict.keys() 
                            if search_text in item.lower()]
            
            # Calculate appropriate height (1 to 10 items)
            num_items = len(filtered_items)

            display_height = min(max(num_items, 1), 8.75)
            
            # Insert filtered items
            for item in filtered_items:
                listbox.insert(tk.END, item)
            
            # Position and show listbox
            input_widget = self.window[input_key].Widget
            x = input_widget.winfo_rootx() - self.window.TKroot.winfo_rootx()
            y = input_widget.winfo_rooty() - self.window.TKroot.winfo_rooty() + input_widget.winfo_height()
            
            # Calculate height in pixels (approximately 20 pixels per item)
            total_height = display_height * 20
            
            listbox.place(x=x, y=y, width=400, height=total_height)
            
            # Start or reset the inactivity checker
            self.start_inactivity_checker(
                for_casket=for_casket,
                for_keepsake=for_keepsake,
                for_viewing=for_viewing,
                for_limousine=for_limousine,
                for_crematorium=for_crematorium,
                for_other_3=for_other_3,
                for_reception_facilities=for_reception_facilities,
                for_weekend=for_weekend,
                for_urn=for_urn
            )
            
        except Exception as e:
            logging.error(f"Error updating listbox: {str(e)}")
            
    def hide_listbox(self, for_casket=False, for_keepsake=False, for_viewing=False, for_urn=False, for_limousine=False, for_crematorium=False, for_other_3=False, for_reception_facilities=False, for_weekend=False):
        """Hide the floating listbox"""
        if for_weekend:
            listbox = self.weekend_listbox
        elif for_reception_facilities:
            listbox = self.reception_facilities_listbox
        elif for_other_3:
            listbox = self.other_3_listbox
        elif for_crematorium:
            listbox = self.crematorium_listbox
        elif for_viewing:
            listbox = self.viewing_listbox
        elif for_limousine:
            listbox = self.limousine_listbox
        elif for_casket:
            listbox = self.casket_listbox
        elif for_keepsake:
            listbox = self.keepsake_listbox
        elif for_urn:
            listbox = self.urn_listbox
        else:
            return
            
        if listbox:
            listbox.place_forget()

    def start_inactivity_checker(self, **kwargs):
        """Start or reset the inactivity checker"""
        # Cancel existing timer if any
        if self.check_inactivity_timer:
            self.window.TKroot.after_cancel(self.check_inactivity_timer)
        
        # Start new timer
        self.check_inactivity_timer = self.window.TKroot.after(100, lambda: self.check_inactivity(**kwargs))

    def check_inactivity(self, **kwargs):
        """Check for inactivity and hide listbox if inactive for 3 seconds"""
        current_time = time.time()
        if current_time - self.last_activity_time >= 7:
            self.hide_listbox(**kwargs)
        else:
            # Continue checking if not inactive long enough
            self.start_inactivity_checker(**kwargs)


    def on_listbox_select(self, event, for_casket=False, for_keepsake=False, for_viewing=False, for_urn=False, for_limousine=False, for_crematorium=False, for_other_3=False, for_reception_facilities=False, for_weekend=False):
        """Handle listbox selection"""
        if for_weekend:
            listbox = self.weekend_listbox
        elif for_reception_facilities:
            listbox = self.reception_facilities_listbox
        elif for_other_3:
            listbox = self.other_3_listbox
        elif for_crematorium:
            listbox = self.crematorium_listbox
        elif for_viewing:
            listbox = self.viewing_listbox
        elif for_limousine:
            listbox = self.limousine_listbox
        elif for_casket:
            listbox = self.casket_listbox
        elif for_keepsake:
            listbox = self.keepsake_listbox
        elif for_urn:
            listbox = self.urn_listbox
        else:
            return

        selection = listbox.curselection()
        if selection:
            selected_item = listbox.get(selection[0])
            if for_weekend:
                self.window['Weekend or Statutory Holiday'].update(selected_item)
                price = float(self.weekend[selected_item])
                self.window["A7"].update(f"{price:.2f}")
                self.format_dollar_field("A7", f"{price:.2f}")
            if for_reception_facilities:
                self.window['Reception Facilities'].update(selected_item)
                price = float(self.reception_facilities[selected_item])
                self.window["A8"].update(f"{price:.2f}")
                self.format_dollar_field("A8", f"{price:.2f}")
            elif for_other_3:
                self.window['Other_3'].update(selected_item)
                price = float(self.other_3[selected_item])
                self.window["C10"].update(f"{price:.2f}")
                self.format_dollar_field("C10", f"{price:.2f}")
            elif for_crematorium:
                self.window['Crematorium'].update(selected_item)
                price = float(self.crematorium[selected_item])
                self.window["C2"].update(f"{price:.2f}")
                self.format_dollar_field("C2", f"{price:.2f}")
            elif for_viewing:
                self.window['Evening Prayers or Visitation'].update(selected_item)
                price = float(self.viewings[selected_item])
                self.window["A6"].update(f"{price:.2f}")
                self.format_dollar_field("A6", f"{price:.2f}")
            elif for_limousine:
                self.window['Limousine'].update(selected_item)
                price = float(self.limousines[selected_item])
                self.window["A9F"].update(f"{price:.2f}")
                self.format_dollar_field("A9F", f"{price:.2f}")
            elif for_casket:
                self.window['Casket'].update(selected_item)
                price = float(self.caskets[selected_item])
                self.window["B1"].update(f"{price:.2f}")
                self.format_dollar_field("B1", f"{price:.2f}")
            elif for_keepsake:
                self.window['Keepsake'].update(selected_item)
                price = float(self.urns[selected_item])
                self.window["B3"].update(f"{price:.2f}")
                self.format_dollar_field("B3", f"{price:.2f}")
            elif for_urn:
                self.window['Urn'].update(selected_item)
                price = float(self.urns[selected_item])
                self.window["B2"].update(f"{price:.2f}")
                self.format_dollar_field("B2", f"{price:.2f}")
            
            # Trigger calculation after updating values
            self.calculate_grand_total(self.get_current_values())
            
            self.hide_listbox(
                for_casket=for_casket,
                for_keepsake=for_keepsake,
                for_viewing=for_viewing,
                for_urn=for_urn,
                for_limousine=for_limousine,
                for_crematorium=for_crematorium,
                for_other_3=for_other_3,
                for_reception_facilities=for_reception_facilities,
                for_weekend=for_weekend
            )
        else:
            self.last_activity_time = time.time()
         
              
    def ordinal(self, n):
        if 10 <= n % 100 <= 20:
            suffix = 'th'
        else:
            suffix = {1: 'st', 2: 'nd', 3: 'rd'}.get(n % 10, 'th')
        return f"{n}{suffix}"
            
    def apply_package(self, package_name):
        """Apply the selected package to the form"""
        try:
            logging.info(f"Starting apply_package for: {package_name}")
            
            # Clear all fields that might be affected by packages
            fields_to_clear = set()
            for package in self.packages.values():
                fields_to_clear.update(package.keys())
            for field in fields_to_clear:
                if field in self.window.AllKeysDict:
                    self.window[field].update('')
                    logging.info(f"Cleared field {field}")
            
            # Store the selected package name
            self.selected_package = package_name
            # Update the 'Type of Service' field
            self.window['Type of Service'].update(package_name)
            package_data = self.packages[package_name]
            
            # Check if Cadence discount already exists
            cadence_exists = False
            empty_discount_row = None
            
            # Look for existing Cadence discount and empty discount rows
            for i in range(getattr(self.window, 'metadata', 0) + 1):
                if self.window[('-DISCOUNT-ROW-', i)].visible:
                    desc = self.window[('-DISCOUNT-DESC-', i)].get()
                    amt = self.window[('-DISCOUNT-AMT-', i)].get()
                    
                    if desc == "Cadence":
                        cadence_exists = True
                        break
                    elif not desc and not amt:
                        empty_discount_row = i
            
            # Add Cadence discount if it doesn't exist
            if not cadence_exists:
                if empty_discount_row is not None:
                    # Use existing empty row
                    self.window[('-DISCOUNT-DESC-', empty_discount_row)].update("Cadence")
                    self.window[('-DISCOUNT-AMT-', empty_discount_row)].update("400.00")
                    self.handle_dollar_input({('-DISCOUNT-AMT-', empty_discount_row): "400.00"})
                else:
                    # Create new row if no empty one exists
                    new_item_num = getattr(self.window, 'metadata', 0) + 1
                    self.add_discount_field()
                    
                    self.window[('-DISCOUNT-DESC-', new_item_num)].update("Cadence")
                    self.window[('-DISCOUNT-AMT-', new_item_num)].update("400.00")
                    self.handle_dollar_input({('-DISCOUNT-AMT-', new_item_num): "400.00"})
                    
                    self.window[('-DISCOUNT-ROW-', new_item_num)].update(visible=True)
            
            # Apply package data to the form
            for field, value in package_data.items():
                if field in self.window.AllKeysDict:
                    if field in self.dollar_input_keys:
                        self.window[field].update(f"{float(value):,.2f}")
                        self.format_dollar_field(field)
                    else:
                        self.window[field].update(value)
                    logging.info(f"Updated field {field} with value {value}")
            
            # Recalculate totals after applying the package
            data = self.get_current_values()
            self.calculate_grand_total(data)
            logging.info(f"Applied package: {package_name}")
            
        except Exception as e:
            logging.error(f"Error applying package {package_name}: {str(e)}")
            sg.popup_error(f"An error occurred while applying the package: {str(e)}")

    def get_base_path(self):
        if getattr(sys, 'frozen', False):
            # Running as compiled executable
            return sys._MEIPASS
        else:
            # Running as a normal Python script
            return os.path.dirname(os.path.abspath(__file__))

    def setup_logging(self):
        log_dir = os.path.join(self.get_base_path(), "Logs")
        os.makedirs(log_dir, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_file = os.path.join(log_dir, f"pdf_autofill_log_{timestamp}.txt")

        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_file),
                logging.StreamHandler()
            ]
        )
        logging.info("Logging initialized")

    def initialize_pdf_paths(self):
        return {
            1: os.path.join(self.base_path, "Forms", PDFTypes.TRUSTAGE_APPLICATION),
            2: os.path.join(self.base_path, "Forms", PDFTypes.PERSONAL_INFO_SHEET),
            3: os.path.join(self.base_path, "Forms", PDFTypes.INSTRUCTIONS),
            4: os.path.join(self.base_path, "Forms", PDFTypes.PRE_ARRANGED_FULL_SERVICE),
            5: os.path.join(self.base_path, "Forms", PDFTypes.JOURNEY_HOME)
        }
        
    def get_current_values(self):
        """Get current values from all input elements in the window."""
        return {
            key: elem.get() 
            for key, elem in self.window.key_dict.items() 
            if hasattr(elem, 'get')
        }
        
    def handle_dollar_input(self, key):
        """Handle dollar input formatting"""
        try:
            self.format_dollar_field(key)
        
            # Check if age exists and trigger monthly payment calculation
            current_values = self.get_current_values()
            if current_values.get("-AGE-"):
                self.calculate_monthly_payment(current_values)
                self.update_monthly_payments(current_values)
        
        except Exception as e:
            logging.error(f"Error handling dollar input: {str(e)}")
            
    def convert_to_float(self, value):
        """Convert formatted dollar string to float"""
        try:
            if not value:
                return ''
            # Remove currency symbol, commas and spaces
            cleaned = value.replace('$', '').replace(',', '').strip()
            # Convert to float
            return f"{float(cleaned):.2f}"
        except (ValueError, AttributeError):
            return ''
    
    def discount_row(self, item_num):
        """
        Creates a single discount row with a description input, amount input, and delete button
        """
        row = [sg.pin(sg.Col([[
            sg.Input(size=(40, 1), key=('-DISCOUNT-DESC-', item_num), enable_events=True),
            sg.Text("$"),
            sg.Input(size=(self.DOLLAR_WIDTH, 1), key=('-DISCOUNT-AMT-', item_num), enable_events=True),
            sg.Button("×", key=('-DEL-DISCOUNT-', item_num), 
                     size=(2, 1), 
                     button_color=('black', 'lightgray'),
                     tooltip='Delete this discount')
        ]], key=('-DISCOUNT-ROW-', item_num)))]
        return row

    def create_discount_layout(self):
        """
        Creates the initial discount section layout
        """
        return [
            [sg.Text("Discounts", font=("Helvetica", 16, "bold"))],
            [sg.Button("+", key="-ADD-DISCOUNT-", size=(2, 1)), 
             sg.Text("Add Discount", font=("Helvetica", 10))],
            [sg.Column([self.discount_row(0)], key='-DISCOUNT-SECTION-', expand_x=True, expand_y=True)]
        ]   
    
    def update_scroll_region(self):
        """Update the scroll region of the scrollable area"""
        try:
            # Get the sections column widget
            sections_column = self.window['-SECTIONS-COLUMN-'].Widget
            
            # Force geometry update
            sections_column.update_idletasks()
            
            # Get the canvas (first child of the widget)
            canvas = sections_column.winfo_children()[0]
            
            # Get the frame inside the canvas
            frame = canvas.winfo_children()[0]
            
            # Force frame to update its geometry
            frame.update_idletasks()
            
            # Calculate new scroll region based on frame size
            canvas.configure(scrollregion=canvas.bbox('all'))
            
            # Scroll to the bottom
            canvas.yview_moveto(1.0)
            
            logging.info("Scroll region updated successfully")
            
        except Exception as e:
            logging.error(f"Error updating scroll region: {str(e)}")

    def add_discount_field(self):
        """Add a new discount field row"""
        try:
            # Increment counter stored in window metadata
            self.window.metadata = getattr(self.window, 'metadata', 0) + 1
            item_num = self.window.metadata
            
            # Add new row to layout
            self.window.extend_layout(self.window['-DISCOUNT-SECTION-'], 
                                    [self.discount_row(item_num)])
            
            # Use the same tuple format as the initial key
            amount_key = ('-DISCOUNT-AMT-', item_num)
            
            # Add to dollar_input_keys if not already there
            if amount_key not in self.dollar_input_keys:
                self.dollar_input_keys.append(amount_key)
                # Add the same event bindings as in __init__
                self.window[amount_key].Widget.bind('<Key>', lambda e, key=amount_key: self.validate_dollar_input(key))
                self.window[amount_key].Widget.bind('<Return>', lambda e, key=amount_key: self.format_dollar_field(key))
                self.window[amount_key].Widget.bind('<FocusOut>', lambda e, key=amount_key: self.format_dollar_field(key))
                    
            # Initialize last_value for this key
            if amount_key not in self.last_value:
                self.last_value[amount_key] = ''
            
            # Force window to process all pending events
            self.window.refresh()
            
            # Add a small delay before updating scroll region
            self.window.TKroot.after(100, self.update_scroll_region)
            
            logging.info(f"Successfully added discount field {item_num}")
            
        except Exception as e:
            logging.error(f"Error adding discount field: {str(e)}")


    def remove_discount_field(self, item_num):
        """Remove a discount field row or clear it if it's the last one"""
        try:
            # Count total visible discount rows
            visible_rows = sum(1 for i in range(getattr(self.window, 'metadata', 0) + 1) 
                        if self.window[('-DISCOUNT-ROW-', i)].visible)
        
            if visible_rows <= 1:
                # If this is the last visible row, just clear its values
                self.window[('-DISCOUNT-DESC-', item_num)].update('')
                self.window[('-DISCOUNT-AMT-', item_num)].update('')
                logging.info(f"Cleared discount field {item_num} (last remaining row)")
            else:
                # Hide the row and clear its values
                self.window[('-DISCOUNT-ROW-', item_num)].update(visible=False)
                self.window[('-DISCOUNT-DESC-', item_num)].update('')
                self.window[('-DISCOUNT-AMT-', item_num)].update('')
            
                # Remove tracking for amount field
                amount_key = ('-DISCOUNT-AMT-', item_num)
                if amount_key in self.dollar_input_keys:
                    self.dollar_input_keys.remove(amount_key)
                if amount_key in self.discount_amount_keys:
                    self.discount_amount_keys.remove(amount_key)
                if amount_key in self.last_value:
                    del self.last_value[amount_key]
            
                logging.info(f"Successfully removed discount field {item_num}")
        
            # Recalculate totals
            self.calculate_grand_total(self.get_current_values())
            
            # Force window to process all pending events
            self.window.refresh()
            
            # Add a small delay before updating scroll region
            self.window.TKroot.after(100, self.update_scroll_region)
        
        except Exception as e:
            logging.error(f"Error handling discount field {item_num}: {str(e)}")

    
    def calculate_total_discount(self, values):
        """Calculate the total discount from all discount fields"""
        try:
            total_discount = 0
            
            # Iterate through all possible discount rows
            for i in range(self.window.metadata + 1):
                row_key = ('-DISCOUNT-ROW-', i)
                amount_key = ('-DISCOUNT-AMT-', i)
                
                # Only process visible rows
                if self.window[row_key].visible:
                    amount = values.get(amount_key, '0')
                    try:
                        amount = amount.replace('$', '').replace(',', '').strip()
                        total_discount += float(amount or 0)
                    except (ValueError, AttributeError):
                        logging.warning(f"Invalid discount amount: {amount}")
                        continue
            
            # Update the main discount field
            formatted_discount = locale.format_string('%.2f', total_discount, grouping=True)
            self.window["Discount"].update(formatted_discount)
            logging.info(f"Total discount calculated: {formatted_discount}")
            return total_discount
            
        except Exception as e:
            logging.error(f"Error calculating total discount: {str(e)}")
            return 0


    def run(self):
        while True:
            try:
                event, values = self.window.read()
                if event == sg.WINDOW_CLOSED or event == "Exit":
                    break
                elif event == "-REFRESH-":
                    self.refresh_form()
                elif event == "-SINGLE_PAY_JH-":
                    if values.get("-AGE-"):
                        self.handle_single_pay_jh_checkbox(values)
                        self.handle_dollar_input(values)
                    else:
                        self.handle_single_pay_jh_checkbox(values)
                elif event in self.dollar_input_keys and values.get("-AGE-"):
                    self.handle_dollar_input(values)
                elif event in ["4A Single Pay", "4C Single Pay Journey Home", "4D LPR"]:
                    # First handle the normal dollar input formatting
                    if event in self.dollar_input_keys:
                        self.handle_dollar_input(values)
                        # Check if age field is empty
                        if values.get("-AGE-"):
                            # If age exists, calculate section 4 total and monthly payment
                            self.calculate_section_4_total(values)
                            self.calculate_monthly_payment(values)
                            self.update_monthly_payments(values)
                            # If a payment term is already selected, update it with new calculation
                            selected_term = values.get("Payment Term")
                            if selected_term:
                                self.handle_payment_term_selection(selected_term)
                        else:
                            # If no age exists, do full monthly payment calculation
                            self.calculate_section_4_total(values)
                elif event in ["3A Goods and Services", "3B MonumentMarker", 
                            "3C Other Expenses", "3D Final Documents Service", 
                            "3E Journey Home"]:
                    # Handle dollar formatting
                    if event in self.dollar_input_keys:
                        self.handle_dollar_input(values)
                    # Calculate section 3 total
                    self.calculate_section_3_total(values)
                elif event == "Weekend or Statutory Holiday":
                    self.handle_input(event, values, for_weekend=True)
                elif event == "Reception Facilities":
                    self.handle_input(event, values, for_reception_facilities=True)
                elif event == "Other_3":
                    self.handle_input(event, values, for_other_3=True)
                elif event == "Crematorium":
                    self.handle_input(event, values, for_crematorium=True)
                elif event == "Limousine":
                    self.handle_input(event, values, for_limousine=True)
                elif event == "Evening Prayers or Visitation":
                    self.handle_input(event, values, for_viewing=True)
                elif event =='Casket':
                    self.handle_input(event, values, for_casket=True)
                elif event == 'Urn':
                    self.handle_input(event, values, for_urn=True)
                elif event == 'Keepsake':
                    self.handle_input(event, values, for_keepsake=True)
                elif event == "Kearney Location":
                    self.update_establishment_constants(values["Kearney Location"])
                elif event == "-PACKAGE-":
                    logging.info(f"Selected package: {values['-PACKAGE-']}")
                    self.apply_package(values["-PACKAGE-"])
                    self.calculate_grand_total(self.get_current_values())
                elif event == "-BIRTHDATE-":
                    self.update_age(values["-BIRTHDATE-"])
                elif event == "-CALCULATE_MONTHLY_PAYMENT-":
                    self.calculate_monthly_payment(values)
                    self.update_monthly_payments(values)
                elif event == "SIN":
                    self.window["SIN"].update(self.format_sin(values["SIN"]))
                elif event in self.phone_input_keys:
                    self.window[event].update(self.format_phone_number(values[event]))
                elif event in self.postal_input_keys:
                    self.window[event].update(self.format_postal_code(values[event]))
                elif event == "-SAME_ADDRESS-":
                    self.toggle_beneficiary_address_fields(values["-SAME_ADDRESS-"])
                elif event == "Payment Term":
                    self.handle_payment_term_selection(values["Payment Term"])
                    logging.info(f"Selected payment term: {values['Payment Term']}")
                    self.calculate_monthly_payment(values)
                    self.update_monthly_payments(values)
                elif event == "Cards_Qty":
                    self.calculate_cards(values["Cards_Qty"])
                elif event == "Guest_Book_Qty":
                    self.calculate_guest_books(values["Guest_Book_Qty"])
                elif event == "Death_Certificates_Quantity":
                    self.calculate_death_certificates(values["Death_Certificates_Quantity"])
                elif event == "Autofill PDFs":
                    self.autofill_pdfs(values)
                elif event == "4A Single Pay":
                    # First handle the normal dollar input formatting
                    if event in self.dollar_input_keys:
                        self.handle_dollar_input(values)
                    self.calculate_section_4_total(values)
                # New discount field handling
                elif event == "-ADD-DISCOUNT-":
                    self.add_discount_field()
                elif isinstance(event, tuple) and event[0] == '-DEL-DISCOUNT-':
                    self.remove_discount_field(event[1])
                elif isinstance(event, tuple) and event[0] == '-DISCOUNT-AMT-':
                    # Handle changes in discount amount fields
                    self.handle_dollar_input(values)
                    self.calculate_grand_total(self.get_current_values())
                elif event == '-CLEANUP-WIDGETS-':
                    field_to_cleanup = values[event]
                    try:
                        # Cleanup widgets after form refresh
                        for element in field_to_cleanup['row']:
                            try:
                                if hasattr(element, 'Widget') and element.Widget:
                                    element.Widget.destroy()
                                elif hasattr(element, 'TKButton') and element.TKButton:
                                    element.TKButton.destroy()
                            except tk.TclError:
                                pass
                    except Exception as e:
                        logging.warning(f"Error during widget cleanup: {str(e)}")

            except Exception as e:
                logging.error(f"Error in main event loop: {str(e)}")
                sg.popup_error(f"An unexpected error occurred: {str(e)}")

        self.window.close()

    def refresh_form(self):
        """Refresh all form fields by clearing them"""
        try:
            logging.info("Starting form refresh")
            
            # Remove all discount rows except the first one using remove_discount_field
            for i in range(1, getattr(self.window, 'metadata', 0) + 1):  # Start from 1 to keep first row
                if self.window[('-DISCOUNT-ROW-', i)].visible:
                    self.remove_discount_field(i)
                    
            # Clear the first row's values but keep it visible
            self.window[('-DISCOUNT-DESC-', 0)].update('')
            self.window[('-DISCOUNT-AMT-', 0)].update('')            
            
            # Clear all input fields
            for key in self.window.AllKeysDict:
                try:
                    # Skip all discount-related fields (we handle them separately)
                    if isinstance(key, tuple) and key[0] in ['-DISCOUNT-DESC-', '-DISCOUNT-AMT-', '-DEL-DISCOUNT-', '-DISCOUNT-ROW-']:
                        continue
                    widget = self.window[key]
                    if isinstance(widget, sg.Input) or isinstance(widget, sg.Multiline):
                        widget.update('')
                    elif isinstance(widget, sg.Combo):
                        widget.update(value='')
                    elif isinstance(widget, sg.Checkbox):
                        widget.update(value=False)
                except Exception as e:
                    logging.warning(f"Could not clear widget {key}: {str(e)}")
            
            # Reset selections
            self.window["-PACKAGE-"].update(value='')
            self.selected_package = ""
            self.window["Kearney Location"].update(value='')
            
            # Clear calculated fields
            for field in ["Total A", "Total B", "Total C", "Total D", "Total \\(ABC\\)", 
                        "Discount", "GST", "PST", "Total D_2", "Grand Total"]:
                try:
                    self.window[field].update('')
                except Exception as e:
                    continue

            # Reset payment selection fields
            self.window["Payment Term"].update(value='')
            for field in ["4A Single Pay", "4B Time Pay", "4C Single Pay Journey Home", 
                        "4D LPR", "Total 4 \\(ABCD\\)"]:
                try:
                    self.window[field].update('')
                except Exception as e:
                    continue
                    
            # Clear the monthly payments table
            self.window['-MONTHLY_PAYMENTS_TABLE-'].update(values=[['' for _ in range(5)]])
            
            # Reset current_monthly_payments dictionary
            self.current_monthly_payments = {
                '3-year': None,
                '5-year': None,
                '10-year': None,
                '15-year': None,
                '20-year': None
            }

            # Force window refresh at the end
            self.window.refresh()
            logging.info("Form refreshed successfully")
            sg.popup("Form has been refreshed. You can start a new entry.")
            
        except Exception as e:
            logging.error(f"Error refreshing form: {str(e)}")
            sg.popup_error(f"An error occurred while refreshing the form: {str(e)}")
            
            
    def calculate_death_certificates(self, quantity):
        try:
            qty = int(quantity)
            total = qty * 27.00
            self.window["D7"].update(f"{total:.2f}")
            # Trigger calculation after updating values
            self.calculate_grand_total(self.get_current_values())
        except ValueError:
            self.window["D7"].update("")
            # Alsorigger calculation after updating values
            self.calculate_grand_total(self.get_current_values())
    
    def calculate_preplanned_amount(self, values):
        try:
            # Step 1: Calculate Preplanned Amount
            goods_and_services = float(values["3A Goods and Services"].replace('$', '').replace(',', '') or 0)
            monument_marker = float(values["3B MonumentMarker"].replace('$', '').replace(',', '') or 0)
            other_expenses = float(values["3C Other Expenses"].replace('$', '').replace(',', '') or 0)
            final_documents = float(values["3D Final Documents Service"].replace('$', '').replace(',', '') or 0)
            journey_home = float(values["3E Journey Home"].replace('$', '').replace(',', '') or 0)
            total_preplanned = goods_and_services + monument_marker + other_expenses + final_documents + journey_home
            
            
            self.window["Total 3"].update(locale.format_string('%.2f', total_preplanned, grouping=True))
            
            self.window.refresh()
            logging.info("Preplanned Amount calculation completed successfully")

        except Exception as e:
            logging.error(f"Error in preplanned amount calculation: {str(e)}")
            sg.popup_error(f"An error occurred during calculation: {str(e)}")
        
    def calculate_monthly_payment(self, values):
        """Calculate monthly payments for different terms"""
        terms = ['3-year', '5-year', '10-year', '15-year', '20-year']
        
        try:
            # Get the total preplanned amount and single pay amount
            total_preplanned = float(values["Total 3"].replace('$', '').replace(',', '') or 0)
            single_pay = float(values["4A Single Pay"].replace('$', '').replace(',', '') or 0)
            
            # Calculate the amount to be financed
            total_to_finance = math.ceil(total_preplanned - single_pay)
            
            if total_to_finance <= 0:
                logging.info("Amount to finance is zero or negative")
                self.window['-MONTHLY_PAYMENTS_TABLE-'].update(values=[['0.00'] * 5])
                self.current_monthly_payments = {
                    '3-year': 0.00,
                    '5-year': 0.00,
                    '10-year': 0.00,
                    '15-year': 0.00,
                    '20-year': 0.00
                }
                return False

            # Get age
            try:
                age = int(values["-AGE-"].strip())
            except ValueError:
                raise ValueError("Please enter a valid age before calculating monthly payments")

            # Determine the age group
            age_groups = [
                ('0_to_54', lambda x: x <= 54),
                ('55_to_59', lambda x: 55 <= x <= 59),
                ('60_to_64', lambda x: 60 <= x <= 64),
                ('65', lambda x: x == 65),
                ('66_to_69', lambda x: 66 <= x <= 69),
                ('70', lambda x: x == 70),
                ('71_to_74', lambda x: 71 <= x <= 74),
                ('75', lambda x: x == 75),
                ('76_to_79', lambda x: 76 <= x <= 79),
                ('80', lambda x: x == 80),
                ('81_to_82', lambda x: 81 <= x <= 82)
            ]
            age_group = next((group for group, condition in age_groups if condition(age)), None)

            if age_group is None:
                raise ValueError(f"Age {age} is out of supported range")

            # Calculate and store monthly payments
            display_values = []
            self.current_monthly_payments = {}
            
            terms = ['3-year', '5-year', '10-year', '15-year', '20-year']
            for term in terms:
                factor = self.payment_factors[age_group][term]
                if factor is not None:
                    payment = total_to_finance * factor
                    self.current_monthly_payments[term] = payment
                    display_values.append(locale.format_string('%.2f', payment, grouping=True))
                else:
                    self.current_monthly_payments[term] = None
                    display_values.append('N/A')

            # Update the table with the new values
            self.window['-MONTHLY_PAYMENTS_TABLE-'].update(values=[display_values])
            
            # Force a window refresh
            self.window.refresh()
            
            logging.info(f"Monthly payments calculated: {display_values}")
            logging.info(f"Current monthly payments stored: {self.current_monthly_payments}")
            
            # If a payment term is already selected, update it
            selected_term = values.get("Payment Term")
            if selected_term and selected_term in self.current_monthly_payments:
                self.handle_payment_term_selection(selected_term)
            
            return True

        except Exception as e:
            logging.error(f"Error in monthly payment calculation: {str(e)}")
            self.window['-MONTHLY_PAYMENTS_TABLE-'].update(values=[['0.00'] * 5])
            self.current_monthly_payments = {term: 0.00 for term in terms}
            sg.popup_error(f"An error occurred during calculation: {str(e)}")
            return False

    def update_monthly_payments(self, values):
        """Update the monthly payments table with new values"""
        try:
            # Create a list of formatted values from current_monthly_payments
            formatted_values = []
            for term in ['3-year', '5-year', '10-year', '15-year', '20-year']:
                value = self.current_monthly_payments.get(term)
                if value is None:
                    formatted_values.append('N/A')
                else:
                    try:
                        formatted_value = locale.format_string('%.2f', value, grouping=True)
                        formatted_values.append(formatted_value)
                    except (ValueError, TypeError):
                        formatted_values.append('0.00')
            
            # Update the table with a single row of values
            self.window['-MONTHLY_PAYMENTS_TABLE-'].update(values=[formatted_values])
            logging.info(f"Updated monthly payments table with values: {formatted_values}")
            return True
        except Exception as e:
            logging.error(f"Error updating monthly payments table: {str(e)}")
            self.window['-MONTHLY_PAYMENTS_TABLE-'].update(values=[['0.00'] * 5])
            return False
        
    def handle_payment_term_selection(self, selected_term):
        """Handle payment term selection and update relevant fields"""
        try:
            if selected_term in self.current_monthly_payments:
                payment_value = self.current_monthly_payments[selected_term]
                if payment_value is not None:
                    # Format the payment value directly
                    formatted_payment = locale.format_string('%.2f', payment_value, grouping=True)
                    self.window["4B Time Pay"].update(value=formatted_payment)
                    
                    try:
                        # Get other values directly as floats
                        single_pay = float(self.window["4A Single Pay"].get().replace('$', '').replace(',', '') or 0)
                        journey_home = float(self.window["4C Single Pay Journey Home"].get().replace('$', '').replace(',', '') or 0)
                        lpr = float(self.window["4D LPR"].get().replace('$', '').replace(',', '') or 0)
                        
                        # Calculate total
                        total_4_abcd = single_pay + payment_value + journey_home + lpr
                        
                        # Update Total 4
                        self.window["Total 4 \\(ABCD\\)"].update(value=locale.format_string('%.2f', total_4_abcd, grouping=True))
                        self.window.refresh()
                        
                    except ValueError as e:
                        logging.error(f"Error converting values: {str(e)}")
                        return False
                    
                    logging.info(f"Updated 4B Time Pay with {formatted_payment} for {selected_term}")
                    return True
                else:
                    self.window["4B Time Pay"].update(value="")
                    self.window.refresh()
                    logging.info(f"Payment term {selected_term} not available for current age")
                    return False
            
        except Exception as e:
            logging.error(f"Error handling payment term selection: {str(e)}")
            self.window["4B Time Pay"].update(value="")
            self.window.refresh()
            return False
            
    def calculate_grand_total(self, values):
        try:
            # First calculate total discount using existing function
            total_discount = self.calculate_total_discount(values)
            
            # Step 1: Calculate GST and PST
            total_gst_amount = 0  # Sum of all GST applicable amounts
            total_pst = 0
            
            # First sum up all GST applicable amounts
            for field in self.gst_fields:
                try:
                    value = float(values[field].replace('$', '').replace(',', '') or 0)
                    total_gst_amount += value
                    
                    # Still calculate PST as before
                    if field in self.pst_fields:
                        # Skip PST for B1 if Minimum Cremation package is selected
                        if field == "B1" and values.get("Casket", "").strip() == "Basic Cremation Container":
                            continue
                        total_pst += value * 0.07
                except ValueError:
                    continue
            
            # Calculate GST after subtracting total discount
            gst_base = total_gst_amount - total_discount
            total_gst = gst_base * 0.05

            # Step 2: Calculate Totals
            sections = {
                'A': ['A1', 'A2A', 'A2B', 'A2C', 'A2D', 'A3', 'A4A', 'A4B', 'A4C',
                    'A5A', 'A5B', 'A5C', 'A5D', 'A6', 'A7', 'A8', 'A9A', 'A9B',
                    'A9C', 'A9D', 'A9E', 'A9F', 'A9G', 'A9H'],
                'B': ['B1', 'B2', 'B3', 'B4', 'B5', 'B6', 'B7'],
                'C': ['C1', 'C2', 'C3', 'C4', 'C5', 'C6', 'C7', 'C8', 'C9', 'C10'],
                'D': ['D1', 'D2', 'D3', 'D4', 'D5', 'D6', 'D7', 'D8']
            }

            totals = {section: sum(float(values[field].replace('$', '').replace(',', '') or 0) for field in fields)
                    for section, fields in sections.items()}

            # Copy Total D to Total D_2
            self.window["Total D_2"].update(locale.format_string('%.2f', totals['D'], grouping=True))

            # Step 3: Calculate Grand Total
            total_abc = totals['A'] + totals['B'] + totals['C']
            grand_total = total_abc - total_discount + total_gst + total_pst + totals['D']

            # Update all calculated values
            self.window["GST"].update(locale.format_string('%.2f', total_gst, grouping=True))
            self.window["PST"].update(locale.format_string('%.2f', total_pst, grouping=True))
            for section, total in totals.items():
                self.window[f"Total {section}"].update(locale.format_string('%.2f', total, grouping=True))
            self.window["Total \\(ABC\\)"].update(locale.format_string('%.2f', total_abc, grouping=True))
            self.window["Grand Total"].update(locale.format_string('%.2f', grand_total, grouping=True))
            
            # Update 3A Goods and Services with the grand total
            self.window["3A Goods and Services"].update(locale.format_string('%.2f', grand_total, grouping=True))
            
            # Step 4: Calculate Preplanned Amount
            monument_marker = float(values["3B MonumentMarker"].replace('$', '').replace(',', '') or 0)
            other_expenses = float(values["3C Other Expenses"].replace('$', '').replace(',', '') or 0)
            final_documents = float(values["3D Final Documents Service"].replace('$', '').replace(',', '') or 0)
            journey_home = float(values["3E Journey Home"].replace('$', '').replace(',', '') or 0)
            total_preplanned = grand_total + monument_marker + other_expenses + final_documents + journey_home
            
            # Update Total 3
            self.window["Total 3"].update(locale.format_string('%.2f', total_preplanned, grouping=True))

            self.window.refresh()
            logging.info("Grand Total and Preplanned Amount calculations completed successfully")
            
        except Exception as e:
            logging.error(f"Error in calculations: {str(e)}")
            sg.popup_error(f"An error occurred during calculation: {str(e)}")

    def validate_dollar_input(self, key):
        """Validate dollar input to only allow digits, decimal point, and commas"""
        try:
            widget = self.window[key].Widget
            current_value = widget.get()
            last_stored = self.last_value.get(key, '')

            # Quick handling of deletion/backspace
            if len(current_value) < len(last_stored):
                self.last_value[key] = current_value
                return True

            # If empty, allow it
            if not current_value:
                self.last_value[key] = ''
                return True

            # Only check the last character for validity
            last_char = current_value[-1] if current_value else ''
            if not (last_char.isdigit() or last_char in '.,'):
                widget.delete(0, tk.END)
                widget.insert(0, last_stored)
                return "break"

            # Simple decimal point check
            if last_char == '.' and current_value.count('.') > 1:
                widget.delete(0, tk.END)
                widget.insert(0, last_stored)
                return "break"

            # Store valid value
            self.last_value[key] = current_value
            return True

        except Exception as e:
            logging.error(f"Error in validate_dollar_input: {str(e)}")
            return True

    def format_dollar_field(self, key, value=None):
        """Format dollar field to two decimal places"""
        try:
            widget = self.window[key].Widget
            
            # If value is provided, use it; otherwise get from widget
            if value is None:
                value = widget.get()
            else:
                # Update widget with provided value first
                self.window[key].update(value)
                
            # Handle empty or non-digit values
            if not value or not any(c.isdigit() for c in value):
                self.window[key].update('')  # Clear the field
                self.last_value[key] = ''
                
                # Get current values after clearing field
                current_values = self.get_current_values()
                
                # Still trigger calculations even when clearing field
                if any(section_key in key for section_key in ["3A", "3B", "3C", "3D", "3E"]):
                    self.calculate_section_3_total(current_values)
                elif any(section_key in key for section_key in ["4A", "4B", "4C", "4D"]):
                    self.calculate_section_4_total(current_values)
                
                # Always trigger grand total calculation
                self.calculate_grand_total(current_values)
                
                # Check if age exists and trigger monthly payment calculation
                if current_values.get("-AGE-"):
                    self.calculate_monthly_payment(current_values)
                    self.update_monthly_payments(current_values)
                return
                
            # Remove any existing commas and convert to float
            clean_value = float(value.replace(',', ''))
            
            # Format with comma separators and two decimal places
            formatted_value = locale.format_string('%.2f', clean_value, grouping=True)
            
            # Update the field with formatted value
            self.window[key].update(formatted_value)
            self.last_value[key] = formatted_value
            
            # Get current values after formatting
            current_values = self.get_current_values()      
            
            # Trigger calculations after formatting
            if any(section_key in key for section_key in ["3A", "3B", "3C", "3D", "3E"]):
                self.calculate_section_3_total(current_values)
            elif any(section_key in key for section_key in ["4A", "4B", "4C", "4D"]):
                self.calculate_section_4_total(current_values)
            
            # Always trigger grand total calculation
            self.calculate_grand_total(current_values)
            
            # Check if age exists and trigger monthly payment calculation
            if current_values.get("-AGE-"):
                self.calculate_monthly_payment(current_values)
                self.update_monthly_payments(current_values)  
            
        except ValueError as e:
            logging.error(f"Error formatting dollar field {key}: {str(e)}")
            # Revert to last valid value if available
            if key in self.last_value:
                self.window[key].update(self.last_value[key])
                    
    def update_establishment_constants(self, location):
        if location in self.location_data:
            data = self.location_data[location]
            global ESTABLISHMENT_NAME, ESTABLISHMENT_EMAIL, ESTABLISHMENT_PHONE, ESTABLISHMENT_ADDRESS, ESTABLISHMENT_CITY, ESTABLISHMENT_PROVINCE, ESTABLISHMENT_POSTAL_CODE
            
            ESTABLISHMENT_NAME = data['ESTABLISHMENT_NAME']
            ESTABLISHMENT_EMAIL = data['ESTABLISHMENT_EMAIL']
            ESTABLISHMENT_PHONE = data['ESTABLISHMENT_PHONE']
            ESTABLISHMENT_ADDRESS = data['ESTABLISHMENT_ADDRESS']
            ESTABLISHMENT_CITY = data['ESTABLISHMENT_CITY']
            ESTABLISHMENT_PROVINCE = data['ESTABLISHMENT_PROVINCE']
            ESTABLISHMENT_POSTAL_CODE = data['ESTABLISHMENT_POSTAL_CODE']

            logging.info(f"Updated establishment constants for {location}")
        else:
            logging.warning(f"Unknown location selected: {location}")

    def validate_inputs(self, values):
        if values["-EMAIL-"] and not self.validate_email(values["-EMAIL-"]):
            sg.popup_error(
                "Invalid email address. Please enter a valid email or leave it blank.")
            return False
        if values["Email_3"] and not self.validate_email(values["Email_3"]):
            sg.popup_error("Invalid email address for Beneficiary. Please enter a valid email or leave it blank.")
            return False
        if values["Representative Email"] and not self.validate_email(values["Representative Email"]):
            sg.popup_error("Invalid email address for Representative. Please enter a valid email or leave it blank.")
            return False
        if values["-BIRTHDATE-"] and not self.validate_email(values["-BIRTHDATE-"]):
            sg.popup_error(
                "Invalid birthdate. Please use the format 'Month Day, Year' (e.g., January 1, 1990) or leave it blank.")
            return False
        return True

    def validate_email(self, email):
        if not email:  # Allow empty email fields
            return True
        pattern = r'^[\w\.-]+@[\w\.-]+\.\w+$'
        return re.match(pattern, email) is not None

    def validate_birthdate(self, birthdate):
        try:
            parser.parse(birthdate)
            return True
        except ValueError:
            return False
    
    def validate_phone_input(self, event):
        """Validate phone number input to only allow digits and hyphens"""
        try:
            widget = self.window[event].Widget
            value = widget.get()
            
            # Allow backspace/delete
            if len(value) < len(self.last_value.get(event, '')):
                self.last_value[event] = value
                return True
                
            # Only allow digits
            if value and not value[-1].isdigit():
                # Revert to previous valid value
                widget.delete(0, tk.END)
                widget.insert(0, self.last_value.get(event, ''))
                return False
                
            # Format the number if it's valid
            formatted = self.format_phone_number(value)
            if formatted != value:
                widget.delete(0, tk.END)
                widget.insert(0, formatted)
                
            self.last_value[event] = formatted
            return True
            
        except Exception as e:
            logging.error(f"Error in validate_phone_input: {str(e)}")
            return True

    def validate_phone_input(self, event):
        """Validate phone number input to only allow digits and hyphens"""
        try:
            widget = self.window[event].Widget
            value = widget.get()
            
            # Allow backspace/delete
            if len(value) < len(self.last_value.get(event, '')):
                self.last_value[event] = value
                return True
                
            # Only allow digits
            if value and not value[-1].isdigit():
                # Revert to previous valid value
                widget.delete(0, tk.END)
                widget.insert(0, self.last_value.get(event, ''))
                return False
                
            # Format the number if it's valid
            formatted = self.format_phone_number(value)
            if formatted != value:
                widget.delete(0, tk.END)
                widget.insert(0, formatted)
                
            self.last_value[event] = formatted
            return True
            
        except Exception as e:
            logging.error(f"Error in validate_phone_input: {str(e)}")
            return True
    
    def format_phone_number(self, phone):
        # Remove any non-digit characters
        phone = ''.join(filter(str.isdigit, phone))
        # Format as XXX-XXX-XXXX
        if len(phone) <= 3:
            return phone
        elif len(phone) <= 6:
            return f"{phone[:3]}-{phone[3:]}"
        else:
            return f"{phone[:3]}-{phone[3:6]}-{phone[6:10]}"

    def format_postal_code(self, postal_code):
        # Remove any non-alphanumeric characters and convert to uppercase
        postal_code = ''.join(filter(str.isalnum, postal_code)).upper()
        # Format as XXX XXX
        if len(postal_code) <= 3:
            return postal_code
        else:
            return f"{postal_code[:3]} {postal_code[3:6]}"

    def format_sin(self, sin):
        # Remove any non-digit characters
        sin = ''.join(filter(str.isdigit, sin))
        # Format as XXX-XXX-XXX
        if len(sin) <= 3:
            return sin
        elif len(sin) <= 6:
            return f"{sin[:3]}-{sin[3:]}"
        else:
            return f"{sin[:3]}-{sin[3:6]}-{sin[6:9]}"

    def calculate_age(self, birthdate):
        birth_date = parser.parse(birthdate)
        today = datetime.now()
        age = relativedelta(today, birth_date).years
        return age
    
    def update_age(self, birthdate):
        try:
            age = self.calculate_age(birthdate)
            self.window["-AGE-"].update(str(age))
        except ValueError:
            self.window["-AGE-"].update("")

    def toggle_beneficiary_address_fields(self, same_address):
        for key in ["Address \\(if different\\)", "City_4", "Province_4", "Postal Code_4"]:
            self.window[key].update(disabled=same_address)
            
    def show_pdf_open_error(self, filename):
        """Show error popup when PDF is open and can't be overwritten"""
        sg.popup_error(
            f"Cannot save PDF because it is currently open:\n\n"
            f"{filename}\n\n"
            "Please close the PDF and try again.",
            title="PDF File In Use",
            keep_on_top=True
        )

    def autofill_pdfs(self, values):
        if not all(os.path.exists(pdf) for pdf in self.pdf_paths.values()):
            logging.error(f"One or more PDF files not found. Base path: {self.base_path}")
            sg.popup_error(f"One or more PDF files not found. Please check the Forms directory.\nLooking in: {self.base_path}")
            return

        # Create 'Filled_Forms' folder in the same directory as the executable
        if getattr(sys, 'frozen', False):
            output_dir = Path(sys.executable).parent / "Filled Forms"
        else:
            output_dir = Path(os.getcwd()) / "Filled Forms"
        output_dir.mkdir(exist_ok=True)

        output_filenames = {
            1: f"{values['-FIRST-']} {values['-LAST-']} - Protector Plus TruStage Application form.pdf",
            2: f"{values['-FIRST-']} {values['-LAST-']} - Personal Information Sheet.pdf",
            3: f"{values['-FIRST-']} {values['-LAST-']} - Instructions Concerning My Arrangements.pdf",
            4: f"{values['-FIRST-']} {values['-LAST-']} - Pre-Arranged Funeral Service Agreement.pdf",
            5: f"{values['-FIRST-']} {values['-LAST-']} - Journey Home Enrollment Form.pdf"
        }

        output_pdfs = {
            i: output_dir / filename for i, filename in output_filenames.items()
        }

        try:
            age = self.calculate_age(values["-BIRTHDATE-"]) if values["-BIRTHDATE-"] else ""
            data = {k: v for k, v in values.items() if v}
            today = date.today()
            formatted_date = today.strftime("%B %d, %Y")
            formatted_date_short = today.strftime("%d/%m/%y")

            # Create data dictionaries
            data_dicts = self.create_data_dictionaries(data, age, formatted_date, formatted_date_short, today)
            
            # Log the contents of data_dict1
            logging.info(f"Contents of data_dict1: {data_dicts[1]}")

            # Fill PDFs
            for i, (input_pdf, output_pdf) in enumerate(zip(self.pdf_paths.values(), output_pdfs.values()), 1):
                data_dict = data_dicts[i]
                try:
                    fillpdfs.write_fillable_pdf(input_pdf, output_pdf, data_dict)
                    logging.info(f"Filled PDF {i} written to: {output_pdf}")
                    logging.info(f"Filled values for PDF {i}:")
                    for field, value in data_dict.items():
                        logging.info(f"  {field}: {value}")
                except PermissionError:
                    logging.error(f"Permission denied when writing to: {output_pdf}")
                    self.show_pdf_open_error(output_pdf)
                    return

            sg.popup(f"""PDFs have been filled successfully.
    Saved as:
    {chr(10).join([f"{i}. {filename}" for i, filename in output_filenames.items()])}
    In the Filled_Forms directory next to the application.""")
                
        except Exception as e:
            logging.error(f"Error filling PDFs: {str(e)}")
            sg.popup_error(f"Error filling PDFs: {str(e)}")

    def format_birthdate_short(self, birthdate_str):
        try:
            # Parse the input date (e.g., "April 8, 1995")
            birth_date = parser.parse(birthdate_str)
            # Format to dd/mm/yy
            return birth_date.strftime("%d/%m/%y")
        except (ValueError, TypeError):
            return birthdate_str

    def get_discount_descriptions(self):
        """Get all visible discount descriptions combined with ' + '"""
        try:
            descriptions = []
            # Iterate through all possible discount rows
            for i in range(getattr(self.window, 'metadata', 0) + 1):
                row_key = ('-DISCOUNT-ROW-', i)
                desc_key = ('-DISCOUNT-DESC-', i)
                
                # Only process visible rows
                if self.window[row_key].visible:
                    desc = self.window[desc_key].get().strip()
                    if desc:  # Only add non-empty descriptions
                        descriptions.append(desc)
            
            # Combine descriptions with ' + ' if there are multiple
            combined_description = " + ".join(descriptions) if descriptions else ""
            logging.info(f"Combined discount descriptions: {combined_description}")
            return combined_description
            
        except Exception as e:
            logging.error(f"Error getting discount descriptions: {str(e)}")
            return ""

    def get_cadence_discount_amount(self, values):
        """Find the Cadence discount amount from the dynamic discount rows"""
        try:
            # Search through all visible discount rows
            for i in range(getattr(self.window, 'metadata', 0) + 1):
                if self.window[('-DISCOUNT-ROW-', i)].visible:
                    desc = self.window[('-DISCOUNT-DESC-', i)].get().strip().lower()
                    if desc == 'cadence':
                        amount = self.window[('-DISCOUNT-AMT-', i)].get()
                        # Clean up the amount string
                        amount = amount.replace('$', '').replace(',', '')
                        return amount if amount else ''
            return ''
            
        except Exception as e:
            logging.error(f"Error finding Cadence discount: {str(e)}")
            return ''

    def create_data_dictionaries(self, data, age, formatted_date, formatted_date_short, today):
        formatted_birthdate = self.format_birthdate_short(data.get('-BIRTHDATE-', ''))
        
        # Get the selected Kearney Location
        selected_location = data.get("Kearney Location", "")
        establishment = self.location_data.get(selected_location, {})
        
        # Mapping of location names to dictionary keys
        location_mapping = {
        "Kearney Funeral Services (KFS)": "Kearney Vancouver Chapel",
        "Kearney Burnaby Chapel (KBC)": "Kearney Burnaby Chapel",
        "Kearney Burquitlam Funeral Home (BFH)": "Kearney Burquitlam Funeral Home",
        "Kearney Columbia-Bowell Chapel (CBC)": "Kearney ColumbiaBowell Chapel",
        "Kearney Cloverdale & South Surrey (CLO)": "Kearney Cloverdale South Surrey"
        }

        # Data dictionary for the "Protector Plus TruStage Application form - New" PDF
        data_dict1 = {
            'I understand that this is an enrollment into a group policy in order to provide funding for funeral expenses': 'On',
            'Establishment Name': establishment.get('ESTABLISHMENT_NAME', ''),
            'Phone': establishment.get('ESTABLISHMENT_PHONE', ''),
            'Email': establishment.get('ESTABLISHMENT_EMAIL', ''),
            'Address': establishment.get('ESTABLISHMENT_ADDRESS', ''),
            'City': establishment.get('ESTABLISHMENT_CITY', ''),
            'Province': establishment.get('ESTABLISHMENT_PROVINCE', ''),
            'Postal Code': establishment.get('ESTABLISHMENT_POSTAL_CODE', ''),
            'First Name': data.get('-FIRST-', ''),
            'MI': data.get('-MIDDLE-', '')[:1],
            'Last Name': data.get('-LAST-', ''),
            'Birthdate ddmmyy': formatted_birthdate,
            'Age': str(age),
            'Gender': data.get('-GENDER-', ''),
            'SIN': data.get('SIN', ''),
            'Occupation': data.get('-OCCUPATION-', ''),
            'Phone_2': data.get('-PHONE-', ''),
            'Email_2': data.get('-EMAIL-', ''),
            'Mailing Address': data.get('-ADDRESS-', ''),
            'City_2': data.get('-CITY-', ''),
            'Province_2': data.get('-PROVINCE-', ''),
            'Postal Code_2': data.get('-POSTAL-', ''),
            'Name': data.get('Name', ''),
            'Relationship': data.get('Relationship', ''),
            'Phone_3': data.get('Phone_3', ''),
            'Email_3': data.get('Email_3', ''),
            '3A Goods and Services': data.get('3A Goods and Services', ''),
            '3B MonumentMarker': data.get('3B MonumentMarker', ''),
            '3C Other Expenses': data.get('3C Other Expenses', ''),
            '3D Final Documents Service': data.get('3D Final Documents Service', ''),
            '3E Journey Home': data.get('3E Journey Home', ''),
            'Total 3': data.get('Total 3', ''),
            '1-year': '',
            '3-year': '',
            '5-year': '',
            '10-year': '',
            '15-year': '',
            '20-year': '',
            '4A Single Pay': data.get('4A Single Pay', ''),
            '4B Time Pay': data.get('4B Time Pay', ''),
            '4C Single Pay Journey Home': data.get('4C Single Pay Journey Home', ''),
            '4D LPR': data.get('4D LPR', ''),
            'Total 4 \\(ABCD\\)': data.get('Total 4 \\(ABCD\\)', ''),
            'Monthly': 'On',
            'Location Where Signed': ', '.join(filter(None, [
                data.get('Signed City', ''),
                data.get('Signed Province', '')
            ])),
            'Date ddmmyy': formatted_date_short,
            'Representative Name': (f"{data.get('Representative First Name', '')} {data.get('Representative Middle Name', '')} {data.get('Representative Last Name', '')}"
                                 if data.get('Representative Middle Name')
                                 else f"{data.get('Representative First Name', '')} {data.get('Representative Last Name', '')}").strip(),
            'ID': data.get('Representative ID', ''),
            'Phone_5': data.get('Representative Phone', ''),
            'Email_5': data.get('Representative Email', ''),
            'Date ddmmyy_3': formatted_date_short,
            'Payment \\(PAC\\)': 'On',
            'I hereby assign as its interest may lie the death benefit of the certificate applied for and to be issued to the funeral Establishment indicated above to provide funeral goods and': 'On',
            'I request that no new product be offered to me by TruStage Life of Canada or their affiliates or partners': 'On',
            'I also hereby assign the death benefit of the certificate to the funeral Establishment to provide certain cemetery goods and services and elect my certificate to be an EFA': 'On',
            'Protector Plus not available on FEGA IP': 'On'
            
        }
        
        selected_term = data.get('Payment Term', '')
        logging.info(f"Selected Payment Term: {selected_term}")
        if selected_term:
            if selected_term in data_dict1:
                data_dict1[selected_term] = 'On'
                logging.info(f"Updated {selected_term} to 'On'")
            else:
                logging.warning(f"Selected term '{selected_term}' not found in data_dict1")

        # Handle beneficiary address
        if data.get('-SAME_ADDRESS-'):
            data_dict1['Address \\(if different\\)'] = data.get('-ADDRESS-', '')
            data_dict1['City_4'] = data.get('-CITY-', '')
            data_dict1['Province_4'] = data.get('-PROVINCE-', '')
            data_dict1['Postal Code_4'] = data.get('-POSTAL-', '')
        else:
            data_dict1['Address \\(if different\\)'] = data.get('Address \\(if different\\)', '')
            data_dict1['City_4'] = data.get('City_4', '')
            data_dict1['Province_4'] = data.get('Province_4', '')
            data_dict1['Postal Code_4'] = data.get('Postal Code_4', '')

        # Data dictionary for the "Personal Information Sheet - New" PDF
        data_dict2 = {
            'Date': formatted_date,
            'Last name': data.get('-LAST-', ''),
            'First name': data.get('-FIRST-', ''),
            'Middle name': data.get('-MIDDLE-', ''),
            'Address_1': ', '.join(filter(None, [
                data.get('-ADDRESS-', ''),
                data.get('-CITY-', ''),
                data.get('-PROVINCE-', ''),
                data.get('-POSTAL-', '')
            ])),
            'Phone': data.get('-PHONE-', ''),
            'Email': data.get('-EMAIL-', ''),
            'Date of Birth': data.get('-BIRTHDATE-', ''),
            'Occupation': data.get('-OCCUPATION-', ''),
            'SIN': data.get('SIN', ''),
            'Death Certificate #': data.get('Death_Certificates_Quantity', ''),
            'Kearney Vancouver Chapel': '',
            'Kearney ColumbiaBowell Chapel': '',
            'Kearney Burquitlam Funeral Home': '',
            'Kearney Burnaby Chapel': '',
            'Kearney Cloverdale South Surrey': ''
        }
        
        if selected_location in location_mapping:
            data_dict2[location_mapping[selected_location]] = 'Yes'

        # Data dictionary for the "Instructions Concerning My Arrangements - New" PDF
        data_dict3 = {
            'Date': formatted_date,
            'Name': (f"{data.get('-FIRST-', '')} {data.get('-MIDDLE-', '')} {data.get('-LAST-', '')}"
                                 if data.get('-MIDDLE-')
                                 else f"{data.get('-FIRST-', '')} {data.get('-LAST-', '')}").strip(),
            'Phone': data.get('-PHONE-', ''),
            'Email': data.get('-EMAIL-', ''),
            'Type of Service': data.get('Type of Service', ''),
            'Service to be Held at': establishment.get('ESTABLISHMENT_NAME', ''),
            'Address': ', '.join(filter(None, [
                establishment.get('ESTABLISHMENT_ADDRESS', ''),
                establishment.get('ESTABLISHMENT_CITY', ''),
                establishment.get('ESTABLISHMENT_PROVINCE', ''),
                establishment.get('ESTABLISHMENT_POSTAL_CODE', '')
            ])),
            'Death Certificates': data.get('Death_Certificates_Quantity', ''),
            'Casket': '',
            'Urn': '',
            'Kearney Vancouver Chapel': '',
            'Kearney ColumbiaBowell Chapel': '',
            'Kearney Burquitlam Funeral Home': '',
            'Kearney Burnaby Chapel': '',
            'Kearney Cloverdale South Surrey': ''
        }
        
        if selected_location in location_mapping:
            data_dict3[location_mapping[selected_location]] = 'Yes'
        
        if data.get('Casket'):
            if data.get('B1'):
                data_dict3['Casket'] = f"{data.get('Casket')} - ${data.get('B1')}"
            else:
                data_dict3['Casket'] = data.get('Casket')
        
        if data.get('Urn') or data.get('Keepsake'):
            if data.get('Keepsake'):
                if data.get('B3'):
                    data_dict3['Urn'] = f"{data.get('Keepsake')} - ${data.get('B3')}"
                else:
                    data_dict3['Urn'] = data.get('Keepsake')
            elif data.get('Urn'):
                if data.get('B2'):
                    data_dict3['Urn'] = f"{data.get('Urn')} - ${data.get('B2')}"
                else:
                    data_dict3['Urn'] = data.get('Urn')

        # Get Cadence discount amount
        cadence_discount = self.get_cadence_discount_amount(data)        
        
        # Data dictionary for the "Pre-Arranged Funeral Service Agreement - New
        data_dict4 = {
            'Purchaser': (f"{data.get('-FIRST-', '')} {data.get('-MIDDLE-', '')} {data.get('-LAST-', '')}"
                                 if data.get('-MIDDLE-')
                                 else f"{data.get('-FIRST-', '')} {data.get('-LAST-', '')}").strip(),
            'PURCHASERS NAME': (f"{data.get('-FIRST-', '')} {data.get('-MIDDLE-', '')} {data.get('-LAST-', '')}"
                                 if data.get('-MIDDLE-')
                                 else f"{data.get('-FIRST-', '')} {data.get('-LAST-', '')}").strip(),
            'Phone Number': data.get('-PHONE-', ''),
            'Address': ', '.join(filter(None, [data.get('-ADDRESS-', ''), data.get('-CITY-', ''),
                data.get('-PROVINCE-', ''), data.get('-POSTAL-', '')
            ])),
            'Type of Service': data.get('Type of Service', ''),
            'FUNERAL HOME REPRESENTATIVE NAME': (f"{data.get('Representative First Name', '')} {data.get('Representative Middle Name', '')} {data.get('Representative Last Name', '')}"
                                 if data.get('Representative Middle Name')
                                 else f"{data.get('Representative First Name', '')} {data.get('Representative Last Name', '')}").strip(),
            'BENEFICIARY': f"{data.get('-FIRST-', '')} {data.get('-MIDDLE-', '')} {data.get('-LAST-', '')}",
            'DATE OF BIRTH': data.get('-BIRTHDATE-', ''),
            'ADDRESS CITY PROVINCE POSTAL CODE': ', '.join(filter(None, [
                data.get('-ADDRESS-', ''),
                data.get('-CITY-', ''),
                data.get('-PROVINCE-', ''),
                data.get('-POSTAL-', '')
            ])),
            'TELEPHONE NUMBER': data.get('-PHONE-', ''),
            'Day': self.ordinal(today.day),
            'Month': today.strftime("%B"),
            'Year': str(today.year),
            'SIN': data.get('SIN', ''),
            # New fields from the package layout
            'A1': self.convert_to_float(data.get('A1', '')),
            'A2A': self.convert_to_float(data.get('A2A', '')),
            'Pallbearers': data.get('Pallbearers', ''),
            'A2B': self.convert_to_float(data.get('A2B', '')),
            'Alternate Day Interment 1': data.get('Alternate Day Interment 1', ''),
            'A2C': self.convert_to_float(data.get('A2C', '')),
            'Alternate Day Interment 2': data.get('Alternate Day Interment 2', ''),
            'A2D': self.convert_to_float(data.get('A2D', '')),
            'A3': self.convert_to_float(data.get('A3', '')),
            'A4A': self.convert_to_float(data.get('A4A', '')),
            'A4B': self.convert_to_float(data.get('A4B', '')),
            'A4C': self.convert_to_float(data.get('A4C', '')),
            'A5A': self.convert_to_float(data.get('A5A', '')),
            'A5B': self.convert_to_float(data.get('A5B', '')),
            'Pacemaker Removal': data.get('Pacemaker Removal', ''),
            'A5C': self.convert_to_float(data.get('A5C', '')),
            'Autopsy Care': data.get('Autopsy Care', ''),
            'A5D': self.convert_to_float(data.get('A5D', '')),
            'Evening Prayers or Visitation': data.get('Evening Prayers or Visitation', ''),
            'A6': self.convert_to_float(data.get('A6', '')),
            'Weekend or Statutory Holiday': data.get('Weekend or Statutory Holiday', ''),
            'A7': self.convert_to_float(data.get('A7', '')),
            'Reception Facilities': data.get('Reception Facilities', ''),
            'A8': self.convert_to_float(data.get('A8', '')),
            'Delivery of Cremated Remains': data.get('Delivery of Cremated Remains', ''),
            'A9A': self.convert_to_float(data.get('A9A', '')),
            'Transfer to Crematorium or Airport': data.get('Transfer to Crematorium or Airport', ''),
            'A9B': self.convert_to_float(data.get('A9B', '')),
            'Lead Vehicle': data.get('Lead Vehicle', ''),
            'A9C': self.convert_to_float(data.get('A9C', '')),
            'Service Vehicle': data.get('Service Vehicle', ''),
            'A9D': self.convert_to_float(data.get('A9D', '')),
            'Funeral Coach': data.get('Funeral Coach', ''),
            'A9E': self.convert_to_float(data.get('A9E', '')),
            'Limousine': data.get('Limousine', ''),
            'A9F': self.convert_to_float(data.get('A9F', '')),
            'Additional Limousines': data.get('Additional Limousines', ''),
            'A9G': self.convert_to_float(data.get('A9G', '')),
            'Flower Van': data.get('Flower Van', ''),
            'A9H': self.convert_to_float(data.get('A9H', '')),
            'Total A': self.convert_to_float(data.get('Total A', '')),
            'Casket': (f"{data.get('Casket', '')} (Discount - {data.get('Discount_Casket', '')})" 
                      if data.get('Casket') and data.get('Discount_Casket')
                      else data.get('Casket', '')),
            'B1': self.convert_to_float(data.get('B1', '')),
            'Urn': data.get('Urn', ''),
            'B2': self.convert_to_float(data.get('B2', '')),
            'Keepsake': data.get('Keepsake', ''),
            'B3': self.convert_to_float(data.get('B3', '')),
            'Traditional Mourning Items': data.get('Traditional Mourning Items', ''),
            'B4': self.convert_to_float(data.get('B4', '')),
            'Memorial Stationary': (f"Cards ({data.get('Cards_Qty', '')} x $2.95)"
                                 if data.get('Cards_Qty')
                                 else ''),
            'B5': self.convert_to_float(data.get('B5', '')),
            'Funeral Register': (f"Guest Book ({data.get('Guest_Book_Qty', '')} x $75.00)"
                                 if data.get('Guest_Book_Qty')
                                 else ''),
            'B6': self.convert_to_float(data.get('B6', '')),
            'Other_1': data.get('Other_1', ''),
            'B7': self.convert_to_float(data.get('B7', '')),
            'Total B': self.convert_to_float(data.get('Total B', '')),
            'Cemetery': data.get('Cemetery', ''),
            'C1': self.convert_to_float(data.get('C1', '')),
            'Crematorium': data.get('Crematorium', ''),
            'C2': self.convert_to_float(data.get('C2', '')),
            'Obituary Notices': data.get('Obituary Notices', ''),
            'C3': self.convert_to_float(data.get('C3', '')),
            'Flowers': data.get('Flowers', ''),
            'C4': self.convert_to_float(data.get('C4', '')),
            'CPBC Administration Fee': data.get('CPBC Administration Fee', ''),
            'C5': self.convert_to_float(data.get('C5', '')),
            'Hostess': data.get('Hostess', ''),
            'C6': self.convert_to_float(data.get('C6', '')),
            'Markers': data.get('Markers', ''),
            'C7': self.convert_to_float(data.get('C7', '')),
            'Catering': data.get('Catering', ''),
            'C8': self.convert_to_float(data.get('C8', '')),
            'Other_2': (f"{data.get('Other_2', '')} (Discount - ${cadence_discount})"
                       if data.get('Other_2') and cadence_discount
                       else data.get('Other_2', '')),
            'C9': self.convert_to_float(data.get('C9', '')),
            'Other_3': data.get('Other_3', ''),
            'C10': self.convert_to_float(data.get('C10', '')),
            'Total C': self.convert_to_float(data.get('Total C', '')),
            'Clergy Honorarium': data.get('Clergy Honorarium', ''),
            'D1': self.convert_to_float(data.get('D1', '')),
            'Church Honorarium': data.get('Church Honorarium', ''),
            'D2': self.convert_to_float(data.get('D2', '')),
            'Altar Servers': data.get('Altar Servers', ''),
            'D3': self.convert_to_float(data.get('D3', '')),
            'Organist': data.get('Organist', ''),
            'D4': self.convert_to_float(data.get('D4', '')),
            'Soloist': data.get('Soloist', ''),
            'D5': self.convert_to_float(data.get('D5', '')),
            'Harpist': data.get('Harpist', ''),
            'D6': self.convert_to_float(data.get('D6', '')),
            'Death Certificates': (f"{data.get('Death_Certificates_Quantity', '')} x $27.00"
                                 if data.get('Death_Certificates_Quantity')
                                 else ''),
            'D7': self.convert_to_float(data.get('D7', '')),
            'Other_4': data.get('Other_4', ''),
            'D8': self.convert_to_float(data.get('D8', '')),
            'Total D': self.convert_to_float(data.get('Total D', '')),
            'Total \\(ABC\\)': self.convert_to_float(data.get('Total \\(ABC\\)', '')),
            'Discount_description': self.get_discount_descriptions(),
            'Discount': f"({self.convert_to_float(data.get('Discount', ''))})" if data.get('Discount') else '',
            'GST': self.convert_to_float(data.get('GST', '')),
            'PST': self.convert_to_float(data.get('PST', '')),
            'Total D_2': self.convert_to_float(data.get('Total D_2', '')),
            'Grand Total': self.convert_to_float(data.get('Grand Total', '')),
            'City Province': ', '.join(filter(None, [
                data.get('Signed City', ''),
                data.get('Signed Province', '')
            ])),
            'Kearney Vancouver Chapel': '',
            'Kearney ColumbiaBowell Chapel': '',
            'Kearney Burquitlam Funeral Home': '',
            'Kearney Burnaby Chapel': '',
            'Kearney Cloverdale South Surrey': ''
        }
        
        if selected_location in location_mapping:
            data_dict4[location_mapping[selected_location]] = 'On'
        
        # Data dictionary for the "Journey Home Enrollment Form - New" PDF
        data_dict5 = {
            'Purchase Date ddmmyy': formatted_date_short,
            'First Name': data.get('-FIRST-', ''),
            'MI': data.get('-MIDDLE-', '')[:1],
            'Last Name': data.get('-LAST-', ''),
            'Date of Birth': formatted_birthdate,
            'Address': data.get('-ADDRESS-', ''),
            'City': data.get('-CITY-', ''),
            'Province': data.get('-PROVINCE-', ''),
            'Postal Code': data.get('-POSTAL-', ''),
            'Phone Number': data.get('-PHONE-', ''),
            'Email': data.get('-EMAIL-', ''),
            'Rep First Name': data.get('Representative First Name', ''),
            'Rep MI': data.get('Representative Middle Name', '')[:1],
            'Rep Last Name': data.get('Representative Last Name', ''),
            'Representative ID': data.get('Representative ID', ''),
            'Rep Phone Number': data.get('Representative Phone', ''),
            'Funeral Home Name if known': establishment.get('ESTABLISHMENT_NAME', ''),
            'Amount Due': data.get('3E Journey Home', ''),
            'Male': '',
            'Female': ''
        }
        
        gender = data.get('-GENDER-', '').lower().strip()
        if gender in ['m', 'male']:
            data_dict5['Male'] = 'On'
        elif gender in ['f', 'female']:
            data_dict5['Female'] = 'On'

        return {1: data_dict1, 2: data_dict2, 3: data_dict3, 4: data_dict4, 5: data_dict5}

if __name__ == "__main__":
    autofiller = PDFAutofiller()
    autofiller.run()