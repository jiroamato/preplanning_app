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

# Establishment Constants
ESTABLISHMENT_NAME = "Kearney Columbia-Bowell Chapel"
ESTABLISHMENT_EMAIL = "Columbia-Bowell@KearneyFS.com"
ESTABLISHMENT_PHONE = "604-521-4881"
ESTABLISHMENT_ADDRESS = "219-6th Street"
ESTABLISHMENT_CITY = "New Westminster"
ESTABLISHMENT_PROVINCE = "BC"
ESTABLISHMENT_POSTAL_CODE = "V3L 3A3"

# Representative Constants
REPRESENTATIVE_NAME = "Angelina Amato"
REPRESENTATIVE_ID = "502966002"
REPRESENTATIVE_PHONE = "778-869-2512"
REPRESENTATIVE_EMAIL = "a.amato@kearneyfs.com"


class PDFTypes:
    PRE_ARRANGED_FULL_SERVICE = "Pre-Arranged Funeral Service Agreement - Full Funeral Service.pdf"
    TRUSTAGE_APPLICATION = "Protector Plus TruStage Application form.pdf"
    PERSONAL_INFO_SHEET = "Personal Information Sheet.pdf"
    INSTRUCTIONS = "Instructions Concerning My Arrangements.pdf"


class PDFAutofiller:
    def __init__(self):
        self.base_path = self.get_base_path()
        self.setup_logging()
        
        self.location_data = {
            "Kearney Funeral Services (KFS)": {'ESTABLISHMENT_NAME': 'Kearney Funeral Services (KFS)', 'ESTABLISHMENT_EMAIL': 'Vancouver.Chapel@KearneyFS.com', 'ESTABLISHMENT_PHONE': '604-736-2668', 'ESTABLISHMENT_ADDRESS': '450 W 2nd Ave', 'ESTABLISHMENT_CITY': 'Vancouver', 'ESTABLISHMENT_PROVINCE': 'BC', 'ESTABLISHMENT_POSTAL_CODE': 'V5Y 1E2'},
            "Kearney Burnaby Chapel (KBC)": {'ESTABLISHMENT_NAME': 'Kearney Burnaby Chapel', 'ESTABLISHMENT_EMAIL': 'Burnaby@KearneyFS.com', 'ESTABLISHMENT_PHONE': '604-299-6889', 'ESTABLISHMENT_ADDRESS': '4715 Hastings St', 'ESTABLISHMENT_CITY': 'Burnaby', 'ESTABLISHMENT_PROVINCE': 'BC', 'ESTABLISHMENT_POSTAL_CODE': 'V5C 2K8'},
            "Kearney Burquitlam Funeral Home (BFH)": {'ESTABLISHMENT_NAME': 'Kearney Burquitlam Funeral Home', 'ESTABLISHMENT_EMAIL': 'Info@BurquitlamFuneralHome.ca', 'ESTABLISHMENT_PHONE': '604-936-9987', 'ESTABLISHMENT_ADDRESS': '102-200 Bernatchey St', 'ESTABLISHMENT_CITY': 'Coquitlam', 'ESTABLISHMENT_PROVINCE': 'BC', 'ESTABLISHMENT_POSTAL_CODE': 'V3K 0H8'},
            "Kearney Columbia-Bowell Chapel (CBC)": {'ESTABLISHMENT_NAME': 'Kearney Columbia-Bowell Chapel', 'ESTABLISHMENT_EMAIL': 'Columbia-Bowell@KearneyFS.com', 'ESTABLISHMENT_PHONE': '604-521-4881', 'ESTABLISHMENT_ADDRESS': '219 6th Street', 'ESTABLISHMENT_CITY': 'New Westminster', 'ESTABLISHMENT_PROVINCE': 'BC', 'ESTABLISHMENT_POSTAL_CODE': 'V3L 3A3'},
            "Kearney Cloverdale & South Surrey (CLO)": {'ESTABLISHMENT_NAME': 'Kearney Clovedale & South Surrey', 'ESTABLISHMENT_EMAIL': 'Cloverdale@KearneyFS.com', 'ESTABLISHMENT_PHONE': '604-574-2603', 'ESTABLISHMENT_ADDRESS': '17667 57th Ave', 'ESTABLISHMENT_CITY': 'Surrey', 'ESTABLISHMENT_PROVINCE': 'BC', 'ESTABLISHMENT_POSTAL_CODE': 'V3S 1H1'}
        }

        self.kearney_locations = list(self.location_data.keys())
        
        self.gst_fields = [
            "A1", "A2A", "A2B", "A2C", "A2D", "A3", "A4A", "A4B", "A4C",
            "A5A", "A5B", "A5C", "A5D", "A6", "A7", "A8", "A9A", "A9B",
            "A9C", "A9D", "A9E", "A9F", "A9G", "A9H", "B1", "B2", "B3",
            "B4", "B5", "B6", "B7", "C2", "C3", "C4", "C7", "C8", "C9", "C10"
        ]
        
        self.pst_fields = [
            "B1", "B2", "B3", "B4", "B5", "B6", "B7", "C3", "C4", "C7", "C8", "C9", "C10"
        ]

        # Add this new attribute to store package data
        self.packages = {
            "Full Funeral Service - Burial": {
                'A1': '525.00', 'A2A': '2755.00', 'A3': '895.00', 'A4A': '365.00',
                'A5A': '365.00', 'A5B': '695.00', 'A6': '695.00', 'A9D': '295.00',
                'A9E': '525.00', 'C5': '40.00', 'Death Certificates': '4 x $27.00',
                'D7': '108.00'
            },
            "Full Funeral Service - Cremation": {
                'A1': '525.00', 'A2A': '2755.00', 'A3': '895.00', 'A4A': '365.00',
                'A5A': '365.00', 'A5B': '695.00', 'A6': '695.00', 'A9D': '295.00',
                'A9E': '525.00', 'C2': '745.00', 'C5': '40.00', 'Death Certificates': '4 x $27.00',
                'D7': '108.00'
            },
            "Cremation With Viewing": {
                'A1': '525.00', 'A2A': '895.00', 'A3': '350.00', 'A4A': '365.00',
                'A5A': '365.00', 'A5B': '695.00', 'C2': '745.00', 'C5': '40.00', 'Death Certificates': '4 x $27.00',
                'D7': '108.00'
            },
            "Witness Cremation Service": {
                'A1': '525.00', 'A2A': '1295.00', 'A3': '595.00', 'A4A': '365.00',
                'A5A': '365.00', 'A9E': '525.00', 'C2': '1495.00', 'C5': '40.00', 'Death Certificates': '4 x $27.00',
                'D7': '108.00'
            },
            "Minimum Cremation": {
                'A1': '525.00', 'A2A': '765.00', 'A3': '595.00', 'A4A': '365.00',
                'C2': '745.00', 'C5': '40.00', 'Death Certificates': '4 x $27.00',
                'D7': '108.00'
            },
            "Graveside Service": {
                'A1': '525.00', 'A2A': '1195.00', 'A3': '595.00', 'A4A': '365.00',
                'A5A': '365.00', 'A5B': '695.00', 'A6': '695.00', 'A9E': '525.00',
                'C5': '40.00', 'Death Certificates': '4 x $27.00', 'D7': '108.00'
            },
            "Memorial Service (Off Property)": {
                'A1': '525.00', 'A2A': '1895.00', 'A3': '795.00', 'A4A': '365.00',
                'C2': '745.00', 'C5': '40.00', 'Death Certificates': '4 x $27.00',
                'D7': '108.00'
            }
        }
        
        self.payment_factors = {
            '0_to_54': {'3-year': 0.03150, '5-year': 0.01995, '10-year': 0.01155, '15-year': 0.00893, '20-year': 0.00735},
            '55_to_59': {'3-year': 0.03255, '5-year': 0.01995, '10-year': 0.01260, '15-year': 0.00998, '20-year': 0.00840},
            '60_to_64': {'3-year': 0.03255, '5-year': 0.02100, '10-year': 0.01365, '15-year': 0.01050, '20-year': 0.00893},
            '65': {'3-year': 0.03360, '5-year': 0.02100, '10-year': 0.01470, '15-year': 0.01155, '20-year': 0.00998},
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
             sg.Text("Age:"), sg.Input(key="-AGE-", size=(5, 1)),
             sg.Text("Gender:"), sg.Input(key="-GENDER-", size=(20, 1)),
             sg.Text("SIN:"), sg.Input(key="SIN", size=(20, 1), enable_events=True)],
            [sg.Text("Phone:"), sg.Input(key="-PHONE-", size=(20, 1), enable_events=True),
             sg.Text("Email:"), sg.Input(key="-EMAIL-", size=(20, 1)),
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

        # Packages
        TEXT_WIDTH = 40
        INPUT_WIDTH = 20
        DOLLAR_WIDTH = 10

        section_a_layout = [
            # A. Professional Services
            [sg.Text("A. Professional Services", font=("Helvetica", 16, "bold")), sg.Text("(GST applicable)", font=("Helvetica", 10, "italic"))],
            [sg.Column([
                [sg.Text("1. Securing Release of Deceased and Transfer:", size=(TEXT_WIDTH, 1))],
                [sg.Text("2. Services of Licensed Funeral Directors and Staff:", size=(TEXT_WIDTH, 1))],
                [sg.Text("    Pallbearers:", size=(TEXT_WIDTH, 1))],
                [sg.Text("    Alternate Day Interment:", size=(TEXT_WIDTH, 1))],
                [sg.Text("    Alternate Day Interment 2:", size=(TEXT_WIDTH, 1))],
                [sg.Text("3. Administration, Documentation & Registration:", size=(TEXT_WIDTH, 1))],
                [sg.Text("4. Facilities and/or Equipment and Supplies:", size=(TEXT_WIDTH, 1))],
                [sg.Text("    Sheltering of Remains:", size=(TEXT_WIDTH, 1))],
                [sg.Text("    A/V Equipment:", size=(TEXT_WIDTH, 1))],
                [sg.Text("5. Preparation Services", size=(TEXT_WIDTH, 1))],
                [sg.Text("    a) Basic Sanitary Care, Dressing and Casketing:", size=(TEXT_WIDTH, 1))],
                [sg.Text("    b) Embalming:", size=(TEXT_WIDTH, 1))],
                [sg.Text("    c) Pacemaker Removal:", size=(TEXT_WIDTH, 1))],
                [sg.Text("    d) Autopsy Care:", size=(TEXT_WIDTH, 1))],
                [sg.Text("6. Evening Prayers or Visitation:", size=(TEXT_WIDTH, 1))],
                [sg.Text("7. Weekend or Statutory Holiday:", size=(TEXT_WIDTH, 1))],
                [sg.Text("8. Reception Facilities:", size=(TEXT_WIDTH, 1))],
                [sg.Text("9. Vehicles", size=(TEXT_WIDTH, 1))],
                [sg.Text("    Delivery of Cremated Remains:", size=(TEXT_WIDTH, 1))],
                [sg.Text("    Transfer Vehicle for Transfer to Crematorium or Airport:", size=(TEXT_WIDTH, 1))],
                [sg.Text("    Lead Vehicle:", size=(TEXT_WIDTH, 1))],
                [sg.Text("    Service Vehicle:", size=(TEXT_WIDTH, 1))],
                [sg.Text("    Funeral Coach:", size=(TEXT_WIDTH, 1))],
                [sg.Text("    Limousine:", size=(TEXT_WIDTH, 1))],
                [sg.Text("    Additional Limousines:", size=(TEXT_WIDTH, 1))],
                [sg.Text("    Flower Van:", size=(TEXT_WIDTH, 1))],
            ]), sg.Column([
                [sg.Push(), sg.Text("$"), sg.Input(key="A1", size=(DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Text("$"), sg.Input(key="A2A", size=(DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Pallbearers", size=(INPUT_WIDTH, 1)), sg.Text("$"), sg.Input(key="A2B", size=(DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Alternate Day Interment 1", size=(INPUT_WIDTH, 1)), sg.Text("$"), sg.Input(key="A2C", size=(DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Alternate Day Interment 2", size=(INPUT_WIDTH, 1)), sg.Text("$"), sg.Input(key="A2D", size=(DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Text("$"), sg.Input(key="A3", size=(DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Text("$"), sg.Input(key="A4A", size=(DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Text("$"), sg.Input(key="A4B", size=(DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Text("$"), sg.Input(key="A4C", size=(DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Text("")],  # Empty space for "5. Preparation Services"
                [sg.Push(), sg.Text("$"), sg.Input(key="A5A", size=(DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Text("$"), sg.Input(key="A5B", size=(DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Pacemaker Removal", size=(INPUT_WIDTH, 1)), sg.Text("$"), sg.Input(key="A5C", size=(DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Autopsy Care", size=(INPUT_WIDTH, 1)), sg.Text("$"), sg.Input(key="A5D", size=(DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Evening Prayers or Visitation", size=(INPUT_WIDTH, 1)), sg.Text("$"), sg.Input(key="A6", size=(DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Weekend or Statutory Holiday", size=(INPUT_WIDTH, 1)), sg.Text("$"), sg.Input(key="A7", size=(DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Reception Facilities", size=(INPUT_WIDTH, 1)), sg.Text("$"), sg.Input(key="A8", size=(DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Text("")],  # Empty space for "9. Vehicles"
                [sg.Push(), sg.Input(key="Delivery of Cremated Remains", size=(INPUT_WIDTH, 1)), sg.Text("$"), sg.Input(key="A9A", size=(DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Transfer to Crematorium or Airport", size=(INPUT_WIDTH, 1)), sg.Text("$"), sg.Input(key="A9B", size=(DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Lead Vehicle", size=(INPUT_WIDTH, 1)), sg.Text("$"), sg.Input(key="A9C", size=(DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Service Vehicle", size=(INPUT_WIDTH, 1)), sg.Text("$"), sg.Input(key="A9D", size=(DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Funeral Coach", size=(INPUT_WIDTH, 1)), sg.Text("$"), sg.Input(key="A9E", size=(DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Limousine", size=(INPUT_WIDTH, 1)), sg.Text("$"), sg.Input(key="A9F", size=(DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Additional Limousines", size=(INPUT_WIDTH, 1)), sg.Text("$"), sg.Input(key="A9G", size=(DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Flower Van", size=(INPUT_WIDTH, 1)), sg.Text("$"), sg.Input(key="A9H", size=(DOLLAR_WIDTH, 1), enable_events=True)],
            ])],
            [sg.Text("TOTAL A:", font=("Helvetica", 10, "bold"), size=(TEXT_WIDTH, 1)), sg.Push(), sg.Text("$"), sg.Input(key="Total A", size=(DOLLAR_WIDTH, 1), enable_events=True)]
        ]
        
        section_b_layout = [
            # B. Merchandise
            [sg.Text("B. Merchandise", font=("Helvetica", 16, "bold")), sg.Text("(GST and/or PST applicable)", font=("Helvetica", 10, "italic"))],
            [sg.Column([
                [sg.Text("Casket:", size=(TEXT_WIDTH, 1))],
                [sg.Text("Urn:", size=(TEXT_WIDTH, 1))],
                [sg.Text("Keepsake:", size=(TEXT_WIDTH, 1))],
                [sg.Text("Traditional Mourning Items:", size=(TEXT_WIDTH, 1))],
                [sg.Text("Memorial Stationary:", size=(TEXT_WIDTH, 1))],
                [sg.Text("Funeral Register:", size=(TEXT_WIDTH, 1))],
                [sg.Text("Funeral Register 2:", size=(TEXT_WIDTH, 1))],
            ]), sg.Column([
                [sg.Push(), sg.Input(key="Casket", size=(INPUT_WIDTH, 1)), sg.Text("$"), sg.Input(key="B1", size=(DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Urn", size=(INPUT_WIDTH, 1)), sg.Text("$"), sg.Input(key="B2", size=(DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Keepsake", size=(INPUT_WIDTH, 1)), sg.Text("$"), sg.Input(key="B3", size=(DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Traditional Mourning Items", size=(INPUT_WIDTH, 1)), sg.Text("$"), sg.Input(key="B4", size=(DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Memorial Stationary", size=(INPUT_WIDTH, 1)), sg.Text("$"), sg.Input(key="B5", size=(DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Funeral Register 1", size=(INPUT_WIDTH, 1)), sg.Text("$"), sg.Input(key="B6", size=(DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Funeral Register 2", size=(INPUT_WIDTH, 1)), sg.Text("$"), sg.Input(key="B7", size=(DOLLAR_WIDTH, 1), enable_events=True)],
            ])],
            [sg.Text("TOTAL B:", font=("Helvetica", 10, "bold"), size=(TEXT_WIDTH, 1)), sg.Push(), sg.Text("$"), sg.Input(key="Total B", size=(DOLLAR_WIDTH, 1), enable_events=True)]
        ]
        
        section_c_layout = [
            # C. Cash Disbursements
            [sg.Text("C. Cash Disbursements", font=("Helvetica", 16, "bold")), sg.Text("(GST and/or PST applicable)", font=("Helvetica", 10, "italic"))],
            [sg.Column([
                [sg.Text("Cemetery:", size=(TEXT_WIDTH, 1))],
                [sg.Text("Crematorium:", size=(TEXT_WIDTH, 1))],
                [sg.Text("Obituary Notices:", size=(TEXT_WIDTH, 1))],
                [sg.Text("Flowers:", size=(TEXT_WIDTH, 1))],
                [sg.Text("CPBC Administration Fee:", size=(TEXT_WIDTH, 1))],
                [sg.Text("Hostess:", size=(TEXT_WIDTH, 1))],
                [sg.Text("Markers:", size=(TEXT_WIDTH, 1))],
                [sg.Text("Catering:", size=(TEXT_WIDTH, 1))],
                [sg.Text("Catering 2:", size=(TEXT_WIDTH, 1))],
                [sg.Text("Catering 3:", size=(TEXT_WIDTH, 1))],
            ]), sg.Column([
                [sg.Push(), sg.Input(key="Cemetery", size=(INPUT_WIDTH, 1)), sg.Text("$"), sg.Input(key="C1", size=(DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Crematorium", size=(INPUT_WIDTH, 1)), sg.Text("$"), sg.Input(key="C2", size=(DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Obituary Notices", size=(INPUT_WIDTH, 1)), sg.Text("$"), sg.Input(key="C3", size=(DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Flowers", size=(INPUT_WIDTH, 1)), sg.Text("$"), sg.Input(key="C4", size=(DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="CPBC Administration Fee", size=(INPUT_WIDTH, 1)), sg.Text("$"), sg.Input(key="C5", size=(DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Hostess", size=(INPUT_WIDTH, 1)), sg.Text("$"), sg.Input(key="C6", size=(DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Markers", size=(INPUT_WIDTH, 1)), sg.Text("$"), sg.Input(key="C7", size=(DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Catering 1", size=(INPUT_WIDTH, 1)), sg.Text("$"), sg.Input(key="C8", size=(DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Catering 2", size=(INPUT_WIDTH, 1)), sg.Text("$"), sg.Input(key="C9", size=(DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Catering 3", size=(INPUT_WIDTH, 1)), sg.Text("$"), sg.Input(key="C10", size=(DOLLAR_WIDTH, 1), enable_events=True)],
            ])],
            [sg.Text("TOTAL C:", font=("Helvetica", 10, "bold"), size=(TEXT_WIDTH, 1)), sg.Push(), sg.Text("$"), sg.Input(key="Total C", size=(DOLLAR_WIDTH, 1), enable_events=True)]
        ]
        
        section_d_layout = [
            # D. Cash Disbursements
            [sg.Text("D. Cash Disbursements", font=("Helvetica", 16, "bold")), sg.Text("(GST exempt)", font=("Helvetica", 10, "italic"))],
            [sg.Column([
                [sg.Text("Clergy Honorarium:", size=(TEXT_WIDTH, 1))],
                [sg.Text("Church Honorarium:", size=(TEXT_WIDTH, 1))],
                [sg.Text("Altar Servers:", size=(TEXT_WIDTH, 1))],
                [sg.Text("Organist:", size=(TEXT_WIDTH, 1))],
                [sg.Text("Soloist:", size=(TEXT_WIDTH, 1))],
                [sg.Text("Harpist:", size=(TEXT_WIDTH, 1))],
                [sg.Text("Death Certificates:", size=(TEXT_WIDTH, 1))],
                [sg.Text("Other:", size=(TEXT_WIDTH, 1))],
            ]), sg.Column([
                [sg.Push(), sg.Input(key="Clergy Honorarium", size=(INPUT_WIDTH, 1)), sg.Text("$"), sg.Input(key="D1", size=(DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Church Honorarium", size=(INPUT_WIDTH, 1)), sg.Text("$"), sg.Input(key="D2", size=(DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Altar Servers", size=(INPUT_WIDTH, 1)), sg.Text("$"), sg.Input(key="D3", size=(DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Organist", size=(INPUT_WIDTH, 1)), sg.Text("$"), sg.Input(key="D4", size=(DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Soloist", size=(INPUT_WIDTH, 1)), sg.Text("$"), sg.Input(key="D5", size=(DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Harpist", size=(INPUT_WIDTH, 1)), sg.Text("$"), sg.Input(key="D6", size=(DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Death Certificates", size=(INPUT_WIDTH, 1)), sg.Text("$"), sg.Input(key="D7", size=(DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Input(key="Other", size=(INPUT_WIDTH, 1)), sg.Text("$"), sg.Input(key="D8", size=(DOLLAR_WIDTH, 1), enable_events=True)],
            ])],
            [sg.Text("TOTAL D:", font=("Helvetica", 10, "bold"), size=(TEXT_WIDTH, 1)), sg.Push(), sg.Text("$"), sg.Input(key="Total D", size=(DOLLAR_WIDTH, 1), enable_events=True)]
        ]
        
        total_charges_layout = [
            # TOTAL CHARGES
            [sg.Text("TOTAL CHARGES", font=("Helvetica", 16, "bold"))],
            [sg.Column([
                [sg.Text("TOTAL A, B, & C SECTIONS:", size=(TEXT_WIDTH, 1))],
                [sg.Text("DISCOUNT:", size=(TEXT_WIDTH, 1))],
                [sg.Text("G.S.T.:", size=(TEXT_WIDTH, 1))],
                [sg.Text("P.S.T.:", size=(TEXT_WIDTH, 1))],
                [sg.Text("TOTAL D:", size=(TEXT_WIDTH, 1))],
                [sg.Text("GRAND TOTAL:", font=("Helvetica", 10, "bold"), size=(TEXT_WIDTH, 1))],
            ]), sg.Column([
                [sg.Push(), sg.Text("$"), sg.Input(key="Total \\(ABC\\)", size=(DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Text("$"), sg.Input(key="Discount", size=(DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Text("$"), sg.Input(key="GST", size=(DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Text("$"), sg.Input(key="PST", size=(DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Text("$"), sg.Input(key="Total D_2", size=(DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Text("$"), sg.Input(key="Grand Total", size=(DOLLAR_WIDTH, 1), enable_events=True)],
            ])]
        ]
        
        preplanned_amount_layout = [
            [sg.Text("Preplanned Amount", font=("Helvetica", 16, "bold"))],
            [sg.Column([
                [sg.Text("A. Goods and Services:", size=(TEXT_WIDTH, 1))],
                [sg.Text("B. Monument/Marker:", size=(TEXT_WIDTH, 1))],
                [sg.Text("C. Other Expenses:", size=(TEXT_WIDTH, 1))],
                [sg.Text("D. Final Documents Service:", size=(TEXT_WIDTH, 1))],
                [sg.Text("E. Journey Home (Time Pay Only):", size=(TEXT_WIDTH, 1))],
                [sg.Text("Total (A+B+C+D+E):", size=(TEXT_WIDTH, 1))],
                [sg.Text("F. Lifetime Protection Rider:", size=(TEXT_WIDTH, 1))],
            ]), sg.Column([
                [sg.Push(), sg.Text("$"), sg.Input(key="3A Goods and Services", size=(DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Text("$"), sg.Input(key="3B MonumentMarker", size=(DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Text("$"), sg.Input(key="3C Other Expenses", size=(DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Text("$"), sg.Input(key="3D Final Documents Service", size=(DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Text("$"), sg.Input(key="3E Journey Home", size=(DOLLAR_WIDTH, 1), enable_events=True, default_text='595.00')],
                [sg.Push(), sg.Text("$"), sg.Input(key="Total 3", size=(DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Text("$"), sg.Input(key="3F Lifetime Protection Rider", size=(DOLLAR_WIDTH, 1), enable_events=True)]
            ])]
        ]
        
        pmt_select_layout = [
            [sg.Text("Payment Selection", font=("Helvetica", 16, "bold"))],
            [sg.Column([
                [sg.Text("Term:", size=(TEXT_WIDTH, 1))],
                [sg.Text("A. Single Pay:", size=(TEXT_WIDTH, 1))],
                [sg.Text("B. Time Pay:", size=(TEXT_WIDTH, 1))],
                [sg.Text("C. Single Pay Journey Home:", size=(TEXT_WIDTH, 1))],
                [sg.Text("D. LPR:", size=(TEXT_WIDTH, 1))],
                [sg.Text("Total (A+B+C+D):", size=(TEXT_WIDTH, 1))],
            ]), sg.Column([
                [sg.Push(), sg.Combo(["3-year", "5-year", "10-year", "15-year", "20-year"], 
                          key="Payment Term", size=(DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Text("$"), sg.Input(key="4A Single Pay", size=(DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Text("$"), sg.Input(key="4B Time Pay", size=(DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Text("$"), sg.Input(key="4C Single Pay Journey Home", size=(DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Text("$"), sg.Input(key="4D LPR", size=(DOLLAR_WIDTH, 1), enable_events=True)],
                [sg.Push(), sg.Text("$"), sg.Input(key="Total 4 \\(ABCD\\)", size=(DOLLAR_WIDTH, 1), enable_events=True)]
            ])]
        ]

        
        time_pay_layout = [
            [sg.Column(preplanned_amount_layout, expand_x=True)],
            [sg.HorizontalSeparator()],
            [sg.Column(pmt_select_layout, expand_x=True)]
        ]
        
        other_layout = [
            [sg.Text("Other Information", font=("Helvetica", 16, "bold"))],
            [sg.Text("Kearney Location:")],
            [sg.Combo(self.kearney_locations, key="Kearney Location", size=(40, 1), enable_events=True)],
            [sg.Text("Location Where Signed:")],
            [sg.Text("City:"), sg.Input(key="Signed City", size=(20, 1)),
             sg.Text("Province:"), sg.Input(key="Signed Province", size=(10, 1))]
        ]
        
        main_content = [
            [sg.Text("Preplanning PDF Autofiller", font=("Helvetica", 20, "bold"), justification='center', expand_x=True)],
            [sg.HorizontalSeparator()],
            [sg.Column([
                [sg.Frame("Personal Information", personal_info_layout, expand_x=True, expand_y=True)],
                [sg.HorizontalSeparator()],
                [sg.Frame("Packages", [
                    [sg.Text("Select Package:"), sg.Combo(list(self.packages.keys()), key="-PACKAGE-", enable_events=True)],
                    [sg.Text("Type of Service:"), sg.Input(key="Type of Service")],
                    [sg.Column(section_a_layout, scrollable=True, vertical_scroll_only=True, expand_x=True, expand_y=True),
                    sg.VerticalSeparator(),
                    sg.Column([
                        [sg.Frame("Section B", section_b_layout, expand_x=True)],
                        [sg.Frame("Section C", section_c_layout, expand_x=True)],
                        [sg.Frame("Section D", section_d_layout, expand_x=True)],
                        [sg.Frame("Total Charges", total_charges_layout, expand_x=True)]
                    ], scrollable=True, vertical_scroll_only=True, expand_x=True, expand_y=True)]
                ], expand_x=True, expand_y=True)]
            ], expand_x=True, expand_y=True),
            sg.Column([
                [sg.Frame("Time Pay", time_pay_layout, expand_x=True)],
                [sg.Frame("Other Information", other_layout, expand_x=True)]
            ], expand_x=True)]
        ]

        layout = [
            [sg.Column(main_content, scrollable=True, expand_x=True, expand_y=True, key='-MAIN-COLUMN-')],
            [sg.Button("Autofill PDFs", expand_x=True),
             sg.Button("Calculate Monthly Payment", key="-CALCULATE_MONTHLY_PAYMENT-", expand_x=True),
             sg.Button("Refresh Form", key="-REFRESH-", expand_x=True),
             sg.Button("Exit", expand_x=True)]
        ]

        self.window = sg.Window("Preplanning PDF Autofiller", layout, resizable=True, finalize=True)
        self.window.set_min_size((800, 600))  # Set a reasonable minimum size
        self.window.maximize()
        self.pdf_paths = self.initialize_pdf_paths()

        # Add this list of all package input keys
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
            "3F Lifetime Protection Rider", "4A Single Pay", "4B Time Pay",
            "4C Single Pay Journey Home", "4D LPR", "Total 4 \\(ABCD\\)"
    ]

        self.last_value = {key: '' for key in self.dollar_input_keys}
        locale.setlocale(locale.LC_ALL, '')  # Set the locale to the user's default
        
    def ordinal(self, n):
        if 10 <= n % 100 <= 20:
            suffix = 'th'
        else:
            suffix = {1: 'st', 2: 'nd', 3: 'rd'}.get(n % 10, 'th')
        return f"{n}{suffix}"

    def apply_package(self, package_name):
        try:
            # Clear all fields that might be affected by packages
            fields_to_clear = set()
            for package in self.packages.values():
                fields_to_clear.update(package.keys())
            
            for field in fields_to_clear:
                if field in self.window.AllKeysDict:
                    self.window[field].update('')
                    logging.info(f"Cleared field {field}")

            if package_name in self.packages:
                # Store the selected package name
                self.selected_package = package_name
                
                # Update the 'Type of Service' fields
                self.window['Type of Service'].update(package_name)
                
                package_data = self.packages[package_name]
                for key, value in package_data.items():
                    try:
                        if key in self.window.AllKeysDict:
                            self.window[key].update(value)
                            logging.info(f"Updated field {key} with value {value}")
                            # Trigger the format_dollar_field function for dollar amount fields
                            if key in self.dollar_input_keys:
                                self.format_dollar_field(key, value)
                    except Exception as e:
                        logging.error(f"Error updating field {key}: {str(e)}")
            
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
        }

    def run(self):
        while True:
            try:
                event, values = self.window.read()
                if event == sg.WINDOW_CLOSED or event == "Exit":
                    break
                elif event == "-REFRESH-":
                    self.refresh_form()
                elif event == "Kearney Location":
                    self.update_establishment_constants(values["Kearney Location"])
                elif event == "-PACKAGE-":
                    logging.info(f"Selected package: {values['-PACKAGE-']}")
                    self.apply_package(values["-PACKAGE-"])
                elif event == "-BIRTHDATE-":
                    self.update_age(values["-BIRTHDATE-"])
                elif event in self.dollar_input_keys:
                    self.format_dollar_field(event, values[event])
                elif event == "-CALCULATE_MONTHLY_PAYMENT-":
                    self.calculate_monthly_payment(values)
                elif event == "SIN":
                    self.window["SIN"].update(self.format_sin(values["SIN"]))
                elif event in ["-PHONE-", "Phone_3"]:
                    self.window[event].update(self.format_phone_number(values[event]))
                elif event in ["-POSTAL-", "Postal Code_4"]:
                    self.window[event].update(self.format_postal_code(values[event]))
                elif event == "-SAME_ADDRESS-":
                    self.toggle_beneficiary_address_fields(values["-SAME_ADDRESS-"])
                elif event == "Payment Term":
                    logging.info(f"Selected payment term: {values['Payment Term']}")
                elif event == "Autofill PDFs":
                    self.autofill_pdfs(values)
            except Exception as e:
                logging.error(f"Error in main event loop: {str(e)}")
                sg.popup_error(f"An unexpected error occurred: {str(e)}")

        self.window.close()
        
    def refresh_form(self):
        try:
            # Clear all input fields
            for key in self.window.AllKeysDict:
                if isinstance(self.window[key], sg.Input) or isinstance(self.window[key], sg.Multiline):
                    self.window[key].update('')
                elif isinstance(self.window[key], sg.Combo):
                    self.window[key].update(value='')
                elif isinstance(self.window[key], sg.Checkbox):
                    self.window[key].update(value=False)

            # Reset the package selection
            self.window["-PACKAGE-"].update(value='')
            self.selected_package = ""

            # Reset the Kearney Location
            self.window["Kearney Location"].update(value='')

            # Clear calculated fields
            for field in ["Total A", "Total B", "Total C", "Total D", "Total \\(ABC\\)", "Discount", "GST", "PST", "Total D_2", "Grand Total"]:
                self.window[field].update('')

            # Reset preplanned amount fields
            for field in ["3A Goods and Services", "3B MonumentMarker", "3C Other Expenses", "3D Final Documents Service", "Total 3", "3F Lifetime Protection Rider"]:
                self.window[field].update('')

            # Reset payment selection fields
            self.window["Payment Term"].update(value='')
            for field in ["4A Single Pay", "4B Time Pay", "4C Single Pay Journey Home", "4D LPR", "Total 4 \\(ABCD\\)"]:
                self.window[field].update('')

            logging.info("Form refreshed successfully")
            sg.popup("Form has been refreshed. You can start a new entry.")
        except Exception as e:
            logging.error(f"Error refreshing form: {str(e)}")
            sg.popup_error(f"An error occurred while refreshing the form: {str(e)}")
            
    def calculate_monthly_payment(self, values):
        try:
            # Step 1: Calculate GST and PST
            total_gst = 0
            total_pst = 0
            for field in self.gst_fields:
                try:
                    value = float(values[field].replace('$', '').replace(',', ''))
                    total_gst += value * 0.05
                    if field in self.pst_fields:
                        total_pst += value * 0.07
                except ValueError:
                    continue

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
            discount = float(values["Discount"].replace('$', '').replace(',', '') or 0)
            grand_total = total_abc - discount + total_gst + total_pst + totals['D']

            # Step 4: Calculate Preplanned Amount
            goods_and_services = grand_total
            monument_marker = float(values["3B MonumentMarker"].replace('$', '').replace(',', '') or 0)
            other_expenses = float(values["3C Other Expenses"].replace('$', '').replace(',', '') or 0)
            final_documents = float(values["3D Final Documents Service"].replace('$', '').replace(',', '') or 0)
            journey_home = float(values["3E Journey Home"].replace('$', '').replace(',', '') or 0)
            total_preplanned = goods_and_services + monument_marker + other_expenses + final_documents + journey_home

            # Step 5: Calculate Payment
            age = int(values["-AGE-"].strip())
            term = values["Payment Term"]

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

            factor = self.payment_factors[age_group][term]
            if factor is None:
                raise ValueError(f"No factor available for age {age} and term {term}")

            time_pay = total_preplanned * factor

            # Update all calculated values
            self.window["GST"].update(locale.format_string('%.2f', total_gst, grouping=True))
            self.window["PST"].update(locale.format_string('%.2f', total_pst, grouping=True))
            for section, total in totals.items():
                self.window[f"Total {section}"].update(locale.format_string('%.2f', total, grouping=True))
            self.window["Total \\(ABC\\)"].update(locale.format_string('%.2f', total_abc, grouping=True))
            self.window["Grand Total"].update(locale.format_string('%.2f', grand_total, grouping=True))
            self.window["3A Goods and Services"].update(locale.format_string('%.2f', goods_and_services, grouping=True))
            self.window["Total 3"].update(locale.format_string('%.2f', total_preplanned, grouping=True))
            self.window["4B Time Pay"].update(locale.format_string('%.2f', time_pay, grouping=True))

            # Calculate and update Total 4 (ABCD)
            total_4_abcd = sum([
                float(values["4A Single Pay"].replace('$', '').replace(',', '') or 0),
                time_pay,
                float(values["4C Single Pay Journey Home"].replace('$', '').replace(',', '') or 0),
                float(values["4D LPR"].replace('$', '').replace(',', '') or 0)
            ])
            self.window["Total 4 \\(ABCD\\)"].update(locale.format_string('%.2f', total_4_abcd, grouping=True))

            self.window.refresh()
            logging.info("Monthly payment calculation completed successfully")
            sg.popup("Monthly payment calculation completed successfully")
        except Exception as e:
            logging.error(f"Error in monthly payment calculation: {str(e)}")
            sg.popup_error(f"An error occurred during calculation: {str(e)}")

    def format_dollar_field(self, key, value):
        if value != self.last_value[key]:  # Only format if the value has changed
            if value and value.replace(',', '').replace('.', '', 1).isdigit():
                try:
                    # Remove existing commas and convert to float
                    clean_value = float(value.replace(',', ''))
                    
                    # Format the value with comma separators and two decimal places
                    formatted_value = locale.format_string('%.2f', clean_value, grouping=True)
                    
                    # Get the current cursor position
                    cursor_position = self.window[key].Widget.index(tk.INSERT)
                    
                    # Count the number of digits and commas before the cursor
                    original_parts = value[:cursor_position].split('.')
                    original_integer_part = original_parts[0].replace(',', '')
                    digits_and_commas = len(original_integer_part) + (len(original_integer_part) - 1) // 3
                    
                    # Update the input field with the formatted value
                    self.window[key].update(formatted_value)
                    
                    # Calculate new cursor position
                    new_position = min(digits_and_commas + (1 if '.' in value[:cursor_position] else 0), len(formatted_value))
                    
                    # Set the cursor position
                    self.window[key].Widget.icursor(new_position)
                    
                    # Update the last value
                    self.last_value[key] = formatted_value
                except ValueError:
                    pass
            else:
                # If the input is not a valid number, just update the last value
                self.last_value[key] = value
                
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
            1: f"{values['-FIRST-']}_{values['-LAST-']} - Protector Plus TruStage Application form.pdf",
            2: f"{values['-FIRST-']}_{values['-LAST-']} - Personal Information Sheet.pdf",
            3: f"{values['-FIRST-']}_{values['-LAST-']} - Instructions Concerning My Arrangements.pdf",
            4: f"{values['-FIRST-']}_{values['-LAST-']} - Pre-Arranged Funeral Service Agreement - Full Funeral Service.pdf"
        }

        output_pdfs = {
            i: output_dir / filename for i, filename in output_filenames.items()
        }

        try:
            age = self.calculate_age(values["-BIRTHDATE-"]) if values["-BIRTHDATE-"] else ""
            data = {k: v for k, v in values.items() if v}
            today = date.today()
            formatted_date = today.strftime("%B %d, %Y")

            # Create data dictionaries
            data_dicts = self.create_data_dictionaries(data, age, formatted_date, today)
            
            # Log the contents of data_dict1
            logging.info(f"Contents of data_dict1: {data_dicts[1]}")

            # Fill PDFs
            for i, (input_pdf, output_pdf) in enumerate(zip(self.pdf_paths.values(), output_pdfs.values()), 1):
                data_dict = data_dicts[i]
                fillpdfs.write_fillable_pdf(input_pdf, output_pdf, data_dict)
                logging.info(f"Filled PDF {i} written to: {output_pdf}")
                logging.info(f"Filled values for PDF {i}:")
                for field, value in data_dict.items():
                    logging.info(f"  {field}: {value}")

            sg.popup(f"""PDFs have been filled successfully.
Saved as:
{chr(10).join([f"{i}. {filename}" for i, filename in output_filenames.items()])}
In the Filled_Forms directory next to the application.""")
        except Exception as e:
            logging.error(f"Error filling PDFs: {str(e)}")
            sg.popup_error(f"Error filling PDFs: {str(e)}")

    def create_data_dictionaries(self, data, age, formatted_date, today):
        
        # Get the selected Kearney Location
        selected_location = data.get("Kearney Location", "")
        
        # Mapping of location names to dictionary keys
        location_mapping = {
        "Kearney Funeral Services (KFS)": "Kearney Vancouver Chapel",
        "Kearney Burnaby Chapel (KBC)": "Kearney Burnaby Chapel",
        "Kearney Burquitlam Funeral Home (BFH)": "Kearney Burquitlam Funeral Home",
        "Kearney Columbia-Bowell Chapel (CBC)": "Kearney ColumbiaBowell Chapel",
        "Kearney Cloverdale & South Surrey (CLO)": "Kearney Cloverdale South Surrey"
        }

        # Data dictionary for the first PDF
        data_dict1 = {
            'I understand that this is an enrollment into a group policy in order to provide funding for funeral expenses': 'On',
            'Establishment Name': ESTABLISHMENT_NAME,
            'Phone': ESTABLISHMENT_PHONE,
            'Email': ESTABLISHMENT_EMAIL,
            'Address': ESTABLISHMENT_ADDRESS,
            'City': ESTABLISHMENT_CITY,
            'Province': ESTABLISHMENT_PROVINCE,
            'Postal Code': ESTABLISHMENT_POSTAL_CODE,
            'First Name': data.get('-FIRST-', ''),
            'MI': data.get('-MIDDLE-', ''),
            'Last Name': data.get('-LAST-', ''),
            'Birthdate ddmmyy': data.get('-BIRTHDATE-', ''),
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
            '3F Lifetime Protection Rider': data.get('3F Lifetime Protection Rider', ''),
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
            'Location Where Signed': f"{data.get('Signed City', '')}, {data.get('Signed Province')}",
            'Date ddmmyy': formatted_date,
            'Representative Name': REPRESENTATIVE_NAME,
            'ID': REPRESENTATIVE_ID,
            'Phone_5': REPRESENTATIVE_PHONE,
            'Email_5': REPRESENTATIVE_EMAIL,
            'Date ddmmyy_3': formatted_date,
            'Payment \\(PAC\\)': 'On',
            'I hereby assign as its interest may lie the death benefit of the certificate applied for and to be issued to the funeral Establishment indicated above to provide funeral goods and': 'On',
            'I request that no new product be offered to me by TruStage Life of Canada or their affiliates or partners': 'On',
            'I also hereby assign the death benefit of the certificate to the funeral Establishment to provide certain cemetery goods and services and elect my certificate to be an EFA': 'On',
            'Coverage Type: Protector Plus not available on FEGA IP': 'On'
            
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

        # Data dictionary for the second PDF
        data_dict2 = {
            'Last name': data.get('-LAST-', ''),
            'First name': data.get('-FIRST-', ''),
            'Middle name': data.get('-MIDDLE-', ''),
            'Address': f"{data.get('-ADDRESS-', '')}, {data.get('-CITY-', '')}, {data.get('-PROVINCE-', '')}",
            'Postal Code': data.get('-POSTAL-', ''),
            'Phone': data.get('-PHONE-', ''),
            'Date of Birth': data.get('-BIRTHDATE-', ''),
            'Occupation': data.get('-OCCUPATION-', ''),
            'SIN': data.get('SIN', ''),
            'Death Certificate #': data.get('Death Certificates', '')[:1] if data.get('Death Certificates') else '',
            'Kearney ColumbiaBowell Chapel': '',
            'Kearney Vancouver Chapel': '',
            'Kearney Cloverdale South Surrey': '',
            'Kearney Burnaby Chapel': ''
        }
        
        if selected_location in location_mapping:
            data_dict2[location_mapping[selected_location]] = 'Yes'

        # Data dictionary for the third PDF
        data_dict3 = {
            'Date': formatted_date,
            'Name': f"{data.get('-FIRST-', '')} {data.get('-MIDDLE-', '')} {data.get('-LAST-', '')}",
            'Phone': data.get('-PHONE-', ''),
            'Email': data.get('-EMAIL-', ''),
            'Type of Service': data.get('Type of Service', ''),
            'Service to be Held at': ESTABLISHMENT_NAME,
            'Address': f"{ESTABLISHMENT_ADDRESS}, {ESTABLISHMENT_CITY}, {ESTABLISHMENT_PROVINCE}, {ESTABLISHMENT_POSTAL_CODE}",
            'Death Certificates': data.get('Death Certificates', '')[:1] if data.get('Death Certificates') else '',
            'Casket': '',
            'Urn': '',
            'Kearney ColumbiaBowell Chapel': '',
            'Kearney Vancouver Chapel': '',
            'Kearney Cloverdale South Surrey': '',
            'Kearney Burnaby Chapel': ''
        }
        
        if selected_location in location_mapping:
            data_dict3[location_mapping[selected_location]] = 'Yes'
        
        if data.get('Casket'):
            if data.get('B1'):
                data_dict3['Casket'] = f"{data.get('Casket')} - ${data.get('B1')}"
            else:
                data_dict3['Casket'] = data.get('Casket')
        
        if data.get('Urn'):
            if data.get('B2'):
                data_dict3['Urn'] = f"{data.get('Urn')} - ${data.get('B2')}"
            else:
                data_dict3['Urn'] = data.get('Urn')

        # Data dictionary for the fourth PDF (Pre-Arranged Funeral Service Agreement)
        data_dict4 = {
            'Purchaser': f"{data.get('-FIRST-', '')} {data.get('-MIDDLE-', '')} {data.get('-LAST-', '')}",
            'PURCHASERS NAME': f"{data.get('-FIRST-', '')} {data.get('-MIDDLE-', '')} {data.get('-LAST-', '')}",
            'Phone Number': data.get('-PHONE-', ''),
            'Address': f"{data.get('-ADDRESS-', '')}, {data.get('-CITY-', '')}, {data.get('-PROVINCE-', '')}, {data.get('-POSTAL-', '')}",
            'Type of Service': data.get('Type of Service', ''),
            'FUNERAL HOME REPRESENTATIVE NAME': REPRESENTATIVE_NAME,
            'BENEFICIARY': f"{data.get('-FIRST-', '')} {data.get('-MIDDLE-', '')} {data.get('-LAST-', '')}",
            'DATE OF BIRTH': data.get('-BIRTHDATE-', ''),
            'ADDRESS CITY PROVINCE POSTAL CODE': f"{data.get('-ADDRESS-', '')}, {data.get('-CITY-', '')}, {data.get('-PROVINCE-', '')}, {data.get('-POSTAL-', '')}",
            'TELEPHONE NUMBER': data.get('-PHONE-', ''),
            'Day': self.ordinal(today.day),
            'Month': today.strftime("%B"),
            'Year': str(today.year),
            'SIN': data.get('SIN', ''),
            # New fields from the package layout
            'A1': data.get('A1', ''),
            'A2A': data.get('A2A', ''),
            'Pallbearers': data.get('Pallbearers', ''),
            'A2B': data.get('A2B', ''),
            'Alternate Day Interment 1': data.get('Alternate Day Interment 1', ''),
            'A2C': data.get('A2C', ''),
            'Alternate Day Interment 2': data.get('Alternate Day Interment 2', ''),
            'A2D': data.get('A2D', ''),
            'A3': data.get('A3', ''),
            'A4A': data.get('A4A', ''),
            'A4B': data.get('A4B', ''),
            'A4C': data.get('A4C', ''),
            'A5A': data.get('A5A', ''),
            'A5B': data.get('A5B', ''),
            'Pacemaker Removal': data.get('Pacemaker Removal', ''),
            'A5C': data.get('A5C', ''),
            'Autopsy Care': data.get('Autopsy Care', ''),
            'A5D': data.get('A5D', ''),
            'Evening Prayers or Visitation': data.get('Evening Prayers or Visitation', ''),
            'A6': data.get('A6', ''),
            'Weekend or Statutory Holiday': data.get('Weekend or Statutory Holiday', ''),
            'A7': data.get('A7', ''),
            'Reception Facilities': data.get('Reception Facilities', ''),
            'A8': data.get('A8', ''),
            'Delivery of Cremated Remains': data.get('Delivery of Cremated Remains', ''),
            'A9A': data.get('A9A', ''),
            'Transfer to Crematorium or Airport': data.get('Transfer to Crematorium or Airport', ''),
            'A9B': data.get('A9B', ''),
            'Lead Vehicle': data.get('Lead Vehicle', ''),
            'A9C': data.get('A9C', ''),
            'Service Vehicle': data.get('Service Vehicle', ''),
            'A9D': data.get('A9D', ''),
            'Funeral Coach': data.get('Funeral Coach', ''),
            'A9E': data.get('A9E', ''),
            'Limousine': data.get('Limousine', ''),
            'A9F': data.get('A9F', ''),
            'Additional Limousines': data.get('Additional Limousines', ''),
            'A9G': data.get('A9G', ''),
            'Flower Van': data.get('Flower Van', ''),
            'A9H': data.get('A9H', ''),
            'Total A': data.get('Total A', ''),
            'Casket': data.get('Casket', ''),
            'B1': data.get('B1', ''),
            'Urn': data.get('Urn', ''),
            'B2': data.get('B2', ''),
            'Keepsake': data.get('Keepsake', ''),
            'B3': data.get('B3', ''),
            'Traditional Mourning Items': data.get('Traditional Mourning Items', ''),
            'B4': data.get('B4', ''),
            'Memorial Stationary': data.get('Memorial Stationary', ''),
            'B5': data.get('B5', ''),
            'Funeral Register 1': data.get('Funeral Register 1', ''),
            'B6': data.get('B6', ''),
            'Funeral Register 2': data.get('Funeral Register 2', ''),
            'B7': data.get('B7', ''),
            'Total B': data.get('Total B', ''),
            'Cemetery': data.get('Cemetery', ''),
            'C1': data.get('C1', ''),
            'Crematorium': data.get('Crematorium', ''),
            'C2': data.get('C2', ''),
            'Obituary Notices': data.get('Obituary Notices', ''),
            'C3': data.get('C3', ''),
            'Flowers': data.get('Flowers', ''),
            'C4': data.get('C4', ''),
            'CPBC Administration Fee': data.get('CPBC Administration Fee', ''),
            'C5': data.get('C5', ''),
            'Hostess': data.get('Hostess', ''),
            'C6': data.get('C6', ''),
            'Markers': data.get('Markers', ''),
            'C7': data.get('C7', ''),
            'Catering 1': data.get('Catering 1', ''),
            'C8': data.get('C8', ''),
            'Catering 2': data.get('Catering 2', ''),
            'C9': data.get('C9', ''),
            'Catering 3': data.get('Catering 3', ''),
            'C10': data.get('C10', ''),
            'Total C': data.get('Total C', ''),
            'Clergy Honorarium': data.get('Clergy Honorarium', ''),
            'D1': data.get('D1', ''),
            'Church Honorarium': data.get('Church Honorarium', ''),
            'D2': data.get('D2', ''),
            'Altar Servers': data.get('Altar Servers', ''),
            'D3': data.get('D3', ''),
            'Organist': data.get('Organist', ''),
            'D4': data.get('D4', ''),
            'Soloist': data.get('Soloist', ''),
            'D5': data.get('D5', ''),
            'Harpist': data.get('Harpist', ''),
            'D6': data.get('D6', ''),
            'Death Certificates': data.get('Death Certificates', ''),
            'D7': data.get('D7', ''),
            'Other': data.get('Other', ''),
            'D8': data.get('D8', ''),
            'Total D': data.get('Total D', ''),
            'Total \\(ABC\\)': data.get('Total \\(ABC\\)', ''),
            'Discount': data.get('Discount', ''),
            'GST': data.get('GST', ''),
            'PST': data.get('PST', ''),
            'Total D_2': data.get('Total D_2', ''),
            'Grand Total': data.get('Grand Total', ''),
            'City Province': f"{data.get('Signed City', '')}, {data.get('Signed Province')}",
            'Kearney Vancouver Chapel': '',
            'Kearney ColumbiaBowell Chapel': '',
            'Kearney Burquitlam Funeral Home': '',
            'Kearney Burnaby Chapel': '',
            'Kearney Cloverdale South Surrey': ''
        }
        
        if selected_location in location_mapping:
            data_dict4[location_mapping[selected_location]] = 'On'

        return {1: data_dict1, 2: data_dict2, 3: data_dict3, 4: data_dict4}

if __name__ == "__main__":
    autofiller = PDFAutofiller()
    autofiller.run()