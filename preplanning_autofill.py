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
    PERSONAL_INFO_SHEET = "Personal Information Sheet Auto Fill.pdf"
    INSTRUCTIONS = "Instructions Concerning My Arrangements.pdf"


class PDFAutofiller:
    def __init__(self):
        self.setup_logging()
        self.base_path = self.get_base_path()

        try:
            import inflect
            self.inflect_engine = inflect.engine()
        except ImportError:
            logging.error(
                "The 'inflect' library is not installed. Please install it using 'pip install inflect'.")
            sg.popup_error(
                "The 'inflect' library is not installed. Please install it using 'pip install inflect'.")
            return

        applicant_layout = [
            [sg.Text("Applicant Information", font=("Helvetica", 14, "bold"))],
            [sg.Text("First Name:"), sg.Input(key="-FIRST-", size=(20, 1))],
            [sg.Text("Middle Name:"), sg.Input(key="-MIDDLE-", size=(20, 1))],
            [sg.Text("Last Name:"), sg.Input(key="-LAST-", size=(20, 1))],
            [sg.Text("Birthdate:"), sg.Input(key="-BIRTHDATE-", size=(20, 1))],
            [sg.Text("Gender:"), sg.Input(key="-GENDER-", size=(20, 1))],
            [sg.Text("SIN:"), sg.Input(key="SIN", size=(20, 1), enable_events=True)],
            [sg.Text("Phone:"), sg.Input(key="-PHONE-", size=(20, 1), enable_events=True)],
            [sg.Text("Email:"), sg.Input(key="-EMAIL-", size=(20, 1))],
            [sg.Text("Address:"), sg.Input(key="-ADDRESS-", size=(20, 1))],
            [sg.Text("City:"), sg.Input(key="-CITY-", size=(20, 1))],
            [sg.Text("Province:"), sg.Input(key="-PROVINCE-", size=(20, 1))],
            [sg.Text("Postal Code:"), sg.Input(key="-POSTAL-", size=(20, 1), enable_events=True)],
            [sg.Text("Occupation:"), sg.Input(key="-OCCUPATION-", size=(20, 1))]
        ]

        beneficiary_layout = [
            [sg.Text("Beneficiary Information", font=("Helvetica", 14, "bold"))],
            [sg.Text("Full Name:"), sg.Input(key="Name", size=(20, 1))],
            [sg.Text("Relationship:"), sg.Input(key="Relationship", size=(20, 1))],
            [sg.Text("Phone:"), sg.Input(key="Phone_3", size=(20, 1), enable_events=True)],
            [sg.Text("Email:"), sg.Input(key="Email_3", size=(20, 1))],
            [sg.Checkbox("Same address as Applicant", key="-SAME_ADDRESS-", enable_events=True)],
            [sg.Text("Address:"), sg.Input(key="Address \\(if different\\)", size=(20, 1))],
            [sg.Text("City:"), sg.Input(key="City_4", size=(20, 1))],
            [sg.Text("Province:"), sg.Input(key="Province_4", size=(20, 1))],
            [sg.Text("Postal Code:"), sg.Input(key="Postal Code_4", size=(20, 1), enable_events=True)]
        ]

        personal_info_layout = [
            [sg.Column(applicant_layout), sg.VSeperator(), sg.Column(beneficiary_layout)]
        ]

        # Packages
        package_layout = [
            [sg.Text("Packages", font=("Helvetica", 16,
                     "bold"), justification='center', expand_x=True)],
            [sg.HorizontalSeparator()],
            # A. Professional Services
            [sg.Text("A. Professional Services", font=("Helvetica", 16, "bold")), sg.Text("(GST applicable)", font=("Helvetica", 10, "italic"))],
            [sg.Text("    1. Securing Release of Deceased and Transfer:", size=(40, 1)), sg.Text("$"), sg.Input(key="A1", size=(10, 1), enable_events=True)],
            [sg.Text("    2. Services of Licensed Funeral Directors and Staff:", size=(40, 1)), sg.Text("$"), sg.Input(key="A2A", size=(10, 1), enable_events=True)],
            [sg.Text("        Pallbearers:", size=(20, 1)), sg.Input(key="Pallbearers", size=(20, 1)), sg.Text("$"), sg.Input(key="A2B", size=(10, 1), enable_events=True)],
            [sg.Text("        Alternate Day Interment:", size=(40, 1)), sg.Input(key="Alternate Day Interment 1", size=(20, 1)), sg.Text("$"), sg.Input(key="A2C", size=(10, 1), enable_events=True)],
            [sg.Text("        Alternate Day Interment 2:", size=(40, 1)), sg.Input(key="Alternate Day Interment 2", size=(20, 1)), sg.Text("$"), sg.Input(key="A2D", size=(10, 1), enable_events=True)],
            [sg.Text("    3. Administration, Documentation & Registration:", size=(40,1)), sg.Text("$"), sg.Input(key="A3", size=(10,1), enable_events=True)],
            [sg.Text("    4. Facilities and/or Equipment and Supplies:", size=(40,1)), sg.Text("$"), sg.Input(key="A4A", size=(10,1), enable_events=True    )],
            [sg.Text("        Sheltering of Remains:", size=(40,1)), sg.Text("$"), sg.Input(key="A4B", size=(10,1), enable_events=True)],
            [sg.Text("        A/V Equipment:", size=(40,1)), sg.Text("$"), sg.Input(key="A4C", size=(10,1), enable_events=True)],
            [sg.Text("    5. Preparation Services")],
            [sg.Text("        a) Basic Sanitary Care, Dressing and Casketing:", size=(40,1)), sg.Text("$"), sg.Input(key="A5A", size=(10,1), enable_events=True)],
            [sg.Text("        b) Embalming:", size=(40,1)), sg.Text("$"), sg.Input(key="A5B", size=(10,1), enable_events=True)],
            [sg.Text("        c) Pacemaker Removal:", size=(40,1)), sg.Input(key="Pacemaker Removal", size=(20,1)), sg.Text("$"), sg.Input(key="A5C", size=(10,1), enable_events=True)],
            [sg.Text("        d) Autopsy Care:", size=(40,1)), sg.Input(key="Autopsy Care", size=(20,1)), sg.Text("$"), sg.Input(key="A5D", size=(10,1), enable_events=True )],
            [sg.Text("    6. Evening Prayers or Visitation:", size=(40,1)), sg.Input(key="Evening Prayers or Visitation", size=(20,1)), sg.Text("$"), sg.Input(key="A6", size=(10,1), enable_events=True)],
            [sg.Text("    7. Weekend or Statutory Holiday:", size=(40,1)), sg.Input(key="Weekend or Statutory Holiday", size=(20,1)), sg.Text("$"), sg.Input(key="A7", size=(10,1), enable_events=True)],
            [sg.Text("    8. Reception Facilities:", size=(40,1)), sg.Input(key="Reception Facilities", size=(20,1)), sg.Text("$"), sg.Input(key="A8", size=(10,1), enable_events=True)],
            [sg.Text("    9. Vehicles")],
            [sg.Text("        Delivery of Cremated Remains:", size=(40,1)), sg.Input(key="Delivery of Cremated Remains", size=(20, 1)), sg.Text("$"), sg.Input(key="A9A", size=(10,1), enable_events=True)],
            [sg.Text("        Transfer Vehicle for Transfer to Crematorium or Airport:", size=(40,1)), sg.Input(key="Transfer to Crematorium or Airport", size=(20, 1)), sg.Text("$"), sg.Input(key="A9B", size=(10,1), enable_events=True)],
            [sg.Text("        Lead Vehicle:", size=(40,1)), sg.Input(key="Lead Vehicle", size=(20, 1)), sg.Text("$"), sg.Input(key="A9C", size=(10,1), enable_events=True)],
            [sg.Text("        Service Vehicle:", size=(40,1)), sg.Input(key="Service Vehicle", size=(20, 1)), sg.Text("$"), sg.Input(key="A9D", size=(10,1), enable_events=True)],
            [sg.Text("        Funeral Coach:", size=(40,1)), sg.Input(key="Funeral Coach", size=(20, 1)), sg.Text("$"), sg.Input(key="A9E", size=(10,1), enable_events=True)],
            [sg.Text("        Limousine:", size=(40,1)), sg.Input(key="Limousine", size=(20, 1)), sg.Text("$"), sg.Input(key="A9F", size=(10,1), enable_events=True)],
            [sg.Text("        Additional Limousines:", size=(40,1)), sg.Input(key="Additional Limousines", size=(20, 1)), sg.Text("$"), sg.Input(key="A9G", size=(10,1), enable_events=True)],
            [sg.Text("        Flower Van:", size=(40,1)), sg.Input(key="Flower Van", size=(20, 1)), sg.Text("$"), sg.Input(key="A9H", size=(10,1), enable_events=True)],
            [sg.Text("   TOTAL A:", font=("italic"), size=(40,1)), sg.Text("$"), sg.Input(key="Total A", size=(10,1), enable_events=True)],
            [sg.VPush()],
            # B. Merchandise
            [sg.Text("B. Merchandise", font=("Helvetica", 16, "bold")), sg.Text("(GST and/or PST applicable)", font=("Helvetica", 10, "italic"))],
            [sg.Text("    Casket:", size=(20,1)), sg.Input(key="Casket", size=(20,1)), sg.Text("$"), sg.Input(key="B1", size=(10,1), enable_events=True)],
            [sg.Text("    Urn:", size=(20,1)), sg.Input(key="Urn", size=(20,1)), sg.Text("$"), sg.Input(key="B2", size=(10,1), enable_events=True)],
            [sg.Text("    Keepsake:", size=(20,1)), sg.Input(key="Keepsake", size=(20,1)), sg.Text("$"), sg.Input(key="B3", size=(10,1), enable_events=True)],
            [sg.Text("    Traditional Mourning Items:", size=(40,1)), sg.Input(key="Traditional Mourning Items", size=(20,1)), sg.Text("$"), sg.Input(key="B4", size=(10,1), enable_events=True)],
            [sg.Text("    Memorial Stationary:", size=(40,1)), sg.Input(key="Memorial Stationary", size=(20,1)), sg.Text("$"), sg.Input(key="B5", size=(10,1), enable_events=True)],
            [sg.Text("    Funeral Register:", size=(40,1)), sg.Input(key="Funeral Register 1", size=(20,1)), sg.Text("$"), sg.Input(key="B6", size=(10,1), enable_events=True)],
            [sg.Text("    Funeral Register 2:", size=(40,1)), sg.Input(key="Funeral Register 2", size=(20,1)), sg.Text("$"), sg.Input(key="B7", size=(10,1), enable_events=True)],
            [sg.Text("   TOTAL B:", font=("italic"), size=(40,1)), sg.Text("$"), sg.Input(key="Total B", size=(10,1), enable_events=True)],
            [sg.VPush()],
            # C. Cash Disbursements
            [sg.Text("C. Cash Disbursements", font=("Helvetica", 16, "bold")), sg.Text("(GST and/or PST applicable)", font=("Helvetica", 10, "italic"))],
            [sg.Text("    Cemetery:", size=(40,1)), sg.Input(key="Cemetery", size=(20,1)), sg.Text("$"), sg.Input(key="C1", size=(10,1), enable_events=True)],
            [sg.Text("    Crematorium:", size=(40,1)), sg.Input(key="Crematorium", size=(20,1)), sg.Text("$"), sg.Input(key="C2", size=(10,1), enable_events=True)],
            [sg.Text("    Obituary Notices:", size=(40,1)), sg.Input(key="Obituary Notices", size=(20,1)), sg.Text("$"), sg.Input(key="C3", size=(10,1), enable_events=True)],
            [sg.Text("    Flowers:", size=(40,1)), sg.Input(key="Flowers", size=(20,1)), sg.Text("$"), sg.Input(key="C4", size=(10,1), enable_events=True)],
            [sg.Text("    CPBC Administration Fee:", size=(40,1)), sg.Input(key="CPBC Administration Fee", size=(20,1)), sg.Text("$"), sg.Input(key="C5", size=(10,1), enable_events=True)],
            [sg.Text("    Hostess:", size=(40,1)), sg.Input(key="Hostess", size=(20,1)), sg.Text("$"), sg.Input(key="C6", size=(10,1), enable_events=True)],
            [sg.Text("    Markers:", size=(40,1)), sg.Input(key="Markers", size=(20,1)), sg.Text("$"), sg.Input(key="C7", size=(10,1), enable_events=True)],
            [sg.Text("    Catering:", size=(40,1)), sg.Input(key="Catering 1", size=(20,1)), sg.Text("$"), sg.Input(key="C8", size=(10,1), enable_events=True)],
            [sg.Text("    Catering 2:", size=(40,1)), sg.Input(key="Catering 2", size=(20,1)), sg.Text("$"), sg.Input(key="C9", size=(10,1), enable_events=True)],
            [sg.Text("    Catering 3:", size=(40,1)), sg.Input(key="Catering 3", size=(20,1)), sg.Text("$"), sg.Input(key="C10", size=(10,1), enable_events=True)],
            [sg.Text("   TOTAL C:", font=("italic"), size=(40,1)), sg.Text("$"), sg.Input(key="Total C", size=(10,1), enable_events=True)],
            [sg.VPush()],
            # D. Cash Disbursements
            [sg.Text("D. Cash Disbursements", font=("Helvetica", 16, "bold")), sg.Text("(GST exempt)", font=("Helvetica", 10, "italic"))],
            [sg.Text("    Clergy Honorarium:", size=(40,1)), sg.Input(key="Clergy Honorarium", size=(20,1)), sg.Text("$"), sg.Input(key="D1", size=(10,1), enable_events=True)],
            [sg.Text("    Church Honorarium:", size=(40,1)), sg.Input(key="Church Honorarium", size=(20,1)), sg.Text("$"), sg.Input(key="D2", size=(10,1), enable_events=True)],
            [sg.Text("    Altar Servers:", size=(40,1)), sg.Input(key="Altar Servers", size=(20,1)), sg.Text("$"), sg.Input(key="D3", size=(10,1), enable_events=True)],
            [sg.Text("    Organist:", size=(40,1)), sg.Input(key="Organist", size=(20,1)), sg.Text("$"), sg.Input(key="D4", size=(10,1), enable_events=True)],
            [sg.Text("    Soloist:", size=(40,1)), sg.Input(key="Soloist", size=(20,1)), sg.Text("$"), sg.Input(key="D5", size=(10,1), enable_events=True)],
            [sg.Text("    Harpist:", size=(40,1)), sg.Input(key="Harpist", size=(20,1)), sg.Text("$"), sg.Input(key="D6", size=(10,1), enable_events=True)],
            [sg.Text("    Death Certificates:", size=(40,1)), sg.Input(key="Death Certificates", size=(20,1)), sg.Text("$"), sg.Input(key="D7", size=(10,1), enable_events=True)],
            [sg.Text("    Other:", size=(40,1)), sg.Input(key="Other", size=(20,1)), sg.Text("$"), sg.Input(key="D8", size=(10,1), enable_events=True)],
            [sg.Text("   TOTAL D:", size=(40,1)), sg.Text("$"), sg.Input(key="Total D", size=(10,1), enable_events=True)],
            [sg.VPush()],
            # TOTAL CHARGES
            [sg.Text("TOTAL CHARGES", font=("Helvetica", 16, "bold"))],
            [sg.Text("    TOTAL A, B, & C SECTIONS:", size=(40,1)), sg.Text("$"), sg.Input(key="Total \\(ABC\\)", size=(10,1), enable_events=True)],
            [sg.Text("    DISCOUNT:", size=(40,1)), sg.Text("$"), sg.Input(key="Discount", size=(10,1), enable_events=True)],
            [sg.Text("    G.S.T.:", size=(40,1)), sg.Text("$"), sg.Input(key="GST", size=(10,1), enable_events=True)],
            [sg.Text("    P.S.T.:", size=(40,1)), sg.Text("$"), sg.Input(key="PST", size=(10,1), enable_events=True)],
            [sg.Text("    TOTAL D:", size=(40,1)), sg.Text("$"), sg.Input(key="Total D_2", size=(10,1), enable_events=True)],
            [sg.Text("    GRAND TOTAL:", size=(40,1)), sg.Text("$"), sg.Input(key="Grand Total", size=(10,1), enable_events=True)]
        ]

        layout = [
            [sg.Text("Preplanning PDF Autofiller", font=("Helvetica", 20, "bold"), justification='center', expand_x=True)],
            [sg.HorizontalSeparator()],
            [sg.Column(personal_info_layout)],
            [sg.HorizontalSeparator()],
            [sg.TabGroup([[
                sg.Tab("Package Details", package_layout),
            ]])],
            [sg.Button("Autofill PDFs"), sg.Button("Exit")]
        ]

        self.window = sg.Window("Preplanning PDF Autofiller", layout, resizable=True, size=(1050, 800))
        self.pdf_paths = self.initialize_pdf_paths()

        # Add this list of all package input keys
        self.dollar_input_keys = [
            "A1", "A2A", "A2B", "A2C", "A2D", "A3", "A4A", "A4B", "A4C",
            "A5A", "A5B", "A5C", "A5D", "A6", "A7", "A8", "A9A", "A9B",
            "A9C", "A9D", "A9E", "A9F", "A9G", "A9H", "Total A",
            "B1", "B2", "B3", "B4", "B5", "B6", "B7", "Total B",
            "C1", "C2", "C3", "C4", "C5", "C6", "C7", "C8", "C9", "C10", "Total C",
            "D1", "D2", "D3", "D4", "D5", "D6", "D7", "D8", "Total D",
            "Total \\(ABC\\)", "Discount", "GST", "PST", "Total D_2", "Grand Total"
    ]

        self.last_value = {key: '' for key in self.dollar_input_keys}
        locale.setlocale(locale.LC_ALL, '')  # Set the locale to the user's default

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
            event, values = self.window.read()
            if event == sg.WINDOW_CLOSED or event == "Exit":
                break
            elif event in self.dollar_input_keys:
                self.format_dollar_field(event, values[event])
            elif event == "SIN":
                self.window["SIN"].update(self.format_sin(values["SIN"]))
            elif event in ["-PHONE-", "Phone_3"]:
                self.window[event].update(self.format_phone_number(values[event]))
            elif event in ["-POSTAL-", "Postal Code_4"]:
                self.window[event].update(self.format_postal_code(values[event]))
            elif event == "-SAME_ADDRESS-":
                self.toggle_beneficiary_address_fields(values["-SAME_ADDRESS-"])
            elif event == "Autofill PDFs":
                self.autofill_pdfs(values)

        self.window.close()

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

    def toggle_beneficiary_address_fields(self, same_address):
        for key in ["Address \\(if different\\)", "City_4", "Province_4", "Postal Code_4"]:
            self.window[key].update(disabled=same_address)

    def autofill_pdfs(self, values):
        if not all(os.path.exists(pdf) for pdf in self.pdf_paths.values()):
            logging.error(f"One or more PDF files not found. Base path: {self.base_path}")
            sg.popup_error(f"One or more PDF files not found. Please check the Forms directory.\nLooking in: {self.base_path}")
            return

        # Create 'Filled_Forms' folder
        if getattr(sys, 'frozen', False):
            output_dir = Path(sys.executable).parent / "Filled_Forms"
        else:
            output_dir = Path(os.getcwd()) / "Filled_Forms"
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
        # Data dictionary for the first PDF
        data_dict1 = {
            'Establishment Name': ESTABLISHMENT_NAME,
            'Phone': ESTABLISHMENT_PHONE,
            'Email': ESTABLISHMENT_EMAIL,
            'Address': ESTABLISHMENT_ADDRESS,
            'City': ESTABLISHMENT_CITY,
            'Province': ESTABLISHMENT_PROVINCE,
            'Postal Code': ESTABLISHMENT_POSTAL_CODE,
            'First Name': data.get('-FIRST-', ''),
            'MI': data.get('-MIDDLE-', '')[:1] if data.get('-MIDDLE-') else '',
            'Last Name': data.get('-LAST-', ''),
            'Birthdate ddmmyy': data.get('-BIRTHDATE-', ''),
            'Age': str(age),
            'Gender': data.get('-GENDER-', ''),
            'SIN': data.get('SIN', ''),
            'Phone_2': data.get('-PHONE-', ''),
            'Email_2': data.get('-EMAIL-', ''),
            'Mailing Address': data.get('-ADDRESS-', ''),
            'City_2': data.get('-CITY-', ''),
            'Province_2': data.get('-PROVINCE-', ''),
            'Postal Code_2': data.get('-POSTAL-', ''),
            'Occupation': data.get('-OCCUPATION-', ''),
            'Name': data.get('Name', ''),
            'Relationship': data.get('Relationship', ''),
            'Phone_3': data.get('Phone_3', ''),
            'Email_3': data.get('Email_3', ''),
            'Representative Name': REPRESENTATIVE_NAME,
            'ID': REPRESENTATIVE_ID,
            'Phone_5': REPRESENTATIVE_PHONE,
            'Email_5': REPRESENTATIVE_EMAIL,
            'Date ddmmyy_3': formatted_date
        }

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
            'SIN': data.get('SIN', '')
        }

        # Data dictionary for the third PDF
        data_dict3 = {
            'Date': formatted_date,
            'Name': f"{data.get('-FIRST-', '')} {data.get('-MIDDLE-', '')} {data.get('-LAST-', '')}",
            'Phone': data.get('-PHONE-', ''),
            'Email': data.get('-EMAIL-', ''),
            'Address': f"{data.get('-ADDRESS-', '')}, {data.get('-CITY-', '')}, {data.get('-PROVINCE-', '')}, {data.get('-POSTAL-', '')}"
        }

        # Data dictionary for the fourth PDF (Pre-Arranged Funeral Service Agreement)
        data_dict4 = {
            'Purchaser': f"{data.get('-FIRST-', '')} {data.get('-MIDDLE-', '')} {data.get('-LAST-', '')}",
            'PURCHASERS NAME': f"{data.get('-FIRST-', '')} {data.get('-MIDDLE-', '')} {data.get('-LAST-', '')}",
            'Phone Number': data.get('-PHONE-', ''),
            'Address': f"{data.get('-ADDRESS-', '')}, {data.get('-CITY-', '')}, {data.get('-PROVINCE-', '')}, {data.get('-POSTAL-', '')}",
            'FUNERAL HOME REPRESENTATIVE NAME': REPRESENTATIVE_NAME,
            'BENEFICIARY': f"{data.get('-FIRST-', '')} {data.get('-MIDDLE-', '')} {data.get('-LAST-', '')}",
            'DATE OF BIRTH': data.get('-BIRTHDATE-', ''),
            'ADDRESS CITY PROVINCE POSTAL CODE': f"{data.get('-ADDRESS-', '')}, {data.get('-CITY-', '')}, {data.get('-PROVINCE-', '')}, {data.get('-POSTAL-', '')}",
            'TELEPHONE NUMBER': data.get('-PHONE-', ''),
            'Day': self.inflect_engine.ordinal(today.day),
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
            'Grand Total': data.get('Grand Total', '')
        }

        return {1: data_dict1, 2: data_dict2, 3: data_dict3, 4: data_dict4}

if __name__ == "__main__":
    autofiller = PDFAutofiller()
    autofiller.run()