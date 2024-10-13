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

        self.layout = [
            [sg.Text("Preplanning PDF Autofiller", font=(
                "Helvetica", 20, "bold"), justification='center', expand_x=True)],
            [sg.HorizontalSeparator()],
            [sg.VPush()],
            [sg.Text("Applicant Information", font=("Helvetica", 16,
                     "bold"), justification='center', expand_x=True)],
            [sg.HorizontalSeparator()],
            [sg.Text("First Name:"), sg.Input(key="-FIRST-")],
            [sg.Text("Middle Name:"), sg.Input(key="-MIDDLE-")],
            [sg.Text("Last Name:"), sg.Input(key="-LAST-")],
            [sg.Text("Birthdate (e.g., January 1, 1990):"),
             sg.Input(key="-BIRTHDATE-")],
            [sg.Text("Gender:"), sg.Input(key="-GENDER-")],
            [sg.Text("SIN:"), sg.Input(key="SIN", enable_events=True)],
            [sg.Text("Phone:"), sg.Input(key="-PHONE-", enable_events=True)],
            [sg.Text("Email:"), sg.Input(key="-EMAIL-")],
            [sg.Text("Address:"), sg.Input(key="-ADDRESS-")],
            [sg.Text("City:"), sg.Input(key="-CITY-")],
            [sg.Text("Province:"), sg.Input(key="-PROVINCE-")],
            [sg.Text("Postal Code:"), sg.Input(
                key="-POSTAL-", enable_events=True)],
            [sg.Text("Occupation:"), sg.Input(key="-OCCUPATION-")],
            [sg.VPush()],
            [sg.Text("Beneficiary Information", font=("Helvetica", 16,
                     "bold"), justification='center', expand_x=True)],
            [sg.HorizontalSeparator()],
            [sg.Text("Full Name:"), sg.Input(key="Name")],
            [sg.Text("Relationship to the Applicant:"),
             sg.Input(key="Relationship")],
            [sg.Text("Phone:"), sg.Input(key="Phone_3", enable_events=True)],
            [sg.Text("Email:"), sg.Input(key="Email_3")],
            [sg.Checkbox("Beneficiary address same as Applicant's",
                         key="-SAME_ADDRESS-", enable_events=True, default=False)],
            [sg.Text("Address:"), sg.Input(key="Address \\(if different\\)")],
            [sg.Text("City:"), sg.Input(key="City_4")],
            [sg.Text("Province:"), sg.Input(key="Province_4")],
            [sg.Text("Postal Code:"), sg.Input(
                key="Postal Code_4", enable_events=True)],
            [sg.Button("Autofill PDFs"), sg.Button("Exit")]
        ]
        self.window = sg.Window("Preplanning PDF Autofiller", self.layout)
        self.pdf_paths = self.initialize_pdf_paths()

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
            if event == "-PHONE-" or event == "Phone_3":
                values[event] = self.format_phone(values[event])
                self.window[event].update(values[event])
            if event == "SIN":
                values["SIN"] = self.format_sin(values["SIN"])
                self.window["SIN"].update(values["SIN"])
            if event == "-POSTAL-" or event == "Postal Code_4":
                values[event] = self.format_postal_code(values[event])
                self.window[event].update(values[event])
            if event == "-SAME_ADDRESS-":
                self.toggle_beneficiary_address_fields(
                    values["-SAME_ADDRESS-"])
            if event == "Autofill PDFs":
                if not self.validate_email(values["-EMAIL-"]):
                    sg.popup_error("Invalid email format for Applicant Email")
                    continue
                if not self.validate_email(values["Email_3"]):
                    sg.popup_error(
                        "Invalid email format for Beneficiary Email")
                    continue
                self.autofill_pdfs(values)
        self.window.close()

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

    def format_phone(self, phone):
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

        # Format the date components for the fourth PDF
        day_ordinal = self.inflect_engine.ordinal(today.day)
        month = today.strftime("%B")
        year = today.year

        # Data dictionary for the fourth PDF
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
            'Day': day_ordinal,
            'Month': month,
            'Year': str(year),
        }

        return {1: data_dict1, 2: data_dict2, 3: data_dict3, 4: data_dict4}


if __name__ == "__main__":
    autofiller = PDFAutofiller()
    autofiller.run()