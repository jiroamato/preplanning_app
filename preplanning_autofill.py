import sys
import os
import PySimpleGUI as sg
from fillpdf import fillpdfs
from pathlib import Path
import logging
from datetime import datetime
import re
from dateutil import parser
from dateutil.relativedelta import relativedelta

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

class PDFAutofiller:
    def __init__(self):
        self.setup_logging()
        self.base_path = self.get_base_path()
        self.layout = [
            [sg.Text("Preplanning PDF Autofiller", font=("Helvetica", 20, "bold"), justification='center', expand_x=True)],
            [sg.HorizontalSeparator()],
            [sg.VPush()],
            [sg.Text("Applicant Information", font=("Helvetica", 16, "bold"), justification='center', expand_x=True)],
            [sg.HorizontalSeparator()],
            [sg.Text("First Name:"), sg.Input(key="-FIRST-")],
            [sg.Text("Middle Name:"), sg.Input(key="-MIDDLE-")],
            [sg.Text("Last Name:"), sg.Input(key="-LAST-")],
            [sg.Text("Birthdate (e.g., January 1, 1990):"), sg.Input(key="-BIRTHDATE-")],
            [sg.Text("Gender:"), sg.Input(key="-GENDER-")],
            [sg.Text("SIN:"), sg.Input(key="SIN", enable_events=True)],
            [sg.Text("Phone:"), sg.Input(key="-PHONE-", enable_events=True)],
            [sg.Text("Email:"), sg.Input(key="-EMAIL-")],
            [sg.Text("Address:"), sg.Input(key="-ADDRESS-")],
            [sg.Text("City:"), sg.Input(key="-CITY-")],
            [sg.Text("Province:"), sg.Input(key="-PROVINCE-")],
            [sg.Text("Postal Code:"), sg.Input(key="-POSTAL-", enable_events=True)],
            [sg.Text("Occupation:"), sg.Input(key="-OCCUPATION-")],
            [sg.VPush()],
            [sg.Text("Beneficiary Information", font=("Helvetica", 16, "bold"), justification='center', expand_x=True)],
            [sg.HorizontalSeparator()],
            [sg.Text("Full Name:"), sg.Input(key="Name")],
            [sg.Text("Relationship to the Applicant:"), sg.Input(key="Relationship")],
            [sg.Text("Phone:"), sg.Input(key="Phone_3", enable_events=True)],
            [sg.Text("Email:"), sg.Input(key="Email_3")],
            [sg.Checkbox("Beneficiary address same as Applicant's", key="-SAME_ADDRESS-", enable_events=True)],
            [sg.Text("Address:"), sg.Input(key="4 Payment Selection", disabled=True)],
            [sg.Text("City:"), sg.Input(key="City_4", disabled=True)],
            [sg.Text("Province:"), sg.Input(key="Province_4", disabled=True)],
            [sg.Text("Postal Code:"), sg.Input(key="Postal Code_4", disabled=True, enable_events=True)],
            [sg.Button("Autofill PDFs"), sg.Button("Exit")]
        ]
        self.window = sg.Window("Preplanning PDF Autofiller", self.layout)

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
                self.toggle_beneficiary_address_fields(values["-SAME_ADDRESS-"])
                if values["-SAME_ADDRESS-"]:
                    self.window["Postal Code_4"].update(values["-POSTAL-"])
            if event == "Autofill PDFs":
                if not self.validate_email(values["-EMAIL-"]):
                    sg.popup_error("Invalid email format for Applicant Email")
                    continue
                if not self.validate_email(values["Email_3"]):
                    sg.popup_error("Invalid email format for Beneficiary Email")
                    continue
                self.autofill_pdfs(values)
        self.window.close()

    def validate_inputs(self, values):
        if values["-EMAIL-"] and not self.validate_email(values["-EMAIL-"]):
            sg.popup_error("Invalid email address. Please enter a valid email or leave it blank.")
            return False
        if values["-BIRTHDATE-"] and not self.validate_email(values["-BIRTHDATE-"]):
            sg.popup_error("Invalid birthdate. Please use the format 'Month Day, Year' (e.g., January 1, 1990) or leave it blank.")
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
        for key in ["4 Payment Selection", "City_4", "Province_4", "Postal Code_4"]:
            self.window[key].update(disabled=same_address)
            if same_address:
                self.window[key].update("")

    def autofill_pdfs(self, values):
        pdf1_path = os.path.join(self.base_path, "Forms", "Protector Plus TruStage Application form.pdf")
        pdf2_path = os.path.join(self.base_path, "Forms", "Personal Information Sheet Auto Fill.pdf")
        pdf3_path = os.path.join(self.base_path, "Forms", "Instructions Concerning My Arrangements.pdf")
        pdf4_path = os.path.join(self.base_path, "Forms", "Pre-Arranged Funeral Service Agreement - New.pdf")
        
        if not all(os.path.exists(pdf) for pdf in [pdf1_path, pdf2_path, pdf3_path, pdf4_path]):
            logging.error(f"One or more PDF files not found. Base path: {self.base_path}")
            sg.popup_error(f"One or more PDF files not found. Please check the Forms directory.\nLooking in: {self.base_path}")
            return

        # Create 'Filled_Forms' folder in the same directory as the executable or script
        if getattr(sys, 'frozen', False):
            # If running as executable
            output_dir = Path(sys.executable).parent / "Filled_Forms"
        else:
            # If running as script
            output_dir = Path(os.getcwd()) / "Filled_Forms"
        output_dir.mkdir(exist_ok=True)

        output_filename1 = f"{values['-FIRST-']}_{values['-LAST-']} - Protector Plus TruStage Application form.pdf"
        output_filename2 = f"{values['-FIRST-']}_{values['-LAST-']} - Personal Information Sheet.pdf"
        output_filename3 = f"{values['-FIRST-']}_{values['-LAST-']} - Instructions Concerning My Arrangements.pdf"
        output_filename4 = f"{values['-FIRST-']}_{values['-LAST-']} - Pre-Arranged Funeral Service Agreement.pdf"
        
        output_pdf1 = output_dir / output_filename1
        output_pdf2 = output_dir / output_filename2
        output_pdf3 = output_dir / output_filename3
        output_pdf4 = output_dir / output_filename4

        try:
            age = self.calculate_age(values["-BIRTHDATE-"]) if values["-BIRTHDATE-"] else ""
            
            # Create a dictionary with only non-empty values
            data = {k: v for k, v in values.items() if v}
            
            # Data dictionary for the first PDF
            data_dict1 = {
                'Establishment Name': ESTABLISHMENT_NAME,
                'Phone': ESTABLISHMENT_PHONE,
                'Email': ESTABLISHMENT_EMAIL,
                '1 Applicant': ESTABLISHMENT_ADDRESS,
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
                'Representative Name': REPRESENTATIVE_NAME,
                # New beneficiary fields
                'Name': data.get('Name', ''),
                'Relationship': data.get('Relationship', ''),
                'Phone_3': data.get('Phone_3', ''),
                'Email_3': data.get('Email_3', '')
            }
            
            # Handle beneficiary address
            if data.get('-SAME_ADDRESS-'):
                data_dict1['4 Payment Selection'] = data.get('-ADDRESS-', '')
                data_dict1['City_4'] = data.get('-CITY-', '')
                data_dict1['Province_4'] = data.get('-PROVINCE-', '')
                data_dict1['Postal Code_4'] = data.get('-POSTAL-', '')
            else:
                data_dict1['4 Payment Selection'] = data.get('4 Payment Selection', '')
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
                'Name': f"{data.get('-FIRST-', '')} {data.get('-MIDDLE-', '')} {data.get('-LAST-', '')}",
                'Phone': data.get('-PHONE-', ''),
                'Email': data.get('-EMAIL-', ''),
                'Address': f"{data.get('-ADDRESS-', '')}, {data.get('-CITY-', '')}, {data.get('-PROVINCE-', '')}, {data.get('-POSTAL-', '')}"
            }

            # Data dictionary for the fourth PDF
            data_dict4 = {
                'Purchaser': f"{data.get('-FIRST-', '')} {data.get('-MIDDLE-', '')} {data.get('-LAST-', '')}",
                'PURCHASERS NAME': f"{data.get('-FIRST-', '')} {data.get('-MIDDLE-', '')} {data.get('-LAST-', '')}",
                'Phone Number': data.get('-PHONE-', ''),
                'Address': f"{data.get('-ADDRESS-', '')}, {data.get('-CITY-', '')}, {data.get('-PROVINCE-', '')}, {data.get('-POSTAL-', '')}",
                'FUNERAL HOME REPRESENTATIVE NAME': REPRESENTATIVE_NAME
            }

            # Fill the first PDF
            fillpdfs.write_fillable_pdf(pdf1_path, output_pdf1, data_dict1)
            logging.info(f"Filled first PDF written to: {output_pdf1}")

            # Fill the second PDF
            fillpdfs.write_fillable_pdf(pdf2_path, output_pdf2, data_dict2)
            logging.info(f"Filled second PDF written to: {output_pdf2}")

            # Fill the third PDF
            fillpdfs.write_fillable_pdf(pdf3_path, output_pdf3, data_dict3)
            logging.info(f"Filled third PDF written to: {output_pdf3}")

            # Fill the fourth PDF
            fillpdfs.write_fillable_pdf(pdf4_path, output_pdf4, data_dict4)
            logging.info(f"Filled fourth PDF written to: {output_pdf4}")

            logging.info("Filled values for first PDF:")
            for field, value in data_dict1.items():
                logging.info(f"  {field}: {value}")

            logging.info("Filled values for second PDF:")
            for field, value in data_dict2.items():
                logging.info(f"  {field}: {value}")

            logging.info("Filled values for third PDF:")
            for field, value in data_dict3.items():
                logging.info(f"  {field}: {value}")

            logging.info("Filled values for fourth PDF:")
            for field, value in data_dict4.items():
                logging.info(f"  {field}: {value}")

            sg.popup(f"PDFs have been filled successfully.\nSaved as:\n1. {output_filename1}\n2. {output_filename2}\n3. {output_filename3}\n4. {output_filename4}\nIn the Filled_Forms directory next to the application.")
        except Exception as e:
            logging.error(f"Error filling PDFs: {str(e)}")
            sg.popup_error(f"Error filling PDFs: {str(e)}")

if __name__ == "__main__":
    autofiller = PDFAutofiller()
    autofiller.run()