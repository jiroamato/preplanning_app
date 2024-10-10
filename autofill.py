import PySimpleGUI as sg
from PyPDF2 import PdfReader, PdfWriter
import os
import logging

class PDFAutofiller:
    def __init__(self):
        self.layout = [
            [sg.Text("First Name:"), sg.Input(key="-FIRST-")],
            [sg.Text("Middle Name:"), sg.Input(key="-MIDDLE-")],
            [sg.Text("Last Name:"), sg.Input(key="-LAST-")],
            [sg.Button("Autofill PDF"), sg.Button("Exit")]
        ]
        self.window = sg.Window("PDF Autofiller", self.layout)

    def run(self):
        while True:
            event, values = self.window.read()
            if event == sg.WINDOW_CLOSED or event == "Exit":
                break
            if event == "Autofill PDF":
                self.autofill_pdf(values["-FIRST-"], values["-MIDDLE-"], values["-LAST-"])
        self.window.close()

    def autofill_pdf(self, first_name, middle_name, last_name):
        pdf_file = os.path.join("Forms", "Protector Plus TruStage Application form.pdf")
        self.fill_pdf(pdf_file, first_name, middle_name, last_name)
        sg.popup("PDF has been autofilled successfully!")

    def fill_pdf(self, pdf_file, first_name, middle_name, last_name):
        try:
            reader = PdfReader(pdf_file)
            writer = PdfWriter()

            for page_num, page in enumerate(reader.pages):
                writer.add_page(page)
                if '/Annots' in page:
                    for annot in page['/Annots']:
                        try:
                            obj = annot.get_object()
                            if obj.get('/Subtype') == '/Widget':
                                if '/T' in obj:
                                    field_name = obj['/T'].lower()
                                    logging.info(f"Found field: {field_name} on page {page_num + 1}")
                                    if "first name" in field_name:
                                        obj.update({'/V': first_name})
                                        logging.info(f"Filled first name: {first_name}")
                                    elif "middle name" in field_name:
                                        obj.update({'/V': middle_name})
                                        logging.info(f"Filled middle name: {middle_name}")
                                    elif "last name" in field_name:
                                        obj.update({'/V': last_name})
                                        logging.info(f"Filled last name: {last_name}")
                                    elif "name" in field_name:
                                        full_name = f"{first_name} {middle_name} {last_name}".strip()
                                        obj.update({'/V': full_name})
                                        logging.info(f"Filled full name: {full_name}")
                        except Exception as e:
                            logging.error(f"Error processing field on page {page_num + 1}: {str(e)}")

            output_dir = os.path.join("Filled_Forms")
            os.makedirs(output_dir, exist_ok=True)
            output_file_path = os.path.join(output_dir, f"filled_Protector_Plus_TruStage_Application_form.pdf")
            
            with open(output_file_path, "wb") as output_file:
                writer.write(output_file)
            logging.info(f"Wrote filled PDF: {output_file_path}")
        except Exception as e:
            logging.error(f"Error processing file {pdf_file}: {str(e)}")

if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
    autofiller = PDFAutofiller()
    autofiller.run()
