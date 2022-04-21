
""" Class that Handles PDF Manipulation """

from PyPDF2 import PdfFileReader
from pikepdf import Pdf
import fitz
import utils.regextools as rt


class PdfHandler:
    def __init__(self):
        pass

    def read_in_pdf(self):
        pass

    @staticmethod
    def get_pdf_as_list(path):
        with fitz.open(path) as doc:
            text = ""
            for page in doc:
                text += page.get_text()

        return text.split("\n")

    @staticmethod
    def decrypt_pdf(path, file, password):
        """ Decrypts a PDF Document if it's protected """
        pdf_path = f"{path}/{file}"
        pdf_file = PdfFileReader(pdf_path)

        # Check if the opened file is actually Encrypted
        if pdf_file.isEncrypted:
            with Pdf.open(pdf_path, password=password, allow_overwriting_input=True) as pdf:
                pdf.save(pdf_path)
            print("File decrypted Successfully.")

        else:
            print("File already decrypted.")

    @staticmethod
    def read_out_pdf_list(pdf_list):
        for idx, line in enumerate(pdf_list):
            print(f"{idx}: {line}")

    @staticmethod
    def check_if_encrypted(path, file):
        pdf_file = PdfFileReader(f"{path}/{file}")
        return pdf_file.isEncrypted


class IdPdfHandler(PdfHandler):
    def __init__(self):
        super(IdPdfHandler, self).__init__()
        pass

    def read_id_pdf(self, path, document):

        text_list = self.get_pdf_as_list(f"{path}/{document}")
        # self.read_out_pdf_list(text_list)

        sin_place, sin = rt.find_first_pattern(text_list, "[0-9]{3} [0-9]{3} [0-9]{3}")
        email_place, email = rt.find_first_pattern(text_list, "[0-z]*@[0-z]*(.com)")
        start_spot, a = rt.find_first_pattern(text_list, "Client code")
        client_code_place, client_code = rt.find_first_pattern(text_list[start_spot:start_spot + 10], "^[0-9]*$")

        email = rt.check_for_instance(email_place, email, "No Email")
        client_code = rt.check_for_instance(client_code_place, client_code, "")
        first_name = text_list[sin_place - 2]
        last_name = text_list[sin_place - 1]
        sin = sin.replace(' ', '')

        file_name = f"{last_name.replace(' ', '_')}_{first_name.replace(' ', '_')}"
        last_init = last_name[0]
        client_folder_path = f"{last_name.title()}, {first_name.title()}_{client_code}"

        print(f"    Name:{first_name} {last_name}\n"
              f"    SIN: {sin}\n"
              f"    Client Code: {client_code}\n"
              f"    Email: {email}")

        return first_name, last_name, sin, email, client_code, file_name, last_init, client_folder_path
