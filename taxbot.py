import os
import shutil
import win32com.client
import fitz
from pathlib import Path
from PyPDF2 import PdfFileWriter, PdfFileReader

from automation import Automation


class TaxBot(Automation):
    def __init__(self):
        super().__init__()

        # Source Path should be static
        self.source_path = "Q:/Admin - Digital Technology and Risk Advisory/Tax Bot/Tax Test"

    def run(self):

        # while not self.check_for_00_form():
        #     return
        
        self.read_out_pdf()
        return

        # 0. Read in the pdf and set the name variables ####
        first_name, last_name, sin, email = self.read_taxprep_pdf()

        sin = sin.replace(' ', '')
        sin_pw = sin[:-4]  # last 4 digits
        file_name = f"{last_name}_{first_name.replace(' ', '_')}"
        last_init = last_name[0]
        client_folder_path = f"{last_name}, {first_name.split(' ')[0]}"
        year = 2021
        output_folder = "_FINAL T1 DOCS"

        # path_name = "Holman, Kimberly"
        # name = "Holman_Kimberly_A"

        # 1. Create destination path: ####
        destination_path, city = self.select_directory(client_folder_path, year, last_init, output_folder)
        if not destination_path:
            print(f"Can't find the directory for {client_folder_path}")
            return

        print(f"Destination path successfully found for {client_folder_path}")
        # self.destination_path = "Q:/Admin - Digital Technology and Risk Advisory/Tax Bot/Tax Test/Test Folder"
        # self.destination_path = "I:/Toronto/_Personal Tax - TOR/H/Holman, Kimberly_94214/2020/CRA"

        # 2. move the pdf files ####

        # file_1 = file_name + f"_1-Ltr_{year}.pdf"
        # file_2 = file_name + f"_2-T1_({city})_{year}.pdf"
        # file_3 = file_name + f"_3-Docs_Needing_Signature_{year}.pdf"
        # files = [file_1, file_2, file_3]

        files = self.get_matching_files(file_name)
        self.move_pdfs(self.source_path, destination_path, files)

        # 3. Decrypt the letter pdf and extract the information
        # letter_name = f"_1-Ltr_{year}.pdf"
        letter_name = file_name + f"_2-T1_({city})_{year -1}.pdf"
        print(letter_name)
        self.decrypt_pdf(destination_path, letter_name, sin_pw)

        return

        # 3. Send the email with attachments
        to_address = "david.watson@bakertillywm.com"
        subject = f"Tax report for: {first_name.split(' ')[0]}"
        html_body = '<h3>This is HTML Body</h3>'
        body = f"email to send to {email} \n\n" \
               "Dear Kimberly,\n\n" \
               "Please find attached your tax return report \n \n" \
               "Kind Regards"

        self.send_email(to_address=to_address,
                        subject=subject,
                        html_body=html_body,
                        body=body,
                        attachment_files=files,
                        attachment_file_path=destination_path
                        )
        print("email sent")
        return

    @staticmethod
    def move_pdfs(path, destination, files):
        for file in files:
            shutil.move(os.path.join(path, file), destination)

    @staticmethod
    def send_email(**kwargs):

        outlook = win32com.client.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)

        mail.To = kwargs["to_address"]
        mail.Subject = kwargs['subject']
        # mail.HTMLBody = kwargs['html_body']
        mail.Body = kwargs['body']
        for attachment in kwargs["attachment_files"]:
            mail.Attachments.Add(kwargs["attachment_file_path"] + "/" + attachment)
        mail.Send()

    def read_taxprep_pdf(self):
        with fitz.open(self.source_path + "/00-Taxprep-Holman, Kim.pdf") as doc:
            text = ""
            for page in doc:
                text += page.getText()

        text_list = text.split("\n")
        first_name = text_list[29]
        last_name = text_list[30]
        sin = text_list[31]
        email = text_list[108]
        # for idx, line in enumerate(text_list):
        #    print(f"{idx}: {line}")
        print(f"{first_name}, {last_name}: {sin}, {email}")

        return first_name, last_name, sin, email

    def read_out_pdf(self):

        with fitz.open(self.source_path + "/Holman_Kimberly_A_1-Ltr_2020.pdf") as doc:
            text = ""
            for page in doc:
                text += page.getText()

        text_list = text.split("\n")
        # first_name = text_list[29]
        # last_name = text_list[30]
        # sin = text_list[31]
        # email = text_list[108]
        for idx, line in enumerate(text_list):
            print(f"{idx}: {line}")

        # print(f"{first_name}, {last_name}: {sin}, {email}")

        return

    def select_directory(self, path_name, year, last_init, folder):

        tor = self.check_directory(f"I:/Toronto/_Personal Tax - TOR/{last_init}", path_name)
        van = self.check_directory(f"I:/Vancouver/_Personal Tax - VAN/{last_init}", path_name)
        if van:
            path = f"I:/Vancouver/_Personal Tax - VAN/{last_init}/{van}/{year}/{folder}"
            self.create_directory(path)
            city = "Vancouver"
        elif tor:
            path = f"I:/Toronto/_Personal Tax - TOR/{last_init}/{tor}/{year}/{folder}"
            self.create_directory(path)
            city = "Toronto"
        else:
            return False, None

        return path, city

    @staticmethod
    def check_directory(directory, path_to_check):
        """ Checks if the folder/file exists in the directory"""
        for folder in os.listdir(directory):
            if folder.startswith(path_to_check):
                return folder
        return False

    @staticmethod
    def create_directory(directory):
        Path(directory).mkdir(parents=True, exist_ok=True)

    def get_matching_files(self, matching_string):
        files = []
        for file in os.listdir(self.source_path):
            if file.startswith(matching_string):
                files.append(file)

        return files

    def decrypt_pdf(self, path, file, password):

        out = PdfFileWriter()
        if not self.check_directory(path, file):
            print("No Cover Letter Found")
            return

        file = PdfFileReader(f"{path}/{file}")

        # Check if the opened file is actually Encrypted
        if file.isEncrypted:

            file.decrypt(password)
            for idx in range(file.numPages):
                page = file.getPage(idx)
                out.addPage(page)

            # remove the file
            os.remove(f"{path}/{file}")

            with open(f"{path}/{file}", "wb") as f:
                out.write(f)

            print("File decrypted Successfully.")
        else:
            print("File already decrypted.")

# ToDo Test the pdf decrypter
# ToDo Read and get the statement information from the letter